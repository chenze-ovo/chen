import hashlib
from pathlib import Path

from fastapi import APIRouter, Body, Depends, File, Form, HTTPException, Query, UploadFile
from fastapi.responses import FileResponse, Response
from sqlalchemy import func, select
from sqlalchemy.orm import Session, selectinload

from app.config import get_settings
from app.database import get_db
from app.deps import get_current_user, require_roles
from app.models import BATCH_FILE_KINDS, BatchFile, EngineRun, ExceptionReview, ImportBatch, User, utcnow
from app.serialize import batch_detail
from app.services.project1_runner import aggregate_status, convert_files
from app.services.project2_runner import generate_oversea_copy_for_batch, reexport_project2_batch, run_project2_batch

router = APIRouter(prefix="/api/batches", tags=["batches"])

_MAX_BYTES = 30 * 1024 * 1024


def _safe_name(raw: str | None, index: int, used: set[str]) -> str:
    name = Path(raw or "").name.replace("\x00", "")
    if not name or name in {".", ".."}:
        name = f"upload-{index}.bin"
    candidate = name
    seq = 2
    while candidate in used:
        stem = Path(name).stem
        suffix = Path(name).suffix
        candidate = f"{stem}-{seq}{suffix}"
        seq += 1
    used.add(candidate)
    return candidate


def _load_batch(db: Session, batch_id: int) -> ImportBatch | None:
    stmt = (
        select(ImportBatch)
        .where(ImportBatch.id == batch_id)
        .options(selectinload(ImportBatch.files), selectinload(ImportBatch.creator))
    )
    return db.scalars(stmt).first()


@router.get("")
def list_batches(
    page: int = Query(1, ge=1),
    page_size: int = Query(20, ge=1, le=100),
    _user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    total = int(db.scalar(select(func.count(ImportBatch.id))) or 0)
    stmt = (
        select(ImportBatch)
        .options(selectinload(ImportBatch.files), selectinload(ImportBatch.creator))
        .order_by(ImportBatch.id.desc())
        .offset((page - 1) * page_size)
        .limit(page_size)
    )
    rows = db.scalars(stmt).all()
    return {
        "items": [batch_detail(row) for row in rows],
        "total": total,
        "page": page,
        "page_size": page_size,
    }


@router.get("/{batch_id}")
def get_batch(
    batch_id: int,
    _user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    batch = _load_batch(db, batch_id)
    if batch is None:
        raise HTTPException(status_code=404, detail="批次不存在")
    return batch_detail(batch)


@router.post("", status_code=201)
async def create_batch(
    project: str = Form(...),
    note: str = Form(""),
    file_kind: str = Form("unspecified"),
    copy_order_nos: str = Form(""),
    files: list[UploadFile] = File(...),
    same_day: list[UploadFile] | None = File(None),
    user: User = Depends(require_roles("ops")),
    db: Session = Depends(get_db),
) -> dict:
    if project not in {"project1", "project2"}:
        raise HTTPException(status_code=400, detail="project 只能是 project1 或 project2")
    if file_kind not in BATCH_FILE_KINDS:
        raise HTTPException(status_code=400, detail="不支持的文件种类")
    if not files:
        raise HTTPException(status_code=400, detail="请至少上传一个文件")

    payloads: list[tuple[str, bytes]] = []
    used: set[str] = set()
    for index, upload in enumerate(files, start=1):
        data = await upload.read()
        if len(data) > _MAX_BYTES:
            raise HTTPException(status_code=400, detail=f"文件过大：{upload.filename or index}")
        if not data:
            raise HTTPException(status_code=400, detail=f"文件是空的：{upload.filename or index}")
        payloads.append((_safe_name(upload.filename, index, used), data, file_kind))
    for index, upload in enumerate(same_day or [], start=1):
        if not upload.filename and not upload.size:
            continue
        data = await upload.read()
        if not data:
            continue
        if len(data) > _MAX_BYTES:
            raise HTTPException(status_code=400, detail=f"当天店小秘文件过大：{upload.filename or index}")
        payloads.append((_safe_name(upload.filename, index, used), data, "dianxiaomi_same_day"))

    batch = ImportBatch(
        project=project,
        status="created",
        note=note.strip() or None,
        created_by_id=user.id,
        created_at=utcnow(),
    )
    db.add(batch)
    db.flush()

    settings = get_settings()
    upload_dir = settings.data_dir / "uploads" / str(batch.id)
    upload_dir.mkdir(parents=True, exist_ok=True)
    stored_paths: list[Path] = []
    for name, data, kind in payloads:
        relative = Path("uploads") / str(batch.id) / name
        target = settings.data_dir / relative
        target.write_bytes(data)
        db.add(
            BatchFile(
                batch_id=batch.id,
                original_name=name,
                stored_path=relative.as_posix(),
                size_bytes=len(data),
                sha256=hashlib.sha256(data).hexdigest(),
                kind=kind,
            )
        )
        if kind != "dianxiaomi_same_day":
            stored_paths.append(target)
    db.flush()

    if project == "project1":
        summary = convert_files(stored_paths, settings.data_dir / "outputs" / str(batch.id))
        status = aggregate_status(summary)
        downloads = []
        for item in summary.get("files", []):
            for key, label in (("dx", "店小秘"), ("d01", "店管家01"), ("d02", "店管家02")):
                path = (item.get("outputs") or {}).get(key)
                if path:
                    downloads.append({"label": f"{item.get('file')} · {label}", "path": path})
            export_status = item.get("export_status")
            if export_status in {"blocked", "needs_review", "error"}:
                db.add(
                    ExceptionReview(
                        batch_id=batch.id,
                        project="project1",
                        code=str(export_status).upper(),
                        severity="warning" if export_status == "needs_review" else "error",
                        message=(
                            f"{item.get('file')}：order_convert 导出状态 {export_status}。"
                            "ERP 不改写规则，请下载复核文件。"
                        ),
                        status="open",
                        created_at=utcnow(),
                    )
                )
        summary["downloads"] = downloads
    else:
        order_paths = stored_paths if file_kind != "dianxiaomi_same_day" else []
        selected = [x.strip() for x in (copy_order_nos or "").replace("\n", ",").split(",") if x.strip()]
        summary = run_project2_batch(
            db,
            order_paths,
            settings.data_dir / "outputs" / str(batch.id),
            batch.id,
            copy_order_nos=selected or None,
        )
        stats = summary.get("stats") or {}
        if not stats.get("visible_orders"):
            status = "needs_review" if summary.get("anomalies") else "converted"
        elif stats.get("human_review") or stats.get("review"):
            status = "needs_review"
        elif summary.get("executed"):
            status = "converted"
        else:
            status = "failed"

    batch.status = status
    batch.engine_status = status
    batch.engine_summary = summary
    db.add(
        EngineRun(
            project=project,
            batch_id=batch.id,
            status=status,
            summary=summary,
            created_at=utcnow(),
        )
    )
    db.commit()
    stored = _load_batch(db, batch.id)
    if stored is None:
        raise HTTPException(status_code=500, detail="批次写入后读取失败")
    return batch_detail(stored)




@router.post("/{batch_id}/oversea-copy")
def create_oversea_copy(
    batch_id: int,
    payload: dict = Body(...),
    user: User = Depends(require_roles("ops")),
    db: Session = Depends(get_db),
) -> dict:
    """勾选订单号，按已确认规则生成「海外仓订单 - 副本」（仅选中可见行）。"""
    order_nos = payload.get("order_nos") or payload.get("copy_order_nos") or []
    if not isinstance(order_nos, list) or not order_nos:
        raise HTTPException(status_code=400, detail="请至少勾选一个订单号")
    cleaned = [str(x).strip() for x in order_nos if str(x).strip()]
    if not cleaned:
        raise HTTPException(status_code=400, detail="请至少勾选一个订单号")
    batch = _load_batch(db, batch_id)
    if batch is None:
        raise HTTPException(status_code=404, detail="批次不存在")
    if batch.project != "project2":
        raise HTTPException(status_code=400, detail="仅项目2可生成海外仓副本")
    try:
        result = generate_oversea_copy_for_batch(db, batch_id, cleaned)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    stored = _load_batch(db, batch_id)
    return {
        "result": result,
        "batch": batch_detail(stored) if stored is not None else None,
    }


@router.post("/{batch_id}/reexport")
def reexport_batch(
    batch_id: int,
    payload: dict = Body(None),
    user: User = Depends(require_roles("ops", "reviewer")),
    db: Session = Depends(get_db),
) -> dict:
    """审核填完金额后重导海外仓批次（含店管家02）。"""
    payload = payload or {}
    order_nos = payload.get("order_nos") or payload.get("copy_order_nos") or []
    cleaned = [str(x).strip() for x in order_nos if str(x).strip()] if isinstance(order_nos, list) else []
    try:
        summary = reexport_project2_batch(db, batch_id, cleaned or None)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    stored = _load_batch(db, batch_id)
    return {"summary": summary, "batch": batch_detail(stored) if stored is not None else None}


@router.get("/{batch_id}/download-folder")
def download_batch_folder(
    batch_id: int,
    _user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """把本批输出打成 zip（海外仓/运营/店管家/加工说明.txt），不因待审核阻塞。"""
    batch = _load_batch(db, batch_id)
    if batch is None:
        raise HTTPException(status_code=404, detail="批次不存在")
    settings = get_settings()
    data_dir = settings.data_dir.resolve()
    summary = dict(batch.engine_summary or {})
    from app.services.project2_packaging import (
        build_result_zip_bytes,
        ensure_notes_in_summary,
        resolve_out_dir,
    )

    out_dir = resolve_out_dir(data_dir, batch_id, summary)
    if batch.project == "project2":
        summary = ensure_notes_in_summary(out_dir, summary, data_dir, batch_id=batch_id)
        # persist notes path into summary without changing status / review gate
        if summary != (batch.engine_summary or {}):
            batch.engine_summary = summary
            from sqlalchemy.orm.attributes import flag_modified

            flag_modified(batch, "engine_summary")
            db.add(batch)
            db.commit()
        payload, filename = build_result_zip_bytes(out_dir, summary, batch_id=batch_id)
    else:
        # project1: zip flat outputs dir (no nested folders required)
        import io
        import zipfile
        from datetime import datetime
        from zoneinfo import ZoneInfo

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            if out_dir.is_dir():
                for path in sorted(out_dir.rglob("*")):
                    if not path.is_file():
                        continue
                    if path.name.startswith("_") or "_internal" in path.parts:
                        continue
                    zf.write(path, arcname=path.relative_to(out_dir).as_posix())
        stamp = datetime.now(ZoneInfo("Asia/Shanghai")).strftime("%Y%m%d")
        filename = f"店小秘加工结果-批次{batch_id}-{stamp}.zip"
        payload = buf.getvalue()

    headers = {
        "Content-Disposition": f"attachment; filename*=UTF-8''{_url_quote(filename)}",
        "X-Review-Gate": "bypassed",
    }
    return Response(content=payload, media_type="application/zip", headers=headers)



@router.get("/{batch_id}/result-folder")
def get_result_folder(
    batch_id: int,
    _user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    """返回本批结果目录绝对路径（已是普通文件夹，含表与加工说明，不打包）。"""
    batch = _load_batch(db, batch_id)
    if batch is None:
        raise HTTPException(status_code=404, detail="批次不存在")
    settings = get_settings()
    data_dir = settings.data_dir.resolve()
    summary = dict(batch.engine_summary or {})
    from app.services.project2_packaging import ensure_notes_in_summary, resolve_out_dir

    out_dir = resolve_out_dir(data_dir, batch_id, summary)
    if not out_dir.is_dir():
        raise HTTPException(status_code=404, detail="批次结果文件夹不存在")
    if batch.project == "project2":
        summary = ensure_notes_in_summary(out_dir, summary, data_dir, batch_id=batch_id)
        if summary != (batch.engine_summary or {}):
            batch.engine_summary = summary
            from sqlalchemy.orm.attributes import flag_modified

            flag_modified(batch, "engine_summary")
            db.add(batch)
            db.commit()
    files = sorted(
        [
            path.name
            for path in out_dir.iterdir()
            if path.is_file() and not path.name.startswith("_")
        ]
    )
    return {
        "path": str(out_dir.resolve()),
        "batch_id": batch_id,
        "project": batch.project,
        "files": files,
        "notes": "加工说明.txt" if (out_dir / "加工说明.txt").is_file() else None,
    }


@router.post("/{batch_id}/open-result-folder")
def open_result_folder(
    batch_id: int,
    _user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    """用资源管理器打开本批结果文件夹（不下载、不压缩）。"""
    import subprocess
    import sys

    info = get_result_folder(batch_id=batch_id, _user=_user, db=db)
    path = info["path"]
    try:
        if sys.platform == "win32":
            # explorer <dir> 打开文件夹
            subprocess.Popen(["explorer", path], shell=False)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"打开文件夹失败：{exc}") from exc
    return {"ok": True, "path": path, "opened": True}


@router.post("/{batch_id}/export-result-folder")
def export_result_folder(
    batch_id: int,
    payload: dict = Body(None),
    _user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    """把结果文件夹整份复制到桌面或指定目录（普通文件夹，不打 zip）。"""
    import json
    import shutil
    from datetime import datetime
    from zoneinfo import ZoneInfo

    payload = payload or {}
    info = get_result_folder(batch_id=batch_id, _user=_user, db=db)
    src = Path(info["path"])
    settings = get_settings()

    target_hint = str(payload.get("target") or payload.get("dest") or "").strip()
    if not target_hint or target_hint in {"desktop", "桌面"}:
        # 默认桌面
        home = Path.home()
        desktop = home / "Desktop"
        if not desktop.is_dir():
            desktop = home / "桌面"
        if not desktop.is_dir():
            # Windows OneDrive Desktop
            od = home / "OneDrive" / "Desktop"
            desktop = od if od.is_dir() else home
        dest_root = desktop
    elif target_hint in {"export_dir", "设置导出目录"}:
        cfg_path = settings.data_dir / "app_settings.json"
        export_dir = settings.data_dir / "outputs"
        if cfg_path.is_file():
            try:
                data = json.loads(cfg_path.read_text(encoding="utf-8"))
                if data.get("export_dir"):
                    export_dir = Path(str(data["export_dir"]))
            except Exception:
                pass
        dest_root = export_dir
    else:
        dest_root = Path(target_hint)

    stamp = datetime.now(ZoneInfo("Asia/Shanghai")).strftime("%Y%m%d")
    label = "海外仓加工结果" if info.get("project") == "project2" else "店小秘加工结果"
    dest = dest_root / f"{label}-批次{batch_id}-{stamp}"
    # 若已存在则加序号
    base = dest
    n = 2
    while dest.exists():
        dest = Path(f"{base}-{n}")
        n += 1
    try:
        dest_root.mkdir(parents=True, exist_ok=True)
        shutil.copytree(src, dest, ignore=shutil.ignore_patterns("_internal", "_*", "*.pyc"))
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"复制结果文件夹失败：{exc}") from exc

    # 可选：复制后打开
    open_after = bool(payload.get("open", True))
    if open_after:
        import subprocess
        import sys

        try:
            if sys.platform == "win32":
                subprocess.Popen(["explorer", str(dest)])
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(dest)])
            else:
                subprocess.Popen(["xdg-open", str(dest)])
        except Exception:
            pass

    return {"ok": True, "path": str(dest), "source": str(src), "opened": open_after}


def _url_quote(value: str) -> str:
    from urllib.parse import quote

    return quote(value)


@router.get("/{batch_id}/download")
def download_batch_file(
    batch_id: int,
    path: str,
    _user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    batch = _load_batch(db, batch_id)
    if batch is None:
        raise HTTPException(status_code=404, detail="批次不存在")
    allowed = _download_paths(batch.engine_summary)
    if path not in allowed:
        raise HTTPException(status_code=404, detail="该批次没有这个导出文件")
    full = _safe_data_file(path)
    return FileResponse(full, filename=Path(path).name)


def _download_paths(summary: dict | None) -> set[str]:
    found: set[str] = set()

    def walk(value):
        if isinstance(value, dict):
            for item in value.values():
                walk(item)
        elif isinstance(value, list):
            for item in value:
                walk(item)
        elif isinstance(value, str) and value.endswith((".xlsx", ".xlsm", ".xls", ".csv", ".txt")):
            found.add(value)

    walk(summary or {})
    return found


def _safe_data_file(relative: str) -> Path:
    if not relative or relative.startswith(("/", "\\")) or ".." in Path(relative).parts:
        raise HTTPException(status_code=400, detail="非法路径")
    root = get_settings().data_dir.resolve()
    full = (root / relative).resolve()
    if not str(full).startswith(str(root)):
        raise HTTPException(status_code=400, detail="非法路径")
    if not full.is_file():
        raise HTTPException(status_code=404, detail="文件不存在")
    return full
