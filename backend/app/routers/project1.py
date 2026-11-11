"""项目1（店小秘）专用接口：上传 → 加工 → 取结果。

与打单2约定对齐：
  POST /api/project1/upload
  POST /api/project1/run
  GET  /api/project1/result/{batch_id}
"""
from __future__ import annotations

import hashlib
from pathlib import Path

from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field
from sqlalchemy import select
from sqlalchemy.orm import Session, selectinload

from app.config import get_settings
from app.database import get_db
from app.deps import get_current_user, require_roles
from app.models import BATCH_FILE_KINDS, BatchFile, EngineRun, ExceptionReview, ImportBatch, User, utcnow
from app.services.project1_runner import aggregate_status, build_downloads, convert_files, project1_status

router = APIRouter(prefix="/api/project1", tags=["project1"])

_MAX_BYTES = 30 * 1024 * 1024


class RunBody(BaseModel):
    batch_id: int = Field(..., description="upload 返回的批次 id")


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


def _download_paths(summary: dict | None) -> set[str]:
    found: set[str] = set()

    def walk(value):
        if isinstance(value, dict):
            for item in value.values():
                walk(item)
        elif isinstance(value, list):
            for item in value:
                walk(item)
        elif isinstance(value, str) and value.endswith((".xlsx", ".xlsm", ".xls", ".csv")):
            found.add(value)

    walk(summary or {})
    return found


@router.get("/status")
def status(_user: User = Depends(get_current_user)) -> dict:
    """引擎就绪状态（与 /api/engines/project1/status 同源）。"""
    return project1_status()


@router.post("/upload", status_code=201)
async def upload(
    files: list[UploadFile] = File(...),
    note: str = Form(""),
    file_kind: str = Form("raw_order"),
    user: User = Depends(require_roles("ops")),
    db: Session = Depends(get_db),
) -> dict:
    """上传原始 Order / 店小秘加工用 Excel，返回 batch_id（尚未加工）。"""
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
        payloads.append((_safe_name(upload.filename, index, used), data))

    batch = ImportBatch(
        project="project1",
        status="created",
        note=(note.strip() or None),
        created_by_id=user.id,
        created_at=utcnow(),
    )
    db.add(batch)
    db.flush()

    settings = get_settings()
    upload_dir = settings.data_dir / "uploads" / str(batch.id)
    upload_dir.mkdir(parents=True, exist_ok=True)
    stored: list[dict] = []
    for name, data in payloads:
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
                kind=file_kind,
            )
        )
        stored.append(
            {
                "original_name": name,
                "size_bytes": len(data),
                "kind": file_kind,
            }
        )
    db.commit()
    return {
        "batch_id": batch.id,
        "project": "project1",
        "status": "created",
        "files": stored,
        "message": "上传成功，请调用 /api/project1/run 开始加工",
    }


@router.post("/run")
def run(
    body: RunBody,
    user: User = Depends(require_roles("ops")),
    db: Session = Depends(get_db),
) -> dict:
    """按 batch_id 触发店小秘三表加工。"""
    batch = _load_batch(db, body.batch_id)
    if batch is None:
        raise HTTPException(status_code=404, detail="批次不存在")
    if batch.project != "project1":
        raise HTTPException(status_code=400, detail="仅项目1批次可加工")
    if not batch.files:
        raise HTTPException(status_code=400, detail="批次没有可加工的文件")

    settings = get_settings()
    paths: list[Path] = []
    for item in batch.files:
        full = settings.data_dir / item.stored_path
        if not full.is_file():
            raise HTTPException(status_code=400, detail=f"上传文件丢失：{item.original_name}")
        paths.append(full)

    out_dir = settings.data_dir / "outputs" / str(batch.id)
    summary = convert_files(paths, out_dir)
    status = aggregate_status(summary)
    downloads = build_downloads(summary)
    summary["downloads"] = downloads

    for item in summary.get("files", []):
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

    batch.status = status
    batch.engine_status = status
    batch.engine_summary = summary
    db.add(
        EngineRun(
            project="project1",
            batch_id=batch.id,
            status=status,
            summary=summary,
            created_at=utcnow(),
        )
    )
    db.commit()
    stored = _load_batch(db, batch.id)
    return {
        "batch_id": batch.id,
        "status": status,
        "engine_status": status,
        "files": (stored.engine_summary or {}).get("files", []) if stored else summary.get("files", []),
        "downloads": downloads,
        "message": "加工完成",
    }


@router.get("/result/{batch_id}")
def result(
    batch_id: int,
    _user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    """返回三表下载路径 + 预览摘要。"""
    batch = _load_batch(db, batch_id)
    if batch is None:
        raise HTTPException(status_code=404, detail="批次不存在")
    if batch.project != "project1":
        raise HTTPException(status_code=400, detail="仅项目1批次")

    summary = batch.engine_summary or {}
    files_preview = []
    for item in summary.get("files", []):
        files_preview.append(
            {
                "file": item.get("file"),
                "export_status": item.get("export_status"),
                "stats": item.get("stats") or {},
                "error_count": item.get("error_count", 0),
                "warning_count": item.get("warning_count", 0),
                "errors": item.get("errors") or [],
                "warnings": item.get("warnings") or [],
                "anomalies": item.get("anomalies") or [],
                "outputs": item.get("outputs") or {},
                "message": item.get("message"),
            }
        )
    downloads = summary.get("downloads") or build_downloads(summary)
    return {
        "batch_id": batch.id,
        "status": batch.status,
        "engine_status": batch.engine_status,
        "files": files_preview,
        "downloads": downloads,
        "ready_for_download": bool(downloads),
    }


@router.get("/download")
def download(
    batch_id: int,
    path: str,
    _user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """按 result 返回的相对路径下载导出文件。"""
    batch = _load_batch(db, batch_id)
    if batch is None:
        raise HTTPException(status_code=404, detail="批次不存在")
    if batch.project != "project1":
        raise HTTPException(status_code=400, detail="仅项目1批次")
    allowed = _download_paths(batch.engine_summary)
    if path not in allowed:
        raise HTTPException(status_code=404, detail="该批次没有这个导出文件")
    full = _safe_data_file(path)
    return FileResponse(full, filename=Path(path).name)
