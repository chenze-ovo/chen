"""把项目2批次交给独立引擎，并把订单、审核和导出写进数据库。"""
from __future__ import annotations

import sys
from decimal import Decimal
from pathlib import Path

from sqlalchemy import select
from sqlalchemy.orm.attributes import flag_modified
from sqlalchemy.orm import Session

from app.config import get_settings
from app.models import (
    BatchFile,
    ImportBatch,
    ExceptionReview,
    ExportJob,
    ExportProfile,
    OpsGroup,
    Order,
    OrderLine,
    OrderPackage,
    PackageTracking,
    ReviewTask,
    Shop,
    SkuMapping,
    SkuSuffix,
    SplitAmountReview,
    SplitAmountReviewLine,
    EngineRun,
    utcnow,
)
from app.services.project2_stub import project2_status


def _engine():
    root = str(get_settings().erp_repo_root / "engines" / "project2")
    if root not in sys.path:
        sys.path.insert(0, root)
    from project2_engine import process_orders

    return process_orders


def run_project2_batch(
    db: Session,
    order_paths: list[Path],
    out_dir: Path,
    batch_id: int,
    *,
    copy_order_nos: list[str] | None = None,
    manual_amounts: dict | None = None,
) -> dict:
    status = project2_status()
    shops = _load_shops(db)
    config = _config_for_engine()
    yaml_maps = [
        {"source_sku": str(source), "target_sku": str(target), "enabled": True}
        for source, target in (config.get("sku_map") or {}).items()
    ]
    db_maps = [
        {"source_sku": row.source_sku, "target_sku": row.target_sku, "enabled": row.enabled}
        for row in db.scalars(select(SkuMapping).where(SkuMapping.project == "project2")).all()
    ]
    sku_maps = yaml_maps + db_maps
    profiles = db.scalars(select(ExportProfile).where(ExportProfile.project == "project2")).all()
    column_maps = {row.kind: list(row.column_map or []) for row in profiles}
    suffixes = [
        row.suffix
        for row in db.scalars(select(SkuSuffix).where(SkuSuffix.project == "project2", SkuSuffix.enabled.is_(True))).all()
    ]
    repo = get_settings().erp_repo_root
    p2_templates = repo / "engines" / "project2" / "templates"
    p1_templates = (
        repo
        / "engines"
        / "project1"
        / "vendor"
        / "order-convert"
        / "src"
        / "order_convert"
        / "templates"
    )

    def _tpl(name: str) -> str:
        local = p2_templates / name
        if local.is_file():
            return str(local)
        return str(p1_templates / name)

    same_day = _same_day_paths(db, batch_id)
    batch = db.get(ImportBatch, batch_id) if batch_id else None
    export_date = None
    if batch is not None and getattr(batch, "created_at", None) is not None:
        export_date = batch.created_at.strftime("%Y-%m-%d")
    # Prefer gold shell for 02 when present
    d02_tpl = _tpl("店管家02-金标准.xls")
    if not Path(d02_tpl).is_file():
        d02_tpl = _tpl("dian_02_blank.xls")
    gold = (config.get("identical_refs") or {}).get("dian_02")
    if gold and Path(gold).is_file():
        d02_tpl = gold
    summary = _engine()(
        order_paths,
        out_dir,
        config=config,
        same_day_files=same_day,
        shops=shops,
        sku_maps=sku_maps,
        sku_suffixes=suffixes,
        column_maps=column_maps,
        templates={
            "dian_01": _tpl("dian_01_blank.xls") if not Path(_tpl("店管家01样板.xls")).is_file() else _tpl("店管家01样板.xls"),
            "dian_02": d02_tpl,
        },
        export_date=export_date,
        identical_refs=config.get("identical_refs") or {},
        copy_order_nos=copy_order_nos,
        manual_amounts=manual_amounts,
    )
    data_dir = get_settings().data_dir.resolve()
    downloads = []
    for item in summary.get("downloads") or []:
        full = (out_dir / item["path"]).resolve()
        try:
            relative = full.relative_to(data_dir).as_posix()
        except ValueError:
            relative = str(full)
        labeled = {"label": item["label"], "path": relative, "kind": item.get("kind")}
        downloads.append(labeled)
        db.add(
            ExportJob(
                project="project2",
                batch_id=batch_id,
                kind=item["label"],
                status="ready",
                note=relative,
                created_at=utcnow(),
            )
        )
    summary["downloads"] = downloads
    from app.services.project2_packaging import ensure_notes_in_summary

    summary = ensure_notes_in_summary(out_dir, summary, data_dir, batch_id=batch_id)
    downloads = list(summary.get("downloads") or [])
    _persist_orders(db, summary, batch_id)
    for anomaly in summary.get("anomalies") or []:
        db.add(
            ExceptionReview(
                batch_id=batch_id,
                project="project2",
                code=str(anomaly.get("type") or "PROJECT2"),
                severity="warning",
                message=str(anomaly.get("message") or anomaly.get("type")),
                status="open",
                created_at=utcnow(),
            )
        )
        if anomaly.get("task_type"):
            db.add(
                ReviewTask(
                    task_type=anomaly["task_type"],
                    status="pending",
                    batch_id=batch_id,
                    message=str(anomaly.get("message") or ""),
                    created_at=utcnow(),
                )
            )
    summary["status_snapshot"] = {
        "sku_map_entries": status.get("sku_map_entries"),
        "dian_01_field_map_entries": status.get("dian_01_field_map_entries"),
        "force_zero": (status.get("china_direct_amounts") or {}).get("force_zero"),
    }
    return summary


def _config_for_engine() -> dict:
    from app.services.project2_stub import _load_routing

    path = get_settings().project2_routing_path
    config = _load_routing(str(path))
    p2_root = get_settings().erp_repo_root / "engines" / "project2"
    refs = dict(config.get("identical_refs") or {})
    def _abs(value):
        if not value:
            return value
        p = Path(str(value))
        if not p.is_absolute():
            p = p2_root / p
        return str(p) if p.is_file() else None
    if "oversea" in refs:
        refs["oversea"] = _abs(refs.get("oversea"))
    if "oversea_copy" in refs:
        refs["oversea_copy"] = _abs(refs.get("oversea_copy"))
    if "dian_02" in refs:
        resolved = _abs(refs.get("dian_02"))
        if not resolved:
            local = p2_root / "templates" / "店管家02-金标准.xls"
            resolved = str(local) if local.is_file() else None
        refs["dian_02"] = resolved
    ops = dict(refs.get("ops") or {})
    refs["ops"] = {k: _abs(v) for k, v in ops.items()}
    # drop missing so engine falls back cleanly
    refs = {k: v for k, v in refs.items() if v or k == "ops"}
    refs["ops"] = {k: v for k, v in (refs.get("ops") or {}).items() if v}
    config["identical_refs"] = refs
    return config


def _load_shops(db: Session) -> list[dict]:
    rows = db.scalars(select(Shop).where(Shop.project == "project2")).all()
    shops = []
    for row in rows:
        group_name = None
        if row.ops_group_id:
            group = db.get(OpsGroup, row.ops_group_id)
            group_name = group.name if group is not None else None
        shops.append(
            {
                "shop_account": row.shop_account,
                "platform": row.platform,
                "already_in_guanjia": row.already_in_guanjia,
                "destination": row.destination,
                "ops_group": group_name,
            }
        )
    return shops


def _same_day_paths(db: Session, batch_id: int | None = None) -> list[Path]:
    """本批「TikTok/金额原件」栏的文件。

    这些文件的行现在也参与转换（不再只当金额表），所以必须只取本批，
    否则历史批次的行会被带进本批结果。
    """
    root = get_settings().data_dir
    paths = []
    stmt = select(BatchFile).where(BatchFile.kind == "dianxiaomi_same_day")
    if batch_id:
        stmt = stmt.where(BatchFile.batch_id == batch_id)
    rows = db.scalars(stmt).all()
    for row in rows:
        full = root / row.stored_path
        if full.is_file():
            paths.append(full)
    return paths


def _persist_orders(db: Session, summary: dict, batch_id: int) -> None:
    for item in summary.get("orders") or []:
        order_no = item.get("order_no") or ""
        if not order_no:
            continue
        exists = db.scalar(
            select(Order.id).where(Order.project == "project2", Order.external_order_no == order_no)
        )
        if exists is not None:
            continue
        amount = _decimal(item.get("order_amount")) or Decimal("0")
        skus = [sku for sku in (item.get("skus") or []) if sku]
        multi = len(skus) > 1
        order = Order(
            batch_id=batch_id,
            project="project2",
            external_order_no=order_no,
            shop_account=item.get("shop_account") or None,
            platform=item.get("platform") or None,
            status="pending_review" if multi or item.get("destination") in {"review", "human_review"} else "imported",
            buyer_name=item.get("buyer_name") or None,
            country=item.get("country") or None,
            currency=item.get("currency") or "USD",
            order_amount=amount,
            dest_hint=item.get("destination"),
            row_mark=item.get("row_mark"),
            already_in_guanjia=bool(item.get("already_in_guanjia")),
            notes=_order_note(item, multi),
            raw_payload={"destination": item.get("destination"), "china_direct": item.get("china_direct")},
            created_at=utcnow(),
        )
        db.add(order)
        db.flush()
        if not skus:
            skus = ["未填SKU"]
        for index, sku in enumerate(skus, start=1):
            db.add(
                OrderLine(
                    order_id=order.id,
                    line_no=index,
                    sku=sku,
                    sku_name=None,
                    quantity=_decimal(item.get("quantity")) if (not multi and index == 1) else Decimal("0"),
                    unit_price=Decimal("0"),
                    line_amount=amount if not multi else Decimal("0"),
                    currency=order.currency,
                )
            )
        db.flush()
        tracks = item.get("tracking_nos") or []
        if tracks:
            package = OrderPackage(
                order_id=order.id,
                package_no=f"{order_no}-1",
                status="packed",
            )
            db.add(package)
            db.flush()
            for tracking in tracks:
                db.add(
                    PackageTracking(
                        order_id=order.id,
                        package_id=package.id,
                        tracking_no=tracking,
                        carrier=None,
                    )
                )
        if multi:
            review = SplitAmountReview(
                order_id=order.id,
                status="pending",
                reason="检测到多个 SKU。请按店小秘网页核对后手工填写各行金额，系统只核对合计。",
                created_at=utcnow(),
            )
            db.add(review)
            db.flush()
            lines = db.scalars(select(OrderLine).where(OrderLine.order_id == order.id).order_by(OrderLine.line_no)).all()
            for line in lines:
                db.add(
                    SplitAmountReviewLine(
                        review_id=review.id,
                        order_line_id=line.id,
                        sku=line.sku,
                        original_amount=Decimal("0"),
                        proposed_amount=Decimal("0"),
                        note="待手工填写",
                    )
                )


def _order_note(item: dict, multi: bool) -> str:
    parts = [f"去向 {item.get('destination')}"]
    if item.get("china_direct"):
        parts.append("中国直发已按当天店小秘金额对照" if item.get("amount_matched") else "中国直发未对上金额，未改成 0")
    if multi:
        parts.append("多 SKU 金额未拆，等待人工填写")
    if not item.get("ops_group"):
        parts.append("店铺未加入运营分组，所以没有分组表")
    return "。".join(parts)


def _decimal(value) -> Decimal | None:
    if value is None or value == "":
        return None
    try:
        return Decimal(str(value))
    except Exception:
        return None


def generate_oversea_copy_for_batch(
    db: Session,
    batch_id: int,
    copy_order_nos: list[str],
) -> dict:
    """从已有海外仓正本按勾选订单生成副本，并写回 batch.engine_summary.downloads。"""
    from app.services.project2_stub import project2_status  # noqa: F401
    import sys

    batch = db.get(ImportBatch, batch_id)
    if batch is None:
        raise ValueError("批次不存在")
    if batch.project != "project2":
        raise ValueError("仅项目2可生成海外仓副本")
    summary = dict(batch.engine_summary or {})
    downloads = list(summary.get("downloads") or [])
    formal = None
    for item in downloads:
        if item.get("kind") == "oversea" or (item.get("label") == "海外仓订单" and "副本" not in str(item.get("path") or "")):
            formal = item
            break
    if formal is None:
        raise ValueError("该批次还没有海外仓正本，无法生成副本")
    data_dir = get_settings().data_dir.resolve()
    formal_path = (data_dir / formal["path"]).resolve()
    if not formal_path.is_file():
        # path may already be relative under outputs
        alt = data_dir / "outputs" / str(batch_id) / Path(formal["path"]).name
        formal_path = alt if alt.is_file() else formal_path
    if not formal_path.is_file():
        raise ValueError(f"找不到海外仓正本文件：{formal.get('path')}")

    root = str(get_settings().erp_repo_root / "engines" / "project2")
    if root not in sys.path:
        sys.path.insert(0, root)
    from project2_engine import write_oversea_copy_from_formal

    config = _config_for_engine()
    naming = config.get("export_naming") or {}
    # derive date from formal filename if possible
    formal_name = formal_path.name
    date_full = formal_name.replace("海外仓订单.xlsx", "").strip() or "海外仓"
    copy_name = str(naming.get("oversea_copy") or "{date_full}海外仓订单 - 副本.xlsx").format(
        date_full=date_full, date_short=date_full
    )
    out_dir = formal_path.parent
    out_path = out_dir / copy_name
    result = write_oversea_copy_from_formal(formal_path, out_path, copy_order_nos, width_ref=formal_path)
    try:
        relative = out_path.resolve().relative_to(data_dir).as_posix()
    except ValueError:
        relative = str(out_path)
    # replace existing copy download entry
    downloads = [d for d in downloads if d.get("kind") != "oversea_copy" and "副本" not in str(d.get("label") or "")]
    downloads.append({
        "label": "海外仓订单 - 副本",
        "path": relative,
        "kind": "oversea_copy",
        "copy_order_nos": result.get("copy_order_nos"),
        "copy_row_count": result.get("copy_row_count"),
    })
    summary["downloads"] = downloads
    summary["oversea_copy"] = {
        "mode": "selected_orders",
        "copy_order_nos": result.get("copy_order_nos"),
        "copy_row_count": result.get("copy_row_count"),
    }
    from app.services.project2_packaging import ensure_notes_in_summary

    summary = ensure_notes_in_summary(out_dir, summary, data_dir, batch_id=batch_id)
    downloads = list(summary.get("downloads") or [])
    batch.engine_summary = summary
    flag_modified(batch, "engine_summary")
    db.add(
        ExportJob(
            project="project2",
            batch_id=batch_id,
            kind="海外仓订单 - 副本",
            status="ready",
            note=relative,
            created_at=utcnow(),
        )
    )
    db.commit()
    db.refresh(batch)
    return {
        "ok": True,
        "path": relative,
        "filename": out_path.name,
        "copy_order_nos": result.get("copy_order_nos"),
        "copy_row_count": result.get("copy_row_count"),
        "downloads": downloads,
    }


def _approved_manual_amounts(db: Session, batch_id: int) -> dict:
    """订单号|SKU -> 已审金额；用于店管家02合计回填后重导。"""
    out: dict[str, str] = {}
    orders = db.scalars(select(Order).where(Order.batch_id == batch_id, Order.project == "project2")).all()
    for order in orders:
        reviews = db.scalars(
            select(SplitAmountReview)
            .where(SplitAmountReview.order_id == order.id, SplitAmountReview.status == "approved")
            .options()
        ).all()
        for review in reviews:
            lines = db.scalars(
                select(SplitAmountReviewLine).where(SplitAmountReviewLine.review_id == review.id)
            ).all()
            for line in lines:
                key = f"{order.external_order_no}|{line.sku}"
                out[key] = str(line.proposed_amount)
                # also order-level for single-sku china_direct fill via review note amounts
                out[order.external_order_no] = str(line.proposed_amount) if len(lines) == 1 else out.get(order.external_order_no, "")
    return out


def reexport_project2_batch(db: Session, batch_id: int, copy_order_nos: list[str] | None = None) -> dict:
    """填完拆单金额后重跑海外仓批次，店管家02合计用已审金额。"""
    batch = db.get(ImportBatch, batch_id)
    if batch is None:
        raise ValueError("批次不存在")
    if batch.project != "project2":
        raise ValueError("仅海外仓批次可重导")
    settings = get_settings()
    files = db.scalars(select(BatchFile).where(BatchFile.batch_id == batch_id)).all()
    order_paths = []
    for row in files:
        if row.kind == "dianxiaomi_same_day":
            continue
        full = settings.data_dir / row.stored_path
        if full.is_file() and full.suffix.lower() in {".xlsx", ".xlsm"}:
            order_paths.append(full)
    if not order_paths:
        raise ValueError("批次没有可重导的原始 Order 文件")
    out_dir = settings.data_dir / "outputs" / str(batch_id)
    out_dir.mkdir(parents=True, exist_ok=True)
    manual = _approved_manual_amounts(db, batch_id)
    summary = run_project2_batch(
        db,
        order_paths,
        out_dir,
        batch_id,
        copy_order_nos=copy_order_nos,
        manual_amounts=manual,
    )
    stats = summary.get("stats") or {}
    if stats.get("human_review") or stats.get("review"):
        status = "needs_review"
    elif summary.get("executed"):
        status = "converted"
    else:
        status = "failed"
    batch.status = status
    batch.engine_status = status
    batch.engine_summary = summary
    flag_modified(batch, "engine_summary")
    db.add(
        EngineRun(
            project="project2",
            batch_id=batch_id,
            status=status,
            summary=summary,
            created_at=utcnow(),
        )
    )
    db.commit()
    db.refresh(batch)
    return summary
