from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy import func, select
from sqlalchemy.orm import Session, selectinload

from app.database import get_db
from app.deps import get_current_user, require_admin, require_roles
from app.models import EngineRun, ExportJob, ExportProfile, OpsGroup, Order, OrderPackage, Shop, SkuMapping, SkuSuffix, User
from app.schemas import ExportProfileIn, OpsGroupIn, ShopIn, SkuMapIn
from app.serialize import order_summary
from app.services.project1_runner import project1_status
from app.services.project2_stub import project2_status

router = APIRouter(tags=["catalog"])


@router.get("/api/dashboard")
def dashboard(
    _user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    from app.models import ExceptionReview, ImportBatch, SplitAmountReview

    recent_stmt = (
        select(Order)
        .options(
            selectinload(Order.lines),
            selectinload(Order.packages).selectinload(OrderPackage.trackings),
        )
        .order_by(Order.id.desc())
        .limit(5)
    )
    recent = db.scalars(recent_stmt).all()
    return {
        "order_count": int(db.scalar(select(func.count()).select_from(Order)) or 0),
        "pending_orders": int(
            db.scalar(select(func.count()).select_from(Order).where(Order.status == "pending_review")) or 0
        ),
        "open_exceptions": int(
            db.scalar(
                select(func.count()).select_from(ExceptionReview).where(ExceptionReview.status == "open")
            )
            or 0
        ),
        "pending_split_reviews": int(
            db.scalar(
                select(func.count()).select_from(SplitAmountReview).where(SplitAmountReview.status == "pending")
            )
            or 0
        ),
        "batch_count": int(db.scalar(select(func.count()).select_from(ImportBatch)) or 0),
        "recent_orders": [order_summary(row) for row in recent],
        "project1": project1_status(),
        "project2": project2_status(),
    }


@router.get("/api/sku-mappings")
def list_sku_mappings(
    _user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    rows = db.scalars(select(SkuMapping).order_by(SkuMapping.id.desc())).all()
    return {
        "items": [
            {
                "id": row.id,
                "project": row.project,
                "source_sku": row.source_sku,
                "target_sku": row.target_sku,
                "note": row.note,
                "enabled": row.enabled,
            }
            for row in rows
        ]
    }


@router.get("/api/sku-suffixes")
def list_sku_suffixes(
    _user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    rows = db.scalars(select(SkuSuffix).where(SkuSuffix.project == "project2").order_by(SkuSuffix.suffix)).all()
    return {"items": [{"id": row.id, "suffix": row.suffix, "enabled": row.enabled} for row in rows]}


@router.post("/api/sku-suffixes", status_code=201)
def create_sku_suffix(
    body: dict,
    _user: User = Depends(require_roles("ops")),
    db: Session = Depends(get_db),
) -> dict:
    suffix = str(body.get("suffix") or "").strip()
    if not suffix:
        raise HTTPException(status_code=400, detail="请填写后缀")
    row = SkuSuffix(project="project2", suffix=suffix, enabled=True)
    db.add(row)
    db.commit()
    db.refresh(row)
    return {"id": row.id, "suffix": row.suffix}


@router.delete("/api/sku-suffixes/{suffix_id}", status_code=204)
def delete_sku_suffix(
    suffix_id: int,
    _user: User = Depends(require_roles("ops")),
    db: Session = Depends(get_db),
) -> None:
    row = db.get(SkuSuffix, suffix_id)
    if row is None:
        raise HTTPException(status_code=404, detail="后缀不存在")
    db.delete(row)
    db.commit()


@router.post("/api/sku-mappings", status_code=201)
def create_sku_mapping(
    body: SkuMapIn,
    _user: User = Depends(require_roles("ops")),
    db: Session = Depends(get_db),
) -> dict:
    if body.project == "project1":
        raise HTTPException(status_code=400, detail="项目1的 SKU 规则在 order_convert 里，不在这里改")
    row = SkuMapping(
        project=body.project,
        source_sku=body.source_sku.strip(),
        target_sku=body.target_sku.strip(),
        note=body.note,
        enabled=True,
    )
    db.add(row)
    db.commit()
    db.refresh(row)
    return {"id": row.id, "source_sku": row.source_sku, "target_sku": row.target_sku}


@router.delete("/api/sku-mappings/{mapping_id}", status_code=204)
def delete_sku_mapping(
    mapping_id: int,
    _user: User = Depends(require_roles("ops")),
    db: Session = Depends(get_db),
) -> None:
    row = db.get(SkuMapping, mapping_id)
    if row is None:
        raise HTTPException(status_code=404, detail="映射不存在")
    db.delete(row)
    db.commit()


@router.get("/api/shops")
def list_shops(
    _user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    rows = db.scalars(select(Shop).options(selectinload(Shop.ops_group)).order_by(Shop.id.desc())).all()
    return {
        "items": [
            {
                "id": row.id,
                "project": row.project,
                "shop_account": row.shop_account,
                "platform": row.platform,
                "ops_group": row.ops_group.name if row.ops_group is not None else None,
                "ops_group_id": row.ops_group_id,
                "already_in_guanjia": row.already_in_guanjia,
                "destination": row.destination,
                "routing_note": row.routing_note,
            }
            for row in rows
        ]
    }


@router.post("/api/shops", status_code=201)
def create_shop(
    body: ShopIn,
    _user: User = Depends(require_roles("ops")),
    db: Session = Depends(get_db),
) -> dict:
    if body.ops_group_id is not None and db.get(OpsGroup, body.ops_group_id) is None:
        raise HTTPException(status_code=404, detail="运营分组不存在")
    row = Shop(
        project=body.project,
        shop_account=body.shop_account.strip(),
        platform=body.platform,
        ops_group_id=body.ops_group_id,
        already_in_guanjia=body.already_in_guanjia,
        destination=body.destination,
        routing_note=body.routing_note,
    )
    db.add(row)
    db.commit()
    db.refresh(row)
    return {"id": row.id, "shop_account": row.shop_account, "destination": row.destination}


@router.delete("/api/shops/{shop_id}", status_code=204)
def delete_shop(
    shop_id: int,
    _user: User = Depends(require_roles("ops")),
    db: Session = Depends(get_db),
) -> None:
    row = db.get(Shop, shop_id)
    if row is None:
        raise HTTPException(status_code=404, detail="店铺不存在")
    db.delete(row)
    db.commit()


@router.get("/api/ops-groups")
def list_ops_groups(
    _user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    rows = db.scalars(select(OpsGroup).order_by(OpsGroup.id.desc())).all()
    return {
        "items": [
            {"id": row.id, "name": row.name, "project": row.project, "note": row.note}
            for row in rows
        ]
    }


@router.post("/api/ops-groups", status_code=201)
def create_ops_group(
    body: OpsGroupIn,
    _user: User = Depends(require_roles("ops")),
    db: Session = Depends(get_db),
) -> dict:
    row = OpsGroup(name=body.name.strip(), project=body.project, note=body.note)
    db.add(row)
    db.commit()
    db.refresh(row)
    return {"id": row.id, "name": row.name}


@router.get("/api/exports")
def list_exports(
    _user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    rows = db.scalars(select(ExportJob).order_by(ExportJob.id.desc())).all()
    return {
        "items": [
            {
                "id": row.id,
                "project": row.project,
                "batch_id": row.batch_id,
                "kind": row.kind,
                "status": row.status,
                "note": row.note,
            }
            for row in rows
        ]
    }


@router.get("/api/export-profiles")
def list_export_profiles(
    _user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    rows = db.scalars(select(ExportProfile).order_by(ExportProfile.id.desc())).all()
    return {
        "items": [
            {"id": row.id, "project": row.project, "kind": row.kind, "column_map": row.column_map or [], "note": row.note}
            for row in rows
        ]
    }


@router.post("/api/export-profiles", status_code=201)
def create_export_profile(
    body: ExportProfileIn,
    _user: User = Depends(require_roles("ops")),
    db: Session = Depends(get_db),
) -> dict:
    cleaned = []
    for item in body.column_map:
        source = str(item.get("source") or "").strip()
        target = str(item.get("target") or "").strip()
        if source and target:
            cleaned.append({"source": source, "target": target})
    row = ExportProfile(project=body.project, kind=body.kind, column_map=cleaned, note=body.note)
    db.add(row)
    db.commit()
    db.refresh(row)
    return {"id": row.id, "kind": row.kind, "column_map": row.column_map}


@router.get("/api/engine-runs")
def list_engine_runs(
    _user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    rows = db.scalars(select(EngineRun).order_by(EngineRun.id.desc()).limit(50)).all()
    return {
        "items": [
            {
                "id": row.id,
                "project": row.project,
                "batch_id": row.batch_id,
                "status": row.status,
                "created_at": row.created_at.isoformat(timespec="seconds"),
            }
            for row in rows
        ]
    }


@router.get("/api/users")
def list_users(
    _admin: User = Depends(require_admin),
    db: Session = Depends(get_db),
) -> dict:
    rows = db.scalars(select(User).order_by(User.id)).all()
    return {
        "items": [
            {
                "id": row.id,
                "username": row.username,
                "role": row.role,
                "display_name": row.display_name,
                "is_active": row.is_active,
            }
            for row in rows
        ]
    }
