"""设置页读写：SKU 后缀/映射走数据库；仓库物流映射写回 routing.yaml；导出目录存本地 JSON。"""
from __future__ import annotations

import json
from pathlib import Path

import yaml
from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel, Field
from sqlalchemy import select
from sqlalchemy.orm import Session

from app.config import get_settings
from app.database import get_db
from app.deps import get_current_user, require_roles
from app.models import SkuMapping, SkuSuffix, User

router = APIRouter(prefix="/api/settings", tags=["settings"])


class SuffixIn(BaseModel):
    suffix: str
    note: str | None = None


class SuffixesPut(BaseModel):
    items: list[SuffixIn] = Field(default_factory=list)


class SkuMapItem(BaseModel):
    source_sku: str
    target_sku: str
    note: str | None = None
    enabled: bool = True


class SkuMapsPut(BaseModel):
    items: list[SkuMapItem] = Field(default_factory=list)


class WarehouseItem(BaseModel):
    contains: str
    warehouse: str
    logistics: str


class WarehouseMapPut(BaseModel):
    items: list[WarehouseItem] = Field(default_factory=list)


class ExportDirIn(BaseModel):
    path: str


def _routing_path() -> Path:
    return get_settings().project2_routing_path


def _load_routing_fresh() -> dict:
    path = _routing_path()
    if not path.is_file():
        raise HTTPException(status_code=500, detail=f"找不到路由配置：{path}")
    with open(path, encoding="utf-8") as handle:
        loaded = yaml.safe_load(handle) or {}
    if not isinstance(loaded, dict):
        raise HTTPException(status_code=500, detail="routing.yaml 格式错误")
    return loaded


def _save_routing(config: dict) -> None:
    """整文件重写。先备份到 routing.yaml.bak；PyYAML 会丢掉注释，但键值顺序尽量保持。"""
    path = _routing_path()
    bak = path.with_suffix(path.suffix + ".bak")
    if path.is_file() and not bak.exists():
        # 只在尚无 bak 时备份，避免连续保存把「带注释原件」冲掉
        bak.write_bytes(path.read_bytes())
    elif path.is_file():
        # 另存时间戳备份，保留最近一次可回滚
        stamp = path.with_suffix(path.suffix + ".prev")
        stamp.write_bytes(path.read_bytes())
    text = yaml.safe_dump(
        config,
        allow_unicode=True,
        default_flow_style=False,
        sort_keys=False,
        width=120,
    )
    path.write_text(text, encoding="utf-8")
    try:
        from app.services.project2_stub import _load_routing

        _load_routing.cache_clear()
    except Exception:
        pass


def _app_settings_path() -> Path:
    return get_settings().data_dir / "app_settings.json"


def _read_app_settings() -> dict:
    path = _app_settings_path()
    if not path.is_file():
        return {"export_dir": str(get_settings().data_dir / "outputs")}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {"export_dir": str(get_settings().data_dir / "outputs")}
    if not isinstance(data, dict):
        return {"export_dir": str(get_settings().data_dir / "outputs")}
    return data


def _write_app_settings(data: dict) -> None:
    path = _app_settings_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


# ----- SKU 后缀白名单（数据库；剥离逻辑仍一律剥「-」后缀，名单仅展示/管理） -----


@router.get("/sku-suffixes")
def get_sku_suffixes(
    _user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    rows = db.scalars(select(SkuSuffix).where(SkuSuffix.project == "project2").order_by(SkuSuffix.suffix)).all()
    return {
        "items": [
            {"id": row.id, "suffix": row.suffix, "enabled": row.enabled, "note": "已验证" if row.enabled else ""}
            for row in rows
        ],
        "note": "加工时凡带「-」的 SKU 一律剥掉最后一个「-」及后面内容；本表仅管理已知后缀展示，不作为是否剥离的门槛。",
    }


@router.put("/sku-suffixes")
def put_sku_suffixes(
    body: SuffixesPut,
    _user: User = Depends(require_roles("ops")),
    db: Session = Depends(get_db),
) -> dict:
    """整表替换：传入 {items:[{suffix}]}。"""
    cleaned: list[str] = []
    for item in body.items:
        s = item.suffix.strip()
        if s and s not in cleaned:
            cleaned.append(s)
    existing = db.scalars(select(SkuSuffix).where(SkuSuffix.project == "project2")).all()
    for row in existing:
        db.delete(row)
    db.flush()
    for suffix in cleaned:
        db.add(SkuSuffix(project="project2", suffix=suffix, enabled=True))
    db.commit()
    rows = db.scalars(select(SkuSuffix).where(SkuSuffix.project == "project2").order_by(SkuSuffix.suffix)).all()
    return {
        "items": [
            {"id": row.id, "suffix": row.suffix, "enabled": row.enabled, "note": "已验证" if row.enabled else ""}
            for row in rows
        ],
        "note": "加工时凡带「-」的 SKU 一律剥掉最后一个「-」及后面内容；本表仅管理已知后缀展示，不作为是否剥离的门槛。",
    }


@router.post("/sku-suffixes", status_code=201)
def add_sku_suffix(
    body: SuffixIn,
    _user: User = Depends(require_roles("ops")),
    db: Session = Depends(get_db),
) -> dict:
    suffix = body.suffix.strip()
    if not suffix:
        raise HTTPException(status_code=400, detail="请填写后缀")
    row = SkuSuffix(project="project2", suffix=suffix, enabled=True)
    db.add(row)
    db.commit()
    db.refresh(row)
    return {"id": row.id, "suffix": row.suffix, "enabled": row.enabled}


@router.delete("/sku-suffixes/{suffix_id}", status_code=204)
def delete_sku_suffix_setting(
    suffix_id: int,
    _user: User = Depends(require_roles("ops")),
    db: Session = Depends(get_db),
) -> None:
    row = db.get(SkuSuffix, suffix_id)
    if row is None:
        raise HTTPException(status_code=404, detail="后缀不存在")
    db.delete(row)
    db.commit()


# ----- 特殊 SKU 映射：合并 yaml + 数据库，PUT 写回 yaml.sku_map 并同步 DB -----


@router.get("/sku-maps")
def get_sku_maps(
    _user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    cfg = _load_routing_fresh()
    yaml_map = cfg.get("sku_map") or {}
    items = [
        {"id": f"yaml:{source}", "source_sku": str(source), "target_sku": str(target), "source": "yaml", "enabled": True}
        for source, target in yaml_map.items()
    ]
    for row in db.scalars(select(SkuMapping).where(SkuMapping.project == "project2").order_by(SkuMapping.id)).all():
        items.append(
            {
                "id": row.id,
                "source_sku": row.source_sku,
                "target_sku": row.target_sku,
                "note": row.note,
                "enabled": row.enabled,
                "source": "db",
            }
        )
    return {"items": items}


@router.put("/sku-maps")
def put_sku_maps(
    body: SkuMapsPut,
    _user: User = Depends(require_roles("ops")),
    db: Session = Depends(get_db),
) -> dict:
    """写回 routing.yaml 的 sku_map（整表替换 yaml 段）；不删数据库里的额外映射。"""
    mapping: dict[str, str] = {}
    for item in body.items:
        src = item.source_sku.strip()
        tgt = item.target_sku.strip()
        if src and tgt:
            mapping[src] = tgt
    cfg = _load_routing_fresh()
    cfg["sku_map"] = mapping
    _save_routing(cfg)
    rows_out = []
    cfg2 = _load_routing_fresh()
    for source, target in (cfg2.get("sku_map") or {}).items():
        rows_out.append({"id": f"yaml:{source}", "source_sku": str(source), "target_sku": str(target), "source": "yaml", "enabled": True})
    for row in db.scalars(select(SkuMapping).where(SkuMapping.project == "project2").order_by(SkuMapping.id)).all():
        rows_out.append({"id": row.id, "source_sku": row.source_sku, "target_sku": row.target_sku, "note": row.note, "enabled": row.enabled, "source": "db"})
    return {"items": rows_out}


@router.post("/sku-maps", status_code=201)
def add_sku_map(
    body: SkuMapItem,
    _user: User = Depends(require_roles("ops")),
    db: Session = Depends(get_db),
) -> dict:
    src = body.source_sku.strip()
    tgt = body.target_sku.strip()
    if not src or not tgt:
        raise HTTPException(status_code=400, detail="请填写源 SKU 与目标 SKU")
    cfg = _load_routing_fresh()
    sku_map = dict(cfg.get("sku_map") or {})
    sku_map[src] = tgt
    cfg["sku_map"] = sku_map
    _save_routing(cfg)
    return {"source_sku": src, "target_sku": tgt, "source": "yaml"}


@router.delete("/sku-maps/{source_sku}")
def delete_sku_map(
    source_sku: str,
    _user: User = Depends(require_roles("ops")),
    db: Session = Depends(get_db),
) -> dict:
    """优先从 yaml 删除；若是数字 id 则删数据库行。"""
    if source_sku.isdigit():
        row = db.get(SkuMapping, int(source_sku))
        if row is None:
            raise HTTPException(status_code=404, detail="映射不存在")
        db.delete(row)
        db.commit()
        return {"deleted": source_sku, "source": "db"}
    key = source_sku
    if key.startswith("yaml:"):
        key = key[5:]
    cfg = _load_routing_fresh()
    sku_map = dict(cfg.get("sku_map") or {})
    if key not in sku_map:
        # 大小写不敏感试一次
        found = next((k for k in sku_map if str(k).upper() == key.upper()), None)
        if found is None:
            raise HTTPException(status_code=404, detail="映射不存在")
        key = found
    del sku_map[key]
    cfg["sku_map"] = sku_map
    _save_routing(cfg)
    return {"deleted": key, "source": "yaml"}


# ----- 仓库/物流映射 -----


@router.get("/warehouse-map")
def get_warehouse_map(_user: User = Depends(get_current_user)) -> dict:
    cfg = _load_routing_fresh()
    rows = cfg.get("warehouse_logistics_map") or []
    items = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        items.append(
            {
                "contains": str(row.get("contains") or ""),
                "warehouse": str(row.get("warehouse") or ""),
                "logistics": str(row.get("logistics") or ""),
            }
        )
    return {
        "items": items,
        "note": "按 contains 顺序匹配，不区分大小写；更具体的规则请写在前面。",
    }


@router.put("/warehouse-map")
def put_warehouse_map(
    body: WarehouseMapPut,
    _user: User = Depends(require_roles("ops")),
) -> dict:
    cleaned = []
    for item in body.items:
        contains = item.contains.strip()
        if not contains:
            continue
        cleaned.append(
            {
                "contains": contains,
                "warehouse": item.warehouse.strip(),
                "logistics": item.logistics.strip(),
            }
        )
    cfg = _load_routing_fresh()
    cfg["warehouse_logistics_map"] = cleaned
    _save_routing(cfg)
    cfg2 = _load_routing_fresh()
    items = []
    for row in cfg2.get("warehouse_logistics_map") or []:
        if isinstance(row, dict):
            items.append({"contains": str(row.get("contains") or ""), "warehouse": str(row.get("warehouse") or ""), "logistics": str(row.get("logistics") or "")})
    return {"items": items, "note": "按 contains 顺序匹配，不区分大小写；更具体的规则请写在前面。"}


# ----- 导出目录（本地配置；引擎默认仍写 data/outputs） -----


@router.get("/export-dir")
def get_export_dir(_user: User = Depends(get_current_user)) -> dict:
    data = _read_app_settings()
    default = str(get_settings().data_dir / "outputs")
    return {
        "path": str(data.get("export_dir") or default),
        "default": default,
        "note": "保存到本地配置。当前批次结果仍默认写入 data/outputs/<批次号>；此路径供桌面端打开/归档参考。",
    }


@router.put("/export-dir")
def put_export_dir(
    body: ExportDirIn,
    _user: User = Depends(require_roles("ops")),
) -> dict:
    path = body.path.strip()
    if not path:
        raise HTTPException(status_code=400, detail="请填写导出目录")
    data = _read_app_settings()
    data["export_dir"] = path
    _write_app_settings(data)
    data2 = _read_app_settings()
    default = str(get_settings().data_dir / "outputs")
    return {"path": str(data2.get("export_dir") or default), "default": default, "note": "保存到本地配置。当前批次结果仍默认写入 data/outputs/<批次号>；此路径供桌面端打开/归档参考。"}
