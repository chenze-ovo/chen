"""Project2 占位。

不读取 Excel，不匹配店铺，不生成海外仓或店管家文件。
路由、空映射和样本缺口都来自 engines/project2/routing.yaml。
"""
from __future__ import annotations

from functools import lru_cache

import yaml

from app.config import get_settings


def _empty_status(message: str) -> dict:
    return {
        "project": "project2",
        "mode": "stub",
        "ready": False,
        "execute": False,
        "executed": False,
        "sku_map_entries": 0,
        "warehouse_logistics_entries": 0,
        "amount_split_rules": 0,
        "dian_01_field_map_entries": 0,
        "dian_02_field_map_entries": 0,
        "routing_policy": [],
        "processing_constraints": [],
        "outputs_later": [],
        "must_emit_later": [],
        "sample_gaps": [],
        "batch_file_kinds": [],
        "review_task_types": [],
        "china_direct_amounts": {},
        "split_sku_amounts": {},
        "d01_field_rules": {},
        "message": message,
    }


def _count(value) -> int:
    if isinstance(value, dict | list):
        return len(value)
    return 0


@lru_cache
def _load_routing(path_text: str) -> dict:
    path = path_text
    with open(path, encoding="utf-8") as handle:
        loaded = yaml.safe_load(handle) or {}
    if not isinstance(loaded, dict):
        raise ValueError("routing.yaml 必须是映射")
    return loaded


def project2_status() -> dict:
    settings = get_settings()
    path = settings.project2_routing_path
    try:
        cfg = _load_routing(str(path))
    except Exception as exc:
        return _empty_status(f"无法读取路由配置：{type(exc).__name__}: {exc}")

    policies = []
    for item in cfg.get("routing_policy") or []:
        if not isinstance(item, dict):
            continue
        policies.append(
            {
                "id": item.get("id"),
                "destination": item.get("destination"),
                "enabled": bool(item.get("enabled", False)),
                "note": item.get("note"),
                "when": item.get("when") or {},
                "row_marks": item.get("row_marks") or [],
                "action": item.get("action"),
            }
        )
    return {
        "project": "project2",
        "mode": "stub",
        "ready": True,
        "execute": bool(cfg.get("execute", False)),
        "executed": False,
        "config_path": str(path),
        "sku_map_entries": _count(cfg.get("sku_map")),
        "warehouse_logistics_entries": _count(cfg.get("warehouse_logistics_map")),
        "amount_split_rules": _count(cfg.get("amount_split_rules")),
        "dian_01_field_map_entries": _count(cfg.get("dian_01_field_map")),
        "dian_02_field_map_entries": _count(cfg.get("dian_02_field_map")),
        "routing_policy": policies,
        "processing_constraints": cfg.get("processing_constraints") or [],
        "outputs_later": cfg.get("outputs_later") or [],
        "must_emit_later": cfg.get("must_emit_later") or [],
        "sample_gaps": cfg.get("sample_gaps") or [],
        "batch_file_kinds": cfg.get("batch_file_kinds") or [],
        "review_task_types": cfg.get("review_task_types") or [],
        "china_direct_amounts": cfg.get("china_direct_amounts") or {},
        "split_sku_amounts": cfg.get("split_sku_amounts") or {},
        "d01_field_rules": cfg.get("d01_field_rules") or {},
        "row_mark_note": cfg.get("row_mark_note"),
        "message": "Phase 0 不执行 Project2。代码里没有 SKU、仓库物流、拆金额或店管家列映射。",
    }


def run_project2(file_names: list[str]) -> dict:
    # 故意没有“执行转换”的分支。execute 标志在本阶段被忽略。
    info = project2_status()
    info["executed"] = False
    info["execute_ignored"] = True
    info["files_received"] = list(file_names)
    info["reason"] = "文件已归档。Phase 0 不转换、不拆金额、不生成店管家或海外仓文件。"
    return info
