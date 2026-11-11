# -*- coding: utf-8 -*-
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml

from .resources_path import builtin_rules_path


def _project_rules() -> Path:
    return Path(__file__).resolve().parents[2] / "config" / "rules.yaml"


@dataclass
class Rules:
    raw: dict[str, Any]

    @property
    def china_marker(self) -> str:
        return self.raw["china_direct"]["pick_note_marker"]

    @property
    def china_sku(self) -> str:
        return self.raw["china_direct"]["sku"]

    @property
    def china_order_suffix(self) -> str:
        return self.raw["china_direct"]["order_no_suffix"]

    @property
    def strip_suffixes(self) -> set[str]:
        return set(self.raw["sku_suffixes"]["strip"])

    @property
    def japan_suffix(self) -> str:
        return self.raw["sku_suffixes"]["japan"]["suffix"]

    @property
    def japan_note(self) -> str:
        return self.raw["sku_suffixes"]["japan"]["note_contains"]

    @property
    def sku_aliases(self) -> dict[str, str]:
        """精确 SKU 别名映射（去后缀等之后应用）；缺省为空。"""
        raw = self.raw.get("sku_aliases") or {}
        return {str(k): str(v) for k, v in raw.items()}


    @property
    def sku_prefix_rules(self) -> list[dict[str, str]]:
        """SKU 前缀剥离规则：when_starts_with 命中则去掉 strip_prefix；缺省为空。"""
        raw = self.raw.get("sku_prefix_rules") or []
        out: list[dict[str, str]] = []
        for item in raw:
            if not isinstance(item, dict):
                continue
            when = str(item.get("when_starts_with") or "").strip()
            strip = str(item.get("strip_prefix") or "").strip()
            if when and strip:
                out.append({"when_starts_with": when, "strip_prefix": strip})
        return out

    @property
    def logistics_append_skus(self) -> set[str]:
        """基础 SKU 命中时追加到店小秘物流 / 店管家02客服备注（无空格）。"""
        return set(str(x) for x in (self.raw.get("sku_append_to_logistics") or []))


    @property
    def sku_from_note_columns(self) -> list[str]:
        """备注列货号替换：按优先级扫描；缺省 拣货备注→客服备注→订单备注。"""
        cols = self.raw.get("sku_from_note_columns")
        if cols:
            return [str(c) for c in cols]
        return ["拣货备注", "客服备注", "订单备注"]

    @property
    def currency_append_to_logistics(self) -> bool:
        cur = self.raw.get("currency") or {}
        return bool(cur.get("append_to_logistics", False))

    @property
    def currency_columns(self) -> list[str]:
        cur = self.raw.get("currency") or {}
        cols = cur.get("columns") or ["币种缩写", "币种"]
        return [str(c) for c in cols]

    @property
    def currency_skip(self) -> set[str]:
        """跳过追加的币种集合（统一为大写；中文原样保留）。"""
        cur = self.raw.get("currency") or {}
        skip = cur.get("skip") or ["EUR", "欧元"]
        out: set[str] = set()
        for x in skip:
            s = str(x).strip()
            if not s:
                continue
            out.add(s)
            out.add(s.upper())
        return out

    @property
    def amazon_keyword(self) -> str:
        return self.raw["shops"]["amazon_keyword"]

    @property
    def d01_known(self) -> list[str]:
        return list(self.raw["shops"]["d01_known"])

    @property
    def d02_fixed(self) -> dict[str, str]:
        return dict(self.raw["d02_fixed"])

    @property
    def d02_red_columns(self) -> list[str]:
        return list(self.raw["d02_red_columns"])

    @property
    def phone_auto_fallback(self) -> bool:
        return bool(self.raw["phone"]["auto_fallback_to_mobile"])

    @property
    def phone_warn_dest(self) -> str:
        return str(self.raw["phone"]["warn_only_for_dest"])


def load_rules(path: Path | None = None) -> Rules:
    """优先级：显式路径 → 项目 config/rules.yaml → 包内 resources/rules.yaml。"""
    candidates: list[Path] = []
    if path:
        candidates.append(Path(path))
    candidates.append(_project_rules())
    candidates.append(builtin_rules_path())
    for p in candidates:
        if p.exists():
            with open(p, encoding="utf-8") as f:
                return Rules(raw=yaml.safe_load(f))
    raise FileNotFoundError("未找到 rules.yaml（项目 config/ 或包内 resources/）")
