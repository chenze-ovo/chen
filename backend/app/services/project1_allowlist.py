"""中国直发：用户在异常里确认后记下的后缀和店铺名单。

记在 data/project1_allowlist.yaml，转换时与 rules.yaml 合并，不改原规则文件。
"""
from __future__ import annotations

import copy
from pathlib import Path
from typing import Any, Callable

import yaml

from app.config import get_settings

ALLOWLIST_NAME = "project1_allowlist.yaml"
SHOP_KINDS = {"amazon", "other"}


def allowlist_path() -> Path:
    return get_settings().data_dir / ALLOWLIST_NAME


def _empty() -> dict[str, Any]:
    return {"sku_suffixes": [], "shops": []}


def _norm_suffix(raw: str) -> str:
    value = str(raw or "").strip()
    if value.startswith("-"):
        value = value[1:].strip()
    return value


def _norm_shop(raw: str) -> str:
    return str(raw or "").strip()


def load_allowlist() -> dict[str, Any]:
    path = allowlist_path()
    data = _empty()
    if not path.is_file():
        return data
    try:
        loaded = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    except yaml.YAMLError:
        return data
    if not isinstance(loaded, dict):
        return data
    suffixes: list[str] = []
    for item in loaded.get("sku_suffixes") or []:
        value = _norm_suffix(str(item))
        if value and value not in suffixes:
            suffixes.append(value)
    shops: list[dict[str, str]] = []
    seen: set[str] = set()
    for item in loaded.get("shops") or []:
        if not isinstance(item, dict):
            continue
        account = _norm_shop(item.get("account") or item.get("shop_account") or "")
        kind = str(item.get("kind") or "").strip().lower()
        if not account or kind not in SHOP_KINDS:
            continue
        key = account.lower()
        if key in seen:
            shops = [row for row in shops if row["account"].lower() != key]
        seen.add(key)
        shops.append({"account": account, "kind": kind})
    data["sku_suffixes"] = suffixes
    data["shops"] = shops
    return data


def save_allowlist(data: dict[str, Any]) -> dict[str, Any]:
    path = allowlist_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    cleaned = {
        "sku_suffixes": list(data.get("sku_suffixes") or []),
        "shops": list(data.get("shops") or []),
    }
    text = yaml.safe_dump(cleaned, allow_unicode=True, default_flow_style=False, sort_keys=False)
    path.write_text(text, encoding="utf-8")
    return cleaned


def add_suffix(suffix: str) -> dict[str, Any]:
    value = _norm_suffix(suffix)
    if not value:
        raise ValueError("请填写要记下的后缀")
    if not value.isalpha():
        raise ValueError("这个后缀不能这样记下，请改原表货号后再转一次")
    data = load_allowlist()
    existing = {item.lower() for item in data["sku_suffixes"]}
    if value.lower() not in existing:
        data["sku_suffixes"].append(value)
        save_allowlist(data)
    return data


def add_shop(shop_account: str, kind: str) -> dict[str, Any]:
    account = _norm_shop(shop_account)
    kind_value = str(kind or "").strip().lower()
    if not account:
        raise ValueError("请填写店铺账号")
    if kind_value not in SHOP_KINDS:
        raise ValueError("请选择按亚马逊处理，还是和其他店一样处理")
    data = load_allowlist()
    data["shops"] = [row for row in data["shops"] if row["account"].lower() != account.lower()]
    data["shops"].append({"account": account, "kind": kind_value})
    save_allowlist(data)
    return data


def merge_into_rules(rules: Any) -> Any:
    """把已记名单合并进 Rules，不写回 rules.yaml。"""
    data = load_allowlist()
    raw = copy.deepcopy(getattr(rules, "raw", None) or {})
    suffixes = list((raw.get("sku_suffixes") or {}).get("strip") or [])
    known = {str(item) for item in suffixes}
    for item in data["sku_suffixes"]:
        if item not in known:
            suffixes.append(item)
            known.add(item)
    raw.setdefault("sku_suffixes", {})["strip"] = suffixes

    shops = raw.setdefault("shops", {})
    amazon_known = [str(x).strip() for x in (shops.get("amazon_known") or []) if str(x).strip()]
    d01_known = [str(x).strip() for x in (shops.get("d01_known") or []) if str(x).strip()]
    amazon_l = {item.lower() for item in amazon_known}
    d01_l = {item.lower() for item in d01_known}
    for item in data["shops"]:
        account = item["account"]
        key = account.lower()
        if item["kind"] == "amazon":
            d01_known = [row for row in d01_known if row.lower() != key]
            d01_l.discard(key)
            if key not in amazon_l:
                amazon_known.append(account)
                amazon_l.add(key)
        else:
            amazon_known = [row for row in amazon_known if row.lower() != key]
            amazon_l.discard(key)
            if key not in d01_l:
                d01_known.append(account)
                d01_l.add(key)
    shops["amazon_known"] = amazon_known
    shops["d01_known"] = d01_known
    cls = type(rules)
    return cls(raw=raw)


def load_merged_rules(load_rules: Callable[..., Any]) -> Any:
    return merge_into_rules(load_rules())


def public_allowlist() -> dict[str, Any]:
    data = load_allowlist()
    shops = []
    for item in data["shops"]:
        kind_label = "按亚马逊处理" if item["kind"] == "amazon" else "和其他店一样处理（店管家01）"
        shops.append({"account": item["account"], "kind": item["kind"], "kind_label": kind_label})
    return {"sku_suffixes": data["sku_suffixes"], "shops": shops}
