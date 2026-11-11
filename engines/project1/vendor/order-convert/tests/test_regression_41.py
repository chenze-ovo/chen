# -*- coding: utf-8 -*-
"""41 单样本回归：与人工三表比对；非欧元币种接物流为新规则，比对时忽略该后缀差。"""
from __future__ import annotations

import csv
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]

from order_convert.compare import diff_d01, diff_d02, diff_dianxiaomi
from order_convert.config import load_rules
from order_convert.pipeline import convert_order_file
from order_convert.resources_path import default_template_d01, default_template_d02


@pytest.fixture(scope="module")
def converted():
    out = ROOT / "out"
    out.mkdir(exist_ok=True)
    # 故意不传 template_*，验证内置空白模板
    assert default_template_d01().exists()
    assert default_template_d02().exists()
    result = convert_order_file(
        ROOT / "raw_order.xlsx",
        out,
        rules=load_rules(),
        exports=("dx", "d01", "d02"),
    )
    return result


def test_builtin_templates_used(converted):
    assert converted.formal_export_blocked is False
    assert converted.export_status == "ok"
    assert converted.d01_path == ROOT / "out" / "dian_01.xls"
    assert converted.d02_path == ROOT / "out" / "dian_02.xls"
    assert converted.stats["errors"] == 0


def test_row_counts(converted):
    assert converted.stats["raw_orders"] == 41
    assert converted.stats["dx_orders"] == 41
    assert converted.stats["d01_orders"] == 6
    assert converted.stats["d02_orders"] == 35
    assert converted.stats["pending_orders"] == 0
    assert converted.stats["china_direct"] == 27


def _strip_known_currency_suffix(text: str, skip: set[str]) -> str:
    """若文案末尾是非 skip 币种缩写，去掉后用于与历史金标比对。"""
    s = "" if text is None else str(text)
    for cur in ("USD", "GBP", "JPY", "CAD", "AUD", "HKD", "CNY", "SGD"):
        if cur in skip:
            continue
        if s.endswith(cur):
            return s[: -len(cur)]
    return s


def _filter_currency_only_diffs(diffs: list[dict], field: str, skip: set[str]) -> list[dict]:
    """保留非「仅币种后缀」差异；金标为人工作业时尚无币种拼接规则。"""
    kept = []
    for d in diffs:
        if d.get("field") != field:
            kept.append(d)
            continue
        gen = "" if d.get("gen") is None else str(d.get("gen"))
        man = "" if d.get("man") is None else str(d.get("man"))
        gen_base = _strip_known_currency_suffix(gen, skip)
        if gen_base == man:
            continue  # 仅多了币种后缀
        kept.append(d)
    return kept


def test_dianxiaomi_zero_diff(converted):
    diffs = diff_dianxiaomi(converted.dx_path, ROOT / "dianxiaomi.xlsx")
    skip = load_rules().currency_skip
    # 新规则：非欧元币种接到物流方式后；与历史金标比对时忽略此类后缀差
    diffs = _filter_currency_only_diffs(diffs, "物流方式", skip)
    assert diffs == [], diffs[:10]


def test_d01_zero_diff(converted):
    diffs = diff_d01(converted.d01_path, ROOT / "dian_01.xls")
    assert diffs == [], diffs[:10]


def test_d02_zero_diff(converted):
    rules = load_rules()
    diffs = diff_d02(converted.d02_path, ROOT / "dian_02.xls", rules.d02_red_columns)
    diffs = _filter_currency_only_diffs(diffs, "客服备注", rules.currency_skip)
    assert diffs == [], diffs[:10]


def test_no_phone_warnings_for_d01_only_cases(converted):
    phone_warns = [a for a in converted.anomalies if a["type"] == "PHONE_FALLBACK_UNCONFIRMED"]
    assert phone_warns == [], phone_warns


def test_m00016_marker_in_logistics(converted):
    rules = load_rules()
    marker = rules.china_marker
    china = [m for m in converted.mappings if m["is_china_direct"]]
    assert len(china) == 27
    skip = rules.currency_skip
    for m in china:
        log = m["dx_物流方式"]
        assert marker in log
        # 币种在中国直发标记之后（欧元不追加）
        assert log.startswith(m["原始物流方式"]) or m["原始物流方式"] in log
        base = log
        for cur in ("USD", "GBP", "JPY", "CAD", "AUD"):
            if cur not in skip and base.endswith(cur):
                base = base[: -len(cur)]
                break
        assert base.endswith(marker)
        assert "---" not in m["dx_物流方式"]


def test_row_mapping_csv_columns(converted):
    path = ROOT / "out" / "row_mapping.csv"
    with open(path, encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    assert len(rows) == 41
    required = {"原始SKU", "店管家01行号", "店管家02行号", "dx_seq", "运单号"}
    assert required <= set(rows[0].keys())
    assert sum(1 for r in rows if r["店管家01行号"]) == 6
    assert sum(1 for r in rows if r["店管家02行号"]) == 35
