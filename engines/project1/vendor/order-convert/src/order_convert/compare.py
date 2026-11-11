# -*- coding: utf-8 -*-
"""与人工样板逐项对比（内容、顺序、红字、空白列）。"""
from __future__ import annotations

from pathlib import Path
from typing import Any

from .io_excel import load_xls_nonzero, load_xlsx_rows, norm, vals_equal


D2_RED_COLS = [
    "订单编号", "客户网名", "收货人", "电话", "国家/地区", "地址",
    "客服备注", "客户备注", "编号", "条码", "数量", "合计", "订单来源", "店铺",
]


def diff_dianxiaomi(generated: Path, manual: Path) -> list[dict]:
    headers, gen = load_xlsx_rows(generated)
    _, man = load_xlsx_rows(manual)
    man = sorted(man, key=lambda x: x["_excel_row"])
    diffs: list[dict] = []
    if len(gen) != len(man):
        diffs.append({"field": "ROW_COUNT", "gen": len(gen), "man": len(man)})
        return diffs
    for i, (g, m) in enumerate(zip(gen, man)):
        if norm(g.get("运单号")) != norm(m.get("运单号")):
            diffs.append({"pos": i + 1, "field": "ORDER", "gen": g.get("运单号"), "man": m.get("运单号")})
            continue
        for h in headers:
            if not vals_equal(g.get(h), m.get(h)):
                diffs.append({"pos": i + 1, "field": h, "ono": g.get("订单号"), "gen": g.get(h), "man": m.get(h)})
        if g.get("_all_red") != m.get("_all_red"):
            diffs.append({"pos": i + 1, "field": "RED", "ono": g.get("订单号"), "gen": g.get("_all_red"), "man": m.get("_all_red")})
        elif g.get("_all_red") and set(m.get("_red_cols") or []) != set(headers):
            diffs.append({"pos": i + 1, "field": "RED_PARTIAL", "ono": g.get("订单号"), "man_red": len(m.get("_red_cols") or [])})
    return diffs


def diff_d01(generated: Path, manual: Path) -> list[dict]:
    h_g, gen = load_xls_nonzero(generated)
    h_m, man = load_xls_nonzero(manual)
    diffs: list[dict] = []
    if len(gen) != len(man):
        diffs.append({"field": "ROW_COUNT", "gen": len(gen), "man": len(man)})
        return diffs
    # compare all header columns from manual
    headers = h_m
    for i, (g, m) in enumerate(zip(gen, man)):
        if norm(g.get("客户备注")) != norm(m.get("客户备注")):
            diffs.append({"pos": i + 1, "field": "ORDER", "gen": g.get("客户备注"), "man": m.get("客户备注")})
            continue
        for h in headers:
            if not vals_equal(g.get(h, ""), m.get(h, "")):
                diffs.append({"pos": i + 1, "field": h, "gen": g.get(h), "man": m.get(h)})
    return diffs


def diff_d02(generated: Path, manual: Path, red_columns: list[str] | None = None) -> list[dict]:
    red_columns = red_columns or D2_RED_COLS
    _, gen = load_xls_nonzero(generated)
    h_m, man = load_xls_nonzero(manual)
    diffs: list[dict] = []
    if len(gen) != len(man):
        diffs.append({"field": "ROW_COUNT", "gen": len(gen), "man": len(man)})
        return diffs
    for i, (g, m) in enumerate(zip(gen, man)):
        if norm(g.get("客户备注")) != norm(m.get("客户备注")):
            diffs.append({"pos": i + 1, "field": "ORDER", "gen": g.get("客户备注"), "man": m.get("客户备注")})
            continue
        for h in h_m:
            if not vals_equal(g.get(h, ""), m.get(h, "")):
                diffs.append({"pos": i + 1, "field": h, "ono": g.get("订单编号"), "gen": g.get(h), "man": m.get(h)})
        if set(g.get("_red_cols") or []) != set(red_columns):
            diffs.append({"pos": i + 1, "field": "RED_COLS_GEN", "ono": g.get("订单编号"), "gen": g.get("_red_cols")})
        if set(m.get("_red_cols") or []) != set(red_columns):
            diffs.append({"pos": i + 1, "field": "RED_COLS_MAN", "ono": m.get("订单编号"), "man": m.get("_red_cols")})
    return diffs
