# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path
from typing import Any

import openpyxl
import xlrd


def norm(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and abs(v - int(v)) < 1e-9:
        return str(int(v))
    return str(v).strip()


def load_xlsx_rows(path: Path) -> tuple[list[str], list[dict]]:
    wb = openpyxl.load_workbook(path, data_only=False)
    ws = wb.active
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    rows: list[dict] = []
    for r in range(2, ws.max_row + 1):
        row: dict = {"_excel_row": r}
        red: list[str] = []
        for c, h in enumerate(headers):
            cell = ws.cell(r, c + 1)
            row[h] = cell.value
            color = None
            if cell.font and cell.font.color and cell.font.color.type == "rgb" and cell.font.color.rgb:
                color = str(cell.font.color.rgb)
            # 仅 RGB 末 6 位为 FF0000 才算红；避免 FF000000（黑）被 "FF0000" in s 误伤
            if color:
                s = color.upper().lstrip("#")
                rgb6 = s[2:] if len(s) == 8 else s
                if rgb6 == "FF0000":
                    red.append(h)
        row["_red_cols"] = red
        row["_all_red"] = len(red) == len(headers)
        rows.append(row)
    return headers, rows


def load_xls_nonzero(path: Path, sheet_idx: int = 0) -> tuple[list[str], list[dict]]:
    wb = xlrd.open_workbook(str(path), formatting_info=True)
    sh = wb.sheet_by_index(sheet_idx)
    headers = [sh.cell_value(0, c) for c in range(sh.ncols)]
    rows: list[dict] = []
    for r in range(1, sh.nrows):
        vals = [sh.cell_value(r, c) for c in range(sh.ncols)]
        if not any(v not in ("", None) for v in vals):
            continue
        row = {headers[c]: vals[c] for c in range(len(headers))}
        row["_excel_row"] = r + 1
        red: list[str] = []
        for c in range(sh.ncols):
            xf = wb.xf_list[sh.cell_xf_index(r, c)]
            if wb.font_list[xf.font_index].colour_index == 10:
                red.append(headers[c])
        row["_red_cols"] = red
        rows.append(row)
    return headers, rows


def vals_equal(a: Any, b: Any) -> bool:
    if a == b or norm(a) == norm(b):
        return True
    try:
        return float(a) == float(b)
    except Exception:
        return False
