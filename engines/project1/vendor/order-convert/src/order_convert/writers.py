# -*- coding: utf-8 -*-
"""
写出三份 Excel。

店管家 01/02（.xls）：复制内置空白模板（xlutils.copy）再写入数据行，
以尽量保留列宽、表头样式、工作表结构；不带入旧订单（模板已清空数据）。

店小秘（.xlsx）：复制空白表头模板后写入；亚马逊整行红字（ARGB FFFF0000）；
数据行强制宋体 11 + 居中 + 自动换行；默认行高 25（样板实测，约双行）。

店管家行高（样板实测）：
- 01：有内容数据行多为 540 twips（27pt，≈默认 285 的双倍）
- 02：有内容数据行 1080/1350/2430 twips，按长文本列大致分级

.xls 技术限制：BIFF/xlwt 无法完美克隆全部 Excel 特性；
红字用 colour_index=10；数据行样式由写入端按样板显式构造。
"""
from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Font
import xlrd
import xlwt
from xlutils.copy import copy as xl_copy

from .resources_path import default_template_d01, default_template_d02, default_template_dx

# 店小秘样板实测
DX_FONT_NAME = "宋体"
DX_FONT_SIZE = 11
DX_DEFAULT_ROW_HEIGHT = 25.0  # 样板 sheet_format.defaultRowHeight
DX_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

# xlwt Font.height 单位为 twips（1/20 pt）：180=9pt, 220=11pt, 240=12pt
XLS_PT9 = 180
XLS_PT11 = 220
XLS_PT12 = 240

# 店管家样板行高（twips）
D01_DATA_ROW_HEIGHT = 540   # 27.0pt ≈ 双行
D02_DATA_ROW_HEIGHT_BASE = 1080   # 54.0pt（样板最短有内容行）
D02_DATA_ROW_HEIGHT_MID = 1350    # 67.5pt
D02_DATA_ROW_HEIGHT_TALL = 2430   # 121.5pt（长备注/地址常见）


def normalize_argb(rgb: str) -> str:
    """统一为 8 位 ARGB（如 FFFF0000）。兼容传入 FF0000 / #FF0000 / 00FF0000。"""
    s = (rgb or "").strip().lstrip("#").upper()
    if len(s) == 6:
        return "FF" + s
    if len(s) == 8:
        # 历史 00FF0000 透明红在部分软件不可见 → 强制不透明红
        if s.endswith("FF0000") and s[:2] == "00":
            return "FFFF0000"
        return s
    return "FFFF0000"


def xls_headers(template: Path, sheet_idx: int = 0) -> list[str]:
    rb = xlrd.open_workbook(str(template), formatting_info=True)
    sh = rb.sheet_by_index(sheet_idx)
    return [sh.cell_value(0, c) for c in range(sh.ncols)]


def _dx_font(*, red: bool, red_rgb: str) -> Font:
    if red:
        return Font(name=DX_FONT_NAME, size=DX_FONT_SIZE, color=normalize_argb(red_rgb))
    # 与样板一致：非红行不写死颜色（避免 FF000000 被旧检测误判）
    return Font(name=DX_FONT_NAME, size=DX_FONT_SIZE)


def _xls_set_row_height(sheet: Any, row_idx: int, height_twips: int) -> None:
    """xlwt：设置行高（twips），须同时设 height_mismatch=1 才会写入文件。"""
    row = sheet.row(row_idx)
    row.height = int(height_twips)
    row.height_mismatch = 1


def write_dianxiaomi(
    path: Path,
    headers: list[str],
    rows: list[dict],
    *,
    template: Path | None = None,
    red_rgb: str = "FFFF0000",
) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tpl = Path(template) if template else default_template_dx()
    if tpl.exists():
        shutil.copy2(tpl, path)
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        for c, h in enumerate(headers, 1):
            ws.cell(1, c, h)
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "order_1"
        for c, h in enumerate(headers, 1):
            cell = ws.cell(1, c, h)
            cell.font = Font(name=DX_FONT_NAME, size=DX_FONT_SIZE, bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    # 与人工样板一致：默认行高约双行，数据行显式同高 + 自动换行
    ws.sheet_format.defaultRowHeight = DX_DEFAULT_ROW_HEIGHT

    red_font = _dx_font(red=True, red_rgb=red_rgb)
    black_font = _dx_font(red=False, red_rgb=red_rgb)
    for i, row in enumerate(rows):
        use_red = bool(row.get("_all_red"))
        font = red_font if use_red else black_font
        excel_row = i + 2
        for c, h in enumerate(headers, 1):
            cell = ws.cell(excel_row, c, row.get(h))
            cell.font = font
            cell.alignment = DX_ALIGN
        ws.row_dimensions[excel_row].height = DX_DEFAULT_ROW_HEIGHT
    wb.save(path)


def _xls_font(name: str = "宋体", height: int = XLS_PT11, colour_index: int = 32767) -> xlwt.Font:
    fnt = xlwt.Font()
    fnt.name = name
    fnt.height = height
    fnt.colour_index = colour_index
    return fnt


def _xls_style(
    *,
    name: str = "宋体",
    height: int = XLS_PT11,
    colour_index: int = 32767,
    num_format_str: str = "General",
    horz: int = xlwt.Alignment.HORZ_CENTER,
    vert: int = xlwt.Alignment.VERT_CENTER,
    wrap: int = 1,
) -> xlwt.XFStyle:
    st = xlwt.XFStyle()
    st.font = _xls_font(name=name, height=height, colour_index=colour_index)
    st.num_format_str = num_format_str
    al = xlwt.Alignment()
    al.horz = horz
    al.vert = vert
    al.wrap = wrap
    st.alignment = al
    return st


# 店管家01：有内容列≈样板 宋体11 居中；其余列 宋体9 文本@
_D01_CONTENT_COLS = {"原始单号", "客户备注", "客服备注", "自定义1", "订单编号"}


def write_dian_01(
    path: Path,
    rows: list[dict],
    template: Path | None = None,
) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tpl = Path(template) if template else default_template_d01()
    if not tpl.exists():
        raise FileNotFoundError(f"店管家01模板不存在: {tpl}")

    rb = xlrd.open_workbook(str(tpl), formatting_info=True)
    headers = [rb.sheet_by_index(0).cell_value(0, c) for c in range(rb.sheet_by_index(0).ncols)]
    wb = xl_copy(rb)
    sheet = wb.get_sheet(0)

    style_content = _xls_style(
        height=XLS_PT11,
        num_format_str="General",
        horz=xlwt.Alignment.HORZ_CENTER,
        vert=xlwt.Alignment.VERT_CENTER,
        wrap=1,
    )
    style_empty = _xls_style(
        height=XLS_PT9,
        num_format_str="@",
        horz=xlwt.Alignment.HORZ_GENERAL,
        vert=xlwt.Alignment.VERT_BOTTOM,
        wrap=0,
    )
    # 长单号强制文本，避免科学计数
    style_text = _xls_style(
        height=XLS_PT11,
        num_format_str="@",
        horz=xlwt.Alignment.HORZ_CENTER,
        vert=xlwt.Alignment.VERT_CENTER,
        wrap=1,
    )

    for ri, row in enumerate(rows, 1):
        for c, h in enumerate(headers):
            v = row.get(h, "")
            if v is None:
                v = ""
            if h in ("原始单号", "客户备注", "自定义1", "订单编号"):
                sheet.write(ri, c, str(v) if v != "" else "", style_text)
            elif h in _D01_CONTENT_COLS:
                sheet.write(ri, c, v if v != "" else "", style_content if v != "" else style_empty)
            else:
                sheet.write(ri, c, v if v != "" else "", style_content if v != "" else style_empty)
        # 样板有内容数据行多为 540 twips（双行）
        _xls_set_row_height(sheet, ri, D01_DATA_ROW_HEIGHT)
    wb.save(str(path))


def _d02_row_height_twips(row: dict) -> int:
    """按样板分级：短内容 1080，中等 1350，长备注/地址 2430。"""
    note = str(row.get("客服备注") or "")
    addr = str(row.get("地址") or "")
    custom = str(row.get("自定义1") or "")
    longest = max(len(note), len(addr), len(custom), 0)
    has_newline = ("\n" in note) or ("\n" in addr) or ("\n" in custom)
    if has_newline or len(note) >= 30 or longest >= 80:
        return D02_DATA_ROW_HEIGHT_TALL
    if len(note) > 20 or len(addr) > 70 or longest >= 60:
        return D02_DATA_ROW_HEIGHT_MID
    return D02_DATA_ROW_HEIGHT_BASE


def write_dian_02(
    path: Path,
    rows: list[dict],
    template: Path | None = None,
    red_columns: list[str] | None = None,
) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tpl = Path(template) if template else default_template_d02()
    if not tpl.exists():
        raise FileNotFoundError(f"店管家02模板不存在: {tpl}")
    red_columns = red_columns or []

    rb = xlrd.open_workbook(str(tpl), formatting_info=True)
    headers = [rb.sheet_by_index(0).cell_value(0, c) for c in range(rb.sheet_by_index(0).ncols)]
    wb = xl_copy(rb)
    sheet = wb.get_sheet(0)

    # 样板：红字有内容列 宋体11 + colour_index=10 + General 居中 wrap
    style_red = _xls_style(
        height=XLS_PT11,
        colour_index=10,
        num_format_str="General",
        horz=xlwt.Alignment.HORZ_CENTER,
        vert=xlwt.Alignment.VERT_CENTER,
        wrap=1,
    )
    style_red_text = _xls_style(
        height=XLS_PT11,
        colour_index=10,
        num_format_str="@",
        horz=xlwt.Alignment.HORZ_CENTER,
        vert=xlwt.Alignment.VERT_CENTER,
        wrap=1,
    )
    # 非红有内容：宋体11 黑
    style_norm = _xls_style(
        height=XLS_PT11,
        colour_index=32767,
        num_format_str="General",
        horz=xlwt.Alignment.HORZ_CENTER,
        vert=xlwt.Alignment.VERT_CENTER,
        wrap=1,
    )
    style_norm_text = _xls_style(
        height=XLS_PT11,
        colour_index=32767,
        num_format_str="@",
        horz=xlwt.Alignment.HORZ_CENTER,
        vert=xlwt.Alignment.VERT_CENTER,
        wrap=1,
    )
    # 空列近似样板：宋体12 + @
    style_empty = _xls_style(
        height=XLS_PT12,
        colour_index=32767,
        num_format_str="@",
        horz=xlwt.Alignment.HORZ_GENERAL,
        vert=xlwt.Alignment.VERT_BOTTOM,
        wrap=0,
    )

    text_cols = {"订单编号", "客户备注", "自定义1", "编号", "条码", "电话", "客户网名", "收货人"}

    for ri, row in enumerate(rows, 1):
        for c, h in enumerate(headers):
            v = row.get(h, "")
            if v is None:
                v = ""
            is_red = h in red_columns
            empty = v == ""
            if empty:
                st = style_red_text if is_red else style_empty
                sheet.write(ri, c, "", st)
                continue
            if h in text_cols:
                st = style_red_text if is_red else style_norm_text
                sheet.write(ri, c, str(v), st)
            else:
                st = style_red if is_red else style_norm
                sheet.write(ri, c, v, st)
        _xls_set_row_height(sheet, ri, _d02_row_height_twips(row))
    wb.save(str(path))
