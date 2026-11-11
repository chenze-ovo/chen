# -*- coding: utf-8 -*-
"""从人工样板提取空白模板到包内 templates/。

- 店小秘 xlsx：删光数据行，保留表头字体/对齐/列宽。
- 店管家 xls：复制后删掉全部数据行（只留表头），保留表头 XF 与列宽。
  注意：勿对数据格 write("")，那会冲掉 XF 变成 Arial；数据行样式由 writers 按样板重建。
"""
from pathlib import Path

from xlrd import open_workbook
from xlutils.copy import copy
import openpyxl

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "src" / "order_convert" / "templates"


def blank_xls(src: Path, dst: Path) -> None:
    rb = open_workbook(str(src), formatting_info=True)
    wb = copy(rb)
    for si in range(rb.nsheets):
        wsh = wb.get_sheet(si)
        rows = wsh._Worksheet__rows
        for k in list(rows.keys()):
            if k != 0:
                del rows[k]
    dst.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(dst))


def main() -> None:
    blank_xls(ROOT / "dian_01.xls", OUT / "dian_01_blank.xls")
    blank_xls(ROOT / "dian_02.xls", OUT / "dian_02_blank.xls")
    wb = openpyxl.load_workbook(ROOT / "dianxiaomi.xlsx")
    ws = wb.active
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)
    wb.save(OUT / "dianxiaomi_blank.xlsx")
    print("templates ->", OUT)


if __name__ == "__main__":
    main()
