# -*- coding: utf-8 -*-
"""边界与异常分级测试。"""
from __future__ import annotations

from pathlib import Path
from shutil import copyfile

import openpyxl
import pytest

ROOT = Path(__file__).resolve().parents[1]

from order_convert.config import load_rules
from order_convert.pipeline import convert_order_file
from order_convert.transform import is_china_direct, is_sku_like, d01_original_order_no, split_sku_lines
from order_convert.writers import normalize_argb
from order_convert.io_excel import load_xlsx_rows


def _clone_raw(tmp_path: Path, mutator) -> Path:
    dst = tmp_path / "raw.xlsx"
    copyfile(ROOT / "raw_order.xlsx", dst)
    wb = openpyxl.load_workbook(dst)
    ws = wb.active
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    mutator(ws, headers)
    wb.save(dst)
    return dst


def _col(headers, name):
    return headers.index(name) + 1


def _make_minimal_raw(tmp_path: Path, rows: list[dict]) -> Path:
    """合成最小原始表（不写死真实盲测单号业务分支）。"""
    src = openpyxl.load_workbook(ROOT / "raw_order.xlsx")
    headers = [src.active.cell(1, c).value for c in range(1, src.active.max_column + 1)]
    # 取第2行作模板字段
    template = {h: src.active.cell(2, i + 1).value for i, h in enumerate(headers)}
    dst = tmp_path / "synth.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    for c, h in enumerate(headers, 1):
        ws.cell(1, c, h)
    for ri, override in enumerate(rows, 2):
        for c, h in enumerate(headers, 1):
            ws.cell(ri, c, override.get(h, template.get(h)))
    wb.save(dst)
    return dst


def test_china_marker_incomplete_not_trigger(tmp_path):
    """拣货备注仅「中国直发M00016」无完整句 → 不触发。"""
    def mut(ws, headers):
        c = _col(headers, "拣货备注")
        ws.cell(2, c, "中国直发M00016")
    raw = _clone_raw(tmp_path, mut)
    rules = load_rules()
    _, rows = load_xlsx_rows(raw)
    r = rows[0]
    assert "中国直发M00016" in str(r["拣货备注"])
    assert rules.china_marker not in str(r["拣货备注"])
    assert is_china_direct(r, rules) is False
    out = tmp_path / "out"
    result = convert_order_file(raw, out)
    m = next(x for x in result.mappings if x["raw_excel_row"] == 2)
    assert m["is_china_direct"] is False
    assert not str(m["dx_订单号"]).endswith("-1")


def test_china_marker_full_with_different_logistics_still_trigger(tmp_path):
    """完整句在备注中，即使物流改成别的仍触发。"""
    def mut(ws, headers):
        ws.cell(2, _col(headers, "拣货备注"), "前缀 " + load_rules().china_marker + " 后缀")
        ws.cell(2, _col(headers, "物流方式"), "完全不同的物流名称XYZ")
    raw = _clone_raw(tmp_path, mut)
    result = convert_order_file(raw, tmp_path / "out")
    m = next(x for x in result.mappings if x["raw_excel_row"] == 2)
    assert m["is_china_direct"] is True
    assert m["dx_订单号"].endswith("-1")
    assert m["dx_SKU"] == "M00016"
    assert m["dx_物流方式"].endswith(load_rules().china_marker)


def test_unknown_sku_suffix_needs_review_no_formal(tmp_path):
    def mut(ws, headers):
        for r in range(2, ws.max_row + 1):
            note = ws.cell(r, _col(headers, "拣货备注")).value or ""
            shop = str(ws.cell(r, _col(headers, "店铺账号")).value or "")
            if "中国直发M00016 海外仓发产品" not in str(note) and "亚马逊" in shop:
                ws.cell(r, _col(headers, "SKU"), "WX99999-UNKNOWN")
                break
    raw = _clone_raw(tmp_path, mut)
    out = tmp_path / "out"
    result = convert_order_file(raw, out)
    warns = [a for a in result.warnings if a["type"] == "UNKNOWN_SKU_SUFFIX"]
    assert warns, result.anomalies
    assert result.export_status == "needs_review"
    assert result.formal_export_blocked is False
    assert not (out / "dian_01.xls").exists()
    assert not (out / "dian_02.xls").exists()
    assert not (out / "dianxiaomi.xlsx").exists()
    assert (out / "review" / "dianxiaomi_NEEDS_REVIEW.xlsx").exists()
    assert (out / "review" / "dian_01_NEEDS_REVIEW.xls").exists()
    assert (out / "review" / "dian_02_NEEDS_REVIEW.xls").exists()
    m = next(x for x in result.mappings if x["原始SKU"] == "WX99999-UNKNOWN")
    # 未知后缀：输出基础SKU，仍 needs_review
    assert m["dx_SKU"] == "WX99999"
    assert m["base_sku"] == "WX99999"
    assert any(a.get("suffix") == "UNKNOWN" for a in warns)


def test_multi_sku_newline_detects_each_suffix_needs_review(tmp_path):
    """换行多SKU：分别识别后缀；不拆明细；待复核。"""
    rows = [
        {
            "订单号": "SYNTH-MULTI-001",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "M0019-HKCQ\nM0003-CQCS",
            "产品总数": 3,
            "拣货备注": "",
            "物流方式": "云途全球专线挂号",
            "运单号": "TRACK-MULTI-001",
            "订单金额": 10.5,
        }
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    out = tmp_path / "out"
    result = convert_order_file(raw, out)
    multi = [a for a in result.warnings if a["type"] == "MULTI_SKU"]
    assert len(multi) == 1
    assert multi[0]["line_count"] == 2
    suffixes = {
        a["suffix"]
        for a in result.warnings
        if a["type"] == "UNKNOWN_SKU_SUFFIX" and a["订单号"] == "SYNTH-MULTI-001"
    }
    assert suffixes == {"HKCQ", "CQCS"}
    assert result.export_status == "needs_review"
    assert not (out / "dian_02.xls").exists()
    assert (out / "review" / "dian_02_NEEDS_REVIEW.xls").exists()
    m = next(x for x in result.mappings if x["原始订单号"] == "SYNTH-MULTI-001")
    assert m["is_multi_sku"] is True
    assert m["sku_line_count"] == 2
    # 不拆成多行销售明细；各行已去后缀
    assert result.stats["d02_orders"] == 1
    assert m["dx_SKU"] == "M0019\nM0003"
    assert m["base_sku"] == "M0019\nM0003"


def test_multi_sku_no_suffix_still_needs_review(tmp_path):
    rows = [
        {
            "订单号": "SYNTH-MULTI-002",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "M0002\nM0004",
            "产品总数": 4,
            "拣货备注": "",
            "物流方式": "燕文航空挂号",
            "运单号": "TRACK-MULTI-002",
            "订单金额": 8,
        }
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    result = convert_order_file(raw, tmp_path / "out")
    assert any(a["type"] == "MULTI_SKU" for a in result.warnings)
    assert result.export_status == "needs_review"
    assert result.stats["d02_orders"] == 1  # 不拆


def test_quantity_greater_than_one(tmp_path):
    rows = [
        {
            "订单号": "SYNTH-QTY-003",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "WX10001-LYF",
            "产品总数": 5,
            "拣货备注": "",
            "物流方式": "云途全球专线挂号",
            "运单号": "TRACK-QTY-003",
            "订单金额": 20,
        }
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    result = convert_order_file(raw, tmp_path / "out")
    assert result.export_status == "ok"
    assert result.stats["d02_orders"] == 1
    # 数量写入 02
    import xlrd
    sh = xlrd.open_workbook(str(result.d02_path)).sheet_by_index(0)
    headers = [sh.cell_value(0, c) for c in range(sh.ncols)]
    qty_c = headers.index("数量")
    assert sh.cell_value(1, qty_c) == 5
    # 白名单后缀已剥离
    m = result.mappings[0]
    assert m["dx_SKU"] == "WX10001"
    assert m["base_sku"] == "WX10001"


def test_new_logistics_pinyin_sort_stable(tmp_path):
    """新物流方式仍按拼音稳定排序（不写死盲测单号）。"""
    rows = [
        {
            "订单号": "SYNTH-LOG-A",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "A-LYF",
            "产品总数": 1,
            "拣货备注": "",
            "物流方式": "浙中测试专线",
            "运单号": "T-A",
            "订单金额": 1,
        },
        {
            "订单号": "SYNTH-LOG-B",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "B-LYF",
            "产品总数": 1,
            "拣货备注": "",
            "物流方式": "安徽测试专线",
            "运单号": "T-B",
            "订单金额": 1,
        },
        {
            "订单号": "SYNTH-LOG-C",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "C-LYF",
            "产品总数": 1,
            "拣货备注": "",
            "物流方式": "北京测试专线",
            "运单号": "T-C",
            "订单金额": 1,
        },
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    result = convert_order_file(raw, tmp_path / "out")
    assert result.export_status == "ok"
    order_nos = [m["原始订单号"] for m in result.mappings]
    # 安徽 anhui < 北京 beijing < 浙中 zhezhong
    assert order_nos == ["SYNTH-LOG-B", "SYNTH-LOG-C", "SYNTH-LOG-A"]


def test_known_and_unknown_suffix_mix_in_batch(tmp_path):
    rows = [
        {
            "订单号": "SYNTH-SUF-OK",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "WX1-XYTU",
            "产品总数": 1,
            "拣货备注": "",
            "物流方式": "云途A",
            "运单号": "T-OK",
            "订单金额": 1,
        },
        {
            "订单号": "SYNTH-SUF-BAD",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "WX2-ZZZZ",
            "产品总数": 2,
            "拣货备注": "",
            "物流方式": "云途B",
            "运单号": "T-BAD",
            "订单金额": 2,
        },
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    out = tmp_path / "out"
    result = convert_order_file(raw, out)
    assert result.export_status == "needs_review"
    assert any(a["suffix"] == "ZZZZ" for a in result.warnings if a["type"] == "UNKNOWN_SKU_SUFFIX")
    m_ok = next(x for x in result.mappings if x["原始订单号"] == "SYNTH-SUF-OK")
    assert m_ok["dx_SKU"] == "WX1"
    m_bad = next(x for x in result.mappings if x["原始订单号"] == "SYNTH-SUF-BAD")
    assert m_bad["dx_SKU"] == "WX2"
    assert m_bad["base_sku"] == "WX2"
    assert not (out / "dianxiaomi.xlsx").exists()


def test_unknown_suffix_shuu_strips_to_base_needs_review(tmp_path):
    """用户要求：WX17023-SHUU → WX17023；未知后缀仍待复核。"""
    rows = [
        {
            "订单号": "SYNTH-SHUU-001",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "WX17023-SHUU",
            "产品总数": 1,
            "拣货备注": "",
            "物流方式": "云途全球专线挂号",
            "运单号": "TRACK-SHUU-001",
            "订单金额": 9.9,
        }
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    out = tmp_path / "out"
    result = convert_order_file(raw, out)
    warns = [a for a in result.warnings if a["type"] == "UNKNOWN_SKU_SUFFIX"]
    assert any(a.get("suffix") == "SHUU" for a in warns)
    assert result.export_status == "needs_review"
    m = result.mappings[0]
    assert m["原始SKU"] == "WX17023-SHUU"
    assert m["dx_SKU"] == "WX17023"
    assert m["base_sku"] == "WX17023"
    assert not (out / "dianxiaomi.xlsx").exists()
    assert (out / "review" / "dianxiaomi_NEEDS_REVIEW.xlsx").exists()


def test_split_sku_lines_helper():
    assert split_sku_lines("A-HKCQ\nB-CQCS") == ["A-HKCQ", "B-CQCS"]
    assert split_sku_lines("A\r\nB\rC") == ["A", "B", "C"]
    assert split_sku_lines("  ") == []


def test_red_font_argb_normalized():
    assert normalize_argb("FF0000") == "FFFF0000"
    assert normalize_argb("00FF0000") == "FFFF0000"
    assert normalize_argb("FFFF0000") == "FFFF0000"
    assert normalize_argb("#ff0000") == "FFFF0000"


def test_dianxiaomi_red_is_ffff0000(tmp_path):
    rows = [
        {
            "订单号": "SYNTH-RED-001",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "R-LYF",
            "产品总数": 1,
            "拣货备注": "",
            "物流方式": "云途",
            "运单号": "T-RED",
            "订单金额": 1,
        }
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    result = convert_order_file(raw, tmp_path / "out")
    assert result.export_status == "ok"
    wb = openpyxl.load_workbook(result.dx_path)
    cell = wb.active.cell(2, 1)
    assert cell.font.name == "宋体"
    assert cell.font.size == 11
    assert cell.alignment.horizontal == "center"
    assert cell.alignment.vertical == "center"
    rgb = str(cell.font.color.rgb) if cell.font.color and cell.font.color.type == "rgb" else ""
    assert rgb.upper() == "FFFF0000"


def test_dianxiaomi_non_amazon_black_songti(tmp_path):
    rows = [
        {
            "订单号": "SYNTH-BLK-001",
            "店铺账号": "ebay宇星店德国（自营）",
            "SKU": "W-LYF",
            "产品总数": 1,
            "拣货备注": "",
            "物流方式": "eBay线上",
            "运单号": "T-BLK",
            "订单金额": 1,
            "买家指定物流": "DE",
        }
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    result = convert_order_file(raw, tmp_path / "out")
    assert result.export_status == "ok"
    wb = openpyxl.load_workbook(result.dx_path)
    cell = wb.active.cell(2, 1)
    assert cell.font.name == "宋体"
    assert cell.font.size == 11
    rgb = str(cell.font.color.rgb) if cell.font.color and cell.font.color.type == "rgb" else ""
    assert rgb.upper() in ("", "FF000000", "00000000") or not rgb.upper().endswith("FF0000")


def test_d01_d02_data_font_is_songti(tmp_path):
    """店管家数据行须为宋体（非 xlwt 默认 Arial）。"""
    import xlrd

    result = convert_order_file(ROOT / "raw_order.xlsx", tmp_path / "out")
    assert result.export_status == "ok"
    for path, label in [(result.d01_path, "d01"), (result.d02_path, "d02")]:
        rb = xlrd.open_workbook(str(path), formatting_info=True)
        sh = rb.sheet_by_index(0)
        assert sh.nrows >= 2, label
        # 找第一个有内容的数据行
        found = False
        for r in range(1, sh.nrows):
            if any(str(sh.cell_value(r, c)).strip() for c in range(sh.ncols)):
                for c in range(min(8, sh.ncols)):
                    if str(sh.cell_value(r, c)).strip() == "":
                        continue
                    xf = rb.xf_list[sh.cell(r, c).xf_index]
                    font = rb.font_list[xf.font_index]
                    assert font.name == "宋体", (label, r, c, font.name)
                    found = True
                    break
                break
        assert found, label


def test_dianxiaomi_wrap_and_double_row_height(tmp_path):
    """店小秘数据行须 wrap_text，行高≈样板默认 25（双行）。"""
    result = convert_order_file(ROOT / "raw_order.xlsx", tmp_path / "out")
    assert result.export_status == "ok"
    wb = openpyxl.load_workbook(result.dx_path)
    ws = wb.active
    assert ws.sheet_format.defaultRowHeight == 25.0
    assert ws.max_row >= 2
    for r in range(2, min(ws.max_row + 1, 6)):
        h = ws.row_dimensions[r].height
        assert h is not None and abs(h - 25.0) <= 1.0, (r, h)
        cell = ws.cell(r, 7)  # 物流方式
        assert cell.alignment.wrap_text is True, (r, cell.alignment.wrap_text)
        assert cell.alignment.horizontal == "center"
        assert cell.alignment.vertical == "center"


def test_d01_d02_data_row_height_and_wrap(tmp_path):
    """店管家数据行：换行开启；01 行高 540；02 行高 ≥1080（样板分级）。"""
    import xlrd

    result = convert_order_file(ROOT / "raw_order.xlsx", tmp_path / "out")
    assert result.export_status == "ok"

    rb1 = xlrd.open_workbook(str(result.d01_path), formatting_info=True)
    sh1 = rb1.sheet_by_index(0)
    assert sh1.nrows >= 2
    for r in range(1, sh1.nrows):
        if not any(str(sh1.cell_value(r, c)).strip() for c in range(sh1.ncols)):
            continue
        ri = sh1.rowinfo_map.get(r)
        assert ri is not None and ri.height == 540, (r, None if ri is None else ri.height)
        # 有内容列应 text_wrapped
        for c in range(min(sh1.ncols, 8)):
            if str(sh1.cell_value(r, c)).strip() == "":
                continue
            xf = rb1.xf_list[sh1.cell(r, c).xf_index]
            assert xf.alignment.text_wrapped == 1, (r, c)
            break
        break

    rb2 = xlrd.open_workbook(str(result.d02_path), formatting_info=True)
    sh2 = rb2.sheet_by_index(0)
    assert sh2.nrows >= 2
    seen = False
    for r in range(1, sh2.nrows):
        if not any(str(sh2.cell_value(r, c)).strip() for c in range(sh2.ncols)):
            continue
        ri = sh2.rowinfo_map.get(r)
        assert ri is not None and ri.height >= 1080, (r, None if ri is None else ri.height)
        assert ri.height in (1080, 1350, 2430) or ri.height >= 1080
        for c in range(min(sh2.ncols, 10)):
            if str(sh2.cell_value(r, c)).strip() == "":
                continue
            xf = rb2.xf_list[sh2.cell(r, c).xf_index]
            assert xf.alignment.text_wrapped == 1, (r, c)
            seen = True
            break
        if seen:
            break
    assert seen


def test_unknown_shop_blocks_formal_d01_d02(tmp_path):
    def mut(ws, headers):
        ws.cell(2, _col(headers, "店铺账号"), "未知测试店XYZ")
    raw = _clone_raw(tmp_path, mut)
    out = tmp_path / "out"
    result = convert_order_file(raw, out, exports=("dx", "d01", "d02"))
    assert result.export_status == "blocked"
    assert result.formal_export_blocked is True
    assert any(a["type"] == "UNKNOWN_SHOP" and a["severity"] == "error" for a in result.errors)
    assert not (out / "dian_01.xls").exists()
    assert not (out / "dian_02.xls").exists()
    assert (out / "dianxiaomi.xlsx").exists()  # blocked 时店小秘预览仍可
    assert (out / "anomalies.json").exists()


def test_unknown_shop_review_export_not_formal(tmp_path):
    def mut(ws, headers):
        ws.cell(2, _col(headers, "店铺账号"), "未知测试店XYZ")
    raw = _clone_raw(tmp_path, mut)
    out = tmp_path / "out"
    result = convert_order_file(
        raw, out, exports=("d01", "d02"), allow_blocked_review_export=True
    )
    assert result.formal_export_blocked
    assert result.export_status == "blocked"
    assert not (out / "dian_01.xls").exists()
    assert (out / "review" / "dian_01_NEEDS_REVIEW.xls").exists()
    assert (out / "review" / "dian_02_NEEDS_REVIEW.xls").exists()


def test_ebay_order_no_various_lengths():
    rules = load_rules()
    cases = [
        ("2447420-43933", "43933"),
        ("12-3", "3"),
        ("onlysegment", "onlysegment"),
        ("a-b-c-d", "d"),
    ]
    for ono, exp in cases:
        raw = {"订单号": ono, "店铺账号": "ebay宇星店德国（自营）"}
        assert d01_original_order_no(raw, rules) == exp
    raw = {"订单号": "111-222", "店铺账号": "Shirley（自营）不贴logo"}
    assert d01_original_order_no(raw, rules) == "111-222"


def test_same_sku_multiple_orders(tmp_path):
    """同一 SKU 多订单应都能转换（中国直发多单本就同 SKU）。"""
    result = convert_order_file(ROOT / "raw_order.xlsx", tmp_path / "out")
    m16 = [m for m in result.mappings if m["dx_SKU"] == "M00016"]
    assert len(m16) == 27
    assert result.export_status == "ok"
    assert result.formal_export_blocked is False


def test_empty_orders(tmp_path):
    dst = tmp_path / "empty.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    src = openpyxl.load_workbook(ROOT / "raw_order.xlsx")
    headers = [src.active.cell(1, c).value for c in range(1, src.active.max_column + 1)]
    for c, h in enumerate(headers, 1):
        ws.cell(1, c, h)
    wb.save(dst)
    out = tmp_path / "out"
    result = convert_order_file(dst, out)
    assert result.export_status == "blocked"
    assert any(a["type"] == "EMPTY_INPUT" for a in result.errors)
    assert not (out / "dian_01.xls").exists()


def test_empty_logistics_is_error(tmp_path):
    def mut(ws, headers):
        ws.cell(2, _col(headers, "物流方式"), "")
    raw = _clone_raw(tmp_path, mut)
    result = convert_order_file(raw, tmp_path / "out")
    assert any(a["type"] == "SORT_KEY_EMPTY" for a in result.errors)
    assert result.export_status == "blocked"


def test_missing_required_column(tmp_path):
    dst = tmp_path / "miss.xlsx"
    copyfile(ROOT / "raw_order.xlsx", dst)
    wb = openpyxl.load_workbook(dst)
    ws = wb.active
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    c = headers.index("运单号") + 1
    ws.cell(1, c, "运单号_REMOVED")
    wb.save(dst)
    result = convert_order_file(dst, tmp_path / "out")
    assert any(a["type"] == "MISSING_COLUMNS" for a in result.errors)
    assert result.export_status == "blocked"
    assert not (tmp_path / "out" / "dianxiaomi.xlsx").exists()


def test_duplicate_order_no(tmp_path):
    """原始订单号重复：警告，不阻断；由 D02 去重处理。"""
    def mut(ws, headers):
        ono_c = _col(headers, "订单号")
        ws.cell(3, ono_c, ws.cell(2, ono_c).value)
    raw = _clone_raw(tmp_path, mut)
    result = convert_order_file(raw, tmp_path / "out")
    assert any(a["type"] == "DUPLICATE_ORDER_NO" for a in result.warnings)
    assert not any(a["type"] == "DUPLICATE_ORDER_NO" for a in result.errors)
    assert result.export_status != "blocked"


def test_duplicate_tracking(tmp_path):
    def mut(ws, headers):
        tc = _col(headers, "运单号")
        ws.cell(3, tc, ws.cell(2, tc).value)
    raw = _clone_raw(tmp_path, mut)
    result = convert_order_file(raw, tmp_path / "out")
    assert any(a["type"] == "DUPLICATE_TRACKING" for a in result.errors)


def test_invalid_quantity_and_amount(tmp_path):
    def mut(ws, headers):
        ws.cell(2, _col(headers, "产品总数"), "不是数字")
        ws.cell(2, _col(headers, "订单金额"), "xx")
        ws.cell(2, _col(headers, "拣货备注"), "")
    raw = _clone_raw(tmp_path, mut)
    result = convert_order_file(raw, tmp_path / "out")
    types = {a["type"] for a in result.errors}
    assert "INVALID_QUANTITY" in types
    assert "INVALID_AMOUNT" in types
    assert result.export_status == "blocked"


def test_only_dx_does_not_require_writing_d01_d02(tmp_path):
    out = tmp_path / "out"
    result = convert_order_file(ROOT / "raw_order.xlsx", out, exports=("dx",))
    assert result.dx_path and result.dx_path.exists()
    assert result.d01_path is None
    assert result.d02_path is None
    assert not (out / "dian_01.xls").exists()
    assert not (out / "dian_02.xls").exists()


def test_only_d01_with_blocker_writes_nothing_formal(tmp_path):
    def mut(ws, headers):
        ws.cell(2, _col(headers, "店铺账号"), "未知店")
    raw = _clone_raw(tmp_path, mut)
    out = tmp_path / "out"
    result = convert_order_file(raw, out, exports=("d01",))
    assert result.export_status == "blocked"
    assert not (out / "dian_01.xls").exists()


def test_is_sku_like_helper():
    assert is_sku_like("WX17123") is True
    assert is_sku_like("M0006S") is True
    assert is_sku_like("QDD1") is True
    assert is_sku_like("  WX17123  ") is True
    assert is_sku_like("QDD") is False  # 无数字
    assert is_sku_like("12345") is False  # 无字母
    assert is_sku_like("贴日本标签") is False
    assert is_sku_like("中国直发M00016 海外仓发产品") is False
    assert is_sku_like("WX17123 备用") is False
    assert is_sku_like("") is False
    assert is_sku_like("ab") is False  # 太短


def test_sku_from_pick_note_replaces_dx_and_d02(tmp_path):
    """拣货备注=WX17123 替换原 SKU；店小秘与店管家02编号/条码均为 WX17123。"""
    rows = [
        {
            "订单号": "SYNTH-NOTE-SKU-001",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "WX17023-SHUU",
            "产品总数": 1,
            "拣货备注": "WX17123",
            "物流方式": "云途全球专线挂号",
            "运单号": "TRACK-NOTE-SKU-001",
            "订单金额": 12.3,
        }
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    out = tmp_path / "out"
    result = convert_order_file(raw, out)
    infos = [a for a in result.infos if a["type"] == "SKU_FROM_NOTE"]
    assert len(infos) == 1
    assert infos[0]["原始SKU"] == "WX17023-SHUU"
    assert infos[0]["dx_SKU"] == "WX17123"
    assert infos[0]["source_column"] == "拣货备注"
    assert result.export_status == "ok"
    assert not any(a["type"] == "UNKNOWN_SKU_SUFFIX" for a in result.warnings)
    m = result.mappings[0]
    assert m["dx_SKU"] == "WX17123"
    assert m["base_sku"] == "WX17123"
    assert m["sku_from_pick_note"] is True
    assert m["拣货备注"] == "WX17123"
    wb = openpyxl.load_workbook(result.dx_path)
    headers = [wb.active.cell(1, c).value for c in range(1, wb.active.max_column + 1)]
    sku_c = headers.index("SKU") + 1
    note_c = headers.index("拣货备注") + 1
    assert wb.active.cell(2, sku_c).value == "WX17123"
    assert wb.active.cell(2, note_c).value == "WX17123"
    import xlrd
    sh = xlrd.open_workbook(str(result.d02_path)).sheet_by_index(0)
    h2 = [sh.cell_value(0, c) for c in range(sh.ncols)]
    assert sh.cell_value(1, h2.index("编号")) == "WX17123"
    assert sh.cell_value(1, h2.index("条码")) == "WX17123"


def test_china_direct_not_overridden_by_sku_like_note(tmp_path):
    """中国直发完整标记优先；备注不是 SKU-like，走 M00016。"""
    marker = load_rules().china_marker
    rows = [
        {
            "订单号": "SYNTH-NOTE-CHINA-001",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "WX99999-LYF",
            "产品总数": 1,
            "拣货备注": marker,
            "物流方式": "燕文航空挂号",
            "运单号": "TRACK-NOTE-CHINA-001",
            "订单金额": 5,
        }
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    result = convert_order_file(raw, tmp_path / "out")
    assert not any(a["type"] == "SKU_FROM_NOTE" for a in result.anomalies)
    m = result.mappings[0]
    assert m["is_china_direct"] is True
    assert m["dx_SKU"] == "M00016"
    assert m["base_sku"] == "M00016"
    assert m["sku_from_pick_note"] is False


def test_empty_pick_note_keeps_normal_sku_strip(tmp_path):
    """普通无备注：不去改 SKU，仅去后缀。"""
    rows = [
        {
            "订单号": "SYNTH-NOTE-NONE-001",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "WX10001-LYF",
            "产品总数": 1,
            "拣货备注": "",
            "物流方式": "云途全球专线挂号",
            "运单号": "TRACK-NOTE-NONE-001",
            "订单金额": 3,
        }
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    result = convert_order_file(raw, tmp_path / "out")
    assert result.export_status == "ok"
    assert not any(a["type"] == "SKU_FROM_NOTE" for a in result.anomalies)
    m = result.mappings[0]
    assert m["dx_SKU"] == "WX10001"
    assert m["sku_from_pick_note"] is False


def test_multi_sku_not_overridden_by_pick_note(tmp_path):
    """多 SKU 待复核；即使备注是单个 SKU-like 也不覆盖。"""
    rows = [
        {
            "订单号": "SYNTH-NOTE-MULTI-001",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": 'M0019-HKCQ\nM0003-CQCS',
            "产品总数": 3,
            "拣货备注": "WX17123",
            "物流方式": "云途全球专线挂号",
            "运单号": "TRACK-NOTE-MULTI-001",
            "订单金额": 10,
        }
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    result = convert_order_file(raw, tmp_path / "out")
    assert any(a["type"] == "MULTI_SKU" for a in result.warnings)
    assert not any(a["type"] == "SKU_FROM_NOTE" for a in result.anomalies)
    assert result.export_status == "needs_review"
    m = result.mappings[0]
    assert m["sku_from_pick_note"] is False
    assert m["dx_SKU"] == 'M0019\nM0003'
    assert m["base_sku"] == 'M0019\nM0003'


def test_pick_note_sku_with_suffix_strips_after_replace(tmp_path):
    """备注本身带白名单后缀时，替换后再去后缀。"""
    rows = [
        {
            "订单号": "SYNTH-NOTE-SUF-001",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "OLD999-XYTU",
            "产品总数": 1,
            "拣货备注": "WX20001-LYF",
            "物流方式": "云途A",
            "运单号": "TRACK-NOTE-SUF-001",
            "订单金额": 1,
        }
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    result = convert_order_file(raw, tmp_path / "out")
    assert result.export_status == "ok"
    infos = [a for a in result.infos if a["type"] == "SKU_FROM_NOTE"]
    assert len(infos) == 1
    assert infos[0]["source_column"] == "拣货备注"
    assert result.mappings[0]["dx_SKU"] == "WX20001"
    assert result.mappings[0]["base_sku"] == "WX20001"


def test_sku_alias_m00007_plain_and_suffixed(tmp_path):
    """M00007 / M00007-XYTU → 输出 M0007；店小秘与店管家02一致；其它 SKU 不变。"""
    rows = [
        {
            "订单号": "SYNTH-ALIAS-M00007-A",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "M00007-XYTU",
            "产品总数": 1,
            "拣货备注": "",
            "物流方式": "云途A",
            "运单号": "TRACK-ALIAS-M00007-A",
            "订单金额": 9,
        },
        {
            "订单号": "SYNTH-ALIAS-M00007-B",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "M00007",
            "产品总数": 1,
            "拣货备注": "",
            "物流方式": "云途B",
            "运单号": "TRACK-ALIAS-M00007-B",
            "订单金额": 8,
        },
        {
            "订单号": "SYNTH-ALIAS-OTHER",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "WX10001-LYF",
            "产品总数": 1,
            "拣货备注": "",
            "物流方式": "云途C",
            "运单号": "TRACK-ALIAS-OTHER",
            "订单金额": 7,
        },
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    result = convert_order_file(raw, tmp_path / "out")
    assert result.export_status == "ok"
    by_ono = {m["原始订单号"]: m for m in result.mappings}
    assert by_ono["SYNTH-ALIAS-M00007-A"]["dx_SKU"] == "M0007"
    assert by_ono["SYNTH-ALIAS-M00007-A"]["base_sku"] == "M0007"
    assert by_ono["SYNTH-ALIAS-M00007-B"]["dx_SKU"] == "M0007"
    assert by_ono["SYNTH-ALIAS-OTHER"]["dx_SKU"] == "WX10001"
    import xlrd
    sh = xlrd.open_workbook(str(result.d02_path)).sheet_by_index(0)
    h2 = [sh.cell_value(0, c) for c in range(sh.ncols)]
    codes = {sh.cell_value(r, h2.index("订单编号")): sh.cell_value(r, h2.index("编号")) for r in range(1, sh.nrows)}
    assert codes["SYNTH-ALIAS-M00007-A"] == "M0007"
    assert codes["SYNTH-ALIAS-M00007-B"] == "M0007"
    assert codes["SYNTH-ALIAS-OTHER"] == "WX10001"
    # 中国直发 M00016 不被别名误伤（配置无此键）
    assert load_rules().sku_aliases.get("M00016") is None
    assert load_rules().sku_aliases["M00007"] == "M0007"


def test_china_direct_m00016_not_aliased(tmp_path):
    marker = load_rules().china_marker
    rows = [
        {
            "订单号": "SYNTH-ALIAS-CHINA-16",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "M00007-XYTU",
            "产品总数": 1,
            "拣货备注": marker,
            "物流方式": "燕文A",
            "运单号": "TRACK-ALIAS-CHINA-16",
            "订单金额": 1,
        }
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    result = convert_order_file(raw, tmp_path / "out")
    m = result.mappings[0]
    assert m["is_china_direct"] is True
    assert m["dx_SKU"] == "M00016"
    assert m["base_sku"] == "M00016"



def test_sku_prefix_wxm_strips_wx_only(tmp_path):
    """WXM… 去 WX 得 M…；WX+数字不动；中国直发 M00016 优先。"""
    marker = load_rules().china_marker
    rows = [
        {
            "订单号": "SYNTH-WXM-PLAIN",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "WXM0021",
            "产品总数": 1,
            "拣货备注": "",
            "物流方式": "云途WXM1",
            "运单号": "TRACK-WXM-PLAIN",
            "订单金额": 1,
        },
        {
            "订单号": "SYNTH-WXM-SUF",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "WXM0021-XYTU",
            "产品总数": 1,
            "拣货备注": "",
            "物流方式": "云途WXM2",
            "运单号": "TRACK-WXM-SUF",
            "订单金额": 2,
        },
        {
            "订单号": "SYNTH-WX-DIGITS-SUF",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "WX10225-LYF",
            "产品总数": 1,
            "拣货备注": "",
            "物流方式": "云途WX1",
            "运单号": "TRACK-WX-DIGITS-SUF",
            "订单金额": 3,
        },
        {
            "订单号": "SYNTH-WX-DIGITS",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "WX21005",
            "产品总数": 1,
            "拣货备注": "",
            "物流方式": "云途WX2",
            "运单号": "TRACK-WX-DIGITS",
            "订单金额": 4,
        },
        {
            "订单号": "SYNTH-WXM-CHINA",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "WXM0021-XYTU",
            "产品总数": 1,
            "拣货备注": marker,
            "物流方式": "燕文WXM",
            "运单号": "TRACK-WXM-CHINA",
            "订单金额": 5,
        },
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    result = convert_order_file(raw, tmp_path / "out")
    assert result.export_status == "ok"
    by_ono = {m["原始订单号"]: m for m in result.mappings}
    assert by_ono["SYNTH-WXM-PLAIN"]["dx_SKU"] == "M0021"
    assert by_ono["SYNTH-WXM-PLAIN"]["base_sku"] == "M0021"
    assert by_ono["SYNTH-WXM-SUF"]["dx_SKU"] == "M0021"
    assert by_ono["SYNTH-WXM-SUF"]["base_sku"] == "M0021"
    assert by_ono["SYNTH-WX-DIGITS-SUF"]["dx_SKU"] == "WX10225"
    assert by_ono["SYNTH-WX-DIGITS-SUF"]["base_sku"] == "WX10225"
    assert by_ono["SYNTH-WX-DIGITS"]["dx_SKU"] == "WX21005"
    assert by_ono["SYNTH-WX-DIGITS"]["base_sku"] == "WX21005"
    assert by_ono["SYNTH-WXM-CHINA"]["dx_SKU"] == "M00016"
    assert by_ono["SYNTH-WXM-CHINA"]["base_sku"] == "M00016"
    strips = [a for a in result.infos if a["type"] == "SKU_STRIP_WXM_PREFIX"]
    befores = {a["before"] for a in strips}
    assert "WXM0021" in befores
    assert all(a["after"] == "M0021" for a in strips if a["before"] == "WXM0021")
    rules = load_rules()
    assert rules.sku_prefix_rules[0]["when_starts_with"] == "WXM"
    assert rules.sku_prefix_rules[0]["strip_prefix"] == "WX"



def test_qdx_appends_to_dx_logistics_and_d02_kefu(tmp_path):
    """基础 SKU=QDX（含去后缀）时：店小秘物流无空格追加 QDX；02客服备注无空格追加；02物流仍云途。"""
    rows = [
        {
            "订单号": "SYNTH-QDX-SUF-001",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "QDX-XYTU",
            "产品总数": 1,
            "拣货备注": "",
            "物流方式": "宇星科技-FBM SHIP+云途特快（普货）",
            "运单号": "TRACK-QDX-SUF-001",
            "订单金额": 11,
        },
        {
            "订单号": "SYNTH-QDX-PLAIN-001",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "QDX",
            "产品总数": 2,
            "拣货备注": "",
            "物流方式": "速卖通线上-无忧物流-标准",
            "运单号": "TRACK-QDX-PLAIN-001",
            "订单金额": 12,
        },
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    result = convert_order_file(raw, tmp_path / "out")
    assert result.export_status == "ok"
    by = {m["原始订单号"]: m for m in result.mappings}
    assert by["SYNTH-QDX-SUF-001"]["dx_SKU"] == "QDX"
    assert by["SYNTH-QDX-SUF-001"]["base_sku"] == "QDX"
    assert by["SYNTH-QDX-SUF-001"]["dx_物流方式"] == "宇星科技-FBM SHIP+云途特快（普货）QDX"
    assert by["SYNTH-QDX-PLAIN-001"]["dx_物流方式"] == "速卖通线上-无忧物流-标准QDX"
    wb = openpyxl.load_workbook(result.dx_path)
    headers = [wb.active.cell(1, c).value for c in range(1, wb.active.max_column + 1)]
    log_c = headers.index("物流方式") + 1
    sku_c = headers.index("SKU") + 1
    ono_c = headers.index("订单号") + 1
    for r in range(2, wb.active.max_row + 1):
        ono = wb.active.cell(r, ono_c).value
        if ono == "SYNTH-QDX-SUF-001":
            assert wb.active.cell(r, sku_c).value == "QDX"
            assert wb.active.cell(r, log_c).value.endswith("QDX")
            assert " QDX" not in str(wb.active.cell(r, log_c).value)
    import xlrd
    sh = xlrd.open_workbook(str(result.d02_path)).sheet_by_index(0)
    h2 = [sh.cell_value(0, c) for c in range(sh.ncols)]
    for r in range(1, sh.nrows):
        ono = sh.cell_value(r, h2.index("订单编号"))
        if ono == "SYNTH-QDX-SUF-001":
            assert sh.cell_value(r, h2.index("编号")) == "QDX"
            assert sh.cell_value(r, h2.index("条码")) == "QDX"
            assert sh.cell_value(r, h2.index("物流方式")) == "云途"
            assert sh.cell_value(r, h2.index("客服备注")) == "宇星科技-FBM SHIP+云途特快（普货）QDX"
        if ono == "SYNTH-QDX-PLAIN-001":
            assert sh.cell_value(r, h2.index("客服备注")) == "速卖通线上-无忧物流-标准QDX"


def test_qdx_does_not_override_china_direct(tmp_path):
    """中国直发优先：SKU 即使是 QDX 也不走 QDX 追加，物流接中国直发标记。"""
    marker = load_rules().china_marker
    rows = [
        {
            "订单号": "SYNTH-QDX-CHINA-001",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "QDX-XYTU",
            "产品总数": 1,
            "拣货备注": marker,
            "物流方式": "燕文测试物流",
            "运单号": "TRACK-QDX-CHINA-001",
            "订单金额": 3,
        }
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    result = convert_order_file(raw, tmp_path / "out")
    m = result.mappings[0]
    assert m["is_china_direct"] is True
    assert m["dx_SKU"] == "M00016"
    assert m["dx_物流方式"] == "燕文测试物流" + marker
    assert not m["dx_物流方式"].endswith("QDX")


def test_iterative_strip_qhde_fbm_keeps_numeric_segment(tmp_path):
    """18MB-020-QHDE-FBM / 18MB-020-QHDE → 18MB-020；不剥成 18MB；白名单无 UNKNOWN。"""
    rows = [
        {
            "订单号": "SYNTH-18MB-FBM-001",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "18MB-020-QHDE-FBM",
            "产品总数": 1,
            "拣货备注": "",
            "物流方式": "云途全球专线挂号",
            "运单号": "TRACK-18MB-FBM-001",
            "订单金额": 1,
        },
        {
            "订单号": "SYNTH-18MB-QHDE-001",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "18MB-020-QHDE",
            "产品总数": 1,
            "拣货备注": "",
            "物流方式": "云途全球专线挂号",
            "运单号": "TRACK-18MB-QHDE-001",
            "订单金额": 1,
        },
        {
            "订单号": "SYNTH-18MB-BASE-001",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "18MB-020",
            "产品总数": 1,
            "拣货备注": "",
            "物流方式": "云途全球专线挂号",
            "运单号": "TRACK-18MB-BASE-001",
            "订单金额": 1,
        },
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    out = tmp_path / "out"
    result = convert_order_file(raw, out)
    by = {m["原始订单号"]: m for m in result.mappings}
    assert by["SYNTH-18MB-FBM-001"]["dx_SKU"] == "18MB-020"
    assert by["SYNTH-18MB-FBM-001"]["base_sku"] == "18MB-020"
    assert by["SYNTH-18MB-QHDE-001"]["dx_SKU"] == "18MB-020"
    assert by["SYNTH-18MB-BASE-001"]["dx_SKU"] == "18MB-020"
    # 白名单剥离，不应因 QHDE/FBM 记 UNKNOWN
    unk = [
        a
        for a in result.warnings
        if a["type"] == "UNKNOWN_SKU_SUFFIX"
        and a["订单号"].startswith("SYNTH-18MB")
    ]
    assert unk == []
    # 店管家02 编号/条码
    import xlrd

    review = out / "review" / "dian_02_NEEDS_REVIEW.xls"
    # 本合成表无未知后缀时可能正式导出；若 needs_review 则看 review
    path02 = out / "dian_02.xls"
    if not path02.exists():
        path02 = review
    assert path02.exists()
    sh = xlrd.open_workbook(path02, formatting_info=False).sheet_by_index(0)
    h2 = [sh.cell_value(0, c) for c in range(sh.ncols)]
    for r in range(1, sh.nrows):
        ono = sh.cell_value(r, h2.index("订单编号"))
        if ono in ("SYNTH-18MB-FBM-001", "SYNTH-18MB-QHDE-001", "SYNTH-18MB-BASE-001"):
            assert sh.cell_value(r, h2.index("编号")) == "18MB-020"
            assert sh.cell_value(r, h2.index("条码")) == "18MB-020"


def test_iterative_strip_unknown_then_alias(tmp_path):
    """WX17023-SHUU → WX17023 + UNKNOWN；M00007-XYTU → M0007。"""
    rows = [
        {
            "订单号": "SYNTH-ITER-SHUU-001",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "WX17023-SHUU",
            "产品总数": 1,
            "拣货备注": "",
            "物流方式": "云途A",
            "运单号": "TRACK-ITER-SHUU",
            "订单金额": 1,
        },
        {
            "订单号": "SYNTH-ITER-ALIAS-001",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "M00007-XYTU",
            "产品总数": 1,
            "拣货备注": "",
            "物流方式": "云途B",
            "运单号": "TRACK-ITER-ALIAS",
            "订单金额": 1,
        },
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    result = convert_order_file(raw, tmp_path / "out")
    by = {m["原始订单号"]: m for m in result.mappings}
    assert by["SYNTH-ITER-SHUU-001"]["dx_SKU"] == "WX17023"
    assert any(
        a["type"] == "UNKNOWN_SKU_SUFFIX" and a.get("suffix") == "SHUU"
        for a in result.warnings
    )
    assert by["SYNTH-ITER-ALIAS-001"]["dx_SKU"] == "M0007"
    assert by["SYNTH-ITER-ALIAS-001"]["base_sku"] == "M0007"
    assert "QHDE" in load_rules().strip_suffixes
    assert "FBM" in load_rules().strip_suffixes


def test_sku_from_kefu_note_column(tmp_path):
    """货号在客服备注：拣货备注空时用客服备注替换。"""
    rows = [
        {
            "订单号": "SYNTH-NOTE-KEFU-001",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "OLD-XYTU",
            "产品总数": 1,
            "拣货备注": "",
            "客服备注": "WX30001",
            "订单备注": "",
            "物流方式": "云途A",
            "运单号": "TRACK-NOTE-KEFU-001",
            "订单金额": 1,
            "币种缩写": "EUR",
        }
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    result = convert_order_file(raw, tmp_path / "out")
    infos = [a for a in result.infos if a["type"] == "SKU_FROM_NOTE"]
    assert len(infos) == 1
    assert infos[0]["source_column"] == "客服备注"
    assert result.mappings[0]["dx_SKU"] == "WX30001"
    assert result.mappings[0]["sku_from_note_column"] == "客服备注"
    assert result.export_status == "ok"


def test_sku_from_order_note_column(tmp_path):
    """货号仅在订单备注。"""
    rows = [
        {
            "订单号": "SYNTH-NOTE-ORD-001",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "OLD-XYTU",
            "产品总数": 1,
            "拣货备注": "",
            "客服备注": "",
            "订单备注": "WX30002",
            "物流方式": "云途A",
            "运单号": "TRACK-NOTE-ORD-001",
            "订单金额": 1,
            "币种缩写": "EUR",
        }
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    result = convert_order_file(raw, tmp_path / "out")
    infos = [a for a in result.infos if a["type"] == "SKU_FROM_NOTE"]
    assert len(infos) == 1
    assert infos[0]["source_column"] == "订单备注"
    assert result.mappings[0]["dx_SKU"] == "WX30002"


def test_sku_from_note_priority_pick_over_kefu_and_order(tmp_path):
    """三列都有货号时只取优先级最高的拣货备注。"""
    rows = [
        {
            "订单号": "SYNTH-NOTE-PRIO-001",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "OLD-XYTU",
            "产品总数": 1,
            "拣货备注": "WX11111",
            "客服备注": "WX22222",
            "订单备注": "WX33333",
            "物流方式": "云途A",
            "运单号": "TRACK-NOTE-PRIO-001",
            "订单金额": 1,
            "币种缩写": "EUR",
        }
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    result = convert_order_file(raw, tmp_path / "out")
    infos = [a for a in result.infos if a["type"] == "SKU_FROM_NOTE"]
    assert len(infos) == 1
    assert infos[0]["source_column"] == "拣货备注"
    assert infos[0]["note_value"] == "WX11111"
    assert result.mappings[0]["dx_SKU"] == "WX11111"
    assert result.mappings[0]["dx_SKU"] != "WX22222"


def test_currency_append_usd_gbp_skip_eur(tmp_path):
    """USD/GBP 接到物流后；EUR/欧元不接；与 QDX、中国直发共存不丢前缀。"""
    marker = load_rules().china_marker
    rows = [
        {
            "订单号": "SYNTH-CUR-USD-001",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "WX1-LYF",
            "产品总数": 1,
            "拣货备注": "",
            "物流方式": "云途特快",
            "运单号": "TRACK-CUR-USD",
            "订单金额": 1,
            "币种缩写": "USD",
        },
        {
            "订单号": "SYNTH-CUR-GBP-001",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "WX2-LYF",
            "产品总数": 1,
            "拣货备注": "",
            "物流方式": "燕文挂号",
            "运单号": "TRACK-CUR-GBP",
            "订单金额": 1,
            "币种缩写": "GBP",
        },
        {
            "订单号": "SYNTH-CUR-EUR-001",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "WX3-LYF",
            "产品总数": 1,
            "拣货备注": "",
            "物流方式": "云途标准",
            "运单号": "TRACK-CUR-EUR",
            "订单金额": 1,
            "币种缩写": "EUR",
        },
        {
            "订单号": "SYNTH-CUR-CN-EUR-001",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "WX4-LYF",
            "产品总数": 1,
            "拣货备注": "",
            "物流方式": "云途标准2",
            "运单号": "TRACK-CUR-CN-EUR",
            "订单金额": 1,
            "币种缩写": "欧元",
        },
        {
            "订单号": "SYNTH-CUR-QDX-USD-001",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "QDX-XYTU",
            "产品总数": 1,
            "拣货备注": "",
            "物流方式": "宇星科技-FBM SHIP",
            "运单号": "TRACK-CUR-QDX-USD",
            "订单金额": 1,
            "币种缩写": "USD",
        },
        {
            "订单号": "SYNTH-CUR-CHINA-USD-001",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "ANY-LYF",
            "产品总数": 1,
            "拣货备注": marker,
            "物流方式": "燕文测试",
            "运单号": "TRACK-CUR-CHINA-USD",
            "订单金额": 1,
            "币种缩写": "USD",
        },
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    result = convert_order_file(raw, tmp_path / "out")
    by = {m["原始订单号"]: m for m in result.mappings}
    assert by["SYNTH-CUR-USD-001"]["dx_物流方式"] == "云途特快USD"
    assert by["SYNTH-CUR-GBP-001"]["dx_物流方式"] == "燕文挂号GBP"
    assert by["SYNTH-CUR-EUR-001"]["dx_物流方式"] == "云途标准"
    assert by["SYNTH-CUR-CN-EUR-001"]["dx_物流方式"] == "云途标准2"
    assert by["SYNTH-CUR-QDX-USD-001"]["dx_物流方式"] == "宇星科技-FBM SHIPQDXUSD"
    assert by["SYNTH-CUR-CHINA-USD-001"]["dx_物流方式"] == "燕文测试" + marker + "USD"
    import xlrd
    sh = xlrd.open_workbook(str(result.d02_path)).sheet_by_index(0)
    h2 = [sh.cell_value(0, c) for c in range(sh.ncols)]
    for r in range(1, sh.nrows):
        ono = sh.cell_value(r, h2.index("订单编号"))
        kefu = sh.cell_value(r, h2.index("客服备注"))
        logcol = sh.cell_value(r, h2.index("物流方式"))
        assert logcol == "云途"
        if ono == "SYNTH-CUR-USD-001":
            assert kefu == "云途特快USD"
        if ono == "SYNTH-CUR-QDX-USD-001":
            assert kefu == "宇星科技-FBM SHIPQDXUSD"
        if ono == "SYNTH-CUR-CHINA-USD-001":
            assert kefu == "燕文测试 " + marker + "USD"
        if ono == "SYNTH-CUR-EUR-001":
            assert kefu == "云途标准"


def test_d02_duplicate_order_no_gets_suffix(tmp_path):
    """两条同订单号进 D02：第一条原号，第二条 原号-1；店小秘不去重。"""
    rows = [
        {
            "订单号": "SYNTH-DEDUP-A",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "WX10001-XYTU",
            "物流方式": "云途特快",
            "拣货备注": "",
            "运单号": "TRACK-DEDUP-A1",
            "产品总数": 1,
            "订单金额": 10,
            "收货人姓名": "Alice",
            "收货人国家": "US",
            "详细地址": "Addr A",
            "币种": "EUR",
        },
        {
            "订单号": "SYNTH-DEDUP-A",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "WX10002-XYTU",
            "物流方式": "云途特快",
            "拣货备注": "",
            "运单号": "TRACK-DEDUP-A2",
            "产品总数": 1,
            "订单金额": 11,
            "收货人姓名": "Bob",
            "收货人国家": "US",
            "详细地址": "Addr B",
            "币种": "EUR",
        },
        {
            "订单号": "SYNTH-DEDUP-B",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "WX10003-XYTU",
            "物流方式": "燕文挂号",
            "拣货备注": "",
            "运单号": "TRACK-DEDUP-B1",
            "产品总数": 1,
            "订单金额": 12,
            "收货人姓名": "Carol",
            "收货人国家": "US",
            "详细地址": "Addr C",
            "币种": "EUR",
        },
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    result = convert_order_file(raw, tmp_path / "out")
    assert result.export_status != "blocked"
    assert result.stats["d02_orders"] == 3
    dx_onos = [m["dx_订单号"] for m in result.mappings if m["原始订单号"] == "SYNTH-DEDUP-A"]
    assert dx_onos == ["SYNTH-DEDUP-A", "SYNTH-DEDUP-A"]
    d02_nos = sorted(
        m["d02_订单编号"] for m in result.mappings if m["原始订单号"] == "SYNTH-DEDUP-A"
    )
    assert d02_nos == ["SYNTH-DEDUP-A", "SYNTH-DEDUP-A-1"]
    assert any(a["type"] == "D02_ORDER_NO_DEDUP" for a in result.infos)
    assert any(a["type"] == "DUPLICATE_ORDER_NO" for a in result.warnings)
    import xlrd
    sh = xlrd.open_workbook(str(result.d02_path)).sheet_by_index(0)
    h2 = [sh.cell_value(0, c) for c in range(sh.ncols)]
    col = h2.index("订单编号")
    written = [sh.cell_value(r, col) for r in range(1, sh.nrows)]
    assert written.count("SYNTH-DEDUP-A") == 1
    assert "SYNTH-DEDUP-A-1" in written
    assert "SYNTH-DEDUP-B" in written
    assert len(written) == len(set(written))


def test_d02_dedup_three_same_order_no(tmp_path):
    """三条同号：原号 / -1 / -2。"""
    rows = [
        {
            "订单号": "SYNTH-DEDUP3",
            "店铺账号": "亚马逊德国店（测试）",
            "SKU": f"SKU{i}-XYTU",
            "物流方式": "云途",
            "拣货备注": "",
            "运单号": f"TRACK-DEDUP3-{i}",
            "产品总数": 1,
            "订单金额": 9 + i,
            "收货人姓名": f"P{i}",
            "收货人国家": "DE",
            "详细地址": f"Addr{i}",
            "币种": "EUR",
        }
        for i in range(1, 4)
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    result = convert_order_file(raw, tmp_path / "out")
    nos = [m["d02_订单编号"] for m in result.mappings]
    assert set(nos) == {"SYNTH-DEDUP3", "SYNTH-DEDUP3-1", "SYNTH-DEDUP3-2"}
    assert result.stats["d02_order_no_dedup"] == 2


def test_d02_dedup_china_suffix_collision(tmp_path):
    """中国直发已带 -1 的同号两条：第一条保 -1，第二条为 -2；店小秘仍都是 原号-1。"""
    marker = load_rules().china_marker
    rows = [
        {
            "订单号": "SYNTH-CN-DEDUP",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "M0003-CQCS",
            "物流方式": "燕文",
            "拣货备注": marker,
            "运单号": "TRACK-CN-DEDUP-1",
            "产品总数": 1,
            "订单金额": 20,
            "收货人姓名": "N1",
            "收货人国家": "US",
            "详细地址": "A1",
            "币种": "EUR",
        },
        {
            "订单号": "SYNTH-CN-DEDUP",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "M0003-CQCS",
            "物流方式": "燕文",
            "拣货备注": marker,
            "运单号": "TRACK-CN-DEDUP-2",
            "产品总数": 1,
            "订单金额": 21,
            "收货人姓名": "N2",
            "收货人国家": "US",
            "详细地址": "A2",
            "币种": "EUR",
        },
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    result = convert_order_file(raw, tmp_path / "out")
    dx = [m["dx_订单号"] for m in result.mappings]
    assert dx == ["SYNTH-CN-DEDUP-1", "SYNTH-CN-DEDUP-1"]
    d02 = sorted(m["d02_订单编号"] for m in result.mappings)
    assert d02 == ["SYNTH-CN-DEDUP-1", "SYNTH-CN-DEDUP-2"]
    assert result.export_status != "blocked"
    assert any(a["type"] == "D02_ORDER_NO_DEDUP" for a in result.infos)


def test_d02_dedup_avoids_existing_minus_one(tmp_path):
    """同组普通两条去重占用 -1 后，同原始单号的中国直发行（已是 基号-1）需再递增为 -2。"""
    marker = load_rules().china_marker
    rows = [
        {
            "订单号": "SYNTH-MIX-BASE",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "WX20001-XYTU",
            "物流方式": "云途",
            "拣货备注": "",
            "运单号": "TRACK-MIX-1",
            "产品总数": 1,
            "订单金额": 10,
            "收货人姓名": "M1",
            "收货人国家": "US",
            "详细地址": "X1",
            "币种": "EUR",
        },
        {
            "订单号": "SYNTH-MIX-BASE",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "WX20002-XYTU",
            "物流方式": "云途",
            "拣货备注": "",
            "运单号": "TRACK-MIX-2",
            "产品总数": 1,
            "订单金额": 11,
            "收货人姓名": "M2",
            "收货人国家": "US",
            "详细地址": "X2",
            "币种": "EUR",
        },
        {
            "订单号": "SYNTH-MIX-BASE",
            "店铺账号": "亚马逊美国店（测试）",
            "SKU": "M0003-CQCS",
            "物流方式": "燕文",
            "拣货备注": marker,
            "运单号": "TRACK-MIX-3",
            "产品总数": 1,
            "订单金额": 12,
            "收货人姓名": "M3",
            "收货人国家": "US",
            "详细地址": "X3",
            "币种": "EUR",
        },
    ]
    raw = _make_minimal_raw(tmp_path, rows)
    result = convert_order_file(raw, tmp_path / "out")
    d02 = [m["d02_订单编号"] for m in result.mappings]
    assert len(d02) == len(set(d02))
    assert "SYNTH-MIX-BASE" in d02
    assert "SYNTH-MIX-BASE-1" in d02
    assert "SYNTH-MIX-BASE-2" in d02
