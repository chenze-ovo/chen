"""项目2规则验收：Amazon 正本、中国直发金额、备注、SKU 映射、eBay 分单、行高、拣货备注。"""
from __future__ import annotations

import re
import sys
import zipfile
from decimal import Decimal
from pathlib import Path

import yaml
from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from project2_engine.pipeline import (  # noqa: E402
    OVERSEA_ROW_HEIGHT,
    _append_logistics,
    _d02_total_amount,
    _dian_02_values,
    _exclude_ebay_split_from_dian_01,
    _map_warehouse_logistics,
    _partition,
    _prepare_row,
    process_orders,
)


def _load_config() -> dict:
    return yaml.safe_load((ROOT / "routing.yaml").read_text())


def _cfg() -> dict:
    cfg = _load_config()
    cfg["identical_refs"] = {}  # 测试不依赖本机金标准壳
    return cfg


def _raw(**cells) -> dict:
    base = {
        "订单号": "X-1",
        "店铺账号": "亚马逊测试店",
        "平台": "Amazon",
        "SKU": "WX10001",
        "物流方式": "谷仓仓储-鑫赢科技-捷克仓-DHL德国派送",
        "运单号": "TRK1",
        "包裹号": "XM1",
        "拣货备注": "",
        "客服备注": "",
        "订单备注": "",
        "订单金额": 10,
        "币种缩写": "EUR",
        "产品总数": "1",
        "收货人国家": "DE",
        "收货人姓名": "Ada",
    }
    base.update(cells)
    return {"cells": base, "row_mark": "white"}


def test_amazon_oversea_swaps_tracking_and_d02_maps_source_logistics():
    cfg = _cfg()
    source_logistics = "Amazon-亚马逊陈艾生德国（自营）-德国站-Expedited-加急配送"
    raw = _raw(
        订单号="AMZ-1",
        物流方式=source_logistics,
        运单号="960D21340391F",
        包裹号="XM68AXJ07979",
        币种缩写="EUR",
    )
    row = _prepare_row(raw, cfg, [], None, set(), {}, [])
    assert row["cells"]["运单号"] == "XM68AXJ07979"
    assert row["cells"]["物流方式"].endswith("960D21340391F")
    assert source_logistics in row["cells"]["物流方式"]
    assert row["raw_logistics"] == source_logistics
    assert row["tracking_nos"] == ["960D21340391F"]
    assert row["ops_cells"]["运单号"] == "960D21340391F"
    assert row["ops_cells"]["物流方式"] == source_logistics

    values = _dian_02_values(
        row, cfg, {}, custom1="26113001", anomalies_out=[], order_counts={"AMZ-1": 1}
    )
    assert values["物流方式"] == "亚马逊FBA德国"
    assert values["客户备注"] == "960D21340391F"
    assert values["物流单号"] == "960D21340391F"


def test_amazon_skips_append_when_tracking_equals_package():
    cfg = _cfg()
    raw = _raw(
        物流方式="Amazon FBA-亚马逊香港创趣美国（自营）-美国站",
        运单号="XM68AXJ09011",
        包裹号="XM68AXJ09011",
        币种缩写="USD",
    )
    row = _prepare_row(raw, cfg, [], None, set(), {}, [])
    assert row["cells"]["运单号"] == "XM68AXJ09011"
    assert "XM68AXJ09011" not in (row["cells"]["物流方式"].replace(row["cells"]["包裹号"], ""))
    # 运单号等于包裹号时不往物流方式里再追加
    assert not row["cells"]["物流方式"].endswith("XM68AXJ09011") or row["cells"]["物流方式"].count("XM68AXJ09011") == 0


def test_china_direct_keeps_source_amount_not_price_or_zero_d02_backfills():
    cfg = _cfg()
    raw = _raw(
        订单号="303-9011940-1431562",
        拣货备注="中国直发M00016 海外仓发产品---亚马逊宇星欧洲（自营）",
        订单金额=18.8,
        订单价格=99.9,
        物流方式="谷仓仓储-深圳承思-捷克仓-CZ_GC_Parcel",
    )
    same_day = {"303-9011940-1431562": Decimal("36.5")}
    row = _prepare_row(raw, cfg, [], None, set(), same_day, [])
    assert row["china_direct"] is True
    assert row["amount_matched"] is True
    assert row["amount"] == Decimal("36.5")
    assert float(row["cells"]["订单金额"]) == 18.8
    assert row["cells"].get("订单价格") == 99.9
    total = _d02_total_amount(
        row, row["cells"], order_counts={"303-9011940-1431562": 1}, anomalies_out=[], order_no="303-9011940-1431562"
    )
    assert total == 36.5


def test_source_remarks_kept_on_oversea_row():
    cfg = _cfg()
    raw = _raw(
        物流方式="谷仓仓储-得业科技-捷克仓-DHL德国派送",
        拣货备注="千红捷克仓",
        订单备注="贴日本标签",
        客服备注="",
        币种缩写="EUR",
    )
    row = _prepare_row(raw, cfg, [], None, set(), {}, [])
    assert row["cells"]["拣货备注"] == "千红捷克仓"
    assert row["cells"]["订单备注"] == "贴日本标签"
    assert "千红捷克仓" in row["cells"]["物流方式"]


def test_special_sku_map_oversea_mapped_ops_keeps_raw():
    cfg = _cfg()
    raw = _raw(SKU="WX13111", 订单号="SKU-1")
    row = _prepare_row(raw, cfg, [], None, set(), {}, [])
    assert row["cells"]["SKU"] == "WX13199"
    assert row["skus"] == ["WX13199"]
    assert row["ops_cells"]["SKU"] == "WX13111"

    raw2 = _raw(SKU="WX10021-ZEE", 订单号="SKU-2")
    row2 = _prepare_row(raw2, cfg, [], None, set(), {}, [])
    assert row2["cells"]["SKU"] == "WX27020"
    assert row2["ops_cells"]["SKU"] == "WX10021-ZEE"


def test_ebay_split_note_24474201_skipped_from_dian_01():
    cfg = _cfg()
    raw = _raw(
        订单号="24474201-44270",
        店铺账号="eBay宇星店德国",
        平台="eBay",
        拣货备注="分单 请勿导入",
    )
    row = _prepare_row(raw, cfg, [], None, set(), {}, [])
    assert row["destination"] == "dian_01"
    partition = _partition([row], [])
    anomalies: list = []
    _exclude_ebay_split_from_dian_01(partition, anomalies)
    assert partition["dian_01"] == []
    assert partition["review"][0]["order_no"] == "24474201-44270"
    assert any(item.get("type") == "EBAY_SPLIT_SKIP_D01" and item.get("order_no") == "24474201-44270" for item in anomalies)


def test_ebay_split_two_rows_same_order_skipped():
    cfg = _cfg()
    rows = []
    for sku in ("WX10001", "WX10002"):
        raw = _raw(订单号="24474201-44270", 店铺账号="eBay宇星店德国", 平台="eBay", SKU=sku, 拣货备注="")
        rows.append(_prepare_row(raw, cfg, [], None, set(), {}, []))
    partition = _partition(rows, [])
    _exclude_ebay_split_from_dian_01(partition, [])
    assert partition["dian_01"] == []
    assert len(partition["review"]) == 2


def test_descriptive_pick_note_appended_but_zhizhan_and_china_direct_skipped():
    cfg = _cfg()
    note = "WX10171发一只美国仓，PayPal收36美金，记宇星美国托管"
    text = _append_logistics("谷仓仓储-宇星科技-加州-快递", note, "USD", cfg, ["WX10171"])
    assert note in text
    assert text.endswith("USD") or "USD" in text.split()

    zhizhan = _append_logistics("云途专线", "制单--YT1234567890", "USD", cfg, ["WX10001"])
    assert "YT1234567890" not in zhizhan
    already = _append_logistics("云途专线", "已制单--YT999", "USD", cfg, ["WX10001"])
    assert "YT999" not in already

    china = _append_logistics("谷仓仓储-深圳承思-捷克仓", "中国直发M00016 海外仓发产品", "EUR", cfg, ["WX17049"])
    assert "中国直发" not in china


def test_oversea_xlsx_row_height_matches_gold():
    cfg = _cfg()
    tmp = Path("/tmp/p2-rule-tests")
    src_dir = tmp / "in"
    out_dir = tmp / "out"
    src_dir.mkdir(parents=True, exist_ok=True)
    out_dir.mkdir(parents=True, exist_ok=True)
    src = src_dir / "orders.xlsx"
    headers = [
        "下单时间", "付款时间", "订单状态", "店铺账号", "订单号", "买家指定物流", "物流方式", "运单号",
        "SKU", "产品总数", "客服备注", "订单备注", "拣货备注", "买家账号", "买家Email", "收货人姓名",
        "详细地址", "收货人城市", "收货人州/省", "邮编", "收货人国家", "收货人电话", "收货人手机",
        "英文报关名", "中文报关名", "报关重量", "申报单价", "地址1", "地址2", "中文国家名", "包裹号",
        "订单金额", "币种缩写",
    ]
    wb = Workbook()
    ws = wb.active
    for c, h in enumerate(headers, start=1):
        ws.cell(1, c, h)
    row_vals = {
        "店铺账号": "亚马逊测试店",
        "订单号": "H-1",
        "物流方式": "谷仓仓储-鑫赢科技-捷克仓-DHL德国派送",
        "运单号": "TRK-H",
        "SKU": "WX10001",
        "产品总数": 1,
        "订单备注": "贴日本标签",
        "拣货备注": "千红捷克仓",
        "收货人姓名": "Ada",
        "收货人国家": "DE",
        "包裹号": "XM-H",
        "订单金额": 10,
        "币种缩写": "EUR",
    }
    for c, h in enumerate(headers, start=1):
        ws.cell(2, c, row_vals.get(h, ""))
    wb.save(src)
    wb.close()

    summary = process_orders([src], out_dir, config=cfg, export_date="2026-11-13")
    oversea = next(item for item in summary["downloads"] if item["kind"] == "oversea")
    path = out_dir / oversea["path"]

    # openpyxl 读到的默认行高 / 显式行高
    got = load_workbook(path)
    sheet = got.active
    assert float(sheet.sheet_format.defaultRowHeight) == float(OVERSEA_ROW_HEIGHT)
    assert bool(sheet.sheet_format.customHeight) is True
    for r in range(1, sheet.max_row + 1):
        height = sheet.row_dimensions[r].height
        assert height is not None, f"row {r} missing explicit height"
        assert float(height) == float(OVERSEA_ROW_HEIGHT)
    # 备注写进结果
    headers_out = [sheet.cell(1, c).value for c in range(1, sheet.max_column + 1)]
    data = {headers_out[i]: sheet.cell(2, i + 1).value for i in range(len(headers_out))}
    assert data["拣货备注"] == "千红捷克仓"
    assert data["订单备注"] == "贴日本标签"
    assert "千红捷克仓" in str(data["物流方式"])
    got.close()

    # XML 层：sheetFormatPr + 行 ht/customHeight，贴近金标准人工表
    with zipfile.ZipFile(path) as zf:
        xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
    fmt = re.search(r"<sheetFormatPr[^>]*>", xml)
    assert fmt, xml[:500]
    assert 'defaultRowHeight="25' in fmt.group(0)
    assert "customHeight=" in fmt.group(0)
    row_tags = re.findall(r"<row [^>]*>", xml)
    assert len(row_tags) >= 2
    for tag in row_tags:
        assert 'ht="25' in tag, tag
        assert "customHeight=" in tag, tag
