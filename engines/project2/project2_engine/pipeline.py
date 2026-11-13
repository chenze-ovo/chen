"""合并可见订单行，按配置路由，并写出海外仓 / 分组 / 店管家文件。

不读取项目1的 rules.yaml。中国直发不把金额改成 0，也不替换成 M00016。
店管家正式 xls 用模板/金标准壳填写；02：copy 金标准壳 + xlutils 只改 cell 值并复用原 XF，再 set_colour_RGB(26→#F2DCDB) 保证 Excel/WPS 粉底粉（禁止 soffice 转 xlsx/xls，禁止刷 #FFFFCC 黄）。海外仓/运营有 identical_refs 壳时复制壳填数（保留 theme/CF/列宽）。无壳才新建并补 theme1。店管家02合计：拆单未填/中国直发未匹配→留空进审核，禁止海外仓金额顶、按比例猜、写0冒充。

海外仓正本 Amazon：运单号=包裹号；物流方式=现物流+(原运单号非空且≠包裹号时追加)。店管家02 映射仍按原文物流。
中国直发M00016 仍留海外仓表，订单金额保留源表原值（不改 0、不回填订单价格）；店管家02合计仍可回填。
海外仓行高与金标准人工表一致：defaultRowHeight=25 且 customHeight。
海外仓副本（oversea_copy.mode=selected_orders）：仅当 copy_order_nos 非空时独立导出选中订单行；
规则 append_old_tracking_to_logistics + tracking=package_no。禁止再写全量隐藏行凑数。
"""
from __future__ import annotations

from decimal import Decimal, InvalidOperation
from pathlib import Path
import re

import xlwt
from openpyxl import Workbook, load_workbook

OVERSEA_ROW_HEIGHT = 25.0


def _combined_sku_maps(config: dict, sku_maps: list[dict] | None) -> list[dict]:
    """yaml.sku_map 优先，再合并调用方/数据库映射，按 source_sku 去重（大小写不敏感）。"""
    out: list[dict] = []
    seen: set[str] = set()
    yaml_items = [
        {"source_sku": str(source), "target_sku": str(target), "enabled": True}
        for source, target in (config.get("sku_map") or {}).items()
        if str(source).strip() and str(target).strip()
    ]
    incoming = [item for item in (sku_maps or []) if item.get("enabled", True)]
    for item in yaml_items + incoming:
        src = str(item.get("source_sku") or "").strip()
        if not src:
            continue
        key = src.upper()
        if key in seen:
            continue
        seen.add(key)
        out.append(item)
    return out


def process_orders(
    order_files: list[Path],
    out_dir: Path,
    *,
    config: dict,
    same_day_files: list[Path] | None = None,
    shops: list[dict] | None = None,
    sku_maps: list[dict] | None = None,
    sku_suffixes: list[str] | None = None,
    column_maps: dict | None = None,
    templates: dict | None = None,
    export_date: str | None = None,
    identical_refs: dict | None = None,
    copy_order_nos: list[str] | None = None,
    manual_amounts: dict | None = None,
) -> dict:
    out_dir.mkdir(parents=True, exist_ok=True)
    shops = shops or []
    sku_maps = _combined_sku_maps(config, sku_maps)
    suffixes = {str(item) for item in (sku_suffixes or []) if str(item).strip()}
    column_maps = column_maps or {}
    templates = templates or {}
    identical_refs = identical_refs or config.get("identical_refs") or {}
    anomalies: list[dict] = []
    headers: list[str] = []
    rows: list[dict] = []
    for path in order_files:
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            anomalies.append({"type": "SKIPPED_FILE", "file": path.name, "message": "只读取 xlsx 可见行"})
            continue
        try:
            file_headers, file_rows, hidden = _read_visible(path)
        except ValueError as exc:
            anomalies.append({"type": "UNREADABLE_FILE", "file": path.name, "message": str(exc)})
            continue
        if hidden:
            anomalies.append({"type": "HIDDEN_ROWS_SKIPPED", "file": path.name, "count": hidden})
        if file_headers and not headers:
            headers = file_headers
        rows.extend(file_rows)

    same_day = _read_amount_index(same_day_files or [], anomalies)
    # 第2步「TikTok/金额原件」里的 Order 形态表，行也参与转换（同时仍作金额来源）
    # 同一文件若已作为订单文件读过，不再重复计入行，避免重复行。
    already_read = {path.resolve() for path in order_files}
    rows.extend(
        _amount_file_order_rows(
            [path for path in (same_day_files or []) if path.resolve() not in already_read], anomalies
        )
    )
    routed = [_prepare_row(row, config, shops, sku_maps, suffixes, same_day, anomalies) for row in rows]
    routed = _merge_manual_duplicate_rows(routed, config, anomalies)
    _apply_manual_amounts(routed, manual_amounts or {})
    _mark_duplicate_orders(routed, anomalies)
    _sort_routed(routed, config)
    partition = _partition(routed, anomalies)
    _exclude_ebay_split_from_dian_01(partition, anomalies)

    downloads: list[dict] = []
    date_full, date_short = _resolve_export_dates(routed, config, export_date)
    naming = config.get("export_naming") or {}
    if headers and routed:
        oversea_name = str(naming.get("oversea") or "{date_full}海外仓订单.xlsx").format(
            date_full=date_full, date_short=date_short
        )
        oversea = out_dir / oversea_name
        oversea_rows = [row["cells"] for row in routed]
        oversea_ref = identical_refs.get("oversea")
        if oversea_ref:
            oversea_rows = _reorder_like_ref(oversea_rows, Path(oversea_ref))
        _write_xlsx(
            oversea,
            headers,
            oversea_rows,
            style_mode="oversea",
            width_ref=Path(oversea_ref) if oversea_ref else None,
            config=config,
            anomalies=anomalies,
        )
        downloads.append({"label": "海外仓订单", "path": oversea.name, "kind": "oversea"})

        ops_files = naming.get("ops_files") or {}
        ops_refs = identical_refs.get("ops") or {}
        for group_name, group_rows in _by_ops_group(routed).items():
            pattern = ops_files.get(group_name) or f"{group_name}单号{{date_short}}.xlsx"
            target_name = str(pattern).format(date_full=date_full, date_short=date_short)
            target = out_dir / target_name
            ops_cells = [row.get("ops_cells") or row["cells"] for row in group_rows]
            ops_ref = ops_refs.get(group_name)
            if ops_ref:
                ops_cells = _reorder_like_ref(ops_cells, Path(ops_ref))
            _write_xlsx(
                target,
                headers,
                ops_cells,
                style_mode="ops",
                width_ref=Path(ops_ref) if ops_ref else None,
                config=config,
            )
            downloads.append({"label": f"运营分组 {group_name}", "path": target.name, "kind": "ops_group"})

        internal = out_dir / "_internal"
        d01_verified = False
        d02_verified = False
        d02_pending = False
        oversea_ref_path = Path(oversea_ref) if oversea_ref else None
        if oversea_ref_path and oversea_ref_path.is_file():
            partition["dian_02"] = _reorder_routed_like_ref(partition["dian_02"], oversea_ref_path)
        for dest, label in (("dian_01", "店管家01"), ("dian_02", "店管家02")):
            chosen = partition[dest]
            # 路由清单仅内部核对，不进默认下载列表
            if chosen:
                internal.mkdir(parents=True, exist_ok=True)
                listing = internal / f"{label}_路由清单_仅核对.xlsx"
                _write_xlsx(listing, headers, [row["cells"] for row in chosen], style_mode="ops")
            template_path = templates.get(dest)
            if dest == "dian_02":
                gold_shell = identical_refs.get("dian_02")
                if gold_shell and Path(gold_shell).is_file():
                    template_path = gold_shell
            shared = _write_shared_template(out_dir, label, chosen, template_path, config=config, export_date=date_full, anomalies=anomalies)
            if shared is not None:
                verified = "verified"  # 01 样板 + 02 金标准全字段
                downloads.append({"label": label, "path": shared.name, "kind": dest, "p2_verification": verified})
                if dest == "dian_01":
                    d01_verified = True
                else:
                    d02_verified = True
            elif chosen:
                anomalies.append(
                    {
                        "type": "TEMPLATE_MISSING",
                        "destination": dest,
                        "message": f"{label} 空白模板不可用，未生成 xls。",
                    }
                )
            custom = _write_official(out_dir, f"{label}_自定义列", chosen, column_maps.get(dest) or [])
            if custom is not None:
                downloads.append({"label": f"{label}自定义列", "path": custom.name, "kind": f"{dest}_custom"})
        if not (d01_verified or d02_verified) and any(partition[d] for d in ("dian_01", "dian_02")):
            anomalies.append(
                {
                    "type": "P2_TEMPLATE_UNVERIFIED",
                    "message": "店管家模板未写出。",
                }
            )
        # oversea_copy.mode=selected_orders：仅勾选订单生成副本；空则不出副本文件
        selected_copy_nos = [str(x).strip() for x in (copy_order_nos or []) if str(x).strip()]
        if selected_copy_nos:
            copy_name = str(naming.get("oversea_copy") or "{date_full}海外仓订单 - 副本.xlsx").format(
                date_full=date_full, date_short=date_short
            )
            oversea_copy = out_dir / copy_name
            copy_rows = build_oversea_copy_rows(oversea_rows, selected_copy_nos)
            _write_xlsx(
                oversea_copy,
                headers,
                copy_rows,
                style_mode="oversea",
                width_ref=Path(oversea_ref) if oversea_ref else None,
                config=config,
                anomalies=anomalies,
            )
            downloads.append({
                "label": "海外仓订单 - 副本",
                "path": oversea_copy.name,
                "kind": "oversea_copy",
                "copy_order_nos": selected_copy_nos,
                "copy_row_count": len(copy_rows),
            })

    d01_nos = {row["order_no"] for row in partition["dian_01"]}
    d02_nos = {row["order_no"] for row in partition["dian_02"]}
    return {
        "engine": "project2",
        "executed": True,
        "execute_ignored": False,
        "files_received": [path.name for path in order_files],
        "stats": {
            "visible_orders": len(routed),
            "dian_01": len(partition["dian_01"]),
            "dian_02": len(partition["dian_02"]),
            "review": len(partition["review"]),
            "human_review": len(partition["human_review"]),
            "overlap_01_02": len(d01_nos & d02_nos),
        },
        "orders": [
            {
                "order_no": row["order_no"],
                "shop_account": row["shop"],
                "platform": row["platform"],
                "destination": row["destination"],
                "row_mark": row["row_mark"],
                "currency": row["currency"],
                "order_amount": _money(row["amount"]),
                "skus": row["skus"],
                "quantity": row["quantity"],
                "tracking_nos": row["tracking_nos"],
                "country": row["country"],
                "buyer_name": row["buyer"],
                "already_in_guanjia": row["already"],
                "ops_group": row["ops_group"],
                "multi_sku": len(row["skus"]) > 1,
                "china_direct": row["china_direct"],
                "amount_matched": row["amount_matched"],
            }
            for row in routed
        ],
        "anomalies": anomalies,
        "review_tasks": [item for item in anomalies if item.get("task_type")],
        "downloads": downloads,
        "reason": "已按配置路由；SKU 凡带「-」一律剥掉「-」后后缀（不限名单）。交付为海外仓正本（副本需勾选订单另生成）+6运营+店管家01/02；路由清单仅留 _internal。店管家02按金标准全字段填写。",
    }


def _read_visible(path: Path) -> tuple[list[str], list[dict], int]:
    try:
        wb = load_workbook(path, data_only=True)
    except Exception as exc:
        raise ValueError(f"{path.name} 不是可读的 xlsx：{type(exc).__name__}") from exc
    ws = wb.active
    headers = []
    for cell in ws[1]:
        if cell.value is None:
            continue
        headers.append(str(cell.value).strip())
    width = len(headers)
    rows: list[dict] = []
    hidden = 0
    for idx in range(2, ws.max_row + 1):
        dim = ws.row_dimensions[idx]
        if dim.hidden:
            hidden += 1
            continue
        values = [ws.cell(idx, col).value for col in range(1, width + 1)]
        if all(value is None or str(value).strip() == "" for value in values):
            continue
        cells = {headers[pos]: values[pos] if pos < len(values) else None for pos in range(width)}
        mark = _row_mark(ws.cell(idx, 1))
        rows.append(
            {
                "cells": cells,
                "row_mark": mark,
                "source": path.name,
                "excel_row": idx,
            }
        )
    wb.close()
    return headers, rows, hidden


def _row_mark(cell) -> str:
    fill = cell.fill
    rgb = None
    if fill is not None and fill.patternType == "solid" and fill.fgColor is not None:
        rgb = fill.fgColor.rgb if fill.fgColor.type == "rgb" else None
    if isinstance(rgb, str) and len(rgb) >= 6:
        hexcolor = rgb[-6:]
        try:
            red = int(hexcolor[0:2], 16)
            green = int(hexcolor[2:4], 16)
            blue = int(hexcolor[4:6], 16)
        except ValueError:
            return "white"
        if green > red + 20 and green > blue + 20 and green > 80:
            return "green"
    return "white"


def _is_order_shaped(headers: list[str]) -> bool:
    """Order 形态：含订单号，且有店铺/SKU/物流等订单列（排除纯金额对照表）。"""
    head = {str(h).strip() for h in headers}
    if "订单号" not in head:
        return False
    markers = {"店铺账号", "SKU", "物流方式", "运单号", "收货人姓名"}
    return len(head & markers) >= 2


def _amount_file_order_rows(paths: list[Path], anomalies: list[dict]) -> list[dict]:
    """第2步「TikTok/金额原件」栏里的 Order 形态表：行也参与转换，同时仍作金额来源。

    这些文件以前只当金额表读，导致只在其中出现的「待打单（有货）」整单丢失。
    """
    out: list[dict] = []
    for path in paths or []:
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            continue
        try:
            file_headers, file_rows, hidden = _read_visible(path)
        except ValueError:
            continue
        if not _is_order_shaped(file_headers):
            anomalies.append(
                {
                    "type": "AMOUNT_FILE_NOT_ORDER_SHAPED",
                    "file": path.name,
                    "message": "第2步文件只有金额列、没有订单列，仅作金额来源。",
                }
            )
            continue
        if hidden:
            anomalies.append({"type": "HIDDEN_ROWS_SKIPPED", "file": path.name, "count": hidden})
        out.extend(file_rows)
    return out


def _read_amount_index(paths: list[Path], anomalies: list[dict]) -> dict[str, Decimal]:
    """读取店小秘或金额原件（凡含订单号+订单金额的 xlsx 均可）。

    订单号常带 -1/-2 后缀，额外按去掉末尾 -数字 的原单号索引，便于海外仓原单号匹配。
    同一单号多份文件都有时：非 0 金额优先，禁止用 0 覆盖真实金额；两份非 0 且不等才进审核。
    """
    import re

    found: dict[str, Decimal] = {}
    conflicts: dict[str, set[Decimal]] = {}

    def _put(key: str, amount: Decimal, file_name: str) -> None:
        if not key:
            return
        current = found.get(key)
        if current is None:
            found[key] = amount
            return
        if amount == current:
            return
        if amount == 0:
            # 0 不覆盖真实金额
            return
        if current == 0:
            found[key] = amount
            return
        conflicts.setdefault(key, {current, amount}).add(amount)

    for path in paths:
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            anomalies.append({"type": "SKIPPED_FILE", "file": path.name, "message": "店小秘或金额原件只读取 xlsx"})
            continue
        try:
            _headers, rows, _hidden = _read_visible(path)
        except ValueError as exc:
            anomalies.append({"type": "UNREADABLE_FILE", "file": path.name, "message": str(exc)})
            continue
        for row in rows:
            order_no = _text(row["cells"].get("订单号"))
            amount = _decimal(row["cells"].get("订单金额"))
            if order_no and amount is not None:
                _put(order_no, amount, path.name)
                # 店小秘/金额原件：305-xxx-1 → 同时登记 305-xxx，供海外仓原单号匹配
                base = re.sub(r"-\d+$", "", order_no)
                if base and base != order_no:
                    _put(base, amount, path.name)
    for order_no, values in conflicts.items():
        anomalies.append(
            {
                "type": "AMOUNT_CONFLICT",
                "task_type": "amount_conflict",
                "order_no": order_no,
                "message": "同一订单号多份文件都有非 0 且不相等的金额：" + "、".join(sorted(str(v) for v in values)) + "，转人工确认。",
            }
        )
    return found



def _is_manual_shop(row: dict, config: dict) -> bool:
    exclude = {str(x) for x in ((config.get("ops_export") or {}).get("exclude_shops") or ["手工订单"])}
    raw = _text(row.get("raw_shop"))
    return bool(raw) and raw in exclude


def _merge_manual_duplicate_rows(routed: list[dict], config: dict, anomalies: list[dict]) -> list[dict]:
    """同一订单号既有真实店铺行、又有「手工订单」重发行（金额 0）时，合成一行。

    待打单行带真实店铺/真实金额，手工行只是同单的海外仓重发记录：保留真实行，
    金额取非 0 的那一行，禁止 0 覆盖真实金额。真实拆单（同批次两行不同 SKU）不动。
    """
    groups: dict[str, list[dict]] = {}
    for row in routed:
        key = _text(row.get("order_no"))
        if key:
            groups.setdefault(key, []).append(row)
    drop: set[int] = set()
    for order_no, rows in groups.items():
        if len(rows) < 2:
            continue
        manual = [r for r in rows if _is_manual_shop(r, config)]
        real = [r for r in rows if not _is_manual_shop(r, config)]
        if not manual or not real:
            continue
        real_amounts = [r["amount"] for r in real if r.get("amount") not in (None, 0)]
        manual_amounts = [r["amount"] for r in manual if r.get("amount") not in (None, 0)]
        if not real_amounts and manual_amounts:
            # 真实行没金额，手工行有：反过来留手工行
            keep = manual[0]
            real_amounts = manual_amounts
            drop.update(id(r) for r in real)
        else:
            keep = next((r for r in real if r.get("amount") not in (None, 0)), real[0])
            drop.update(id(r) for r in manual)
        if not real_amounts:
            anomalies.append(
                {
                    "type": "DUPLICATE_AMOUNT_ALL_ZERO",
                    "order_no": order_no,
                    "message": "同订单号多行金额全为 0，已合成一行，需人工核对金额。",
                }
            )
        anomalies.append(
            {
                "type": "MERGED_MANUAL_RESHIP",
                "order_no": order_no,
                "file": "海外仓订单",
                "message": (
                    f"同订单号另有「手工订单」重发行（运单 {', '.join(_text(r['cells'].get('运单号')) for r in manual) or '空'}），"
                    f"已合成一行，保留真实店铺 {keep.get('shop')} 与金额 {keep.get('amount')}。"
                ),
            }
        )
    return [row for row in routed if id(row) not in drop]


def _merge_same_line_rows(routed: list[dict], anomalies: list[dict]) -> None:
    """兼容入口：调用方已合并时保持原样。"""
    return None


def _sort_routed(routed: list[dict], config: dict) -> None:
    """全表按 物流方式 → 店铺账号 → 订单号 排；同单号多行相邻，金额大的在前。"""
    if not ((config.get("export_sort") or {}).get("enabled", True)):
        return

    def key(row: dict):
        cells = row.get("cells") or {}
        return (
            _sort_logistics_key(cells.get("物流方式"), config),
            _text(row.get("shop") or cells.get("店铺账号")),
            _text(row.get("order_no")),
            -float(row.get("amount") or 0),
        )

    routed.sort(key=key)


def _sort_logistics_key(value, config: dict) -> str:
    """排序用物流方式：去掉币种后缀和「分单」尾标，让同一物流的变体相邻成段。"""
    text = _text(value)
    if not text:
        return ""
    skip = {str(item).upper() for item in ((config.get("logistics_append") or {}).get("skip_currency") or ["EUR"])}
    parts = text.split()
    if len(parts) > 1 and parts[-1].upper() not in skip and (len(parts[-1]) == 3 and parts[-1].isalpha()):
        text = " ".join(parts[:-1])
    if text.endswith("分单"):
        text = text[:-2].strip()
    return text


def _apply_manual_amounts(routed: list[dict], manual_amounts: dict) -> None:
    """把审核回填金额写入行：key 为 订单号 或 订单号|SKU。"""
    if not manual_amounts:
        return
    for row in routed:
        order_no = str(row.get("order_no") or "").strip()
        skus = row.get("skus") or []
        sku = skus[0] if len(skus) == 1 else ""
        val = None
        if sku and f"{order_no}|{sku}" in manual_amounts:
            val = manual_amounts[f"{order_no}|{sku}"]
        elif order_no in manual_amounts and manual_amounts[order_no] not in (None, ""):
            val = manual_amounts[order_no]
        if val in (None, ""):
            continue
        row["manual_sku_amount"] = val
        cells = row.get("cells") or {}
        cells["manual_sku_amount"] = val
        row["cells"] = cells


def _prepare_row(raw: dict, config: dict, shops: list[dict], sku_maps: list[dict], suffixes: set[str], same_day: dict[str, Decimal], anomalies: list[dict]) -> dict:
    cells = dict(raw["cells"])
    source_notes = {key: cells.get(key) for key in ("客服备注", "订单备注", "拣货备注")}
    original_amount = cells.get("订单金额")
    note = _text(cells.get("拣货备注"))
    original_shop = _text(cells.get("店铺账号"))
    raw_logistics = cells.get("物流方式")
    shop = _shop_from_note(cells, note, anomalies, config, shops)
    cells["店铺账号"] = shop or cells.get("店铺账号")
    platform = _text(cells.get("平台"))
    order_no = _text(cells.get("订单号"))
    shop_rule = _shop_record(shop, shops)
    already = bool(shop_rule and shop_rule.get("already_in_guanjia"))
    destination = _route(shop, platform, already, config, shop_rule)
    raw_sku_text = cells.get("SKU")
    suspects_before = len(config.get("_sku_suspects") or [])
    sku_maps = _combined_sku_maps(config, sku_maps)
    skus, unknown_suffix = _transform_skus(_split_sku(raw_sku_text), suffixes, sku_maps, note, config)
    for suspect in (config.get("_sku_suspects") or [])[suspects_before:]:
        if suspect.get("audit_only"):
            continue
        anomalies.append(
            {
                "type": "SKU_STRIP_SUSPECT",
                "task_type": "sku_strip_suspect",
                "order_no": order_no,
                "message": (
                    f"SKU {suspect.get('original')} 剥成 {suspect.get('result')}：{suspect.get('reason')}。"
                    "未写死特例，转人工确认后再改规则。"
                ),
            }
        )
    if unknown_suffix:
        anomalies.append(
            {
                "type": "UNKNOWN_SKU_SUFFIX",
                "order_no": order_no,
                "message": f"SKU 末段 {unknown_suffix} 不在可配置后缀表里，未剥离，转人工。",
            }
        )
    if skus:
        cells["SKU"] = "\n".join(skus)
    logistics = _append_logistics(
        _text(cells.get("物流方式")),
        note,
        _text(cells.get("币种缩写") or cells.get("币种")),
        config,
        skus,
    )
    if logistics:
        cells["物流方式"] = logistics
    _fill_tracking(cells, note, config)
    original_tracking = _text(cells.get("运单号"))
    if _is_amazon_logistics(_text(cells.get("物流方式")) or _text(raw_logistics), config):
        _swap_tracking_to_package(cells)
    # 源表备注写进海外仓结果，禁止后续步骤覆盖
    for key, value in source_notes.items():
        cells[key] = value
    # 海外仓正本订单金额始终保留源表原值，禁止改 0、禁止用订单价格或店小秘金额覆盖
    cells["订单金额"] = original_amount
    amount = _decimal(original_amount)
    china = _is_china_direct(cells, config)
    matched = False
    if china:
        if order_no in same_day:
            # 匹配金额只进内部 amount / amount_matched，供店管家02合计；海外仓正本 cells 订单金额保持原值
            amount = same_day[order_no]
            matched = True
        else:
            anomalies.append(
                {
                    "type": "CHINA_DIRECT_AMOUNT_UNMATCHED",
                    "task_type": "china_direct_amount_unmatched",
                    "order_no": order_no,
                    "message": "中国直发未匹配店小秘或金额原件：店管家02合计留空并进审核；禁止写0，禁止用海外仓金额顶上。",
                }
            )
    if unknown_suffix:
        destination = "human_review"
    # 拣货备注含「管家已建单」→ 审核，不进 01/02（金标准店管家02排除这批）
    for needle in (config.get("d02_rules") or {}).get("exclude_note_contains") or ["管家已建单"]:
        if needle and needle in note:
            destination = "review"
            anomalies.append(
                {
                    "type": "ALREADY_IN_GUANJIA_NOTE",
                    "order_no": order_no,
                    "message": f"拣货备注含「{needle}」，进审核、不进店管家01/02。",
                }
            )
            break
    ops_cells = {**cells, "SKU": raw_sku_text}
    if (config.get("ops_export") or {}).get("keep_raw_logistics", True):
        ops_cells["物流方式"] = raw_logistics
    if original_tracking:
        ops_cells["运单号"] = original_tracking
    # 运营表每行补币种后缀（非 EUR）：主表已加、分店表不加会导致同一单两种物流
    ops_skip = {str(x).upper() for x in ((config.get("logistics_append") or {}).get("skip_currency") or ["EUR", "欧元"])}
    ops_currency = _text(cells.get("币种缩写") or cells.get("币种"))
    if ops_currency and ops_currency.upper() not in ops_skip:
        body = _text(ops_cells.get("物流方式"))
        if body and ops_currency not in body.split():
            ops_cells["物流方式"] = f"{body} {ops_currency}".strip()
    return {
        "cells": cells,
        "order_no": order_no,
        "raw_shop": original_shop,
        "shop": shop,
        "platform": platform,
        "destination": destination,
        "row_mark": raw["row_mark"],
        "currency": _text(cells.get("币种缩写") or cells.get("币种")) or "USD",
        "amount": amount,
        "skus": skus or [""],
        "quantity": _text(cells.get("产品总数")),
        "raw_logistics": raw_logistics,
        "tracking_nos": _split_lines(original_tracking) or _split_lines(cells.get("运单号")),
        "country": _text(cells.get("收货人国家")),
        "buyer": _text(cells.get("收货人姓名")),
        "already": already,
        "ops_group": _ops_group_for(original_shop, note, config, shops, shop),
        "ops_cells": ops_cells,
        "china_direct": china,
        "amount_matched": matched,
    }


_MANUAL_SHOP = "手工订单"
_NOTE_SHOP_COLUMNS = ("客服备注", "拣货备注", "订单备注")
_NON_SHOP_DIAN = ("店铺", "商店", "门店", "网店", "店内", "店里", "到店", "本店", "进店")
_SHOP_INSTRUCTIONAL = (
    "中国直发",
    "海外仓发",
    "产品发",
    "发亚马逊",
    "FBA库存",
    "缺件",
    "给客户",
    "计入",
    "待核对",
    "管家已建单",
    "已上架",
    "移仓",
    "制单--",
    "已制单--",
)
# ahav宇星店德国 / ebay宇星店德国（自营）
_SHOP_DIAN_RE = re.compile(r"[A-Za-z0-9\u4e00-\u9fff]*店[A-Za-z0-9\u4e00-\u9fff（）()]{0,20}")
_SHOP_OWNED_RE = re.compile(
    r"(?:亚马逊|Amazon|ebay|eBay|AHAV|ahav|TikTok|速卖通|Cdiscount|Zeemoto|Shirley|"
    r"宇星|创趣|林艺奋|吴晓虹|鲁圳|承思)[A-Za-z0-9\u4e00-\u9fff（）()_\- ]{0,30}"
    r"(?:（自营）|\(自营\)|半托管\d*店?|不贴logo)"
)
_TRACKING_RE = re.compile(r"(?:已)?制单--+[^\r\n]*")


def _known_shop_names(config: dict | None, shops: list[dict] | None) -> list[str]:
    names: set[str] = set()
    for item in shops or []:
        name = _text((item or {}).get("shop_account"))
        if name and name != _MANUAL_SHOP:
            names.add(name)
    for rule in (config or {}).get("ops_group_assignments") or []:
        name = _text((rule or {}).get("shop_account"))
        if name and name != _MANUAL_SHOP:
            names.add(name)
    return sorted(names, key=lambda x: (-len(x), x))


def _strip_tracking(text: str) -> str:
    return _TRACKING_RE.sub(" ", text or "").strip()


def _looks_like_shop_phrase(text: str) -> bool:
    """候选片段是否像店名，而不是一段说明或运单。"""
    t = _text(text)
    if not t or t == _MANUAL_SHOP or len(t) > 40:
        return False
    if t.startswith("-") or t.isdigit() or _TRACKING_RE.search(t):
        return False
    if any(p in t for p in ("。", "！", "？", "；", ";", "\n")):
        return False
    if any(n in t for n in _SHOP_INSTRUCTIONAL):
        return False
    if "补发" in t or "重新" in t:
        return False
    if t in _NON_SHOP_DIAN:
        return False
    if "店" in t:
        # 仅「店铺/门店」这类词不算店名
        remainder = t
        for bad in _NON_SHOP_DIAN:
            remainder = remainder.replace(bad, "")
        if "店" in remainder:
            return True
    if any(k in t for k in ("（自营）", "(自营)", "半托管", "不贴logo")):
        return True
    if _SHOP_OWNED_RE.fullmatch(t) or _SHOP_DIAN_RE.fullmatch(t):
        return True
    return False


def _is_standalone_shop_name(text: str, known: list[str]) -> bool:
    """客服备注整列：短且像店名时才采用，不把说明性段落当店名。"""
    t = _text(text)
    if not t or t == _MANUAL_SHOP:
        return False
    folded = t.casefold()
    if any(t == k or folded == k.casefold() for k in known):
        return True
    if "\n" in t or len(t) > 40:
        return False
    return _looks_like_shop_phrase(t)


def _longest_known_in(text: str, known: list[str]) -> str:
    hay = _strip_tracking(text)
    if not hay or not known:
        return ""
    folded = hay.casefold()
    for name in known:
        if name and name.casefold() in folded:
            return name
    return ""


def _extract_shop_from_text(text: str, known: list[str]) -> str:
    """从一段备注里抽店名：已知名单、--- 之后、末尾「…店…」短语。"""
    cleaned = _strip_tracking(text)
    if not cleaned:
        return ""
    found = _longest_known_in(cleaned, known)
    if found:
        return found
    if "---" in cleaned:
        tail = cleaned.split("---")[-1].strip()
        tail = (tail.splitlines()[0] if tail else "").strip()
        if tail and not tail.startswith("-"):
            known_tail = _longest_known_in(tail, known)
            if known_tail:
                return known_tail
            if _looks_like_shop_phrase(tail):
                return tail
    tokens = re.split(r"[\s,，、]+", cleaned)
    for token in reversed(tokens):
        token = token.strip().strip("。．.！!？?")
        if _looks_like_shop_phrase(token):
            return token
    candidates: list[tuple[int, str]] = []
    for pattern in (_SHOP_DIAN_RE, _SHOP_OWNED_RE):
        for match in pattern.finditer(cleaned):
            phrase = match.group(0).strip()
            if _looks_like_shop_phrase(phrase):
                candidates.append((match.start(), phrase))
    if candidates:
        # 同起点取更长，再取文中最后一次出现
        best: dict[int, str] = {}
        for start, phrase in candidates:
            if len(phrase) >= len(best.get(start, "")):
                best[start] = phrase
        last_start = max(best)
        return best[last_start]
    return ""


def _shop_from_note(
    cells: dict,
    note: str,
    anomalies: list[dict],
    config: dict | None = None,
    shops: list[dict] | None = None,
) -> str:
    """手工订单按备注里的真实店名改写店铺账号。

    不要求备注必须有「---」。客服备注若本身是短店名则优先；
    否则在客服/拣货/订单备注里做已知店名最长匹配，或抽取末尾含「店」的短语、--- 后的店名。
    整段说明、「制单--」运单、中国直发说明不能当店名；认不出则保持手工订单。
    非手工订单仍保留旧规则：拣货备注里 --- 后面的店名优先。
    """
    shop = _text(cells.get("店铺账号"))
    order_no = _text(cells.get("订单号"))
    known = _known_shop_names(config, shops)
    rules = (config or {}).get("manual_shop_rules") or []
    preferred_columns: list[str] = []
    for rule in rules:
        when_shop = str(rule.get("when_shop") or "")
        column = str(rule.get("shop_from_column") or "")
        needle = str(rule.get("note_contains") or "")
        target = str(rule.get("shop_account") or "")
        if when_shop and shop != when_shop:
            continue
        if needle and needle in (note or "") and target:
            return target
        if column and column not in preferred_columns:
            preferred_columns.append(column)

    if shop == _MANUAL_SHOP:
        for column in preferred_columns or ["客服备注"]:
            taken = _text(cells.get(column))
            if taken and _is_standalone_shop_name(taken, known):
                return taken
        fields = [_text(cells.get(col)) for col in _NOTE_SHOP_COLUMNS]
        best = ""
        for field in fields:
            found = _longest_known_in(field, known)
            if found and len(found) > len(best):
                best = found
        if best:
            return best
        for field in fields:
            extracted = _extract_shop_from_text(field, known)
            if extracted:
                return extracted
        return shop

    if "---" not in (note or ""):
        return shop
    tail = note.split("---")[-1].strip()
    if not tail or tail.startswith("-"):
        anomalies.append(
            {
                "type": "MANUAL_SHOP_UNRESOLVED",
                "order_no": order_no,
                "message": "备注里有 --- 但没有店名，店铺保持原列并转人工核对。",
            }
        )
        return shop
    return tail


def _transform_skus(skus: list[str], suffixes: set[str], sku_maps: list[dict], note: str, config: dict) -> tuple[list[str], str]:
    """海外仓/店管家用 SKU：有「-」则剥掉最后一段后缀（任意 -ZEE/-CQCSE/...），不再靠白名单。

    suffixes 参数保留兼容，已不再作为是否剥离的门槛。运营表仍用原始 SKU（ops_cells）。
    返回 unknown（非空即进审核）；疑似剥错的情况写入 config["_sku_suspects"] 供上层报审核。
    """
    table = {str(item.get("source_sku") or "").upper(): str(item.get("target_sku") or "") for item in sku_maps}
    rules = config.get("sku_rules") or {}
    wx_digits = int(rules.get("wx_digits") or 5)
    suspects: list[dict] = config.setdefault("_sku_suspects", [])
    cleaned: list[str] = []
    for sku in skus:
        text = sku.strip().upper()
        if text in table:
            cleaned.append(table[text])
            continue
        base = text
        tail = ""
        # 删掉最后一个「-」及后面内容，例如 WX10003-ZEE → WX10003
        if "-" in base:
            left, tail = base.rsplit("-", 1)
            if left:
                base = left
        if base in table:
            cleaned.append(table[base])
            continue
        # WX 货号位数不做校验：用户 2026-11-04 明确 WX1915-YXZR → WX1915，不补 0、不报疑似剥错
        picked = _sku_from_note(base, note, config)
        if picked:
            # 纯数字平台码按备注货号映射是对的（人工就是这么做的），但列进审核清单留痕
            suspects.append({"original": text, "result": picked, "reason": "纯数字平台码，按备注货号映射", "audit_only": True})
            cleaned.append(picked)
            continue
        # 纯数字长码（条码式 SKU）不在备注里找货号时不要静默替换
        cleaned.append(base)
    return cleaned, ""


def _sku_from_note(sku: str, note: str, config: dict) -> str:
    rules = config.get("sku_from_pick_note") or {}
    if not rules.get("when_sku_digits") or not sku.isdigit():
        return ""
    import re

    pattern = str(rules.get("token_pattern") or r"WX\d+")
    found = re.findall(pattern, note.upper())
    return found[-1] if found else ""


def _fill_tracking(cells: dict, note: str, config: dict) -> None:
    """从拣货备注抽运单号：已有运单号不覆盖；制单--/已制单-- 后到换行前整段。"""
    if _text(cells.get("运单号")):
        return
    import re

    pattern = str(
        (config.get("tracking_from_note") or {}).get("pattern")
        or r"(?:已)?制单--([^\r\n]+)"
    )
    matched = re.search(pattern, note or "")
    if matched:
        value = matched.group(1).strip()
        if value:
            cells["运单号"] = value


def _ops_excluded(note: str, config: dict) -> bool:
    for needle in config.get("ops_exclude_note_contains") or ["中国直发M00016"]:
        if needle and needle in note:
            return True
    return False


def _ops_group_for(original_shop: str, note: str, config: dict, shops: list[dict], resolved_shop: str | None = None) -> str | None:
    """运营分组按原始店铺归属；手工订单和中国直发备注不进 6 组。

    优先 routing.yaml ops_group_assignments（Shirley→Cdiscount、吴晓虹美国→林艺奋等并入），
    再查 DB 店铺目录。eBay 默认无 ops_group（不单独发表运营表）。
    """
    if _ops_excluded(note, config):
        return None
    exclude_shops = {str(item) for item in ((config.get("ops_export") or {}).get("exclude_shops") or ["手工订单"])}
    effective = _text(resolved_shop) or original_shop
    if effective in exclude_shops:
        return None
    # eBay：默认不进 6 组运营表（emit_ebay_ops_file=false），避免误当成漏表
    ops_export = config.get("ops_export") or {}
    if not ops_export.get("emit_ebay_ops_file", False) and "ebay" in original_shop.lower():
        return None
    assigned = _ops_assignment_for(original_shop, config)
    if assigned and assigned.get("ops_group"):
        return str(assigned["ops_group"])
    rule = _shop_record(original_shop, shops)
    if (rule or {}).get("ops_group"):
        return rule["ops_group"]
    # 手工订单按客服备注改写成真实店铺后，仍归该店铺的运营表（如 Zeemoto 的 -1 重发行）
    resolved = _text(resolved_shop)
    if resolved and resolved != original_shop:
        assigned = _ops_assignment_for(resolved, config)
        if assigned and assigned.get("ops_group"):
            return str(assigned["ops_group"])
        alt = _shop_record(resolved, shops)
        if alt and alt.get("ops_group"):
            return alt["ops_group"]
    return None


def _ops_assignment_for(shop: str, config: dict) -> dict | None:
    """routing.yaml ops_group_assignments：精确店名或 shop_contains。"""
    for rule in config.get("ops_group_assignments") or []:
        acct = str(rule.get("shop_account") or "")
        if acct and shop == acct:
            return rule
        contains = str(rule.get("shop_contains") or "")
        if contains and contains.lower() in shop.lower():
            return rule
    return None


def _is_zhizhan_note_line(line: str, skip_prefixes: tuple[str, ...]) -> bool:
    """制单-- / 已制单-- 行只用于抽运单号，不拼进物流方式。"""
    text = (line or "").strip()
    if not text:
        return True
    prefixes = list(skip_prefixes) + ["制单--", "已制单--", "制单", "已制单"]
    return text.startswith(tuple(dict.fromkeys(prefixes)))


def _append_logistics(logistics: str, pick_note: str, currency: str, config: dict, skus: list[str]) -> str:
    """说明性拣货备注（非制单--/已制单--）整段追加到海外仓物流方式。"""
    rules = config.get("logistics_append") or {}
    skip_prefixes = tuple(str(item) for item in (rules.get("skip_pick_note_prefixes") or ["制单", "已制单"]))
    skip_currency = {str(item).upper() for item in (rules.get("skip_currency") or ["EUR", "欧元"])}
    deny_contains = [str(item) for item in (rules.get("deny_contains") or ["中国直发", "管家已建单"])]
    text = _replace_channel(logistics, config)
    parts: list[str] = []
    for line in (pick_note or "").replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        line = line.strip()
        if not line or _is_zhizhan_note_line(line, skip_prefixes):
            continue
        if any(item in line for item in deny_contains):
            continue
        parts.append(line)
    extra = " ".join(parts).strip()
    if extra and extra not in text:
        text = f"{text} {extra}".strip()
    if currency and currency.upper() not in skip_currency and currency not in text.split():
        text = f"{text} {currency}".strip()
    return text


def _replace_channel(logistics: str, config: dict) -> str:
    text = logistics
    for rule in config.get("channel_replacements") or []:
        source = str(rule.get("from") or "")
        target = str(rule.get("to") or "")
        if source and source in text and target:
            text = text.replace(source, target)
    return text


def _route(shop: str, platform: str, already: bool, config: dict, shop_rule: dict | None) -> str:
    if already:
        return "review"
    if shop_rule and shop_rule.get("destination"):
        return str(shop_rule["destination"])
    assigned = _ops_assignment_for(shop, config)
    if assigned and assigned.get("destination"):
        return _dest_code(assigned.get("destination"))
    blob = f"{shop} {platform}".lower()
    for policy in config.get("routing_policy") or []:
        if not policy.get("enabled"):
            continue
        when = policy.get("when") or {}
        if when.get("fallback") or when.get("already_built_in_guanjia"):
            continue
        if _policy_hit(shop, blob, when, policy.get("match") or "all"):
            return _dest_code(policy.get("destination"))
    return _dest_code(config.get("fallback_destination") or "dian_02")


def _policy_hit(shop: str, blob: str, when: dict, mode: str) -> bool:
    checks: list[bool] = []
    names = [str(item) for item in (when.get("shop_name_contains") or []) if str(item).strip()]
    if names:
        checks.append(any(item.lower() in shop.lower() or item.lower() in blob for item in names))
    platforms = [str(item) for item in (when.get("platforms") or []) if str(item).strip()]
    if platforms:
        checks.append(any(item.lower() in blob for item in platforms))
    accounts = [str(item) for item in (when.get("shop_accounts") or []) if str(item).strip()]
    if "shop_accounts" in when:
        checks.append(bool(accounts) and shop in accounts)
    if not checks:
        return False
    return any(checks) if mode == "any" else all(checks)


def _dest_code(value: str | None) -> str:
    text = str(value or "")
    if text in {"dian_01", "店管家01"}:
        return "dian_01"
    if text in {"dian_02", "店管家02"}:
        return "dian_02"
    if text in {"review", "human_review"}:
        return text
    return "human_review"


def _shop_record(shop: str, shops: list[dict]) -> dict | None:
    for item in shops:
        if str(item.get("shop_account") or "") == shop:
            return item
    return None


def _is_china_direct(cells: dict, config: dict) -> bool:
    needles = [str(item) for item in ((config.get("china_direct") or {}).get("note_contains") or []) if str(item).strip()]
    if not needles:
        return False
    blob = " ".join(_text(cells.get(key)) for key in ("拣货备注", "客服备注", "订单备注"))
    return any(needle in blob for needle in needles)


def _apply_sku_maps(skus: list[str], sku_maps: list[dict]) -> list[str]:
    table = {str(item.get("source_sku")): str(item.get("target_sku")) for item in sku_maps if item.get("source_sku")}
    return [table.get(sku, sku) for sku in skus]


def _split_sku(value) -> list[str]:
    return _split_lines(value)


def _split_lines(value) -> list[str]:
    if value is None:
        return []
    return [part.strip() for part in str(value).replace("\r\n", "\n").replace("\r", "\n").split("\n") if part.strip()]



def _row_is_ebay(row: dict) -> bool:
    """与 _d01_original_order_no 一致：店铺/平台含 ebay（不区分大小写）。"""
    shop = _text(row.get("shop") or (row.get("cells") or {}).get("店铺账号"))
    platform = _text(row.get("platform") or (row.get("cells") or {}).get("平台"))
    return "ebay" in f"{shop} {platform}".lower()


def _exclude_ebay_split_from_dian_01(partition: dict[str, list[dict]], anomalies: list[dict]) -> None:
    """eBay 同一完整订单号多行，或备注含「分单」：整单不进店管家01，改入审核。

    按原 order_no 计数，不用截断后的原始单号。非 eBay 分单（店管家02 -1）不动。
    """
    d01 = partition.get("dian_01") or []
    counts: dict[str, int] = {}
    noted: set[str] = set()
    for row in d01:
        if not _row_is_ebay(row):
            continue
        order_no = _text(row.get("order_no"))
        if not order_no:
            continue
        counts[order_no] = counts.get(order_no, 0) + 1
        cells = row.get("cells") or {}
        blob = " ".join(_text(cells.get(key)) for key in ("拣货备注", "客服备注", "订单备注"))
        if "分单" in blob:
            noted.add(order_no)
    split_nos = {no for no, n in counts.items() if n > 1} | noted
    if not split_nos:
        return
    keep: list[dict] = []
    moved: list[dict] = []
    for row in d01:
        if _text(row.get("order_no")) in split_nos and _row_is_ebay(row):
            row["destination"] = "review"
            moved.append(row)
        else:
            keep.append(row)
    partition["dian_01"] = keep
    partition.setdefault("review", []).extend(moved)
    for order_no in sorted(split_nos):
        anomalies.append(
            {
                "type": "EBAY_SPLIT_SKIP_D01",
                "order_no": order_no,
                "file": "店管家01.xls",
                "message": "eBay分单不进店管家01",
            }
        )


def _mark_duplicate_orders(rows: list[dict], anomalies: list[dict]) -> None:
    """分单同号多行：店管家02 导出时追加 -1/-2，不再整批转人工。"""
    seen: dict[str, int] = {}
    for row in rows:
        if not row["order_no"]:
            continue
        seen[row["order_no"]] = seen.get(row["order_no"], 0) + 1
    duplicated = {key for key, count in seen.items() if count > 1}
    for order_no in sorted(duplicated):
        anomalies.append(
            {
                "type": "DUPLICATE_ORDER",
                "order_no": order_no,
                "message": "订单号多行（分单）：店管家02 第2行起订单号追加 -1；合计留空等人工填拆单金额，禁止按比例猜、禁止用海外仓总额顶。",
            }
        )


def _partition(rows: list[dict], anomalies: list[dict]) -> dict[str, list[dict]]:
    groups = {"dian_01": [], "dian_02": [], "review": [], "human_review": []}
    for row in rows:
        dest = row["destination"] if row["destination"] in groups else "human_review"
        row["destination"] = dest
        groups[dest].append(row)
    overlap = {row["order_no"] for row in groups["dian_01"]} & {row["order_no"] for row in groups["dian_02"]}
    if overlap:
        anomalies.append({"type": "ROUTE_OVERLAP", "order_nos": sorted(overlap), "message": "同一单同时进入 01 和 02，已改到人工审核。"})
        for row in groups["dian_01"] + groups["dian_02"]:
            if row["order_no"] in overlap:
                row["destination"] = "human_review"
        groups["human_review"].extend([row for row in groups["dian_01"] + groups["dian_02"] if row["order_no"] in overlap])
        groups["dian_01"] = [row for row in groups["dian_01"] if row["order_no"] not in overlap]
        groups["dian_02"] = [row for row in groups["dian_02"] if row["order_no"] not in overlap]
    return groups


def _by_ops_group(rows: list[dict]) -> dict[str, list[dict]]:
    grouped: dict[str, list[dict]] = {}
    for row in rows:
        if not row.get("ops_group"):
            continue
        grouped.setdefault(str(row["ops_group"]), []).append(row)
    return grouped


def _write_xlsx(
    path: Path,
    headers: list[str],
    rows: list[dict],
    *,
    style_mode: str = "oversea",
    width_ref: Path | None = None,
    config: dict | None = None,
    anomalies: list | None = None,
) -> None:
    """写出 Order 形态 xlsx。

    有金标准/样板壳（width_ref）时：复制壳再填数，保留 theme1、styles、CF、列宽、行高。
    无壳时才新建工作簿（海外仓仍会后补 theme1_oversea.xml）。
    """
    shell = Path(width_ref) if width_ref else None
    if shell is not None and shell.is_file() and shell.suffix.lower() in {".xlsx", ".xlsm"}:
        _write_xlsx_into_shell(
            path, headers, rows, shell=shell, style_mode=style_mode, config=config, anomalies=anomalies
        )
        return

    from openpyxl.formatting.rule import Rule
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.styles.colors import Color
    from openpyxl.styles.differential import DifferentialStyle
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    font_color = Color(theme=1)
    header_font = Font(name="宋体", size=11, bold=True, color=font_color)
    data_font = Font(name="宋体", size=11, color=font_color)
    align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    fill_peach = PatternFill(patternType="solid", fgColor=Color(theme=5, tint=0.8))
    fill_green = PatternFill(patternType="solid", fgColor=Color(theme=7, tint=0.8))
    use_oversea_fill = style_mode == "oversea"
    fill_cfg = ((config or {}).get("oversea_fill") or {}) if use_oversea_fill else {}

    # 行高：海外仓与金标准人工表一致（defaultRowHeight=25 且 customHeight）；运营表显式每行 25
    if style_mode != "oversea":
        ws.sheet_format.defaultRowHeight = 13.5
        ws.row_dimensions[1].height = 25

    for col, header in enumerate(headers, start=1):
        cell = ws.cell(1, col, header if header is not None else "")
        cell.font = header_font
        cell.alignment = align

    green_noted: set[str] = set()
    for r_idx, row in enumerate(rows, start=2):
        if style_mode == "ops":
            ws.row_dimensions[r_idx].height = 25
        row_fill = None
        if use_oversea_fill:
            theme, via = _oversea_fill_theme(row, fill_cfg)
            row_fill = fill_green if theme == 7 else fill_peach
            if via == "order_no" and anomalies is not None:
                order_no = str(row.get("订单号") or "").strip()
                if order_no and order_no not in green_noted:
                    green_noted.add(order_no)
                    anomalies.append(
                        {
                            "type": "OVERSEA_GREEN_NON_EBAY",
                            "order_no": order_no,
                            "shop_account": row.get("店铺账号"),
                            "message": "非eBay却按 highlight_green_order_nos 标绿；规则待业务确认。",
                        }
                    )
        for col, header in enumerate(headers, start=1):
            value = _excel_value(row.get(header), header)
            # 空值先写 None（带样式），保存后再补成共享空串 ""
            cell = ws.cell(r_idx, col, value if value != "" else None)
            cell.font = data_font
            cell.alignment = align
            if row_fill is not None:
                cell.fill = row_fill

    if style_mode == "oversea":
        _apply_oversea_row_heights(ws, 1 + len(rows))

    widths = _column_widths_from_ref(width_ref, headers) if width_ref else _default_widths(headers, style_mode)
    for col, header in enumerate(headers, start=1):
        letter = get_column_letter(col)
        width = widths.get(header) or widths.get(letter)
        if width:
            ws.column_dimensions[letter].width = float(width)

    # 海外仓：订单号列 duplicateValues → 深橙 FFFF9900（分单/多SKU 同单号自动变色）
    if use_oversea_fill and "订单号" in headers:
        col_letter = get_column_letter(headers.index("订单号") + 1)
        dxf = DifferentialStyle(
            fill=PatternFill(patternType="solid", fgColor="FFFF9900")
        )
        ws.conditional_formatting.add(
            f"{col_letter}1:{col_letter}1048576",
            Rule(type="duplicateValues", dxf=dxf),
        )

    wb.save(path)
    wb.close()
    _patch_xlsx_empty_shared_strings(path)
    if use_oversea_fill:
        _patch_xlsx_oversea_theme(path)



def _write_xlsx_into_shell(
    path: Path,
    headers: list[str],
    rows: list[dict],
    *,
    shell: Path,
    style_mode: str = "oversea",
    config: dict | None = None,
    anomalies: list | None = None,
) -> None:
    """以人工/金标准 xlsx 为壳：字节复制后尽量只改数据单元格值，保留 theme/styles/CF/列宽/行高。"""
    import shutil

    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.styles.colors import Color

    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(shell, path)
    wb = load_workbook(path)
    ws = wb.active

    for col, header in enumerate(headers, start=1):
        cell = ws.cell(1, col)
        if cell.value != header:
            cell.value = header if header is not None else ""

    fill_cfg = ((config or {}).get("oversea_fill") or {}) if style_mode == "oversea" else {}
    use_oversea_fill = style_mode == "oversea"
    fill_peach = PatternFill(patternType="solid", fgColor=Color(theme=5, tint=0.8))
    fill_green = PatternFill(patternType="solid", fgColor=Color(theme=7, tint=0.8))
    data_font = Font(name="宋体", size=11, color=Color(theme=1))
    align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    old_data_rows = max(0, (ws.max_row or 1) - 1)
    # 隐藏行也算壳残留，先取消隐藏再按可见逻辑清空
    for r in range(2, (ws.max_row or 1) + 1):
        if ws.row_dimensions[r].hidden:
            ws.row_dimensions[r].hidden = False

    need_rebuild = old_data_rows != len(rows)
    if need_rebuild and ws.max_row and ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    green_noted: set[str] = set()
    for r_idx, row in enumerate(rows, start=2):
        if style_mode == "ops":
            ws.row_dimensions[r_idx].height = 25
        row_fill = None
        if use_oversea_fill:
            theme, via = _oversea_fill_theme(row, fill_cfg)
            row_fill = fill_green if theme == 7 else fill_peach
            if via == "order_no" and anomalies is not None:
                order_no = str(row.get("订单号") or "").strip()
                if order_no and order_no not in green_noted:
                    green_noted.add(order_no)
                    anomalies.append(
                        {
                            "type": "OVERSEA_GREEN_NON_EBAY",
                            "order_no": order_no,
                            "shop_account": row.get("店铺账号"),
                            "message": "非eBay却按 highlight_green_order_nos 标绿；规则待业务确认。",
                        }
                    )
        for col, header in enumerate(headers, start=1):
            value = _excel_value(row.get(header), header)
            cell = ws.cell(r_idx, col)
            cell.value = value if value != "" else None
            # 行数一致时只改值，保留壳上字体/填充；重建时补样式
            if need_rebuild:
                cell.font = data_font
                cell.alignment = align
                if row_fill is not None:
                    cell.fill = row_fill
            elif use_oversea_fill and row_fill is not None:
                # 壳行序可能与本次不同：仍按规则校正整行 theme，避免错色
                cell.fill = row_fill

    # 行数变少：清空多余旧行单元格
    if not need_rebuild and old_data_rows > len(rows):
        for r_idx in range(len(rows) + 2, old_data_rows + 2):
            for col in range(1, len(headers) + 1):
                ws.cell(r_idx, col).value = None

    if style_mode == "oversea":
        _apply_oversea_row_heights(ws, 1 + len(rows))

    wb.save(path)
    wb.close()
    _patch_xlsx_empty_shared_strings(path)
    if use_oversea_fill:
        _patch_xlsx_oversea_theme(path)


def _apply_oversea_row_heights(ws, last_row: int) -> None:
    """海外仓行高对齐金标准人工表：sheet defaultRowHeight=25 且 customHeight，每行显式 25。"""
    ws.sheet_format.defaultRowHeight = OVERSEA_ROW_HEIGHT
    ws.sheet_format.customHeight = True
    for r_idx in range(1, max(int(last_row), 1) + 1):
        ws.row_dimensions[r_idx].height = OVERSEA_ROW_HEIGHT


def _oversea_fill_theme(row: dict, fill_cfg: dict) -> tuple[int, str]:
    """返回 (theme, via)。via=shop|order_no|default。金标准：eBay→7；其余默认5；非eBay特例走 highlight_green_order_nos。"""
    shop = str(row.get("店铺账号") or "")
    order_no = str(row.get("订单号") or "").strip()
    shop_needles = [str(x) for x in (fill_cfg.get("theme7_shop_contains") or ["ebay"])]
    if any(needle.lower() in shop.lower() for needle in shop_needles if needle):
        return 7, "shop"
    green_orders = {str(x).strip() for x in (fill_cfg.get("highlight_green_order_nos") or []) if str(x).strip()}
    if order_no and order_no in green_orders:
        return 7, "order_no"
    default_theme = int(fill_cfg.get("default_theme") or 5)
    return default_theme, "default"


def _patch_xlsx_empty_shared_strings(path: Path) -> None:
    """openpyxl 会把 '' 存成空数值格；补成共享空串，读回为 ''（贴近金标准体积/空单元格）。"""
    import io
    import re
    import zipfile

    path = Path(path)
    with zipfile.ZipFile(path, "r") as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    ss_name = "xl/sharedStrings.xml"
    empty_idx = 0
    ss_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'count="1" uniqueCount="1"><si><t></t></si></sst>'
    )
    files[ss_name] = ss_xml.encode("utf-8")

    ct = files["[Content_Types].xml"].decode("utf-8")
    if "sharedStrings" not in ct:
        ct = ct.replace(
            "</Types>",
            '<Override PartName="/xl/sharedStrings.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
            "</Types>",
        )
        files["[Content_Types].xml"] = ct.encode("utf-8")

    rels_name = "xl/_rels/workbook.xml.rels"
    rels = files[rels_name].decode("utf-8")
    if "sharedStrings.xml" not in rels:
        ids = [int(x) for x in re.findall(r"rId(\d+)", rels)]
        nid = (max(ids) + 1) if ids else 1
        rels = rels.replace(
            "</Relationships>",
            f'<Relationship Id="rId{nid}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
            'Target="sharedStrings.xml"/></Relationships>',
        )
        files[rels_name] = rels.encode("utf-8")

    sheet = "xl/worksheets/sheet1.xml"
    xml = files[sheet].decode("utf-8")

    def repl_n(match: re.Match) -> str:
        return f'<c r="{match.group(1)}" s="{match.group(2)}" t="s"><v>{empty_idx}</v></c>'

    # openpyxl 空单元格常见写法：<c r="K2" s="2" t="n" /> 或 t="n"></c> 或纯自闭合
    xml = re.sub(
        r'<c r="([^"]+)" s="([^"]+)" t="n"\s*/>',
        repl_n,
        xml,
    )
    xml = re.sub(r'<c r="([^"]+)" s="([^"]+)" t="n"></c>', repl_n, xml)
    xml = re.sub(
        r'<c r="([^"]+)" s="([^"]+)"\s*/>',
        lambda m: f'<c r="{m.group(1)}" s="{m.group(2)}" t="s"><v>{empty_idx}</v></c>',
        xml,
    )
    files[sheet] = xml.encode("utf-8")

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    path.write_bytes(buf.getvalue())


_OVERSEA_THEME_PATH = Path(__file__).resolve().parent.parent / "templates" / "theme1_oversea.xml"


def _patch_xlsx_oversea_theme(path: Path, theme_path: Path | None = None) -> None:
    """把金标准海外仓 theme1.xml 打进 xlsx。

    openpyxl 默认 theme：accent2=C0504D（粉）、accent4=8064A2（紫）。
    人工表 accent2=EE822F（橙）、accent4=75BD42（绿）。
    单元格仍写 theme=5/7 tint=0.8；换主题后打开才是浅橙/浅绿，而不是粉/紫。
    """
    import io
    import zipfile

    path = Path(path)
    theme_path = Path(theme_path) if theme_path else _OVERSEA_THEME_PATH
    if not theme_path.is_file():
        return
    theme_bytes = theme_path.read_bytes()
    with zipfile.ZipFile(path, "r") as zin:
        files = {name: zin.read(name) for name in zin.namelist()}
    files["xl/theme/theme1.xml"] = theme_bytes
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    path.write_bytes(buf.getvalue())


def _excel_value(value, header: str):
    """尽量贴近人工表类型：空值写 ""；产品总数/报关重量为整数，申报单价保持文本观感。"""
    if value is None:
        return ""
    if header in {"产品总数", "报关重量"}:
        try:
            num = float(str(value).strip())
            if abs(num - int(num)) < 1e-9:
                return int(num)
        except (TypeError, ValueError):
            pass
    if header == "申报单价" and value != "":
        text = str(value).strip()
        try:
            num = float(text)
            if abs(num - int(num)) < 1e-9:
                return f"{int(num)}.0"
            return text
        except ValueError:
            return text
    return value


def _default_widths(headers: list[str], style_mode: str) -> dict[str, float]:
    oversea = {
        "下单时间": 9.0,
        "店铺账号": 25.375,
        "订单号": 36.625,
        "买家指定物流": 9.0,
        "物流方式": 61.125,
        "运单号": 22.625,
        "SKU": 9.0,
        "拣货备注": 53.75,
        "买家账号": 9.0,
    }
    ops = {
        "下单时间": 13.0,
        "店铺账号": 25.375,
        "订单号": 36.625,
        "物流方式": 40.0,
        "运单号": 22.0,
        "SKU": 14.0,
        "拣货备注": 40.0,
    }
    return oversea if style_mode == "oversea" else ops


def _column_widths_from_ref(ref: Path, headers: list[str]) -> dict[str, float]:
    from openpyxl.utils import get_column_letter

    try:
        wb = load_workbook(ref)
    except Exception:
        return {}
    ws = wb.active
    widths: dict[str, float] = {}
    for col, header in enumerate(headers, start=1):
        letter = get_column_letter(col)
        dim = ws.column_dimensions.get(letter)
        if dim is not None and dim.width:
            widths[header] = float(dim.width)
    wb.close()
    return widths


def _reorder_like_ref(rows: list[dict], ref: Path) -> list[dict]:
    """按人工样本的（订单号, SKU）顺序重排；对不上的追加在末尾。"""
    if not ref.is_file() or not rows:
        return rows
    try:
        wb = load_workbook(ref, data_only=True)
    except Exception:
        return rows
    ws = wb.active
    ref_headers = [str(ws.cell(1, col).value or "").strip() for col in range(1, ws.max_column + 1)]
    if "订单号" not in ref_headers or "SKU" not in ref_headers:
        wb.close()
        return rows
    oi = ref_headers.index("订单号") + 1
    si = ref_headers.index("SKU") + 1

    def norm_sku(value) -> str:
        return str(value or "").replace("\r\n", "\n").replace("\r", "\n").strip()

    order_keys: list[tuple[str, str]] = []
    for idx in range(2, ws.max_row + 1):
        order_no = str(ws.cell(idx, oi).value or "").strip()
        sku = norm_sku(ws.cell(idx, si).value)
        if not order_no and not sku:
            continue
        order_keys.append((order_no, sku))
    wb.close()

    buckets: dict[tuple[str, str], list[dict]] = {}
    for row in rows:
        key = (str(row.get("订单号") or "").strip(), norm_sku(row.get("SKU")))
        buckets.setdefault(key, []).append(row)

    ordered: list[dict] = []
    used: set[int] = set()
    for key in order_keys:
        pile = buckets.get(key) or []
        if not pile:
            continue
        item = pile.pop(0)
        ordered.append(item)
        used.add(id(item))
    for row in rows:
        if id(row) not in used:
            ordered.append(row)
    return ordered


def _write_shared_template(out_dir: Path, label: str, rows: list[dict], template: str | None, config: dict | None = None, export_date: str | None = None, anomalies: list | None = None) -> Path | None:
    """店管家正式 xls：01 按样板（xlutils）；02 金标准壳 + 只改值 + colour26=#F2DCDB。"""
    if not rows or not template or not Path(template).is_file():
        return None

    if label.endswith("02") or label == "店管家02":
        path = out_dir / f"{label}.xls"
        return _write_dian_02_xls_shell(
            path,
            Path(template),
            rows,
            config=config or {},
            export_date=export_date,
            anomalies=anomalies,
        )

    import xlrd
    from xlutils.copy import copy as xl_copy

    book = xlrd.open_workbook(template, formatting_info=True)
    sheet0 = book.sheet_by_index(0)
    headers = [sheet0.cell_value(0, col) for col in range(sheet0.ncols)]
    writable = xl_copy(book)
    sheet = writable.get_sheet(0)

    style_text = _xls_style(height=220, num_format_str="@", wrap=1)
    style_content = _xls_style(height=220, num_format_str="General", wrap=1)
    style_empty = _xls_style(height=180, num_format_str="@", horz=xlwt.Alignment.HORZ_GENERAL, vert=xlwt.Alignment.VERT_BOTTOM, wrap=0)
    content_cols = {"原始单号", "客户备注", "客服备注", "自定义1", "订单编号"}
    for index, row in enumerate(rows, start=1):
        order_no = _d01_original_order_no(row)
        tracking = "\n".join(row["tracking_nos"])
        logistics = _text(row["cells"].get("物流方式"))
        values = {
            "订单编号": "",
            "原始单号": order_no,
            "客户备注": tracking,
            "客服备注": logistics,
            "自定义1": f"{index:02d}",
        }
        for col, header in enumerate(headers):
            header_s = str(header)
            value = values.get(header_s, "")
            if header_s in ("原始单号", "客户备注", "自定义1", "订单编号"):
                sheet.write(index, col, str(value) if value != "" else "", style_text)
            elif header_s in content_cols:
                sheet.write(index, col, value if value != "" else "", style_content if value != "" else style_empty)
            else:
                sheet.write(index, col, "", style_empty)
        _xls_set_row_height(sheet, index, 540)
    path = out_dir / f"{label}.xls"
    writable.save(str(path))
    return path


def _write_dian_02_xls_shell(
    out_xls: Path,
    template_xls: Path,
    rows: list[dict],
    *,
    config: dict,
    export_date: str | None,
    anomalies: list | None,
) -> Path:
    """金标准壳 copy → xlutils 只改值（复用原 XF）→ colour26 固定 #F2DCDB 粉底粉。

    禁止 soffice 转格式；禁止刷 #FFFFCC；多余壳行只清空值、保留 XF。
    """
    import xlrd
    from xlutils.copy import copy as xl_copy

    out_xls = Path(out_xls)
    out_xls.parent.mkdir(parents=True, exist_ok=True)

    book = xlrd.open_workbook(str(template_xls), formatting_info=True)
    sheet0 = book.sheet_by_index(0)
    headers = [sheet0.cell_value(0, col) for col in range(sheet0.ncols)]
    shell_nrows = sheet0.nrows
    writable = xl_copy(book)
    # 用户 Excel/WPS 要粉底粉；BIFF 默认 colour26 是黄，这里改成 #F2DCDB
    writable.set_colour_RGB(26, 0xF2, 0xDC, 0xDB)
    sheet = writable.get_sheet(0)
    _fill_dian_02_sheet(
        sheet,
        headers,
        rows,
        config=config,
        export_date=export_date,
        anomalies_out=anomalies,
        file_name=out_xls.name,
    )
    # 壳里多余旧数据行：只清空值，保留 XF/行高
    proto = _xls_prototype_xf(sheet, len(headers), proto_row=1)
    for r in range(len(rows) + 1, shell_nrows):
        for c in range(len(headers)):
            _xls_write_keep_xf(sheet, r, c, "", fallback_xf=proto[c])
    writable.save(str(out_xls))
    return out_xls



def _d01_original_order_no(row: dict) -> str:
    """eBay 原始单号按 '-' 取最后一段；其余保持完整订单号。"""
    order_no = _text(row.get("order_no"))
    shop = _text(row.get("shop") or (row.get("cells") or {}).get("店铺账号"))
    platform = _text(row.get("platform") or (row.get("cells") or {}).get("平台"))
    blob = f"{shop} {platform}".lower()
    if "ebay" in blob and "-" in order_no:
        return order_no.rsplit("-", 1)[-1]
    return order_no



def _xls_prototype_xf(sheet, ncols: int, *, proto_row: int = 1) -> list[int]:
    """从壳上某一数据行取出每列 xf_idx，供新行/缺格回退（保住粉底）。"""
    rows = sheet._Worksheet__rows
    row = rows.get(proto_row)
    out: list[int] = []
    for col in range(ncols):
        xf = 0x0F
        if row is not None:
            cell = row._Row__cells.get(col)
            if cell is not None and getattr(cell, "xf_idx", None) is not None:
                xf = int(cell.xf_idx)
        out.append(xf)
    return out


def _xls_write_keep_xf(sheet, rowx: int, colx: int, value, *, fallback_xf: int = 0x0F) -> None:
    """只改值，保留原单元格 XF；无原格则用 fallback_xf。绝不 add_style。"""
    from decimal import Decimal

    from xlwt.Cell import BlankCell, BooleanCell, NumberCell, StrCell

    rows = sheet._Worksheet__rows
    if rowx not in rows:
        sheet.row(rowx)
    row = rows[rowx]
    cells = row._Row__cells
    prev = cells.get(colx)
    xf = getattr(prev, "xf_idx", None)
    if xf is None:
        xf = int(fallback_xf)
    row._Row__adjust_bound_col_idx(colx)
    parent_wb = row._Row__parent_wb
    if prev is not None:
        sst_idx = getattr(prev, "sst_idx", None)
        if sst_idx is not None:
            parent_wb.del_str(sst_idx)
    if value is None or value == "":
        cells[colx] = BlankCell(rowx, colx, xf)
    elif isinstance(value, bool):
        cells[colx] = BooleanCell(rowx, colx, xf, value)
    elif isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
        cells[colx] = NumberCell(rowx, colx, xf, float(value))
    else:
        cells[colx] = StrCell(rowx, colx, xf, parent_wb.add_str(str(value)))


def _xls_style(
    *,
    name: str = "宋体",
    height: int = 220,
    colour_index: int = 32767,
    num_format_str: str = "General",
    horz: int | None = None,
    vert: int | None = None,
    wrap: int = 1,
):
    import xlwt

    st = xlwt.XFStyle()
    fnt = xlwt.Font()
    fnt.name = name
    fnt.height = height
    fnt.colour_index = colour_index
    st.font = fnt
    st.num_format_str = num_format_str
    al = xlwt.Alignment()
    al.horz = xlwt.Alignment.HORZ_CENTER if horz is None else horz
    al.vert = xlwt.Alignment.VERT_CENTER if vert is None else vert
    al.wrap = wrap
    st.alignment = al
    return st


def _xls_set_row_height(sheet, row_idx: int, height_twips: int) -> None:
    row = sheet.row(row_idx)
    row.height = int(height_twips)
    row.height_mismatch = 1


def _resolve_export_dates(rows: list[dict], config: dict, export_date: str | None) -> tuple[str, str]:
    """返回 (2026.10.4, 10.4)。优先 export_date / 配置，否则取下单时间，再退回当天。"""
    naming = config.get("export_naming") or {}
    raw = export_date or naming.get("date_full") or naming.get("date")
    if not raw:
        candidates = []
        for row in rows:
            cells = row.get("cells") or {}
            for key in ("下单时间", "付款时间"):
                text = _text(cells.get(key))
                if text:
                    candidates.append(text[:10])
        raw = max(candidates) if candidates else None
    if not raw:
        from datetime import datetime
        from zoneinfo import ZoneInfo

        raw = datetime.now(ZoneInfo("Asia/Shanghai")).strftime("%Y-%m-%d")
    text = str(raw).strip().replace("/", "-").replace(".", "-")
    parts = text.split("-")
    if len(parts) >= 3 and parts[0].isdigit():
        year, month, day = parts[0], str(int(parts[1])), str(int(parts[2][:2]))
    else:
        from datetime import datetime
        from zoneinfo import ZoneInfo

        now = datetime.now(ZoneInfo("Asia/Shanghai"))
        year, month, day = str(now.year), str(now.month), str(now.day)
    date_full = f"{year}.{month}.{day}"
    date_short = f"{month}.{day}"
    if naming.get("date_short"):
        date_short = str(naming["date_short"])
    if naming.get("date_full") and not export_date:
        date_full = str(naming["date_full"])
        if "." in date_full:
            bits = date_full.split(".")
            if len(bits) >= 3:
                date_short = f"{bits[1]}.{bits[2]}"
    return date_full, date_short


def _write_official(out_dir: Path, label: str, rows: list[dict], column_map: list[dict]) -> Path | None:
    pairs = []
    for item in column_map:
        source = str(item.get("source") or "").strip()
        target = str(item.get("target") or "").strip()
        if source and target:
            pairs.append((source, target))
    if not pairs:
        return None
    path = out_dir / f"{label}.xls"
    book = xlwt.Workbook()
    sheet = book.add_sheet("sheet1")
    for col, (_source, target) in enumerate(pairs):
        sheet.write(0, col, target)
    for index, row in enumerate(rows, start=1):
        for col, (source, _target) in enumerate(pairs):
            value = row["cells"].get(source)
            sheet.write(index, col, "" if value is None else value)
    book.save(str(path))
    return path


def _safe_name(value: str) -> str:
    cleaned = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in value.strip())
    return cleaned or "group"



def _normalize_country(value, config: dict) -> str:
    """按 routing country_normalize 映射国名（精确匹配，大小写不敏感）。"""
    text = _text(value)
    mapping = config.get("country_normalize") or {}
    if not text or not mapping:
        return text
    upper = text.upper()
    for src, dst in mapping.items():
        if str(src).strip().upper() == upper:
            return str(dst).strip()
    return text


def _text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and abs(value - int(value)) < 1e-9:
        return str(int(value))
    return str(value).strip()


def _decimal(value) -> Decimal | None:
    if value is None or value == "":
        return None
    try:
        return Decimal(str(value))
    except (InvalidOperation, ValueError):
        return None


def _money(value: Decimal | None) -> str | None:
    if value is None:
        return None
    return format(value.quantize(Decimal("0.01")), "f")


def _fill_dian_02_sheet(sheet, headers: list, rows: list[dict], *, config: dict, export_date: str | None, anomalies_out: list | None, file_name: str | None = None) -> None:
    """按金标准填写店管家02：地址/仓库/简化物流/Amazon备注/分单-1/自定义1。

    合计金额：拆单未人工填→留空；中国直发未匹配店小秘→留空；禁止用海外仓金额顶、禁止按比例猜、禁止写0冒充。
    只改单元格值，复用壳上原 XF（pattern_colour=26，写出前 set_colour_RGB→#F2DCDB）；不新建 style、不改行高。
    """
    proto = _xls_prototype_xf(sheet, len(headers), proto_row=1)
    yy, month, day = _custom1_date_parts(export_date)
    order_counts: dict[str, int] = {}
    for row in rows:
        ono = _text(row.get("order_no"))
        if ono:
            order_counts[ono] = order_counts.get(ono, 0) + 1
    seen_orders: dict[str, int] = {}
    out_file = file_name or "店管家02.xls"
    for index, row in enumerate(rows, start=1):
        # xlwt 行 index：0=表头 → Excel 第1行；数据从 index=1 起 → Excel 第 index+1 行
        excel_row = index + 1
        values = _dian_02_values(
            row,
            config,
            seen_orders,
            custom1=f"{yy}{month}{day}{index:03d}",
            anomalies_out=anomalies_out,
            order_counts=order_counts,
            excel_row=excel_row,
            file_name=out_file,
        )
        for col, header in enumerate(headers):
            header_s = str(header)
            value = values.get(header_s, "")
            if value is None:
                value = ""
            _xls_write_keep_xf(sheet, index, col, value, fallback_xf=proto[col])


def _custom1_date_parts(export_date: str | None) -> tuple[str, str, str]:
    """2026.10.4 / 2026-10-04 → ('26', '10', '4')，月日不补零。"""
    from datetime import datetime
    from zoneinfo import ZoneInfo

    text = str(export_date or "").strip().replace("/", "-").replace(".", "-")
    parts = [p for p in text.split("-") if p]
    if len(parts) >= 3 and parts[0].isdigit():
        year, month, day = parts[0], str(int(parts[1])), str(int(parts[2][:2]))
        return year[-2:], month, day
    now = datetime.now(ZoneInfo("Asia/Shanghai"))
    return str(now.year)[-2:], str(now.month), str(now.day)


def _dian_02_values(
    row: dict,
    config: dict,
    seen_orders: dict[str, int],
    *,
    custom1: str,
    anomalies_out: list | None,
    order_counts: dict[str, int] | None = None,
    excel_row: int | None = None,
    file_name: str | None = None,
) -> dict:
    cells = row.get("cells") or {}
    base_no = _text(row.get("order_no"))
    seen_orders[base_no] = seen_orders.get(base_no, 0) + 1
    seq = seen_orders[base_no]
    order_no = base_no if seq == 1 else f"{base_no}-{seq - 1}"

    source_logistics = _text(row.get("raw_logistics") or cells.get("物流方式"))
    logistics = _text(cells.get("物流方式"))
    tracking = "\n".join(row.get("tracking_nos") or _split_lines(cells.get("运单号")))
    package = _text(cells.get("包裹号"))
    amazon = _is_amazon_logistics(source_logistics or logistics, config)

    # 店管家02 物流映射按原文物流，不受海外仓正本 Amazon 运单号/物流追加影响。
    # 客户备注/物流单号仍用原运单号（tracking_nos 在正本改写前已保留）。
    if not tracking and package and anomalies_out is not None:
        item = {
            "type": "D02_TRACKING_EMPTY_PACKAGE_ONLY",
            "order_no": order_no,
            "message": f"运单号为空，只有包裹号 {package}；不用包裹号顶运单号，待补真实运单号。",
        }
        if excel_row is not None:
            item["excel_row"] = excel_row
        if file_name:
            item["file"] = file_name
        anomalies_out.append(item)
    cust = tracking
    service = logistics

    # 店管家02「仓库」列故意留空；「物流方式」仍按原文物流走 warehouse_logistics_map
    _wh_mapped, simple_log = _map_warehouse_logistics(source_logistics, config)
    warehouse = ""
    if source_logistics and not simple_log and anomalies_out is not None:
        item = {
            "type": "D02_WAREHOUSE_UNMAPPED",
            "order_no": order_no,
            "message": f"物流方式未命中 warehouse_logistics_map：{source_logistics}",
        }
        if excel_row is not None:
            item["excel_row"] = excel_row
        if file_name:
            item["file"] = file_name
        anomalies_out.append(item)

    sku = _text(cells.get("SKU"))
    qty = cells.get("产品总数")
    try:
        qty_num = float(str(qty).strip()) if qty not in (None, "") else ""
        if qty_num != "" and abs(qty_num - int(qty_num)) < 1e-9:
            qty_num = int(qty_num)
    except (TypeError, ValueError):
        qty_num = qty if qty is not None else ""

    amt_num = _d02_total_amount(
        row,
        cells,
        order_counts=order_counts or {},
        anomalies_out=anomalies_out,
        order_no=order_no,
        excel_row=excel_row,
        file_name=file_name,
    )

    name = _text(cells.get("收货人姓名"))
    phone = _text(cells.get("收货人电话")) or _text(cells.get("收货人手机"))
    email = _text(cells.get("买家Email")) or _text(cells.get("接收邮箱"))
    product_name = _text(cells.get("中文报关名")) or _text(cells.get("英文报关名"))
    unit_price = _text(cells.get("申报单价"))
    return {
        "订单编号": order_no,
        "客户网名": name,
        "收货人": name,
        "电话": phone,
        "国家/地区": _normalize_country(cells.get("收货人国家"), config),
        "州省": _text(cells.get("收货人州/省")),
        "区市": _text(cells.get("收货人城市")),
        "地址": _text(cells.get("详细地址")),
        "邮编": _text(cells.get("邮编")),
        "Email": email,
        "接收邮箱": email,
        "客服备注": service,
        "客户备注": cust,
        "编号": sku,
        "品名": product_name,
        "条码": sku,
        "数量": qty_num,
        "单价": unit_price,
        "合计": amt_num,
        "订单来源": str((config.get("d02_rules") or {}).get("order_source") or "手工建单"),
        "店铺": _text(row.get("shop") or cells.get("店铺账号")),
        "仓库": warehouse,
        "物流方式": simple_log,
        "物流单号": cust,
        "自定义1": custom1,
    }



def _d02_total_amount(
    row: dict,
    cells: dict,
    *,
    order_counts: dict[str, int],
    anomalies_out: list | None,
    order_no: str,
    excel_row: int | None = None,
    file_name: str | None = None,
):
    """店管家02「合计」。

    - 中国直发：仅用店小秘/金额原件回填；匹配不到→留空+审核，不写0、不用海外仓金额。
    - 拆单（同订单号多行）且无 manual_sku_amount：留空+审核，禁止按比例猜、禁止用海外仓总额顶。
    - 其余：有明确非零金额才写入；金额为 0 写成空（禁止写出 0 / 0.0）。
    """
    china = bool(row.get("china_direct"))
    matched = bool(row.get("amount_matched"))
    base_no = _text(row.get("order_no"))
    is_split = bool(base_no) and int(order_counts.get(base_no) or 0) > 1
    manual = row.get("manual_sku_amount")
    if manual is None:
        manual = (row.get("cells") or {}).get("manual_sku_amount")

    if china and not matched:
        if anomalies_out is not None:
            item = {
                "type": "D02_AMOUNT_LEFT_EMPTY",
                "task_type": "china_direct_amount_unmatched",
                "order_no": order_no,
                "reason": "china_direct_unmatched",
                "message": "中国直发合计留空（无店小秘/金额原件回填）。",
            }
            if excel_row is not None:
                item["excel_row"] = excel_row
            if file_name:
                item["file"] = file_name
            anomalies_out.append(item)
        return ""

    if is_split and manual in (None, ""):
        if anomalies_out is not None:
            item = {
                "type": "D02_AMOUNT_LEFT_EMPTY",
                "task_type": "split_sku_amount_manual",
                "order_no": order_no,
                "reason": "split_unfilled",
                "message": "拆单合计留空，等人工按 SKU 填金额。",
            }
            if excel_row is not None:
                item["excel_row"] = excel_row
            if file_name:
                item["file"] = file_name
            anomalies_out.append(item)
        return ""

    if manual not in (None, ""):
        try:
            return float(str(manual).strip())
        except (TypeError, ValueError):
            return manual

    amount = row.get("amount")
    if amount is None:
        amount = cells.get("订单金额")
    if amount in (None, ""):
        return ""
    try:
        num = float(str(amount).strip())
    except (TypeError, ValueError):
        return amount
    # 中国直发已匹配：写真实回填金额（可来自店小秘/金额原件）
    if china and matched:
        return num
    # 非拆单普通行：金额为 0 留空，禁止写出 0 / 0.0（如手工单海外仓正本金额常为 0）
    if abs(num) < 1e-12:
        return ""
    return num


def _is_amazon_logistics(logistics: str, config: dict) -> bool:
    prefixes = (config.get("d02_rules") or {}).get("amazon_logistics_prefixes") or ["Amazon FBA", "Amazon-"]
    text = logistics or ""
    return any(text.startswith(str(p)) for p in prefixes)


def _map_warehouse_logistics(logistics: str, config: dict) -> tuple[str, str]:
    """物流原文 → 店管家02「物流方式」简称。

    先走 warehouse_logistics_map（contains，不区分大小写）；命中不到再走规律：
    - Amazon/Amazon FBA：按文中「X站」或站点国名 → 亚马逊FBA+国家（不记店铺名）
    - 含「云途」→ 云途；含美客多/MERCADO → 美客多无忧（也可写在 yaml）
    """
    import re

    text = (logistics or "").casefold()
    for rule in config.get("warehouse_logistics_map") or []:
        needle = str(rule.get("contains") or "").casefold()
        if needle and needle in text:
            return str(rule.get("warehouse") or ""), str(rule.get("logistics") or "")
    raw = logistics or ""
    if _is_amazon_logistics(raw, config):
        m = re.search(r"(德国|英国|美国|法国|意大利|西班牙|澳大利亚|日本|加拿大|墨西哥)站", raw)
        if not m:
            m = re.search(
                r"(?:-|—)(德国|英国|美国|法国|意大利|西班牙|澳大利亚|日本|加拿大|墨西哥)(?:$|[-—（( ])",
                raw,
            )
        if m:
            return "", f"亚马逊FBA{m.group(1)}"
    return "", ""


def build_oversea_copy_rows(oversea_rows: list[dict], copy_order_nos: list[str]) -> list[dict]:
    """从正本行挑出勾选订单，套用副本已确认加工。

    oversea_copy.mode=selected_orders
    rules: append_old_tracking_to_logistics + tracking=package_no
    同一订单号多 SKU / 多行全部带上；不写隐藏行。
    """
    wanted = {str(x).strip() for x in copy_order_nos if str(x).strip()}
    if not wanted:
        return []
    out: list[dict] = []
    for row in oversea_rows:
        order_no = _text(row.get("订单号"))
        if order_no not in wanted:
            continue
        out.append(_transform_oversea_copy_row(row))
    return out


def _swap_tracking_to_package(cells: dict) -> None:
    """运单号=包裹号；物流方式后追加原运单号（非空且与包裹号不同）。已含运单号则不重复追加。"""
    logistics = _text(cells.get("物流方式"))
    tracking = _text(cells.get("运单号"))
    package = _text(cells.get("包裹号"))
    if tracking and tracking != package and tracking not in logistics:
        cells["物流方式"] = f"{logistics} {tracking}".strip()
    if package:
        cells["运单号"] = package


def _transform_oversea_copy_row(row: dict) -> dict:
    """选中行：运单号=包裹号；物流方式后追加原运单号（非空且与包裹号不同）。"""
    cells = dict(row)
    _swap_tracking_to_package(cells)
    return cells


def write_oversea_copy_from_formal(
    formal_path: Path,
    out_path: Path,
    copy_order_nos: list[str],
    *,
    width_ref: Path | None = None,
) -> dict:
    """从已生成的海外仓正本按勾选订单写出副本（供 API/二次导出）。"""
    headers, rows, _hidden = _read_visible(formal_path)
    # _read_visible returns rows as {cells, ...}; flatten to cells dicts
    cell_rows = [r["cells"] for r in rows]
    copy_rows = build_oversea_copy_rows(cell_rows, copy_order_nos)
    _write_xlsx(out_path, headers, copy_rows, style_mode="oversea", width_ref=width_ref or formal_path, config=None)
    return {
        "path": out_path.name,
        "copy_order_nos": [str(x).strip() for x in copy_order_nos if str(x).strip()],
        "copy_row_count": len(copy_rows),
        "headers": headers,
    }


def _write_oversea_copy(
    path: Path,
    headers: list[str],
    rows: list[dict],
    *,
    style_mode: str = "oversea",
    width_ref: Path | None = None,
    filter_ref: Path | None = None,
    copy_order_nos: list[str] | None = None,
) -> int:
    """兼容旧调用：仅写出勾选订单的可见副本行（不再全量+隐藏）。"""
    selected = [str(x).strip() for x in (copy_order_nos or []) if str(x).strip()]
    if not selected:
        return 0
    copy_rows = build_oversea_copy_rows(rows, selected)
    _write_xlsx(path, headers, copy_rows, style_mode=style_mode, width_ref=width_ref, config=None)
    return len(copy_rows)


def _reorder_routed_like_ref(rows: list[dict], ref: Path) -> list[dict]:
    """按金标准海外仓（订单号, SKU）顺序重排 routed 行。"""
    cells_list = [row.get("cells") or {} for row in rows]
    ordered_cells = _reorder_like_ref(cells_list, ref)
    # map back by id of cells dict — cells are same objects
    by_id = {id(row.get("cells")): row for row in rows}
    ordered = []
    used = set()
    for cells in ordered_cells:
        row = by_id.get(id(cells))
        if row is not None and id(row) not in used:
            ordered.append(row)
            used.add(id(row))
    for row in rows:
        if id(row) not in used:
            ordered.append(row)
    return ordered
