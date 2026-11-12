# -*- coding: utf-8 -*-
"""转换流水线：读原始 Order → 排序/改写 → 写三表 + 统计/异常/映射。"""
from __future__ import annotations

import csv
import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Literal

from .config import Rules, load_rules
from .io_excel import load_xlsx_rows, norm
from .resources_path import default_template_d01, default_template_d02, default_template_dx
from .transform import (
    ERROR_TYPES,
    REVIEW_TYPES,
    Anomaly,
    classify_dest,
    d01_original_order_no,
    d02_amount,
    d02_kefu_note,
    d02_phone,
    dedupe_d02_order_nos,
    is_amazon,
    is_china_direct,
    is_sku_like,
    find_sku_from_note,
    parse_quantity,
    sort_raw_rows,
    split_sku_lines,
    transform_logistics_dx,
    transform_order_no_dx,
    transform_sku,
    validate_duplicates,
    validate_headers,
)
from .writers import write_dian_01, write_dian_02, write_dianxiaomi

ExportStatus = Literal["ok", "needs_review", "blocked"]


@dataclass
class ConvertResult:
    dx_path: Path | None = None
    d01_path: Path | None = None
    d02_path: Path | None = None
    stats: dict[str, Any] = field(default_factory=dict)
    anomalies: list[dict] = field(default_factory=list)
    mappings: list[dict] = field(default_factory=list)
    pending_orders: list[dict] = field(default_factory=list)
    formal_export_blocked: bool = False
    export_status: ExportStatus = "ok"
    errors: list[dict] = field(default_factory=list)
    warnings: list[dict] = field(default_factory=list)
    infos: list[dict] = field(default_factory=list)


def convert_order_file(
    raw_path: Path,
    out_dir: Path,
    *,
    rules: Rules | None = None,
    template_d01: Path | None = None,
    template_d02: Path | None = None,
    template_dx: Path | None = None,
    exports: Iterable[str] = ("dx", "d01", "d02"),
    write_sidecars: bool = True,
    allow_blocked_review_export: bool = False,
) -> ConvertResult:
    """
    导出状态三分法：
      - ok：正式写出 out/dianxiaomi.xlsx、dian_01.xls、dian_02.xls
      - needs_review：未知SKU/多SKU等 → 仅写 out/review/*_NEEDS_REVIEW.*，
        正式路径不得出现可误导入文件名
      - blocked：严重错误 → 不写正式 01/02；可选 --allow-review-export 写 review

    allow_blocked_review_export:
        仅对 blocked 生效；needs_review 时默认就会写 review 目录。
    """
    rules = rules or load_rules()
    exports = set(exports)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    template_d01 = Path(template_d01) if template_d01 else default_template_d01()
    template_d02 = Path(template_d02) if template_d02 else default_template_d02()
    template_dx = Path(template_dx) if template_dx else default_template_dx()

    headers, raw_rows = load_xlsx_rows(Path(raw_path))
    anomalies: list[Anomaly] = []

    validate_headers(headers, anomalies)
    if not raw_rows:
        anomalies.append(
            Anomaly(
                type="EMPTY_INPUT",
                severity="error",
                订单号="",
                action="严重错误：原始表无订单数据行",
            )
        )
    validate_duplicates(raw_rows, anomalies)

    hard_stop = any(a.type in ("MISSING_COLUMNS", "EMPTY_INPUT") for a in anomalies)

    dx_rows: list[dict] = []
    d01_rows: list[dict] = []
    d02_rows: list[dict] = []
    mappings: list[dict] = []
    pending: list[dict] = []
    d01_line = 0
    d02_line = 0

    if not hard_stop:
        sorted_raw = sort_raw_rows(raw_rows, rules, anomalies)
        for seq, raw in enumerate(sorted_raw, start=1):
            dest = classify_dest(raw, rules, anomalies)
            dx_sku, base_sku = transform_sku(raw, rules, anomalies)
            dx_ono = transform_order_no_dx(raw, rules)
            dx_log = transform_logistics_dx(raw, rules, base_sku)

            qty = parse_quantity(raw, anomalies)
            # 输出基础 SKU 为中国直发 sku（M00016）时，数量一律改为 1；不报错、不阻断
            if base_sku == rules.china_sku:
                qty = 1
            amt = d02_amount(raw, rules, anomalies)
            sku_lines = split_sku_lines(raw.get("SKU"))
            is_multi = len(sku_lines) > 1 and not is_china_direct(raw, rules)
            note_hit = (
                None
                if is_multi or is_china_direct(raw, rules)
                else find_sku_from_note(raw, rules)
            )

            dx_row = {h: raw.get(h) for h in headers}
            dx_row["订单号"] = dx_ono
            dx_row["物流方式"] = dx_log
            dx_row["SKU"] = dx_sku
            if "产品总数" in dx_row:
                dx_row["产品总数"] = qty
            if "报关重量" in dx_row and isinstance(dx_row["报关重量"], float):
                w = dx_row["报关重量"]
                if abs(w - int(w)) < 1e-9:
                    dx_row["报关重量"] = int(w)
            dx_row["_all_red"] = is_amazon(raw, rules)
            dx_row["_seq"] = seq
            dx_row["_track"] = norm(raw.get("运单号"))
            dx_row["_dest"] = dest
            dx_row["_base_sku"] = base_sku
            dx_rows.append(dx_row)

            d01_excel_row = None
            d02_excel_row = None

            if dest == "D01":
                d01_line += 1
                d01_excel_row = d01_line + 1
                d01_rows.append(
                    {
                        "订单编号": "",
                        "原始单号": d01_original_order_no(raw, rules),
                        "客户备注": norm(raw.get("运单号")),
                        "客服备注": norm(raw.get("物流方式")),
                        "自定义1": f"{seq:02d}",
                    }
                )
            elif dest == "D02":
                d02_line += 1
                d02_excel_row = d02_line + 1
                phone = d02_phone(raw, rules, anomalies, dest="D02")
                d02_rows.append(
                    {
                        "订单编号": dx_ono,
                        "客户网名": norm(raw.get("收货人姓名")),
                        "收货人": norm(raw.get("收货人姓名")),
                        "电话": phone,
                        "国家/地区": norm(raw.get("收货人国家")),
                        "地址": norm(raw.get("详细地址")),
                        "客服备注": d02_kefu_note(raw, rules, base_sku),
                        "客户备注": norm(raw.get("运单号")),
                        "编号": base_sku,
                        "条码": base_sku,
                        "数量": qty,
                        "合计": amt,
                        "订单来源": rules.d02_fixed["订单来源"],
                        "店铺": norm(raw.get("店铺账号")),
                        "物流方式": rules.d02_fixed["物流方式"],
                        "自定义1": f"{seq:02d}",
                    }
                )
            else:
                pending.append(
                    {
                        "订单号": norm(raw.get("订单号")),
                        "店铺账号": raw.get("店铺账号"),
                        "运单号": norm(raw.get("运单号")),
                        "dx_seq": seq,
                    }
                )

            mappings.append(
                {
                    "dx_seq": seq,
                    "raw_excel_row": raw.get("_excel_row"),
                    "运单号": norm(raw.get("运单号")),
                    "原始订单号": norm(raw.get("订单号")),
                    "dx_订单号": dx_ono,
                    "d02_订单编号": dx_ono if dest == "D02" else "",
                    "店铺账号": raw.get("店铺账号"),
                    "原始SKU": ("" if raw.get("SKU") is None else str(raw.get("SKU")).strip()),
                    "dx_SKU": dx_sku,
                    "base_sku": base_sku,
                    "sku_line_count": len(sku_lines),
                    "is_multi_sku": is_multi,
                    "原始物流方式": norm(raw.get("物流方式")),
                    "dx_物流方式": dx_log,
                    "is_china_direct": is_china_direct(raw, rules),
                    "dest": dest,
                    "店管家01行号": d01_excel_row if d01_excel_row is not None else "",
                    "店管家02行号": d02_excel_row if d02_excel_row is not None else "",
                    "dx_all_red": dx_row["_all_red"],
                    "拣货备注": norm(raw.get("拣货备注")),
                    "客服备注_原": norm(raw.get("客服备注")),
                    "订单备注_原": norm(raw.get("订单备注")),
                    "sku_from_note_column": note_hit[0] if note_hit else "",
                    "sku_from_pick_note": bool(note_hit),
                }
            )

    # 店管家02订单编号去重（原始单号重复不阻断；仅改 D02，不动店小秘）
    dedupe_d02_order_nos(d02_rows, anomalies, mappings)

    anom_dicts = [a.to_dict() for a in anomalies]
    for a in anom_dicts:
        if a.get("type") in ERROR_TYPES:
            a["severity"] = "error"
        a.setdefault("severity", "warning")
    errors = [a for a in anom_dicts if a["severity"] == "error"]
    warnings = [a for a in anom_dicts if a["severity"] == "warning"]
    infos = [a for a in anom_dicts if a["severity"] == "info"]
    needs_review_flags = [a for a in anom_dicts if a.get("type") in REVIEW_TYPES]

    if errors:
        export_status: ExportStatus = "blocked"
    elif needs_review_flags:
        export_status = "needs_review"
    else:
        export_status = "ok"

    blocked = export_status == "blocked"

    result = ConvertResult(
        anomalies=anom_dicts,
        mappings=mappings,
        pending_orders=pending,
        formal_export_blocked=blocked,
        export_status=export_status,
        errors=errors,
        warnings=warnings,
        infos=infos,
    )

    red_rgb = rules.raw.get("dianxiaomi_red_font_rgb", "FFFF0000")

    def _write_dx(target: Path) -> None:
        write_dianxiaomi(
            target,
            headers,
            dx_rows,
            template=template_dx,
            red_rgb=red_rgb,
        )

    def _write_d01(target: Path) -> None:
        write_dian_01(target, d01_rows, template_d01)

    def _write_d02(target: Path) -> None:
        write_dian_02(target, d02_rows, template_d02, rules.d02_red_columns)

    # 清理可能残留的正式文件名，避免待复核/阻断时误留可导入文件
    for name in ("dianxiaomi.xlsx", "dian_01.xls", "dian_02.xls",
                 "店小秘订单.xlsx", "01店管家更新订单信息.xls", "02店管家销售导入.xls"):
        p = out_dir / name
        if p.exists():
            p.unlink()

    if hard_stop:
        pass  # 不写业务表
    elif export_status == "ok":
        if "dx" in exports:
            result.dx_path = out_dir / "dianxiaomi.xlsx"
            _write_dx(result.dx_path)
        if "d01" in exports:
            result.d01_path = out_dir / "dian_01.xls"
            _write_d01(result.d01_path)
        if "d02" in exports:
            result.d02_path = out_dir / "dian_02.xls"
            _write_d02(result.d02_path)
    elif export_status == "needs_review":
        # 默认写入 review，文件名带 NEEDS_REVIEW；正式路径不得出现
        review = out_dir / "review"
        review.mkdir(exist_ok=True)
        if "dx" in exports:
            result.dx_path = review / "dianxiaomi_NEEDS_REVIEW.xlsx"
            _write_dx(result.dx_path)
        if "d01" in exports:
            result.d01_path = review / "dian_01_NEEDS_REVIEW.xls"
            _write_d01(result.d01_path)
        if "d02" in exports:
            result.d02_path = review / "dian_02_NEEDS_REVIEW.xls"
            _write_d02(result.d02_path)
    else:  # blocked
        # 店小秘预览仍可写到 out（非正式导入名也可接受作排查）；01/02 禁止正式名
        if "dx" in exports:
            result.dx_path = out_dir / "dianxiaomi.xlsx"
            _write_dx(result.dx_path)
        if allow_blocked_review_export:
            review = out_dir / "review"
            review.mkdir(exist_ok=True)
            if "d01" in exports:
                p = review / "dian_01_NEEDS_REVIEW.xls"
                _write_d01(p)
                result.d01_path = p
            if "d02" in exports:
                p = review / "dian_02_NEEDS_REVIEW.xls"
                _write_d02(p)
                result.d02_path = p

    result.stats = {
        "raw_orders": len(raw_rows),
        "dx_orders": len(dx_rows),
        "d01_orders": len(d01_rows),
        "d02_orders": len(d02_rows),
        "pending_orders": len(pending),
        "china_direct": sum(1 for m in mappings if m["is_china_direct"]),
        "multi_sku_orders": sum(1 for m in mappings if m.get("is_multi_sku")),
        "sku_from_pick_note": sum(1 for m in mappings if m.get("sku_from_pick_note")),
        "sku_from_note": sum(1 for m in mappings if m.get("sku_from_pick_note")),
        "d02_order_no_dedup": sum(1 for a in anom_dicts if a.get("type") == "D02_ORDER_NO_DEDUP"),
        "anomalies": len(anom_dicts),
        "errors": len(errors),
        "warnings": len(warnings),
        "infos": len(infos),
        "review_flags": len(needs_review_flags),
        "formal_export_blocked": blocked,
        "export_status": export_status,
        "anomaly_types": {},
    }
    for a in anom_dicts:
        t = a["type"]
        result.stats["anomaly_types"][t] = result.stats["anomaly_types"].get(t, 0) + 1

    if write_sidecars:
        _write_sidecars(out_dir, result)

    return result


def _write_sidecars(out_dir: Path, result: ConvertResult) -> None:
    (out_dir / "stats.json").write_text(
        json.dumps(result.stats, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (out_dir / "anomalies.json").write_text(
        json.dumps(
            {
                "export_status": result.export_status,
                "formal_export_blocked": result.formal_export_blocked,
                "errors": result.errors,
                "warnings": result.warnings,
                "infos": result.infos,
                "anomalies": result.anomalies,
                "pending_orders": result.pending_orders,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    if result.mappings:
        fields = list(result.mappings[0].keys())
        with open(out_dir / "row_mapping.csv", "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=fields)
            w.writeheader()
            w.writerows(result.mappings)
        (out_dir / "row_mapping.json").write_text(
            json.dumps(result.mappings, ensure_ascii=False, indent=2), encoding="utf-8"
        )
