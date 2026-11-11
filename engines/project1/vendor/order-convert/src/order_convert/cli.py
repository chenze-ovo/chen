# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import json
from pathlib import Path

from .config import load_rules
from .pipeline import convert_order_file


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="原始 Order Excel → 店小秘 / 店管家01 / 店管家02")
    p.add_argument("-i", "--input", required=True, type=Path, help="原始 Order xlsx")
    p.add_argument("-o", "--out-dir", required=True, type=Path, help="输出目录")
    p.add_argument(
        "--only",
        choices=["dx", "d01", "d02", "all"],
        default="all",
        help="导出哪些表（默认全部）",
    )
    p.add_argument("--config", type=Path, default=None, help="rules.yaml 路径")
    p.add_argument("--template-d01", type=Path, default=None, help="覆盖内置01空白模板")
    p.add_argument("--template-d02", type=Path, default=None, help="覆盖内置02空白模板")
    p.add_argument(
        "--allow-review-export",
        action="store_true",
        help="严重错误(blocked)时仍写入 out/review/*_NEEDS_REVIEW.xls；"
        "needs_review 状态默认就会写 review，无需此开关",
    )
    args = p.parse_args(argv)

    exports = ("dx", "d01", "d02") if args.only == "all" else (args.only,)
    rules = load_rules(args.config) if args.config else load_rules()
    result = convert_order_file(
        args.input,
        args.out_dir,
        rules=rules,
        template_d01=args.template_d01,
        template_d02=args.template_d02,
        exports=exports,
        allow_blocked_review_export=args.allow_review_export,
    )
    print(json.dumps(result.stats, ensure_ascii=False, indent=2))
    status = result.export_status
    if status == "blocked":
        print(
            f"[BLOCKED] export_status=blocked，存在 {len(result.errors)} 条严重错误，"
            f"未写出正式店管家01/02。详见 {args.out_dir / 'anomalies.json'}"
        )
        return 1
    if status == "needs_review":
        print(
            f"[NEEDS_REVIEW] export_status=needs_review，"
            f"存在未知SKU/多SKU等未确认问题；已写入 out/review/*_NEEDS_REVIEW.*，"
            f"正式可导入文件名未写出。详见 {args.out_dir / 'anomalies.json'}"
        )
        return 2
    if result.warnings:
        print(f"[warnings] {len(result.warnings)} 条（未触发待复核）→ {args.out_dir / 'anomalies.json'}")
        return 0
    return 0
