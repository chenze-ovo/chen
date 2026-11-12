from __future__ import annotations

from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP

from app.models import (
    ExceptionReview,
    ImportBatch,
    Order,
    ReviewTask,
    SplitAmountReview,
)

_MONEY = Decimal("0.01")
_QTY = Decimal("0.01")
_WEIGHT = Decimal("0.001")


def _q(value: Decimal | None, places: Decimal) -> str | None:
    if value is None:
        return None
    return format(Decimal(value).quantize(places, rounding=ROUND_HALF_UP), "f")


def money(value: Decimal | None) -> str | None:
    return _q(value, _MONEY)


def qty(value: Decimal | None) -> str | None:
    return _q(value, _QTY)


def weight(value: Decimal | None) -> str | None:
    return _q(value, _WEIGHT)


def iso(value: datetime | None) -> str | None:
    if value is None:
        return None
    return value.isoformat(timespec="seconds")


def order_summary(order: Order) -> dict:
    line_sum = sum((line.line_amount for line in order.lines), Decimal("0"))
    tracking_count = sum(len(package.trackings) for package in order.packages)
    return {
        "id": order.id,
        "project": order.project,
        "external_order_no": order.external_order_no,
        "shop_account": order.shop_account,
        "platform": order.platform,
        "status": order.status,
        "buyer_name": order.buyer_name,
        "currency": order.currency,
        "order_amount": money(order.order_amount),
        "line_amount_sum": money(line_sum),
        "line_count": len(order.lines),
        "package_count": len(order.packages),
        "tracking_count": tracking_count,
        "row_mark": order.row_mark,
        "already_in_guanjia": order.already_in_guanjia,
        "created_at": iso(order.created_at),
    }


def order_detail(order: Order) -> dict:
    body = order_summary(order)
    body.update(
        {
            "country": order.country,
            "dest_hint": order.dest_hint,
            "notes": order.notes,
            "batch_id": order.batch_id,
            "lines": [
                {
                    "id": line.id,
                    "line_no": line.line_no,
                    "sku": line.sku,
                    "sku_name": line.sku_name,
                    "quantity": qty(line.quantity),
                    "unit_price": money(line.unit_price),
                    "line_amount": money(line.line_amount),
                    "currency": line.currency,
                }
                for line in order.lines
            ],
            "packages": [
                {
                    "id": package.id,
                    "package_no": package.package_no,
                    "warehouse_code": package.warehouse_code,
                    "logistics_method": package.logistics_method,
                    "weight_kg": weight(package.weight_kg),
                    "status": package.status,
                    "trackings": [
                        {
                            "id": tracking.id,
                            "tracking_no": tracking.tracking_no,
                            "carrier": tracking.carrier,
                        }
                        for tracking in package.trackings
                    ],
                }
                for package in order.packages
            ],
        }
    )
    return body



_INFO_ISSUE_TYPES = frozenset(
    {
        "SKU_FROM_NOTE",
        "SKU_STRIP_PREFIX",
        "SKU_STRIP_WXM_PREFIX",
        "D02_ORDER_NO_DEDUP",
    }
)


def _is_info_issue(item: dict) -> bool:
    typ = str(item.get("type") or item.get("code") or "")
    severity = str(item.get("severity") or "").lower()
    return typ in _INFO_ISSUE_TYPES or severity == "info"


def batch_exception_count(batch: ImportBatch) -> int:
    """给列表页一个轻量异常数，不改也不回传整份 summary 的结构。"""
    summary = batch.engine_summary or {}
    if batch.project == "project1":
        listed = 0
        seen: set[tuple] = set()
        counted = 0
        for item in summary.get("files") or []:
            counted += int(item.get("error_count") or 0) + int(item.get("warning_count") or 0)
            buckets = list(item.get("errors") or []) + list(item.get("warnings") or [])
            file_name = str(item.get("file") or "")
            for raw in buckets:
                if not isinstance(raw, dict):
                    key = (file_name, "", str(raw))
                else:
                    if _is_info_issue(raw):
                        continue
                    key = (
                        file_name,
                        str(raw.get("订单号") or raw.get("order_no") or ""),
                        str(raw.get("type") or raw.get("code") or ""),
                        str(raw.get("suffix") or raw.get("SKU") or raw.get("店铺账号") or raw.get("message") or ""),
                    )
                if key in seen:
                    continue
                seen.add(key)
                listed += 1
        if listed:
            return max(listed, counted) if counted > listed else listed
        return counted
    stats = summary.get("stats") or {}
    if stats.get("human_review") not in (None, ""):
        return int(stats.get("human_review") or 0)
    if stats.get("review") not in (None, ""):
        return int(stats.get("review") or 0)
    anomalies = summary.get("anomalies") or []
    return len(anomalies)

def batch_detail(batch: ImportBatch) -> dict:
    creator = batch.creator
    return {
        "id": batch.id,
        "project": batch.project,
        "status": batch.status,
        "note": batch.note,
        "engine_status": batch.engine_status,
        "engine_summary": batch.engine_summary,
        "exception_count": batch_exception_count(batch),
        "created_by": creator.username if creator is not None else None,
        "created_at": iso(batch.created_at),
        "files": [
            {
                "id": item.id,
                "original_name": item.original_name,
                "size_bytes": item.size_bytes,
                "sha256": item.sha256,
                "kind": item.kind,
            }
            for item in batch.files
        ],
    }


def split_review_detail(review: SplitAmountReview) -> dict:
    order = review.order
    return {
        "id": review.id,
        "order_id": review.order_id,
        "external_order_no": order.external_order_no if order is not None else None,
        "project": order.project if order is not None else None,
        "status": review.status,
        "reason": review.reason,
        "review_note": review.review_note,
        "created_at": iso(review.created_at),
        "decided_at": iso(review.decided_at),
        "lines": [
            {
                "id": line.id,
                "order_line_id": line.order_line_id,
                "sku": line.sku,
                "original_amount": money(line.original_amount),
                "proposed_amount": money(line.proposed_amount),
                "note": line.note,
            }
            for line in review.lines
        ],
    }


def review_task_detail(task: ReviewTask) -> dict:
    return {
        "id": task.id,
        "task_type": task.task_type,
        "status": task.status,
        "order_id": task.order_id,
        "batch_id": task.batch_id,
        "batch_file_id": task.batch_file_id,
        "message": task.message,
        "created_at": iso(task.created_at),
        "resolved_at": iso(task.resolved_at),
    }


def exception_detail(review: ExceptionReview) -> dict:
    order = review.order
    return {
        "id": review.id,
        "order_id": review.order_id,
        "external_order_no": order.external_order_no if order is not None else None,
        "batch_id": review.batch_id,
        "project": review.project,
        "code": review.code,
        "severity": review.severity,
        "message": review.message,
        "status": review.status,
        "created_at": iso(review.created_at),
        "resolved_at": iso(review.resolved_at),
    }
