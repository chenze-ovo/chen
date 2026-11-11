from __future__ import annotations

from decimal import Decimal
from typing import Literal

from pydantic import BaseModel, Field


class LoginIn(BaseModel):
    username: str = Field(min_length=1, max_length=64)
    password: str = Field(min_length=1, max_length=128)


class TrackingIn(BaseModel):
    tracking_no: str = Field(min_length=1, max_length=128)
    carrier: str | None = Field(default=None, max_length=64)


class PackageIn(BaseModel):
    package_no: str = Field(min_length=1, max_length=64)
    warehouse_code: str | None = Field(default=None, max_length=64)
    logistics_method: str | None = Field(default=None, max_length=128)
    weight_kg: Decimal | None = Field(default=None, ge=0)
    status: str = Field(default="packed", max_length=32)
    trackings: list[TrackingIn] = Field(default_factory=list)


class LineIn(BaseModel):
    sku: str = Field(min_length=1, max_length=128)
    sku_name: str | None = Field(default=None, max_length=255)
    quantity: Decimal = Field(gt=0)
    unit_price: Decimal = Field(ge=0)
    line_amount: Decimal = Field(ge=0)
    currency: str | None = Field(default=None, max_length=8)


class OrderCreate(BaseModel):
    project: Literal["project1", "project2"]
    external_order_no: str = Field(min_length=1, max_length=128)
    shop_account: str | None = Field(default=None, max_length=128)
    platform: str | None = Field(default=None, max_length=64)
    status: str = "imported"
    buyer_name: str | None = Field(default=None, max_length=128)
    country: str | None = Field(default=None, max_length=64)
    currency: str = Field(default="USD", min_length=1, max_length=8)
    dest_hint: str | None = Field(default=None, max_length=64)
    row_mark: Literal["green", "white"] | None = None
    already_in_guanjia: bool = False
    notes: str | None = Field(default=None, max_length=2000)
    lines: list[LineIn] = Field(min_length=1)
    packages: list[PackageIn] = Field(default_factory=list)


class SplitLineIn(BaseModel):
    order_line_id: int | None = None
    sku: str = Field(min_length=1, max_length=128)
    original_amount: Decimal = Field(ge=0)
    proposed_amount: Decimal = Field(ge=0)
    note: str | None = Field(default=None, max_length=500)


class SplitReviewCreate(BaseModel):
    order_id: int
    reason: str | None = Field(default=None, max_length=2000)
    lines: list[SplitLineIn] = Field(min_length=1)


class SkuMapIn(BaseModel):
    project: Literal["project1", "project2"]
    source_sku: str = Field(min_length=1, max_length=128)
    target_sku: str = Field(min_length=1, max_length=128)
    note: str | None = None


class OpsGroupIn(BaseModel):
    name: str = Field(min_length=1, max_length=128)
    project: Literal["project1", "project2"] | None = "project2"
    note: str | None = None


class ShopIn(BaseModel):
    project: Literal["project1", "project2"]
    shop_account: str = Field(min_length=1, max_length=128)
    platform: str | None = None
    ops_group_id: int | None = None
    already_in_guanjia: bool = False
    destination: Literal["dian_01", "dian_02", "review", "human_review"] | None = None
    routing_note: str | None = None


class ExportProfileIn(BaseModel):
    project: Literal["project2"] = "project2"
    kind: Literal["dian_01", "dian_02"]
    column_map: list[dict] = Field(default_factory=list)
    note: str | None = None


class ReviewTaskCreate(BaseModel):
    task_type: Literal["china_direct_amount_unmatched", "split_sku_amount_manual"]
    order_id: int | None = None
    batch_id: int | None = None
    batch_file_id: int | None = None
    message: str | None = Field(default=None, max_length=2000)


class SplitFillLine(BaseModel):
    id: int
    proposed_amount: Decimal = Field(ge=0)


class DecisionIn(BaseModel):
    status: Literal["approved", "rejected"]
    review_note: str | None = Field(default=None, max_length=2000)
    lines: list[SplitFillLine] | None = None
