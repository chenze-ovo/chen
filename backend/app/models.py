"""PostgreSQL 导向的表结构。SQLite 仅作本地回退，类型保持兼容。

一单多行：order_lines 存 SKU 与行金额。
一单多包裹：order_packages。
一包裹多运单：package_trackings，同时冗余 order_id 便于按单查询。
拆单金额只建审核表。员工核对店小秘网页后在 ERP 手工录入，不调用 API，也不爬页面。
row_mark 预留绿标/白标。项目2以后 Shirley（速卖通）和 eBay 转店管家01 时两种标记都要覆盖。
order_lines 支持一单多 SKU。10 月 4 日样本只作参考，不从中写入 SKU、金额、仓库或物流映射。
Phase 0 不解析颜色，也不生成店管家文件。

batch_files.kind 与 review_tasks.task_type 只预留种类，不是 SKU / 仓库 / 列映射。
项目2中国直发金额禁止写成 0；当天店小秘订单 Excel 用 kind=dianxiaomi_same_day 归档，
对不上再开 china_direct_amount_unmatched。本阶段不执行对照。
"""
from __future__ import annotations

from datetime import datetime, timezone
from decimal import Decimal

from sqlalchemy import (
    Boolean,
    CheckConstraint,
    DateTime,
    ForeignKey,
    Integer,
    Numeric,
    String,
    Text,
    UniqueConstraint,
)
from sqlalchemy.dialects.postgresql import JSONB
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship
from sqlalchemy.types import JSON


def utcnow() -> datetime:
    return datetime.now(timezone.utc).replace(tzinfo=None)


def json_type():
    return JSON().with_variant(JSONB(), "postgresql")


# 文件角色。样本包不是 12 份都算原始订单：4 原始、2 海外仓、6 运营分组。
# dianxiaomi_same_day 只表示「当天店小秘订单 Excel」，供以后对照中国直发金额，现在不解析。
BATCH_FILE_KINDS = (
    "unspecified",
    "raw_order",
    "oversea_working",
    "ops_group",
    "dianxiaomi_same_day",
)

# 人工任务种类。不包含外部抓取。
REVIEW_TASK_TYPES = (
    "china_direct_amount_unmatched",
    "split_sku_amount_manual",
)


def _sql_in(column: str, values: tuple[str, ...]) -> str:
    listed = ", ".join(f"'{item}'" for item in values)
    return f"{column} in ({listed})"


class Base(DeclarativeBase):
    pass


class User(Base):
    __tablename__ = "users"
    __table_args__ = (
        CheckConstraint("role in ('admin', 'reviewer', 'ops')", name="ck_users_role"),
    )

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    username: Mapped[str] = mapped_column(String(64), unique=True, index=True)
    password_hash: Mapped[str] = mapped_column(String(255))
    role: Mapped[str] = mapped_column(String(32))
    display_name: Mapped[str] = mapped_column(String(64))
    is_active: Mapped[bool] = mapped_column(Boolean, default=True, nullable=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow, nullable=False)


class ImportBatch(Base):
    __tablename__ = "import_batches"
    __table_args__ = (
        CheckConstraint("project in ('project1', 'project2')", name="ck_batches_project"),
    )

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    project: Mapped[str] = mapped_column(String(16), index=True)
    status: Mapped[str] = mapped_column(String(32), default="created", nullable=False)
    note: Mapped[str | None] = mapped_column(Text, nullable=True)
    created_by_id: Mapped[int] = mapped_column(ForeignKey("users.id"))
    engine_status: Mapped[str | None] = mapped_column(String(32), nullable=True)
    engine_summary: Mapped[dict | None] = mapped_column(json_type(), nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow, nullable=False)

    creator: Mapped[User] = relationship()
    files: Mapped[list[BatchFile]] = relationship(
        back_populates="batch",
        cascade="all, delete-orphan",
        order_by="BatchFile.id",
    )


class BatchFile(Base):
    """上传文件。kind 只标记角色，不触发转换。"""

    __tablename__ = "batch_files"
    __table_args__ = (
        CheckConstraint(_sql_in("kind", BATCH_FILE_KINDS), name="ck_batch_files_kind"),
    )

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    batch_id: Mapped[int] = mapped_column(ForeignKey("import_batches.id"))
    original_name: Mapped[str] = mapped_column(String(255))
    stored_path: Mapped[str] = mapped_column(String(512))
    size_bytes: Mapped[int] = mapped_column(Integer)
    sha256: Mapped[str] = mapped_column(String(64))
    # unspecified | raw_order | oversea_working | ops_group | dianxiaomi_same_day
    kind: Mapped[str] = mapped_column(String(32), default="unspecified", nullable=False)

    batch: Mapped[ImportBatch] = relationship(back_populates="files")


class Order(Base):
    """销售订单头。行、包裹、运单都在子表。"""

    __tablename__ = "orders"
    __table_args__ = (
        UniqueConstraint("project", "external_order_no", name="uq_orders_project_no"),
        CheckConstraint("project in ('project1', 'project2')", name="ck_orders_project"),
        CheckConstraint(
            "row_mark is null or row_mark in ('green', 'white')",
            name="ck_orders_row_mark",
        ),
    )

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    batch_id: Mapped[int | None] = mapped_column(ForeignKey("import_batches.id"), nullable=True)
    project: Mapped[str] = mapped_column(String(16), index=True)
    external_order_no: Mapped[str] = mapped_column(String(128), index=True)
    shop_account: Mapped[str | None] = mapped_column(String(128), nullable=True)
    platform: Mapped[str | None] = mapped_column(String(64), nullable=True)
    status: Mapped[str] = mapped_column(String(32), default="imported", nullable=False)
    buyer_name: Mapped[str | None] = mapped_column(String(128), nullable=True)
    country: Mapped[str | None] = mapped_column(String(64), nullable=True)
    currency: Mapped[str] = mapped_column(String(8), default="USD", nullable=False)
    order_amount: Mapped[Decimal] = mapped_column(Numeric(14, 2), nullable=False)
    dest_hint: Mapped[str | None] = mapped_column(String(64), nullable=True)
    # green / white / null。Project2 绿标与白标后续都要处理，P0 只存值。
    row_mark: Mapped[str | None] = mapped_column(String(16), nullable=True)
    already_in_guanjia: Mapped[bool] = mapped_column(Boolean, default=False, nullable=False)
    notes: Mapped[str | None] = mapped_column(Text, nullable=True)
    raw_payload: Mapped[dict | None] = mapped_column(json_type(), nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow, nullable=False)

    lines: Mapped[list[OrderLine]] = relationship(
        back_populates="order",
        cascade="all, delete-orphan",
        order_by="OrderLine.line_no",
    )
    packages: Mapped[list[OrderPackage]] = relationship(
        back_populates="order",
        cascade="all, delete-orphan",
        order_by="OrderPackage.id",
    )
    split_reviews: Mapped[list[SplitAmountReview]] = relationship(back_populates="order")


class OrderLine(Base):
    __tablename__ = "order_lines"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    order_id: Mapped[int] = mapped_column(ForeignKey("orders.id"), index=True)
    line_no: Mapped[int] = mapped_column(Integer)
    sku: Mapped[str] = mapped_column(String(128))
    sku_name: Mapped[str | None] = mapped_column(String(255), nullable=True)
    quantity: Mapped[Decimal] = mapped_column(Numeric(12, 2))
    unit_price: Mapped[Decimal] = mapped_column(Numeric(14, 2))
    line_amount: Mapped[Decimal] = mapped_column(Numeric(14, 2))
    currency: Mapped[str] = mapped_column(String(8))

    order: Mapped[Order] = relationship(back_populates="lines")


class OrderPackage(Base):
    __tablename__ = "order_packages"
    __table_args__ = (
        UniqueConstraint("order_id", "package_no", name="uq_package_no_per_order"),
    )

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    order_id: Mapped[int] = mapped_column(ForeignKey("orders.id"), index=True)
    package_no: Mapped[str] = mapped_column(String(64))
    warehouse_code: Mapped[str | None] = mapped_column(String(64), nullable=True)
    logistics_method: Mapped[str | None] = mapped_column(String(128), nullable=True)
    weight_kg: Mapped[Decimal | None] = mapped_column(Numeric(12, 3), nullable=True)
    status: Mapped[str] = mapped_column(String(32), default="packed", nullable=False)

    order: Mapped[Order] = relationship(back_populates="packages")
    trackings: Mapped[list[PackageTracking]] = relationship(
        back_populates="package",
        cascade="all, delete-orphan",
        order_by="PackageTracking.id",
    )


class PackageTracking(Base):
    __tablename__ = "package_trackings"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    order_id: Mapped[int] = mapped_column(ForeignKey("orders.id"), index=True)
    package_id: Mapped[int] = mapped_column(ForeignKey("order_packages.id"), index=True)
    tracking_no: Mapped[str] = mapped_column(String(128))
    carrier: Mapped[str | None] = mapped_column(String(64), nullable=True)

    package: Mapped[OrderPackage] = relationship(back_populates="trackings")


class SplitAmountReview(Base):
    """拆单金额审核预留。不调用店小秘接口。"""

    __tablename__ = "split_amount_reviews"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    order_id: Mapped[int] = mapped_column(ForeignKey("orders.id"), index=True)
    status: Mapped[str] = mapped_column(String(32), default="pending", nullable=False)
    reason: Mapped[str | None] = mapped_column(Text, nullable=True)
    reviewer_id: Mapped[int | None] = mapped_column(ForeignKey("users.id"), nullable=True)
    review_note: Mapped[str | None] = mapped_column(Text, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow, nullable=False)
    decided_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)

    order: Mapped[Order] = relationship(back_populates="split_reviews")
    lines: Mapped[list[SplitAmountReviewLine]] = relationship(
        back_populates="review",
        cascade="all, delete-orphan",
        order_by="SplitAmountReviewLine.id",
    )


class SplitAmountReviewLine(Base):
    __tablename__ = "split_amount_review_lines"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    review_id: Mapped[int] = mapped_column(ForeignKey("split_amount_reviews.id"), index=True)
    order_line_id: Mapped[int | None] = mapped_column(ForeignKey("order_lines.id"), nullable=True)
    sku: Mapped[str] = mapped_column(String(128))
    original_amount: Mapped[Decimal] = mapped_column(Numeric(14, 2))
    proposed_amount: Mapped[Decimal] = mapped_column(Numeric(14, 2))
    note: Mapped[str | None] = mapped_column(Text, nullable=True)

    review: Mapped[SplitAmountReview] = relationship(back_populates="lines")


class ExceptionReview(Base):
    __tablename__ = "exception_reviews"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    order_id: Mapped[int | None] = mapped_column(ForeignKey("orders.id"), nullable=True)
    batch_id: Mapped[int | None] = mapped_column(ForeignKey("import_batches.id"), nullable=True)
    project: Mapped[str] = mapped_column(String(16))
    code: Mapped[str] = mapped_column(String(64))
    severity: Mapped[str] = mapped_column(String(16), default="warning", nullable=False)
    message: Mapped[str] = mapped_column(Text)
    status: Mapped[str] = mapped_column(String(32), default="open", nullable=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow, nullable=False)
    resolved_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)

    order: Mapped[Order | None] = relationship()


class SkuSuffix(Base):
    """项目2可编辑的 SKU 后缀。只剥离名单里的末段，不猜测仓库或金额。"""

    __tablename__ = "sku_suffixes"
    __table_args__ = (UniqueConstraint("project", "suffix", name="uq_sku_suffix"),)

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    project: Mapped[str] = mapped_column(String(16), index=True)
    suffix: Mapped[str] = mapped_column(String(32))
    enabled: Mapped[bool] = mapped_column(Boolean, default=True, nullable=False)


class SkuMapping(Base):
    """ERP 侧补充映射。Project2 禁止预置。项目1规则留在 order_convert。"""

    __tablename__ = "sku_mappings"
    __table_args__ = (
        UniqueConstraint("project", "source_sku", name="uq_sku_map_source"),
    )

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    project: Mapped[str] = mapped_column(String(16), index=True)
    source_sku: Mapped[str] = mapped_column(String(128))
    target_sku: Mapped[str] = mapped_column(String(128))
    note: Mapped[str | None] = mapped_column(Text, nullable=True)
    enabled: Mapped[bool] = mapped_column(Boolean, default=True, nullable=False)


class OpsGroup(Base):
    __tablename__ = "ops_groups"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    name: Mapped[str] = mapped_column(String(128), unique=True)
    project: Mapped[str | None] = mapped_column(String(16), nullable=True)
    note: Mapped[str | None] = mapped_column(Text, nullable=True)

    shops: Mapped[list[Shop]] = relationship(back_populates="ops_group")


class Shop(Base):
    __tablename__ = "shops"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    project: Mapped[str] = mapped_column(String(16), index=True)
    shop_account: Mapped[str] = mapped_column(String(128))
    platform: Mapped[str | None] = mapped_column(String(64), nullable=True)
    ops_group_id: Mapped[int | None] = mapped_column(ForeignKey("ops_groups.id"), nullable=True)
    already_in_guanjia: Mapped[bool] = mapped_column(Boolean, default=False, nullable=False)
    # dian_01 | dian_02 | review | human_review | null（null 表示没有单独指定）
    destination: Mapped[str | None] = mapped_column(String(32), nullable=True)
    routing_note: Mapped[str | None] = mapped_column(Text, nullable=True)

    ops_group: Mapped[OpsGroup | None] = relationship(back_populates="shops")


class ExportProfile(Base):
    """项目2店管家导出的列对照。空表表示还没有模板规则，不能猜列。"""

    __tablename__ = "export_profiles"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    project: Mapped[str] = mapped_column(String(16), index=True)
    kind: Mapped[str] = mapped_column(String(32))
    column_map: Mapped[list | None] = mapped_column(json_type(), nullable=True)
    template_path: Mapped[str | None] = mapped_column(String(512), nullable=True)
    note: Mapped[str | None] = mapped_column(Text, nullable=True)


class ExportJob(Base):
    __tablename__ = "export_jobs"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    project: Mapped[str] = mapped_column(String(16))
    batch_id: Mapped[int | None] = mapped_column(ForeignKey("import_batches.id"), nullable=True)
    kind: Mapped[str] = mapped_column(String(64))
    status: Mapped[str] = mapped_column(String(32), default="reserved", nullable=False)
    note: Mapped[str | None] = mapped_column(Text, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow, nullable=False)


class ReviewTask(Base):
    """预留的人工任务。

    china_direct_amount_unmatched：
        项目2中国直发金额不能改成 0。以后用当天店小秘订单 Excel 对照，
        对不上就记这条任务，等人处理。Phase 0 不对照。
    split_sku_amount_manual：
        员工看过店小秘网页后，在 ERP 里手工填拆分金额。
        不调用店小秘 API，也不爬页面。
    """

    __tablename__ = "review_tasks"
    __table_args__ = (
        CheckConstraint(_sql_in("task_type", REVIEW_TASK_TYPES), name="ck_review_tasks_type"),
    )

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    task_type: Mapped[str] = mapped_column(String(64), index=True)
    status: Mapped[str] = mapped_column(String(32), default="pending", nullable=False)
    order_id: Mapped[int | None] = mapped_column(ForeignKey("orders.id"), nullable=True)
    batch_id: Mapped[int | None] = mapped_column(ForeignKey("import_batches.id"), nullable=True)
    batch_file_id: Mapped[int | None] = mapped_column(ForeignKey("batch_files.id"), nullable=True)
    message: Mapped[str | None] = mapped_column(Text, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow, nullable=False)
    resolved_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)


class EngineRun(Base):
    __tablename__ = "engine_runs"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    project: Mapped[str] = mapped_column(String(16), index=True)
    batch_id: Mapped[int] = mapped_column(ForeignKey("import_batches.id"))
    status: Mapped[str] = mapped_column(String(32))
    summary: Mapped[dict | None] = mapped_column(json_type(), nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow, nullable=False)
