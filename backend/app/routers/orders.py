from decimal import Decimal

from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy import func, or_, select
from sqlalchemy.orm import Session, selectinload

from app.database import get_db
from app.deps import get_current_user, require_roles
from app.models import Order, OrderLine, OrderPackage, PackageTracking, User
from app.schemas import OrderCreate
from app.serialize import order_detail, order_summary

router = APIRouter(prefix="/api/orders", tags=["orders"])

_ALLOWED_STATUS = {"imported", "pending_review", "exception", "converted"}
_ORDER_LOAD = (
    selectinload(Order.lines),
    selectinload(Order.packages).selectinload(OrderPackage.trackings),
)


def load_order(db: Session, order_id: int) -> Order | None:
    stmt = select(Order).where(Order.id == order_id).options(*_ORDER_LOAD)
    return db.scalars(stmt).first()


def _filters(project: str | None, q: str | None) -> list:
    filters = []
    if project:
        if project not in {"project1", "project2"}:
            raise HTTPException(status_code=400, detail="project 只能是 project1 或 project2")
        filters.append(Order.project == project)
    if q and q.strip():
        like = f"%{q.strip()}%"
        filters.append(
            or_(
                Order.external_order_no.ilike(like),
                Order.shop_account.ilike(like),
                Order.platform.ilike(like),
            )
        )
    return filters


@router.get("")
def list_orders(
    project: str | None = None,
    q: str | None = None,
    page: int = Query(1, ge=1),
    page_size: int = Query(20, ge=1, le=100),
    _user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    filters = _filters(project, q)
    count_stmt = select(func.count(Order.id))
    stmt = select(Order).options(*_ORDER_LOAD).order_by(Order.id.desc())
    for item in filters:
        count_stmt = count_stmt.where(item)
        stmt = stmt.where(item)
    total = int(db.scalar(count_stmt) or 0)
    rows = db.scalars(stmt.offset((page - 1) * page_size).limit(page_size)).all()
    return {
        "items": [order_summary(row) for row in rows],
        "total": total,
        "page": page,
        "page_size": page_size,
    }


@router.get("/{order_id}")
def get_order(
    order_id: int,
    _user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    order = load_order(db, order_id)
    if order is None:
        raise HTTPException(status_code=404, detail="订单不存在")
    return order_detail(order)


@router.post("", status_code=201)
def create_order(
    body: OrderCreate,
    _user: User = Depends(require_roles("ops")),
    db: Session = Depends(get_db),
) -> dict:
    if body.status not in _ALLOWED_STATUS:
        raise HTTPException(status_code=400, detail="订单状态不在允许范围内")
    package_nos = [package.package_no for package in body.packages]
    if len(package_nos) != len(set(package_nos)):
        raise HTTPException(status_code=400, detail="同一订单里包裹号不能重复")
    for package in body.packages:
        tracking_nos = [item.tracking_no for item in package.trackings]
        if len(tracking_nos) != len(set(tracking_nos)):
            raise HTTPException(status_code=400, detail="同一包裹里运单号不能重复")

    exists = db.scalar(
        select(Order.id).where(
            Order.project == body.project,
            Order.external_order_no == body.external_order_no,
        )
    )
    if exists is not None:
        raise HTTPException(status_code=409, detail="该项目下订单号已存在")

    amount = sum((line.line_amount for line in body.lines), Decimal("0.00"))
    order = Order(
        project=body.project,
        external_order_no=body.external_order_no.strip(),
        shop_account=body.shop_account,
        platform=body.platform,
        status=body.status,
        buyer_name=body.buyer_name,
        country=body.country,
        currency=body.currency,
        order_amount=amount,
        dest_hint=body.dest_hint,
        row_mark=body.row_mark,
        already_in_guanjia=body.already_in_guanjia,
        notes=body.notes,
        raw_payload={"source": "api"},
    )
    db.add(order)
    db.flush()
    for index, line in enumerate(body.lines, start=1):
        db.add(
            OrderLine(
                order_id=order.id,
                line_no=index,
                sku=line.sku.strip(),
                sku_name=line.sku_name,
                quantity=line.quantity,
                unit_price=line.unit_price,
                line_amount=line.line_amount,
                currency=line.currency or body.currency,
            )
        )
    for package in body.packages:
        pkg = OrderPackage(
            order_id=order.id,
            package_no=package.package_no.strip(),
            warehouse_code=package.warehouse_code,
            logistics_method=package.logistics_method,
            weight_kg=package.weight_kg,
            status=package.status or "packed",
        )
        db.add(pkg)
        db.flush()
        for tracking in package.trackings:
            db.add(
                PackageTracking(
                    order_id=order.id,
                    package_id=pkg.id,
                    tracking_no=tracking.tracking_no.strip(),
                    carrier=tracking.carrier,
                )
            )
    db.commit()
    stored = load_order(db, order.id)
    if stored is None:
        raise HTTPException(status_code=500, detail="订单写入后读取失败")
    return order_detail(stored)
