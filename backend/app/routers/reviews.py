"""审核接口。拆单金额只落库，不请求店小秘。"""
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy import select
from sqlalchemy.orm import Session, selectinload

from app.database import get_db
from app.deps import get_current_user, require_roles
from decimal import Decimal

from app.models import (
    REVIEW_TASK_TYPES,
    ExceptionReview,
    OrderLine,
    Order,
    ReviewTask,
    SplitAmountReview,
    SplitAmountReviewLine,
    User,
    utcnow,
)
from app.schemas import DecisionIn, ReviewTaskCreate, SplitReviewCreate
from app.serialize import exception_detail, review_task_detail, split_review_detail

router = APIRouter(prefix="/api/reviews", tags=["reviews"])


def _load_split(db: Session, review_id: int) -> SplitAmountReview | None:
    stmt = (
        select(SplitAmountReview)
        .where(SplitAmountReview.id == review_id)
        .options(selectinload(SplitAmountReview.lines), selectinload(SplitAmountReview.order))
    )
    return db.scalars(stmt).first()


@router.get("/split-amounts")
def list_split_reviews(
    _user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    stmt = (
        select(SplitAmountReview)
        .options(selectinload(SplitAmountReview.lines), selectinload(SplitAmountReview.order))
        .order_by(SplitAmountReview.id.desc())
    )
    rows = db.scalars(stmt).all()
    return {"items": [split_review_detail(row) for row in rows]}


@router.post("/split-amounts", status_code=201)
def create_split_review(
    body: SplitReviewCreate,
    _user: User = Depends(require_roles("ops")),
    db: Session = Depends(get_db),
) -> dict:
    order = db.get(Order, body.order_id)
    if order is None:
        raise HTTPException(status_code=404, detail="订单不存在")
    review = SplitAmountReview(
        order_id=order.id,
        status="pending",
        reason=body.reason,
        created_at=utcnow(),
    )
    db.add(review)
    db.flush()
    for line in body.lines:
        db.add(
            SplitAmountReviewLine(
                review_id=review.id,
                order_line_id=line.order_line_id,
                sku=line.sku,
                original_amount=line.original_amount,
                proposed_amount=line.proposed_amount,
                note=line.note,
            )
        )
    db.commit()
    stored = _load_split(db, review.id)
    if stored is None:
        raise HTTPException(status_code=500, detail="审核单写入后读取失败")
    return split_review_detail(stored)


@router.post("/split-amounts/{review_id}/decision")
def decide_split_review(
    review_id: int,
    body: DecisionIn,
    user: User = Depends(require_roles("reviewer")),
    db: Session = Depends(get_db),
) -> dict:
    # 只更新审核状态。不回写订单金额，也不调用外部接口。
    review = _load_split(db, review_id)
    if review is None:
        raise HTTPException(status_code=404, detail="审核单不存在")
    if review.status != "pending":
        raise HTTPException(status_code=409, detail="该审核已处理")
    if body.lines:
        incoming = {item.id: item.proposed_amount for item in body.lines}
        for line in review.lines:
            if line.id in incoming:
                line.proposed_amount = incoming[line.id]
        db.flush()
    if body.status == "approved":
        total = sum((line.proposed_amount for line in review.lines), Decimal("0"))
        order_total = review.order.order_amount if review.order is not None else Decimal("0")
        if abs(total - order_total) > Decimal("0.01"):
            raise HTTPException(
                status_code=400,
                detail=f"行金额合计 {total} 与订单金额 {order_total} 不一致，请按店小秘页面修改后再提交",
            )
        for line in review.lines:
            if line.order_line_id is None:
                continue
            order_line = db.get(OrderLine, line.order_line_id)
            if order_line is not None:
                order_line.line_amount = line.proposed_amount
    review.status = body.status
    review.review_note = body.review_note
    review.reviewer_id = user.id
    review.decided_at = utcnow()
    db.commit()
    stored = _load_split(db, review.id)
    if stored is None:
        raise HTTPException(status_code=500, detail="审核单更新后读取失败")
    return split_review_detail(stored)


@router.get("/tasks")
def list_review_tasks(
    _user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    rows = db.scalars(select(ReviewTask).order_by(ReviewTask.id.desc())).all()
    return {"items": [review_task_detail(row) for row in rows]}


@router.post("/tasks", status_code=201)
def create_review_task(
    body: ReviewTaskCreate,
    _user: User = Depends(require_roles("ops", "reviewer")),
    db: Session = Depends(get_db),
) -> dict:
    # 只建任务。不读 Excel，不把金额改成 0，不请求店小秘。
    if body.task_type not in REVIEW_TASK_TYPES:
        raise HTTPException(status_code=400, detail="不支持的审核任务类型")
    task = ReviewTask(
        task_type=body.task_type,
        status="pending",
        order_id=body.order_id,
        batch_id=body.batch_id,
        batch_file_id=body.batch_file_id,
        message=body.message,
        created_at=utcnow(),
    )
    db.add(task)
    db.commit()
    db.refresh(task)
    return review_task_detail(task)


@router.get("/exceptions")
def list_exceptions(
    _user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> dict:
    stmt = (
        select(ExceptionReview)
        .options(selectinload(ExceptionReview.order))
        .order_by(ExceptionReview.id.desc())
    )
    rows = db.scalars(stmt).all()
    return {"items": [exception_detail(row) for row in rows]}


@router.post("/exceptions/{review_id}/resolve")
def resolve_exception(
    review_id: int,
    user: User = Depends(require_roles("reviewer")),
    db: Session = Depends(get_db),
) -> dict:
    review = db.scalars(
        select(ExceptionReview)
        .where(ExceptionReview.id == review_id)
        .options(selectinload(ExceptionReview.order))
    ).first()
    if review is None:
        raise HTTPException(status_code=404, detail="异常单不存在")
    if review.status != "open":
        raise HTTPException(status_code=409, detail="该异常已处理")
    review.status = "resolved"
    review.resolved_at = utcnow()
    _ = user
    db.commit()
    db.refresh(review)
    return exception_detail(review)
