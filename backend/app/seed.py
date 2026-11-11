"""初始账号与示例订单。示例数字是演示数据，不是店铺或仓库对照表。"""
from __future__ import annotations

from decimal import Decimal

from sqlalchemy import func, select
from sqlalchemy.orm import Session, sessionmaker

from app.models import (
    ExceptionReview,
    OpsGroup,
    Order,
    OrderLine,
    OrderPackage,
    PackageTracking,
    Shop,
    SkuMapping,
    SkuSuffix,
    SplitAmountReview,
    SplitAmountReviewLine,
    User,
    utcnow,
)
from app.security import hash_password

# 明文只出现在这里和 README。上线前必须更换。
SEED_USERS = (
    ("admin", "admin123", "admin", "系统管理员"),
    ("reviewer", "reviewer123", "reviewer", "审核员"),
    ("ops", "ops123", "ops", "运营"),
)


def seed(session_factory: sessionmaker) -> None:
    db = session_factory()
    try:
        _seed_users(db)
        # 不再灌演示订单/拆单审核/异常，空库启动审核页应为空
        # _seed_orders(db)
        _seed_project2_config(db)
        db.commit()
    except Exception:
        db.rollback()
        raise
    finally:
        db.close()


def _seed_users(db: Session) -> None:
    existing = db.scalar(select(User.id).where(User.username == "admin"))
    if existing is not None:
        return
    for username, password, role, display_name in SEED_USERS:
        db.add(
            User(
                username=username,
                password_hash=hash_password(password),
                role=role,
                display_name=display_name,
                is_active=True,
                created_at=utcnow(),
            )
        )
    db.flush()


def _seed_orders(db: Session) -> None:
    count = db.scalar(select(func.count()).select_from(Order))
    if count:
        return

    first, first_lines = _create_order(
        db,
        {
            "project": "project1",
            "external_order_no": "P1-SEED-10086",
            "shop_account": "示例速卖通店",
            "platform": "速卖通",
            "status": "pending_review",
            "buyer_name": "Jane Cooper",
            "country": "US",
            "currency": "USD",
            "notes": "示例订单，用来展示一单多 SKU、多包裹、多运单和行金额。",
            "raw_payload": {"seed": True},
        },
        [
            {
                "sku": "DEMO-BOX",
                "sku_name": "桌面收纳盒",
                "quantity": Decimal("2.00"),
                "unit_price": Decimal("8.00"),
                "line_amount": Decimal("16.00"),
                "currency": "USD",
            },
            {
                "sku": "DEMO-FILTER",
                "sku_name": "替换滤芯",
                "quantity": Decimal("1.00"),
                "unit_price": Decimal("12.50"),
                "line_amount": Decimal("12.50"),
                "currency": "USD",
            },
            {
                "sku": "DEMO-STICKER",
                "sku_name": "贴纸套装",
                "quantity": Decimal("3.00"),
                "unit_price": Decimal("6.00"),
                "line_amount": Decimal("18.00"),
                "currency": "USD",
            },
        ],
        [
            {
                "package_no": "PKG-10086-A",
                "warehouse_code": "WH-SAMPLE-A",
                "logistics_method": "示例专线",
                "weight_kg": Decimal("1.250"),
                "trackings": [
                    {"tracking_no": "YT10086001", "carrier": "示例承运A"},
                    {"tracking_no": "YT10086001B", "carrier": "示例承运A"},
                ],
            },
            {
                "package_no": "PKG-10086-B",
                "warehouse_code": "WH-SAMPLE-B",
                "logistics_method": "示例小包",
                "weight_kg": Decimal("0.420"),
                "trackings": [
                    {"tracking_no": "SY10086002", "carrier": "示例承运B"},
                ],
            },
        ],
    )
    review = SplitAmountReview(
        order_id=first.id,
        status="pending",
        reason="示例：行金额待人工拆分。本阶段只记审核，不调用店小秘。",
        created_at=utcnow(),
    )
    db.add(review)
    db.flush()
    proposals = (
        (Decimal("16.00"), Decimal("10.00"), "示例拆出部分金额"),
        (Decimal("12.50"), Decimal("12.50"), "维持原行金额"),
        (Decimal("18.00"), Decimal("24.00"), "示例并入其他行"),
    )
    for line, (original, proposed, note) in zip(first_lines, proposals, strict=True):
        db.add(
            SplitAmountReviewLine(
                review_id=review.id,
                order_line_id=line.id,
                sku=line.sku,
                original_amount=original,
                proposed_amount=proposed,
                note=note,
            )
        )

    second, _ = _create_order(
        db,
        {
            "project": "project2",
            "external_order_no": "P2-SEED-20011",
            "shop_account": "示例店铺-待确认",
            "platform": "eBay",
            "status": "pending_review",
            "buyer_name": "Alex Kim",
            "country": "DE",
            "currency": "EUR",
            "row_mark": "green",
            "already_in_guanjia": False,
            "notes": "行标记仅演示绿标字段。Project2 引擎不会据此决定去向。",
            "raw_payload": {"seed": True, "row_mark": "green"},
        },
        [
            {
                "sku": "DEMO-LAMP",
                "sku_name": "台灯",
                "quantity": Decimal("1.00"),
                "unit_price": Decimal("21.00"),
                "line_amount": Decimal("21.00"),
                "currency": "EUR",
            },
            {
                "sku": "DEMO-BULB",
                "sku_name": "灯泡",
                "quantity": Decimal("2.00"),
                "unit_price": Decimal("3.50"),
                "line_amount": Decimal("7.00"),
                "currency": "EUR",
            },
        ],
        [
            {
                "package_no": "PKG-20011-A",
                "warehouse_code": "WH-SAMPLE-C",
                "logistics_method": "示例空运",
                "weight_kg": Decimal("0.800"),
                "trackings": [
                    {"tracking_no": "EB20011001", "carrier": "示例承运C"},
                    {"tracking_no": "EB20011002", "carrier": "示例承运C"},
                ],
            }
        ],
    )
    db.add(
        ExceptionReview(
            order_id=second.id,
            project="project2",
            code="ROUTE_UNRESOLVED",
            severity="warning",
            message="Project2 仍是占位引擎，未执行路由。需要人工确认去向。",
            status="open",
            created_at=utcnow(),
        )
    )

    _create_order(
        db,
        {
            "project": "project2",
            "external_order_no": "P2-SEED-20012",
            "shop_account": "示例-已建档",
            "platform": "速卖通",
            "status": "pending_review",
            "buyer_name": "林晓",
            "country": "FR",
            "currency": "EUR",
            "row_mark": "white",
            "already_in_guanjia": True,
            "notes": "已建档标记只说明数据能存。Phase 0 不会自动拦截或创建店管家订单。",
            "raw_payload": {"seed": True, "row_mark": "white"},
        },
        [
            {
                "sku": "DEMO-MUG",
                "sku_name": "马克杯",
                "quantity": Decimal("4.00"),
                "unit_price": Decimal("5.25"),
                "line_amount": Decimal("21.00"),
                "currency": "EUR",
            }
        ],
        [
            {
                "package_no": "PKG-20012-A",
                "warehouse_code": None,
                "logistics_method": "示例海运",
                "weight_kg": Decimal("2.000"),
                "trackings": [{"tracking_no": "AE20012001", "carrier": "示例承运D"}],
            },
            {
                "package_no": "PKG-20012-B",
                "warehouse_code": None,
                "logistics_method": "示例海运",
                "weight_kg": Decimal("1.100"),
                "trackings": [
                    {"tracking_no": "AE20012002", "carrier": "示例承运D"},
                    {"tracking_no": "AE20012003", "carrier": "示例承运D"},
                ],
            },
        ],
    )


# 从 10 月 4 日运营拆分表和加工表对照里能直接看到的末段后缀，可在界面增删。
OBSERVED_SKU_SUFFIXES = (
    "CQCD",
    "CQCSD",
    "CQCS",
    "CQCES",
    "CECD",
    "HKCQ",
    "CDCS",
    "CDCF",
    "CASD",
    "YXZR",
    "YXZTT",
    "YZXR",
    "JX",
    "XYTU",
    "LYFUS",
    "LYF",
    "TTLNUS2",
    "TTLNUS",
    "TKTK",
    "TK",
)

# 人工海外仓表对得上的 SKU，界面可改。TTLNUS2 要排在 TTLNUS 前面剥离。
OBSERVED_SKU_MAPS = (
    ("JG134", "WX13134"),
    ("JG135", "WX13135"),
    ("1729677372332743562", "WX13201"),
    ("WX13108N-CDCS", "WX13108S"),
)

# 六个运营分组来自样本里已经分开的店铺集合，不是引擎里的写死分支。
OBSERVED_OPS_GROUPS = (
    ("创趣欧洲", ("亚马逊香港创趣欧洲（自营）",)),
    ("宇星欧洲", ("亚马逊宇星欧洲（自营）",)),
    # Shirley 不贴logo 并入 Cdiscount 运营表；店管家仍走 01（见 routing.yaml）
    ("Cdiscount", ("Zeemoto Cdiscount", "Shirley（自营）不贴logo")),
    ("英美半托管", ("宇星美国半托管", "承思英国半托管")),
    # 吴晓虹美国并入林艺奋美国运营表
    ("林艺奋", ("亚马逊林艺奋美国（自营）", "亚马逊林艺奋欧洲（自营）", "亚马逊吴晓虹美国（自营）")),
    ("Zeemoto", ("Zeemoto欧洲半托管", "Zeemoto欧洲半托管2店", "Zeemoto遥控车半托管")),
)


def _seed_project2_config(db: Session) -> None:
    if db.scalar(select(func.count()).select_from(SkuSuffix)) == 0:
        for suffix in OBSERVED_SKU_SUFFIXES:
            db.add(SkuSuffix(project="project2", suffix=suffix, enabled=True))
    if db.scalar(select(func.count()).select_from(SkuMapping).where(SkuMapping.project == "project2")) == 0:
        for source, target in OBSERVED_SKU_MAPS:
            db.add(SkuMapping(project="project2", source_sku=source, target_sku=target, enabled=True, note="10月4日人工海外仓对照"))
    if db.scalar(select(func.count()).select_from(OpsGroup)) == 0:
        for name, accounts in OBSERVED_OPS_GROUPS:
            group = OpsGroup(name=name, project="project2", note="来自 10 月 4 日已拆开的店铺集合，可改")
            db.add(group)
            db.flush()
            for account in accounts:
                # Shirley 进店管家01；其余运营组默认店管家02
                dest = "dian_01" if "Shirley" in account else "dian_02"
                note = (
                    "Shirley 不贴logo：店管家01 + 并入 Cdiscount 运营表"
                    if "Shirley" in account
                    else (
                        "吴晓虹美国：并入林艺奋运营表"
                        if "吴晓虹美国" in account
                        else "样本中的运营分组店铺，去向可在页面修改"
                    )
                )
                db.add(
                    Shop(
                        project="project2",
                        shop_account=account,
                        ops_group_id=group.id,
                        destination=dest,
                        already_in_guanjia=False,
                        routing_note=note,
                    )
                )


def _create_order(db: Session, header: dict, lines: list[dict], packages: list[dict]):
    amount = sum((line["line_amount"] for line in lines), Decimal("0.00"))
    order = Order(order_amount=amount, **header)
    db.add(order)
    db.flush()
    stored: list[OrderLine] = []
    for index, line in enumerate(lines, start=1):
        row = OrderLine(order_id=order.id, line_no=index, **line)
        db.add(row)
        stored.append(row)
    db.flush()
    for package in packages:
        pkg = OrderPackage(
            order_id=order.id,
            package_no=package["package_no"],
            warehouse_code=package.get("warehouse_code"),
            logistics_method=package.get("logistics_method"),
            weight_kg=package.get("weight_kg"),
            status=package.get("status", "packed"),
        )
        db.add(pkg)
        db.flush()
        for tracking in package["trackings"]:
            db.add(
                PackageTracking(
                    order_id=order.id,
                    package_id=pkg.id,
                    tracking_no=tracking["tracking_no"],
                    carrier=tracking.get("carrier"),
                )
            )
    return order, stored
