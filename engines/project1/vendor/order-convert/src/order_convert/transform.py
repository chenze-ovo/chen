# -*- coding: utf-8 -*-
"""字段转换与分流（纯逻辑，无 IO）。"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Literal

from pypinyin import Style, lazy_pinyin

from .config import Rules
from .io_excel import norm

Severity = Literal["error", "warning", "info"]

# 严重错误类型：存在任一时禁止正式导出（blocked）
ERROR_TYPES = frozenset(
    {
        "UNKNOWN_SHOP",
        "MISSING_COLUMNS",
        "EMPTY_INPUT",
        # DUPLICATE_ORDER_NO：原始单号重复改为 warning，由店管家02订单编号去重处理
        "DUPLICATE_TRACKING",
        "SORT_KEY_EMPTY",
        "INVALID_QUANTITY",
        "INVALID_AMOUNT",
    }
)

# 待复核类型：存在任一时标记 needs_review，禁止正式可导入文件名
REVIEW_TYPES = frozenset(
    {
        "UNKNOWN_SKU_SUFFIX",
        "MULTI_SKU",
    }
)


@dataclass
class Anomaly:
    type: str
    订单号: str
    action: str
    severity: Severity = "warning"
    detail: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict:
        d = {
            "type": self.type,
            "severity": self.severity,
            "订单号": self.订单号,
            "action": self.action,
        }
        d.update(self.detail)
        return d


def pinyin_key(s: str) -> str:
    return "".join(lazy_pinyin(s, style=Style.NORMAL))



_SKU_LIKE_RE = re.compile(r"^[A-Za-z0-9_-]+$")


def is_sku_like(text: Any) -> bool:
    """备注列整格是否像单个 SKU/货号（用于替换原 SKU）。

    规则：trim 后单 token、仅字母数字与 -_、长度 3–32、至少各含一个字母与一个数字。
    排除纯中文、含空格长句、「贴日本标签」等。
    """
    s = norm(text) if not isinstance(text, str) else text.strip()
    if not s:
        return False
    if any(ch.isspace() for ch in s):
        return False
    if len(s) < 3 or len(s) > 32:
        return False
    if not _SKU_LIKE_RE.fullmatch(s):
        return False
    has_letter = any(c.isalpha() for c in s)
    has_digit = any(c.isdigit() for c in s)
    return has_letter and has_digit



def find_sku_from_note(raw: dict, rules: Rules) -> tuple[str, str] | None:
    """按 sku_from_note_columns 优先级找第一个「像 SKU」的备注列。

    返回 (列名, 货号文本)；都未命中则 None。先命中优先，不互相覆盖。
    """
    for col in rules.sku_from_note_columns:
        if col not in raw and raw.get(col) is None:
            # 列可能缺失：仍尝试 get（缺省 None）
            pass
        val = norm(raw.get(col))
        if is_sku_like(val):
            return col, val
    return None


def raw_currency(raw: dict, rules: Rules) -> str:
    """读取原表币种（按配置列名优先级）。"""
    for col in rules.currency_columns:
        if col in raw and raw.get(col) not in (None, ""):
            return norm(raw.get(col))
    # 列名可能存在但值为空：继续；若表头都没有也返回空
    for col in rules.currency_columns:
        if col in raw:
            return norm(raw.get(col))
    return ""


def currency_suffix_for_logistics(raw: dict, rules: Rules) -> str:
    """仅亚马逊店铺追加非欧元等 skip 币种；非亚马逊返回空串。"""
    if not rules.currency_append_to_logistics:
        return ""
    if not is_amazon(raw, rules):
        return ""
    cur = raw_currency(raw, rules)
    if not cur:
        return ""
    if cur in rules.currency_skip or cur.upper() in rules.currency_skip:
        return ""
    return cur


def is_china_direct(raw: dict, rules: Rules) -> bool:
    """仅看原始拣货备注是否包含完整标记文案。"""
    return rules.china_marker in norm(raw.get("拣货备注"))


def is_amazon(raw: dict, rules: Rules) -> bool:
    return rules.amazon_keyword in str(raw.get("店铺账号") or "")


def is_ebay(raw: dict) -> bool:
    return "ebay" in str(raw.get("店铺账号") or "").lower()


def is_known_d01(raw: dict, rules: Rules) -> bool:
    shop = str(raw.get("店铺账号") or "")
    shop_l = shop.lower()
    for key in rules.d01_known:
        if key.lower() in shop_l or key in shop:
            return True
    return False


def classify_dest(raw: dict, rules: Rules, anomalies: list[Anomaly]) -> str:
    if is_amazon(raw, rules):
        return "D02"
    if is_known_d01(raw, rules):
        return "D01"
    anomalies.append(
        Anomaly(
            type="UNKNOWN_SHOP",
            severity="error",
            订单号=norm(raw.get("订单号")),
            action="严重错误：未知店铺进入待核对；禁止正式导出店管家01/02",
            detail={"店铺账号": raw.get("店铺账号")},
        )
    )
    return "PENDING"


def split_sku_lines(sku_cell: Any) -> list[str]:
    """按换行拆分 SKU 单元格；忽略空行。不擅自拆数量。"""
    if sku_cell is None:
        return []
    text = str(sku_cell).replace("\r\n", "\n").replace("\r", "\n")
    return [ln.strip() for ln in text.split("\n") if ln.strip()]


def _transform_one_sku(
    sku: str,
    note: str,
    rules: Rules,
    anomalies: list[Anomaly],
    order_no: str,
) -> tuple[str, str]:
    """单行 SKU → (店小秘SKU, 基础SKU)。

    从右向左迭代剥离「-后缀」：
      - 日本后缀 + 备注命中 → 店小秘追加备注，02 仅基础码，然后停止
      - 后缀在 strip 白名单 → 剥掉，继续
      - 后缀纯字母但不在白名单 → 剥掉并记 UNKNOWN_SKU_SUFFIX（needs_review），继续
      - 后缀含数字（如 020）→ 停止，保留为产品码一部分（避免 18MB-020 被剥成 18MB）
    别名（如 M00007→M0007）在调用方 apply_sku_aliases 中处理。
    """
    working = sku
    while "-" in working:
        base, suf = working.rsplit("-", 1)
        if suf == rules.japan_suffix and rules.japan_note in note:
            return f"{base} {note}", base
        if suf in rules.strip_suffixes:
            working = base
            continue
        if suf.isalpha():
            anomalies.append(
                Anomaly(
                    type="UNKNOWN_SKU_SUFFIX",
                    severity="warning",
                    订单号=order_no,
                    action="已去掉未知SKU后缀，仅输出基础SKU；需确认后缀白名单，禁止正式导入",
                    detail={"SKU": sku, "suffix": suf},
                )
            )
            working = base
            continue
        # 含数字或其它非纯字母段：视为产品码一部分，停止剥离
        break
    return working, working



def apply_sku_aliases(dx_sku: str, base_sku: str, rules: Rules) -> tuple[str, str]:
    """对最终输出 SKU 应用配置别名（精确匹配；多行逐行；日本标签仅替换首段基础码）。

    在去后缀 / 拣货备注替换 / 日本标签处理之后调用。
    中国直发固定 M00016 不在别名表中时保持不变。
    """
    aliases = rules.sku_aliases
    if not aliases:
        return dx_sku, base_sku

    def alias_base_line(line: str) -> str:
        return aliases.get(line, line)

    def alias_dx_line(line: str) -> str:
        # 日本标签等：形如 "BASE 备注文字"
        if " " in line:
            first, rest = line.split(" ", 1)
            if first in aliases:
                return aliases[first] + " " + rest
        return aliases.get(line, line)

    dx_parts = dx_sku.split("\n") if dx_sku is not None else [""]
    base_parts = base_sku.split("\n") if base_sku is not None else [""]
    return (
        "\n".join(alias_dx_line(x) for x in dx_parts),
        "\n".join(alias_base_line(x) for x in base_parts),
    )



def apply_sku_prefix_rules(
    dx_sku: str,
    base_sku: str,
    rules: Rules,
    anomalies: list[Anomaly],
    order_no: str,
) -> tuple[str, str]:
    """对最终输出 SKU 应用前缀剥离（精确 starts_with；多行逐行；日本标签仅处理首段）。

    在 sku_aliases 之后调用。例：when_starts_with=WXM + strip_prefix=WX → WXM0021→M0021；
    WX10225 不以 WXM 开头，不动。中国直发 M00016 不命中。
    """
    prefix_rules = rules.sku_prefix_rules
    if not prefix_rules:
        return dx_sku, base_sku

    def strip_token(token: str) -> tuple[str, dict | None]:
        for rule in prefix_rules:
            when = rule["when_starts_with"]
            strip = rule["strip_prefix"]
            if token.startswith(when) and token.startswith(strip):
                after = token[len(strip) :]
                meta = {
                    "when_starts_with": when,
                    "strip_prefix": strip,
                    "before": token,
                    "after": after,
                }
                return after, meta
        return token, None

    def strip_dx_line(line: str) -> tuple[str, dict | None]:
        if " " in line:
            first, rest = line.split(" ", 1)
            new_first, meta = strip_token(first)
            return new_first + " " + rest, meta
        return strip_token(line)

    dx_parts = dx_sku.split("\n") if dx_sku is not None else [""]
    base_parts = base_sku.split("\n") if base_sku is not None else [""]
    new_dx: list[str] = []
    new_base: list[str] = []
    for line in dx_parts:
        n, meta = strip_dx_line(line)
        new_dx.append(n)
        if meta:
            anom_type = (
                "SKU_STRIP_WXM_PREFIX"
                if meta["when_starts_with"] == "WXM" and meta["strip_prefix"] == "WX"
                else "SKU_STRIP_PREFIX"
            )
            anomalies.append(
                Anomaly(
                    type=anom_type,
                    severity="info",
                    订单号=order_no,
                    action="SKU 前缀规则：已去掉配置前缀以对齐店管家货号",
                    detail=meta,
                )
            )
    for line in base_parts:
        n, _meta = strip_token(line)
        new_base.append(n)
    return "\n".join(new_dx), "\n".join(new_base)


def finalize_skus(
    dx_sku: str,
    base_sku: str,
    rules: Rules,
    anomalies: list[Anomaly],
    order_no: str,
) -> tuple[str, str]:
    """别名 → 前缀规则，统一收口。"""
    dx_sku, base_sku = apply_sku_aliases(dx_sku, base_sku, rules)
    return apply_sku_prefix_rules(dx_sku, base_sku, rules, anomalies, order_no)


def transform_sku(raw: dict, rules: Rules, anomalies: list[Anomaly]) -> tuple[str, str]:
    """返回 (店小秘SKU, 02基础SKU/条码)。

    优先级：
      1) 中国直发 M00016（拣货备注含完整标记）→ 固定 M00016，不被备注替换覆盖
      2) 非多 SKU：按 sku_from_note_columns 优先级，整格像 SKU/货号则替换输出 SKU；
         替换后再去 -后缀；备注列保持原文；记 SKU_FROM_NOTE（info，含 source_column）
      3) 否则现有去后缀 / 日本标签逻辑
      4) 对输出基础 SKU 应用 sku_aliases（如 M00007→M0007）；多行逐行
      5) 再应用 sku_prefix_rules（如 WXM…→去 WX 得 M…；WX+数字不动）

    多行 SKU：对各行分别去后缀；输出用换行拼接；不拆明细；标 MULTI_SKU 待复核；
    不用单个备注 SKU 覆盖多行。
    """
    order_no = norm(raw.get("订单号"))
    pick_note = norm(raw.get("拣货备注"))
    if is_china_direct(raw, rules):
        return finalize_skus(rules.china_sku, rules.china_sku, rules, anomalies, order_no)

    raw_sku = raw.get("SKU")
    lines = split_sku_lines(raw_sku)
    original = "" if raw_sku is None else str(raw_sku).strip()

    if len(lines) > 1:
        per_line: list[dict[str, str]] = []
        dx_parts: list[str] = []
        base_parts: list[str] = []
        for ln in lines:
            dx_one, base_one = _transform_one_sku(ln, pick_note, rules, anomalies, order_no)
            dx_one, base_one = finalize_skus(dx_one, base_one, rules, anomalies, order_no)
            per_line.append({"sku": ln, "dx_sku": dx_one, "base_sku": base_one})
            dx_parts.append(dx_one)
            base_parts.append(base_one)
        anomalies.append(
            Anomaly(
                type="MULTI_SKU",
                severity="warning",
                订单号=order_no,
                action=(
                    "待复核：SKU单元格含多个SKU（换行）；"
                    "不拆销售明细、不臆造各SKU数量；禁止正式导入"
                ),
                detail={
                    "SKU": original,
                    "sku_lines": lines,
                    "line_count": len(lines),
                    "per_line": per_line,
                },
            )
        )
        # 输出去后缀后的各行（换行拼接），仍不拆成多条销售明细；不用备注覆盖多行
        return finalize_skus("\n".join(dx_parts), "\n".join(base_parts), rules, anomalies, order_no)

    # 单 SKU（或空 SKU 单元格）：按配置列优先级找货号替换
    hit = find_sku_from_note(raw, rules)
    if hit:
        source_column, note_sku = hit
        dx_sku, base_sku = _transform_one_sku(
            note_sku, pick_note, rules, anomalies, order_no
        )
        dx_sku, base_sku = finalize_skus(dx_sku, base_sku, rules, anomalies, order_no)
        anomalies.append(
            Anomaly(
                type="SKU_FROM_NOTE",
                severity="info",
                订单号=order_no,
                action=(
                    f"{source_column}为SKU/货号，已替换店小秘SKU与店管家编号/条码；"
                    "备注列保持原文"
                ),
                detail={
                    "原始SKU": original,
                    "source_column": source_column,
                    "note_value": note_sku,
                    "dx_SKU": dx_sku,
                    "base_sku": base_sku,
                },
            )
        )
        return dx_sku, base_sku

    if len(lines) == 1:
        dx_one, base_one = _transform_one_sku(
            lines[0], pick_note, rules, anomalies, order_no
        )
        return finalize_skus(dx_one, base_one, rules, anomalies, order_no)

    return finalize_skus(original, original, rules, anomalies, order_no)

def transform_order_no_dx(raw: dict, rules: Rules) -> str:
    ono = norm(raw.get("订单号"))
    if is_china_direct(raw, rules):
        return ono + rules.china_order_suffix
    return ono


def should_append_sku_to_logistics(base_sku: str, rules: Rules) -> str | None:
    """若基础 SKU（整格，非多行）在 append 列表中，返回要追加的标记；否则 None。"""
    if not base_sku or "\n" in base_sku:
        return None
    marker = base_sku.strip()
    if marker in rules.logistics_append_skus:
        return marker
    return None


def _append_currency(text: str, raw: dict, rules: Rules) -> str:
    """在业务后缀之后无空格追加币种（仅亚马逊；欧元等 skip 不追加）。"""
    suf = currency_suffix_for_logistics(raw, rules)
    if not suf:
        return text
    return text + suf


def transform_logistics_dx(raw: dict, rules: Rules, base_sku: str = "") -> str:
    log = norm(raw.get("物流方式"))
    if is_china_direct(raw, rules):
        return _append_currency(log + rules.china_marker, raw, rules)
    marker = should_append_sku_to_logistics(base_sku, rules)
    if marker:
        return _append_currency(log + marker, raw, rules)  # 无空格
    return _append_currency(log, raw, rules)


def d02_kefu_note(raw: dict, rules: Rules, base_sku: str = "") -> str:
    log = norm(raw.get("物流方式"))
    if is_china_direct(raw, rules):
        return _append_currency(log + " " + rules.china_marker, raw, rules)
    marker = should_append_sku_to_logistics(base_sku, rules)
    if marker:
        return _append_currency(log + marker, raw, rules)  # 无空格
    return _append_currency(log, raw, rules)


def d01_original_order_no(raw: dict, rules: Rules) -> str:
    ono = norm(raw.get("订单号"))
    if is_ebay(raw) and rules.raw.get("ebay", {}).get("truncate_order_no", True) and "-" in ono:
        return ono.rsplit("-", 1)[-1]
    return ono


def d02_amount(raw: dict, rules: Rules, anomalies: list[Anomaly]) -> float:
    if is_china_direct(raw, rules) and rules.raw["china_direct"].get("d02_amount_zero", True):
        return 0.0
    amt = raw.get("订单金额")
    if amt in ("", None):
        return 0.0
    try:
        return float(amt)
    except (TypeError, ValueError):
        anomalies.append(
            Anomaly(
                type="INVALID_AMOUNT",
                severity="error",
                订单号=norm(raw.get("订单号")),
                action="严重错误：订单金额无法解析",
                detail={"订单金额": amt},
            )
        )
        return 0.0


def parse_quantity(raw: dict, anomalies: list[Anomaly]) -> Any:
    qty = raw.get("产品总数")
    if qty in ("", None):
        anomalies.append(
            Anomaly(
                type="INVALID_QUANTITY",
                severity="error",
                订单号=norm(raw.get("订单号")),
                action="严重错误：产品总数为空",
                detail={},
            )
        )
        return 0
    try:
        f = float(qty)
        return int(f) if abs(f - int(f)) < 1e-9 else f
    except (TypeError, ValueError):
        anomalies.append(
            Anomaly(
                type="INVALID_QUANTITY",
                severity="error",
                订单号=norm(raw.get("订单号")),
                action="严重错误：产品总数无法解析",
                detail={"产品总数": qty},
            )
        )
        return 0


def d02_phone(raw: dict, rules: Rules, anomalies: list[Anomaly], dest: str) -> str:
    phone = norm(raw.get("收货人电话"))
    mobile = norm(raw.get("收货人手机"))
    if rules.phone_auto_fallback and not phone and mobile:
        return mobile
    if dest == rules.phone_warn_dest and not phone and mobile and not rules.phone_auto_fallback:
        anomalies.append(
            Anomaly(
                type="PHONE_FALLBACK_UNCONFIRMED",
                severity="warning",
                订单号=norm(raw.get("订单号")),
                action="警告：D02 电话空且手机非空；不自动回退",
                detail={"收货人手机": mobile},
            )
        )
    return phone


def sort_raw_rows(raw_rows: list[dict], rules: Rules, anomalies: list[Anomaly]) -> list[dict]:
    indexed = []
    for i, r in enumerate(raw_rows):
        log = norm(r.get("物流方式"))
        if not log:
            anomalies.append(
                Anomaly(
                    type="SORT_KEY_EMPTY",
                    severity="error",
                    订单号=norm(r.get("订单号")),
                    action="严重错误：原始物流方式为空，无法可靠排序",
                )
            )
        indexed.append((pinyin_key(log), i, r))
    indexed.sort(key=lambda t: (t[0], t[1]))
    return [t[2] for t in indexed]


REQUIRED_COLUMNS = [
    "订单号",
    "店铺账号",
    "SKU",
    "物流方式",
    "运单号",
    "拣货备注",
    "产品总数",
    "订单金额",
    "收货人姓名",
    "收货人国家",
    "详细地址",
    "收货人电话",
    "收货人手机",
]


def validate_headers(headers: list[str], anomalies: list[Anomaly]) -> None:
    missing = [c for c in REQUIRED_COLUMNS if c not in headers]
    if missing:
        anomalies.append(
            Anomaly(
                type="MISSING_COLUMNS",
                severity="error",
                订单号="",
                action="严重错误：原始表缺少必要列，中止正式导出",
                detail={"missing": missing},
            )
        )


def validate_duplicates(raw_rows: list[dict], anomalies: list[Anomaly]) -> None:
    seen_ono: dict[str, int] = {}
    seen_track: dict[str, int] = {}
    for r in raw_rows:
        ono = norm(r.get("订单号"))
        track = norm(r.get("运单号"))
        if ono:
            if ono in seen_ono:
                anomalies.append(
                    Anomaly(
                        type="DUPLICATE_ORDER_NO",
                        severity="warning",
                        订单号=ono,
                        action="警告：原始订单号重复；店管家02将按规则去重后缀，不阻断导出",
                        detail={"first_row": seen_ono[ono], "row": r.get("_excel_row")},
                    )
                )
            else:
                seen_ono[ono] = r.get("_excel_row")
        if track:
            if track in seen_track:
                anomalies.append(
                    Anomaly(
                        type="DUPLICATE_TRACKING",
                        severity="error",
                        订单号=ono,
                        action="严重错误：运单号重复",
                        detail={"运单号": track, "first_row": seen_track[track]},
                    )
                )
            else:
                seen_track[track] = r.get("_excel_row")


_TRAILING_NUM_SUFFIX_RE = re.compile(r"^(.*)-(\d+)$")


def _split_trailing_num_suffix(order_no: str) -> tuple[str, int | None]:
    """若订单号已以 -数字 结尾，返回 (基号, 数字)；否则 (原号, None)。"""
    m = _TRAILING_NUM_SUFFIX_RE.match(order_no)
    if not m:
        return order_no, None
    return m.group(1), int(m.group(2))


def _next_d02_order_no(desired: str, used: set[str]) -> str:
    """在 used 集合中为 desired 找下一个不冲突的编号。

    - desired 未被占用 → 原样返回
    - 否则：若 desired 形如 A-n，从 A-(n+1) 起递增；否则从 desired-1 起递增
    """
    if desired and desired not in used:
        return desired
    base, existing_n = _split_trailing_num_suffix(desired)
    n = (existing_n + 1) if existing_n is not None else 1
    while True:
        cand = f"{base}-{n}"
        if cand not in used:
            return cand
        n += 1


def _forced_d02_suffix(original: str, used: set[str]) -> str:
    """后续重复条：相对原号强制追加 -n（原号已是 -k 则从 k+1 起），并避开 used。"""
    base, existing_n = _split_trailing_num_suffix(original)
    n = (existing_n + 1) if existing_n is not None else 1
    while True:
        cand = f"{base}-{n}"
        if cand not in used:
            return cand
        n += 1


def dedupe_d02_order_nos(
    d02_rows: list[dict],
    anomalies: list[Anomaly],
    mappings: list[dict] | None = None,
) -> list[dict]:
    """店管家02「订单编号」去重：同号多条时第一条保持，其后加 -1/-2…

    若原号已以 -数字 结尾（如中国直发已带 -1），继续递增避免碰撞（A-1 → A-2）。
    仅改 D02 订单编号；店小秘订单号不动。不因本规则阻断导出。
    """
    if mappings is not None:
        for m in mappings:
            m.setdefault("d02_订单编号", "")

    if not d02_rows:
        return []

    from collections import OrderedDict

    groups: "OrderedDict[str, list[int]]" = OrderedDict()
    for i, row in enumerate(d02_rows):
        ono = norm(row.get("订单编号"))
        row["订单编号"] = ono
        groups.setdefault(ono, []).append(i)

    used: set[str] = set()
    changes: list[dict] = []

    def _record(idx: int, original: str, final: str, reason: str) -> None:
        changes.append(
            {"row_index": idx, "original": original, "final": final, "reason": reason}
        )
        anomalies.append(
            Anomaly(
                type="D02_ORDER_NO_DEDUP",
                severity="info",
                订单号=original,
                action="店管家02订单编号去重：避免同号冲突",
                detail={
                    "original_订单编号": original,
                    "final_订单编号": final,
                    "d02_data_line": idx + 1,
                    "reason": reason,
                },
            )
        )

    for original, indices in groups.items():
        first_idx = indices[0]
        first_final = _next_d02_order_no(original, used)
        d02_rows[first_idx]["订单编号"] = first_final
        used.add(first_final)
        if first_final != original:
            _record(first_idx, original, first_final, "first_collision")

        for idx in indices[1:]:
            final = _forced_d02_suffix(original, used)
            d02_rows[idx]["订单编号"] = final
            used.add(final)
            _record(idx, original, final, "duplicate")

    if mappings is not None:
        # 店管家02行号 = d02_line+1（表头占第1行，首条数据为2）
        for m in mappings:
            r = m.get("店管家02行号")
            if r == "" or r is None:
                m["d02_订单编号"] = ""
                continue
            try:
                excel_row = int(r)
            except (TypeError, ValueError):
                m["d02_订单编号"] = ""
                continue
            idx = excel_row - 2
            if 0 <= idx < len(d02_rows):
                m["d02_订单编号"] = d02_rows[idx]["订单编号"]
            else:
                m["d02_订单编号"] = ""

    return changes
