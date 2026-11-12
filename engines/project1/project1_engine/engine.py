"""调用原样接入的 order_convert，整理状态与导出结果。

禁止在此重写 transform / pipeline / rules.yaml。
"""
from __future__ import annotations

import hashlib
import sys
from pathlib import Path
from typing import Any


def _package_root() -> Path:
    # engines/project1/project1_engine/engine.py → engines/project1
    return Path(__file__).resolve().parents[1]


def vendor_src_path() -> Path:
    return _package_root() / "vendor" / "order-convert" / "src"


def _ensure_vendor_on_path() -> None:
    src = str(vendor_src_path())
    if src not in sys.path:
        sys.path.insert(0, src)


def import_order_convert():
    _ensure_vendor_on_path()
    import order_convert

    return order_convert


def _sha256(path: Path) -> str | None:
    if not path.exists() or not path.is_file():
        return None
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _rules_fingerprint() -> dict[str, Any]:
    package = import_order_convert()
    pkg_dir = Path(package.__file__).resolve().parent
    builtin = pkg_dir / "resources" / "rules.yaml"
    # config.py：包文件 parents[2] / config/rules.yaml
    project_rules = pkg_dir.parents[1] / "config" / "rules.yaml"
    builtin_sha = _sha256(builtin)
    project_sha = _sha256(project_rules)
    return {
        "builtin_rules": str(builtin),
        "builtin_sha256": builtin_sha,
        "project_rules": str(project_rules),
        "project_sha256": project_sha,
        "rules_identical": bool(builtin_sha and builtin_sha == project_sha),
    }


# 仅记录性、不进入「异常」列表的类型（即使 severity 未标 info 也不展示）
_INFO_TYPES = frozenset(
    {
        "SKU_FROM_NOTE",
        "SKU_STRIP_PREFIX",
        "SKU_STRIP_WXM_PREFIX",
        "D02_ORDER_NO_DEDUP",
    }
)

def _as_dict(item: Any) -> dict[str, Any]:
    if isinstance(item, dict):
        return item
    to_dict = getattr(item, "to_dict", None)
    if callable(to_dict):
        return to_dict()
    return {"message": str(item)}


def _issue_brief(items: list) -> list[dict[str, Any]]:
    """把 anomalies / errors / warnings 压成前端可展示的摘要。

    保留 Anomaly.to_dict 铺平进来的明细（SKU、suffix、店铺账号 等），
    不丢字段。不在此处静默截断；条数以返回列表长度为准。
    """
    out: list[dict[str, Any]] = []
    for item in items:
        raw = _as_dict(item)
        row = dict(raw)
        row["type"] = raw.get("type") or raw.get("code") or ""
        row["severity"] = raw.get("severity") or ""
        row["订单号"] = raw.get("订单号") or raw.get("order_no") or ""
        row["message"] = (
            raw.get("action") or raw.get("message") or raw.get("msg") or ""
        )
        out.append(row)
    return out


def is_visible_issue(item: Any) -> bool:
    """errors / 待复核 / warning 可见；info 级记录不进入异常列表。"""
    if not isinstance(item, dict):
        return bool(str(item).strip())
    typ = str(item.get("type") or item.get("code") or "")
    severity = str(item.get("severity") or "").lower()
    if typ in _INFO_TYPES or severity == "info":
        return False
    return True


def _dedupe_key(item: dict[str, Any], file_name: str = "") -> tuple:
    typ = str(item.get("type") or item.get("code") or "")
    order_no = str(item.get("订单号") or item.get("order_no") or "")
    suffix = str(item.get("suffix") or "")
    sku = str(item.get("SKU") or item.get("sku") or "")
    extra = suffix or sku or str(item.get("店铺账号") or item.get("运单号") or item.get("message") or item.get("action") or "")
    return (file_name, order_no, typ, extra)


def explain_issue(item: dict[str, Any], *, file_name: str = "") -> dict[str, Any]:
    """把一条异常翻成白话：订单号、什么问题、该怎么处理。界面不展示英文码。"""
    typ = str(item.get("type") or item.get("code") or "")
    order_no = str(item.get("订单号") or item.get("order_no") or "")
    problem, how_to = _plain_words(typ, item, file_name=file_name)
    return {
        "订单号": order_no,
        "file": file_name,
        "problem": problem,
        "how_to": how_to,
    }


def _plain_words(typ: str, item: dict[str, Any], *, file_name: str = "") -> tuple[str, str]:
    sku = str(item.get("SKU") or item.get("sku") or "").strip()
    suffix = str(item.get("suffix") or "").strip()

    if typ == "UNKNOWN_SKU_SUFFIX":
        base = ""
        if sku and suffix and sku.endswith(f"-{suffix}"):
            base = sku[: -(len(suffix) + 1)]
        elif sku and "-" in sku:
            base = sku.rsplit("-", 1)[0]
        if sku and suffix and base:
            problem = f"货号带了系统不认识的后缀「{suffix}」（原货号 {sku}），已先剥成 {base}。"
        elif sku and suffix:
            problem = f"货号带了系统不认识的后缀「{suffix}」（原货号 {sku}），已先剥成基础货号。"
        elif suffix:
            problem = f"货号带了系统不认识的后缀「{suffix}」，已先剥成基础货号。"
        else:
            problem = "货号带了系统不认识的后缀，已先剥成基础货号。"
        how = (
            "请确认这个后缀要不要加进中国直发的白名单；确认后告诉我们，或改好原表货号再转一次。"
            "这几单暂时不要直接拿去导入。"
        )
        return problem, how

    if typ == "MULTI_SKU":
        lines = item.get("sku_lines")
        extra = ""
        if isinstance(lines, list) and lines:
            extra = f"（{ '、'.join(str(x) for x in lines) }）"
        elif sku:
            extra = f"（{sku}）"
        problem = f"这一单有多个货号{extra}，需要打开结果表核对货号和数量。"
        how = "请打开结果表，核对每个货号和对应数量后再决定是否导入。这几单暂时不要直接拿去导入。"
        return problem, how

    if typ == "UNKNOWN_SHOP":
        shop = str(item.get("店铺账号") or "").strip()
        if shop:
            problem = f"店铺账号「{shop}」不在已知名单里，需要确认是哪一类店。"
        else:
            problem = "店铺账号不在已知名单里，需要确认是哪一类店。"
        how = "请确认这家店是亚马逊还是其他已对接的店铺；确认后告诉我们。这几单暂时不要直接拿去导入。"
        return problem, how

    if typ == "MISSING_COLUMNS":
        missing = item.get("missing") or []
        if isinstance(missing, list) and missing:
            cols = "、".join(str(x) for x in missing)
            problem = f"原表缺少必要的列（{cols}）。"
        else:
            problem = "原表缺少必要的列。"
        how = "请在原表补上这些列后再转一次。"
        return problem, how

    if typ == "EMPTY_INPUT":
        problem = "上传的文件里没有订单。"
        if file_name:
            problem = f"文件「{file_name}」里没有订单。"
        how = "请检查原表是否有数据，确认后再重新上传转换。"
        return problem, how

    if typ == "INVALID_QUANTITY":
        qty = item.get("产品总数")
        if qty not in (None, ""):
            problem = f"这一单的产品总数无法识别（填写的是「{qty}」）。"
        else:
            problem = "这一单的产品总数是空的或无法识别。"
        how = "请打开原表核对产品总数，改好后再转一次。"
        return problem, how

    if typ == "INVALID_AMOUNT":
        amt = item.get("订单金额")
        if amt not in (None, ""):
            problem = f"这一单的订单金额无法识别（填写的是「{amt}」）。"
        else:
            problem = "这一单的订单金额无法识别。"
        how = "请打开原表核对订单金额，改好后再转一次。"
        return problem, how

    if typ == "DUPLICATE_TRACKING":
        track = str(item.get("运单号") or "").strip()
        if track:
            problem = f"运单号「{track}」和其他订单重复。"
        else:
            problem = "这一单的运单号和其他订单重复。"
        how = "请打开原表核对运单号，改好后再转一次。这几单暂时不要直接拿去导入。"
        return problem, how

    if typ == "DUPLICATE_ORDER_NO":
        problem = "这一单的订单号和其他订单重复。"
        how = "店管家02会按规则给重复编号加后缀，请打开结果表核对后再导入。"
        return problem, how

    if typ == "SORT_KEY_EMPTY":
        problem = "这一单的物流方式是空的，没法可靠排序。"
        how = "请打开原表补上物流方式后再转一次。"
        return problem, how

    if typ == "PHONE_FALLBACK_UNCONFIRMED":
        mobile = str(item.get("收货人手机") or "").strip()
        if mobile:
            problem = f"店管家02 的电话是空的，只填了手机（{mobile}），系统没有自动替过去。"
        else:
            problem = "店管家02 的电话是空的，只填了手机，系统没有自动替过去。"
        how = "请打开结果表确认电话是否要用这个手机号。"
        return problem, how

    if typ in {"CONVERT_FAILED", "error"} or not typ:
        msg = str(item.get("message") or item.get("action") or item.get("msg") or "").strip()
        if file_name:
            problem = f"文件「{file_name}」转换失败。"
        else:
            problem = "这份文件转换失败了。"
        if msg and _looks_plain_chinese(msg) and len(msg) < 80:
            problem = f"{problem}{msg}" if problem.endswith("。") else f"{problem}：{msg}"
        how = "请检查原表后重新上传再转一次。"
        return problem, how

    msg = str(item.get("message") or item.get("action") or item.get("msg") or "").strip()
    if msg and _looks_plain_chinese(msg):
        problem = msg
    else:
        problem = "转换时发现需要核对的问题。"
    how = "请打开结果表核对该单，确认无误后再导入。"
    return problem, how


def _looks_plain_chinese(text: str) -> bool:
    if not text:
        return False
    # 避免把英文状态码直接展示给用户
    if any(ch.isascii() and ch.isupper() for ch in text) and "_" in text and not any(
        "\u4e00" <= ch <= "\u9fff" for ch in text
    ):
        return False
    return any("\u4e00" <= ch <= "\u9fff" for ch in text)


def collect_visible_issues(files: list[dict[str, Any]] | None) -> list[dict[str, Any]]:
    """从各文件的 errors + warnings 收集去重后的白话异常列表。"""
    rows: list[dict[str, Any]] = []
    seen: set[tuple] = set()
    for file_item in files or []:
        name = str(file_item.get("file") or "")
        bucket = list(file_item.get("errors") or []) + list(file_item.get("warnings") or [])
        if not bucket and file_item.get("message") and file_item.get("export_status") in {
            "error",
            "blocked",
            "failed",
        }:
            bucket = [
                {
                    "type": "CONVERT_FAILED",
                    "severity": "error",
                    "订单号": "",
                    "message": file_item.get("message"),
                }
            ]
        for raw in bucket:
            item = raw if isinstance(raw, dict) else {"message": str(raw)}
            if not is_visible_issue(item):
                continue
            key = _dedupe_key(item, name)
            if key in seen:
                continue
            seen.add(key)
            rows.append(explain_issue(item, file_name=name))
    return rows


def count_visible_issues(files: list[dict[str, Any]] | None) -> tuple[int, int]:
    """返回 (列表可见条数, 引擎计数上限)。后者用于历史批次被截断时提示总数。"""
    visible = collect_visible_issues(files)
    listed = len(visible)
    counted = 0
    for file_item in files or []:
        counted += int(file_item.get("error_count") or 0) + int(file_item.get("warning_count") or 0)
    return listed, max(listed, counted)


def render_exception_notes(issues: list[dict[str, Any]]) -> str:
    """一行一单的白话说明，供下载「异常说明.txt」。"""
    lines = [
        "中国直发转换异常说明",
        f"共 {len(issues)} 条需要查看。请先核对，暂时不要直接拿去导入。",
        "",
    ]
    for index, issue in enumerate(issues, start=1):
        order_no = str(issue.get("订单号") or "").strip() or "（整份文件）"
        problem = str(issue.get("problem") or "").strip()
        how_to = str(issue.get("how_to") or "").strip()
        line = f"{index}. 订单号 {order_no}：{problem}"
        if how_to:
            line += f" {how_to}"
        lines.append(line)
    lines.append("")
    return "\n".join(lines)


def _write_exception_notes(out_dir: Path, issues: list[dict[str, Any]]) -> Path | None:
    if not issues:
        return None
    path = out_dir / "异常说明.txt"
    path.write_text(render_exception_notes(issues), encoding="utf-8")
    return path


def status() -> dict[str, Any]:
    """引擎就绪、版本、rules 指纹。"""
    try:
        package = import_order_convert()
        from order_convert.resources_path import (
            default_template_d01,
            default_template_d02,
            default_template_dx,
        )

        templates = {
            "dianxiaomi": default_template_dx(),
            "dian_01": default_template_d01(),
            "dian_02": default_template_d02(),
        }
        fingerprint = _rules_fingerprint()
        return {
            "project": "project1",
            "engine": "order_convert",
            "version": package.__version__,
            "ready": all(path.exists() for path in templates.values())
            and fingerprint["rules_identical"],
            "business_rules_modified_by_erp": False,
            "rules_identical": fingerprint["rules_identical"],
            "rules": fingerprint,
            "templates_present": {name: path.exists() for name, path in templates.items()},
            "note": "ERP 只调用 convert_order_file，不改 transform / pipeline / rules.yaml。",
        }
    except Exception as exc:
        return {
            "project": "project1",
            "engine": "order_convert",
            "ready": False,
            "business_rules_modified_by_erp": False,
            "error": f"{type(exc).__name__}: {exc}",
        }


def convert(input_path: Path, out_dir: Path) -> dict[str, Any]:
    """单文件转换，返回状态、统计、告警摘要与三表路径（绝对路径字符串）。"""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    import_order_convert()
    from order_convert import convert_order_file

    result = convert_order_file(Path(input_path), out_dir)
    dx = str(result.dx_path.resolve()) if result.dx_path else None
    d01 = str(result.d01_path.resolve()) if result.d01_path else None
    d02 = str(result.d02_path.resolve()) if result.d02_path else None
    errors = _issue_brief(list(result.errors))
    warnings = _issue_brief(list(result.warnings))
    anomalies = _issue_brief(list(result.anomalies))
    file_name = Path(input_path).name
    issues = collect_visible_issues(
        [
            {
                "file": file_name,
                "errors": errors,
                "warnings": warnings,
            }
        ]
    )
    notes_path = _write_exception_notes(out_dir, issues)
    notes = str(notes_path.resolve()) if notes_path else None
    return {
        "file": file_name,
        "export_status": result.export_status,
        "stats": result.stats,
        "outputs": {"dx": dx, "d01": d01, "d02": d02, "notes": notes},
        "notes_path": notes,
        "error_count": len(result.errors),
        "warning_count": len(result.warnings),
        "errors": errors,
        "warnings": warnings,
        "anomalies": anomalies,
        "anomaly_count": len(result.anomalies),
        "mapping_count": len(result.mappings),
        "formal_export_blocked": bool(result.formal_export_blocked),
        "issues": issues,
        "exception_count": len(issues),
    }


def convert_files(paths: list[Path], out_dir: Path) -> dict[str, Any]:
    """逐个文件调用引擎。单个失败不影响其他文件。"""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    files: list[dict[str, Any]] = []
    for path in paths:
        target = out_dir / Path(path).stem
        try:
            files.append(convert(Path(path), target))
        except Exception as exc:
            files.append(
                {
                    "file": Path(path).name,
                    "export_status": "error",
                    "message": f"{type(exc).__name__}: {exc}",
                    "outputs": {"dx": None, "d01": None, "d02": None, "notes": None},
                    "error_count": 1,
                    "warning_count": 0,
                    "errors": [
                        {
                            "type": "CONVERT_FAILED",
                            "severity": "error",
                            "订单号": "",
                            "message": f"{type(exc).__name__}: {exc}",
                        }
                    ],
                    "warnings": [],
                    "anomalies": [],
                }
            )
    all_issues = collect_visible_issues(files)
    return {
        "engine": "order_convert",
        "files": files,
        "issues": all_issues,
        "exception_count": len(all_issues),
    }
