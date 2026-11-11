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


def _issue_brief(items: list) -> list[dict[str, Any]]:
    """把 anomalies / errors / warnings 压成前端可展示的短摘要。"""
    out: list[dict[str, Any]] = []
    for item in items[:30]:
        if isinstance(item, dict):
            out.append(
                {
                    "type": item.get("type") or item.get("code") or "",
                    "severity": item.get("severity") or "",
                    "订单号": item.get("订单号") or item.get("order_no") or "",
                    "message": item.get("action")
                    or item.get("message")
                    or item.get("msg")
                    or "",
                }
            )
        else:
            out.append({"message": str(item)})
    return out


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
    return {
        "file": Path(input_path).name,
        "export_status": result.export_status,
        "stats": result.stats,
        "outputs": {"dx": dx, "d01": d01, "d02": d02},
        "error_count": len(result.errors),
        "warning_count": len(result.warnings),
        "errors": _issue_brief(list(result.errors)),
        "warnings": _issue_brief(list(result.warnings)),
        "anomalies": _issue_brief(list(result.anomalies)),
        "anomaly_count": len(result.anomalies),
        "mapping_count": len(result.mappings),
        "formal_export_blocked": bool(result.formal_export_blocked),
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
                    "outputs": {"dx": None, "d01": None, "d02": None},
                    "error_count": 1,
                    "warning_count": 0,
                }
            )
    return {"engine": "order_convert", "files": files}
