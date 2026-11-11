"""把上传文件交给项目1封装层（再调 vendor order_convert）。

不要在这里重写 transform、pipeline 或 rules.yaml。
"""
from __future__ import annotations

import sys
from pathlib import Path

from app.config import get_settings


def _ensure_project1_engine() -> None:
    """把 engines/project1 加入 path，以便 import project1_engine。"""
    root = str(get_settings().erp_repo_root / "engines" / "project1")
    if root not in sys.path:
        sys.path.insert(0, root)


def _engine():
    _ensure_project1_engine()
    import project1_engine

    return project1_engine


def import_order_convert():
    """兼容旧调用：仍返回 vendor order_convert 包。"""
    return _engine().import_order_convert()


def project1_status() -> dict:
    return _engine().status()


def convert_files(paths: list[Path], out_dir: Path) -> dict:
    """逐个文件调用封装层，并把输出路径改成相对 data_dir。"""
    summary = _engine().convert_files(list(paths), Path(out_dir))
    for item in summary.get("files", []):
        outputs = item.get("outputs") or {}
        item["outputs"] = {
            "dx": _rel_str(outputs.get("dx")),
            "d01": _rel_str(outputs.get("d01")),
            "d02": _rel_str(outputs.get("d02")),
        }
    return summary


def aggregate_status(summary: dict) -> str:
    statuses = [item.get("export_status") for item in summary.get("files", [])]
    if not statuses or any(item == "error" for item in statuses):
        return "failed"
    if any(item == "blocked" for item in statuses):
        return "blocked"
    if any(item == "needs_review" for item in statuses):
        return "needs_review"
    if all(item == "ok" for item in statuses):
        return "converted"
    return "failed"


def _rel_str(path: str | Path | None) -> str | None:
    if path is None:
        return None
    data_dir = get_settings().data_dir.resolve()
    resolved = Path(path).resolve()
    try:
        return resolved.relative_to(data_dir).as_posix()
    except ValueError:
        return str(resolved)


def build_downloads(summary: dict) -> list[dict]:
    """从 convert 摘要生成下载列表（店小秘 / 店管家01 / 店管家02）。"""
    downloads: list[dict] = []
    labels = (("dx", "店小秘"), ("d01", "店管家01"), ("d02", "店管家02"))
    for item in summary.get("files", []):
        outputs = item.get("outputs") or {}
        for key, label in labels:
            path = outputs.get(key)
            if path:
                downloads.append(
                    {
                        "label": label,
                        "path": path,
                        "kind": key,
                        "source_file": item.get("file"),
                    }
                )
    return downloads
