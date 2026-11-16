"""中国直发下载结果：只给用户要导入的表和异常说明，并用中文文件名。"""
from __future__ import annotations

import io
import shutil
import zipfile
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

_SHANGHAI = ZoneInfo("Asia/Shanghai")

_INTERNAL_NAMES = {
    "anomalies.json",
    "stats.json",
    "row_mapping.json",
    "row_mapping.csv",
}

_PUBLIC_NAMES = {
    "dianxiaomi.xlsx": "店小秘.xlsx",
    "dianxiaomi_NEEDS_REVIEW.xlsx": "店小秘-待核对.xlsx",
    "dian_01.xls": "店管家01.xls",
    "dian_02.xls": "店管家02.xls",
    "dian_01_NEEDS_REVIEW.xls": "店管家01-待核对.xls",
    "dian_02_NEEDS_REVIEW.xls": "店管家02-待核对.xls",
    "异常说明.txt": "异常说明.txt",
}


def is_internal_result_file(path: Path) -> bool:
    name = path.name
    if name.startswith(".") or name.startswith("_"):
        return True
    if name in _INTERNAL_NAMES:
        return True
    if name.lower().endswith((".json", ".csv")):
        return True
    return False


def public_result_name(path: Path) -> str:
    return _PUBLIC_NAMES.get(path.name, path.name)


def iter_user_result_files(out_dir: Path) -> list[Path]:
    if not out_dir.is_dir():
        return []
    found: list[Path] = []
    for path in sorted(out_dir.rglob("*")):
        if not path.is_file() or is_internal_result_file(path):
            continue
        if "_internal" in path.parts or path.name.startswith("_"):
            continue
        if path.suffix.lower() in {".xlsx", ".xls", ".txt"}:
            found.append(path)
    return found


def build_project1_zip_bytes(out_dir: Path, *, batch_id: int) -> tuple[bytes, str]:
    buf = io.BytesIO()
    used: set[str] = set()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in iter_user_result_files(out_dir):
            name = public_result_name(path)
            if name in used:
                stem = Path(name)
                n = 2
                while True:
                    candidate = f"{stem.stem}-{n}{stem.suffix}"
                    if candidate not in used:
                        name = candidate
                        break
                    n += 1
            used.add(name)
            zf.write(path, arcname=name)
    stamp = datetime.now(_SHANGHAI).strftime("%Y%m%d")
    filename = f"中国直发加工结果-批次{batch_id}-{stamp}.zip"
    return buf.getvalue(), filename


def copy_project1_user_files(src: Path, dest: Path) -> None:
    dest.mkdir(parents=True, exist_ok=True)
    used: set[str] = set()
    for path in iter_user_result_files(src):
        name = public_result_name(path)
        if name in used:
            stem = Path(name)
            n = 2
            while True:
                candidate = f"{stem.stem}-{n}{stem.suffix}"
                if candidate not in used:
                    name = candidate
                    break
                n += 1
        used.add(name)
        shutil.copy2(path, dest / name)
