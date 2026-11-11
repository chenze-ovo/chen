"""项目2结果打包：有人工项才写「加工说明.txt」；可打 zip。"""
from __future__ import annotations

import io
import zipfile
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

_SHANGHAI = ZoneInfo("Asia/Shanghai")

_KIND_FOLDER = {
    "oversea": "海外仓",
    "oversea_copy": "海外仓",
    "ops_group": "运营",
    "ops": "运营",
    "dian_01": "店管家",
    "dian_02": "店管家",
    "dian_01_custom": "店管家",
    "dian_02_custom": "店管家",
}

_NOTES_NAME = "加工说明.txt"
_SKIP_ZIP_NAMES = {"_internal", "_summary.json", "_result.zip", "__pycache__"}

# 只写进加工说明的「要动手」类型；其余（规则、统计、副本、隐藏行等）一律不写
_ACTIONABLE_TYPES = {
    "D02_WAREHOUSE_UNMAPPED",
    "D02_AMOUNT_LEFT_EMPTY",
    "CHINA_DIRECT_AMOUNT_UNMATCHED",
    "UNKNOWN_SKU_SUFFIX",
    "EBAY_SPLIT_SKIP_D01",
    "SKU_STRIP_SUSPECT",
    "D02_TRACKING_EMPTY_PACKAGE_ONLY",
}


def notes_filename() -> str:
    return _NOTES_NAME


def folder_for_kind(kind: str | None, label: str | None = None) -> str:
    if kind and kind in _KIND_FOLDER:
        return _KIND_FOLDER[kind]
    text = f"{kind or ''} {label or ''}"
    if "副本" in text or "海外仓" in text:
        return "海外仓"
    if "运营" in text:
        return "运营"
    if "店管家" in text:
        return "店管家"
    return "其他"


def collect_actionable_note_lines(summary: dict) -> list[str]:
    """只收集需要人工动手的短行；无则返回空列表。"""
    anomalies = list(summary.get("anomalies") or [])
    downloads = list(summary.get("downloads") or [])
    file_by_kind = {
        str(d.get("kind") or ""): Path(str(d.get("path") or "")).name
        for d in downloads
        if d.get("path") and d.get("kind") not in {None, "notes"}
    }
    default_d02 = file_by_kind.get("dian_02") or "店管家02.xls"
    default_d01 = file_by_kind.get("dian_01") or "店管家01.xls"

    # 中国直发：若已有 D02_AMOUNT 的 china 条，跳过上游 CHINA_DIRECT，避免重复
    has_d02_china = any(
        a.get("type") == "D02_AMOUNT_LEFT_EMPTY" and a.get("reason") == "china_direct_unmatched"
        for a in anomalies
    )

    lines: list[str] = []
    seen: set[str] = set()
    for a in anomalies:
        typ = str(a.get("type") or "")
        if typ not in _ACTIONABLE_TYPES:
            continue
        if typ == "CHINA_DIRECT_AMOUNT_UNMATCHED" and has_d02_china:
            continue

        order_no = str(a.get("order_no") or "").strip()
        file_name = str(a.get("file") or a.get("file_name") or "").strip()
        if not file_name:
            if typ.startswith("D02_") or typ == "CHINA_DIRECT_AMOUNT_UNMATCHED":
                file_name = default_d02
            elif typ == "UNKNOWN_SKU_SUFFIX":
                dest = str(a.get("destination") or "")
                file_name = default_d01 if dest == "dian_01" else default_d02
            elif typ == "EBAY_SPLIT_SKIP_D01":
                file_name = default_d01
            elif typ == "SKU_STRIP_SUSPECT":
                file_name = "海外仓订单"
            else:
                file_name = default_d02

        row = a.get("excel_row")
        if row is None:
            row = a.get("row")
        try:
            row_n = int(row) if row is not None and str(row).strip() != "" else None
        except (TypeError, ValueError):
            row_n = None

        action = _action_phrase(a)
        bits = [file_name]
        if row_n is not None:
            bits.append(f"第{row_n}行")
        if order_no:
            bits.append(f"订单{order_no}")
        bits.append(action)
        line = " ".join(bits)
        if line in seen:
            continue
        seen.add(line)
        lines.append(line)
    return lines


def _action_phrase(a: dict) -> str:
    typ = str(a.get("type") or "")
    reason = str(a.get("reason") or "")
    if typ == "D02_WAREHOUSE_UNMAPPED":
        return "物流方式未映射"
    if typ == "D02_AMOUNT_LEFT_EMPTY" and reason == "split_unfilled":
        return "拆单需补金额"
    if typ == "D02_AMOUNT_LEFT_EMPTY" and reason == "china_direct_unmatched":
        return "需补金额（中国直发）"
    if typ == "CHINA_DIRECT_AMOUNT_UNMATCHED":
        return "需补金额（中国直发）"
    if typ == "D02_AMOUNT_LEFT_EMPTY":
        return "需补金额"
    if typ == "UNKNOWN_SKU_SUFFIX":
        return "未知SKU需修正"
    if typ == "EBAY_SPLIT_SKIP_D01":
        return "eBay分单不进店管家01"
    if typ == "SKU_STRIP_SUSPECT":
        return "SKU疑似剥错需确认"
    if typ == "D02_TRACKING_EMPTY_PACKAGE_ONLY":
        return "缺运单号（只有包裹号）"
    return "需人工处理"


def build_processing_notes_text(summary: dict, *, batch_id: int | None = None) -> str:
    """有人工项才返回短文本；否则空字符串。batch_id 保留兼容，不写入正文。"""
    _ = batch_id
    lines = collect_actionable_note_lines(summary)
    if not lines:
        return ""
    return "\n".join(lines) + "\n"


def write_processing_notes(out_dir: Path, summary: dict, *, batch_id: int | None = None) -> Path | None:
    """有需人工条目才写「加工说明.txt」；否则不写并删掉旧文件。返回路径或 None。"""
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / _NOTES_NAME
    text = build_processing_notes_text(summary, batch_id=batch_id)
    if not text.strip():
        if path.is_file():
            try:
                path.unlink()
            except OSError:
                pass
        return None
    path.write_text(text, encoding="utf-8")
    return path


def resolve_out_dir(data_dir: Path, batch_id: int, summary: dict | None = None) -> Path:
    """推断批次输出目录。"""
    candidate = data_dir / "outputs" / str(batch_id)
    if candidate.is_dir():
        return candidate
    for item in (summary or {}).get("downloads") or []:
        rel = str(item.get("path") or "")
        if not rel:
            continue
        full = (data_dir / rel).resolve()
        if full.is_file():
            return full.parent
    candidate.mkdir(parents=True, exist_ok=True)
    return candidate


def ensure_notes_in_summary(
    out_dir: Path,
    summary: dict,
    data_dir: Path,
    *,
    batch_id: int | None = None,
) -> dict:
    """有人工项才写加工说明并加入 downloads；否则不写、不展示、删旧文件。"""
    notes_path = write_processing_notes(out_dir, summary, batch_id=batch_id)
    downloads = [
        d
        for d in (summary.get("downloads") or [])
        if d.get("kind") != "notes" and Path(str(d.get("path") or "")).name != _NOTES_NAME
    ]
    summary = dict(summary)
    if notes_path is None:
        summary["downloads"] = downloads
        summary.pop("notes_path", None)
        return summary
    try:
        relative = notes_path.resolve().relative_to(data_dir.resolve()).as_posix()
    except ValueError:
        relative = str(notes_path)
    downloads.append({"label": "加工说明", "path": relative, "kind": "notes"})
    summary["downloads"] = downloads
    summary["notes_path"] = relative
    return summary


def build_result_zip_bytes(
    out_dir: Path,
    summary: dict,
    *,
    batch_id: int | None = None,
) -> tuple[bytes, str]:
    """打成 zip 字节；仅当有人工项时才含加工说明.txt。"""
    write_processing_notes(out_dir, summary, batch_id=batch_id)

    buf = io.BytesIO()
    used_arcnames: set[str] = set()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        notes = out_dir / _NOTES_NAME
        if notes.is_file():
            zf.write(notes, arcname=_NOTES_NAME)
            used_arcnames.add(_NOTES_NAME)

        mapped_files: set[Path] = set()
        for item in summary.get("downloads") or []:
            if item.get("kind") == "notes":
                continue
            rel_or_name = str(item.get("path") or "")
            if not rel_or_name:
                continue
            name = Path(rel_or_name).name
            full = out_dir / name
            if not full.is_file():
                alt = Path(rel_or_name)
                if alt.is_file():
                    full = alt
                else:
                    continue
            if full.name == _NOTES_NAME:
                continue
            folder = folder_for_kind(item.get("kind"), item.get("label"))
            arc = f"{folder}/{full.name}"
            arc = _unique_arc(arc, used_arcnames)
            zf.write(full, arcname=arc)
            mapped_files.add(full.resolve())

        for path in sorted(out_dir.iterdir()):
            if path.name in _SKIP_ZIP_NAMES or path.name.startswith("."):
                continue
            if path.is_dir():
                if path.name == "_internal":
                    continue
                for child in sorted(path.rglob("*")):
                    if not child.is_file():
                        continue
                    if child.name in _SKIP_ZIP_NAMES:
                        continue
                    if child.resolve() in mapped_files:
                        continue
                    arc = child.relative_to(out_dir).as_posix()
                    if arc == _NOTES_NAME or arc.endswith(f"/{_NOTES_NAME}"):
                        if _NOTES_NAME in used_arcnames:
                            continue
                    arc = _unique_arc(arc, used_arcnames)
                    zf.write(child, arcname=arc)
                    mapped_files.add(child.resolve())
                continue
            if path.resolve() in mapped_files:
                continue
            if path.name == _NOTES_NAME:
                continue
            folder = _guess_folder_from_name(path.name)
            arc = f"{folder}/{path.name}" if folder else path.name
            arc = _unique_arc(arc, used_arcnames)
            zf.write(path, arcname=arc)

    stamp = datetime.now(_SHANGHAI).strftime("%Y%m%d")
    if batch_id is not None:
        filename = f"海外仓加工结果-批次{batch_id}-{stamp}.zip"
    else:
        filename = f"海外仓加工结果-{stamp}.zip"
    return buf.getvalue(), filename


def _unique_arc(arc: str, used: set[str]) -> str:
    if arc not in used:
        used.add(arc)
        return arc
    stem = Path(arc)
    parent = stem.parent.as_posix() if str(stem.parent) not in {".", ""} else ""
    base, suf = stem.stem, stem.suffix
    n = 2
    while True:
        name = f"{base}-{n}{suf}"
        candidate = f"{parent}/{name}" if parent else name
        if candidate not in used:
            used.add(candidate)
            return candidate
        n += 1


def _guess_folder_from_name(name: str) -> str:
    if "海外仓" in name:
        return "海外仓"
    if name.startswith("店管家") or "店管家" in name:
        return "店管家"
    if name.endswith(".xlsx") or name.endswith(".xls"):
        if "单号" in name or "运营" in name:
            return "运营"
    return "其他"
