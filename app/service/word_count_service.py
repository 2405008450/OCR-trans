# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import os
import re
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Awaitable, Callable, Iterable, Optional
from zipfile import ZipFile

from lxml import etree
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from app.core.config import settings
from app.service.libreoffice_service import (
    convert_doc_to_docx_via_libreoffice,
    convert_presentation_to_pptx_via_libreoffice,
    convert_spreadsheet_to_xlsx_via_libreoffice,
)

ProgressCallback = Callable[[int, str], Awaitable[None]]

WORD_COUNT_TASK_TYPE = "word_count"
COUNTABLE_EXTENSIONS = {".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt", ".pdf", ".txt"}
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".tif", ".tiff"}
CAD_EXTENSIONS = {".dwg", ".dxf", ".dwt", ".dws"}
CANDIDATE_EXTENSIONS = COUNTABLE_EXTENSIONS | IMAGE_EXTENSIONS | CAD_EXTENSIONS
DEFAULT_SCAN_EXTENSIONS = sorted(CANDIDATE_EXTENSIONS)

STATUS_COUNTED = "counted"
STATUS_NEEDS_OCR = "needs_ocr"
STATUS_NEEDS_CAD_PARSER = "needs_cad_parser"
STATUS_SKIPPED = "skipped"
STATUS_FAILED = "failed"

EXTRA_SOURCE_TYPES = {"header", "footer", "footnote", "endnote", "chart", "comment"}

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "v": "urn:schemas-microsoft-com:vml",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
REL_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"


@dataclass(frozen=True)
class TextMetrics:
    word_count: int
    non_space_chars: int
    raw_chars: int


@dataclass(frozen=True)
class TextItem:
    source_type: str
    source_label: str
    text: str
    is_extra: bool = False
    paragraph_count: int = 1
    line_count: int = 1
    image_count: int = 0


@dataclass(frozen=True)
class ExtractedContent:
    items: list[TextItem]
    page_count: int = 0
    paragraph_count: int = 0
    line_count: int = 0
    image_count: int = 0
    stat_method: str = ""
    file_type: str = ""


def normalize_scan_extensions(extensions: Optional[Iterable[str]]) -> list[str]:
    if not extensions:
        return DEFAULT_SCAN_EXTENSIONS
    normalized: set[str] = set()
    for item in extensions:
        value = str(item or "").strip().lower()
        if not value:
            continue
        if not value.startswith("."):
            value = f".{value}"
        if value in CANDIDATE_EXTENSIONS:
            normalized.add(value)
    return sorted(normalized or CANDIDATE_EXTENSIONS)


def get_word_count_config() -> dict[str, Any]:
    allowed_roots = []
    for raw_root in settings.WORD_COUNT_ALLOWED_ROOTS:
        root = Path(raw_root).expanduser()
        mapped_root = _map_unc_path_to_local(str(raw_root))
        mapped_path = mapped_root[0] if mapped_root is not None else None
        scope_only = _is_unc_server_root(root) and not root.exists()
        exists = mapped_path.exists() if mapped_path is not None else root.exists() or scope_only
        allowed_roots.append(
            {
                "path": str(root),
                "exists": exists,
                "scope_only": scope_only,
                "mount_path": str(mapped_path) if mapped_path is not None else "",
            }
        )
    unc_mount_mappings = [
        {
            "unc": _normalize_unc_text(unc_path) or unc_path,
            "mount": mount_path,
        }
        for unc_path, mount_path in settings.WORD_COUNT_UNC_MOUNT_MAP.items()
    ]
    if settings.WORD_COUNT_ALLOW_LOCAL_PATHS_ENABLED:
        allowed_roots.append(
            {
                "path": "本地路径（测试放开：所有本机盘符和挂载路径）",
                "exists": True,
            }
        )
    return {
        "allowed_roots": allowed_roots,
        "countable_extensions": sorted(COUNTABLE_EXTENSIONS),
        "image_extensions": sorted(IMAGE_EXTENSIONS),
        "cad_extensions": sorted(CAD_EXTENSIONS),
        "default_extensions": DEFAULT_SCAN_EXTENSIONS,
        "max_files": settings.WORD_COUNT_MAX_FILES,
        "max_file_mb": settings.WORD_COUNT_MAX_FILE_MB,
        "unc_mount_mappings": unc_mount_mappings,
        "unc_auto_mount_roots": settings.WORD_COUNT_UNC_AUTO_MOUNT_ROOTS,
        "follow_symlinks": settings.WORD_COUNT_FOLLOW_SYMLINKS_ENABLED,
        "allow_local_paths": settings.WORD_COUNT_ALLOW_LOCAL_PATHS_ENABLED,
        "count_policy": "Word 近似口径：中日韩字符逐字计数，拉丁/数字按连续词计数，标点和空白不计。",
    }


def prepare_word_count_request(
    *,
    directory_path: str,
    recursive: bool = True,
    include_hidden: bool = False,
    extensions: Optional[Iterable[str]] = None,
) -> dict[str, Any]:
    resolved_dir, matched_root = _resolve_allowed_directory(directory_path)
    normalized_extensions = normalize_scan_extensions(extensions)
    params = {
        "directory_path": str(resolved_dir),
        "recursive": bool(recursive),
        "include_hidden": bool(include_hidden),
        "extensions": normalized_extensions,
        "max_files": settings.WORD_COUNT_MAX_FILES,
        "max_file_mb": settings.WORD_COUNT_MAX_FILE_MB,
        "follow_symlinks": settings.WORD_COUNT_FOLLOW_SYMLINKS_ENABLED,
    }
    return {
        "filename": str(resolved_dir),
        "params": params,
        "input_files": {
            "directory_path": str(resolved_dir),
            "allowed_root": str(matched_root),
        },
    }


async def execute_word_count_task(
    *,
    task_id: str,
    display_no: Optional[str],
    directory_path: str,
    recursive: bool,
    include_hidden: bool,
    extensions: Iterable[str],
    progress_callback: Optional[ProgressCallback] = None,
    executor=None,
) -> dict[str, Any]:
    import asyncio

    loop = asyncio.get_running_loop()

    def sync_report(progress: int, message: str) -> None:
        if not progress_callback:
            return
        future = asyncio.run_coroutine_threadsafe(progress_callback(progress, message), loop)
        future.result(timeout=30)

    return await loop.run_in_executor(
        executor,
        lambda: run_word_count_task_sync(
            task_id=task_id,
            display_no=display_no,
            directory_path=directory_path,
            recursive=recursive,
            include_hidden=include_hidden,
            extensions=extensions,
            progress_callback=sync_report,
        ),
    )


def run_word_count_task_sync(
    *,
    task_id: str,
    display_no: Optional[str],
    directory_path: str,
    recursive: bool,
    include_hidden: bool,
    extensions: Iterable[str],
    progress_callback: Optional[Callable[[int, str], None]] = None,
) -> dict[str, Any]:
    started_at = datetime.now()
    root, matched_root = _resolve_allowed_directory(directory_path)
    scan_extensions = set(normalize_scan_extensions(extensions))
    output_dir = Path(settings.OUTPUT_DIR) / "word_count" / (display_no or task_id)
    output_dir.mkdir(parents=True, exist_ok=True)
    converted_dir = output_dir / "converted_inputs"
    converted_dir.mkdir(parents=True, exist_ok=True)

    _report(progress_callback, 5, "正在扫描目录...")
    candidates = list(_iter_candidate_files(root, recursive, include_hidden, scan_extensions))
    max_files = max(1, int(settings.WORD_COUNT_MAX_FILES or 5000))
    truncated = len(candidates) > max_files
    if truncated:
        candidates = candidates[:max_files]

    file_results: list[dict[str, Any]] = []
    source_rows: list[dict[str, Any]] = []
    total = len(candidates)
    max_bytes = max(1, int(settings.WORD_COUNT_MAX_FILE_MB or 200)) * 1024 * 1024

    if total == 0:
        _report(progress_callback, 80, "未发现可统计的候选文件，正在生成空报告...")

    for index, file_path in enumerate(candidates, start=1):
        progress = 8 + int(index / max(total, 1) * 72)
        _report(progress_callback, min(progress, 80), f"正在统计 {index}/{total}: {file_path.name}")
        result, rows = _count_single_file(
            file_path=file_path,
            root=root,
            converted_dir=converted_dir,
            max_bytes=max_bytes,
        )
        file_results.append(result)
        source_rows.extend(rows)

    summary = _build_summary(file_results, truncated=truncated, started_at=started_at)
    report_payload = {
        "task_id": task_id,
        "directory_path": str(root),
        "allowed_root": str(matched_root),
        "recursive": bool(recursive),
        "include_hidden": bool(include_hidden),
        "extensions": sorted(scan_extensions),
        "summary": summary,
        "files": file_results,
        "source_details": source_rows,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "rules": [
            "主统计采用 Word 近似口径，不保证与 Word COM 组件 100% 一致。",
            "Word 主数计入正文、正文表格和正文文本框；页眉页脚、脚注尾注、图表文字、批注列为额外内容。",
            "页数和行数采用快速近似口径：Word 优先读取文档属性，PDF/PPT 使用实际页数或幻灯片数，Excel 页数按工作表数近似。",
            "图片数量统计嵌入图片出现次数，用于后续 OCR 流程线索；含图片的可编辑 Office 文件仍保持已统计。",
            "PPT 主数统计幻灯片可见文本框、表格、组合形状和可提取图表文字；备注文本列为额外内容。",
            "PDF 首版只统计可提取文本；无可提取文本的 PDF 标记为需要 OCR。",
            "图片和 CAD 首版仅标记为扩展处理入口，不计入主数。",
        ],
    }

    _report(progress_callback, 86, "正在生成 JSON 报告...")
    json_path = output_dir / "字数统计结果.json"
    json_path.write_text(json.dumps(report_payload, ensure_ascii=False, indent=2), encoding="utf-8")

    _report(progress_callback, 92, "正在生成 Excel 报告...")
    excel_path = output_dir / "字数统计报告.xlsx"
    _write_excel_report(excel_path, report_payload)

    _report(progress_callback, 98, "正在整理输出结果...")
    report_payload["report_json"] = _output_web_path(json_path)
    report_payload["report_excel"] = _output_web_path(excel_path)
    report_payload["summary_text"] = (
        f"已统计 {summary['counted_files']} 个文件，主字数 {summary['total_main_word_count']}，"
        f"额外内容字数 {summary['total_extra_word_count']}。"
    )
    return report_payload


def count_words_word_like(text: str) -> int:
    return _count_text(text).word_count


def _report(callback: Optional[Callable[[int, str], None]], progress: int, message: str) -> None:
    if callback:
        callback(progress, message)


def _resolve_allowed_directory(raw_path: str) -> tuple[Path, Path]:
    raw_text = str(raw_path or "").strip()
    if not raw_text:
        raise ValueError("目录路径不能为空")

    mapped = _map_unc_path_to_local(raw_text)
    if mapped is not None:
        candidate, matched_root, matched_unc = mapped
        candidate = candidate.expanduser().resolve(strict=False)
        matched_root = matched_root.expanduser().resolve(strict=False)
        if not candidate.exists():
            raise FileNotFoundError(
                f"目录不存在: {candidate}（由 {matched_unc} 映射而来；请确认 Docker 已挂载局域网共享目录）"
            )
        if not candidate.is_dir():
            raise ValueError(f"路径不是目录: {candidate}（由 {matched_unc} 映射而来）")
        allowed_root = _match_allowed_root_for_mapped_unc(raw_text, candidate, matched_root)
        if allowed_root is None:
            allowed_text = "；".join(settings.WORD_COUNT_ALLOWED_ROOTS)
            raise PermissionError(f"目录不在允许扫描范围内。允许根目录: {allowed_text}")
        return candidate, allowed_root

    if _is_unc_text(raw_text) and os.name != "nt":
        candidates = _format_unc_auto_candidates(raw_text)
        suffix = f" 自动尝试过: {candidates}" if candidates else ""
        raise FileNotFoundError(
            "Docker/Linux 容器不能直接读取 Windows UNC 路径；/mnt 是容器内挂载点，不是 Windows Server 路径。"
            "请在 docker-compose.yml 中把 Windows Server 共享目录挂载到 /mnt/win-server/服务器资料7，"
            "或配置 WORD_COUNT_UNC_MOUNT_MAP_JSON 将 UNC 映射到实际容器内路径。"
            f"{suffix}"
        )

    candidate = Path(raw_text).expanduser().resolve(strict=False)
    if not candidate.exists():
        raise FileNotFoundError(f"目录不存在: {candidate}")
    if not candidate.is_dir():
        raise ValueError(f"路径不是目录: {candidate}")

    if settings.WORD_COUNT_ALLOW_LOCAL_PATHS_ENABLED and _is_local_absolute_path(candidate):
        return candidate, _local_path_root(candidate)

    for raw_root in settings.WORD_COUNT_ALLOWED_ROOTS:
        allowed_root = Path(raw_root).expanduser().resolve(strict=False)
        if _is_relative_to_path(candidate, allowed_root):
            return candidate, allowed_root
    allowed_text = "；".join(settings.WORD_COUNT_ALLOWED_ROOTS)
    raise PermissionError(f"目录不在允许扫描范围内。允许根目录: {allowed_text}")


def _map_unc_path_to_local(raw_path: str) -> Optional[tuple[Path, Path, str]]:
    raw_unc = _normalize_unc_text(raw_path)
    if not raw_unc:
        return None
    mappings = []
    for unc_path, mount_path in settings.WORD_COUNT_UNC_MOUNT_MAP.items():
        normalized = _normalize_unc_text(unc_path)
        if normalized:
            mappings.append((normalized, str(mount_path)))
    mappings.sort(key=lambda item: len(_unc_parts_from_text(item[0])), reverse=True)
    for unc_prefix, mount_path in mappings:
        remainder = _unc_relative_parts(raw_unc, unc_prefix)
        if remainder is None:
            continue
        mount_root = Path(mount_path)
        return mount_root.joinpath(*remainder), mount_root, unc_prefix
    for unc_prefix, mount_root, candidate in _iter_unc_auto_mapping_candidates(raw_unc):
        if mount_root.exists() or candidate.exists():
            return candidate, mount_root, unc_prefix
    return None


def _match_allowed_root_for_mapped_unc(raw_unc_path: str, candidate: Path, fallback_mapped_root: Optional[Path] = None) -> Optional[Path]:
    raw_unc = _normalize_unc_text(raw_unc_path)
    for raw_root in settings.WORD_COUNT_ALLOWED_ROOTS:
        root_unc = _normalize_unc_text(raw_root)
        if root_unc and raw_unc and _unc_relative_parts(raw_unc, root_unc) is not None:
            mapped_root = _map_unc_path_to_local(root_unc)
            if mapped_root is not None:
                return mapped_root[0].expanduser().resolve(strict=False)
            if fallback_mapped_root is not None:
                return fallback_mapped_root.expanduser().resolve(strict=False)
            return Path(raw_root)
        if not root_unc:
            allowed_root = Path(raw_root).expanduser().resolve(strict=False)
            if _is_relative_to_path(candidate, allowed_root):
                return allowed_root
    return None


def _iter_unc_auto_mapping_candidates(raw_unc: str) -> Iterable[tuple[str, Path, Path]]:
    parts = _unc_parts_from_text(raw_unc)
    if len(parts) < 2:
        return
    server, share = parts[0], parts[1]
    remainder = parts[2:]
    share_unc = _normalize_unc_text(f"\\\\{server}\\{share}\\")
    server_unc = _normalize_unc_text(f"\\\\{server}\\")
    for root_text in settings.WORD_COUNT_UNC_AUTO_MOUNT_ROOTS:
        base = Path(root_text).expanduser()
        share_mounts = [
            base / server / share,
            base / share,
        ]
        seen: set[str] = set()
        for mount_root in share_mounts:
            key = os.path.normcase(str(mount_root))
            if key in seen:
                continue
            seen.add(key)
            yield share_unc, mount_root, mount_root.joinpath(*remainder)
        server_mount = base / server
        server_key = os.path.normcase(str(server_mount))
        if server_key not in seen:
            yield server_unc, server_mount, server_mount.joinpath(share, *remainder)


def _format_unc_auto_candidates(raw_path: str, limit: int = 8) -> str:
    raw_unc = _normalize_unc_text(raw_path)
    if not raw_unc:
        return ""
    candidates: list[str] = []
    seen: set[str] = set()
    for _, _, candidate in _iter_unc_auto_mapping_candidates(raw_unc):
        text = str(candidate)
        key = os.path.normcase(text)
        if key in seen:
            continue
        seen.add(key)
        candidates.append(text)
        if len(candidates) >= limit:
            break
    return "；".join(candidates)


def _is_relative_to_path(candidate: Path, root: Path) -> bool:
    if _is_unc_server_root(root):
        return _unc_server_name(candidate) == _unc_server_name(root)

    try:
        candidate.resolve(strict=False).relative_to(root.resolve(strict=False))
        return True
    except ValueError:
        pass

    try:
        candidate_text = os.path.normcase(str(candidate.resolve(strict=False)))
        root_text = os.path.normcase(str(root.resolve(strict=False)))
        return os.path.commonpath([candidate_text, root_text]) == root_text
    except (ValueError, OSError):
        return False


def _is_local_absolute_path(path: Path) -> bool:
    return path.is_absolute() and not _is_unc_path(path)


def _local_path_root(path: Path) -> Path:
    anchor = path.anchor
    return Path(anchor) if anchor else path


def _is_unc_path(path: Path) -> bool:
    return _is_unc_text(str(path))


def _is_unc_server_root(path: Path) -> bool:
    return len(_unc_parts(path)) == 1


def _unc_server_name(path: Path) -> Optional[str]:
    parts = _unc_parts(path)
    return parts[0].lower() if parts else None


def _unc_parts(path: Path) -> list[str]:
    return _unc_parts_from_text(str(path))


def _is_unc_text(value: str) -> bool:
    return str(value or "").strip().replace("/", "\\").startswith("\\\\")


def _normalize_unc_text(value: str) -> str:
    text = str(value or "").strip().replace("/", "\\")
    if not text.startswith("\\\\"):
        return ""
    parts = _unc_parts_from_text(text)
    if not parts:
        return ""
    return "\\\\" + "\\".join(parts) + "\\"


def _unc_parts_from_text(value: str) -> list[str]:
    text = str(value or "").strip().replace("/", "\\")
    if not text.startswith("\\\\"):
        return []
    return [part for part in text.strip("\\").split("\\") if part]


def _unc_relative_parts(candidate_unc: str, root_unc: str) -> Optional[list[str]]:
    candidate_parts = _unc_parts_from_text(candidate_unc)
    root_parts = _unc_parts_from_text(root_unc)
    if not candidate_parts or not root_parts or len(candidate_parts) < len(root_parts):
        return None
    candidate_head = [part.lower() for part in candidate_parts[: len(root_parts)]]
    root_head = [part.lower() for part in root_parts]
    if candidate_head != root_head:
        return None
    return candidate_parts[len(root_parts):]


def _iter_candidate_files(
    root: Path,
    recursive: bool,
    include_hidden: bool,
    extensions: set[str],
) -> Iterable[Path]:
    followlinks = settings.WORD_COUNT_FOLLOW_SYMLINKS_ENABLED

    if not recursive:
        for child in sorted(root.iterdir(), key=lambda item: item.name.lower()):
            if not include_hidden and _is_hidden_name(child.name):
                continue
            if child.is_file() and child.suffix.lower() in extensions:
                if followlinks or not child.is_symlink():
                    yield child
        return

    for current_dir, dir_names, file_names in os.walk(root, followlinks=followlinks):
        if not include_hidden:
            dir_names[:] = [name for name in dir_names if not _is_hidden_name(name)]
            file_names = [name for name in file_names if not _is_hidden_name(name)]
        dir_names.sort(key=str.lower)
        for file_name in sorted(file_names, key=str.lower):
            file_path = Path(current_dir) / file_name
            if file_path.suffix.lower() not in extensions:
                continue
            if not followlinks and file_path.is_symlink():
                continue
            yield file_path


def _is_hidden_name(name: str) -> bool:
    return name.startswith(".") or name.startswith("~$")


def _count_single_file(
    *,
    file_path: Path,
    root: Path,
    converted_dir: Path,
    max_bytes: int,
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    now_status = STATUS_COUNTED
    ext = file_path.suffix.lower()
    relative_path = _relative_path(file_path, root)
    error = ""
    warning = ""
    source_rows: list[dict[str, Any]] = []

    try:
        stat = file_path.stat()
        size_bytes = int(stat.st_size)
        modified_at = datetime.fromtimestamp(stat.st_mtime).isoformat(timespec="seconds")
    except OSError as exc:
        return (
            _base_file_result(file_path, relative_path, ext, STATUS_FAILED, error=f"无法读取文件信息: {exc}"),
            [],
        )

    base = _base_file_result(
        file_path,
        relative_path,
        ext,
        now_status,
        size_bytes=size_bytes,
        modified_at=modified_at,
    )

    if ext in IMAGE_EXTENSIONS:
        base.update(
            status=STATUS_NEEDS_OCR,
            image_count=1,
            message="图片文件需要 OCR 后再统计",
            stat_method=_stat_method_for_extension(ext),
            counted_at=_now_iso(),
        )
        return base, []
    if ext in CAD_EXTENSIONS:
        base.update(
            status=STATUS_NEEDS_CAD_PARSER,
            message="CAD 文件需要专用解析器后再统计",
            stat_method=_stat_method_for_extension(ext),
            counted_at=_now_iso(),
        )
        return base, []
    if size_bytes > max_bytes:
        base.update(
            status=STATUS_SKIPPED,
            message=f"文件超过大小限制 {settings.WORD_COUNT_MAX_FILE_MB} MB",
            counted_at=_now_iso(),
        )
        return base, []

    try:
        extracted = _extract_content(file_path, converted_dir)
    except Exception as exc:
        base.update(status=STATUS_FAILED, error=str(exc), message="解析失败", counted_at=_now_iso())
        return base, []

    items = extracted.items
    if ext == ".pdf" and not any(item.text.strip() for item in items):
        base.update(
            status=STATUS_NEEDS_OCR,
            page_count=extracted.page_count,
            image_count=extracted.image_count,
            stat_method=extracted.stat_method,
            message="PDF 未提取到可编辑文本，可能是扫描件",
            counted_at=_now_iso(),
        )
        return base, []

    totals = {
        "main_word_count": 0,
        "extra_word_count": 0,
        "main_non_space_chars": 0,
        "main_raw_chars": 0,
        "extra_non_space_chars": 0,
        "extra_raw_chars": 0,
    }
    source_counter: Counter[str] = Counter()
    for item in items:
        metrics = _count_text(item.text)
        if item.is_extra:
            totals["extra_word_count"] += metrics.word_count
            totals["extra_non_space_chars"] += metrics.non_space_chars
            totals["extra_raw_chars"] += metrics.raw_chars
        else:
            totals["main_word_count"] += metrics.word_count
            totals["main_non_space_chars"] += metrics.non_space_chars
            totals["main_raw_chars"] += metrics.raw_chars
        source_counter[item.source_type] += 1
        source_rows.append(
            {
                "file_path": str(file_path),
                "relative_path": relative_path,
                "extension": ext,
                "source_type": item.source_type,
                "source_label": item.source_label,
                "is_extra": item.is_extra,
                "word_count": metrics.word_count,
                "non_space_chars": metrics.non_space_chars,
                "raw_chars": metrics.raw_chars,
                "char_count_no_spaces": metrics.non_space_chars,
                "char_count_with_spaces": metrics.raw_chars,
                "paragraph_count": item.paragraph_count,
                "line_count": item.line_count,
                "image_count": item.image_count,
                "text_preview": _preview_text(item.text),
            }
        )

    if not items:
        warning = "未提取到文本"

    base.update(
        status=STATUS_COUNTED,
        main_word_count=totals["main_word_count"],
        word_count=totals["main_word_count"],
        extra_word_count=totals["extra_word_count"],
        char_count_no_spaces=totals["main_non_space_chars"],
        char_count_with_spaces=totals["main_raw_chars"],
        non_space_chars=totals["main_non_space_chars"],
        raw_chars=totals["main_raw_chars"],
        extra_non_space_chars=totals["extra_non_space_chars"],
        extra_raw_chars=totals["extra_raw_chars"],
        page_count=extracted.page_count,
        paragraph_count=extracted.paragraph_count,
        line_count=extracted.line_count,
        image_count=extracted.image_count,
        stat_method=extracted.stat_method,
        file_type=extracted.file_type or base.get("file_type", ""),
        source_counts=dict(source_counter),
        warning=warning,
        message="统计完成",
        counted_at=_now_iso(),
    )
    return base, source_rows


def _base_file_result(
    file_path: Path,
    relative_path: str,
    extension: str,
    status: str,
    *,
    size_bytes: int = 0,
    modified_at: str = "",
    error: str = "",
) -> dict[str, Any]:
    return {
        "file_path": str(file_path),
        "relative_path": relative_path,
        "filename": file_path.name,
        "extension": extension,
        "file_type": _file_type_for_extension(extension),
        "status": status,
        "message": "",
        "main_word_count": 0,
        "word_count": 0,
        "extra_word_count": 0,
        "char_count_no_spaces": 0,
        "char_count_with_spaces": 0,
        "non_space_chars": 0,
        "raw_chars": 0,
        "extra_non_space_chars": 0,
        "extra_raw_chars": 0,
        "page_count": 0,
        "paragraph_count": 0,
        "line_count": 0,
        "image_count": 0,
        "stat_method": _stat_method_for_extension(extension),
        "counted_at": _now_iso(),
        "size_bytes": size_bytes,
        "modified_at": modified_at,
        "source_counts": {},
        "warning": "",
        "error": error,
    }


def _extract_text_items(file_path: Path, converted_dir: Path) -> list[TextItem]:
    return _extract_content(file_path, converted_dir).items


def _extract_content(file_path: Path, converted_dir: Path) -> ExtractedContent:
    ext = file_path.suffix.lower()
    converted_from = ""
    if ext == ".doc":
        target = converted_dir / f"{_safe_stem(file_path)}.docx"
        file_path = Path(convert_doc_to_docx_via_libreoffice(file_path, target))
        converted_from = "DOC经LibreOffice转换"
        ext = ".docx"
    elif ext == ".xls":
        target = converted_dir / f"{_safe_stem(file_path)}.xlsx"
        file_path = Path(convert_spreadsheet_to_xlsx_via_libreoffice(file_path, target))
        converted_from = "XLS经LibreOffice转换"
        ext = ".xlsx"
    elif ext == ".ppt":
        target = converted_dir / f"{_safe_stem(file_path)}.pptx"
        file_path = Path(convert_presentation_to_pptx_via_libreoffice(file_path, target))
        converted_from = "PPT经LibreOffice转换"
        ext = ".pptx"

    if ext == ".docx":
        content = _extract_docx_content(file_path)
    if ext == ".xlsx":
        content = _extract_xlsx_content(file_path)
    if ext == ".pptx":
        content = _extract_pptx_content(file_path)
    if ext == ".pdf":
        content = _extract_pdf_content(file_path)
    if ext == ".txt":
        content = _extract_txt_content(file_path)
    if ext not in {".docx", ".xlsx", ".pptx", ".pdf", ".txt"}:
        raise ValueError(f"不支持的文件格式: {ext}")
    if converted_from:
        content = ExtractedContent(
            items=content.items,
            page_count=content.page_count,
            paragraph_count=content.paragraph_count,
            line_count=content.line_count,
            image_count=content.image_count,
            stat_method=f"{converted_from}+{content.stat_method}",
            file_type=content.file_type,
        )
    return content


def _safe_stem(path: Path) -> str:
    normalized = re.sub(r"[^0-9A-Za-z\u4e00-\u9fff._-]+", "_", path.stem).strip("._")
    return normalized or "converted"


def _extract_txt_content(path: Path) -> ExtractedContent:
    items = _extract_txt_text_items(path)
    text = "\n".join(item.text for item in items)
    return ExtractedContent(
        items=items,
        page_count=1 if text.strip() else 0,
        paragraph_count=_count_text_paragraphs(text),
        line_count=_count_text_lines(text),
        image_count=0,
        stat_method="TXT文本解析+Word近似计数",
        file_type="TXT",
    )


def _extract_txt_text_items(path: Path) -> list[TextItem]:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "gb18030", "big5"):
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        try:
            import chardet

            detected = chardet.detect(raw).get("encoding") or "utf-8"
            text = raw.decode(detected, errors="replace")
        except Exception:
            text = raw.decode("utf-8", errors="replace")
    return [
        TextItem(
            "txt",
            "文本文件",
            text,
            is_extra=False,
            paragraph_count=_count_text_paragraphs(text),
            line_count=_count_text_lines(text),
        )
    ] if text else []


def _extract_pdf_content(path: Path) -> ExtractedContent:
    import fitz

    items: list[TextItem] = []
    with fitz.open(str(path)) as doc:
        page_count = doc.page_count
        for page_index, page in enumerate(doc, start=1):
            text = page.get_text("text").strip()
            if text:
                items.append(
                    TextItem(
                        "pdf_page",
                        f"第 {page_index} 页",
                        text,
                        is_extra=False,
                        paragraph_count=_count_text_lines(text),
                        line_count=_count_text_lines(text),
                    )
                )
    return ExtractedContent(
        items=items,
        page_count=page_count,
        paragraph_count=sum(item.paragraph_count for item in items),
        line_count=sum(item.line_count for item in items),
        image_count=_count_pdf_images(path),
        stat_method="PDF文本层解析+Word近似计数",
        file_type="PDF",
    )


def _extract_pdf_text_items(path: Path) -> list[TextItem]:
    return _extract_pdf_content(path).items


def _extract_xlsx_content(path: Path) -> ExtractedContent:
    workbook = load_workbook(path, read_only=True, data_only=True)
    items: list[TextItem] = []
    sheet_count = 0
    total_non_empty_rows = 0
    total_non_empty_cells = 0
    try:
        for sheet in workbook.worksheets:
            sheet_count += 1
            values: list[str] = []
            sheet_non_empty_rows = 0
            sheet_non_empty_cells = 0
            for row in sheet.iter_rows():
                row_values: list[str] = []
                for cell in row:
                    if cell.value is None:
                        continue
                    text = str(cell.value).strip()
                    if text:
                        row_values.append(text)
                if row_values:
                    sheet_non_empty_rows += 1
                    sheet_non_empty_cells += len(row_values)
                    values.extend(row_values)
            total_non_empty_rows += sheet_non_empty_rows
            total_non_empty_cells += sheet_non_empty_cells
            if values:
                items.append(
                    TextItem(
                        "excel_cell",
                        f"工作表: {sheet.title}",
                        "\n".join(values),
                        is_extra=False,
                        paragraph_count=sheet_non_empty_cells,
                        line_count=sheet_non_empty_rows,
                    )
                )
    finally:
        workbook.close()
    return ExtractedContent(
        items=items,
        page_count=sheet_count,
        paragraph_count=total_non_empty_cells,
        line_count=total_non_empty_rows,
        image_count=_count_ooxml_image_references(path, ("xl/drawings/",)),
        stat_method="XLSX单元格解析+Word近似计数（页数按工作表数近似）",
        file_type="Excel",
    )


def _extract_xlsx_text_items(path: Path) -> list[TextItem]:
    return _extract_xlsx_content(path).items


def _extract_pptx_content(path: Path) -> ExtractedContent:
    from pptx import Presentation

    prs = Presentation(str(path))
    items: list[TextItem] = []
    for slide_index, slide in enumerate(prs.slides, start=1):
        texts: list[str] = []
        paragraph_count = 0
        line_count = 0
        for shape in _iter_ppt_shapes(slide.shapes):
            shape_texts, shape_paragraphs, shape_lines = _extract_ppt_shape_texts(shape)
            texts.extend(shape_texts)
            paragraph_count += shape_paragraphs
            line_count += shape_lines
        if texts:
            items.append(
                TextItem(
                    "ppt_slide",
                    f"第 {slide_index} 张幻灯片",
                    "\n".join(texts),
                    is_extra=False,
                    paragraph_count=paragraph_count or _count_text_paragraphs("\n".join(texts)),
                    line_count=line_count or _count_text_lines("\n".join(texts)),
                )
            )
        notes_text = _extract_ppt_notes_text(slide)
        if notes_text:
            items.append(
                TextItem(
                    "ppt_notes",
                    f"第 {slide_index} 张幻灯片备注",
                    notes_text,
                    is_extra=True,
                    paragraph_count=_count_text_paragraphs(notes_text),
                    line_count=_count_text_lines(notes_text),
                )
            )
    return ExtractedContent(
        items=items,
        page_count=len(prs.slides),
        paragraph_count=sum(item.paragraph_count for item in items),
        line_count=sum(item.line_count for item in items),
        image_count=_count_ooxml_image_references(path, ("ppt/slides/",)),
        stat_method="PPTX形状文本解析+Word近似计数",
        file_type="PPT",
    )


def _extract_pptx_text_items(path: Path) -> list[TextItem]:
    return _extract_pptx_content(path).items


def _iter_ppt_shapes(shapes) -> Iterable[Any]:
    for shape in shapes:
        yield shape
        nested = getattr(shape, "shapes", None)
        if nested is not None:
            yield from _iter_ppt_shapes(nested)


def _extract_ppt_shape_texts(shape) -> tuple[list[str], int, int]:
    texts: list[str] = []
    paragraph_count = 0
    line_count = 0
    if getattr(shape, "has_text_frame", False):
        text, paragraphs, lines = _extract_text_frame_text(shape.text_frame)
        if text:
            texts.append(text)
            paragraph_count += paragraphs
            line_count += lines
    if getattr(shape, "has_table", False):
        for row in shape.table.rows:
            for cell in row.cells:
                text, paragraphs, lines = _extract_text_frame_text(cell.text_frame)
                if text:
                    texts.append(text)
                    paragraph_count += paragraphs
                    line_count += lines
    if getattr(shape, "has_chart", False):
        chart_text = _extract_ppt_chart_text(shape)
        if chart_text:
            texts.append(chart_text)
            paragraph_count += _count_text_paragraphs(chart_text)
            line_count += _count_text_lines(chart_text)
    return texts, paragraph_count, line_count


def _extract_text_frame_text(text_frame) -> tuple[str, int, int]:
    texts: list[str] = []
    paragraph_count = 0
    line_count = 0
    for paragraph in getattr(text_frame, "paragraphs", []) or []:
        text = (paragraph.text or "").strip()
        if not text:
            continue
        texts.append(text)
        paragraph_count += 1
        line_count += _count_text_lines(text)
    if not texts:
        text = (getattr(text_frame, "text", "") or "").strip()
        if text:
            texts.append(text)
            paragraph_count = _count_text_paragraphs(text)
            line_count = _count_text_lines(text)
    return "\n".join(texts), paragraph_count, line_count


def _extract_ppt_chart_text(shape) -> str:
    try:
        chart_space = getattr(shape.chart, "_chartSpace", None)
        if chart_space is None:
            return ""
        return "\n".join(_extract_chart_texts(etree.tostring(chart_space)))
    except Exception:
        return ""


def _extract_ppt_notes_text(slide) -> str:
    try:
        if not getattr(slide, "has_notes_slide", False):
            return ""
        notes_frame = getattr(slide.notes_slide, "notes_text_frame", None)
        if notes_frame is None:
            return ""
        text, _, _ = _extract_text_frame_text(notes_frame)
        return text
    except Exception:
        return ""


def _qn(tag: str) -> str:
    prefix, local = tag.split(":")
    return f"{{{NS[prefix]}}}{local}"


def _extract_docx_content(path: Path) -> ExtractedContent:
    items = _extract_docx_text_items(path)
    app_props = _read_ooxml_app_properties(path)
    page_count = _to_positive_int(app_props.get("Pages"))
    app_line_count = _to_positive_int(app_props.get("Lines"))
    fallback_line_count = sum(item.line_count for item in items)
    return ExtractedContent(
        items=items,
        page_count=page_count,
        paragraph_count=sum(item.paragraph_count for item in items),
        line_count=app_line_count or fallback_line_count,
        image_count=_count_ooxml_image_references(path, ("word/",)),
        stat_method="DOCX XML解析+Word近似计数",
        file_type="Word",
    )


def _extract_docx_text_items(path: Path) -> list[TextItem]:
    items: list[TextItem] = []
    with ZipFile(path) as zf:
        names = set(zf.namelist())
        if "word/document.xml" not in names:
            raise ValueError("DOCX 结构异常，缺少 word/document.xml")

        document_root = etree.fromstring(zf.read("word/document.xml"))
        body = document_root.find(_qn("w:body"))
        if body is not None:
            for child in body:
                if child.tag == _qn("w:p"):
                    _append_docx_paragraph_items(items, child, "body", "正文")
                elif child.tag == _qn("w:tbl"):
                    _append_docx_table_items(items, child)

        for xml_name in sorted(name for name in names if name.startswith("word/header") and name.endswith(".xml")):
            text = _extract_xml_text(zf.read(xml_name))
            if text.strip():
                items.append(
                    TextItem(
                        "header",
                        "页眉",
                        text,
                        is_extra=True,
                        paragraph_count=_count_text_paragraphs(text),
                        line_count=_count_text_lines(text),
                    )
                )

        for xml_name in sorted(name for name in names if name.startswith("word/footer") and name.endswith(".xml")):
            text = _extract_xml_text(zf.read(xml_name))
            if text.strip():
                items.append(
                    TextItem(
                        "footer",
                        "页脚",
                        text,
                        is_extra=True,
                        paragraph_count=_count_text_paragraphs(text),
                        line_count=_count_text_lines(text),
                    )
                )

        for xml_name, source_type, label in (
            ("word/footnotes.xml", "footnote", "脚注"),
            ("word/endnotes.xml", "endnote", "尾注"),
            ("word/comments.xml", "comment", "批注"),
        ):
            if xml_name in names:
                for text in _extract_note_like_texts(zf.read(xml_name), source_type):
                    items.append(
                        TextItem(
                            source_type,
                            label,
                            text,
                            is_extra=True,
                            paragraph_count=_count_text_paragraphs(text),
                            line_count=_count_text_lines(text),
                        )
                    )

        for xml_name in sorted(name for name in names if name.startswith("word/charts/") and name.endswith(".xml")):
            chart_text = "\n".join(_extract_chart_texts(zf.read(xml_name)))
            if chart_text.strip():
                items.append(
                    TextItem(
                        "chart",
                        "图表",
                        chart_text,
                        is_extra=True,
                        paragraph_count=_count_text_paragraphs(chart_text),
                        line_count=_count_text_lines(chart_text),
                    )
                )

    return items


def _append_docx_paragraph_items(items: list[TextItem], paragraph, source_type: str, label: str) -> None:
    normal_text = _extract_text_from_element(paragraph, skip_textboxes=True)
    if normal_text.strip():
        items.append(
            TextItem(
                source_type,
                label,
                normal_text,
                is_extra=False,
                paragraph_count=1,
                line_count=_count_text_lines(normal_text),
            )
        )
    for textbox_text in _extract_textbox_texts(paragraph):
        if textbox_text.strip():
            items.append(
                TextItem(
                    "textbox",
                    "正文文本框",
                    textbox_text,
                    is_extra=False,
                    paragraph_count=_count_text_paragraphs(textbox_text),
                    line_count=_count_text_lines(textbox_text),
                )
            )


def _append_docx_table_items(items: list[TextItem], table) -> None:
    for row in table.iter(_qn("w:tr")):
        row_texts: list[str] = []
        textbox_texts: list[str] = []
        for cell in row.iter(_qn("w:tc")):
            text = _extract_text_from_element(cell, skip_textboxes=True)
            if text.strip():
                row_texts.append(text)
            textbox_texts.extend(_extract_textbox_texts(cell))
        if row_texts:
            row_text = "\t".join(row_texts)
            items.append(
                TextItem(
                    "table",
                    "正文表格",
                    row_text,
                    is_extra=False,
                    paragraph_count=len(row_texts),
                    line_count=max(1, len(row_texts)),
                )
            )
        for textbox_text in textbox_texts:
            if textbox_text.strip():
                items.append(
                    TextItem(
                        "textbox",
                        "正文文本框",
                        textbox_text,
                        is_extra=False,
                        paragraph_count=_count_text_paragraphs(textbox_text),
                        line_count=_count_text_lines(textbox_text),
                    )
                )


def _extract_xml_text(xml_bytes: bytes) -> str:
    root = etree.fromstring(xml_bytes)
    return _extract_text_from_element(root, skip_textboxes=False)


def _extract_note_like_texts(xml_bytes: bytes, source_type: str) -> list[str]:
    root = etree.fromstring(xml_bytes)
    if source_type == "footnote":
        item_tag = _qn("w:footnote")
    elif source_type == "endnote":
        item_tag = _qn("w:endnote")
    else:
        item_tag = _qn("w:comment")

    texts: list[str] = []
    for elem in root.iter(item_tag):
        note_id = elem.get(_qn("w:id"), "")
        if note_id in {"-1", "0", "1"} and source_type in {"footnote", "endnote"}:
            continue
        text = _extract_text_from_element(elem, skip_textboxes=False)
        if text.strip():
            texts.append(text)
    return texts


def _extract_text_from_element(element, *, skip_textboxes: bool) -> str:
    textbox_text_ids: set[int] = set()
    if skip_textboxes:
        for textbox in _iter_textbox_elements(element):
            for text_node in textbox.iter(_qn("w:t")):
                textbox_text_ids.add(id(text_node))

    deleted_tag = _qn("w:del")
    instr_tag = _qn("w:instrText")
    parts: list[str] = []
    for text_node in element.iter():
        if text_node.tag not in {_qn("w:t"), instr_tag}:
            continue
        if text_node.tag == instr_tag:
            continue
        if id(text_node) in textbox_text_ids:
            continue
        if _has_ancestor(text_node, deleted_tag):
            continue
        if text_node.text:
            parts.append(text_node.text)
    return "".join(parts)


def _iter_textbox_elements(element) -> Iterable[Any]:
    for tag in (_qn("wps:txbx"), _qn("w:txbxContent"), _qn("v:textbox")):
        yield from element.iter(tag)


def _extract_textbox_texts(element) -> list[str]:
    texts: list[str] = []
    seen: set[int] = set()
    for textbox in _iter_textbox_elements(element):
        if id(textbox) in seen:
            continue
        seen.add(id(textbox))
        text = _extract_text_from_element(textbox, skip_textboxes=False)
        if text.strip():
            texts.append(text)
    return texts


def _has_ancestor(node, ancestor_tag: str) -> bool:
    parent = node.getparent()
    while parent is not None:
        if parent.tag == ancestor_tag:
            return True
        parent = parent.getparent()
    return False


def _extract_chart_texts(xml_bytes: bytes) -> list[str]:
    root = etree.fromstring(xml_bytes)
    seen: set[str] = set()
    texts: list[str] = []

    def add(value: Optional[str]) -> None:
        text = (value or "").strip()
        if text and text not in seen:
            seen.add(text)
            texts.append(text)

    for node in root.iter(_qn("a:t")):
        add(node.text)
    for node in root.iter(_qn("c:v")):
        value = (node.text or "").strip()
        if not value:
            continue
        try:
            float(value)
            continue
        except ValueError:
            add(value)
    return texts


def _read_ooxml_app_properties(path: Path) -> dict[str, str]:
    try:
        with ZipFile(path) as zf:
            if "docProps/app.xml" not in zf.namelist():
                return {}
            root = etree.fromstring(zf.read("docProps/app.xml"))
    except Exception:
        return {}
    props: dict[str, str] = {}
    for child in root:
        tag = etree.QName(child).localname
        props[tag] = (child.text or "").strip()
    return props


def _count_ooxml_image_references(path: Path, prefixes: tuple[str, ...]) -> int:
    count = 0
    try:
        with ZipFile(path) as zf:
            for name in zf.namelist():
                if not name.endswith(".xml") or "/_rels/" in name:
                    continue
                if prefixes and not any(name.startswith(prefix) for prefix in prefixes):
                    continue
                try:
                    root = etree.fromstring(zf.read(name))
                except Exception:
                    continue
                for node in root.iter():
                    local_name = etree.QName(node).localname
                    if local_name in {"blip", "imagedata"}:
                        count += 1
    except Exception:
        return 0
    return count


def _count_pdf_images(path: Path) -> int:
    try:
        import fitz

        with fitz.open(str(path)) as doc:
            return sum(len(page.get_images(full=True)) for page in doc)
    except Exception:
        return 0


def _count_text_lines(text: str) -> int:
    return sum(1 for line in re.split(r"\r\n|\r|\n", text or "") if line.strip())


def _count_text_paragraphs(text: str) -> int:
    content = (text or "").strip()
    if not content:
        return 0
    blocks = [block for block in re.split(r"(?:\r?\n\s*){2,}", content) if block.strip()]
    return len(blocks) if blocks else _count_text_lines(content)


def _to_positive_int(value: Any) -> int:
    try:
        number = int(str(value or "").strip())
    except (TypeError, ValueError):
        return 0
    return number if number > 0 else 0


def _now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _file_type_for_extension(extension: str) -> str:
    ext = (extension or "").lower()
    if ext in {".doc", ".docx"}:
        return "Word"
    if ext in {".xls", ".xlsx"}:
        return "Excel"
    if ext in {".ppt", ".pptx"}:
        return "PPT"
    if ext == ".pdf":
        return "PDF"
    if ext == ".txt":
        return "TXT"
    if ext in IMAGE_EXTENSIONS:
        return "图片"
    if ext in CAD_EXTENSIONS:
        return "CAD"
    return "未知"


def _stat_method_for_extension(extension: str) -> str:
    ext = (extension or "").lower()
    if ext == ".docx":
        return "DOCX XML解析+Word近似计数"
    if ext == ".doc":
        return "DOC经LibreOffice转换+DOCX XML解析+Word近似计数"
    if ext == ".xlsx":
        return "XLSX单元格解析+Word近似计数（页数按工作表数近似）"
    if ext == ".xls":
        return "XLS经LibreOffice转换+XLSX单元格解析+Word近似计数（页数按工作表数近似）"
    if ext == ".pptx":
        return "PPTX形状文本解析+Word近似计数"
    if ext == ".ppt":
        return "PPT经LibreOffice转换+PPTX形状文本解析+Word近似计数"
    if ext == ".pdf":
        return "PDF文本层解析+Word近似计数"
    if ext == ".txt":
        return "TXT文本解析+Word近似计数"
    if ext in IMAGE_EXTENSIONS:
        return "图片候选文件，需OCR后统计"
    if ext in CAD_EXTENSIONS:
        return "CAD候选文件，需专用解析器后统计"
    return ""


def _count_text(text: str) -> TextMetrics:
    word_count = 0
    in_token = False
    content = text or ""
    for index, char in enumerate(content):
        if _is_cjk(char):
            word_count += 1
            in_token = False
            continue
        if _is_token_char(char):
            if not in_token:
                word_count += 1
                in_token = True
            continue
        if char in {"'", "’", "-", "_", ".", "/"} and in_token:
            continue
        previous_char = content[index - 1] if index > 0 else ""
        next_char = content[index + 1] if index + 1 < len(content) else ""
        if _is_word_count_cjk_punctuation(char, previous_char, next_char):
            word_count += 1
            in_token = False
            continue
        in_token = False
    return TextMetrics(
        word_count=word_count,
        non_space_chars=sum(1 for char in text or "" if not char.isspace()),
        raw_chars=len(text or ""),
    )


def _is_cjk(char: str) -> bool:
    if not char or len(char) != 1:
        return False
    code = ord(char)
    return (
        0x3400 <= code <= 0x4DBF
        or 0x4E00 <= code <= 0x9FFF
        or 0xF900 <= code <= 0xFAFF
        or 0x3040 <= code <= 0x30FF
        or 0xAC00 <= code <= 0xD7AF
    )


def _is_token_char(char: str) -> bool:
    if not char or _is_cjk(char):
        return False
    category = unicodedata.category(char)
    return category[0] in {"L", "N"}


def _is_word_count_cjk_punctuation(char: str, previous_char: str = "", next_char: str = "") -> bool:
    if not char or char.isspace() or _is_token_char(char):
        return False
    code = ord(char)
    category = unicodedata.category(char)
    if category[0] not in {"P", "S"}:
        return False

    # Word 的中文统计会把中文/全角标点算入“中文字符”，但不会把 © 这类普通符号算入字数。
    if (
        0x3000 <= code <= 0x303F
        or 0xFE10 <= code <= 0xFE1F
        or 0xFE30 <= code <= 0xFE4F
        or 0xFF00 <= code <= 0xFFEF
    ):
        return True

    chinese_context_punctuation = {"“", "”", "‘", "’", "—", "–", "…", "·"}
    if char in chinese_context_punctuation and (_is_cjk(previous_char) or _is_cjk(next_char)):
        return True

    return False


def _build_summary(file_results: list[dict[str, Any]], *, truncated: bool, started_at: datetime) -> dict[str, Any]:
    status_counts = Counter(item.get("status") for item in file_results)
    return {
        "total_files": len(file_results),
        "counted_files": status_counts.get(STATUS_COUNTED, 0),
        "failed_files": status_counts.get(STATUS_FAILED, 0),
        "skipped_files": status_counts.get(STATUS_SKIPPED, 0),
        "needs_ocr_files": status_counts.get(STATUS_NEEDS_OCR, 0),
        "needs_cad_parser_files": status_counts.get(STATUS_NEEDS_CAD_PARSER, 0),
        "total_main_word_count": sum(int(item.get("main_word_count") or 0) for item in file_results),
        "total_word_count": sum(int(item.get("word_count") or item.get("main_word_count") or 0) for item in file_results),
        "total_extra_word_count": sum(int(item.get("extra_word_count") or 0) for item in file_results),
        "total_char_count_no_spaces": sum(int(item.get("char_count_no_spaces") or 0) for item in file_results),
        "total_char_count_with_spaces": sum(int(item.get("char_count_with_spaces") or 0) for item in file_results),
        "total_non_space_chars": sum(int(item.get("char_count_no_spaces") or item.get("non_space_chars") or 0) for item in file_results),
        "total_raw_chars": sum(int(item.get("char_count_with_spaces") or item.get("raw_chars") or 0) for item in file_results),
        "total_page_count": sum(int(item.get("page_count") or 0) for item in file_results),
        "total_paragraph_count": sum(int(item.get("paragraph_count") or 0) for item in file_results),
        "total_line_count": sum(int(item.get("line_count") or 0) for item in file_results),
        "total_image_count": sum(int(item.get("image_count") or 0) for item in file_results),
        "truncated": truncated,
        "started_at": started_at.isoformat(timespec="seconds"),
        "finished_at": datetime.now().isoformat(timespec="seconds"),
    }


def _write_excel_report(path: Path, payload: dict[str, Any]) -> None:
    workbook = Workbook()
    summary_sheet = workbook.active
    summary_sheet.title = "汇总"
    _write_summary_sheet(summary_sheet, payload)
    _write_file_sheet(workbook.create_sheet("文件明细"), payload.get("files") or [])
    _write_source_sheet(workbook.create_sheet("来源明细"), payload.get("source_details") or [])
    _write_fail_sheet(workbook.create_sheet("跳过与失败"), payload.get("files") or [])
    _write_rules_sheet(workbook.create_sheet("规则说明"), payload.get("rules") or [])
    workbook.save(path)


def _style_header(sheet, row: int = 1) -> None:
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)
    for cell in sheet[row]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")


def _autosize(sheet) -> None:
    for column in sheet.columns:
        max_length = 10
        column_letter = get_column_letter(column[0].column)
        for cell in column:
            value = "" if cell.value is None else str(cell.value)
            max_length = max(max_length, min(len(value) + 2, 60))
        sheet.column_dimensions[column_letter].width = max_length


def _write_summary_sheet(sheet, payload: dict[str, Any]) -> None:
    summary = payload.get("summary") or {}
    rows = [
        ("统计目录", payload.get("directory_path", "")),
        ("允许根目录", payload.get("allowed_root", "")),
        ("主字数合计", summary.get("total_main_word_count", 0)),
        ("额外内容字数合计", summary.get("total_extra_word_count", 0)),
        ("字符数(不计空格)合计", summary.get("total_char_count_no_spaces", summary.get("total_non_space_chars", 0))),
        ("字符数(计空格)合计", summary.get("total_char_count_with_spaces", summary.get("total_raw_chars", 0))),
        ("页数合计", summary.get("total_page_count", 0)),
        ("段落数合计", summary.get("total_paragraph_count", 0)),
        ("行数合计", summary.get("total_line_count", 0)),
        ("图片数量合计", summary.get("total_image_count", 0)),
        ("文件总数", summary.get("total_files", 0)),
        ("已统计文件数", summary.get("counted_files", 0)),
        ("失败文件数", summary.get("failed_files", 0)),
        ("跳过文件数", summary.get("skipped_files", 0)),
        ("需 OCR 文件数", summary.get("needs_ocr_files", 0)),
        ("需 CAD 解析文件数", summary.get("needs_cad_parser_files", 0)),
        ("是否截断", "是" if summary.get("truncated") else "否"),
        ("生成时间", payload.get("generated_at", "")),
    ]
    sheet.append(["项目", "值"])
    for row in rows:
        sheet.append(list(row))
    _style_header(sheet)
    _autosize(sheet)


def _write_file_sheet(sheet, files: list[dict[str, Any]]) -> None:
    headers = [
        "相对路径",
        "页数",
        "字数",
        "额外字数",
        "字符数(不计空格)",
        "字符数(计空格)",
        "段落数",
        "行数",
        "图片数量",
        "统计方法",
        "文件类型",
        "状态",
        "错误信息",
        "统计时间",
        "扩展名",
        "文件大小",
        "修改时间",
        "消息",
        "警告",
        "错误",
    ]
    sheet.append(headers)
    for item in files:
        sheet.append(
            [
                item.get("relative_path", ""),
                item.get("page_count", 0),
                item.get("word_count", item.get("main_word_count", 0)),
                item.get("extra_word_count", 0),
                item.get("char_count_no_spaces", item.get("non_space_chars", 0)),
                item.get("char_count_with_spaces", item.get("raw_chars", 0)),
                item.get("paragraph_count", 0),
                item.get("line_count", 0),
                item.get("image_count", 0),
                item.get("stat_method", ""),
                item.get("file_type", ""),
                item.get("status", ""),
                item.get("error", ""),
                item.get("counted_at", ""),
                item.get("extension", ""),
                item.get("size_bytes", 0),
                item.get("modified_at", ""),
                item.get("message", ""),
                item.get("warning", ""),
                item.get("error", ""),
            ]
        )
    _style_header(sheet)
    _autosize(sheet)


def _write_source_sheet(sheet, rows: list[dict[str, Any]]) -> None:
    headers = [
        "相对路径",
        "扩展名",
        "来源类型",
        "来源标签",
        "是否额外内容",
        "字数",
        "字符数(不计空格)",
        "字符数(计空格)",
        "段落数",
        "行数",
        "图片数量",
        "文本预览",
    ]
    sheet.append(headers)
    for row in rows:
        sheet.append(
            [
                row.get("relative_path", ""),
                row.get("extension", ""),
                row.get("source_type", ""),
                row.get("source_label", ""),
                "是" if row.get("is_extra") else "否",
                row.get("word_count", 0),
                row.get("char_count_no_spaces", row.get("non_space_chars", 0)),
                row.get("char_count_with_spaces", row.get("raw_chars", 0)),
                row.get("paragraph_count", 0),
                row.get("line_count", 0),
                row.get("image_count", 0),
                row.get("text_preview", ""),
            ]
        )
    _style_header(sheet)
    _autosize(sheet)


def _write_fail_sheet(sheet, files: list[dict[str, Any]]) -> None:
    sheet.append(["相对路径", "扩展名", "文件类型", "状态", "统计方法", "图片数量", "消息", "警告", "错误", "统计时间"])
    for item in files:
        if item.get("status") == STATUS_COUNTED:
            continue
        sheet.append(
            [
                item.get("relative_path", ""),
                item.get("extension", ""),
                item.get("file_type", ""),
                item.get("status", ""),
                item.get("stat_method", ""),
                item.get("image_count", 0),
                item.get("message", ""),
                item.get("warning", ""),
                item.get("error", ""),
                item.get("counted_at", ""),
            ]
        )
    _style_header(sheet)
    _autosize(sheet)


def _write_rules_sheet(sheet, rules: list[str]) -> None:
    sheet.append(["规则说明"])
    for rule in rules:
        sheet.append([rule])
    _style_header(sheet)
    _autosize(sheet)


def _relative_path(path: Path, root: Path) -> str:
    try:
        return path.resolve(strict=False).relative_to(root.resolve(strict=False)).as_posix()
    except ValueError:
        return path.as_posix()


def _preview_text(text: str, limit: int = 120) -> str:
    normalized = re.sub(r"\s+", " ", text or "").strip()
    return normalized[:limit]


def _output_web_path(path: Path) -> str:
    output_root = Path(settings.OUTPUT_DIR).resolve()
    resolved = path.resolve()
    try:
        return f"outputs/{resolved.relative_to(output_root).as_posix()}"
    except ValueError:
        return str(path).replace("\\", "/")
