# -*- coding: utf-8 -*-
from __future__ import annotations

import asyncio
import re
import shutil
from concurrent.futures import Executor
from pathlib import Path, PurePosixPath
from typing import Any, Awaitable, Callable, Iterable, Optional
from zipfile import ZIP_DEFLATED, ZipFile

import fitz
from pypdf import PdfReader, PdfWriter

from app.core.config import settings
from app.service.pdf_merge_service import get_pdf_merge_config, normalize_output_filename
from app.service.word_count_service import resolve_allowed_shared_input_path


ProgressCallback = Callable[[int, str], Awaitable[None]]
PDF_TOOL_OPERATIONS = {"split", "compress", "extract", "delete", "rotate"}
OPERATION_LABELS = {
    "split": "拆分 PDF",
    "compress": "压缩 PDF",
    "extract": "提取页面",
    "delete": "删除页面",
    "rotate": "旋转页面",
}
COMPRESSION_MODES = {
    "lossless": {
        "label": "无损优化",
        "description": "清理冗余对象并压缩字体、图片和内容流，保持原始画质。",
    },
    "high": {
        "label": "清晰优先",
        "description": "图片降至约 200 DPI、JPEG 质量 85，适合文字和正式阅读。",
    },
    "balanced": {
        "label": "均衡压缩",
        "description": "图片降至约 144 DPI、JPEG 质量 75，兼顾清晰度和体积。",
    },
    "strong": {
        "label": "强力压缩",
        "description": "图片降至约 96 DPI、JPEG 质量 55，适合预览和归档副本。",
    },
}
COPY_CHUNK_SIZE = 1024 * 1024


def get_pdf_tools_config() -> dict[str, Any]:
    shared = get_pdf_merge_config()
    return {
        **shared,
        "operations": [
            {"value": key, "label": value}
            for key, value in OPERATION_LABELS.items()
        ],
        "compression_modes": COMPRESSION_MODES,
        "split_modes": [
            {"value": "every", "label": "每 N 页一个文件"},
            {"value": "ranges", "label": "自定义页组"},
        ],
        "rotate_angles": [90, 180, 270],
        "page_selection_modes": [
            {"value": "custom", "label": "自定义页码"},
            {"value": "odd", "label": "奇数页"},
            {"value": "even", "label": "偶数页"},
        ],
        "page_spec_example": "1-3,5,8",
        "page_groups_example": "1-3;4,6;7-9",
    }


def prepare_pdf_tools_request(
    *,
    directory_path: str,
    relative_path: str,
    operation: str,
    options: Optional[dict[str, Any]] = None,
) -> dict[str, Any]:
    normalized_operation = str(operation or "").strip().lower()
    if normalized_operation not in PDF_TOOL_OPERATIONS:
        raise ValueError(f"不支持的 PDF 操作: {normalized_operation or '空'}")

    source = _resolve_source_pdf(directory_path, relative_path)
    normalized_options = _normalize_operation_options(
        normalized_operation,
        options or {},
        source_name=source["path"].name,
    )
    normalized_relative_path = source["relative_path"]
    return {
        "filename": f"{OPERATION_LABELS[normalized_operation]}：{source['path'].name}",
        "params": {
            "operation": normalized_operation,
            "operation_label": OPERATION_LABELS[normalized_operation],
            "directory_path": str(directory_path or "").strip(),
            "relative_path": normalized_relative_path,
            "options": normalized_options,
        },
        "input_files": {
            "directory_path": str(directory_path or "").strip(),
            "relative_path": normalized_relative_path,
        },
    }


async def execute_pdf_tools_task(
    *,
    task_id: str,
    display_no: Optional[str],
    directory_path: str,
    relative_path: str,
    operation: str,
    options: Optional[dict[str, Any]] = None,
    progress_callback: Optional[ProgressCallback] = None,
    executor: Optional[Executor] = None,
) -> dict[str, Any]:
    loop = asyncio.get_running_loop()
    normalized_operation = str(operation or "").strip().lower()
    if normalized_operation not in PDF_TOOL_OPERATIONS:
        raise ValueError(f"不支持的 PDF 操作: {normalized_operation or '空'}")

    await _report(progress_callback, 5, "正在校验共享路径和源 PDF...")
    source = await loop.run_in_executor(
        executor,
        lambda: _resolve_source_pdf(directory_path, relative_path),
    )
    normalized_options = _normalize_operation_options(
        normalized_operation,
        options or {},
        source_name=source["path"].name,
    )

    task_folder = display_no or task_id
    staging_dir = Path(settings.UPLOAD_DIR) / "pdf_tools_staging" / task_folder
    output_dir = Path(settings.OUTPUT_DIR) / "pdf_tools" / task_folder
    staging_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    staged_path = staging_dir / "source.pdf"

    try:
        await _report(progress_callback, 12, f"正在暂存: {source['relative_path']}")
        await loop.run_in_executor(
            executor,
            lambda: _copy_source_pdf(source["path"], staged_path),
        )
        input_size = staged_path.stat().st_size
        page_count = await loop.run_in_executor(executor, lambda: _pdf_page_count(staged_path))

        if normalized_operation == "split":
            await _report(progress_callback, 35, "正在按设置拆分页面...")
            result = await loop.run_in_executor(
                executor,
                lambda: _split_pdf(staged_path, output_dir, normalized_options, page_count),
            )
        elif normalized_operation == "compress":
            await _report(progress_callback, 35, "正在重写图片并清理冗余对象...")
            result = await loop.run_in_executor(
                executor,
                lambda: _compress_pdf(staged_path, output_dir, normalized_options, page_count),
            )
        else:
            await _report(progress_callback, 35, f"正在{OPERATION_LABELS[normalized_operation]}...")
            result = await loop.run_in_executor(
                executor,
                lambda: _process_pages(
                    staged_path,
                    output_dir,
                    normalized_operation,
                    normalized_options,
                    page_count,
                ),
            )

        await _report(progress_callback, 94, "正在校验输出文件...")
        result.update(
            {
                "operation": normalized_operation,
                "operation_label": OPERATION_LABELS[normalized_operation],
                "source_file": source["relative_path"],
                "input_size": input_size,
                "source_page_count": page_count,
            }
        )
        return result
    finally:
        shutil.rmtree(staging_dir, ignore_errors=True)


def parse_page_spec(spec: str, total_pages: int, *, allow_all: bool = True) -> list[int]:
    """把用户输入的一页式页码表达式解析为从零开始的页码列表。"""
    if total_pages <= 0:
        raise ValueError("PDF 不包含可处理页面")
    raw = str(spec or "").strip().lower()
    if allow_all and raw in {"", "all", "全部", "所有"}:
        return list(range(total_pages))
    if not raw:
        raise ValueError("请填写页码范围")

    normalized = raw.replace("，", ",").replace("、", ",").replace("—", "-").replace("–", "-")
    pages: list[int] = []
    seen: set[int] = set()
    for token in re.split(r"[,\s]+", normalized):
        if not token:
            continue
        match = re.fullmatch(r"(\d+)(?:-(\d+))?", token)
        if not match:
            raise ValueError(f"页码格式无效: {token}")
        start = int(match.group(1))
        end = int(match.group(2) or start)
        if start < 1 or end < start or end > total_pages:
            raise ValueError(f"页码超出范围: {token}（文档共 {total_pages} 页）")
        for page_number in range(start, end + 1):
            page_index = page_number - 1
            if page_index not in seen:
                seen.add(page_index)
                pages.append(page_index)
    if not pages:
        raise ValueError("页码范围没有选中任何页面")
    return pages


def _normalize_operation_options(operation: str, options: dict[str, Any], *, source_name: str) -> dict[str, Any]:
    source_stem = Path(source_name).stem or "PDF"
    if operation == "split":
        split_mode = str(options.get("split_mode") or "every").strip().lower()
        if split_mode not in {"every", "ranges"}:
            raise ValueError(f"不支持的拆分模式: {split_mode}")
        try:
            pages_per_file = int(options.get("pages_per_file") or 1)
        except (TypeError, ValueError) as exc:
            raise ValueError("每份页数必须是整数") from exc
        if not 1 <= pages_per_file <= 1000:
            raise ValueError("每份页数必须在 1 到 1000 之间")
        page_groups = str(options.get("page_groups") or "").strip()
        if split_mode == "ranges" and not page_groups:
            raise ValueError("自定义拆分时请填写页组，例如 1-3;4,6;7-9")
        return {
            "split_mode": split_mode,
            "pages_per_file": pages_per_file,
            "page_groups": page_groups,
            "output_prefix": _normalize_output_prefix(options.get("output_prefix") or f"{source_stem}_拆分"),
        }

    output_filename = normalize_output_filename(
        options.get("output_filename")
        or {
            "compress": f"{source_stem}_压缩.pdf",
            "extract": f"{source_stem}_提取.pdf",
            "delete": f"{source_stem}_删页.pdf",
            "rotate": f"{source_stem}_旋转.pdf",
        }[operation]
    )
    if operation == "compress":
        mode = str(options.get("compression_mode") or "high").strip().lower()
        if mode not in COMPRESSION_MODES:
            raise ValueError(f"不支持的压缩模式: {mode}")
        return {"compression_mode": mode, "output_filename": output_filename}

    if operation in {"extract", "delete"}:
        page_mode = str(options.get("page_mode") or "custom").strip().lower()
        if page_mode not in {"custom", "odd", "even"}:
            raise ValueError(f"不支持的页面选择模式: {page_mode}")
        page_spec = str(options.get("page_spec") or "").strip()
        if page_mode == "custom" and not page_spec:
            raise ValueError("请填写需要处理的页码")
        return {
            "page_mode": page_mode,
            "page_spec": page_spec,
            "output_filename": output_filename,
        }

    page_spec = str(options.get("page_spec") or "all").strip()
    normalized = {"page_spec": page_spec, "output_filename": output_filename}
    if operation == "rotate":
        try:
            angle = int(options.get("angle") or 90)
        except (TypeError, ValueError) as exc:
            raise ValueError("旋转角度必须是 90、180 或 270") from exc
        if angle not in {90, 180, 270}:
            raise ValueError("旋转角度必须是 90、180 或 270")
        normalized["angle"] = angle
    return normalized


def _resolve_source_pdf(directory_path: str, relative_path: str) -> dict[str, Any]:
    root, _, input_kind = resolve_allowed_shared_input_path(directory_path)
    if input_kind != "directory":
        raise ValueError(f"请输入包含 PDF 的目录路径: {root}")

    raw_relative = str(relative_path or "").strip().replace("\\", "/")
    pure_path = PurePosixPath(raw_relative)
    if not raw_relative or pure_path.is_absolute() or any(part in {"", ".", ".."} for part in pure_path.parts):
        raise ValueError(f"PDF 相对路径无效: {relative_path}")
    candidate = root.joinpath(*pure_path.parts).resolve(strict=False)
    if not _is_relative_to(candidate, root):
        raise PermissionError(f"PDF 路径超出所选共享目录: {raw_relative}")
    if candidate.suffix.lower() != ".pdf":
        raise ValueError(f"只支持 PDF 文件: {raw_relative}")
    if not candidate.exists() or not candidate.is_file():
        raise FileNotFoundError(f"PDF 文件不存在或暂时不可访问: {raw_relative}")

    size = candidate.stat().st_size
    max_bytes = max(1, int(settings.PDF_MERGE_MAX_FILE_MB or 500)) * 1024 * 1024
    if size <= 0:
        raise ValueError(f"PDF 文件为空: {raw_relative}")
    if size > max_bytes:
        raise ValueError(f"PDF 超过 {settings.PDF_MERGE_MAX_FILE_MB} MB 限制: {raw_relative}")
    with candidate.open("rb") as source:
        if b"%PDF-" not in source.read(1024):
            raise ValueError(f"文件扩展名为 PDF，但文件头无效: {raw_relative}")
    return {"path": candidate, "relative_path": pure_path.as_posix(), "size": size}


def _copy_source_pdf(source_path: Path, staged_path: Path) -> None:
    partial_path = staged_path.with_suffix(".partial.pdf")
    try:
        with source_path.open("rb") as source, partial_path.open("wb") as target:
            shutil.copyfileobj(source, target, length=COPY_CHUNK_SIZE)
        partial_path.replace(staged_path)
    finally:
        partial_path.unlink(missing_ok=True)


def _pdf_page_count(path: Path) -> int:
    with path.open("rb") as source:
        reader = PdfReader(source, strict=False)
        if reader.is_encrypted and reader.decrypt("") == 0:
            raise ValueError("PDF 已加密且需要密码，暂不支持处理")
        pages = len(reader.pages)
    if pages <= 0:
        raise ValueError("PDF 不包含可处理页面")
    return pages


def _split_pdf(source_path: Path, output_dir: Path, options: dict[str, Any], total_pages: int) -> dict[str, Any]:
    if options["split_mode"] == "every":
        step = int(options["pages_per_file"])
        groups = [list(range(start, min(start + step, total_pages))) for start in range(0, total_pages, step)]
    else:
        raw_groups = [item.strip() for item in re.split(r"[;；\n]+", options["page_groups"]) if item.strip()]
        groups = [parse_page_spec(item, total_pages, allow_all=False) for item in raw_groups]
    if not groups:
        raise ValueError("拆分设置没有生成任何页组")
    if len(groups) > max(2, int(settings.PDF_MERGE_MAX_FILES or 200)):
        raise ValueError(f"拆分结果超过 {settings.PDF_MERGE_MAX_FILES} 个文件限制")

    prefix = options["output_prefix"]
    files: list[dict[str, Any]] = []
    with source_path.open("rb") as source:
        reader = PdfReader(source, strict=False)
        if reader.is_encrypted and reader.decrypt("") == 0:
            raise ValueError("PDF 已加密且需要密码，暂不支持拆分")
        for index, pages in enumerate(groups, start=1):
            page_label = _page_group_label(pages)
            filename = f"{prefix}_{index:03d}_{page_label}.pdf"
            output_path = output_dir / filename
            _write_page_subset(reader, pages, output_path, title=filename[:-4])
            files.append(
                {
                    "path": _output_web_path(output_path),
                    "filename": filename,
                    "pages": [page + 1 for page in pages],
                    "page_count": len(pages),
                    "size": output_path.stat().st_size,
                }
            )

    archive_path = output_dir / f"{prefix}_全部文件.zip"
    with ZipFile(archive_path, "w", compression=ZIP_DEFLATED) as archive:
        for item in files:
            archive.write(output_dir / item["filename"], arcname=item["filename"])
    return {
        "archive_zip": _output_web_path(archive_path),
        "archive_filename": archive_path.name,
        "output_files": files,
        "output_file_count": len(files),
        "output_size": sum(int(item["size"]) for item in files),
        "summary_text": f"已将 {total_pages} 页拆分为 {len(files)} 个 PDF。",
    }


def _compress_pdf(source_path: Path, output_dir: Path, options: dict[str, Any], total_pages: int) -> dict[str, Any]:
    mode = options["compression_mode"]
    output_path = output_dir / options["output_filename"]
    candidate_path = output_path.with_suffix(".optimized.pdf")
    source_size = source_path.stat().st_size
    candidate_path.unlink(missing_ok=True)
    document = fitz.open(source_path)
    try:
        if document.needs_pass:
            raise ValueError("PDF 已加密且需要密码，暂不支持压缩")
        if mode == "high":
            document.rewrite_images(dpi_threshold=240, dpi_target=200, quality=85)
        elif mode == "balanced":
            document.rewrite_images(dpi_threshold=180, dpi_target=144, quality=75)
        elif mode == "strong":
            document.rewrite_images(dpi_threshold=130, dpi_target=96, quality=55)
        if mode != "lossless":
            try:
                document.subset_fonts(verbose=False, fallback=False)
            except Exception:
                pass
        document.save(
            candidate_path,
            garbage=4,
            clean=True,
            deflate=True,
            deflate_images=True,
            deflate_fonts=True,
            use_objstms=1,
            compression_effort=100,
        )
    except Exception:
        try:
            candidate_path.unlink(missing_ok=True)
        except OSError:
            pass
        raise
    finally:
        document.close()

    candidate_size = candidate_path.stat().st_size
    reduced = candidate_size < source_size
    if reduced:
        candidate_path.replace(output_path)
    else:
        shutil.copyfile(source_path, output_path)
        candidate_path.unlink(missing_ok=True)
    output_size = output_path.stat().st_size
    saved_bytes = max(0, source_size - output_size)
    reduction_percent = round(saved_bytes / source_size * 100, 2) if source_size else 0.0
    _verify_output_pdf(output_path, expected_pages=total_pages)
    return {
        "output_pdf": _output_web_path(output_path),
        "output_filename": output_path.name,
        "compression_mode": mode,
        "compression_mode_label": COMPRESSION_MODES[mode]["label"],
        "output_size": output_size,
        "saved_bytes": saved_bytes,
        "reduction_percent": reduction_percent,
        "was_reduced": reduced,
        "total_pages": total_pages,
        "summary_text": (
            f"压缩完成，文件体积减少 {reduction_percent:.2f}%。"
            if reduced
            else "源 PDF 已高度压缩，已保留原始体积，避免输出文件变大。"
        ),
    }


def _process_pages(
    source_path: Path,
    output_dir: Path,
    operation: str,
    options: dict[str, Any],
    total_pages: int,
) -> dict[str, Any]:
    page_mode = options.get("page_mode", "custom")
    if operation in {"extract", "delete"} and page_mode in {"odd", "even"}:
        start_index = 0 if page_mode == "odd" else 1
        selected_pages = list(range(start_index, total_pages, 2))
        if not selected_pages:
            page_label = "奇数页" if page_mode == "odd" else "偶数页"
            raise ValueError(f"文档中没有可处理的{page_label}")
    else:
        selected_pages = parse_page_spec(options["page_spec"], total_pages, allow_all=operation == "rotate")
    selected_set = set(selected_pages)
    output_path = output_dir / options["output_filename"]
    with source_path.open("rb") as source:
        reader = PdfReader(source, strict=False)
        if reader.is_encrypted and reader.decrypt("") == 0:
            raise ValueError("PDF 已加密且需要密码，暂不支持页面处理")
        writer = PdfWriter()
        try:
            if operation == "extract":
                output_pages = selected_pages
            elif operation == "delete":
                output_pages = [index for index in range(total_pages) if index not in selected_set]
                if not output_pages:
                    raise ValueError("不能删除全部页面；如需拆分，请使用拆分 PDF")
            else:
                output_pages = list(range(total_pages))

            for page_index in output_pages:
                page = reader.pages[page_index]
                if operation == "rotate" and page_index in selected_set:
                    page.rotate(int(options["angle"]))
                writer.add_page(page)
            if reader.metadata:
                metadata = {key: str(value) for key, value in reader.metadata.items() if key and value is not None}
                if metadata:
                    writer.add_metadata(metadata)
            with output_path.open("wb") as target:
                writer.write(target)
        finally:
            writer.close()

    output_page_count = _verify_output_pdf(output_path, expected_pages=len(output_pages))
    if operation == "extract":
        mode_label = {"odd": "奇数页", "even": "偶数页"}.get(page_mode, "指定页面")
        summary_text = f"已提取{mode_label} {len(selected_pages)} 页，生成 {output_page_count} 页 PDF。"
    elif operation == "delete":
        mode_label = {"odd": "奇数页", "even": "偶数页"}.get(page_mode, "指定页面")
        summary_text = f"已删除{mode_label} {len(selected_pages)} 页，保留 {output_page_count} 页。"
    else:
        summary_text = f"已将 {len(selected_pages)} 页旋转 {options['angle']}°。"
    return {
        "output_pdf": _output_web_path(output_path),
        "output_filename": output_path.name,
        "selected_pages": [page + 1 for page in selected_pages],
        "processed_page_count": len(selected_pages),
        "page_mode": page_mode,
        "total_pages": output_page_count,
        "output_size": output_path.stat().st_size,
        "angle": options.get("angle"),
        "summary_text": summary_text,
    }


def _write_page_subset(reader: PdfReader, pages: Iterable[int], output_path: Path, *, title: str) -> None:
    page_list = list(pages)
    writer = PdfWriter()
    try:
        for page_index in page_list:
            writer.add_page(reader.pages[page_index])
        writer.add_metadata({"/Title": title, "/Creator": "信实文档处理工作台", "/Producer": "pypdf"})
        with output_path.open("wb") as target:
            writer.write(target)
    finally:
        writer.close()
    _verify_output_pdf(output_path, expected_pages=len(page_list))


def _verify_output_pdf(path: Path, *, expected_pages: int) -> int:
    with path.open("rb") as source:
        page_count = len(PdfReader(source, strict=False).pages)
    if page_count != expected_pages:
        raise RuntimeError(f"输出 PDF 页数校验失败：预期 {expected_pages} 页，实际 {page_count} 页")
    return page_count


def _normalize_output_prefix(raw_name: Any) -> str:
    normalized = normalize_output_filename(str(raw_name or "拆分结果"))
    return Path(normalized).stem[:100] or "拆分结果"


def _page_group_label(pages: list[int]) -> str:
    if not pages:
        return "empty"
    page_numbers = [page + 1 for page in pages]
    if page_numbers == list(range(page_numbers[0], page_numbers[-1] + 1)):
        return f"p{page_numbers[0]}-{page_numbers[-1]}" if len(page_numbers) > 1 else f"p{page_numbers[0]}"
    return "p" + "-".join(str(page) for page in page_numbers[:12])


def _output_web_path(path: Path) -> str:
    output_root = Path(settings.OUTPUT_DIR).resolve()
    resolved = path.resolve()
    try:
        return f"outputs/{resolved.relative_to(output_root).as_posix()}"
    except ValueError:
        return str(resolved)


def _is_relative_to(path: Path, root: Path) -> bool:
    try:
        path.relative_to(root)
        return True
    except ValueError:
        return False


async def _report(callback: Optional[ProgressCallback], progress: int, message: str) -> None:
    if callback is not None:
        await callback(progress, message)
