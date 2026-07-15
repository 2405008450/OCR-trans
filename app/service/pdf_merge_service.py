import asyncio
import os
import re
import shutil
from concurrent.futures import Executor
from contextlib import ExitStack
from datetime import datetime
from pathlib import Path, PurePosixPath
from typing import Any, Awaitable, Callable, Iterable, Optional

from pypdf import PdfReader, PdfWriter

from app.core.config import settings
from app.service.word_count_service import (
    get_word_count_config,
    resolve_allowed_shared_input_path,
)


PDF_EXTENSION = ".pdf"
COPY_CHUNK_SIZE = 1024 * 1024
INVALID_OUTPUT_NAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
NATURAL_SORT_PARTS = re.compile(r"(\d+)")


def get_pdf_merge_config() -> dict[str, Any]:
    """返回 PDF 合并页面需要的共享路径和任务限制。"""
    shared_config = get_word_count_config()
    return {
        "allowed_roots": shared_config.get("allowed_roots", []),
        "unc_mount_mappings": shared_config.get("unc_mount_mappings", []),
        "allow_local_paths": shared_config.get("allow_local_paths", False),
        "max_files": max(2, int(settings.PDF_MERGE_MAX_FILES or 200)),
        "max_file_mb": max(1, int(settings.PDF_MERGE_MAX_FILE_MB or 500)),
        "max_total_mb": max(1, int(settings.PDF_MERGE_MAX_TOTAL_MB or 2048)),
        "output_location": "项目本地 outputs/pdf_merge，完成后可在网页下载",
    }


def discover_pdf_files(*, directory_path: str, recursive: bool = True) -> dict[str, Any]:
    """扫描一个已授权的共享目录，返回可供用户排序选择的 PDF 相对路径。"""
    root, _, input_kind = resolve_allowed_shared_input_path(directory_path)
    if input_kind != "directory":
        raise ValueError(f"请输入包含多个 PDF 的目录路径: {root}")

    max_files = max(2, int(settings.PDF_MERGE_MAX_FILES or 200))
    candidates: list[dict[str, Any]] = []
    truncated = False

    for current_dir, dir_names, file_names in os.walk(root, followlinks=False):
        current = Path(current_dir)
        dir_names[:] = sorted(
            [name for name in dir_names if not _is_hidden_or_temporary(name)],
            key=_natural_sort_key,
        )
        for filename in sorted(file_names, key=_natural_sort_key):
            if _is_hidden_or_temporary(filename) or Path(filename).suffix.lower() != PDF_EXTENSION:
                continue
            candidate = (current / filename).resolve(strict=False)
            if not _is_relative_to(candidate, root) or not candidate.is_file():
                continue
            try:
                stat = candidate.stat()
            except OSError:
                continue
            relative_path = candidate.relative_to(root).as_posix()
            page_count = _read_pdf_page_count(candidate)
            candidates.append(
                {
                    "relative_path": relative_path,
                    "name": candidate.name,
                    "size": stat.st_size,
                    "size_mb": round(stat.st_size / (1024 * 1024), 2),
                    "page_count": page_count,
                    "modified_at": datetime.fromtimestamp(stat.st_mtime).isoformat(timespec="seconds"),
                }
            )
            if len(candidates) > max_files:
                truncated = True
                break
        if truncated or not recursive:
            break

    candidates.sort(key=lambda item: _natural_sort_key(item["relative_path"]))
    visible = candidates[:max_files]
    return {
        "directory_path": str(directory_path).strip(),
        "recursive": bool(recursive),
        "files": visible,
        "file_count": len(visible),
        "total_size": sum(int(item["size"]) for item in visible),
        "truncated": truncated,
        "max_files": max_files,
    }


def _read_pdf_page_count(path: Path) -> Optional[int]:
    """读取扫描列表所需页数；单个异常文件不应中断整个目录扫描。"""
    try:
        with path.open("rb") as source:
            reader = PdfReader(source, strict=False)
            if reader.is_encrypted and reader.decrypt("") == 0:
                return None
            page_count = len(reader.pages)
            return page_count if page_count > 0 else None
    except Exception:
        return None


def prepare_pdf_merge_request(
    *,
    directory_path: str,
    relative_paths: Iterable[str],
    output_filename: str,
) -> dict[str, Any]:
    """校验一次提交，并构造任务队列所需的数据。"""
    normalized_output_name = normalize_output_filename(output_filename)
    normalized_relative_paths = _normalize_relative_paths(relative_paths)
    _resolve_selected_pdf_files(directory_path, normalized_relative_paths)

    return {
        "filename": normalized_output_name,
        "params": {
            "output_filename": normalized_output_name,
            "file_count": len(normalized_relative_paths),
            "directory_path": str(directory_path or "").strip(),
            "relative_paths": normalized_relative_paths,
        },
        "input_files": {
            "directory_path": str(directory_path or "").strip(),
            "relative_paths": normalized_relative_paths,
        },
    }


async def execute_pdf_merge_task(
    *,
    task_id: str,
    display_no: Optional[str],
    directory_path: str,
    relative_paths: Iterable[str],
    output_filename: str,
    progress_callback: Optional[Callable[[int, str], Awaitable[None]]] = None,
    executor: Optional[Executor] = None,
) -> dict[str, Any]:
    """从只读共享目录暂存 PDF，在本地完成合并并返回下载路径。"""
    loop = asyncio.get_running_loop()
    await _report(progress_callback, 4, "正在校验共享目录和所选 PDF...")
    selected_files = await loop.run_in_executor(
        executor,
        lambda: _resolve_selected_pdf_files(directory_path, relative_paths),
    )

    task_folder = display_no or task_id
    staging_dir = Path(settings.UPLOAD_DIR) / "pdf_merge_staging" / task_folder
    output_dir = Path(settings.OUTPUT_DIR) / "pdf_merge" / task_folder
    output_dir.mkdir(parents=True, exist_ok=True)
    staging_dir.mkdir(parents=True, exist_ok=True)
    staged_files: list[dict[str, Any]] = []
    staged_total_bytes = 0
    max_total_bytes = max(1, int(settings.PDF_MERGE_MAX_TOTAL_MB or 2048)) * 1024 * 1024

    try:
        total_files = len(selected_files)
        for index, item in enumerate(selected_files, start=1):
            progress = 8 + int((index - 1) / max(total_files, 1) * 54)
            await _report(
                progress_callback,
                progress,
                f"正在暂存 {index}/{total_files}: {item['relative_path']}",
            )
            staged = await loop.run_in_executor(
                executor,
                lambda current=item, order=index: _stage_pdf_file(current, staging_dir, order),
            )
            staged_total_bytes += int(staged["size"])
            if staged_total_bytes > max_total_bytes:
                raise ValueError(
                    f"暂存后的 PDF 总大小超过 {settings.PDF_MERGE_MAX_TOTAL_MB} MB 限制"
                )
            staged_files.append(staged)

        await _report(progress_callback, 66, "正在合并 PDF 页面和书签...")
        output_name = normalize_output_filename(output_filename)
        output_path = output_dir / output_name
        merge_result = await loop.run_in_executor(
            executor,
            lambda: _merge_staged_pdfs(staged_files, output_path),
        )
        await _report(progress_callback, 96, "正在校验合并结果...")

        return {
            "output_pdf": _output_web_path(output_path),
            "output_filename": output_name,
            "input_file_count": total_files,
            "total_pages": merge_result["total_pages"],
            "input_total_size": sum(int(item["size"]) for item in staged_files),
            "output_size": output_path.stat().st_size,
            "files": merge_result["files"],
            "summary_text": f"已按指定顺序合并 {total_files} 个 PDF，共 {merge_result['total_pages']} 页。",
        }
    finally:
        shutil.rmtree(staging_dir, ignore_errors=True)


def normalize_output_filename(raw_name: str) -> str:
    """生成适合 Windows 和 Linux 的安全 PDF 输出文件名。"""
    basename = str(raw_name or "").strip().replace("\\", "/").rsplit("/", 1)[-1]
    basename = INVALID_OUTPUT_NAME_CHARS.sub("_", basename).strip(" .")
    if not basename:
        basename = "合并结果.pdf"
    if not basename.lower().endswith(PDF_EXTENSION):
        basename = f"{basename}{PDF_EXTENSION}"
    stem = Path(basename).stem.strip(" .") or "合并结果"
    return f"{stem[:120]}{PDF_EXTENSION}"


def _normalize_relative_paths(relative_paths: Iterable[str]) -> list[str]:
    normalized: list[str] = []
    seen: set[str] = set()
    for raw_path in relative_paths or []:
        raw_text = str(raw_path or "").strip().replace("\\", "/")
        pure_path = PurePosixPath(raw_text)
        if not raw_text or pure_path.is_absolute() or any(part in {"", ".", ".."} for part in pure_path.parts):
            raise ValueError(f"PDF 相对路径无效: {raw_path}")
        clean_path = pure_path.as_posix()
        compare_key = clean_path.casefold()
        if compare_key in seen:
            raise ValueError(f"不能重复选择同一个 PDF: {clean_path}")
        seen.add(compare_key)
        normalized.append(clean_path)

    max_files = max(2, int(settings.PDF_MERGE_MAX_FILES or 200))
    if len(normalized) < 2:
        raise ValueError("请至少选择 2 个 PDF 文件")
    if len(normalized) > max_files:
        raise ValueError(f"一次最多合并 {max_files} 个 PDF 文件")
    return normalized


def _resolve_selected_pdf_files(directory_path: str, relative_paths: Iterable[str]) -> list[dict[str, Any]]:
    root, _, input_kind = resolve_allowed_shared_input_path(directory_path)
    if input_kind != "directory":
        raise ValueError(f"请输入包含多个 PDF 的目录路径: {root}")

    normalized_paths = _normalize_relative_paths(relative_paths)
    max_file_bytes = max(1, int(settings.PDF_MERGE_MAX_FILE_MB or 500)) * 1024 * 1024
    max_total_bytes = max(1, int(settings.PDF_MERGE_MAX_TOTAL_MB or 2048)) * 1024 * 1024
    total_bytes = 0
    selected: list[dict[str, Any]] = []

    for relative_path in normalized_paths:
        pure_path = PurePosixPath(relative_path)
        candidate = root.joinpath(*pure_path.parts).resolve(strict=False)
        if not _is_relative_to(candidate, root):
            raise PermissionError(f"PDF 路径超出所选共享目录: {relative_path}")
        if candidate.suffix.lower() != PDF_EXTENSION:
            raise ValueError(f"只支持合并 PDF 文件: {relative_path}")
        if not candidate.exists() or not candidate.is_file():
            raise FileNotFoundError(f"PDF 文件不存在或暂时不可访问: {relative_path}")
        try:
            size = candidate.stat().st_size
            with candidate.open("rb") as source:
                signature = source.read(1024)
        except OSError as exc:
            raise FileNotFoundError(f"无法读取 PDF 文件: {relative_path}（{exc}）") from exc
        if b"%PDF-" not in signature:
            raise ValueError(f"文件扩展名为 PDF，但文件头无效: {relative_path}")
        if size > max_file_bytes:
            raise ValueError(
                f"单个 PDF 超过 {settings.PDF_MERGE_MAX_FILE_MB} MB 限制: {relative_path}"
            )
        total_bytes += size
        if total_bytes > max_total_bytes:
            raise ValueError(f"所选 PDF 总大小超过 {settings.PDF_MERGE_MAX_TOTAL_MB} MB 限制")
        selected.append({"path": candidate, "relative_path": relative_path, "size": size})
    return selected


def _stage_pdf_file(item: dict[str, Any], staging_dir: Path, index: int) -> dict[str, Any]:
    source_path = Path(item["path"])
    safe_name = INVALID_OUTPUT_NAME_CHARS.sub("_", source_path.name).strip(" .") or f"input_{index}.pdf"
    staged_path = staging_dir / f"{index:04d}_{safe_name}"
    partial_path = staged_path.with_suffix(f"{staged_path.suffix}.partial")
    try:
        with source_path.open("rb") as source, partial_path.open("wb") as target:
            shutil.copyfileobj(source, target, length=COPY_CHUNK_SIZE)
        partial_path.replace(staged_path)
    except OSError as exc:
        partial_path.unlink(missing_ok=True)
        raise OSError(f"暂存共享目录文件失败: {item['relative_path']}（{exc}）") from exc

    actual_size = staged_path.stat().st_size
    if actual_size <= 0:
        raise ValueError(f"PDF 文件为空: {item['relative_path']}")
    max_file_bytes = max(1, int(settings.PDF_MERGE_MAX_FILE_MB or 500)) * 1024 * 1024
    if actual_size > max_file_bytes:
        raise ValueError(
            f"暂存后发现单个 PDF 超过 {settings.PDF_MERGE_MAX_FILE_MB} MB 限制: {item['relative_path']}"
        )
    return {
        "path": staged_path,
        "relative_path": item["relative_path"],
        "size": actual_size,
    }


def _merge_staged_pdfs(staged_files: list[dict[str, Any]], output_path: Path) -> dict[str, Any]:
    writer = PdfWriter()
    partial_path = output_path.with_suffix(".partial.pdf")
    total_pages = 0
    file_results: list[dict[str, Any]] = []

    try:
        with ExitStack() as stack:
            for item in staged_files:
                handle = stack.enter_context(Path(item["path"]).open("rb"))
                try:
                    reader = PdfReader(handle, strict=False)
                    if reader.is_encrypted and reader.decrypt("") == 0:
                        raise ValueError("文件已加密且需要密码")
                    page_count = len(reader.pages)
                    if page_count <= 0:
                        raise ValueError("文件不包含可合并页面")
                    writer.append(reader, import_outline=True)
                except Exception as exc:
                    raise ValueError(f"无法合并 {item['relative_path']}: {exc}") from exc
                total_pages += page_count
                file_results.append(
                    {
                        "relative_path": item["relative_path"],
                        "filename": Path(item["relative_path"]).name,
                        "size": item["size"],
                        "page_count": page_count,
                    }
                )

            writer.add_metadata(
                {
                    "/Title": output_path.stem,
                    "/Creator": "信实文档处理工作台",
                    "/Producer": "pypdf",
                }
            )
            with partial_path.open("wb") as output_stream:
                writer.write(output_stream)

        with partial_path.open("rb") as merged_stream:
            verified_reader = PdfReader(merged_stream, strict=False)
            verified_pages = len(verified_reader.pages)
        if verified_pages != total_pages:
            raise RuntimeError(f"合并结果页数校验失败：预期 {total_pages} 页，实际 {verified_pages} 页")
        partial_path.replace(output_path)
        return {"total_pages": total_pages, "files": file_results}
    finally:
        writer.close()
        partial_path.unlink(missing_ok=True)


def _output_web_path(path: Path) -> str:
    output_root = Path(settings.OUTPUT_DIR).resolve()
    resolved = path.resolve()
    try:
        return f"outputs/{resolved.relative_to(output_root).as_posix()}"
    except ValueError:
        return str(resolved)


async def _report(
    callback: Optional[Callable[[int, str], Awaitable[None]]],
    progress: int,
    message: str,
) -> None:
    if callback is not None:
        await callback(progress, message)


def _natural_sort_key(value: str) -> tuple[Any, ...]:
    return tuple(
        (0, int(part)) if part.isdigit() else (1, part.casefold())
        for part in NATURAL_SORT_PARTS.split(str(value).replace("\\", "/"))
        if part != ""
    )


def _is_hidden_or_temporary(name: str) -> bool:
    return name.startswith((".", "~$", "$"))


def _is_relative_to(path: Path, root: Path) -> bool:
    try:
        path.relative_to(root)
        return True
    except ValueError:
        return False
