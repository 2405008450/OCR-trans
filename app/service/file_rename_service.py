from __future__ import annotations

import asyncio
import os
import re
import shutil
from concurrent.futures import Executor
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Awaitable, Callable, Iterable, Literal, Optional

from app.core.config import settings
from app.service.word_count_service import (
    get_word_count_config,
    resolve_allowed_shared_input_path,
)


FILE_RENAME_MODE_NUMBERING = "numbering"
FILE_RENAME_MODE_CLEANUP = "cleanup"
FILE_RENAME_MODE_REGEX = "regex"
FILE_RENAME_MODES = {
    FILE_RENAME_MODE_CLEANUP,
    FILE_RENAME_MODE_NUMBERING,
    FILE_RENAME_MODE_REGEX,
}
FILE_RENAME_COPY_DIR_PREFIX = "文件改名副本_"
COMMON_SYSTEM_FILE_NAMES = {".ds_store", "desktop.ini", "thumbs.db"}
INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
COMPACT_DATE_TIME_PREFIX_PATTERN = re.compile(r"^\d{8}-\d{6}_")
DOTTED_DATE_PREFIX_PATTERN = re.compile(r"^\d{4}\.\d{1,2}\.\d{1,2}_")
WINDOWS_RESERVED_NAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{index}" for index in range(1, 10)),
    *(f"LPT{index}" for index in range(1, 10)),
}


@dataclass(frozen=True)
class RenameCopyOperation:
    source_path: Path
    source_relative_path: str
    target_relative_path: str
    size: int


@dataclass(frozen=True)
class RenameCopyPlan:
    root: Path
    operations: list[RenameCopyOperation]
    preview: dict[str, Any]


def get_file_rename_config() -> dict[str, Any]:
    """复用字数统计的共享路径白名单和 UNC 映射配置。"""
    shared_config = get_word_count_config()
    return {
        "allowed_roots": shared_config.get("allowed_roots", []),
        "unc_mount_mappings": shared_config.get("unc_mount_mappings", []),
        "allow_local_paths": shared_config.get("allow_local_paths", False),
        "max_files": _max_files(),
        "copy_directory_prefix": FILE_RENAME_COPY_DIR_PREFIX,
        "modes": [
            FILE_RENAME_MODE_CLEANUP,
            FILE_RENAME_MODE_NUMBERING,
            FILE_RENAME_MODE_REGEX,
        ],
        "source_policy": "源文件只读；仅把实际需要处理的文件复制到新目录。",
    }


def discover_file_rename_files(
    *,
    directory_path: str,
    recursive: bool = True,
    include_hidden: bool = False,
) -> dict[str, Any]:
    """扫描已授权目录中的普通文件，返回可勾选的候选项。"""
    root = _resolve_directory(directory_path)
    max_files = _max_files()
    files: list[dict[str, Any]] = []
    truncated = False

    for candidate in _iter_candidate_files(root, recursive=recursive, include_hidden=include_hidden):
        if len(files) >= max_files:
            truncated = True
            break
        try:
            stat = candidate.stat()
        except OSError:
            continue
        files.append(
            {
                "relative_path": candidate.relative_to(root).as_posix(),
                "name": candidate.name,
                "extension": candidate.suffix.lower(),
                "size": int(stat.st_size),
                "modified_at": datetime.fromtimestamp(stat.st_mtime).isoformat(timespec="seconds"),
            }
        )

    files.sort(key=lambda item: item["relative_path"].casefold())
    return {
        "directory_path": str(directory_path or "").strip(),
        "recursive": bool(recursive),
        "include_hidden": bool(include_hidden),
        "files": files,
        "file_count": len(files),
        "total_size": sum(int(item["size"]) for item in files),
        "truncated": truncated,
        "max_files": max_files,
    }


def build_file_rename_preview(
    *,
    directory_path: str,
    relative_paths: Iterable[str],
    mode: str,
    recursive: bool = True,
    include_hidden: bool = False,
    regex_pattern: str = "",
    replacement: str = "",
    ignore_case: bool = False,
    cleanup_remove_leading_number: bool = True,
    cleanup_leading_number_max_digits: int = 6,
    cleanup_leading_number_space: bool = True,
    cleanup_leading_number_underscore: bool = True,
    cleanup_remove_datetime: bool = True,
    cleanup_datetime_compact: bool = True,
    cleanup_datetime_dotted: bool = True,
    cleanup_remove_translated: bool = True,
    cleanup_translated_suffix: str = "_translated",
) -> dict[str, Any]:
    return _build_file_rename_plan(
        directory_path=directory_path,
        relative_paths=relative_paths,
        mode=mode,
        recursive=recursive,
        include_hidden=include_hidden,
        regex_pattern=regex_pattern,
        replacement=replacement,
        ignore_case=ignore_case,
        cleanup_remove_leading_number=cleanup_remove_leading_number,
        cleanup_leading_number_max_digits=cleanup_leading_number_max_digits,
        cleanup_leading_number_space=cleanup_leading_number_space,
        cleanup_leading_number_underscore=cleanup_leading_number_underscore,
        cleanup_remove_datetime=cleanup_remove_datetime,
        cleanup_datetime_compact=cleanup_datetime_compact,
        cleanup_datetime_dotted=cleanup_datetime_dotted,
        cleanup_remove_translated=cleanup_remove_translated,
        cleanup_translated_suffix=cleanup_translated_suffix,
    ).preview


def prepare_file_rename_request(
    *,
    directory_path: str,
    relative_paths: Iterable[str],
    mode: str,
    recursive: bool = True,
    include_hidden: bool = False,
    regex_pattern: str = "",
    replacement: str = "",
    ignore_case: bool = False,
    cleanup_remove_leading_number: bool = True,
    cleanup_leading_number_max_digits: int = 6,
    cleanup_leading_number_space: bool = True,
    cleanup_leading_number_underscore: bool = True,
    cleanup_remove_datetime: bool = True,
    cleanup_datetime_compact: bool = True,
    cleanup_datetime_dotted: bool = True,
    cleanup_remove_translated: bool = True,
    cleanup_translated_suffix: str = "_translated",
) -> dict[str, Any]:
    plan = _build_file_rename_plan(
        directory_path=directory_path,
        relative_paths=relative_paths,
        mode=mode,
        recursive=recursive,
        include_hidden=include_hidden,
        regex_pattern=regex_pattern,
        replacement=replacement,
        ignore_case=ignore_case,
        cleanup_remove_leading_number=cleanup_remove_leading_number,
        cleanup_leading_number_max_digits=cleanup_leading_number_max_digits,
        cleanup_leading_number_space=cleanup_leading_number_space,
        cleanup_leading_number_underscore=cleanup_leading_number_underscore,
        cleanup_remove_datetime=cleanup_remove_datetime,
        cleanup_datetime_compact=cleanup_datetime_compact,
        cleanup_datetime_dotted=cleanup_datetime_dotted,
        cleanup_remove_translated=cleanup_remove_translated,
        cleanup_translated_suffix=cleanup_translated_suffix,
    )
    preview = plan.preview
    if preview["conflict_count"] or preview["invalid_count"]:
        raise ValueError("预览中仍有重名冲突或无效文件名，请调整规则或文件选择")
    if not plan.operations:
        raise ValueError("当前规则没有产生需要复制的改名文件")

    normalized_mode = _normalize_mode(mode)
    params = {
        "directory_path": str(directory_path or "").strip(),
        "relative_paths": [operation.source_relative_path for operation in plan.operations],
        "mode": normalized_mode,
        "recursive": bool(recursive),
        "include_hidden": bool(include_hidden),
        "regex_pattern": str(regex_pattern or ""),
        "replacement": str(replacement or ""),
        "ignore_case": bool(ignore_case),
        "cleanup_remove_leading_number": bool(cleanup_remove_leading_number),
        "cleanup_leading_number_max_digits": int(cleanup_leading_number_max_digits),
        "cleanup_leading_number_space": bool(cleanup_leading_number_space),
        "cleanup_leading_number_underscore": bool(cleanup_leading_number_underscore),
        "cleanup_remove_datetime": bool(cleanup_remove_datetime),
        "cleanup_datetime_compact": bool(cleanup_datetime_compact),
        "cleanup_datetime_dotted": bool(cleanup_datetime_dotted),
        "cleanup_remove_translated": bool(cleanup_remove_translated),
        "cleanup_translated_suffix": str(cleanup_translated_suffix or ""),
        "selected_count": preview["selected_count"],
        "process_count": preview["process_count"],
        "estimated_bytes": preview["estimated_bytes"],
    }
    return {
        "filename": f"{plan.root.name or plan.root} · {_mode_label(normalized_mode)}",
        "params": params,
        "input_files": {
            "directory_path": str(directory_path or "").strip(),
            "relative_paths": [operation.source_relative_path for operation in plan.operations],
        },
    }


async def execute_file_rename_copy_task(
    *,
    task_id: str,
    display_no: Optional[str],
    directory_path: str,
    relative_paths: Iterable[str],
    mode: str,
    recursive: bool = True,
    include_hidden: bool = False,
    regex_pattern: str = "",
    replacement: str = "",
    ignore_case: bool = False,
    cleanup_remove_leading_number: bool = True,
    cleanup_leading_number_max_digits: int = 6,
    cleanup_leading_number_space: bool = True,
    cleanup_leading_number_underscore: bool = True,
    cleanup_remove_datetime: bool = True,
    cleanup_datetime_compact: bool = True,
    cleanup_datetime_dotted: bool = True,
    cleanup_remove_translated: bool = True,
    cleanup_translated_suffix: str = "_translated",
    progress_callback: Optional[Callable[[int, str], Awaitable[None]]] = None,
    executor: Optional[Executor] = None,
) -> dict[str, Any]:
    """在源目录内创建可见副本目录，逐个复制并改名，绝不移动源文件。"""
    loop = asyncio.get_running_loop()
    plan = await loop.run_in_executor(
        executor,
        lambda: _build_file_rename_plan(
            directory_path=directory_path,
            relative_paths=relative_paths,
            mode=mode,
            recursive=recursive,
            include_hidden=include_hidden,
            regex_pattern=regex_pattern,
            replacement=replacement,
            ignore_case=ignore_case,
            cleanup_remove_leading_number=cleanup_remove_leading_number,
            cleanup_leading_number_max_digits=cleanup_leading_number_max_digits,
            cleanup_leading_number_space=cleanup_leading_number_space,
            cleanup_leading_number_underscore=cleanup_leading_number_underscore,
            cleanup_remove_datetime=cleanup_remove_datetime,
            cleanup_datetime_compact=cleanup_datetime_compact,
            cleanup_datetime_dotted=cleanup_datetime_dotted,
            cleanup_remove_translated=cleanup_remove_translated,
            cleanup_translated_suffix=cleanup_translated_suffix,
        ),
    )
    if plan.preview["conflict_count"] or plan.preview["invalid_count"]:
        raise ValueError("执行前重新校验失败：存在重名冲突或无效文件名")
    if not plan.operations:
        raise ValueError("没有需要复制的改名文件")

    _validate_available_space(plan.root, plan.preview["estimated_bytes"])
    output_directory = _unique_output_directory(plan.root, display_no or task_id[:8])
    output_directory.mkdir(parents=False, exist_ok=False)
    copied_files: list[dict[str, str]] = []
    failed_files: list[dict[str, str]] = []
    copied_bytes = 0
    try:
        await _report(progress_callback, 3, f"准备生成副本目录：{output_directory.name}")
        total = len(plan.operations)
        for index, operation in enumerate(plan.operations, start=1):
            progress = 5 + int((index - 1) / max(total, 1) * 90)
            await _report(
                progress_callback,
                progress,
                f"正在复制 {index}/{total}: {operation.source_relative_path}",
            )
            target_path = output_directory.joinpath(*operation.target_relative_path.split("/"))
            try:
                target_path.parent.mkdir(parents=True, exist_ok=True)
                await loop.run_in_executor(
                    executor,
                    lambda source=operation.source_path, target=target_path: shutil.copy2(source, target),
                )
                copied_bytes += operation.size
                copied_files.append(
                    {
                        "source_relative_path": operation.source_relative_path,
                        "target_relative_path": operation.target_relative_path,
                    }
                )
            except OSError as exc:
                target_path.unlink(missing_ok=True)
                failed_files.append(
                    {
                        "source_relative_path": operation.source_relative_path,
                        "target_relative_path": operation.target_relative_path,
                        "error": str(exc),
                    }
                )

        if not copied_files:
            first_error = failed_files[0]["error"] if failed_files else "未知复制错误"
            raise OSError(f"所有文件复制失败：{first_error}")

        message = f"副本生成完成：成功 {len(copied_files)} 个"
        if failed_files:
            message += f"，失败 {len(failed_files)} 个"
        await _report(progress_callback, 100, message)
    except BaseException:
        shutil.rmtree(output_directory, ignore_errors=True)
        raise

    return {
        "mode": _normalize_mode(mode),
        "source_directory": str(plan.root),
        "output_directory": str(output_directory),
        "selected_count": plan.preview["selected_count"],
        "planned_count": len(plan.operations),
        "copied_count": len(copied_files),
        "failed_count": len(failed_files),
        "copied_bytes": copied_bytes,
        "copied_files": copied_files,
        "failed_files": failed_files,
    }


def _build_file_rename_plan(
    *,
    directory_path: str,
    relative_paths: Iterable[str],
    mode: str,
    recursive: bool,
    include_hidden: bool,
    regex_pattern: str,
    replacement: str,
    ignore_case: bool,
    cleanup_remove_leading_number: bool,
    cleanup_leading_number_max_digits: int,
    cleanup_leading_number_space: bool,
    cleanup_leading_number_underscore: bool,
    cleanup_remove_datetime: bool,
    cleanup_datetime_compact: bool,
    cleanup_datetime_dotted: bool,
    cleanup_remove_translated: bool,
    cleanup_translated_suffix: str,
) -> RenameCopyPlan:
    normalized_mode = _normalize_mode(mode)
    root = _resolve_directory(directory_path)
    selected_files = _resolve_selected_files(
        root,
        relative_paths,
        recursive=recursive,
        include_hidden=include_hidden,
    )
    selected_files.sort(key=lambda item: item.relative_to(root).as_posix().casefold())

    compiled_pattern: Optional[re.Pattern[str]] = None
    replacement_text = str(replacement or "")
    if normalized_mode == FILE_RENAME_MODE_REGEX:
        pattern_text = str(regex_pattern or "")
        if not pattern_text:
            raise ValueError("请输入正则表达式")
        if len(pattern_text) > 500:
            raise ValueError("正则表达式不能超过 500 个字符")
        try:
            compiled_pattern = re.compile(pattern_text, re.IGNORECASE if ignore_case else 0)
            compiled_pattern.sub(replacement_text, "正则校验")
        except re.error as exc:
            raise ValueError(f"正则表达式或替换内容无效：{exc}") from exc

    width = len(str(len(selected_files))) if selected_files else 0
    entries: list[dict[str, Any]] = []
    operation_by_source: dict[str, RenameCopyOperation] = {}

    for index, source_path in enumerate(selected_files, start=1):
        source_relative = source_path.relative_to(root).as_posix()
        target_name = ""
        status = "ready"
        reason = ""

        if normalized_mode == FILE_RENAME_MODE_NUMBERING:
            target_name = f"{str(index).zfill(width)}_{source_path.name}"
        elif normalized_mode == FILE_RENAME_MODE_CLEANUP:
            source_stem = source_path.stem
            target_stem = _clean_common_filename_stem(
                source_stem,
                remove_leading_number=cleanup_remove_leading_number,
                leading_number_max_digits=cleanup_leading_number_max_digits,
                leading_number_space=cleanup_leading_number_space,
                leading_number_underscore=cleanup_leading_number_underscore,
                remove_datetime=cleanup_remove_datetime,
                datetime_compact=cleanup_datetime_compact,
                datetime_dotted=cleanup_datetime_dotted,
                remove_translated=cleanup_remove_translated,
                translated_suffix=cleanup_translated_suffix,
            )
            target_name = f"{target_stem}{source_path.suffix}"
            if target_name == source_path.name:
                status = "unchanged"
                reason = "所选常用清理规则未改变文件名"
        else:
            assert compiled_pattern is not None
            source_stem = source_path.stem
            match = compiled_pattern.search(source_stem)
            if match is None:
                status = "unmatched"
                reason = "文件主名不匹配正则表达式"
            else:
                try:
                    target_stem = compiled_pattern.sub(replacement_text, source_stem)
                except re.error as exc:
                    raise ValueError(f"替换内容无效：{exc}") from exc
                target_name = f"{target_stem}{source_path.suffix}"
                if target_name == source_path.name:
                    status = "unchanged"
                    reason = "替换后文件名没有变化"

        if status == "ready":
            filename_error = _validate_filename(target_name)
            if filename_error:
                status = "invalid"
                reason = filename_error

        target_relative = ""
        if target_name:
            relative_parent = Path(source_relative).parent
            target_relative = (
                target_name
                if str(relative_parent) == "."
                else (relative_parent / target_name).as_posix()
            )

        entry = {
            "source_relative_path": source_relative,
            "target_relative_path": target_relative,
            "source_name": source_path.name,
            "target_name": target_name,
            "size": int(source_path.stat().st_size),
            "status": status,
            "reason": reason,
        }
        entries.append(entry)
        if status == "ready":
            operation_by_source[source_relative] = RenameCopyOperation(
                source_path=source_path,
                source_relative_path=source_relative,
                target_relative_path=target_relative,
                size=entry["size"],
            )

    target_sources: dict[str, list[str]] = {}
    for source_relative, operation in operation_by_source.items():
        target_sources.setdefault(operation.target_relative_path.casefold(), []).append(source_relative)
    conflicting_sources = {
        source_relative
        for sources in target_sources.values()
        if len(sources) > 1
        for source_relative in sources
    }
    if conflicting_sources:
        for entry in entries:
            if entry["source_relative_path"] in conflicting_sources:
                entry["status"] = "conflict"
                entry["reason"] = "多个文件会生成相同的目标路径"
                operation_by_source.pop(entry["source_relative_path"], None)

    operations = [
        operation_by_source[entry["source_relative_path"]]
        for entry in entries
        if entry["source_relative_path"] in operation_by_source
    ]
    status_counts = {
        status: sum(1 for entry in entries if entry["status"] == status)
        for status in ("ready", "unmatched", "unchanged", "invalid", "conflict")
    }
    preview = {
        "directory_path": str(directory_path or "").strip(),
        "mode": normalized_mode,
        "selected_count": len(selected_files),
        "process_count": len(operations),
        "skipped_count": status_counts["unmatched"] + status_counts["unchanged"],
        "invalid_count": status_counts["invalid"],
        "conflict_count": status_counts["conflict"],
        "estimated_bytes": sum(operation.size for operation in operations),
        "number_width": width if normalized_mode == FILE_RENAME_MODE_NUMBERING else 0,
        "copy_directory_example": f"{FILE_RENAME_COPY_DIR_PREFIX}任务编号_时间",
        "operations": entries,
    }
    return RenameCopyPlan(root=root, operations=operations, preview=preview)


def _resolve_directory(directory_path: str) -> Path:
    root, _, input_kind = resolve_allowed_shared_input_path(directory_path)
    if input_kind != "directory":
        raise ValueError(f"请输入需要批量处理的目录路径：{root}")
    return root.resolve(strict=False)


def _iter_candidate_files(root: Path, *, recursive: bool, include_hidden: bool) -> Iterable[Path]:
    followlinks = settings.WORD_COUNT_FOLLOW_SYMLINKS_ENABLED
    if not recursive:
        for child in sorted(root.iterdir(), key=lambda item: item.name.casefold()):
            if _should_skip_name(child.name, include_hidden=include_hidden, is_directory=child.is_dir()):
                continue
            if child.is_file() and (followlinks or not child.is_symlink()):
                yield child
        return

    for current_dir, dir_names, file_names in os.walk(root, followlinks=followlinks):
        current_path = Path(current_dir)
        dir_names[:] = sorted(
            [
                name
                for name in dir_names
                if not _should_skip_name(name, include_hidden=include_hidden, is_directory=True)
            ],
            key=str.casefold,
        )
        for file_name in sorted(file_names, key=str.casefold):
            if _should_skip_name(file_name, include_hidden=include_hidden, is_directory=False):
                continue
            candidate = current_path / file_name
            if not followlinks and candidate.is_symlink():
                continue
            if candidate.is_file():
                yield candidate


def _should_skip_name(name: str, *, include_hidden: bool, is_directory: bool) -> bool:
    normalized = name.casefold()
    if is_directory and name.startswith(FILE_RENAME_COPY_DIR_PREFIX):
        return True
    if normalized in COMMON_SYSTEM_FILE_NAMES or name.startswith("~$"):
        return not include_hidden
    return not include_hidden and name.startswith(".")


def _resolve_selected_files(
    root: Path,
    relative_paths: Iterable[str],
    *,
    recursive: bool,
    include_hidden: bool,
) -> list[Path]:
    normalized_paths = _normalize_relative_paths(relative_paths)
    followlinks = settings.WORD_COUNT_FOLLOW_SYMLINKS_ENABLED
    resolved: list[Path] = []
    for relative_path in normalized_paths:
        parts = relative_path.split("/")
        if any(part.startswith(FILE_RENAME_COPY_DIR_PREFIX) for part in parts):
            raise ValueError(f"不能处理本工具生成的副本目录：{relative_path}")
        if not include_hidden and any(_should_skip_name(part, include_hidden=False, is_directory=False) for part in parts):
            raise ValueError(f"未启用隐藏文件扫描：{relative_path}")
        if not recursive and len(parts) > 1:
            raise ValueError(f"未启用子目录扫描：{relative_path}")

        unresolved = root.joinpath(*parts)
        if not followlinks:
            current = root
            for part in parts:
                current = current / part
                if current.is_symlink():
                    raise PermissionError(f"当前配置不允许处理符号链接：{relative_path}")
        candidate = unresolved.resolve(strict=False)
        if not _is_relative_to(candidate, root):
            raise PermissionError(f"文件路径超出所选目录：{relative_path}")
        if not candidate.exists() or not candidate.is_file():
            raise FileNotFoundError(f"待处理文件不存在或暂时不可访问：{relative_path}")
        resolved.append(candidate)
    return resolved


def _normalize_relative_paths(relative_paths: Iterable[str]) -> list[str]:
    normalized: list[str] = []
    seen: set[str] = set()
    for raw_path in relative_paths or []:
        raw_text = str(raw_path or "").strip().replace("\\", "/")
        parts = [part for part in raw_text.split("/") if part]
        if (
            not raw_text
            or raw_text.startswith("/")
            or Path(raw_text).is_absolute()
            or any(part in {".", ".."} for part in parts)
        ):
            raise ValueError(f"文件相对路径无效：{raw_path}")
        clean_path = "/".join(parts)
        key = clean_path.casefold()
        if key in seen:
            raise ValueError(f"不能重复选择同一个文件：{clean_path}")
        seen.add(key)
        normalized.append(clean_path)
    if not normalized:
        raise ValueError("请至少勾选一个待处理文件")
    if len(normalized) > _max_files():
        raise ValueError(f"一次最多选择 {_max_files()} 个文件")
    return normalized


def _validate_filename(filename: str) -> str:
    if not filename or filename in {".", ".."}:
        return "替换后的文件名不能为空"
    if INVALID_FILENAME_CHARS.search(filename):
        return "替换后的文件名包含 Windows 不允许的字符"
    if filename.endswith((" ", ".")):
        return "替换后的文件名不能以空格或句点结尾"
    if len(filename) > 255:
        return "替换后的文件名超过 255 个字符"
    reserved_base = filename.split(".", 1)[0].upper()
    if reserved_base in WINDOWS_RESERVED_NAMES:
        return "替换后的文件名是 Windows 保留名称"
    return ""


def _unique_output_directory(root: Path, task_label: str) -> Path:
    safe_task_label = re.sub(r"[^0-9A-Za-z_-]+", "_", task_label).strip("_") or "task"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name = f"{FILE_RENAME_COPY_DIR_PREFIX}{timestamp}_{safe_task_label}"
    for index in range(0, 1000):
        suffix = "" if index == 0 else f"_{index}"
        candidate = root / f"{base_name}{suffix}"
        if not candidate.exists():
            return candidate
    raise OSError("无法创建唯一的副本目录，请稍后重试")


def _validate_available_space(root: Path, required_bytes: int) -> None:
    try:
        free_bytes = shutil.disk_usage(root).free
    except OSError:
        return
    if required_bytes > free_bytes:
        raise OSError(
            f"目标磁盘空间不足：预计需要 {required_bytes} 字节，当前可用 {free_bytes} 字节"
        )


def _normalize_mode(mode: str) -> Literal["cleanup", "numbering", "regex"]:
    normalized = str(mode or "").strip().lower()
    if normalized not in FILE_RENAME_MODES:
        raise ValueError(f"不支持的改名模式：{mode}")
    return normalized  # type: ignore[return-value]


def _mode_label(mode: str) -> str:
    if mode == FILE_RENAME_MODE_CLEANUP:
        return "常用清理"
    return "自动编号" if mode == FILE_RENAME_MODE_NUMBERING else "正则改名"


def _clean_common_filename_stem(
    stem: str,
    *,
    remove_leading_number: bool,
    leading_number_max_digits: int,
    leading_number_space: bool,
    leading_number_underscore: bool,
    remove_datetime: bool,
    datetime_compact: bool,
    datetime_dotted: bool,
    remove_translated: bool,
    translated_suffix: str,
) -> str:
    result = stem
    if remove_leading_number:
        max_digits = int(leading_number_max_digits)
        if max_digits < 1 or max_digits > 12:
            raise ValueError("开头序号最大位数必须在 1 到 12 之间")
        separators = []
        if leading_number_space:
            separators.append(r"\s+")
        if leading_number_underscore:
            separators.append(r"_+")
        if separators:
            leading_pattern = re.compile(rf"^\d{{1,{max_digits}}}(?:{'|'.join(separators)})")
            result = leading_pattern.sub("", result, count=1)
    if remove_datetime:
        if datetime_compact:
            result = COMPACT_DATE_TIME_PREFIX_PATTERN.sub("", result, count=1)
        if datetime_dotted:
            result = DOTTED_DATE_PREFIX_PATTERN.sub("", result, count=1)
    if remove_translated:
        suffix = str(translated_suffix or "")
        if len(suffix) > 80:
            raise ValueError("结尾标记不能超过 80 个字符")
        if suffix:
            result = re.sub(rf"{re.escape(suffix)}$", "", result, count=1, flags=re.IGNORECASE)
    return result


def _max_files() -> int:
    return max(1, int(settings.WORD_COUNT_MAX_FILES or 5000))


def _is_relative_to(candidate: Path, root: Path) -> bool:
    try:
        candidate.resolve(strict=False).relative_to(root.resolve(strict=False))
        return True
    except ValueError:
        return False


async def _report(
    callback: Optional[Callable[[int, str], Awaitable[None]]],
    progress: int,
    message: str,
) -> None:
    if callback:
        await callback(progress, message)
