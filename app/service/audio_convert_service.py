"""共享路径与上传音频的 M4A 转 WAV 服务。"""
from __future__ import annotations

import asyncio
import os
import re
import shutil
import subprocess
import zipfile
from concurrent.futures import Executor
from pathlib import Path, PurePosixPath
from typing import Any

from app.core.config import settings
from app.service.word_count_service import get_word_count_config, resolve_allowed_shared_input_path


class AudioConvertError(RuntimeError):
    def __init__(self, message: str, result: dict | None = None):
        super().__init__(message)
        self.result = result


def limits() -> dict[str, int]:
    return {
        "max_files": max(1, settings.AUDIO_CONVERT_MAX_FILES),
        "max_file_mb": max(1, settings.AUDIO_CONVERT_MAX_FILE_MB),
        "max_total_mb": max(1, settings.AUDIO_CONVERT_MAX_TOTAL_MB),
    }


def ffmpeg_binary() -> str:
    directory = settings.FFMPEG_BIN_DIR.strip()
    executable = shutil.which("ffmpeg", path=directory) if directory else shutil.which("ffmpeg")
    if not executable:
        raise AudioConvertError("找不到 FFmpeg，请安装并配置 FFMPEG_BIN_DIR 或服务进程的 PATH")
    return executable


def get_audio_convert_config() -> dict[str, Any]:
    shared = get_word_count_config()
    try:
        ffmpeg_binary()
        reason = ""
    except AudioConvertError as exc:
        reason = str(exc)
    return {
        **limits(),
        **{key: shared.get(key) for key in ("allowed_roots", "unc_mount_mappings", "allow_local_paths")},
        "allowed_extensions": [".m4a"], "ffmpeg_available": not reason,
        "unavailable_reason": reason, "output_format": "PCM 16 位 WAV，保留原采样率和声道数",
    }


def validate_filename(filename: str) -> None:
    if Path(filename or "").suffix.lower() != ".m4a":
        raise ValueError("仅支持 M4A 文件")


def validate_sizes(sizes: list[int]) -> None:
    config = limits()
    if not sizes or len(sizes) > config["max_files"]:
        raise ValueError(f"每批请选择 1 至 {config['max_files']} 个 M4A 文件")
    if any(size <= 0 for size in sizes):
        raise ValueError("不能转换空文件")
    if any(size > config["max_file_mb"] * 1048576 for size in sizes):
        raise ValueError(f"单文件不能超过 {config['max_file_mb']} MB")
    if sum(sizes) > config["max_total_mb"] * 1048576:
        raise ValueError(f"整批输入不能超过 {config['max_total_mb']} MB")


def _relative_path(raw: str) -> str:
    value = str(raw).replace("\\", "/")
    path = PurePosixPath(value)
    if not value or path.is_absolute() or any(part in ("", ".", "..") for part in value.split("/")) or ":" in value:
        raise ValueError(f"无效的相对路径：{raw}")
    return path.as_posix()


def _is_hidden(path: PurePosixPath) -> bool:
    return any(part.startswith((".", "~$", "$")) for part in path.parts)


def resolve_selected_file(directory_path: str, relative_path: str, recursive: bool = True) -> Path:
    root, _, kind = resolve_allowed_shared_input_path(directory_path)
    relative = _relative_path(relative_path)
    base = root if kind == "directory" else root.parent
    candidate = (base / relative).resolve()
    if not candidate.is_relative_to(base) or (kind == "file" and candidate != root):
        raise PermissionError("所选文件超出共享路径范围")
    if _is_hidden(PurePosixPath(relative)) or (kind == "directory" and not recursive and len(PurePosixPath(relative).parts) > 1):
        raise ValueError("所选文件不符合当前扫描选项")
    validate_filename(candidate.name)
    if not candidate.is_file():
        raise FileNotFoundError(f"文件已删除或无法访问：{relative}")
    return candidate


def discover_audio_files(*, directory_path: str, recursive: bool = True) -> dict[str, Any]:
    root, _, kind = resolve_allowed_shared_input_path(directory_path)
    config = limits()
    files = []
    truncated = False

    def candidates():
        if kind == "file":
            validate_filename(root.name)
            yield root, root.name
            return
        def walk_error(exc):
            raise OSError(f"无法读取共享目录：{exc}") from exc
        for current, directories, names in os.walk(root, followlinks=False, onerror=walk_error):
            directories[:] = sorted(name for name in directories if not _is_hidden(PurePosixPath(name)))
            for name in sorted(names):
                path = Path(current) / name
                if path.suffix.lower() == ".m4a" and not _is_hidden(PurePosixPath(name)):
                    yield path, path.relative_to(root).as_posix()
            if not recursive:
                break

    for candidate, relative in candidates():
        if not candidate.resolve().is_relative_to(root if kind == "directory" else root.parent):
            continue
        if len(files) >= config["max_files"]:
            truncated = True
            break
        try:
            size = candidate.stat().st_size
            reason = "空文件" if size == 0 else ("超过单文件大小限制" if size > config["max_file_mb"] * 1048576 else "")
        except OSError:
            size, reason = 0, "文件无法读取"
        files.append({"name": candidate.name, "relative_path": relative, "size": size,
                      "selectable": not reason, "reason": reason})
    return {"files": files, "truncated": truncated, "max_files": config["max_files"],
            "selectable_count": sum(item["selectable"] for item in files)}


def prepare_audio_convert_request(*, directory_path: str, relative_paths: list[str], recursive: bool = True) -> dict:
    ffmpeg_binary()
    if not relative_paths or len(relative_paths) > limits()["max_files"]:
        raise ValueError(f"每批请选择 1 至 {limits()['max_files']} 个 M4A 文件")
    normalized = [_relative_path(value) for value in relative_paths]
    if len({value.casefold() for value in normalized}) != len(normalized):
        raise ValueError("不能重复选择同一个文件")
    manifest = []
    for relative in normalized:
        path = resolve_selected_file(directory_path, relative, recursive)
        stat = path.stat()
        manifest.append({"relative_path": relative, "size": stat.st_size, "mtime_ns": stat.st_mtime_ns})
    validate_sizes([item["size"] for item in manifest])
    return {"input_source": "path", "directory_path": directory_path.strip(), "recursive": recursive,
            "relative_paths": normalized, "manifest": manifest}


def _safe_output_name(relative: str) -> Path:
    parts = [re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", part).strip(" .") or "音频"
             for part in PurePosixPath(relative.replace("\\", "/")).parts]
    return Path(*parts).with_suffix(".wav")


def convert_file(source: Path, output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    try:
        completed = subprocess.run(
            [ffmpeg_binary(), "-nostdin", "-hide_banner", "-loglevel", "error", "-y",
             "-i", str(source), "-map", "0:a:0", "-vn", "-c:a", "pcm_s16le", str(output)],
            capture_output=True, text=True, encoding="utf-8", errors="replace", check=False,
            timeout=max(1, settings.AUDIO_CONVERT_TIMEOUT_SECONDS),
        )
        if completed.returncode or not output.is_file() or output.stat().st_size <= 44:
            detail = (completed.stderr or "未生成有效 WAV 文件").strip()[-600:]
            raise AudioConvertError(f"转换失败，文件可能损坏或没有音轨：{detail}")
    except subprocess.TimeoutExpired as exc:
        output.unlink(missing_ok=True)
        raise AudioConvertError("音频转换超时，请缩短音频后重试") from exc
    except Exception:
        output.unlink(missing_ok=True)
        raise


def _stage_shared_file(source: Path, target: Path, remaining_bytes: int) -> int:
    """限制实际复制的字节数，防止扫描后文件增长绕过大小限制。"""
    max_bytes = min(limits()["max_file_mb"] * 1048576, remaining_bytes)
    size = 0
    with source.open("rb") as reader, target.open("wb") as writer:
        while chunk := reader.read(1024 * 1024):
            size += len(chunk)
            if size > max_bytes:
                raise ValueError("文件或整批实际大小超过限制")
            writer.write(chunk)
    if size == 0:
        raise ValueError("不能转换空文件")
    return size


async def execute_audio_convert_task(*, task_id: str, display_no: str | None, input_files: dict,
                                     params: dict, progress_callback=None, executor: Executor | None = None) -> dict:
    loop = asyncio.get_running_loop()
    async def run(function, *args):
        return await loop.run_in_executor(executor, function, *args)
    async def report(progress, message):
        if progress_callback:
            await progress_callback(progress, message)

    await run(ffmpeg_binary)
    folder = display_no or task_id
    output_dir = Path(settings.OUTPUT_DIR) / "audio_convert" / folder
    staging = Path(settings.UPLOAD_DIR) / "audio_convert_staging" / folder
    output_dir.mkdir(parents=True, exist_ok=True)
    staging.mkdir(parents=True, exist_ok=True)
    is_shared = params.get("input_source") == "path"
    entries = [{"relative_path": value} for value in params["relative_paths"]] if is_shared else input_files["files"]
    results, outputs, used_names = [], [], set()
    staged_bytes = 0
    try:
        for index, item in enumerate(entries):
            name = item.get("relative_path") or item["filename"]
            await report(5 + int(index / len(entries) * 85), f"正在转换 {index + 1}/{len(entries)}：{name}")
            try:
                if is_shared:
                    source = await run(resolve_selected_file, params["directory_path"], name, params.get("recursive", True))
                    staged = staging / f"{index}.m4a"
                    staged_bytes += await run(_stage_shared_file, source, staged, limits()["max_total_mb"] * 1048576 - staged_bytes)
                    source = staged
                else:
                    source = Path(item["path"])
                    if not source.resolve().is_relative_to((Path(settings.UPLOAD_DIR) / "audio_convert" / folder).resolve()):
                        raise PermissionError("上传文件路径无效")
                    validate_filename(name)
                    validate_sizes([source.stat().st_size])
                    staged_bytes += source.stat().st_size
                    if staged_bytes > limits()["max_total_mb"] * 1048576:
                        raise ValueError("整批实际大小超过限制")
                relative = _safe_output_name(name if is_shared else name.replace("\\", "/").rsplit("/", 1)[-1])
                original = relative
                counter = 2
                while relative.as_posix().casefold() in used_names:
                    relative = original.with_name(f"{original.stem}_{counter}.wav")
                    counter += 1
                used_names.add(relative.as_posix().casefold())
                output = output_dir / relative
                await run(convert_file, source, output)
                web_path = "outputs/" + output.resolve().relative_to(Path(settings.OUTPUT_DIR).resolve()).as_posix()
                outputs.append({"filename": relative.as_posix(), "path": web_path})
                results.append({"filename": name, "status": "done", "output_path": web_path, "size": output.stat().st_size})
            except (OSError, ValueError, AudioConvertError) as exc:
                results.append({"filename": name, "status": "failed", "error": str(exc)})
            finally:
                if is_shared:
                    (staging / f"{index}.m4a").unlink(missing_ok=True)
        success = len(outputs)
        failed = len(results) - success
        summary = f"转换完成：成功 {success} 个，失败 {failed} 个"
        result = {"files": results, "output_files": outputs, "success_count": success,
                  "failed_count": failed, "summary_text": summary, "archive_zip": None}
        if not success:
            raise AudioConvertError("全部文件转换失败", result)
        await report(94, "正在打包 WAV 文件")
        archive = output_dir / "音频转换结果.zip"
        def write_archive():
            with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED, allowZip64=True) as bundle:
                for item in outputs:
                    bundle.write(output_dir / item["filename"], arcname=item["filename"])
        await run(write_archive)
        result["archive_zip"] = "outputs/" + archive.resolve().relative_to(Path(settings.OUTPUT_DIR).resolve()).as_posix()
        await report(99, summary)
        return result
    finally:
        shutil.rmtree(staging, ignore_errors=True)
