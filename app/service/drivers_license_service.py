import asyncio
import contextlib
import io
import os
import sys
from concurrent.futures import Executor
from pathlib import Path
from typing import Any, Awaitable, Callable, Dict, List, Optional

from app.core.config import settings

ProgressCallback = Callable[[int, str], Awaitable[None]]
SUPPORTED_PROCESSING_MODES = {
    "single": {"label": "单图处理", "min_files": 1, "allow_multi": False},
    "merge": {"label": "多图合并", "min_files": 1, "allow_multi": True},
    "batch": {"label": "多图批量", "min_files": 1, "allow_multi": True},
}


def _drivers_license_root() -> Path:
    return Path(__file__).resolve().parents[2] / "Driver's_License"


def _prepare_drivers_license_path() -> None:
    root = _drivers_license_root()
    src_dir = root / "src"
    if not root.exists() or not src_dir.exists():
        raise FileNotFoundError(
            f"驾驶证模块目录缺失: {src_dir}. "
            "请确认部署环境已包含 Driver's_License 目录，并在 Dockerfile 中执行 COPY Driver's_License/ ./Driver's_License/"
        )
    root_str = str(root)
    if root_str not in sys.path:
        sys.path.insert(0, root_str)


async def _maybe_report(progress_callback: Optional[ProgressCallback], progress: int, message: str) -> None:
    if progress_callback:
        await progress_callback(progress, message)


def _normalize_path(path: Path | str) -> str:
    return str(path).replace("\\", "/")


def get_drivers_license_config() -> Dict[str, Any]:
    return {
        "processing_modes": SUPPORTED_PROCESSING_MODES,
        "accept_types": [".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"],
        "max_files": 20,
    }


def _build_pipeline():
    _prepare_drivers_license_path()
    os.environ["GLM_API_KEY"] = settings.GLM_API_KEY or os.getenv("GLM_API_KEY", "")
    os.environ["DEEPSEEK_API_KEY"] = settings.DEEPSEEK_API_KEY or os.getenv("DEEPSEEK_API_KEY", "")

    from src.config import TranslatorConfig
    from src.translator_pipeline import TranslatorPipeline

    config = TranslatorConfig.from_env()
    if not config.glm_api_key:
        raise ValueError("未配置全局 .env 中的 GLM_API_KEY")
    if not config.deepseek_api_key:
        raise ValueError("未配置全局 .env 中的 DEEPSEEK_API_KEY")
    return TranslatorPipeline(config.glm_api_key, config.deepseek_api_key)


@contextlib.contextmanager
def _capture_stdout():
    stream = io.StringIO()
    with contextlib.redirect_stdout(stream), contextlib.redirect_stderr(stream):
        yield stream


def _run_single_sync(input_path: str, output_dir: Path) -> tuple[str, str]:
    pipeline = _build_pipeline()
    with _capture_stdout() as logs:
        output_path = pipeline.translate_image(input_path, str(output_dir))
    return output_path, logs.getvalue()


def _run_merge_sync(input_paths: List[str], output_dir: Path) -> tuple[str, str]:
    pipeline = _build_pipeline()
    with _capture_stdout() as logs:
        output_path = pipeline.translate_merge(input_paths, str(output_dir))
    return output_path, logs.getvalue()


def _run_batch_sync(input_paths: List[str], output_dir: Path) -> tuple[Dict[str, str], str]:
    pipeline = _build_pipeline()
    with _capture_stdout() as logs:
        result = pipeline.translate_batch(input_paths, str(output_dir))
    return result, logs.getvalue()


async def execute_drivers_license_task(
    *,
    task_id: str,
    display_no: Optional[str],
    input_paths: List[str],
    original_filenames: List[str],
    processing_mode: str,
    progress_callback: Optional[ProgressCallback] = None,
    executor: Optional[Executor] = None,
) -> Dict[str, Any]:
    if processing_mode not in SUPPORTED_PROCESSING_MODES:
        raise ValueError(f"不支持的处理模式: {processing_mode}")
    if not input_paths:
        raise ValueError("未提供图片文件")
    if processing_mode == "single" and len(input_paths) != 1:
        raise ValueError("单图处理模式只能上传 1 张图片")

    loop = asyncio.get_running_loop()
    folder_name = display_no or task_id
    output_dir = Path(settings.OUTPUT_DIR) / "drivers_license" / folder_name
    output_dir.mkdir(parents=True, exist_ok=True)

    await _maybe_report(progress_callback, 5, "正在初始化驾驶证翻译流水线...")

    if processing_mode == "single":
        await _maybe_report(progress_callback, 15, "正在识别并生成驾驶证 Word...")
        output_path, logs = await loop.run_in_executor(executor, lambda: _run_single_sync(input_paths[0], output_dir))
        await _maybe_report(progress_callback, 95, "正在整理输出结果...")
        return {
            "task_id": task_id,
            "processing_mode": processing_mode,
            "filename": original_filenames[0],
            "input_count": 1,
            "output_docx": _normalize_path(output_path),
            "items": [
                {
                    "input_filename": original_filenames[0],
                    "output_docx": _normalize_path(output_path),
                    "status": "done",
                }
            ],
            "stream_log": logs,
        }

    if processing_mode == "merge":
        await _maybe_report(progress_callback, 15, f"正在合并处理 {len(input_paths)} 张驾驶证图片...")
        output_path, logs = await loop.run_in_executor(executor, lambda: _run_merge_sync(input_paths, output_dir))
        await _maybe_report(progress_callback, 95, "正在整理输出结果...")
        return {
            "task_id": task_id,
            "processing_mode": processing_mode,
            "filename": " | ".join(original_filenames),
            "input_count": len(input_paths),
            "output_docx": _normalize_path(output_path),
            "items": [
                {
                    "input_filename": name,
                    "status": "merged",
                }
                for name in original_filenames
            ],
            "stream_log": logs,
        }

    await _maybe_report(progress_callback, 15, f"正在批量处理 {len(input_paths)} 张驾驶证图片...")
    result_map, logs = await loop.run_in_executor(executor, lambda: _run_batch_sync(input_paths, output_dir))
    await _maybe_report(progress_callback, 95, "正在整理批量输出结果...")

    items: List[Dict[str, Any]] = []
    success_count = 0
    fail_count = 0
    for path_str, result in result_map.items():
        input_name = Path(path_str).name
        if isinstance(result, str) and result.startswith("ERROR"):
            fail_count += 1
            items.append(
                {
                    "input_filename": input_name,
                    "status": "failed",
                    "error": result,
                }
            )
            continue
        success_count += 1
        items.append(
            {
                "input_filename": input_name,
                "status": "done",
                "output_docx": _normalize_path(result),
            }
        )

    return {
        "task_id": task_id,
        "processing_mode": processing_mode,
        "filename": " | ".join(original_filenames),
        "input_count": len(input_paths),
        "success_count": success_count,
        "failed_count": fail_count,
        "items": items,
        "stream_log": logs,
    }
