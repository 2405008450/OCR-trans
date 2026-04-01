import asyncio
import os
import uuid
from concurrent.futures import Executor
from typing import Any, Awaitable, Callable, Dict, Optional

from fastapi import UploadFile

from app.core.config import settings
from app.service.image_processor import convert_input_to_images, process_image

ai_process_lock = asyncio.Lock()

os.makedirs(settings.UPLOAD_DIR, exist_ok=True)
os.makedirs(settings.OUTPUT_DIR, exist_ok=True)
os.makedirs(settings.TEMP_IMAGES_DIR, exist_ok=True)

ProgressCallback = Callable[[int, str], Awaitable[None]]


async def _maybe_report(progress_callback: Optional[ProgressCallback], progress: int, message: str):
    if progress_callback:
        await progress_callback(progress, message)


def _normalize_path(path: Optional[str]) -> Optional[str]:
    return path.replace("\\", "/") if path else None


async def execute_ocr_task_from_path(
    *,
    task_id: str,
    display_no: Optional[str] = None,
    input_path: str,
    original_filename: str,
    from_lang: str = "zh",
    to_lang: str = "en",
    enable_correction: bool = False,
    enable_visualization: bool = True,
    card_side: str = "front",
    progress_callback: Optional[ProgressCallback] = None,
    executor: Optional[Executor] = None,
    **legacy_params: Any,
) -> Dict[str, Any]:
    loop = asyncio.get_running_loop()
    output_dir = os.path.join(settings.OUTPUT_DIR, "ocr", display_no or task_id)
    os.makedirs(output_dir, exist_ok=True)

    if legacy_params.get("doc_type") not in (None, "", "id_card"):
        raise ValueError("当前 OCR 模块仅支持身份证，不再支持结婚证")

    await _maybe_report(progress_callback, 5, "文件已入队，准备解析")
    image_paths = await loop.run_in_executor(executor, convert_input_to_images, input_path, settings.TEMP_IMAGES_DIR)
    if not image_paths:
        file_ext = os.path.splitext(input_path)[1].lower()
        raise ValueError(f"不支持的文件格式: {file_ext}")

    results = []
    total_images = len(image_paths)

    for idx, img_path in enumerate(image_paths):
        base_progress = 15 + int((idx / max(total_images, 1)) * 70)
        await _maybe_report(progress_callback, base_progress, f"正在处理第 {idx + 1}/{total_images} 张图片")

        async with ai_process_lock:
            await _maybe_report(progress_callback, base_progress + 5, "已获得处理槽位，开始执行身份证 OCR 任务")
            result = await loop.run_in_executor(
                executor,
                lambda: process_image(
                    input_path=img_path,
                    output_dir=output_dir,
                    from_lang=from_lang,
                    to_lang=to_lang,
                    enable_correction=enable_correction,
                    enable_visualization=enable_visualization,
                    card_side=card_side,
                ),
            )

        results.append(
            {
                "corrected_image": None,
                "visualization_image": _normalize_path(result.get("visualization")),
                "translated_image": _normalize_path(result.get("final_output")),
                "ocr_json": _normalize_path(result.get("raw_ocr_json")),
                "translated_json": _normalize_path(result.get("translated_json")),
            }
        )

    await _maybe_report(progress_callback, 95, "正在整理输出结果")
    return {
        "task_id": task_id,
        "filename": original_filename,
        "results": results,
        "total_images": total_images,
    }


async def run_llm_task(
    file: UploadFile,
    from_lang: str = "zh",
    to_lang: str = "en",
    enable_correction: bool = False,
    enable_visualization: bool = True,
    card_side: str = "front",
    **legacy_params: Any,
) -> Dict[str, Any]:
    task_id = str(uuid.uuid4())
    file_ext = os.path.splitext(file.filename or "")[1].lower() or ".jpg"
    input_path = os.path.join(settings.UPLOAD_DIR, f"{task_id}{file_ext}")

    with open(input_path, "wb") as f:
        f.write(await file.read())

    return await execute_ocr_task_from_path(
        task_id=task_id,
        input_path=input_path,
        original_filename=file.filename or os.path.basename(input_path),
        from_lang=from_lang,
        to_lang=to_lang,
        enable_correction=enable_correction,
        enable_visualization=enable_visualization,
        card_side=card_side,
        **legacy_params,
    )

