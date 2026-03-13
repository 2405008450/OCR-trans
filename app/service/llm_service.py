import asyncio
import os
import uuid
from typing import Any, Awaitable, Callable, Dict, Optional

from fastapi import UploadFile

from app.core.config import settings
from app.service.image_processor import convert_input_to_images, process_image
from app.service.marriage_cert_processor import process_marriage_cert_image

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
    input_path: str,
    original_filename: str,
    from_lang: str = "zh",
    to_lang: str = "en",
    enable_correction: bool = False,
    enable_visualization: bool = True,
    card_side: str = "front",
    doc_type: str = "id_card",
    marriage_page_template: str = "page2",
    registrar_signature_text: Optional[str] = None,
    registered_by_text: Optional[str] = None,
    registered_by_offset_x: int = 0,
    registered_by_offset_y: int = 0,
    registrar_signature_offset_x: int = 36,
    registrar_signature_offset_y: int = -12,
    enable_merge: bool = True,
    enable_overlap_fix: bool = True,
    enable_colon_fix: bool = False,
    font_size: Optional[int] = None,
    progress_callback: Optional[ProgressCallback] = None,
) -> Dict[str, Any]:
    loop = asyncio.get_running_loop()

    await _maybe_report(progress_callback, 5, "文件已入队，准备解析")
    image_paths = await loop.run_in_executor(None, convert_input_to_images, input_path, settings.TEMP_IMAGES_DIR)
    if not image_paths:
        file_ext = os.path.splitext(input_path)[1].lower()
        raise ValueError(f"不支持的文件格式: {file_ext}")

    results = []
    total_images = len(image_paths)

    for idx, img_path in enumerate(image_paths):
        base_progress = 15 + int((idx / max(total_images, 1)) * 70)
        await _maybe_report(progress_callback, base_progress, f"正在处理第 {idx + 1}/{total_images} 张图片")

        if doc_type == "marriage_cert":
            template = (marriage_page_template or "page2").lower().strip()
            confidence_threshold = 0.5
            _em, _eo, _ec = enable_merge, enable_overlap_fix, enable_colon_fix
            _rs, _rb = registrar_signature_text, registered_by_text
            _rbx, _rby = registered_by_offset_x, registered_by_offset_y
            _rsx, _rsy = registrar_signature_offset_x, registrar_signature_offset_y

            if template == "page1":
                _em, _eo, _ec = True, True, False
                confidence_threshold = 0.8
            elif template == "page2":
                _em, _eo, _ec = False, True, True
            elif template == "page3":
                _em, _eo, _ec = True, True, False

            if template != "page1":
                _rs = _rb = None
                _rbx = _rby = 0
                _rsx, _rsy = 36, -12

            async with ai_process_lock:
                await _maybe_report(progress_callback, base_progress + 5, "已获得处理槽位，开始执行结婚证任务")
                result = await loop.run_in_executor(
                    None,
                    lambda: process_marriage_cert_image(
                        input_path=img_path,
                        output_dir=settings.OUTPUT_DIR,
                        from_lang=from_lang,
                        to_lang=to_lang,
                        enable_correction=enable_correction,
                        enable_visualization=enable_visualization,
                        enable_merge=_em,
                        enable_overlap_fix=_eo,
                        enable_colon_fix=_ec,
                        font_size=font_size if font_size else 18,
                        confidence_threshold=confidence_threshold,
                        page_template=template,
                        registrar_signature_text=_rs,
                        registered_by_text=_rb,
                        registered_by_offset_x=_rbx,
                        registered_by_offset_y=_rby,
                        registrar_signature_offset_x=_rsx,
                        registrar_signature_offset_y=_rsy,
                    ),
                )
        else:
            async with ai_process_lock:
                await _maybe_report(progress_callback, base_progress + 5, "已获得处理槽位，开始执行 OCR 任务")
                result = await loop.run_in_executor(
                    None,
                    lambda: process_image(
                        input_path=img_path,
                        output_dir=settings.OUTPUT_DIR,
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
    doc_type: str = "id_card",
    marriage_page_template: str = "page2",
    registrar_signature_text: Optional[str] = None,
    registered_by_text: Optional[str] = None,
    registered_by_offset_x: int = 0,
    registered_by_offset_y: int = 0,
    registrar_signature_offset_x: int = 36,
    registrar_signature_offset_y: int = -12,
    enable_merge: bool = True,
    enable_overlap_fix: bool = True,
    enable_colon_fix: bool = False,
    font_size: Optional[int] = None,
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
        doc_type=doc_type,
        marriage_page_template=marriage_page_template,
        registrar_signature_text=registrar_signature_text,
        registered_by_text=registered_by_text,
        registered_by_offset_x=registered_by_offset_x,
        registered_by_offset_y=registered_by_offset_y,
        registrar_signature_offset_x=registrar_signature_offset_x,
        registrar_signature_offset_y=registrar_signature_offset_y,
        enable_merge=enable_merge,
        enable_overlap_fix=enable_overlap_fix,
        enable_colon_fix=enable_colon_fix,
        font_size=font_size,
    )

