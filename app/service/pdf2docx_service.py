import asyncio
from concurrent.futures import Executor
from pathlib import Path
from typing import Any, Awaitable, Callable, Dict, Optional

from app.core.config import settings
from app.core.file_naming import build_user_visible_filename, ensure_unique_path
from app.service.gemini_service import GEMINI_ROUTE_OPENROUTER, ensure_gemini_route_configured
from pdf2docx import convert_text_to_word_via_libreoffice, ocr_file

ProgressCallback = Callable[[int, str], Awaitable[None]]
PDF2DOCX_DEFAULT_GEMINI_ROUTE = GEMINI_ROUTE_OPENROUTER
PDF2DOCX_DEFAULT_MODEL = "google/gemini-3-flash-preview"

PDF2DOCX_MODELS: Dict[str, Dict[str, str]] = {
    "gemini-3.1-flash-lite-preview": {
        "label": "极速版V2",
        "description": "更轻量的 OCR 模型，适合追求速度的 PDF / 图片转 Word 场景。",
    },
    "google/gemini-3-flash-preview": {
        "label": "Google Gemini 3 Flash Preview",
        "description": "速度更快，适合常规 PDF / 图片转 Word 场景。",
    },
    "google/gemini-3.1-pro-preview": {
        "label": "Google Gemini 3.1 Pro Preview",
        "description": "更强的复杂版面与细节理解能力，适合高难度文档。",
    },
}


async def _maybe_report(
    progress_callback: Optional[ProgressCallback],
    progress: int,
    message: str,
):
    if progress_callback:
        await progress_callback(progress, message)


def _normalize_path(path: Path) -> str:
    return str(path).replace("\\", "/")


def get_pdf2docx_models() -> Dict[str, Dict[str, str]]:
    return PDF2DOCX_MODELS


def _normalize_ocr_payload(ocr_result: Any, fallback_total_pages: int = 0) -> Dict[str, Any]:
    if isinstance(ocr_result, dict):
        raw_text = str(ocr_result.get("text") or "")
        blank_pages = [
            int(page)
            for page in (ocr_result.get("blank_pages") or [])
            if str(page).strip().isdigit() and int(page) > 0
        ]
        blank_page_count = ocr_result.get("blank_page_count")
        try:
            blank_page_count = int(blank_page_count)
        except (TypeError, ValueError):
            blank_page_count = len(blank_pages)

        total_pages = ocr_result.get("total_pages")
        try:
            total_pages = int(total_pages)
        except (TypeError, ValueError):
            total_pages = fallback_total_pages

        return {
            "text": raw_text,
            "total_pages": max(total_pages, 0),
            "blank_page_count": max(blank_page_count, len(blank_pages), 0),
            "blank_pages": blank_pages,
        }

    return {
        "text": str(ocr_result or ""),
        "total_pages": max(int(fallback_total_pages or 0), 0),
        "blank_page_count": 0,
        "blank_pages": [],
    }


async def execute_pdf2docx_task_from_path(
    *,
    task_id: str,
    display_no: Optional[str] = None,
    input_path: str,
    original_filename: str,
    model: str = PDF2DOCX_DEFAULT_MODEL,
    gemini_route: str = PDF2DOCX_DEFAULT_GEMINI_ROUTE,
    progress_callback: Optional[ProgressCallback] = None,
    executor: Optional[Executor] = None,
) -> Dict[str, Any]:
    if model not in PDF2DOCX_MODELS:
        raise ValueError(f"不支持的模型: {model}")
    gemini_route = ensure_gemini_route_configured(gemini_route)

    loop = asyncio.get_running_loop()
    input_file = Path(input_path)
    task_output_dir = Path(settings.OUTPUT_DIR) / "pdf2docx" / (display_no or task_id)
    task_output_dir.mkdir(parents=True, exist_ok=True)

    raw_output_path = task_output_dir / f"{input_file.stem}_raw.txt"
    html_output_path = task_output_dir / f"{input_file.stem}.html"
    docx_output_path = task_output_dir / f"{input_file.stem}.docx"

    await _maybe_report(progress_callback, 5, "文件已入队，准备调用视觉模型")

    page_state: Dict[str, Any] = {"current": 0, "total": 0}

    def page_callback(current: int, total: int):
        page_state["current"] = current
        page_state["total"] = total
        pct = 5 + int(current / max(total, 1) * 60)
        pct = min(pct, 65)
        if progress_callback:
            future = asyncio.run_coroutine_threadsafe(
                progress_callback(pct, f"正在 OCR 第 {current}/{total} 页"),
                loop,
            )
            future.result(timeout=30)

    def ocr_status_callback(message: str):
        if not progress_callback:
            return
        current = page_state.get("current", 0)
        total = page_state.get("total", 0)
        pct = 8 if not total else min(65, 5 + int(current / max(total, 1) * 60))
        future = asyncio.run_coroutine_threadsafe(
            progress_callback(pct, message),
            loop,
        )
        future.result(timeout=30)

    ocr_result = await loop.run_in_executor(
        executor,
        lambda: ocr_file(
            file_path=input_path,
            model=model,
            gemini_route=gemini_route,
            page_progress_callback=page_callback,
            return_metadata=True,
            ocr_status_callback=ocr_status_callback,
        ),
    )
    ocr_payload = _normalize_ocr_payload(ocr_result, fallback_total_pages=page_state.get("total", 0))
    raw_text = ocr_payload["text"]

    total_pages = ocr_payload["total_pages"] or page_state.get("total", 0)
    blank_page_count = ocr_payload["blank_page_count"]
    blank_pages = ocr_payload["blank_pages"]
    if total_pages:
        if blank_page_count:
            pages_msg = f"OCR 完成（共 {total_pages} 页，空白页 {blank_page_count} 页），正在整理中间文本"
        else:
            pages_msg = f"OCR 完成（共 {total_pages} 页，未检测到空白页），正在整理中间文本"
    else:
        pages_msg = "OCR 完成，正在整理中间文本"
    await _maybe_report(progress_callback, 70, pages_msg)
    raw_output_path.write_text(raw_text, encoding="utf-8")

    await _maybe_report(progress_callback, 85, "正在生成 Word 文档")
    await loop.run_in_executor(
        executor,
        lambda: convert_text_to_word_via_libreoffice(
            raw_text,
            str(docx_output_path),
            html_output_path=str(html_output_path),
            title=input_file.stem,
        ),
    )

    final_docx_path = ensure_unique_path(
        task_output_dir / build_user_visible_filename(original_filename, ext=".docx"),
        existing_path=docx_output_path,
    )
    if docx_output_path != final_docx_path:
        docx_output_path.replace(final_docx_path)
        docx_output_path = final_docx_path

    await _maybe_report(progress_callback, 95, "正在整理输出结果")
    return {
        "task_id": task_id,
        "filename": original_filename,
        "model": model,
        "gemini_route": gemini_route,
        "raw_output_txt": _normalize_path(raw_output_path),
        "output_html": _normalize_path(html_output_path),
        "output_docx": _normalize_path(docx_output_path),
        "total_pages": total_pages,
        "blank_page_count": blank_page_count,
        "blank_pages": blank_pages,
    }
