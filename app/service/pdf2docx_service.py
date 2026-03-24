import asyncio
from concurrent.futures import Executor
from pathlib import Path
from typing import Any, Awaitable, Callable, Dict, Optional

from app.core.config import settings
from app.service.gemini_service import ensure_gemini_route_configured
from pdf2docx import convert_text_to_word_via_libreoffice, ocr_file

ProgressCallback = Callable[[int, str], Awaitable[None]]

PDF2DOCX_MODELS: Dict[str, Dict[str, str]] = {
    "gemini-3.1-flash-lite-preview": {
        "label": "极速版V2",
        "description": "更轻量的极速 OCR 模型，适合追求速度的 PDF / 图片转 Word 场景。",
    },
    "google/gemini-3-flash-preview": {
        "label": "Google Gemini 3 Flash Preview",
        "description": "速度更快，适合常规 PDF / 图片转 Word 场景。",
    },
    "google/gemini-3.1-pro-preview": {
        "label": "Google Gemini 3.1 Pro Preview",
        "description": "更强调复杂版面与细节理解，适合高难度文档。",
    },
}


async def _maybe_report(progress_callback: Optional[ProgressCallback], progress: int, message: str):
    if progress_callback:
        await progress_callback(progress, message)


def _normalize_path(path: Path) -> str:
    return str(path).replace("\\", "/")


def get_pdf2docx_models() -> Dict[str, Dict[str, str]]:
    return PDF2DOCX_MODELS


async def execute_pdf2docx_task_from_path(
    *,
    task_id: str,
    display_no: Optional[str] = None,
    input_path: str,
    original_filename: str,
    model: str = "google/gemini-3-flash-preview",
    gemini_route: str = "google",
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

    _page_state: Dict[str, Any] = {"current": 0, "total": 0}

    def _page_cb(current: int, total: int):
        _page_state["current"] = current
        _page_state["total"] = total
        pct = 5 + int(current / max(total, 1) * 60)
        pct = min(pct, 65)
        if progress_callback:
            try:
                fut = asyncio.run_coroutine_threadsafe(
                    progress_callback(pct, f"正在 OCR 第 {current}/{total} 页"),
                    loop,
                )
                fut.result(timeout=30)
            except Exception:
                pass

    raw_text = await loop.run_in_executor(
        executor,
        lambda: ocr_file(
            file_path=input_path,
            model=model,
            gemini_route=gemini_route,
            page_progress_callback=_page_cb,
        ),
    )

    total_pages = _page_state.get("total", 0)
    pages_msg = f"OCR 完成（共 {total_pages} 页），正在整理中间文本" if total_pages else "OCR 完成，正在整理中间文本"
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

    await _maybe_report(progress_callback, 95, "正在整理输出结果")
    return {
        "task_id": task_id,
        "filename": original_filename,
        "model": model,
        "gemini_route": gemini_route,
        "raw_output_txt": _normalize_path(raw_output_path),
        "output_html": _normalize_path(html_output_path),
        "output_docx": _normalize_path(docx_output_path),
    }
