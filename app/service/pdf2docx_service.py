import asyncio
from concurrent.futures import Executor
from pathlib import Path
from typing import Any, Awaitable, Callable, Dict, Optional

from app.core.config import settings
from pdf2docx import HybridToDocxConverter, ocr_file

ProgressCallback = Callable[[int, str], Awaitable[None]]

PDF2DOCX_MODELS: Dict[str, Dict[str, str]] = {
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
    input_path: str,
    original_filename: str,
    model: str = "google/gemini-3-flash-preview",
    progress_callback: Optional[ProgressCallback] = None,
    executor: Optional[Executor] = None,
) -> Dict[str, Any]:
    if model not in PDF2DOCX_MODELS:
        raise ValueError(f"不支持的模型: {model}")
    if not settings.OPENROUTER_API_KEY:
        raise ValueError("未配置 OPENROUTER_API_KEY，无法执行 PDF 转 Word 任务")

    loop = asyncio.get_running_loop()
    input_file = Path(input_path)
    task_output_dir = Path(settings.OUTPUT_DIR) / "pdf2docx" / task_id
    task_output_dir.mkdir(parents=True, exist_ok=True)

    raw_output_path = task_output_dir / f"{input_file.stem}_raw.txt"
    docx_output_path = task_output_dir / f"{input_file.stem}.docx"

    await _maybe_report(progress_callback, 5, "文件已入队，准备调用视觉模型")
    raw_text = await loop.run_in_executor(
        executor,
        lambda: ocr_file(
            file_path=input_path,
            api_key=settings.OPENROUTER_API_KEY,
            model=model,
        ),
    )

    await _maybe_report(progress_callback, 70, "OCR 完成，正在整理中间文本")
    raw_output_path.write_text(raw_text, encoding="utf-8")

    await _maybe_report(progress_callback, 85, "正在生成 Word 文档")
    await loop.run_in_executor(
        executor,
        lambda: HybridToDocxConverter().convert(raw_text, str(docx_output_path)),
    )

    await _maybe_report(progress_callback, 95, "正在整理输出结果")
    return {
        "task_id": task_id,
        "filename": original_filename,
        "model": model,
        "raw_output_txt": _normalize_path(raw_output_path),
        "output_docx": _normalize_path(docx_output_path),
    }
