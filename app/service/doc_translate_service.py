# -*- coding: utf-8 -*-
"""
文档翻译服务

流程：
1. 复用 pdf2docx.py 的 OCR 能力，将 PDF/图片转为 raw text
2. 调用 DeepSeek LLM 将 raw text 翻译为目标语言
3. 使用 HybridToDocxConverter 将翻译文本导出为 Word 文档
4. 支持一对多翻译（如同时翻译成英文 + 西班牙文）
"""

import asyncio
import time
from concurrent.futures import Executor
from pathlib import Path
from typing import Any, Awaitable, Callable, Dict, List, Optional

from openai import OpenAI

from app.core.config import settings
from app.service.gemini_service import ensure_gemini_route_configured
from pdf2docx import convert_text_to_word_via_libreoffice, ocr_file

ProgressCallback = Callable[[int, str], Awaitable[None]]

# ============================================================
# 支持的语种配置
# ============================================================
SUPPORTED_LANGUAGES: Dict[str, Dict[str, str]] = {
    "zh": {"name": "中文", "english_name": "Chinese"},
    "en": {"name": "英文", "english_name": "English"},
    "ja": {"name": "日文", "english_name": "Japanese"},
    "ko": {"name": "韩文", "english_name": "Korean"},
    "es": {"name": "西班牙文", "english_name": "Spanish"},
    "fr": {"name": "法文", "english_name": "French"},
    "de": {"name": "德文", "english_name": "German"},
    "ru": {"name": "俄文", "english_name": "Russian"},
    "ar": {"name": "阿拉伯文", "english_name": "Arabic"},
    "pt": {"name": "葡萄牙文", "english_name": "Portuguese"},
    "it": {"name": "意大利文", "english_name": "Italian"},
    "th": {"name": "泰文", "english_name": "Thai"},
    "vi": {"name": "越南文", "english_name": "Vietnamese"},
}

# OCR 模型配置（复用 pdf2docx 的模型）
DOC_TRANSLATE_MODELS: Dict[str, Dict[str, str]] = {
    "google/gemini-3-flash-preview": {
        "label": "快速版V2",
        "description": "速度更快，适合常规 PDF / 图片文档。",
    },
    "google/gemini-3.1-pro-preview": {
        "label": "增强版V2",
        "description": "更强调复杂版面与细节理解，适合高难度文档。",
    },
}

# DeepSeek 翻译模型
DEEPSEEK_TRANSLATE_MODEL = "deepseek-chat"


# ============================================================
# 辅助函数
# ============================================================
async def _maybe_report(progress_callback: Optional[ProgressCallback], progress: int, message: str):
    if progress_callback:
        await progress_callback(progress, message)


def _normalize_path(path: Path) -> str:
    return str(path).replace("\\", "/")


def get_doc_translate_models() -> Dict[str, Dict[str, str]]:
    return DOC_TRANSLATE_MODELS


def get_supported_languages() -> Dict[str, Dict[str, str]]:
    return SUPPORTED_LANGUAGES


# ============================================================
# LLM 翻译：调用 DeepSeek API
# ============================================================
def _translate_text_with_llm(
    raw_text: str,
    source_lang: str,
    target_lang: str,
    retries: int = 3,
) -> str:
    """
    调用 DeepSeek API 翻译文本。
    分段处理防止超长文本导致单次调用失败。
    """
    source_name = SUPPORTED_LANGUAGES.get(source_lang, {}).get("name", source_lang)
    target_name = SUPPORTED_LANGUAGES.get(target_lang, {}).get("name", target_lang)

    system_prompt = f"""你是一个专业的文档翻译专家。请将以下文档内容从{source_name}翻译为{target_name}。

翻译规则：
1. 保留原文的所有格式标记（HTML 标签、Markdown 标记等），只翻译文字内容
2. 保持原文的段落结构和排版
3. 专有名词（如公司名称、人名、地名）应提供准确翻译
4. 日期、数字、编号等保持原格式不变
5. 仅返回翻译后的内容，不要添加任何解释或注释"""

    # 分段：每段不超过 6000 字符
    MAX_CHUNK_SIZE = 6000
    chunks = _split_text_into_chunks(raw_text, MAX_CHUNK_SIZE)
    translated_parts = []

    client = OpenAI(
        api_key=settings.DEEPSEEK_API_KEY,
        base_url=settings.DEEPSEEK_BASE_URL,
    )

    for i, chunk in enumerate(chunks):
        chunk_text = chunk.strip()
        if not chunk_text:
            translated_parts.append("")
            continue

        for attempt in range(retries):
            try:
                response = client.chat.completions.create(
                    model=DEEPSEEK_TRANSLATE_MODEL,
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": chunk_text},
                    ],
                    temperature=0.1,
                    max_tokens=8192,
                )
                result = response.choices[0].message.content or ""
                translated_parts.append(result)
                break
            except Exception as e:
                if attempt < retries - 1:
                    wait = 3 * (attempt + 1)
                    print(f"⚠️ 翻译请求失败({e.__class__.__name__})，{wait}秒后重试 [{attempt + 1}/{retries}]...")
                    time.sleep(wait)
                else:
                    raise RuntimeError(f"翻译失败（已重试 {retries} 次）: {e}")

    return "\n\n".join(translated_parts)


def _split_text_into_chunks(text: str, max_size: int) -> List[str]:
    """按 page_break 标记和段落分割文本，确保每段不超过 max_size"""
    # 先按 page_break 分割
    pages = text.split("<page_break/>")
    chunks = []
    current_chunk = ""

    for page in pages:
        page = page.strip()
        if not page:
            continue

        if len(current_chunk) + len(page) + 20 <= max_size:
            if current_chunk:
                current_chunk += "\n\n<page_break/>\n\n" + page
            else:
                current_chunk = page
        else:
            if current_chunk:
                chunks.append(current_chunk)
            # 如果单页就超限，按段落再分
            if len(page) > max_size:
                paragraphs = page.split("\n\n")
                sub_chunk = ""
                for para in paragraphs:
                    if len(sub_chunk) + len(para) + 2 <= max_size:
                        sub_chunk = sub_chunk + "\n\n" + para if sub_chunk else para
                    else:
                        if sub_chunk:
                            chunks.append(sub_chunk)
                        sub_chunk = para
                if sub_chunk:
                    current_chunk = sub_chunk
                else:
                    current_chunk = ""
            else:
                current_chunk = page

    if current_chunk:
        chunks.append(current_chunk)

    return chunks if chunks else [text]


# ============================================================
# 主流程：OCR → 翻译 → DOCX
# ============================================================
async def execute_doc_translate_task(
    *,
    task_id: str,
    display_no: Optional[str] = None,
    input_path: str,
    original_filename: str,
    source_lang: str = "zh",
    target_langs: List[str],
    ocr_model: str = "google/gemini-3-flash-preview",
    gemini_route: str = "google",
    progress_callback: Optional[ProgressCallback] = None,
    executor: Optional[Executor] = None,
) -> Dict[str, Any]:
    """
    执行文档翻译任务完整流程。

    Args:
        task_id: 任务 ID
        display_no: 显示编号
        input_path: 输入文件路径
        original_filename: 原始文件名
        source_lang: 源语言代码
        target_langs: 目标语言代码列表（支持多个）
        ocr_model: OCR 模型
        progress_callback: 进度回调
        executor: 线程池
    """
    if ocr_model not in DOC_TRANSLATE_MODELS:
        raise ValueError(f"不支持的 OCR 模型: {ocr_model}")
    gemini_route = ensure_gemini_route_configured(gemini_route)
    if not settings.DEEPSEEK_API_KEY:
        raise ValueError("未配置 DEEPSEEK_API_KEY，无法执行翻译")

    loop = asyncio.get_running_loop()
    task_output_dir = Path(settings.OUTPUT_DIR) / "doc_translate" / (display_no or task_id)
    task_output_dir.mkdir(parents=True, exist_ok=True)

    input_file = Path(input_path)
    stem = input_file.stem

    # ----------------------------------------------------------
    # Step 1: OCR 提取原始文本
    # ----------------------------------------------------------
    await _maybe_report(progress_callback, 5, "正在调用视觉模型进行 OCR 识别...")

    raw_text = await loop.run_in_executor(
        executor,
        lambda: ocr_file(
            file_path=input_path,
            model=ocr_model,
            gemini_route=gemini_route,
        ),
    )

    # 保存原始 OCR 文本
    raw_output_path = task_output_dir / f"{stem}_raw.txt"
    raw_output_path.write_text(raw_text, encoding="utf-8")

    await _maybe_report(progress_callback, 35, "OCR 识别完成，准备翻译...")

    # ----------------------------------------------------------
    # Step 2: 逐语种翻译 + 生成 DOCX
    # ----------------------------------------------------------
    total_langs = len(target_langs)
    results_per_lang: Dict[str, Dict[str, str]] = {}

    for idx, lang in enumerate(target_langs):
        lang_name = SUPPORTED_LANGUAGES.get(lang, {}).get("name", lang)
        lang_progress_base = 35 + int((idx / max(total_langs, 1)) * 55)

        await _maybe_report(
            progress_callback,
            lang_progress_base,
            f"正在翻译为{lang_name}... ({idx + 1}/{total_langs})",
        )

        # 翻译
        translated_text = await loop.run_in_executor(
            executor,
            lambda l=lang: _translate_text_with_llm(raw_text, source_lang, l),
        )

        # 保存翻译后文本
        translated_txt_path = task_output_dir / f"{stem}_{lang}.txt"
        translated_txt_path.write_text(translated_text, encoding="utf-8")

        # 生成 DOCX
        await _maybe_report(
            progress_callback,
            lang_progress_base + int(30 / max(total_langs, 1)),
            f"正在生成{lang_name} Word 文档...",
        )

        html_path = task_output_dir / f"{stem}_{lang}.html"
        docx_path = task_output_dir / f"{stem}_{lang}.docx"
        await loop.run_in_executor(
            executor,
            lambda txt=translated_text, out=str(docx_path), html=str(html_path): convert_text_to_word_via_libreoffice(
                txt,
                out,
                html_output_path=html,
                title=f"{stem}_{lang}",
            ),
        )

        results_per_lang[lang] = {
            "lang_code": lang,
            "lang_name": lang_name,
            "translated_txt": _normalize_path(translated_txt_path),
            "output_html": _normalize_path(html_path),
            "output_docx": _normalize_path(docx_path),
        }

    await _maybe_report(progress_callback, 95, "正在整理输出结果...")

    return {
        "task_id": task_id,
        "filename": original_filename,
        "ocr_model": ocr_model,
        "gemini_route": gemini_route,
        "source_lang": source_lang,
        "raw_output_txt": _normalize_path(raw_output_path),
        "translations": results_per_lang,
    }
