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
import posixpath
import time
import zipfile
from concurrent.futures import Executor
from io import BytesIO
from pathlib import Path
from typing import Any, Awaitable, Callable, Dict, List, Optional
from xml.etree import ElementTree as ET

from openai import OpenAI
from PIL import Image, UnidentifiedImageError

from app.core.config import settings
from app.core.file_naming import build_user_visible_filename, ensure_unique_path
from app.service.gemini_service import GEMINI_ROUTE_OPENROUTER, ensure_gemini_route_configured
from app.service.libreoffice_service import convert_doc_to_docx_via_libreoffice
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
    "tr": {"name": "土耳其文", "english_name": "Turkish"},
    "nl": {"name": "荷兰文", "english_name": "Dutch"},
    "id": {"name": "印尼文", "english_name": "Indonesian"},
    "ms": {"name": "马来文", "english_name": "Malay"},
    "pl": {"name": "波兰文", "english_name": "Polish"},
    "uk": {"name": "乌克兰文", "english_name": "Ukrainian"},
    "ro": {"name": "罗马尼亚文", "english_name": "Romanian"},
    "cs": {"name": "捷克文", "english_name": "Czech"},
    "hu": {"name": "匈牙利文", "english_name": "Hungarian"},
    "el": {"name": "希腊文", "english_name": "Greek"},
}

# OCR 模型配置（复用 pdf2docx 的模型）
DOC_TRANSLATE_DEFAULT_MODEL = "google/gemini-3-flash-preview"
DOC_TRANSLATE_DEFAULT_GEMINI_ROUTE = GEMINI_ROUTE_OPENROUTER

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

DOC_TRANSLATE_ALLOWED_EXTENSIONS = (
    ".pdf",
    ".png",
    ".jpg",
    ".jpeg",
    ".bmp",
    ".gif",
    ".webp",
    ".tif",
    ".tiff",
    ".doc",
    ".docx",
)
DOC_TRANSLATE_WORD_EXTENSIONS = {".doc", ".docx"}
DOC_TRANSLATE_OCR_IMAGE_EXTENSIONS = {
    ".png",
    ".jpg",
    ".jpeg",
    ".bmp",
    ".gif",
    ".webp",
    ".tif",
    ".tiff",
}
DOCX_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DOCX_OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DOCX_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
DOCX_VML_NS = "urn:schemas-microsoft-com:vml"


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


def get_doc_translate_allowed_extensions() -> List[str]:
    return list(DOC_TRANSLATE_ALLOWED_EXTENSIONS)


def get_supported_languages() -> Dict[str, Dict[str, str]]:
    return SUPPORTED_LANGUAGES


def _is_word_file(path: Path) -> bool:
    return path.suffix.lower() in DOC_TRANSLATE_WORD_EXTENSIONS


def _resolve_docx_relationship_target(base_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    base_dir = posixpath.dirname(base_part)
    return posixpath.normpath(posixpath.join(base_dir, target))


def _read_docx_image_relationships(zip_file: zipfile.ZipFile, part_name: str) -> Dict[str, str]:
    rels_name = posixpath.join(
        posixpath.dirname(part_name),
        "_rels",
        f"{posixpath.basename(part_name)}.rels",
    )
    try:
        rels_xml = zip_file.read(rels_name)
    except KeyError:
        return {}

    root = ET.fromstring(rels_xml)
    rel_tag = f"{{{DOCX_REL_NS}}}Relationship"
    image_relationships: Dict[str, str] = {}
    for rel in root.findall(rel_tag):
        rel_id = rel.get("Id")
        rel_type = rel.get("Type", "")
        target = rel.get("Target", "")
        if rel_id and target and rel_type.endswith("/image"):
            image_relationships[rel_id] = _resolve_docx_relationship_target(part_name, target)
    return image_relationships


def _iter_docx_image_targets(zip_file: zipfile.ZipFile):
    part_names = [
        "word/document.xml",
        *(
            sorted(
                name
                for name in zip_file.namelist()
                if name.startswith("word/header") and name.endswith(".xml")
            )
        ),
        *(
            sorted(
                name
                for name in zip_file.namelist()
                if name.startswith("word/footer") and name.endswith(".xml")
            )
        ),
    ]

    embed_attr = f"{{{DOCX_OFFICE_REL_NS}}}embed"
    id_attr = f"{{{DOCX_OFFICE_REL_NS}}}id"
    blip_tag = f"{{{DOCX_DRAWING_NS}}}blip"
    image_tag = f"{{{DOCX_VML_NS}}}imagedata"

    for part_name in part_names:
        if part_name not in zip_file.namelist():
            continue
        relationships = _read_docx_image_relationships(zip_file, part_name)
        if not relationships:
            continue

        root = ET.fromstring(zip_file.read(part_name))
        for node in root.iter():
            if node.tag == blip_tag:
                rel_id = node.get(embed_attr)
            elif node.tag == image_tag:
                rel_id = node.get(id_attr)
            else:
                continue

            target = relationships.get(rel_id or "")
            if target:
                yield target


def _write_embedded_image(blob: bytes, original_name: str, output_dir: Path, index: int) -> Optional[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    source_suffix = Path(original_name).suffix.lower()
    if source_suffix in DOC_TRANSLATE_OCR_IMAGE_EXTENSIONS:
        output_path = output_dir / f"word_image_{index:03d}{source_suffix}"
        output_path.write_bytes(blob)
        return output_path

    try:
        with Image.open(BytesIO(blob)) as image:
            output_path = output_dir / f"word_image_{index:03d}.png"
            if image.mode not in {"RGB", "RGBA", "L"}:
                image = image.convert("RGB")
            image.save(output_path, format="PNG")
            return output_path
    except (UnidentifiedImageError, OSError):
        return None


def _extract_word_images_for_ocr(docx_path: Path, output_dir: Path) -> List[Path]:
    extracted_paths: List[Path] = []
    with zipfile.ZipFile(docx_path) as zip_file:
        for index, target_name in enumerate(_iter_docx_image_targets(zip_file), start=1):
            try:
                image_blob = zip_file.read(target_name)
            except KeyError:
                continue
            image_path = _write_embedded_image(image_blob, target_name, output_dir, index)
            if image_path is not None:
                extracted_paths.append(image_path)
    return extracted_paths


def _prepare_word_input_for_ocr(input_path: Path, work_dir: Path) -> Path:
    if input_path.suffix.lower() != ".doc":
        return input_path
    work_dir.mkdir(parents=True, exist_ok=True)
    converted_path = work_dir / f"{input_path.stem}.docx"
    return Path(convert_doc_to_docx_via_libreoffice(input_path, converted_path))


def _join_text_segments(segments: List[str]) -> str:
    return "\n\n<page_break/>\n\n".join("" if segment is None else segment for segment in segments)


def _write_text_segments(output_dir: Path, segments: List[str]) -> List[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    paths: List[Path] = []
    for index, segment in enumerate(segments, start=1):
        segment_path = output_dir / f"part_{index:03d}.txt"
        segment_path.write_text("" if segment is None else segment, encoding="utf-8")
        paths.append(segment_path)
    return paths


# ============================================================
# LLM 翻译：调用 DeepSeek API
# ============================================================
def _build_translation_system_prompt(source_lang: str, target_lang: str) -> str:
    source_name = SUPPORTED_LANGUAGES.get(source_lang, {}).get("name", source_lang)
    target_name = SUPPORTED_LANGUAGES.get(target_lang, {}).get("name", target_lang)
    return f"""你是一个专业的文档翻译专家。请将以下文档内容从{source_name}翻译为{target_name}。

翻译规则：
1. 保留原文的所有格式标记（HTML 标签、Markdown 标记等），只翻译文字内容
2. 保持原文的段落结构和排版
3. 专有名词（如公司名称、人名、地名）应提供准确翻译
4. 数字、编号等保持原格式不变
5. 如果原文包含多语种对照、双语标签或同义重复项（如土耳其语+英语、中文+英语），输出时只保留目标语言版本，不要同时保留未翻译的对照文本
6. 对于“护照/PASSPORT”“国籍/NATIONALITY”这类证件固定栏位，如果多个源语言表达的是同一含义，只输出一次目标语言译文
7. 除证件号码、姓名拼写、机器可读码(MRZ)、URL、邮箱、品牌或机构官方缩写，以及用户明确要求保留的字段外，不要保留任何源语言原文
8. 仅返回翻译后的内容，不要添加任何解释或注释"""


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
    system_prompt = _build_translation_system_prompt(source_lang, target_lang)

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
    ocr_model: str = DOC_TRANSLATE_DEFAULT_MODEL,
    gemini_route: str = DOC_TRANSLATE_DEFAULT_GEMINI_ROUTE,
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
    source_images: List[str] = []
    ocr_segments: List[str] = []
    raw_part_paths: List[str] = []

    if _is_word_file(input_file):
        word_image_dir = task_output_dir / f"{stem}_word_images"
        if input_file.suffix.lower() == ".doc":
            await _maybe_report(progress_callback, 5, "正在转换 Word 文档格式...")
        else:
            await _maybe_report(progress_callback, 5, "正在提取 Word 文档中的图片...")

        prepared_word_path = await loop.run_in_executor(
            executor,
            lambda: _prepare_word_input_for_ocr(input_file, word_image_dir),
        )
        await _maybe_report(progress_callback, 10, "正在提取 Word 文档中的图片...")
        extracted_images = await loop.run_in_executor(
            executor,
            lambda: _extract_word_images_for_ocr(prepared_word_path, word_image_dir),
        )

        if not extracted_images:
            raise ValueError("Word 文档中未找到可处理的图片")

        total_images = len(extracted_images)
        await _maybe_report(progress_callback, 12, f"已提取 {total_images} 张图片，开始逐张 OCR...")

        for image_index, image_path in enumerate(extracted_images, start=1):
            image_progress = 12 + int(image_index / max(total_images, 1) * 20)
            await _maybe_report(
                progress_callback,
                min(image_progress, 32),
                f"正在处理 Word 图片 {image_index}/{total_images}...",
            )
            part_text = await loop.run_in_executor(
                executor,
                lambda path=str(image_path): ocr_file(
                    file_path=path,
                    model=ocr_model,
                    gemini_route=gemini_route,
                ),
            )
            ocr_segments.append(part_text or "")

        if not any(segment.strip() for segment in ocr_segments):
            raise ValueError("Word 文档中的图片未识别出可用文本")

        source_images = [_normalize_path(path) for path in extracted_images]
        raw_part_paths = [
            _normalize_path(path)
            for path in _write_text_segments(task_output_dir / f"{stem}_raw_parts", ocr_segments)
        ]
        raw_text = _join_text_segments(ocr_segments)
    else:
        await _maybe_report(progress_callback, 5, "正在调用视觉模型进行 OCR 识别...")
        raw_text = await loop.run_in_executor(
            executor,
            lambda: ocr_file(
                file_path=input_path,
                model=ocr_model,
                gemini_route=gemini_route,
            ),
        )
        ocr_segments = [raw_text]

    raw_output_path = task_output_dir / f"{stem}_raw.txt"
    raw_output_path.write_text(raw_text, encoding="utf-8")

    await _maybe_report(progress_callback, 35, "OCR 识别完成，准备翻译...")

    # ----------------------------------------------------------
    # Step 2: 逐语种翻译 + 生成 DOCX
    # ----------------------------------------------------------
    total_langs = len(target_langs)
    results_per_lang: Dict[str, Dict[str, Any]] = {}

    for idx, lang in enumerate(target_langs):
        lang_name = SUPPORTED_LANGUAGES.get(lang, {}).get("name", lang)
        lang_progress_base = 35 + int((idx / max(total_langs, 1)) * 55)
        translated_part_paths: List[str] = []

        if len(ocr_segments) > 1:
            total_segments = len(ocr_segments)
            translated_segments: List[str] = []
            for segment_index, segment_text in enumerate(ocr_segments, start=1):
                segment_progress = lang_progress_base + int((segment_index - 1) / max(total_segments, 1) * 20)
                await _maybe_report(
                    progress_callback,
                    min(segment_progress, 54),
                    f"正在翻译为{lang_name}（图片 {segment_index}/{total_segments}）... ({idx + 1}/{total_langs})",
                )
                translated_segment = await loop.run_in_executor(
                    executor,
                    lambda txt=segment_text, l=lang: _translate_text_with_llm(txt, source_lang, l) if txt.strip() else "",
                )
                translated_segments.append(translated_segment)

            translated_text = _join_text_segments(translated_segments)
            translated_part_paths = [
                _normalize_path(path)
                for path in _write_text_segments(task_output_dir / f"{stem}_{lang}_parts", translated_segments)
            ]
        else:
            await _maybe_report(
                progress_callback,
                lang_progress_base,
                f"正在翻译为{lang_name}... ({idx + 1}/{total_langs})",
            )
            translated_text = await loop.run_in_executor(
                executor,
                lambda l=lang: _translate_text_with_llm(raw_text, source_lang, l),
            )

        translated_txt_path = task_output_dir / f"{stem}_{lang}.txt"
        translated_txt_path.write_text(translated_text, encoding="utf-8")

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

        final_docx_path = ensure_unique_path(
            task_output_dir / build_user_visible_filename(original_filename, suffix=lang, ext=".docx"),
            existing_path=docx_path,
        )
        if docx_path != final_docx_path:
            docx_path.replace(final_docx_path)
            docx_path = final_docx_path

        results_per_lang[lang] = {
            "lang_code": lang,
            "lang_name": lang_name,
            "translated_txt": _normalize_path(translated_txt_path),
            "translated_parts": translated_part_paths,
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
        "raw_parts": raw_part_paths,
        "source_image_count": len(source_images),
        "source_images": source_images,
        "translations": results_per_lang,
    }
