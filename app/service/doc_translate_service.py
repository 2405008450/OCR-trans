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
import re
import time
import zipfile
from concurrent.futures import Executor
from io import BytesIO
from pathlib import Path
from typing import Any, Awaitable, Callable, Dict, List, Optional, Tuple
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
    # ── 东亚 ──────────────────────────────────────────────
    "zh": {"name": "中文（简体）", "english_name": "Chinese (Simplified)"},
    "zh-TW": {"name": "中文（繁体）", "english_name": "Chinese (Traditional)"},
    "ja": {"name": "日文", "english_name": "Japanese"},
    "ko": {"name": "韩文", "english_name": "Korean"},
    # ── 东南亚 ────────────────────────────────────────────
    "th": {"name": "泰文", "english_name": "Thai"},
    "vi": {"name": "越南文", "english_name": "Vietnamese"},
    "id": {"name": "印尼文", "english_name": "Indonesian"},
    "ms": {"name": "马来文", "english_name": "Malay"},
    "fil": {"name": "菲律宾文", "english_name": "Filipino"},
    "my": {"name": "缅甸文", "english_name": "Burmese"},
    "km": {"name": "高棉文", "english_name": "Khmer"},
    "lo": {"name": "老挝文", "english_name": "Lao"},
    # ── 南亚 ──────────────────────────────────────────────
    "hi": {"name": "印地文", "english_name": "Hindi"},
    "bn": {"name": "孟加拉文", "english_name": "Bengali"},
    "ur": {"name": "乌尔都文", "english_name": "Urdu"},
    "ta": {"name": "泰米尔文", "english_name": "Tamil"},
    "te": {"name": "泰卢固文", "english_name": "Telugu"},
    "ml": {"name": "马拉雅拉姆文", "english_name": "Malayalam"},
    "si": {"name": "僧伽罗文", "english_name": "Sinhala"},
    "ne": {"name": "尼泊尔文", "english_name": "Nepali"},
    # ── 西亚 / 中东 ───────────────────────────────────────
    "ar": {"name": "阿拉伯文", "english_name": "Arabic"},
    "fa": {"name": "波斯文", "english_name": "Persian"},
    "he": {"name": "希伯来文", "english_name": "Hebrew"},
    "tr": {"name": "土耳其文", "english_name": "Turkish"},
    "az": {"name": "阿塞拜疆文", "english_name": "Azerbaijani"},
    "ka": {"name": "格鲁吉亚文", "english_name": "Georgian"},
    "am": {"name": "阿姆哈拉文", "english_name": "Amharic"},
    # ── 中亚 ──────────────────────────────────────────────
    "kk": {"name": "哈萨克文", "english_name": "Kazakh"},
    "uz": {"name": "乌兹别克文", "english_name": "Uzbek"},
    "ky": {"name": "柯尔克孜文", "english_name": "Kyrgyz"},
    "tg": {"name": "塔吉克文", "english_name": "Tajik"},
    "tk": {"name": "土库曼文", "english_name": "Turkmen"},
    "mn": {"name": "蒙古文", "english_name": "Mongolian"},
    # ── 西欧 ──────────────────────────────────────────────
    "en": {"name": "英文", "english_name": "English"},
    "fr": {"name": "法文", "english_name": "French"},
    "de": {"name": "德文", "english_name": "German"},
    "es": {"name": "西班牙文", "english_name": "Spanish"},
    "pt": {"name": "葡萄牙文", "english_name": "Portuguese"},
    "it": {"name": "意大利文", "english_name": "Italian"},
    "nl": {"name": "荷兰文", "english_name": "Dutch"},
    "sv": {"name": "瑞典文", "english_name": "Swedish"},
    "no": {"name": "挪威文", "english_name": "Norwegian"},
    "da": {"name": "丹麦文", "english_name": "Danish"},
    "fi": {"name": "芬兰文", "english_name": "Finnish"},
    "is": {"name": "冰岛文", "english_name": "Icelandic"},
    "ga": {"name": "爱尔兰文", "english_name": "Irish"},
    "cy": {"name": "威尔士文", "english_name": "Welsh"},
    "eu": {"name": "巴斯克文", "english_name": "Basque"},
    "ca": {"name": "加泰罗尼亚文", "english_name": "Catalan"},
    "gl": {"name": "加利西亚文", "english_name": "Galician"},
    # ── 东欧 ──────────────────────────────────────────────
    "ru": {"name": "俄文", "english_name": "Russian"},
    "pl": {"name": "波兰文", "english_name": "Polish"},
    "uk": {"name": "乌克兰文", "english_name": "Ukrainian"},
    "cs": {"name": "捷克文", "english_name": "Czech"},
    "sk": {"name": "斯洛伐克文", "english_name": "Slovak"},
    "hu": {"name": "匈牙利文", "english_name": "Hungarian"},
    "ro": {"name": "罗马尼亚文", "english_name": "Romanian"},
    "bg": {"name": "保加利亚文", "english_name": "Bulgarian"},
    "hr": {"name": "克罗地亚文", "english_name": "Croatian"},
    "sr": {"name": "塞尔维亚文", "english_name": "Serbian"},
    "sl": {"name": "斯洛文尼亚文", "english_name": "Slovenian"},
    "bs": {"name": "波斯尼亚文", "english_name": "Bosnian"},
    "mk": {"name": "马其顿文", "english_name": "Macedonian"},
    "sq": {"name": "阿尔巴尼亚文", "english_name": "Albanian"},
    "el": {"name": "希腊文", "english_name": "Greek"},
    "lt": {"name": "立陶宛文", "english_name": "Lithuanian"},
    "lv": {"name": "拉脱维亚文", "english_name": "Latvian"},
    "et": {"name": "爱沙尼亚文", "english_name": "Estonian"},
    "be": {"name": "白俄罗斯文", "english_name": "Belarusian"},
    # ── 非洲 ──────────────────────────────────────────────
    "sw": {"name": "斯瓦希里文", "english_name": "Swahili"},
    "yo": {"name": "约鲁巴文", "english_name": "Yoruba"},
    "ig": {"name": "伊博文", "english_name": "Igbo"},
    "ha": {"name": "豪萨文", "english_name": "Hausa"},
    "zu": {"name": "祖鲁文", "english_name": "Zulu"},
    "so": {"name": "索马里文", "english_name": "Somali"},
    # ── 美洲 ──────────────────────────────────────────────
    "pt-BR": {"name": "葡萄牙文（巴西）", "english_name": "Portuguese (Brazil)"},
    "es-419": {"name": "西班牙文（拉丁美洲）", "english_name": "Spanish (Latin America)"},
    "qu": {"name": "克丘亚文", "english_name": "Quechua"},
    "ht": {"name": "海地克里奥尔文", "english_name": "Haitian Creole"},
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
DOC_TRANSLATE_DEFAULT_MODE = "standard"
DOC_TRANSLATE_TRANSLATE_MODES: Dict[str, Dict[str, str]] = {
    "standard": {
        "label": "标准翻译",
        "description": "仅输出目标语言译文",
    },
    "bilingual": {
        "label": "双语对照",
        "description": "按原文和译文换行对照输出，目标语言内容原样保留",
    },
}

# 文本翻译模型
DOC_TRANSLATE_TRANSLATION_MODEL = "deepseek-v4-flash"
DOC_TRANSLATE_TRANSLATION_MAX_TOKENS = 384000
DOC_TRANSLATE_TRANSLATION_REQUEST_MAX_TOKENS = 384000

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
DOCX_WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
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


def get_doc_translate_modes() -> Dict[str, Dict[str, str]]:
    return DOC_TRANSLATE_TRANSLATE_MODES


def get_supported_languages() -> Dict[str, Dict[str, str]]:
    return SUPPORTED_LANGUAGES


def normalize_doc_translate_mode(translate_mode: Optional[str]) -> str:
    normalized = str(translate_mode or DOC_TRANSLATE_DEFAULT_MODE).strip().lower()
    if normalized not in DOC_TRANSLATE_TRANSLATE_MODES:
        raise ValueError(f"不支持的翻译模式: {translate_mode}")
    return normalized


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

    try:
        root = ET.fromstring(rels_xml)
    except ET.ParseError:
        return {}
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


def _iter_docx_content_parts(zip_file: zipfile.ZipFile) -> List[str]:
    part_names = ["word/document.xml"]
    part_names.extend(
        sorted(
            name
            for name in zip_file.namelist()
            if (
                (name.startswith("word/header") or name.startswith("word/footer"))
                and name.endswith(".xml")
            )
        )
    )
    part_names.extend(
        name
        for name in ("word/footnotes.xml", "word/endnotes.xml")
        if name in zip_file.namelist()
    )
    return [name for name in part_names if name in zip_file.namelist()]


def _normalize_word_text_segment(text: str) -> str:
    normalized = (text or "").replace("\r", "\n").replace("\xa0", " ")
    lines = [re.sub(r"[ \t]+$", "", line) for line in normalized.splitlines()]
    normalized = "\n".join(lines)
    normalized = re.sub(r"\n{3,}", "\n\n", normalized)
    return normalized.strip()


def _append_word_text_segment(segments: List[Dict[str, Any]], text_parts: List[str]) -> None:
    text = _normalize_word_text_segment("".join(text_parts))
    text_parts.clear()
    if text:
        segments.append({"type": "text", "text": text})


def _extract_docx_part_segments(
    zip_file: zipfile.ZipFile,
    part_name: str,
    output_dir: Path,
    image_start_index: int,
) -> Tuple[List[Dict[str, Any]], int, List[str]]:
    relationships = _read_docx_image_relationships(zip_file, part_name)
    try:
        root = ET.fromstring(zip_file.read(part_name))
    except ET.ParseError as exc:
        return [], image_start_index, [f"Word 内容 XML 解析失败，已跳过 {part_name}: {exc}"]

    w_text_tag = f"{{{DOCX_WORD_NS}}}t"
    w_tab_tag = f"{{{DOCX_WORD_NS}}}tab"
    w_break_tag = f"{{{DOCX_WORD_NS}}}br"
    w_carriage_return_tag = f"{{{DOCX_WORD_NS}}}cr"
    w_paragraph_tag = f"{{{DOCX_WORD_NS}}}p"
    w_table_cell_tag = f"{{{DOCX_WORD_NS}}}tc"
    w_table_row_tag = f"{{{DOCX_WORD_NS}}}tr"
    drawing_text_tag = f"{{{DOCX_DRAWING_NS}}}t"
    blip_tag = f"{{{DOCX_DRAWING_NS}}}blip"
    image_tag = f"{{{DOCX_VML_NS}}}imagedata"
    embed_attr = f"{{{DOCX_OFFICE_REL_NS}}}embed"
    link_attr = f"{{{DOCX_OFFICE_REL_NS}}}link"
    id_attr = f"{{{DOCX_OFFICE_REL_NS}}}id"

    segments: List[Dict[str, Any]] = []
    warnings: List[str] = []
    text_parts: List[str] = []
    image_index = image_start_index

    def walk(node: ET.Element) -> None:
        nonlocal image_index
        tag = node.tag

        if tag in {w_text_tag, drawing_text_tag}:
            if node.text:
                text_parts.append(node.text)
            return

        if tag == w_tab_tag:
            text_parts.append("\t")
            return

        if tag in {w_break_tag, w_carriage_return_tag}:
            text_parts.append("\n")
            return

        if tag in {blip_tag, image_tag}:
            rel_id = node.get(embed_attr) or node.get(link_attr) or node.get(id_attr)
            target_name = relationships.get(rel_id or "")
            _append_word_text_segment(segments, text_parts)
            if target_name:
                try:
                    image_blob = zip_file.read(target_name)
                except KeyError:
                    warnings.append(f"Word 图片资源缺失，已跳过: {target_name}")
                    return

                image_path = _write_embedded_image(image_blob, target_name, output_dir, image_index)
                if image_path is None:
                    warnings.append(f"Word 图片格式暂不支持 OCR，已跳过: {target_name}")
                    return

                segments.append(
                    {
                        "type": "image",
                        "path": image_path,
                        "target": target_name,
                    }
                )
                image_index += 1
            return

        for child in node:
            walk(child)

        if tag == w_paragraph_tag:
            text_parts.append("\n")
        elif tag == w_table_cell_tag:
            text_parts.append("\t")
        elif tag == w_table_row_tag:
            text_parts.append("\n")

    walk(root)
    _append_word_text_segment(segments, text_parts)
    return segments, image_index, warnings


def _extract_word_segments_for_translation(
    docx_path: Path,
    output_dir: Path,
) -> Tuple[List[Dict[str, Any]], List[Path], List[str]]:
    segments: List[Dict[str, Any]] = []
    warnings: List[str] = []
    image_index = 1

    with zipfile.ZipFile(docx_path) as zip_file:
        for part_name in _iter_docx_content_parts(zip_file):
            part_segments, image_index, part_warnings = _extract_docx_part_segments(
                zip_file,
                part_name,
                output_dir,
                image_index,
            )
            segments.extend(part_segments)
            warnings.extend(part_warnings)

    image_paths = [
        segment["path"]
        for segment in segments
        if segment.get("type") == "image" and isinstance(segment.get("path"), Path)
    ]
    return segments, image_paths, warnings


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


_PAGE_BREAK_PATTERN = re.compile(r"\s*<page_break\s*/>\s*", flags=re.IGNORECASE)


def _split_ocr_text_segments(raw_text: str) -> List[str]:
    text = "" if raw_text is None else str(raw_text)
    if "<page_break" not in text.lower():
        stripped = text.strip()
        return [stripped] if stripped else [""]

    parts = _PAGE_BREAK_PATTERN.split(text)
    if len(parts) <= 1:
        stripped = text.strip()
        return [stripped] if stripped else [""]
    return [part.strip() for part in parts]


def _normalize_translation_max_tokens(max_tokens: int) -> int:
    try:
        requested = int(max_tokens)
    except (TypeError, ValueError):
        requested = DOC_TRANSLATE_TRANSLATION_REQUEST_MAX_TOKENS
    return max(1, min(requested, DOC_TRANSLATE_TRANSLATION_REQUEST_MAX_TOKENS))


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
8. 日期格式需要翻译成目标语种常用日期格式，但必须严格遵守源文档国家/语种的日期顺序，不要猜测或改写数字
9. 法文/法国或欧盟驾驶证中的点分短日期通常是“日.月.年”（DD.MM.YY 或 DD.MM.YYYY），两位年份在最后一段；例如“10.01.13”应译为“2013年1月10日”
10. 看到“19.01.13”这类三段两位点分数字时，除非上下文明确第一段是年份，否则禁止译为“2019年01月13日”；应按证件上下文核对为日.月.年，无法确认时保留原格式
11. 驾驶证字段编号（如 1、2、3、4a、4b、10、11、12）与日期相邻时，字段编号不是日期的一部分
12. 仅返回翻译后的内容，不要添加任何解释或注释"""


def _build_bilingual_system_prompt(source_lang: str, target_lang: str) -> str:
    source_name = SUPPORTED_LANGUAGES.get(source_lang, {}).get("name", source_lang)
    target_name = SUPPORTED_LANGUAGES.get(target_lang, {}).get("name", target_lang)
    return f"""你是一个专业的文档翻译专家。请对以下文档执行双语对照翻译。

处理规则：
1. 对于{source_name}内容：在原文下方紧跟一行{target_name}译文，形成“原文\\n译文”的逐段对照格式
2. 对于已经是{target_name}的内容：原样保留，不要翻译，也不要重复补写对照内容
3. 保留原文的所有格式标记（HTML 标签、Markdown 标记等），只翻译需要翻译的文字内容
4. 保持原文的段落结构、分页标记和排版顺序不变
5. 专有名词应提供准确翻译
6. 数字、编号、证件号码、URL、邮箱、品牌或机构官方缩写保持原格式不变
7. 日期格式保持原样；尤其不要把法文/法国或欧盟驾驶证中的“DD.MM.YY”误改写为“YY.MM.DD”
8. 仅返回处理后的内容，不要添加任何解释或注释
9. 双语对照格式示例：
   原文段落A
   Translation of paragraph A

   原文段落B
   Translation of paragraph B

   Already English paragraph (kept as-is)"""


def _translate_text_with_llm(
    raw_text: str,
    source_lang: str,
    target_lang: str,
    retries: int = 3,
    translate_mode: str = DOC_TRANSLATE_DEFAULT_MODE,
) -> str:
    """
    调用 DeepSeek API 翻译文本。
    分段处理防止超长文本导致单次调用失败。
    """
    resolved_translate_mode = normalize_doc_translate_mode(translate_mode)
    if resolved_translate_mode == "bilingual":
        system_prompt = _build_bilingual_system_prompt(source_lang, target_lang)
    else:
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
                    model=DOC_TRANSLATE_TRANSLATION_MODEL,
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": chunk_text},
                    ],
                    max_tokens=_normalize_translation_max_tokens(DOC_TRANSLATE_TRANSLATION_MAX_TOKENS),
                )
                result = response.choices[0].message.content or ""
                translated_parts.append(result)
                break
            except Exception as e:
                if attempt < retries - 1:
                    wait = 3 * (attempt + 1)
                    print(f"⚠️ DeepSeek 翻译请求失败({e.__class__.__name__})，{wait}秒后重试 [{attempt + 1}/{retries}]...")
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
    translate_mode: str = DOC_TRANSLATE_DEFAULT_MODE,
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
    translate_mode = normalize_doc_translate_mode(translate_mode)
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
    processing_warnings: List[str] = []
    source_segment_label = "页"

    if _is_word_file(input_file):
        source_segment_label = "Word 片段"
        word_asset_dir = task_output_dir / f"{stem}_word_assets"
        if input_file.suffix.lower() == ".doc":
            await _maybe_report(progress_callback, 5, "正在转换 Word 文档格式...")
        else:
            await _maybe_report(progress_callback, 5, "正在读取 Word 文档内容...")

        prepared_word_path = await loop.run_in_executor(
            executor,
            lambda: _prepare_word_input_for_ocr(input_file, word_asset_dir),
        )
        await _maybe_report(progress_callback, 10, "正在提取 Word 正文和图片...")
        word_segments, extracted_images, word_warnings = await loop.run_in_executor(
            executor,
            lambda: _extract_word_segments_for_translation(prepared_word_path, word_asset_dir),
        )
        processing_warnings.extend(word_warnings)

        text_segment_count = sum(1 for segment in word_segments if segment.get("type") == "text")
        total_images = len(extracted_images)

        if not word_segments:
            detail = f"；{processing_warnings[0]}" if processing_warnings else ""
            raise ValueError(f"Word 文档中未找到可处理的文本或图片{detail}")

        if total_images:
            await _maybe_report(
                progress_callback,
                12,
                f"已读取 {text_segment_count} 段文本、提取 {total_images} 张图片，开始逐张 OCR...",
            )
        else:
            await _maybe_report(progress_callback, 12, f"已读取 {text_segment_count} 段 Word 文本...")

        image_counter = 0
        failed_image_count = 0
        for segment in word_segments:
            if segment.get("type") == "text":
                ocr_segments.append(segment.get("text") or "")
                continue

            image_path = segment.get("path")
            if not isinstance(image_path, Path):
                continue

            image_counter += 1
            image_progress = 12 + int(image_counter / max(total_images, 1) * 20)
            await _maybe_report(
                progress_callback,
                min(image_progress, 32),
                f"正在处理 Word 图片 {image_counter}/{total_images}...",
            )
            try:
                part_text = await loop.run_in_executor(
                    executor,
                    lambda path=str(image_path): ocr_file(
                        file_path=path,
                        model=ocr_model,
                        gemini_route=gemini_route,
                    ),
                )
            except Exception as exc:
                failed_image_count += 1
                warning = f"Word 图片 {image_counter}/{total_images} OCR 失败，已继续处理后续内容: {exc}"
                processing_warnings.append(warning)
                await _maybe_report(progress_callback, min(image_progress, 32), warning)
                ocr_segments.append("")
                continue

            ocr_segments.append(part_text or "")

        if not any(segment.strip() for segment in ocr_segments):
            if failed_image_count:
                raise ValueError(f"Word 文档未识别出可用文本，{failed_image_count} 张图片 OCR 失败")
            raise ValueError("Word 文档中未识别出可用文本")

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
        ocr_segments = _split_ocr_text_segments(raw_text)
        if len(ocr_segments) > 1:
            raw_part_paths = [
                _normalize_path(path)
                for path in _write_text_segments(task_output_dir / f"{stem}_raw_parts", ocr_segments)
            ]

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
            segment_label = source_segment_label
            for segment_index, segment_text in enumerate(ocr_segments, start=1):
                segment_progress = lang_progress_base + int((segment_index - 1) / max(total_segments, 1) * 20)
                await _maybe_report(
                    progress_callback,
                    min(segment_progress, 54),
                    f"正在翻译为{lang_name}（{segment_label} {segment_index}/{total_segments}）... ({idx + 1}/{total_langs})",
                )
                translated_segment = await loop.run_in_executor(
                    executor,
                    lambda txt=segment_text, l=lang, mode=translate_mode: _translate_text_with_llm(
                        txt,
                        source_lang,
                        l,
                        translate_mode=mode,
                    ) if txt.strip() else "",
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
                lambda l=lang, mode=translate_mode: _translate_text_with_llm(
                    raw_text,
                    source_lang,
                    l,
                    translate_mode=mode,
                ),
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
        "translation_model": DOC_TRANSLATE_TRANSLATION_MODEL,
        "gemini_route": gemini_route,
        "source_lang": source_lang,
        "translate_mode": translate_mode,
        "raw_output_txt": _normalize_path(raw_output_path),
        "raw_parts": raw_part_paths,
        "source_image_count": len(source_images),
        "source_images": source_images,
        "warnings": processing_warnings,
        "translations": results_per_lang,
    }





