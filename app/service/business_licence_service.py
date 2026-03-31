from __future__ import annotations

import asyncio
import importlib.util
import json
import re
from concurrent.futures import Executor
from functools import lru_cache
from pathlib import Path
from typing import Any, Awaitable, Callable, Dict, Optional

from app.core.config import settings
from app.service.gemini_service import (
    GEMINI_ROUTE_OPENROUTER,
    ensure_gemini_route_configured,
    generate_vision_html,
)

ProgressCallback = Callable[[int, str], Awaitable[None]]

BUSINESS_LICENCE_DEFAULT_MODEL = "google/gemini-3.1-pro-preview"
BUSINESS_LICENCE_DEFAULT_ROUTE = GEMINI_ROUTE_OPENROUTER

BUSINESS_LICENCE_MODELS: Dict[str, Dict[str, str]] = {
    "google/gemini-3.1-pro-preview": {
        "label": "Gemini 3.1 Pro",
        "description": "适合复杂版面和印章、二维码等细节识别。",
    },
    "google/gemini-3-flash-preview": {
        "label": "Gemini 3 Flash",
        "description": "速度更快，适合常规营业执照图片。",
    },
}

BUSINESS_LICENCE_SYSTEM_PROMPT = ""

BUSINESS_LICENCE_USER_PROMPT = """请仔细识别这张图片中的所有文字内容，提取出字段标签和对应的字段值，并将中文翻译成英文。

请以JSON格式返回结果，格式如下：
{
    "document_type": "文档类型（如：Business License, Invoice, ID Card等）",
    "layout": "文档布局（single_column 或 two_column）",
    "fields": [
        {
            "label_cn": "中文字段标签",
            "label_en": "English Field Label",
            "value_cn": "中文字段值",
            "value_en": "English Field Value",
            "position": "字段在原图中的位置（left 表示左栏，right 表示右栏，full_width 表示跨越整行，如果是单栏布局则都填 left）",
            "row": 行号（从1开始，同一行的字段行号相同）,
            "importance": "字段重要性（primary 表示主要字段如名称/类型/法人/经营范围/注册资本/成立日期/住所，secondary 表示次要字段）"
        }
    ],
    "seal_text": {
        "organization_cn": "印章中的机构名称（中文）",
        "organization_en": "Organization name in seal (English, use standard government agency translation format)",
        "date_cn": "印章覆盖的日期（中文，通常在印章下方或被印章覆盖的日期）",
        "date_en": "Date covered by seal (English)"
    },
    "credit_code": {
        "code": "统一社会信用代码的编号（如：91440000MA4XXXXXX）",
        "full_text_cn": "完整的信用代码文字（包括'统一社会信用代码'标签和编号）",
        "full_text_en": "Full credit code text in English (e.g., 'Unified Social Credit Code: 91440000MA4XXXXXX')"
    },
    "registration_no": {
        "no": "注册号/编号（如果文档中有单独的注册号或编号字段，提取其值；如果没有则为空字符串）",
        "full_text_en": "Registration No.: XXXXXXXX（完整的英文格式，如果没有则为空字符串）"
    },
    "qr_code": {
        "exists": true或false（图片中是否存在二维码）,
        "bbox": [x1, y1, x2, y2]（二维码的边界框坐标，像素值，左上角为原点。如果不存在则为null）,
        "position_description": "二维码在图片中的位置描述（如：右上角、左下角等）"
    },
    "qr_text": {
        "text_cn": "二维码旁边的说明文字（中文）",
        "text_en": "QR code description text (English translation)"
    },
    "duplicate_number": {
        "exists": true或false（图片中是否存在副本编号，如"副本 1-1"、"第一副本"等）,
        "number_cn": "副本编号的中文原文（如：1-1、第一副本等）",
        "number_en": "副本编号的英文格式（如：1-1）"
    }
}

注意：
1. 字段标签是指表单中的标题或提示文字（如"名称"、"地址"、"注册号"等）
2. 字段值是指对应标签后面的具体内容
3. 请识别所有可见的字段
4. 如果某个字段值为空，请用空字符串表示
5. 翻译要准确、专业
6. 只返回JSON，不要有其他解释文字
7. 特别注意识别红色印章中的文字（通常是圆形或椭圆形的公章），提取机构名称
8. 印章覆盖的日期通常在印章附近或被印章部分覆盖，格式如"2016年04月13日"，翻译成英文格式如"April 13, 2016"
9. 如果没有印章或印章文字不清晰，seal_text中的字段可以为空字符串
16. 重要：印章中的机构名称翻译必须使用标准政府机构翻译格式，例如：
    - "XX市XX区市场监督管理局" -> "Administration for Market Regulation of XX District, XX City"
    - "XX市市场监督管理局" -> "Administration for Market Regulation of XX City"
    - "XX省XX市工商行政管理局" -> "Administration for Industry and Commerce of XX City, XX Province"
    - 翻译时保持地名的拼音形式（如 Foshan, Zhuhai, Guangzhou 等），机构名称使用标准英文翻译
    - 严禁使用 "Municipality" 翻译"市"，必须使用 "City"
    - 严禁使用 "Bureau" 翻译"局"，必须使用 "Administration"
    - 正确格式示例：佛山市市场监督管理局 -> "Administration for Market Regulation of Foshan City"
    - 错误格式示例：佛山市市场监督管理局 -> "Foshan Municipality Market Supervision Bureau"（禁止）
10. 重要：请判断文档是单栏还是双栏布局，并标注每个字段的位置（left/right/full_width）和行号
11. 对于双栏布局，左栏和右栏同一水平位置的字段应该有相同的行号
12. 对于跨越整行的字段（如经营范围通常占据整行），position 应标记为 full_width
13. 特别注意识别"统一社会信用代码"（通常在文档左上角），提取完整的代码编号
13. 如果文档中有单独的"注册号"或"编号"字段（不同于统一社会信用代码），请提取到 registration_no 中
14. 特别注意识别二维码旁边的说明文字（通常在文档右上角，描述扫码用途），并翻译成英文。重要：只识别图片中实际存在的文字，如果二维码旁边没有任何说明文字，qr_text 的 text_cn 和 text_en 必须返回空字符串，不要猜测或推断可能存在的文字
15. 重要：请识别图片中的二维码位置，返回其边界框坐标（像素值）。坐标格式为[x1, y1, x2, y2]，其中(x1,y1)是左上角，(x2,y2)是右下角
17. 特别注意识别营业执照上的副本编号（通常在文档右上角或标题附近，如"副本"、"副本 1-1"、"第一副本"等）。如果存在副本编号，提取其编号格式（如1-1）；如果只写"副本"没有具体编号，或者是"正本"，则 duplicate_number.exists 为 false"""


async def _maybe_report(progress_callback: Optional[ProgressCallback], progress: int, message: str):
    if progress_callback:
        await progress_callback(progress, message)


def _normalize_path(path: Path) -> str:
    return str(path).replace("\\", "/")


def get_business_licence_models() -> Dict[str, Dict[str, str]]:
    return BUSINESS_LICENCE_MODELS


def _guess_image_mime_type(image_path: Path) -> str:
    suffix = image_path.suffix.lower()
    if suffix in {".jpg", ".jpeg"}:
        return "image/jpeg"
    if suffix == ".png":
        return "image/png"
    if suffix == ".webp":
        return "image/webp"
    if suffix == ".gif":
        return "image/gif"
    if suffix == ".bmp":
        return "image/bmp"
    if suffix in {".tif", ".tiff"}:
        return "image/tiff"
    return "image/jpeg"


def _extract_json_block(text: str) -> str:
    candidates: list[str] = []
    fenced = re.findall(r"```(?:json)?\s*([\s\S]*?)\s*```", text, flags=re.IGNORECASE)
    candidates.extend(fenced)

    stripped = text.strip()
    if stripped:
        candidates.append(stripped)

    start = stripped.find("{")
    end = stripped.rfind("}")
    if start != -1 and end != -1 and end > start:
        candidates.append(stripped[start : end + 1])

    for candidate in candidates:
        payload = candidate.strip().lstrip("\ufeff")
        if not payload:
            continue
        try:
            json.loads(payload)
            return payload
        except json.JSONDecodeError:
            continue
    raise json.JSONDecodeError("No valid JSON object found", stripped, 0)


def _replace_wai_in_code(value: str) -> str:
    return value.replace("外", "W.") if value else value


def normalize_business_licence_result(payload: Dict[str, Any]) -> Dict[str, Any]:
    data = dict(payload or {})
    data["document_type"] = data.get("document_type") or "Business License"
    data["layout"] = data.get("layout") or "two_column"
    data["fields"] = data.get("fields") if isinstance(data.get("fields"), list) else []

    data.setdefault(
        "seal_text",
        {"organization_cn": "", "organization_en": "", "date_cn": "", "date_en": ""},
    )
    data.setdefault(
        "credit_code",
        {"code": "", "full_text_cn": "", "full_text_en": ""},
    )
    data.setdefault(
        "registration_no",
        {"no": "", "full_text_en": ""},
    )
    data.setdefault(
        "qr_code",
        {"exists": False, "bbox": None, "position_description": ""},
    )
    data.setdefault(
        "qr_text",
        {"text_cn": "", "text_en": ""},
    )
    data.setdefault(
        "duplicate_number",
        {"exists": False, "number_cn": "", "number_en": ""},
    )

    credit_code = data.get("credit_code") or {}
    registration_no = data.get("registration_no") or {}
    if isinstance(credit_code, dict):
        if credit_code.get("code"):
            credit_code["code"] = _replace_wai_in_code(str(credit_code["code"]))
        if credit_code.get("full_text_en"):
            credit_code["full_text_en"] = _replace_wai_in_code(str(credit_code["full_text_en"]))
    if isinstance(registration_no, dict):
        if registration_no.get("no"):
            registration_no["no"] = _replace_wai_in_code(str(registration_no["no"]))
        if registration_no.get("full_text_en"):
            registration_no["full_text_en"] = _replace_wai_in_code(str(registration_no["full_text_en"]))

    return data


def parse_business_licence_response(raw_text: str) -> Dict[str, Any]:
    json_text = _extract_json_block(raw_text)
    return normalize_business_licence_result(json.loads(json_text))


def generate_business_licence_response(
    image_path: str | Path,
    *,
    model: str = BUSINESS_LICENCE_DEFAULT_MODEL,
    gemini_route: str = BUSINESS_LICENCE_DEFAULT_ROUTE,
    log_callback=None,
) -> str:
    image_file = Path(image_path)
    route = ensure_gemini_route_configured(gemini_route)
    return generate_vision_html(
        system_prompt=BUSINESS_LICENCE_SYSTEM_PROMPT,
        user_prompt=BUSINESS_LICENCE_USER_PROMPT,
        image_bytes=image_file.read_bytes(),
        mime_type=_guess_image_mime_type(image_file),
        model=model,
        route=route,
        temperature=0.0,
        max_output_tokens=20000,
        log_callback=log_callback,
    )


def extract_business_licence_data(
    image_path: str | Path,
    *,
    model: str = BUSINESS_LICENCE_DEFAULT_MODEL,
    gemini_route: str = BUSINESS_LICENCE_DEFAULT_ROUTE,
    log_callback=None,
) -> Dict[str, Any]:
    raw_text = generate_business_licence_response(
        image_path,
        model=model,
        gemini_route=gemini_route,
        log_callback=log_callback,
    )
    return parse_business_licence_response(raw_text)


@lru_cache(maxsize=1)
def _load_legacy_business_licence_module():
    module_path = Path(__file__).resolve().parents[2] / "businesslicence" / "start.py"
    spec = importlib.util.spec_from_file_location("businesslicence_start_legacy", module_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"无法加载营业执照脚本: {module_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


async def execute_business_licence_task(
    *,
    task_id: str,
    display_no: Optional[str] = None,
    input_path: str,
    original_filename: str,
    model: str = BUSINESS_LICENCE_DEFAULT_MODEL,
    gemini_route: str = BUSINESS_LICENCE_DEFAULT_ROUTE,
    progress_callback: Optional[ProgressCallback] = None,
    executor: Optional[Executor] = None,
) -> Dict[str, Any]:
    if model not in BUSINESS_LICENCE_MODELS:
        raise ValueError(f"不支持的模型: {model}")

    route = ensure_gemini_route_configured(gemini_route)
    loop = asyncio.get_running_loop()
    input_file = Path(input_path)
    output_dir = Path(settings.OUTPUT_DIR) / "business_licence" / (display_no or task_id)
    output_dir.mkdir(parents=True, exist_ok=True)

    output_docx_path = output_dir / f"{input_file.stem}_translated.docx"

    await _maybe_report(progress_callback, 5, "正在准备营业执照识别任务")

    raw_text = await loop.run_in_executor(
        executor,
        lambda: generate_business_licence_response(
            input_file,
            model=model,
            gemini_route=route,
        ),
    )

    await _maybe_report(progress_callback, 45, "模型识别完成，正在解析结构化结果")
    parsed_data = parse_business_licence_response(raw_text)

    await _maybe_report(progress_callback, 70, "正在套用营业执照模板生成 Word")
    legacy_module = _load_legacy_business_licence_module()
    selected_template = legacy_module.select_template(parsed_data, str(input_file))
    selected_template_path = Path(selected_template)
    default_template_path = Path(legacy_module.TEMPLATE_PATH)
    vertical_template_path = Path(legacy_module.TEMPLATE_PATH_VERTICAL)
    if not selected_template_path.exists():
        selected_template_path = default_template_path

    def _render_docx() -> None:
        if selected_template_path == vertical_template_path:
            legacy_module.fill_vertical_template(
                str(selected_template_path),
                str(output_docx_path),
                parsed_data,
                str(input_file),
            )
            return
        legacy_module.fill_template(
            str(selected_template_path),
            str(output_docx_path),
            parsed_data,
            str(input_file),
        )

    await loop.run_in_executor(executor, _render_docx)

    await _maybe_report(progress_callback, 95, "正在整理营业执照输出文件")
    return {
        "task_id": task_id,
        "filename": original_filename,
        "model": model,
        "gemini_route": route,
        "output_docx": _normalize_path(output_docx_path),
        "selected_template": _normalize_path(selected_template_path),
    }
