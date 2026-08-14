from __future__ import annotations

import asyncio
import json
import math
import re
import subprocess
import tempfile
from collections import Counter
from concurrent.futures import Executor
from pathlib import Path
from typing import Any, Awaitable, Callable, Optional

from lxml import etree

from app.core.config import settings
from app.core.file_naming import build_user_visible_filename, ensure_unique_path
from app.service.fixed_layout_docx_service import (
    _CdpClient,
    _find_free_port,
    _wait_for_page_websocket,
    resolve_browser_path,
)
from app.service.gemini_service import ensure_gemini_route_configured, generate_vision_html
from app.service.pdf2docx_service import (
    PDF2DOCX_DEFAULT_GEMINI_ROUTE,
    PDF2DOCX_DEFAULT_MODEL,
    get_pdf2docx_models,
    normalize_pdf2docx_model,
)

ProgressCallback = Callable[[int, str], Awaitable[None]]

SVG_NS = "http://www.w3.org/2000/svg"
SVG_EDITABLE_DEFAULT_MODEL = PDF2DOCX_DEFAULT_MODEL
SVG_EDITABLE_DEFAULT_ROUTE = PDF2DOCX_DEFAULT_GEMINI_ROUTE
SVG_EDITABLE_MAX_BYTES = 10 * 1024 * 1024
SVG_EDITABLE_MAX_ELEMENTS = 50_000
SVG_EDITABLE_MAX_PIXELS = 16_000_000
SVG_EDITABLE_CONFIDENCE_THRESHOLD = 0.82

_FORBIDDEN_TAGS = {"script", "foreignObject", "iframe", "object", "embed", "audio", "video"}
_URL_FUNCTION_RE = re.compile(r"url\s*\(\s*(['\"]?)(.*?)\1\s*\)", re.I | re.S)
_JSON_FENCE_RE = re.compile(r"^\s*```(?:json)?\s*|\s*```\s*$", re.I)

SVG_TEXT_RECOGNITION_PROMPT = """You recover editable text from an SVG rendering.

Return strict JSON only, with no markdown or explanation:
{
  "lines": [
    {
      "text": "exact visible text on one visual line",
      "bbox": [left, top, right, bottom],
      "confidence": 0.0,
      "font_size": 24,
      "fill": "#000000",
      "font_family": "Arial"
    }
  ]
}

Coordinates must be absolute image pixels. Include only readable text that appears to be
outlined/vector text. Exclude logos, icons, decorative marks, photographs, charts and any
uncertain pseudo-text. Preserve punctuation, capitalization and spaces. Split multiline text
into separate line objects. confidence must reflect character and bbox certainty.
"""


def get_svg_editable_config() -> dict[str, Any]:
    return {
        "models": get_pdf2docx_models(),
        "default_model": SVG_EDITABLE_DEFAULT_MODEL,
        "default_route": SVG_EDITABLE_DEFAULT_ROUTE,
        "confidence_threshold": SVG_EDITABLE_CONFIDENCE_THRESHOLD,
        "max_file_mb": SVG_EDITABLE_MAX_BYTES // (1024 * 1024),
    }


def validate_svg_filename(filename: str | None) -> None:
    if Path(filename or "").suffix.lower() != ".svg":
        raise ValueError("仅支持 .svg 文件")


def _local_name(value: str) -> str:
    return value.rsplit("}", 1)[-1]


def _is_external_reference(value: str) -> bool:
    candidate = (value or "").strip()
    if not candidate or candidate.startswith("#"):
        return False
    lowered = candidate.lower()
    return not lowered.startswith(("data:image/png;base64,", "data:image/jpeg;base64,", "data:image/webp;base64,", "data:image/gif;base64,"))


def sanitize_svg(svg_bytes: bytes) -> tuple[etree._ElementTree, dict[str, Any]]:
    if len(svg_bytes) > SVG_EDITABLE_MAX_BYTES:
        raise ValueError("SVG 文件超过 10MB 限制")
    lowered = svg_bytes[:65536].lower()
    if b"<!doctype" in lowered or b"<!entity" in lowered:
        raise ValueError("SVG 不允许包含 DOCTYPE 或 XML 实体")

    parser = etree.XMLParser(resolve_entities=False, no_network=True, load_dtd=False, recover=False, huge_tree=False)
    try:
        root = etree.fromstring(svg_bytes, parser=parser)
    except etree.XMLSyntaxError as exc:
        raise ValueError(f"SVG XML 格式无效: {exc}") from exc
    if _local_name(root.tag).lower() != "svg":
        raise ValueError("文件根元素不是 <svg>")

    removed_elements = 0
    removed_attributes = 0
    elements = list(root.iter())
    if len(elements) > SVG_EDITABLE_MAX_ELEMENTS:
        raise ValueError("SVG 元素数量超过安全限制")

    for element in elements:
        if element is not root and _local_name(element.tag) in _FORBIDDEN_TAGS:
            parent = element.getparent()
            if parent is not None:
                parent.remove(element)
                removed_elements += 1
            continue
        for name, value in list(element.attrib.items()):
            attr_name = _local_name(name).lower()
            if attr_name.startswith("on"):
                del element.attrib[name]
                removed_attributes += 1
                continue
            if attr_name in {"href", "src"} and _is_external_reference(value):
                del element.attrib[name]
                removed_attributes += 1
                continue
            if attr_name == "style":
                unsafe = any(_is_external_reference(match.group(2)) for match in _URL_FUNCTION_RE.finditer(value))
                if unsafe or "@import" in value.lower() or "javascript:" in value.lower():
                    del element.attrib[name]
                    removed_attributes += 1
        if _local_name(element.tag).lower() == "style" and element.text:
            css = element.text
            unsafe = any(_is_external_reference(match.group(2)) for match in _URL_FUNCTION_RE.finditer(css))
            if unsafe or "@import" in css.lower() or "javascript:" in css.lower():
                element.text = ""
                removed_elements += 1

    path_count = 0
    text_count = 0
    for element in root.iter():
        name = _local_name(element.tag).lower()
        if name == "path":
            path_count += 1
            element.set("data-svg-edit-id", f"p{path_count}")
        elif name == "text":
            text_count += 1
            element.set("data-svg-edit-existing-text", f"t{text_count}")

    return etree.ElementTree(root), {
        "element_count": len(elements),
        "path_count": path_count,
        "existing_text_count": text_count,
        "removed_elements": removed_elements,
        "removed_attributes": removed_attributes,
    }


def _serialize_svg(tree: etree._ElementTree) -> bytes:
    return etree.tostring(tree, encoding="utf-8", xml_declaration=True, pretty_print=True)


def _render_svg(svg_bytes: bytes) -> tuple[bytes, dict[str, float]]:
    try:
        import fitz
    except ImportError as exc:
        raise RuntimeError("请先安装 PyMuPDF") from exc
    svg_root = etree.fromstring(svg_bytes, parser=etree.XMLParser(resolve_entities=False, no_network=True))
    view_box = [part for part in re.split(r"[\s,]+", svg_root.get("viewBox", "").strip()) if part]
    svg_x = svg_y = 0.0
    svg_width = svg_height = 0.0
    if len(view_box) == 4:
        try:
            svg_x, svg_y, svg_width, svg_height = [float(part) for part in view_box]
        except ValueError:
            svg_width = svg_height = 0.0

    document = fitz.open(stream=svg_bytes, filetype="svg")
    try:
        if len(document) != 1:
            raise ValueError("SVG 渲染结果不是单页")
        page = document[0]
        width = max(float(page.rect.width), 1.0)
        height = max(float(page.rect.height), 1.0)
        scale = min(3.0, max(1.0, math.sqrt(SVG_EDITABLE_MAX_PIXELS / (width * height))))
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
        return pix.tobytes("png"), {
            "page_width": width,
            "page_height": height,
            "pixel_width": float(pix.width),
            "pixel_height": float(pix.height),
            "svg_x": svg_x,
            "svg_y": svg_y,
            "svg_width": svg_width if svg_width > 0 else width,
            "svg_height": svg_height if svg_height > 0 else height,
        }
    finally:
        document.close()


def _collect_browser_elements(svg_path: Path, width: int, height: int) -> dict[str, Any]:
    browser = resolve_browser_path()
    port = _find_free_port()
    with tempfile.TemporaryDirectory(prefix="svg-edit-browser-", ignore_cleanup_errors=True) as profile_dir:
        command = [
            browser,
            "--headless=new",
            "--disable-gpu",
            "--hide-scrollbars",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-javascript",
            "--disable-background-networking",
            "--disable-extensions",
            "--force-device-scale-factor=1",
            f"--remote-debugging-port={port}",
            f"--user-data-dir={profile_dir}",
            f"--window-size={max(width, 320)},{max(height, 240)}",
            svg_path.as_uri(),
        ]
        process = subprocess.Popen(
            command,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        client: _CdpClient | None = None
        try:
            client = _CdpClient(_wait_for_page_websocket(port, svg_path))
            client.call("Runtime.enable")
            result = client.evaluate(
                """
(() => {
  const root = document.documentElement;
  if (!root || root.localName !== 'svg') return null;
  const rootMatrix = root.getScreenCTM();
  if (!rootMatrix) return null;
  const inverse = rootMatrix.inverse();
  function info(element) {
    try {
      const box = element.getBBox();
      const matrix = element.getScreenCTM();
      if (!matrix || box.width <= 0 || box.height <= 0) return null;
      const points = [
        new DOMPoint(box.x, box.y), new DOMPoint(box.x + box.width, box.y),
        new DOMPoint(box.x + box.width, box.y + box.height), new DOMPoint(box.x, box.y + box.height)
      ].map(point => point.matrixTransform(matrix).matrixTransform(inverse));
      const xs = points.map(point => point.x), ys = points.map(point => point.y);
      const style = getComputedStyle(element);
      return {
        id: element.getAttribute('data-svg-edit-id') || element.getAttribute('data-svg-edit-existing-text'),
        bbox: [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)],
        fill: style.fill || '#000000', font_family: style.fontFamily || 'Arial',
        font_size: parseFloat(style.fontSize) || 16
      };
    } catch (_) { return null; }
  }
  return {
    paths: Array.from(document.querySelectorAll('path[data-svg-edit-id]')).map(info).filter(Boolean),
    texts: Array.from(document.querySelectorAll('text[data-svg-edit-existing-text]')).map(info).filter(Boolean)
  };
})()
"""
            )
            if not isinstance(result, dict):
                raise RuntimeError("浏览器无法读取 SVG 元素坐标")
            return result
        finally:
            if client is not None:
                try:
                    client.call("Browser.close")
                except Exception:
                    pass
                try:
                    client.close()
                except Exception:
                    pass
            if process.poll() is None:
                process.terminate()
            try:
                process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait(timeout=5)


def _parse_json_response(text: str) -> dict[str, Any]:
    cleaned = _JSON_FENCE_RE.sub("", (text or "").strip())
    try:
        payload = json.loads(cleaned)
    except json.JSONDecodeError:
        start, end = cleaned.find("{"), cleaned.rfind("}")
        if start < 0 or end <= start:
            raise ValueError("文字识别结果不是有效 JSON")
        payload = json.loads(cleaned[start : end + 1])
    if not isinstance(payload, dict):
        raise ValueError("文字识别结果不是 JSON 对象")
    return payload


def _recognize_text(png_bytes: bytes, pixel_width: int, pixel_height: int, model: str, route: str) -> list[dict[str, Any]]:
    response = generate_vision_html(
        system_prompt=SVG_TEXT_RECOGNITION_PROMPT,
        user_prompt=f"Analyze this {pixel_width}x{pixel_height} SVG rendering and return the JSON object.",
        image_bytes=png_bytes,
        mime_type="image/png",
        model=model,
        route=route,
        temperature=0,
    )
    payload = _parse_json_response(response)
    lines: list[dict[str, Any]] = []
    for item in payload.get("lines") or []:
        if not isinstance(item, dict):
            continue
        text = str(item.get("text") or "").strip()
        bbox = item.get("bbox")
        if not text or not isinstance(bbox, list) or len(bbox) != 4:
            continue
        try:
            numbers = [float(value) for value in bbox]
            confidence = max(0.0, min(float(item.get("confidence") or 0), 1.0))
        except (TypeError, ValueError):
            continue
        if numbers[2] <= numbers[0] or numbers[3] <= numbers[1]:
            continue
        lines.append({**item, "text": text, "bbox": numbers, "confidence": confidence})
    return lines


def _intersection_ratio(inner: list[float], outer: list[float]) -> float:
    left, top = max(inner[0], outer[0]), max(inner[1], outer[1])
    right, bottom = min(inner[2], outer[2]), min(inner[3], outer[3])
    intersection = max(0.0, right - left) * max(0.0, bottom - top)
    area = max(1e-6, (inner[2] - inner[0]) * (inner[3] - inner[1]))
    return intersection / area


def _pixel_bbox_to_svg(bbox: list[float], render: dict[str, float]) -> list[float]:
    origin_x = render.get("svg_x", 0.0)
    origin_y = render.get("svg_y", 0.0)
    width = render.get("svg_width", render["page_width"])
    height = render.get("svg_height", render["page_height"])
    return [
        origin_x + bbox[0] * width / render["pixel_width"],
        origin_y + bbox[1] * height / render["pixel_height"],
        origin_x + bbox[2] * width / render["pixel_width"],
        origin_y + bbox[3] * height / render["pixel_height"],
    ]


def _append_editable_text(tree: etree._ElementTree, lines: list[dict[str, Any]], browser_data: dict[str, Any], render: dict[str, float], threshold: float) -> dict[str, Any]:
    root = tree.getroot()
    paths_by_id = {str(item.get("id")): item for item in browser_data.get("paths") or [] if item.get("id")}
    existing_texts = [item for item in browser_data.get("texts") or [] if isinstance(item.get("bbox"), list)]
    elements_by_id = {element.get("data-svg-edit-id"): element for element in root.iter() if element.get("data-svg-edit-id")}
    layer = etree.Element(f"{{{SVG_NS}}}g")
    layer.set("id", "svg-editable-text-layer")
    layer.set("data-generated-by", "fastapi-llm-demo")
    accepted, review_lines, hidden_ids = 0, [], set()

    for index, item in enumerate(lines, start=1):
        svg_bbox = _pixel_bbox_to_svg(item["bbox"], render)
        review = {**item, "svg_bbox": [round(value, 3) for value in svg_bbox], "status": "skipped", "matched_path_ids": []}
        if item["confidence"] < threshold:
            review["reason"] = "置信度低于阈值"
            review_lines.append(review)
            continue
        if any(_intersection_ratio(svg_bbox, existing["bbox"]) >= 0.45 for existing in existing_texts):
            review["reason"] = "原 SVG 已存在可编辑 text 元素"
            review_lines.append(review)
            continue

        pad_x = max((svg_bbox[2] - svg_bbox[0]) * 0.08, 1.0)
        pad_y = max((svg_bbox[3] - svg_bbox[1]) * 0.18, 1.0)
        expanded = [svg_bbox[0] - pad_x, svg_bbox[1] - pad_y, svg_bbox[2] + pad_x, svg_bbox[3] + pad_y]
        candidates = []
        for path_id, path_info in paths_by_id.items():
            path_bbox = path_info.get("bbox")
            if not isinstance(path_bbox, list) or len(path_bbox) != 4:
                continue
            path_area = max(0.0, (path_bbox[2] - path_bbox[0]) * (path_bbox[3] - path_bbox[1]))
            line_area = max(1.0, (expanded[2] - expanded[0]) * (expanded[3] - expanded[1]))
            if path_area > line_area * 1.5:
                continue
            if _intersection_ratio(path_bbox, expanded) >= 0.65:
                candidates.append(path_id)
        if not candidates:
            review["reason"] = "未匹配到文字轮廓路径"
            review_lines.append(review)
            continue

        fills = [str(paths_by_id[path_id].get("fill") or "") for path_id in candidates]
        fill = Counter(value for value in fills if value and value != "none").most_common(1)
        fill_value = fill[0][0] if fill else str(item.get("fill") or "#000000")
        font_size = float(item.get("font_size") or 0)
        if font_size > 0:
            font_size *= render.get("svg_height", render["page_height"]) / render["pixel_height"]
        if font_size <= 0:
            font_size = max((svg_bbox[3] - svg_bbox[1]) * 0.88, 1.0)
        text_element = etree.SubElement(layer, f"{{{SVG_NS}}}text")
        text_element.set("id", f"editable-text-{index}")
        text_element.set("x", f"{svg_bbox[0]:.3f}")
        text_element.set("y", f"{svg_bbox[3] - max((svg_bbox[3] - svg_bbox[1]) * 0.12, 0.5):.3f}")
        text_element.set("font-size", f"{font_size:.3f}")
        text_element.set("font-family", str(item.get("font_family") or "Arial, sans-serif"))
        text_element.set("fill", fill_value)
        text_element.set("textLength", f"{max(svg_bbox[2] - svg_bbox[0], 1.0):.3f}")
        text_element.set("lengthAdjust", "spacingAndGlyphs")
        text_element.set("data-confidence", f"{item['confidence']:.3f}")
        text_element.text = item["text"]

        for path_id in candidates:
            element = elements_by_id.get(path_id)
            if element is None or path_id in hidden_ids:
                continue
            old_style = element.get("style") or ""
            element.set("data-svg-edit-original-style", old_style)
            element.set("data-svg-edit-replaced-by", text_element.get("id"))
            element.set("style", f"{old_style.rstrip(';')};display:none!important")
            hidden_ids.add(path_id)

        accepted += 1
        review.update({"status": "converted", "matched_path_ids": candidates, "fill": fill_value})
        review_lines.append(review)

    if len(layer):
        root.append(layer)
    return {"converted_line_count": accepted, "hidden_path_count": len(hidden_ids), "lines": review_lines}


async def execute_svg_editable_task_from_path(
    *,
    task_id: str,
    display_no: Optional[str],
    input_path: str,
    original_filename: str,
    model: str = SVG_EDITABLE_DEFAULT_MODEL,
    gemini_route: str = SVG_EDITABLE_DEFAULT_ROUTE,
    confidence_threshold: float = SVG_EDITABLE_CONFIDENCE_THRESHOLD,
    progress_callback: Optional[ProgressCallback] = None,
    executor: Optional[Executor] = None,
) -> dict[str, Any]:
    validate_svg_filename(original_filename)
    model = normalize_pdf2docx_model(model)
    if model not in get_pdf2docx_models():
        raise ValueError(f"不支持的模型: {model}")
    route = ensure_gemini_route_configured(gemini_route)
    threshold = max(0.5, min(float(confidence_threshold), 0.99))
    loop = asyncio.get_running_loop()

    async def report(progress: int, message: str) -> None:
        if progress_callback:
            await progress_callback(progress, message)

    input_file = Path(input_path)
    output_dir = Path(settings.OUTPUT_DIR) / "svg_editable" / (display_no or task_id)
    output_dir.mkdir(parents=True, exist_ok=True)
    sanitized_path = output_dir / "sanitized.svg"
    preview_path = output_dir / "preview.png"
    output_svg = ensure_unique_path(output_dir / build_user_visible_filename(original_filename, suffix="editable", ext=".svg"))
    review_path = ensure_unique_path(output_dir / build_user_visible_filename(original_filename, suffix="review", ext=".json"))

    await report(8, "正在安全解析 SVG")
    tree, security = await loop.run_in_executor(executor, lambda: sanitize_svg(input_file.read_bytes()))
    sanitized_bytes = _serialize_svg(tree)
    sanitized_path.write_bytes(sanitized_bytes)

    await report(22, "正在渲染 SVG 预览")
    png_bytes, render = await loop.run_in_executor(executor, lambda: _render_svg(sanitized_bytes))
    preview_path.write_bytes(png_bytes)

    await report(36, "正在读取矢量路径坐标")
    browser_data = await loop.run_in_executor(
        executor,
        lambda: _collect_browser_elements(sanitized_path, int(render["pixel_width"]), int(render["pixel_height"])),
    )

    if security["path_count"]:
        await report(52, "正在识别转曲文字")
        lines = await loop.run_in_executor(
            executor,
            lambda: _recognize_text(png_bytes, int(render["pixel_width"]), int(render["pixel_height"]), model, route),
        )
    else:
        lines = []

    await report(78, "正在匹配文字与轮廓路径")
    conversion = _append_editable_text(tree, lines, browser_data, render, threshold)
    output_svg.write_bytes(_serialize_svg(tree))
    review = {
        "source_filename": original_filename,
        "model": model,
        "gemini_route": route,
        "confidence_threshold": threshold,
        "security": security,
        "render": render,
        **conversion,
        "notes": [
            "高置信度匹配路径已隐藏但未删除，可通过 data-svg-edit-original-style 恢复。",
            "字体名称、字距和艺术字效果为近似恢复，建议在矢量编辑器中复核。",
        ],
    }
    review_path.write_text(json.dumps(review, ensure_ascii=False, indent=2), encoding="utf-8")
    try:
        sanitized_path.unlink(missing_ok=True)
    except OSError:
        pass
    await report(95, "可编辑 SVG 已生成")
    return {
        "task_id": task_id,
        "filename": original_filename,
        "model": model,
        "gemini_route": route,
        "output_svg": str(output_svg).replace("\\", "/"),
        "review_json": str(review_path).replace("\\", "/"),
        "preview_png": str(preview_path).replace("\\", "/"),
        "converted_line_count": conversion["converted_line_count"],
        "hidden_path_count": conversion["hidden_path_count"],
        "existing_text_count": security["existing_text_count"],
        "warnings": review["notes"],
    }
