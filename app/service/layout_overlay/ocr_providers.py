"""云端 OCR 适配；所有坐标换算到输入图片像素，禁止凭纯文本猜坐标。"""
from __future__ import annotations

import base64
import io
import math
from typing import Protocol

import cv2
import httpx
from PIL import Image

from app.core.config import settings
from .models import OcrLine, union_bbox, validate_bbox


class OcrProvider(Protocol):
    def recognize(self, image_bytes: bytes) -> list[OcrLine]: ...


def _size(data: bytes):
    with Image.open(io.BytesIO(data)) as image:
        return image.size


def _line(text, polygon, width, height, confidence=None, **kwargs):
    if not str(text).strip() or not polygon:
        return None
    polygon = [(float(x), float(y)) for x, y in polygon]
    if not all(math.isfinite(v) for point in polygon for v in point):
        raise ValueError("OCR 多边形含非法坐标")
    # Google 行可能由多个 word 的四边形拼接，取凸包避免自交掩码。
    if len(polygon) > 4:
        import numpy as np
        polygon = [tuple(point) for point in cv2.convexHull(np.array(polygon, dtype=np.float32)).reshape(-1, 2).tolist()]
    bbox = validate_bbox((min(x for x, _ in polygon), min(y for _, y in polygon),
                          max(x for x, _ in polygon), max(y for _, y in polygon)), width, height)
    if confidence is not None:
        confidence = float(confidence)
        if not math.isfinite(confidence):
            confidence = None
        else:
            confidence = max(0, min(1, confidence))
    return OcrLine(str(text).strip(), polygon, bbox, confidence, **kwargs)


def _vertices(box, width, height):
    normalized = "normalizedVertices" in box
    return [(float(p.get("x", 0)) * (width if normalized else 1),
             float(p.get("y", 0)) * (height if normalized else 1))
            for p in box.get("normalizedVertices" if normalized else "vertices", [])]


def parse_google_response(payload: dict, width: int, height: int) -> list[OcrLine]:
    response = (payload.get("responses") or [{}])[0]
    if response.get("error"):
        raise ValueError("Google Vision 返回识别错误")
    lines = []
    for page in response.get("fullTextAnnotation", {}).get("pages", []):
        sx, sy = width / (page.get("width") or width), height / (page.get("height") or height)
        for block_index, block in enumerate(page.get("blocks", [])):
            for para in block.get("paragraphs", []):
                text, points, words, confidences = "", [], [], []
                for word in para.get("words", []):
                    symbols = word.get("symbols", [])
                    word_text = "".join(s.get("text", "") for s in symbols)
                    vertices = _vertices(word.get("boundingBox", {}), page.get("width") or width, page.get("height") or height)
                    points.extend([(x * sx, y * sy) for x, y in vertices])
                    text += word_text
                    words.append({"text": word_text, "polygon": [(x * sx, y * sy) for x, y in vertices]})
                    if "confidence" in word:
                        confidences.append(word["confidence"])
                    br = (symbols[-1].get("property", {}).get("detectedBreak", {}).get("type") if symbols else None)
                    if br in {"SPACE", "SURE_SPACE", "EOL_SURE_SPACE"}:
                        text += " "
                    if br in {"LINE_BREAK", "EOL_SURE_SPACE", "HYPHEN"}:
                        line = _line(text, points, width, height, sum(confidences) / len(confidences) if confidences else None,
                                     words=words, block_id=str(block_index))
                        if line:
                            lines.append(line)
                        text, points, words, confidences = "", [], [], []
                line = _line(text, points, width, height, sum(confidences) / len(confidences) if confidences else None,
                             words=words, block_id=str(block_index))
                if line:
                    lines.append(line)
    if not lines:
        for item in response.get("textAnnotations", [])[1:]:
            line = _line(item.get("description", ""), _vertices(item.get("boundingPoly", {}), width, height), width, height)
            if line:
                lines.append(line)
    return lines


def parse_qwen_response(payload: dict, width: int, height: int) -> list[OcrLine]:
    lines = []
    for choice in payload.get("output", {}).get("choices", []):
        for part in choice.get("message", {}).get("content", []):
            if str(part.get("text") or "").strip() and not part.get("ocr_result", {}).get("words_info"):
                raise ValueError("Qwen 返回非空文字但缺少行坐标")
            for item in part.get("ocr_result", {}).get("words_info", []):
                location = item.get("location", [])
                if len(location) == 8:
                    polygon = list(zip(location[::2], location[1::2]))
                elif item.get("rotate_rect") and len(item["rotate_rect"]) == 5:
                    cx, cy, w, h, angle = map(float, item["rotate_rect"])
                    polygon = cv2.boxPoints(((cx, cy), (w, h), angle)).tolist()
                else:
                    continue
                line = _line(item.get("text", ""), polygon, width, height, item.get("confidence"))
                if line:
                    lines.append(line)
    return lines


def parse_glm_response(payload: dict, width: int, height: int) -> list[OcrLine]:
    lines = []
    # 同时兼容旧 files/ocr 的 location 与新 layout_parsing 的像素 bbox_2d。
    regions = payload.get("words_result", [])
    details = payload.get("layout_details", [])
    if details:
        regions = details[0]
    page_info = (payload.get("data_info", {}).get("pages") or [{}])[0]
    sx, sy = width / (page_info.get("width") or width), height / (page_info.get("height") or height)
    for item in regions:
        if item.get("label") in {"image", "seal", "chart"}:
            continue
        location = item.get("location")
        if isinstance(location, dict):
            x, y = location.get("left", 0), location.get("top", 0)
            box = (x, y, x + location.get("width", 0), y + location.get("height", 0))
        else:
            box = item.get("bbox_2d")
        if not box or len(box) != 4:
            continue
        x1, y1, x2, y2 = float(box[0]) * sx, float(box[1]) * sy, float(box[2]) * sx, float(box[3]) * sy
        confidence = item.get("confidence", item.get("probability"))
        if isinstance(confidence, dict):
            confidence = confidence.get("average")
        line = _line(item.get("words", item.get("content", "")), [(x1, y1), (x2, y1), (x2, y2), (x1, y2)],
                     width, height, confidence)
        if line:
            lines.append(line)
    if not lines and str(payload.get("text") or "").strip():
        raise ValueError("GLM 返回非空文字但缺少布局坐标")
    return lines


def _post(url, **kwargs):
    timeout = kwargs.pop("timeout", settings.LAYOUT_OVERLAY_API_TIMEOUT_SECONDS)
    response = httpx.post(url, timeout=timeout, **kwargs)
    # 不输出包含 key 的 URL、请求体或原始响应。
    if response.is_error:
        raise RuntimeError(f"OCR 服务 HTTP {response.status_code}")
    payload = response.json()
    if payload.get("error") or payload.get("code"):
        raise RuntimeError("OCR 服务返回业务错误")
    return payload


class _HttpProvider:
    def __init__(self, timeout=None):
        self.timeout = timeout or settings.LAYOUT_OVERLAY_API_TIMEOUT_SECONDS


class GoogleVisionProvider(_HttpProvider):
    def recognize(self, image_bytes):
        key = settings.GOOGLE_VISION_API_KEY or settings.GOOGLE_API_KEY
        headers, params = {}, {}
        if key:
            params["key"] = key
        else:
            import google.auth
            from google.auth.transport.requests import Request
            credentials, project = google.auth.default(scopes=["https://www.googleapis.com/auth/cloud-platform"])
            request = Request()
            credentials.refresh(lambda *args, **kwargs: request(*args, **dict(kwargs, timeout=self.timeout)))
            headers = {"Authorization": f"Bearer {credentials.token}",
                       "x-goog-user-project": settings.VERTEX_PROJECT_ID or project or ""}
        payload = _post("https://vision.googleapis.com/v1/images:annotate", params=params, headers=headers, timeout=self.timeout,
                        json={"requests": [{"image": {"content": base64.b64encode(image_bytes).decode("ascii")},
                                            "features": [{"type": "DOCUMENT_TEXT_DETECTION"}]}]})
        return parse_google_response(payload, *_size(image_bytes))


class QwenOcrProvider(_HttpProvider):
    def recognize(self, image_bytes):
        if not settings.DASHSCOPE_API_KEY:
            raise ValueError("缺少 DASHSCOPE_API_KEY")
        # 自己显式缩放，确保供应商预处理不会改变返回坐标的参照系。
        width, height = _size(image_bytes)
        with Image.open(io.BytesIO(image_bytes)) as image:
            scale = max(math.sqrt(3072 / (width * height)), min(1, math.sqrt(8_000_000 / (width * height))))
            if scale != 1:
                image = image.resize((max(1, int(width * scale)), max(1, int(height * scale))))
            request_width, request_height = image.size
            output = io.BytesIO()
            image.convert("RGB").save(output, "PNG")
        payload = _post(settings.DASHSCOPE_BASE_URL.rstrip("/") + "/services/aigc/multimodal-generation/generation",
                        headers={"Authorization": f"Bearer {settings.DASHSCOPE_API_KEY}"}, timeout=self.timeout,
                        json={"model": settings.QWEN_OCR_MODEL,
                              "input": {"messages": [{"role": "user", "content": [{
                                  "image": "data:image/png;base64," + base64.b64encode(output.getvalue()).decode("ascii"),
                                  "min_pixels": 3072, "max_pixels": 8_388_608, "enable_rotate": False}]}]},
                              "parameters": {"ocr_options": {"task": "advanced_recognition"}}})
        lines = parse_qwen_response(payload, request_width, request_height)
        sx, sy = width / request_width, height / request_height
        for line in lines:
            line.polygon = [(x * sx, y * sy) for x, y in line.polygon]
            line.bbox = (line.bbox[0] * sx, line.bbox[1] * sy, line.bbox[2] * sx, line.bbox[3] * sy)
        return lines


class GlmOcrProvider(_HttpProvider):
    def recognize(self, image_bytes):
        if not settings.GLM_API_KEY:
            raise ValueError("缺少 GLM_API_KEY")
        payload = _post(settings.LAYOUT_OVERLAY_GLM_OCR_URL,
                        headers={"Authorization": f"Bearer {settings.GLM_API_KEY}"}, timeout=self.timeout,
                        json={"model": "glm-ocr", "file": "data:image/png;base64," + base64.b64encode(image_bytes).decode("ascii")})
        return parse_glm_response(payload, *_size(image_bytes))


PROVIDERS = {"google": GoogleVisionProvider, "qwen": QwenOcrProvider, "glm": GlmOcrProvider}


class OcrRouter:
    def __init__(self, primary=None, fallbacks=None):
        self.order = list(dict.fromkeys([primary or settings.LAYOUT_OVERLAY_OCR_PROVIDER] +
                                       (fallbacks if fallbacks is not None else settings.LAYOUT_OVERLAY_OCR_FALLBACKS.split(","))))
        self.order = [name.strip() for name in self.order if name.strip()]
        if any(name not in PROVIDERS for name in self.order):
            raise ValueError("OCR 引擎必须为 google、qwen 或 glm")
        self.last_provider = ""
        self.warnings = []

    def recognize(self, image_bytes, timeout=None, allow_empty=False):
        import time
        deadline = time.monotonic() + (timeout or settings.LAYOUT_OVERLAY_API_TIMEOUT_SECONDS * len(self.order))
        for name in self.order:
            try:
                remaining = deadline - time.monotonic()
                if remaining < 1:
                    raise TimeoutError("OCR 预算耗尽")
                lines = PROVIDERS[name](timeout=min(remaining, settings.LAYOUT_OVERLAY_API_TIMEOUT_SECONDS)).recognize(image_bytes)
                if not lines and not allow_empty:
                    raise ValueError("识别结果为空或缺少坐标")
                self.last_provider = name
                return lines
            except Exception as exc:
                self.warnings.append(f"{name} 识别失败（{type(exc).__name__}），尝试下一个引擎")
        raise RuntimeError("所有 OCR 引擎均失败，请检查 API 配置与图片质量")


def extract_pdf_text_lines(page, image_width: int, image_height: int) -> list[OcrLine]:
    import fitz
    sx, sy = image_width / page.rect.width, image_height / page.rect.height
    lines = []
    for block_index, block in enumerate(page.get_text("dict").get("blocks", [])):
        for item in block.get("lines", []):
            spans = [s for s in item.get("spans", []) if s.get("text", "").strip()]
            if not spans:
                continue
            rect = fitz.Rect(union_bbox([s["bbox"] for s in spans])) * page.rotation_matrix
            x1, y1, x2, y2 = rect.x0 * sx, rect.y0 * sy, rect.x1 * sx, rect.y1 * sy
            dominant = max(spans, key=lambda s: len(s["text"]))
            line = _line("".join(s["text"] for s in spans), [(x1, y1), (x2, y1), (x2, y2), (x1, y2)],
                         image_width, image_height, 1.0, font_size=dominant["size"],
                         color=f"{dominant.get('color', 0):06X}", bold=bool(dominant.get("flags", 0) & 16),
                         block_id=str(block_index))
            if line:
                lines.append(line)
    return lines
