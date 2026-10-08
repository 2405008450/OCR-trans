"""文字擦除、原文样式估计与可重复的字号拟合。"""
from __future__ import annotations

import math
import os
from pathlib import Path

import cv2
import numpy as np
from PIL import ImageFont

from app.core.config import settings
from .models import OverlayPage, Segment, overlaps


def resolve_font(bold=False) -> tuple[str, str]:
    configured = settings.LAYOUT_OVERLAY_FONT_PATH
    if configured:
        candidates = [Path(configured)]
    else:
        fonts = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts"
        candidates = [fonts / ("arialbd.ttf" if bold else "arial.ttf"),
                      Path("/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf"),
                      Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")]
    for candidate in candidates:
        if candidate.is_file():
            return str(candidate), ImageFont.truetype(str(candidate), 12).getname()[0]
    raise FileNotFoundError("未找到可测量字体，请设置 LAYOUT_OVERLAY_FONT_PATH")


def _pixel_box(box, width, height, padding=0):
    return (max(0, math.floor(box[0]) - padding), max(0, math.floor(box[1]) - padding),
            min(width, math.ceil(box[2]) + padding), min(height, math.ceil(box[3]) + padding))


def estimate_style(image: np.ndarray, segment: Segment, page: OverlayPage):
    x1, y1, x2, y2 = _pixel_box(segment.lines[0].bbox, page.width, page.height)
    roi = image[y1:y2, x1:x2]
    if roi.size == 0:
        return
    pixels = roi.reshape(-1, 3)
    gray = cv2.cvtColor(roi, cv2.COLOR_RGB2GRAY)
    background = np.median(pixels, axis=0)
    distance = np.linalg.norm(roi.astype(float) - background, axis=2)
    ink = pixels[distance.reshape(-1) > max(30, np.percentile(distance, 75))]
    line = segment.lines[0]
    color = np.median(ink, axis=0).astype(int) if len(ink) else np.array([0, 0, 0])
    segment.color = line.color or "".join(f"{c:02X}" for c in color)
    segment.bold = line.bold or float(np.mean(gray < max(0, np.median(gray) - 40))) > 0.28
    segment.font_size = line.font_size or max(settings.LAYOUT_OVERLAY_MIN_FONT_PT,
                                            (y2 - y1) * page.height_pt / page.height * 0.85)


def erase_page(original: np.ndarray, page: OverlayPage) -> np.ndarray:
    """从原图重新生成底图，掩码减去所有保护区，修补不会累计损伤。"""
    protected = np.zeros(original.shape[:2], dtype=np.uint8)
    for region in page.protected_regions:
        x1, y1, x2, y2 = _pixel_box(region["bbox"], page.width, page.height)
        protected[y1:y2, x1:x2] = 255
    for segment in page.segments:
        if segment.protected or segment.translation.strip() == segment.text.strip():
            x1, y1, x2, y2 = _pixel_box(segment.bbox, page.width, page.height)
            protected[y1:y2, x1:x2] = 255
    clean = original.copy()
    for segment in page.segments:
        if segment.protected or segment.translation.strip() == segment.text.strip():
            continue
        for line in segment.lines:
            # inpaint 需要掩码外的已知像素；预留背景环，避免整块 ROI 都被遮罩。
            x1, y1, x2, y2 = _pixel_box(line.bbox, page.width, page.height, segment.erase_padding + 8)
            roi = original[y1:y2, x1:x2]
            if not roi.size:
                continue
            polygon = np.array([(x - x1, y - y1) for x, y in line.polygon], np.int32)
            kernel = np.ones((max(1, min(5, segment.erase_padding)) * 2 + 1,) * 2, np.uint8)
            allowed = np.zeros(roi.shape[:2], np.uint8)
            cv2.fillPoly(allowed, [polygon], 255)
            allowed = cv2.dilate(allowed, kernel)
            gray = cv2.cvtColor(roi, cv2.COLOR_RGB2GRAY)
            border = np.concatenate([roi[0], roi[-1], roi[:, 0], roi[:, -1]])
            background = np.median(border, axis=0)
            background_var = float(np.mean(np.var(border.astype(float), axis=0)))
            distance = np.linalg.norm(roi.astype(float) - background, axis=2)
            if background_var < 200:
                ink = (distance > 25).astype(np.uint8) * 255
            else:
                mode = cv2.THRESH_BINARY_INV if background.mean() > 128 else cv2.THRESH_BINARY
                ink = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, mode, 15, 8)
                # 粗笔画内部可能与局部均值相同，补充明显偏离底纹颜色的像素。
                border_distance = np.linalg.norm(border.astype(float) - background, axis=1)
                strong_ink = (distance > max(40, np.percentile(border_distance, 90) + 25)).astype(np.uint8) * 255
                ink = cv2.bitwise_or(ink, strong_ink)
            ink = cv2.dilate(ink, kernel)
            ink = cv2.bitwise_and(ink, allowed)
            ink[protected[y1:y2, x1:x2] > 0] = 0
            if background_var < 200:
                clean[y1:y2, x1:x2][ink > 0] = background.astype(np.uint8)
            else:
                local = cv2.inpaint(roi, ink, 3, cv2.INPAINT_TELEA)
                clean[y1:y2, x1:x2][ink > 0] = local[ink > 0]
    clean[protected > 0] = original[protected > 0]
    return clean


def wrap_text(text: str, font, width: float) -> list[str]:
    result = []
    for paragraph in text.split("\n"):
        if not paragraph:
            result.append("")
            continue
        current = ""
        for char in paragraph:
            if current and font.getlength(current + char) > width:
                # 优先在空格处分行；单个超长单词仍允许按字符换行。
                split = current.rfind(" ")
                if split > 0:
                    result.append(current[:split])
                    current = current[split + 1:] + char
                else:
                    result.append(current)
                    current = char
            else:
                current += char
        result.append(current.strip())
    return result


def fit_segment(segment: Segment, page: OverlayPage, *, min_font_pt=None) -> None:
    minimum = math.ceil(max(1.0, min_font_pt or settings.LAYOUT_OVERLAY_MIN_FONT_PT) * 2) / 2
    path, _ = resolve_font(segment.bold)
    px_per_pt = page.height / page.height_pt
    width = (segment.bbox[2] - segment.bbox[0]) * settings.LAYOUT_OVERLAY_FIT_WIDTH_RATIO
    height = segment.bbox[3] - segment.bbox[1]
    size = max(minimum, min(72.0, segment.font_size * segment.font_scale))
    while True:
        # 与 Word 的半磅字号保持一致。
        size = math.floor(size * 2) / 2
        font = ImageFont.truetype(path, max(1, round(size * px_per_pt)))
        lines = wrap_text(segment.translation, font, width)
        fits = (max((font.getlength(line) for line in lines), default=0) <= width and
                len(lines) * size * px_per_pt * 1.15 <= height)
        if fits or size <= minimum:
            segment.font_size = size
            segment.font_scale = 1.0
            segment.fitted_text = "\n".join(lines)
            segment.overflow = not fits
            return
        size = max(minimum, size - 0.5)


def safe_overlay(segment: Segment, page: OverlayPage) -> bool:
    return (not segment.protected and segment.translation.strip() != segment.text.strip() and
            not any(overlaps(segment.bbox, r["bbox"]) for r in page.protected_regions))
