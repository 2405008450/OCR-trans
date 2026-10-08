from __future__ import annotations

import json
import math
from dataclasses import dataclass, field
from pathlib import Path

BBox = tuple[float, float, float, float]
PROTECTED_KINDS = {"photo", "seal", "qr", "signature", "no_translate"}


def validate_bbox(value, width: float, height: float) -> BBox:
    if not isinstance(value, (list, tuple)) or len(value) != 4:
        raise ValueError("bbox 必须为 [x1, y1, x2, y2]")
    x1, y1, x2, y2 = map(float, value)
    if not all(math.isfinite(v) for v in (x1, y1, x2, y2)):
        raise ValueError("bbox 含非法数值")
    x1, y1, x2, y2 = max(0, x1), max(0, y1), min(width, x2), min(height, y2)
    if x2 <= x1 or y2 <= y1:
        raise ValueError("bbox 没有有效面积")
    return x1, y1, x2, y2


def union_bbox(boxes) -> BBox:
    return min(b[0] for b in boxes), min(b[1] for b in boxes), max(b[2] for b in boxes), max(b[3] for b in boxes)


def overlaps(a: BBox, b: BBox) -> bool:
    return min(a[2], b[2]) > max(a[0], b[0]) and min(a[3], b[3]) > max(a[1], b[1])


def parse_json(text: str):
    text = text.strip()
    if text.startswith("```"):
        text = text.split("\n", 1)[1].rsplit("```", 1)[0].strip()
    return json.loads(text)


@dataclass
class OcrLine:
    text: str
    polygon: list[tuple[float, float]]
    bbox: BBox
    confidence: float | None = None
    words: list[dict] = field(default_factory=list)
    font_size: float | None = None
    color: str | None = None
    bold: bool = False
    block_id: str = ""


@dataclass
class Segment:
    segment_id: str
    text: str
    bbox: BBox
    lines: list[OcrLine]
    kind: str = "text"
    translation: str = ""
    font_size: float = 10.0  # 单位为磅
    color: str = "000000"
    bold: bool = False
    font_scale: float = 1.0
    fitted_text: str = ""
    overflow: bool = False
    erase_padding: int = 2

    @property
    def protected(self) -> bool:
        return self.kind in PROTECTED_KINDS


@dataclass
class OverlayPage:
    image_path: Path
    width: int
    height: int
    width_pt: float
    height_pt: float
    lines: list[OcrLine]
    segments: list[Segment] = field(default_factory=list)
    protected_regions: list[dict] = field(default_factory=list)
    clean_path: Path | None = None
    ocr_provider: str = ""
