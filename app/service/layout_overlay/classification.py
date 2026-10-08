import json
import re

from .models import PROTECTED_KINDS, Segment, overlaps, union_bbox, validate_bbox
from .translation import call_vision


def is_identifier(text):
    compact = text.strip()
    return bool(re.fullmatch(r"[0-9\s./:\-年月日]+", compact) or
                ("<" in compact and re.fullmatch(r"[A-Z0-9<\s]+", compact)) or
                (re.fullmatch(r"[A-Z0-9\-]+", compact) and any(c.isdigit() for c in compact) and len(compact) >= 5))


def classify_page(page, page_index, model, route):
    lines = [{"line_id": i, "text": line.text, "bbox": line.bbox, "block_id": line.block_id}
             for i, line in enumerate(page.lines)]
    prompt = (
        f"原图大小 {page.width}x{page.height}，坐标单位为原图像素，格式为 [x1,y1,x2,y2]。"
        "合并紧邻文字行为语义块；字段标签与字段值尽量分开，禁止跨列/远距离合并。"
        "保护人像、照片、印章、二维码、手写签名、证件编号、日期和 MRZ，类型依次为 "
        "photo/seal/qr/signature/no_translate。保护区包含全部相关像素。所有输入 line_id 必须恰好出现一次。"
        '输出 {"segments":[{"line_ids":[0],"kind":"text"}],'
        '"protected_regions":[{"kind":"photo","bbox":[0,0,10,10]}]}。'
        "不要改写文字或坐标，不要推断原图以外内容。OCR 数据：\n" + json.dumps(lines, ensure_ascii=False))
    result = call_vision(page.image_path.read_bytes(), prompt, model, route)
    if not isinstance(result, dict) or not isinstance(result.get("segments"), list) or not isinstance(result.get("protected_regions"), list):
        raise ValueError("版式分类响应格式不正确，已停止擦除以保留原图")
    regions = []
    for region in result["protected_regions"]:
        if not isinstance(region, dict) or region.get("kind") not in PROTECTED_KINDS:
            raise ValueError("保护区类型无效")
        regions.append({"kind": region["kind"], "bbox": validate_bbox(region.get("bbox"), page.width, page.height)})
    segments, seen = [], set()
    for item in result["segments"]:
        if not isinstance(item, dict):
            raise ValueError("语义块格式无效")
        ids, kind = item.get("line_ids"), item.get("kind", "text")
        if kind not in PROTECTED_KINDS | {"text"} or not isinstance(ids, list) or not ids:
            raise ValueError("语义块类型或 line_ids 无效")
        for lid in ids:
            if type(lid) is not int or lid < 0 or lid >= len(page.lines) or lid in seen:
                raise ValueError("版式分类包含重复/未知 line_id")
            seen.add(lid)
        selected = [page.lines[i] for i in ids]
        selected.sort(key=lambda line: (line.bbox[1], line.bbox[0]))
        # 大跨度或包含编号的合并回退到逐行，避免整块擦掉中间图像或编号。
        box = union_bbox([line.bbox for line in selected])
        occupied = sum((l.bbox[2] - l.bbox[0]) * (l.bbox[3] - l.bbox[1]) for l in selected)
        area = (box[2] - box[0]) * (box[3] - box[1])
        split = occupied < area * 0.45 or any(is_identifier(l.text) for l in selected)
        groups = [[line] for line in selected] if split else [selected]
        for group in groups:
            bbox = union_bbox([line.bbox for line in group])
            text = "\n".join(line.text for line in group)
            actual_kind = "no_translate" if is_identifier(text) else kind
            if any(overlaps(bbox, r["bbox"]) for r in regions):
                actual_kind = "no_translate"
            segment = Segment(f"p{page_index + 1}s{len(segments) + 1}", text, bbox, group, actual_kind)
            segments.append(segment)
    if seen != set(range(len(page.lines))):
        raise ValueError("版式分类遗漏文字行，已停止擦除以保留原图")
    page.segments, page.protected_regions = segments, regions
