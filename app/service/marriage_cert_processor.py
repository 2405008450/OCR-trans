"""
结婚证（Marriage Certificate）OCR翻译处理模块

功能：
1. 使用PaddleOCR自动识别图片文本
2. 从OCR结果中过滤置信度>=0.5的文本
3. 智能合并连续文本框（法律文书等连续性文本）
4. 使用DeepSeek模型进行中译英（智能映射 + 混合翻译 + LLM翻译）
5. 使用Lama/OpenCV智能涂抹原文区域
6. 将翻译后的英文文本回填到图片上

集成方式：
- 复用 image_processor.py 中的 OCR、DeepSeek、LaMa 等全局实例
- 通过 llm_service.py 的 doc_type 参数路由到本模块
- 四个布尔开关（merge/overlap_fix/colon_fix/use_lama）通过API参数控制
"""

import os
import json
import re
import time
import numpy as np
import cv2
from PIL import Image, ImageDraw, ImageFont
from typing import List, Dict, Any, Optional, Tuple

from app.core.config import settings

# 从 image_processor 导入共享资源和工具函数
from app.service.image_processor import (
    ocr, deepseek_client, lama_model, LAMA_AVAILABLE,
    add_watermark,
)

# 结婚证专用翻译配置
from app.service.marriage_cert_config import (
    MC_CERTIFICATE_KEYWORDS, MC_FIELD_TRANSLATIONS,
    MC_GENDER_TRANSLATIONS, MC_NATIONALITY_TRANSLATIONS,
    MC_LABEL_FIELDS, MC_MERGE_KEYWORDS,
)


# ============================
# 常量
# ============================
VALUE_FIELD_OFFSET = 40  # 标签右侧值框的右移像素数，避免与标签翻译重叠
PAGE1_SPECIAL_CN_TEXT = "中华人民共和国民政部监制"
PAGE1_SPECIAL_BOX_EXTRA_BOTTOM = 28  # 第一页特殊长文本下边界额外拉伸像素

# 第1页固定字段白名单（仅处理这些字段，其他OCR文本全部忽略）
PAGE1_ALLOWED_FIELDS = [
    "中华人民共和国民政部监制",
    "结婚申请，符合《中华人民",
    "共和国民法典》规定，予以登记，",
    "发给此证。",
    "婚姻登记员",
    "结婚申请，符合《中华人民"
]


# ============================
# 工具函数：Box 格式转换
# ============================

def _poly_to_rect(poly) -> List[int]:
    """
    多边形/不规则box → 矩形 [x_min, y_min, x_max, y_max]
    支持 [[x1,y1],[x2,y2],[x3,y3],[x4,y4]] 和 [x_min,y_min,x_max,y_max] 两种格式
    """
    if isinstance(poly, np.ndarray):
        poly = poly.tolist()
    # 已经是 rect 格式 [x_min, y_min, x_max, y_max]
    if len(poly) == 4 and not isinstance(poly[0], (list, tuple)):
        return [int(v) for v in poly]
    # polygon 格式 [[x1,y1],[x2,y2],...]
    xs = [p[0] for p in poly]
    ys = [p[1] for p in poly]
    return [int(min(xs)), int(min(ys)), int(max(xs)), int(max(ys))]


def _rect_to_poly(rect) -> List[List[int]]:
    """矩形 [x_min, y_min, x_max, y_max] → 多边形 [[x1,y1],[x2,y2],[x3,y3],[x4,y4]]"""
    x_min, y_min, x_max, y_max = [int(v) for v in rect]
    return [[x_min, y_min], [x_max, y_min], [x_max, y_max], [x_min, y_max]]


# ============================
# OCR 数据提取
# ============================

def _extract_ocr_data(result) -> Dict[str, list]:
    """
    从PaddleOCR结果中提取OCR数据，统一为 rect 格式 [x_min, y_min, x_max, y_max]。
    兼容 PaddleOCR 2.x / 3.x 多种返回格式。
    """
    all_texts: List[str] = []
    all_scores: List[float] = []
    all_boxes: List[List[int]] = []

    for elem in result:
        texts: list = []
        scores: list = []
        boxes: list = []

        if isinstance(elem, dict):
            res = elem.get("res", elem)
            texts = list(res.get("rec_texts", []))
            raw_scores = res.get("rec_scores", [])
            scores = raw_scores.tolist() if hasattr(raw_scores, 'tolist') else list(raw_scores)

            # 优先取 rec_boxes（rect格式），其次取 rec_polys 转换
            raw_boxes = res.get("rec_boxes")
            if raw_boxes is not None:
                if hasattr(raw_boxes, 'tolist'):
                    raw_boxes = raw_boxes.tolist()
                boxes = [_poly_to_rect(b) for b in raw_boxes]
            else:
                raw_polys = res.get("rec_polys", [])
                if hasattr(raw_polys, 'tolist'):
                    raw_polys = raw_polys.tolist()
                boxes = [_poly_to_rect(p) for p in raw_polys]

        elif hasattr(elem, 'rec_texts'):
            # PaddleOCR 3.x OCRResult 对象
            texts = list(getattr(elem, 'rec_texts', []))
            raw_scores = getattr(elem, 'rec_scores', [])
            scores = raw_scores.tolist() if hasattr(raw_scores, 'tolist') else list(raw_scores)
            raw_boxes = getattr(elem, 'rec_boxes', None)
            if raw_boxes is not None:
                if hasattr(raw_boxes, 'tolist'):
                    raw_boxes = raw_boxes.tolist()
                boxes = [_poly_to_rect(b) for b in raw_boxes]
            else:
                raw_polys = getattr(elem, 'rec_polys', [])
                if hasattr(raw_polys, 'tolist'):
                    raw_polys = raw_polys.tolist()
                boxes = [_poly_to_rect(p) for p in raw_polys]

        elif isinstance(elem, (list, tuple)):
            # PaddleOCR 2.x 旧格式: [[[box], (text, score)], ...]
            for row in elem:
                if isinstance(row, (list, tuple)) and len(row) >= 2:
                    box_data = row[0]
                    pair = row[1]
                    if isinstance(pair, (list, tuple)) and len(pair) >= 2:
                        text = pair[0]
                        score = float(pair[1])
                    else:
                        text = str(pair)
                        score = 0.0
                    rect = _poly_to_rect(box_data)
                    texts.append(text)
                    scores.append(score)
                    boxes.append(rect)

        all_texts.extend(texts)
        all_scores.extend(scores)
        all_boxes.extend(boxes)

    return {
        'rec_texts': all_texts,
        'rec_scores': all_scores,
        'rec_boxes': all_boxes
    }


# ============================
# 置信度过滤
# ============================

def _filter_by_confidence(ocr_data: dict, threshold: float = 0.5) -> List[Dict]:
    """
    过滤置信度 < threshold 的文本
    返回 [{index, text, score, box}, ...]
    """
    rec_texts = ocr_data['rec_texts']
    rec_scores = ocr_data['rec_scores']
    rec_boxes = ocr_data['rec_boxes']

    filtered = []
    for i, (text, score, box) in enumerate(zip(rec_texts, rec_scores, rec_boxes)):
        if score >= threshold:
            filtered.append({
                'index': i,
                'text': text,
                'score': score,
                'box': box  # [x_min, y_min, x_max, y_max]
            })

    print(f"   原始文本: {len(rec_texts)} 个，过滤后: {len(filtered)} 个 (阈值>={threshold})")
    dropped = [(t, f"{s:.4f}") for t, s in zip(rec_texts, rec_scores) if s < threshold]
    if dropped:
        print(f"   被过滤:")
        for text, score in dropped:
            print(f"     - {text} (置信度: {score})")

    return filtered


def _normalize_cn_for_match(text: str) -> str:
    """
    中文文本匹配归一化：
    - 去除空白、常见中英文标点、引号、星号
    - 用于第一页固定字段的容错匹配
    """
    if not text:
        return ""
    normalized = re.sub(r'[\s,，。；;：:《》“”"\'`·•★\-\(\)\[\]{}]+', '', text)
    return normalized.strip()


def _filter_page1_fixed_fields(filtered_data: List[Dict]) -> List[Dict]:
    """
    第1页固定字段过滤：
    仅保留白名单字段相关文本，其他OCR项全部忽略。
    """
    targets = [(_t, _normalize_cn_for_match(_t)) for _t in PAGE1_ALLOWED_FIELDS]
    kept: List[Dict] = []
    dropped_count = 0
    min_partial_len = 4  # 只允许长度>=4的片段参与模糊匹配，过滤“★”“婚”“证”等单字噪音

    for item in filtered_data:
        text = item.get("text", "")
        norm_text = _normalize_cn_for_match(text)
        keep = False

        # 空串或极短内容直接丢弃（常见噪音）
        if len(norm_text) < 2:
            dropped_count += 1
            continue

        for _, norm_target in targets:
            # 规则1: 完整命中
            if norm_text == norm_target:
                keep = True
                break

            # 规则2: 仅对较长文本做双向包含，兼容 OCR 截断/粘连
            if len(norm_text) >= min_partial_len and (norm_target in norm_text or norm_text in norm_target):
                keep = True
                break

        if keep:
            kept.append(item)
        else:
            dropped_count += 1

    print(f"   第1页白名单过滤: 保留 {len(kept)} 个，忽略 {dropped_count} 个")
    for it in kept:
        print(f"     ✓ {it.get('text', '')}")
    return kept


# ============================
# 文本合并
# ============================

def _should_merge_texts(text_group: List[str]) -> bool:
    """判断一组文本是否需要合并翻译"""
    combined = ''.join(text_group)

    # 包含合并关键词
    for keyword in MC_MERGE_KEYWORDS:
        if keyword in combined:
            return True

    # 未完成的句子（以连接词/标点结尾）
    for text in text_group[:-1]:
        if text.rstrip().endswith(('，', '、', '；', '的', '和', '与', '及')):
            return True

    return False


def _are_boxes_vertically_adjacent(boxes: List[List[int]], max_gap: int = 30) -> bool:
    """检查文本框是否在垂直方向上相邻"""
    if len(boxes) < 2:
        return True
    sorted_boxes = sorted(boxes, key=lambda b: b[1])
    for i in range(len(sorted_boxes) - 1):
        gap = sorted_boxes[i + 1][1] - sorted_boxes[i][3]
        if gap > max_gap:
            return False
    return True


def _are_boxes_horizontally_related(
    boxes: List[List[int]],
    min_overlap_ratio: float = 0.2,
    max_left_shift: int = 80
) -> bool:
    """
    检查文本框是否属于同一列/段落的横向关系。
    避免“仅上下相邻”导致跨列误合并（如“婚姻登记员”与“中华人民共和国民政部监制”）。
    """
    if len(boxes) < 2:
        return True

    sorted_boxes = sorted(boxes, key=lambda b: b[1])
    for i in range(len(sorted_boxes) - 1):
        b1 = sorted_boxes[i]
        b2 = sorted_boxes[i + 1]

        x1_min, _, x1_max, _ = b1
        x2_min, _, x2_max, _ = b2

        w1 = max(1, x1_max - x1_min)
        w2 = max(1, x2_max - x2_min)
        overlap = max(0, min(x1_max, x2_max) - max(x1_min, x2_min))
        overlap_ratio = overlap / min(w1, w2)

        # 有重叠，或左边界接近（轻微缩进）都视为同一列
        if overlap_ratio >= min_overlap_ratio:
            continue
        if abs(x1_min - x2_min) <= max_left_shift:
            continue
        return False

    return True


def _merge_consecutive_texts(filtered_data: List[Dict]) -> List[Dict]:
    """
    将需要合并的连续文本合并成组。
    返回合并后的数据列表（包含 is_merged / sub_boxes 标记）。
    """
    if not filtered_data:
        return []

    print("\n" + "=" * 60)
    print("分析文本合并需求...")
    print("=" * 60)

    merged_groups = []
    i = 0

    while i < len(filtered_data):
        best_group = [filtered_data[i]]
        best_group_size = 1

        # 向前查看最多 10 行，寻找最大可合并组
        for look_ahead in range(2, min(11, len(filtered_data) - i + 1)):
            text_group = [filtered_data[i + j]['text'] for j in range(look_ahead)]
            if _should_merge_texts(text_group):
                boxes = [filtered_data[i + j]['box'] for j in range(look_ahead)]
                if _are_boxes_vertically_adjacent(boxes) and _are_boxes_horizontally_related(boxes):
                    best_group = filtered_data[i:i + look_ahead]
                    best_group_size = look_ahead

        if best_group_size > 1:
            merged_text = ''.join([item['text'] for item in best_group])
            merged_boxes = [item['box'] for item in best_group]
            x_mins = [b[0] for b in merged_boxes]
            y_mins = [b[1] for b in merged_boxes]
            x_maxs = [b[2] for b in merged_boxes]
            y_maxs = [b[3] for b in merged_boxes]
            merged_box = [min(x_mins), min(y_mins), max(x_maxs), max(y_maxs)]

            merged_groups.append({
                'text': merged_text,
                'box': merged_box,
                'merged_indices': [item['index'] for item in best_group],
                'sub_boxes': merged_boxes,
                'is_merged': True
            })

            print(f"  合并 ({best_group_size}行):")
            for item in best_group:
                print(f"    - {item['text']}")
            print(f"    → {merged_text}")
            i += best_group_size
        else:
            merged_groups.append({
                'text': filtered_data[i]['text'],
                'box': filtered_data[i]['box'],
                'merged_indices': [filtered_data[i]['index']],
                'sub_boxes': [filtered_data[i]['box']],
                'is_merged': False
            })
            i += 1

    print(f"\n  原始: {len(filtered_data)} 个文本块，合并后: {len(merged_groups)} 个")
    print("=" * 60)

    return merged_groups


# ============================
# 翻译相关
# ============================

def _llm_translate(prompt: str, max_tokens: int = 1024) -> Optional[str]:
    """调用DeepSeek LLM翻译"""
    try:
        response = deepseek_client.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.2,
            max_tokens=max_tokens
        )
        return response.choices[0].message.content.strip().strip('"').strip("'")
    except Exception as e:
        print(f"    [LLM调用失败] {e}")
        return None


def _convert_date_format(chinese_date: str) -> str:
    """转换中文日期为英文格式。如 "2023年05月15日" → "May 15, 2023" """
    month_names = [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ]

    # YYYY年MM月DD日
    match = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', chinese_date)
    if match:
        year, month, day = match.groups()
        return f"{month_names[int(month) - 1]} {int(day)}, {year}"

    # YYYYMMDD
    match = re.search(r'(\d{4})(\d{2})(\d{2})', chinese_date)
    if match:
        year, month, day = match.groups()
        return f"{month_names[int(month) - 1]} {int(day)}, {year}"

    return chinese_date


def _smart_translate_field(chinese_text: str) -> Optional[str]:
    """
    智能翻译证件字段（使用直接映射，不调用LLM）。
    成功则返回英文翻译，无法识别返回 None（需交给 LLM 处理）。
    """
    for field_cn, field_en in MC_FIELD_TRANSLATIONS.items():
        if not chinese_text.startswith(field_cn):
            continue

        value = chinese_text[len(field_cn):].strip()

        # 无值：证件号类不加冒号，其他加冒号
        if not value:
            if field_en.rstrip('.').endswith('No'):
                return field_en
            return f"{field_en}:"

        # 性别
        if '性别' in field_cn:
            value_en = MC_GENDER_TRANSLATIONS.get(value, value)
            return f"{field_en}: {value_en}"

        # 国籍
        if '国籍' in field_cn:
            value_en = MC_NATIONALITY_TRANSLATIONS.get(value, value)
            return f"{field_en}: {value_en}"

        # 日期
        if '日期' in field_cn or '出生' in field_cn:
            value_en = _convert_date_format(value)
            return f"{field_en}: {value_en}"

        # 结婚证字号 + 值 → 需要 LLM（混合翻译）
        if field_cn in ('结婚证字号', '结婚证号') and value:
            return None

        # 证件号（保持不变）
        if '号' in field_cn or '身份证' in field_cn:
            if value.startswith(':'):
                value = value[1:].strip()
            if field_en.endswith('No'):
                return f"{field_en}. {value}"
            return f"{field_en}: {value}"

        # 备注
        if '备注' in field_cn:
            if not value or value in ['无', '空', '-']:
                return f"{field_en}:"
            return f"{field_en}: {value}"

        # 持证人/姓名含人名 → 需要 LLM
        if field_cn in ('持证人', '姓名') and value:
            return None

        return f"{field_en}: {value}"

    return None


def _try_hybrid_translate(chinese_text: str) -> Optional[Tuple[str, str, str]]:
    """
    尝试混合翻译：字段名走映射，值走 LLM。
    返回 (field_en, value, field_cn) 或 None。
    """
    hybrid_fields = ['持证人', '姓名', '结婚证字号', '结婚证号']

    for field_cn in hybrid_fields:
        if chinese_text.startswith(field_cn):
            value = chinese_text[len(field_cn):].strip()
            if not value:
                return None
            field_en = MC_FIELD_TRANSLATIONS.get(field_cn)
            if not field_en:
                return None
            return (field_en, value, field_cn)

    return None


def _translate_value_only(value: str, field_cn: str = '') -> str:
    """仅翻译/转写值部分（用于混合翻译）"""
    if field_cn in ('持证人', '姓名'):
        prompt = (
            f"请将以下中文人名转写为英文拼音格式（首字母大写）。"
            f"只返回转写结果，不要添加任何解释。\n\n"
            f"中文人名：{value}\n\n英文："
        )
    elif field_cn in ('结婚证字号', '结婚证号'):
        prompt = (
            f"请保持以下证件号不变，原样输出。只返回证件号本身，不要添加任何解释。\n\n"
            f"证件号：{value}\n\n输出："
        )
    else:
        prompt = (
            f"请将以下内容翻译或转写为英文。保持数字、代码不变。只返回结果，不要添加解释。\n\n"
            f"内容：{value}\n\n英文："
        )

    result = _llm_translate(prompt, max_tokens=256)
    return result if result else value


def _translate_manual_signature_text(signature_text: str) -> str:
    """
    翻译/转写婚姻登记员手写签名（用户手动输入）。
    - 中文姓名优先转写为拼音（首字母大写）
    - 英文或代码样式内容尽量保持
    """
    if not signature_text:
        return signature_text

    prompt = (
        "请将以下“婚姻登记员手写签名文本”翻译或转写为英文。\n"
        "规则：\n"
        "1) 若是中文姓名，请转写为拼音并首字母大写\n"
        "2) 若原文本身是英文/字母缩写，尽量保持原样\n"
        "3) 只返回最终结果，不要解释\n\n"
        f"文本：{signature_text}\n\n英文："
    )
    result = _llm_translate(prompt, max_tokens=80)
    return result if result else signature_text


def _translate_registered_by_text(registered_by_text: str) -> str:
    """
    翻译/转写 Registered by 的 xxx 文本（用户手动输入）。
    输出用于：Registered by: (xxx)
    """
    if not registered_by_text:
        return registered_by_text

    prompt = (
        "请将以下“Registered by 字段中的姓名/文本”翻译或转写为英文。\n"
        "规则：\n"
        "1) 若是中文姓名，请转写为拼音并首字母大写\n"
        "2) 若原文本身是英文/缩写，尽量保持原样\n"
        "3) 只返回结果文本，不要加解释\n\n"
        f"文本：{registered_by_text}\n\n英文："
    )
    result = _llm_translate(prompt, max_tokens=80)
    return result if result else registered_by_text


def _fix_missing_colons(english_text: str, chinese_text: str) -> str:
    """自动修正翻译结果中缺失的冒号"""
    is_certificate = any(kw in chinese_text for kw in MC_CERTIFICATE_KEYWORDS)
    if not is_certificate:
        return english_text

    all_fields = set(MC_FIELD_TRANSLATIONS.values()) | {
        'Name', 'Gender', 'Nationality', 'ID Number', 'Date of Birth',
        'Registration Date', 'Registration', 'Certificate No',
        'Marriage Certificate No', 'Remarks', 'Certificate Holder',
        'Address', 'Tel', 'Phone', 'Email', 'Ethnicity', 'Blood Type'
    }

    fixed_text = english_text
    for field in all_fields:
        # 模式1: "字段名 值" → "字段名: 值"
        pattern1 = re.compile(rf'\b({re.escape(field)})\s+([A-Z][a-zA-Z])', re.IGNORECASE)
        match1 = pattern1.search(fixed_text)
        if match1 and ':' not in match1.group(0):
            fixed_text = pattern1.sub(r'\1: \2', fixed_text)
            print(f"    [冒号修正] 添加冒号: {field}")

        # 模式2: 单独字段名末尾
        pattern2 = re.compile(rf'\b({re.escape(field)})$', re.IGNORECASE)
        if pattern2.search(fixed_text):
            fixed_text = pattern2.sub(r'\1:', fixed_text)

    return fixed_text


def _translate_all(merged_data: List[Dict], enable_colon_fix: bool = True) -> List[Dict]:
    """
    智能翻译所有文本块：
    1. 证件固定字段 → 直接映射（不调用LLM）
    2. 含人名/证件号的字段 → 混合翻译（字段名映射 + 值走LLM）
    3. 复杂文本 → 完整LLM翻译
    """
    print("\n" + "=" * 60)
    print("开始智能翻译...")
    print("=" * 60)

    total = len(merged_data)
    direct_count = 0
    llm_count = 0

    for i, data in enumerate(merged_data):
        chinese_text = data['text']

        # ---- 1. 尝试直接映射 ----
        english_text = _smart_translate_field(chinese_text)

        if english_text:
            data['translation'] = english_text
            data['original_text'] = chinese_text
            direct_count += 1
            is_merged_label = "(合并)" if data.get('is_merged') else ""
            print(f"[{i+1}/{total}] [映射] {is_merged_label} {chinese_text} → {english_text}")
            continue

        # ---- 2. 尝试混合翻译 ----
        hybrid = _try_hybrid_translate(chinese_text)

        if hybrid:
            field_en, value, field_cn = hybrid
            try:
                value_en = _translate_value_only(value, field_cn)
                if field_en.rstrip('.').endswith('No'):
                    english_text = f"{field_en} {value_en}"
                else:
                    english_text = f"{field_en}: {value_en}"
                data['translation'] = english_text
                data['original_text'] = chinese_text
                llm_count += 1
                is_merged_label = "(合并)" if data.get('is_merged') else ""
                print(f"[{i+1}/{total}] [混合] {is_merged_label} {chinese_text}")
                print(f"                   字段: {field_cn} → {field_en}")
                print(f"                   值: {value} → {value_en}")
                print(f"                   译文: {english_text}")
            except Exception as e:
                print(f"[{i+1}/{total}] 混合翻译失败: {chinese_text}, 错误: {e}")
                data['translation'] = chinese_text
                data['original_text'] = chinese_text
            continue

        # ---- 3. 完整 LLM 翻译 ----
        prompt = (
            f"请将以下中文翻译成英文。\n"
            f"要求：简洁、准确，符合英文表达习惯，只返回翻译结果。\n\n"
            f"中文：{chinese_text}\n\n英文："
        )
        result = _llm_translate(prompt)

        if result:
            if enable_colon_fix:
                result = _fix_missing_colons(result, chinese_text)
            data['translation'] = result
        else:
            data['translation'] = chinese_text

        data['original_text'] = chinese_text
        llm_count += 1
        is_merged_label = "(合并)" if data.get('is_merged') else ""
        print(f"[{i+1}/{total}] [LLM] {is_merged_label} {chinese_text} → {data['translation']}")
        time.sleep(0.2)  # 避免 API 限流

    print("\n" + "=" * 60)
    pct_direct = direct_count / total * 100 if total else 0
    pct_llm = llm_count / total * 100 if total else 0
    print(f"翻译统计:")
    print(f"  - 直接映射: {direct_count}/{total} ({pct_direct:.1f}%)")
    print(f"  - LLM翻译: {llm_count}/{total} ({pct_llm:.1f}%)")
    print(f"  - 节省API调用: {direct_count} 次")
    print("=" * 60 + "\n")

    return merged_data


# ============================
# 图像涂抹
# ============================

def _inpaint_regions(image: np.ndarray, filtered_data: List[Dict], use_lama: bool = True) -> np.ndarray:
    """使用 LaMa 或 OpenCV 涂抹文本区域"""
    print("\n涂抹文本区域...")

    mask = np.zeros(image.shape[:2], dtype=np.uint8)
    padding = 5

    for data in filtered_data:
        box = data['box']
        x_min, y_min, x_max, y_max = [int(v) for v in box]
        cv2.rectangle(
            mask,
            (max(0, x_min - padding), max(0, y_min - padding)),
            (min(image.shape[1], x_max + padding), min(image.shape[0], y_max + padding)),
            255, -1
        )

    if use_lama and LAMA_AVAILABLE and lama_model is not None:
        try:
            img_rgb = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
            img_pil = Image.fromarray(img_rgb)
            mask_pil = Image.fromarray(mask)
            result_pil = lama_model(img_pil, mask_pil)
            inpainted = cv2.cvtColor(np.array(result_pil), cv2.COLOR_RGB2BGR)
            print("   LaMa 涂抹完成")
            return inpainted
        except Exception as e:
            print(f"   LaMa 失败，回退到 OpenCV: {e}")

    inpainted = cv2.inpaint(image, mask, inpaintRadius=5, flags=cv2.INPAINT_TELEA)
    print("   OpenCV 涂抹完成")
    return inpainted


# ============================
# 字体加载
# ============================

def _load_font(font_path: Optional[str], size: int) -> ImageFont.FreeTypeFont:
    """加载字体，优先使用指定路径，否则尝试系统字体"""
    candidates = [
        font_path,
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/Arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
    ]
    for path in candidates:
        if path and os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                continue
    try:
        return ImageFont.truetype("arial.ttf", size)
    except Exception:
        return ImageFont.load_default()


# ============================
# 文本绘制
# ============================

def _calculate_box_offsets(merged_data: List[Dict], y_threshold: int = 20, min_x_gap: int = 100) -> Dict[int, int]:
    """
    检测box间的潜在重叠，为右侧box设置固定偏移量。
    返回 {box_index: left_offset}
    """
    offsets: Dict[int, int] = {}

    single_boxes = []
    for i, data in enumerate(merged_data):
        if not data.get('is_merged', False) or len(data.get('sub_boxes', [])) <= 1:
            single_boxes.append((i, data['box']))

    for i, (idx_i, box_i) in enumerate(single_boxes):
        y1_center = (box_i[1] + box_i[3]) / 2

        for j, (idx_j, box_j) in enumerate(single_boxes):
            if i >= j:
                continue

            y2_center = (box_j[1] + box_j[3]) / 2

            # 同一水平线
            if abs(y1_center - y2_center) > y_threshold:
                continue
            # box_j 在 box_i 右侧
            if box_j[0] <= box_i[2]:
                continue
            # 间距过小
            x_gap = box_j[0] - box_i[2]
            if x_gap < min_x_gap:
                offsets[idx_j] = VALUE_FIELD_OFFSET

    return offsets


def _draw_single_text(
    draw: ImageDraw.ImageDraw, text: str, box: List[int],
    font_path: Optional[str], font_size: Optional[int] = None, left_offset: int = 5,
    force_wrap: bool = False
):
    """在单个框中绘制文本"""
    x_min, y_min, x_max, y_max = [int(v) for v in box]
    box_height = y_max - y_min
    box_width = x_max - x_min

    adjusted_size = font_size if font_size else max(8, int(box_height * 0.65))
    font = _load_font(font_path, adjusted_size)

    try:
        bbox = draw.textbbox((0, 0), text, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
    except AttributeError:
        text_width = int(len(text) * adjusted_size * 0.6)
        text_height = adjusted_size

    available_width = max(40, box_width - left_offset - 5)
    need_wrap = force_wrap or (text_width > available_width)

    if not need_wrap:
        x_pos = x_min + left_offset
        y_pos = y_min + (box_height - text_height) // 2
        draw.text((x_pos, y_pos), text, font=font, fill=(0, 0, 0))
        return

    # 单框自动换行（用于第一页长文本等场景）
    lines = _wrap_text_for_width(draw, text, font, available_width)
    line_h = _text_size(draw, "Ay", font)[1]
    line_spacing = max(2, int(line_h * 0.15))
    total_h = len(lines) * line_h + (len(lines) - 1) * line_spacing
    y_pos = y_min + max(2, (box_height - total_h) // 2)

    for line in lines:
        draw.text((x_min + left_offset, y_pos), line, font=font, fill=(0, 0, 0))
        y_pos += line_h + line_spacing


def _apply_page1_special_adjustments(merged_data: List[Dict]) -> None:
    """
    第一页特殊优化：
    - 命中“中华人民共和国民政部监制”时，强制英文换行回填
    - 适度下拉 box/sub_boxes 的下边界，提供更多绘制高度
    """
    for data in merged_data:
        original_text = data.get('original_text', data.get('text', ''))
        if PAGE1_SPECIAL_CN_TEXT not in original_text:
            continue

        data['force_wrap'] = True

        # 下拉主框下边界
        if isinstance(data.get('box'), list) and len(data['box']) == 4:
            data['box'][3] = int(data['box'][3]) + PAGE1_SPECIAL_BOX_EXTRA_BOTTOM

        # 下拉子框下边界
        if isinstance(data.get('sub_boxes'), list):
            for sb in data['sub_boxes']:
                if isinstance(sb, list) and len(sb) == 4:
                    sb[3] = int(sb[3]) + PAGE1_SPECIAL_BOX_EXTRA_BOTTOM

        print(
            f"   第一页特殊优化已应用: '{PAGE1_SPECIAL_CN_TEXT}' "
            f"(下边界+{PAGE1_SPECIAL_BOX_EXTRA_BOTTOM}px, 强制换行)"
        )
        break


def _draw_merged_text(
    draw: ImageDraw.ImageDraw, text: str, sub_boxes: List[List[int]],
    font_path: Optional[str], font_size: Optional[int] = None
):
    """将合并后的翻译文本在合并大框中重新排版（自动换行）"""
    if not sub_boxes:
        return

    sorted_boxes = sorted(sub_boxes, key=lambda b: b[1])

    x_min = min(b[0] for b in sorted_boxes)
    y_min = min(b[1] for b in sorted_boxes)
    x_max = max(b[2] for b in sorted_boxes)
    y_max = max(b[3] for b in sorted_boxes)

    merged_width = int(x_max - x_min)
    merged_height = int(y_max - y_min)

    if font_size:
        actual_size = font_size
    else:
        avg_h = sum(b[3] - b[1] for b in sorted_boxes) / len(sorted_boxes)
        actual_size = max(8, int(avg_h * 0.65))

    font = _load_font(font_path, actual_size)

    try:
        sample_bbox = draw.textbbox((0, 0), "Ay", font=font)
        text_height = sample_bbox[3] - sample_bbox[1]
    except AttributeError:
        text_height = actual_size

    line_spacing = int(text_height * 0.2)
    available_width = merged_width - 10  # 左右留白

    # 按单词自动换行
    words = text.split()
    lines: List[str] = []

    while words:
        line_words: List[str] = []
        while words:
            test_line = ' '.join(line_words + [words[0]])
            try:
                bbox = draw.textbbox((0, 0), test_line, font=font)
                test_width = bbox[2] - bbox[0]
            except AttributeError:
                test_width = len(test_line) * actual_size * 0.6

            if test_width <= available_width:
                line_words.append(words.pop(0))
            else:
                break

        # 防止死循环：至少放一个单词
        if not line_words and words:
            line_words.append(words.pop(0))

        if line_words:
            lines.append(' '.join(line_words))

    # 计算总高度并检查是否超出
    total_lines = len(lines)
    required_height = total_lines * text_height + (total_lines - 1) * line_spacing

    if required_height > merged_height and total_lines > 1:
        # 压缩行间距
        adjusted_spacing = max(0, (merged_height - total_lines * text_height) // (total_lines - 1))
        line_spacing = adjusted_spacing
        print(f"    压缩行间距至 {adjusted_spacing}px (高度不足)")

    # 绘制每一行
    current_y = int(y_min) + 5
    for line_text in lines:
        draw.text((int(x_min) + 5, current_y), line_text, font=font, fill=(0, 0, 0))
        current_y += text_height + line_spacing

    print(f"    分为 {total_lines} 行绘制")


def _draw_translated_texts(
    inpainted_image: np.ndarray, merged_data: List[Dict],
    font_path: Optional[str] = None, font_size: int = 18,
    enable_overlap_fix: bool = True,
    manual_registrar_signature_en: Optional[str] = None,
    manual_registered_by_en: Optional[str] = None,
    registered_by_offset_x: int = 0,
    registered_by_offset_y: int = 0,
    registrar_signature_offset_x: int = 36,
    registrar_signature_offset_y: int = -12
) -> np.ndarray:
    """将翻译后的文本绘制到涂抹后的图片上"""
    print("\n回填翻译文本...")

    img_h, img_w = inpainted_image.shape[:2]
    pil_image = Image.fromarray(cv2.cvtColor(inpainted_image, cv2.COLOR_BGR2RGB))
    draw = ImageDraw.Draw(pil_image)

    # 计算重叠偏移
    box_offsets = _calculate_box_offsets(merged_data) if enable_overlap_fix else {}
    signature_drawn = False

    for i, data in enumerate(merged_data):
        translated_text = data.get('translation', '')
        if not translated_text:
            continue

        if data.get('is_merged', False) and len(data['sub_boxes']) > 1:
            _draw_merged_text(draw, translated_text, data['sub_boxes'], font_path, font_size)
            orig = data.get('original_text', data['text'])
            print(f"  回填(合并): '{translated_text[:60]}' ← '{orig[:30]}'")
        else:
            left_offset = box_offsets.get(i, 5)
            _draw_single_text(
                draw,
                translated_text,
                data['box'],
                font_path,
                font_size,
                left_offset,
                force_wrap=bool(data.get('force_wrap', False))
            )
            orig = data.get('original_text', data['text'])
            suffix = f" (右移{left_offset}px)" if left_offset > 5 else ""
            print(f"  回填: '{translated_text[:60]}' ← '{orig[:30]}'{suffix}")

            # 若用户手动输入了“婚姻登记员签名”，则在“婚姻登记员”右侧补填
            if (
                manual_registrar_signature_en
                and not signature_drawn
                and ("婚姻登记员" in orig)
            ):
                x_min, y_min, x_max, y_max = [int(v) for v in data['box']]
                sign_w = max(140, int((x_max - x_min) * 2.2))
                sign_box = [
                    x_max + max(0, int(registrar_signature_offset_x)),
                    max(0, y_min - 2 + int(registrar_signature_offset_y)),
                    min(
                        img_w - 8,
                        x_max + max(0, int(registrar_signature_offset_x)) + sign_w
                    ),
                    min(img_h - 4, y_max + 12 + int(registrar_signature_offset_y)),
                ]

                # 在签名框上方补填：
                # 第一行: Registered by:
                # 第二行: (xxx)  红色斜体，并自动换行避免超出右边界
                if manual_registered_by_en:
                    label_text = "Registered by:"
                    value_text = f"({manual_registered_by_en})"
                    label_font = _get_font_regular(max(12, int(font_size * 0.85)))
                    value_font = _get_font_italic(max(12, int(font_size * 0.9)))
                    _, label_h = _text_size(draw, label_text, label_font)
                    value_line_h = _text_size(draw, "Ay", value_font)[1]
                    reg_y = max(0, sign_box[1] - max(26, int(font_size * 1.7)))
                    reg_x = sign_box[0]
                    max_width = max(60, img_w - reg_x - 8)

                    # 第一行：标签
                    draw.text((reg_x, reg_y), label_text, font=label_font, fill=(0, 0, 0))

                    # 第二行：(xxx) 自动换行
                    value_y = reg_y + label_h + 2
                    wrapped_lines = _wrap_text_for_width(draw, value_text, value_font, max_width)
                    for line in wrapped_lines:
                        draw.text((reg_x, value_y), line, font=value_font, fill=(255, 0, 0))
                        value_y += value_line_h + 2

                    print(f"  Registered by补填: {label_text} + {len(wrapped_lines)} 行手写文本")

                _draw_single_text(
                    draw,
                    manual_registrar_signature_en,
                    sign_box,
                    font_path,
                    font_size,
                    left_offset=2,
                    force_wrap=False
                )
                signature_drawn = True
                print(
                    f"  手动签名补填: '{manual_registrar_signature_en}' "
                    f"(右侧补填, offset_x={registrar_signature_offset_x}px, offset_y={registrar_signature_offset_y}px)"
                )

    final_image = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)
    print("文本回填完成")
    return final_image


# ============================
# 可视化
# ============================

def _try_paddle_vis_native(result, img_path: str, output_dir: str) -> Optional[str]:
    """仅使用 PaddleOCR 原生可视化（save_to_img），生成 {输入名}_ocr_res_img.jpg。"""
    if not result:
        return None
    try:
        for res in result:
            if hasattr(res, "save_to_img"):
                res.save_to_img(output_dir)
                break
        stem = os.path.splitext(os.path.basename(img_path))[0]
        candidate = os.path.join(output_dir, stem + "_ocr_res_img.jpg")
        if os.path.exists(candidate):
            print(f"   使用 PaddleOCR 原生可视化: {candidate}")
            return candidate
    except Exception as e:
        print(f"   PaddleOCR 原生可视化失败: {e}")
    return None

def _draw_mc_visualization(
    img_path: str, filtered_data: List[Dict], save_path: Optional[str] = None
) -> Optional[str]:
    """
    兼容函数：已停用自绘可视化。
    结婚证流程改为仅使用 PaddleOCR 原生可视化 (*_ocr_res_img.jpg)。
    """
    _ = (img_path, filtered_data, save_path)
    return None


# ============================
# 第一页：左侧文字填充
# ============================

# 第一页左侧填充的文字配置
_PAGE1_TITLE_TEXT = "Certificate of Marriage"
_PAGE1_SEAL_TEXT = "(With the seal of the Ministry of Civil Affairs, P.R.C.)"
_PAGE1_TITLE_FONT_SIZE = 28
_PAGE1_SEAL_FONT_SIZE = 18
_PAGE1_LEFT_PADDING = 20   # 距左边距
_PAGE1_LINE_GAP = 12       # 两行文字之间的间距
_PAGE1_TOP_RATIO = 0.14    # 距顶部比例（靠上）


def _get_font_regular(size: int):
    """加载常规（非斜体）英文字体"""
    paths = [
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/Arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
    ]
    for path in paths:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                continue
    try:
        return ImageFont.truetype("arial.ttf", size)
    except Exception:
        return ImageFont.load_default()


def _get_font_italic(size: int):
    """加载斜体英文字体"""
    paths = [
        "C:/Windows/Fonts/ariali.ttf",
        "C:/Windows/Fonts/ARIALI.TTF",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Italic.ttf",
        "/System/Library/Fonts/Supplemental/Arial Italic.ttf",
    ]
    for path in paths:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                continue
    # 回退到常规字体
    return _get_font_regular(size)


def _text_size(draw_obj, text, font):
    """兼容不同 Pillow 版本的文本尺寸计算"""
    try:
        bbox = draw_obj.textbbox((0, 0), text, font=font)
        return bbox[2] - bbox[0], bbox[3] - bbox[1]
    except AttributeError:
        try:
            return draw_obj.textsize(text, font=font)
        except Exception:
            return len(text) * 10, 20


def _draw_page1_left_text(img_path: str) -> None:
    """
    在结婚证第一页图片左半屏、靠上位置填充：
      第1行: "Certificate of Marriage"          (黑色常规体)
      第2行: "(With the seal of the Ministry of Civil Affairs, P.R.C.)"  (红色斜体)

    直接覆盖原图文件。
    """
    try:
        img = Image.open(img_path).convert("RGBA")
    except Exception as e:
        print(f"   第一页文字填充失败: 无法打开图片 {img_path}: {e}")
        return

    w, h = img.size
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)

    # 可用区域：图片左半屏，避免文字跑到右半屏
    region_width = w // 2
    max_text_width = region_width - _PAGE1_LEFT_PADDING * 2

    # 加载字体
    font_title = _get_font_regular(_PAGE1_TITLE_FONT_SIZE)
    font_seal = _get_font_italic(_PAGE1_SEAL_FONT_SIZE)

    # 对标题做自动换行（限制在左半屏内）
    title_lines = _wrap_text_for_width(draw, _PAGE1_TITLE_TEXT, font_title,
                                       max_text_width)
    seal_lines = _wrap_text_for_width(draw, _PAGE1_SEAL_TEXT, font_seal,
                                      max_text_width)

    # 计算各行高度
    title_line_sizes = [_text_size(draw, line, font_title) for line in title_lines]
    seal_line_sizes = [_text_size(draw, line, font_seal) for line in seal_lines]

    title_total_h = sum(s[1] for s in title_line_sizes) + max(0, len(title_lines) - 1) * 4
    seal_total_h = sum(s[1] for s in seal_line_sizes) + max(0, len(seal_lines) - 1) * 4
    total_h = title_total_h + _PAGE1_LINE_GAP + seal_total_h

    # 靠上布局（不是垂直居中）
    y_start = int(h * _PAGE1_TOP_RATIO)
    y_start = max(20, y_start)

    current_y = y_start

    # 绘制标题行（黑色）
    title_color = (0, 0, 0, 255)
    for i, line in enumerate(title_lines):
        _, lh = title_line_sizes[i]
        # 左对齐到左半屏边距，确保不会向右偏
        x = _PAGE1_LEFT_PADDING
        try:
            draw.text((x, current_y), line, font=font_title, fill=title_color)
        except Exception:
            draw.text((x, current_y), line, fill=title_color)
        current_y += lh + 4

    current_y += _PAGE1_LINE_GAP

    # 绘制印章行（红色斜体）
    seal_color = (255, 0, 0, 255)
    for i, line in enumerate(seal_lines):
        _, lh = seal_line_sizes[i]
        # 左对齐到左半屏边距，确保不会向右偏
        x = _PAGE1_LEFT_PADDING
        try:
            draw.text((x, current_y), line, font=font_seal, fill=seal_color)
        except Exception:
            draw.text((x, current_y), line, fill=seal_color)
        current_y += lh + 4

    # 合成
    out = Image.alpha_composite(img, overlay)
    out_rgb = out.convert("RGB")
    out_rgb.save(img_path, "JPEG", quality=95)
    print(f"   第一页左侧文字已填充: {img_path}")


def _wrap_text_for_width(draw, text: str, font, max_width: int) -> List[str]:
    """将文本按单词自动换行到指定宽度"""
    words = text.split()
    lines: List[str] = []
    current_words: List[str] = []

    for word in words:
        test_line = ' '.join(current_words + [word])
        tw, _ = _text_size(draw, test_line, font)
        if tw <= max_width:
            current_words.append(word)
        else:
            if current_words:
                lines.append(' '.join(current_words))
                current_words = [word]
            else:
                # 单词本身就超宽，强制放入
                lines.append(word)

    if current_words:
        lines.append(' '.join(current_words))

    return lines if lines else [text]


# ============================
# 主处理函数
# ============================

def process_marriage_cert_image(
    input_path: str,
    output_dir: str = None,
    from_lang: str = 'zh',
    to_lang: str = 'en',
    enable_correction: bool = False,
    enable_visualization: bool = True,
    enable_merge: bool = True,
    enable_overlap_fix: bool = True,
    enable_colon_fix: bool = False,
    font_size: int = 18,
    confidence_threshold: float = 0.5,
    page_template: str = 'page2',
    registrar_signature_text: Optional[str] = None,
    registered_by_text: Optional[str] = None,
    registrar_signature_offset_x: int = 36,
    registrar_signature_offset_y: int = -12,
) -> Dict[str, Any]:
    """
    结婚证图片完整处理流程

    Args:
        input_path: 输入图片路径
        output_dir: 输出目录
        from_lang: 源语言（默认 'zh'）
        to_lang: 目标语言（默认 'en'）
        enable_correction: 是否启用透视矫正（已停用，仅兼容旧参数）
        enable_visualization: 是否生成可视化图片
        enable_merge: 框体合并（True=连续文本框合并翻译，False=每个框单独翻译）
        enable_overlap_fix: 重叠修正（True=自动检测并右移重叠框，False=保持原位）
        enable_colon_fix: 冒号修正（True=为字段名自动添加冒号，False=保持原样）
        font_size: 字体大小（像素），建议范围 8-30
        confidence_threshold: 置信度过滤阈值（第一页用0.8，其他用0.5）
        page_template: 模板页标识（'page1'/'page2'/'page3'）
        registrar_signature_text: 婚姻登记员手写签名手动输入文本（可选）
        registered_by_text: Registered by中的xxx手动输入文本（可选）
        registrar_signature_offset_x: 婚姻登记员手写签名右移偏移(px)
        registrar_signature_offset_y: 婚姻登记员手写签名纵向偏移(px)

    Returns:
        包含处理结果路径的字典，格式与 process_image() 一致
    """
    if output_dir is None:
        output_dir = settings.OUTPUT_DIR
    os.makedirs(output_dir, exist_ok=True)

    base_name = os.path.splitext(os.path.basename(input_path))[0]

    print("\n" + "=" * 80)
    print(" " * 15 + "结婚证 OCR 翻译与图像处理系统")
    print("=" * 80)
    print(f"输入图片: {input_path}")
    print(f"输出目录: {output_dir}")
    print(f"模板页: {page_template}")
    print(f"置信度阈值: {confidence_threshold}")
    if registrar_signature_text:
        print(f"手动签名输入: {registrar_signature_text}")
        print(f"手动签名右移偏移: {registrar_signature_offset_x}px")
        print(f"手动签名纵向偏移: {registrar_signature_offset_y}px")
    if registered_by_text:
        print(f"Registered by输入: {registered_by_text}")
    print(f"框体合并: {'开启' if enable_merge else '关闭'}")
    print(f"重叠修正: {'开启' if enable_overlap_fix else '关闭'}")
    print(f"冒号修正: {'开启' if enable_colon_fix else '关闭'}")
    print(f"字体大小: {font_size}px")
    print(f"使用LaMa: {LAMA_AVAILABLE and lama_model is not None}")
    print("=" * 80)

    # ---- 步骤0: 直接使用原图，避免预处理导致清晰度下降 ----
    # enable_correction 参数保留仅为兼容旧接口，不再执行矫正。
    img_path = input_path
    if enable_correction:
        print("\n提示: 透视矫正步骤已停用，为保证清晰度将直接使用原图。")

    # ---- 步骤1: OCR识别 ----
    print("\n步骤1: OCR识别...")
    try:
        result = ocr.predict(img_path)
    except (AttributeError, NotImplementedError):
        result = ocr.ocr(img_path)

    ocr_data = _extract_ocr_data(result)
    print(f"   识别到 {len(ocr_data['rec_texts'])} 个文本块")

    # ---- 步骤2: 置信度过滤 ----
    print(f"\n步骤2: 置信度过滤 (阈值={confidence_threshold})...")
    filtered_data = _filter_by_confidence(ocr_data, threshold=confidence_threshold)

    # 第1页：固定字段白名单过滤（去除星号、噪音字符等非目标文本）
    if (page_template or "").strip().lower() in ("page1", "first", "cover", "第一页"):
        print("\n步骤2.1: 第1页固定字段白名单过滤...")
        filtered_data = _filter_page1_fixed_fields(filtered_data)

    if not filtered_data:
        print("没有符合条件的文本，处理终止")
        return {
            "input_path": input_path,
            "processed_image": img_path,
            "raw_ocr_json": None,
            "translated_json": None,
            "visualization": None,
            "final_output": None,
            "items_count": 0
        }

    # 保存原始OCR结果
    out_json = os.path.join(output_dir, f"{base_name}_raw_ocr.json")
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(filtered_data, f, ensure_ascii=False, indent=4, default=str)
    print(f"   OCR结果已保存: {out_json}")

    # ---- 步骤3: 可视化 ----
    vis_path = None
    if enable_visualization:
        print("\n步骤3: 生成OCR可视化（PaddleOCR原生）...")
        vis_path = _try_paddle_vis_native(result, img_path, output_dir)
        if vis_path is None:
            print("   未生成 PaddleOCR 原生可视化图片（期望文件: *_ocr_res_img.jpg）")

    # ---- 步骤4: 文本合并 ----
    if enable_merge:
        print("\n步骤4: 文本合并...")
        merged_data = _merge_consecutive_texts(filtered_data)
    else:
        merged_data = [{
            'text': item['text'],
            'box': item['box'],
            'merged_indices': [item['index']],
            'sub_boxes': [item['box']],
            'is_merged': False
        } for item in filtered_data]
        print("\n步骤4: 框体合并已关闭，每个文本框单独处理")

    # ---- 步骤5: 智能翻译 ----
    print("\n步骤5: 智能翻译...")
    merged_data = _translate_all(merged_data, enable_colon_fix=enable_colon_fix)

    # 第1页特殊长文本处理：拉伸框高并强制换行
    if (page_template or "").strip().lower() in ("page1", "first", "cover", "第一页"):
        _apply_page1_special_adjustments(merged_data)

    # 保存翻译结果
    trans_json = os.path.join(output_dir, f"{base_name}_translated.json")
    with open(trans_json, "w", encoding="utf-8") as f:
        json.dump(merged_data, f, ensure_ascii=False, indent=4, default=str)
    print(f"   翻译结果已保存: {trans_json}")

    # ---- 步骤6: 涂抹文本区域 ----
    print("\n步骤6: 涂抹文本区域...")
    image = cv2.imread(img_path)
    inpainted = _inpaint_regions(image, filtered_data, use_lama=True)

    # ---- 步骤7: 回填翻译文本 ----
    print("\n步骤7: 回填翻译文本...")
    manual_registrar_signature_en = None
    manual_registered_by_en = None
    if registrar_signature_text and registrar_signature_text.strip():
        manual_registrar_signature_en = _translate_manual_signature_text(registrar_signature_text.strip())
        print(f"   手动签名翻译结果: {manual_registrar_signature_en}")
    if registered_by_text and registered_by_text.strip():
        manual_registered_by_en = _translate_registered_by_text(registered_by_text.strip())
        print(f"   Registered by翻译结果: {manual_registered_by_en}")

    final_image = _draw_translated_texts(
        inpainted, merged_data,
        font_path=None,
        font_size=font_size,
        enable_overlap_fix=enable_overlap_fix,
        manual_registrar_signature_en=manual_registrar_signature_en,
        manual_registered_by_en=manual_registered_by_en,
        registrar_signature_offset_x=registrar_signature_offset_x,
        registrar_signature_offset_y=registrar_signature_offset_y
    )

    # ---- 步骤8: 输出缩放 ----
    output_width = settings.TARGET_IMAGE_WIDTH
    h, w = final_image.shape[:2]
    if w != output_width:
        scale = output_width / w
        new_h = int(round(h * scale))
        final_image = cv2.resize(final_image, (output_width, new_h), interpolation=cv2.INTER_AREA)

    # ---- 步骤9: 保存结果 ----
    output_img = os.path.join(output_dir, f"{base_name}_translated.jpg")
    cv2.imwrite(output_img, final_image)
    print(f"\n   翻译图片已保存: {output_img}")

    # ---- 步骤10: 添加水印 ----
    add_watermark(output_img)

    # ---- 步骤11: 第一/封面专属 — 左侧填充 Certificate of Marriage 文字 ----
    # 兼容触发条件：
    # 1) page_template 明确为 page1
    # 2) 或当前阈值>=0.8（第一页模板特征），避免前端模板参数未透传时丢失填充
    template_key = (page_template or "").strip().lower()
    if template_key in ("page1", "first", "cover", "第一页") or confidence_threshold >= 0.8:
        print("\n步骤11: 第一页左侧文字填充...")
        _draw_page1_left_text(output_img)

    print("\n" + "=" * 80)
    print("   结婚证处理全部完成！")
    print("=" * 80)

    return {
        "input_path": input_path,
        "processed_image": img_path,
        "raw_ocr_json": out_json,
        "translated_json": trans_json,
        "visualization": vis_path,
        "final_output": output_img,
        "items_count": len(filtered_data)
    }
