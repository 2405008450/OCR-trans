"""
图片处理服务模块
包含OCR、翻译、图像修复等功能
"""
import os
# 必须在 import paddle 之前设置，避免 OneDNN/PIR 兼容性导致的 NotImplementedError
os.environ["FLAGS_use_mkldnn"] = "0"
os.environ["FLAGS_use_new_executor"] = "0"

import json
import numpy as np
from paddleocr import PaddleOCR
import cv2
from PIL import Image, ImageDraw, ImageFont
import re
import time
from openai import OpenAI
from typing import List, Dict, Any, Optional

from app.core.config import settings

try:
    from simple_lama_inpainting import SimpleLama
    LAMA_AVAILABLE = True
except ImportError:
    LAMA_AVAILABLE = False
    print("⚠️ simple_lama_inpainting 未安装，将使用OpenCV作为备选方案")

# -----------------------------
# 输入转图片列表
# -----------------------------

def convert_input_to_images(input_path: str, output_dir: str = None) -> List[str]:
    """
    📂 输入处理器：仅支持图片，直接返回路径列表；其他格式返回空列表。
    """
    file_ext = os.path.splitext(input_path)[1].lower()
    if file_ext in ['.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff']:
        return [input_path]
    print("❌ 不支持的文件格式，仅支持图片: .jpg, .jpeg, .png, .bmp, .tif, .tiff")
    return []


# -----------------------------
# 自动矫正
# -----------------------------

def order_points(pts):
    """对四个角点进行排序：左上，右上，右下，左下"""
    rect = np.zeros((4, 2), dtype="float32")
    s = pts.sum(axis=1)
    rect[0] = pts[np.argmin(s)]  # 左上
    rect[2] = pts[np.argmax(s)]  # 右下
    diff = np.diff(pts, axis=1)
    rect[1] = pts[np.argmin(diff)]  # 右上
    rect[3] = pts[np.argmax(diff)]  # 左下
    return rect


def auto_correct_perspective(img_path: str) -> str:
    """
    📐 自动透视矫正：把歪斜的身份证拉正
    """
    img = cv2.imread(img_path)
    if img is None:
        return img_path

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    edged = cv2.Canny(blurred, 75, 200)

    cnts, _ = cv2.findContours(edged.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not cnts:
        print("⚠️ 未检测到轮廓，跳过矫正")
        return img_path

    cnts = sorted(cnts, key=cv2.contourArea, reverse=True)[:5]

    screenCnt = None
    for c in cnts:
        peri = cv2.arcLength(c, True)
        approx = cv2.approxPolyDP(c, 0.02 * peri, True)

        if len(approx) == 4 and cv2.contourArea(c) > 50000:
            screenCnt = approx
            break

    if screenCnt is None:
        print("⚠️ 未找到矩形卡证区域，跳过矫正")
        return img_path

    print("📐 检测到倾斜卡证，正在进行透视变换矫正...")

    pts = screenCnt.reshape(4, 2)
    rect = order_points(pts)
    (tl, tr, br, bl) = rect

    widthA = np.sqrt(((br[0] - bl[0]) ** 2) + ((br[1] - bl[1]) ** 2))
    widthB = np.sqrt(((tr[0] - tl[0]) ** 2) + ((tr[1] - tl[1]) ** 2))
    maxWidth = max(int(widthA), int(widthB))

    heightA = np.sqrt(((tr[0] - br[0]) ** 2) + ((tr[1] - br[1]) ** 2))
    heightB = np.sqrt(((tl[0] - bl[0]) ** 2) + ((tl[1] - bl[1]) ** 2))
    maxHeight = max(int(heightA), int(heightB))

    dst = np.array([
        [0, 0],
        [maxWidth - 1, 0],
        [maxWidth - 1, maxHeight - 1],
        [0, maxHeight - 1]], dtype="float32")

    M = cv2.getPerspectiveTransform(rect, dst)
    warped = cv2.warpPerspective(img, M, (maxWidth, maxHeight))

    name, ext = os.path.splitext(img_path)
    new_path = f"{name}_corrected{ext}"
    cv2.imwrite(new_path, warped)

    print(f"✅ 矫正完成: {new_path}")
    return new_path


def preprocess_resize_image(input_path: str, target_width: int = None) -> str:
    """
    🖼️ 图像预处理：将图片等比例缩放至指定宽度
    """
    if target_width is None:
        target_width = settings.TARGET_IMAGE_WIDTH

    if not os.path.exists(input_path):
        print(f"⚠️ 文件不存在: {input_path}")
        return input_path

    img = cv2.imread(input_path)
    if img is None:
        print(f"⚠️ 无法读取图片: {input_path}")
        return input_path

    h, w = img.shape[:2]

    if abs(w - target_width) < 5:
        print(f"📏 图片宽度 ({w}px) 已符合标准，跳过缩放。")
        return input_path

    scale = target_width / w
    new_h = int(h * scale)

    print(f"🔄正在预处理缩放: {w}x{h} → {target_width}x{new_h} (缩放比: {scale:.2f})")

    interpolation = cv2.INTER_AREA if scale < 1 else cv2.INTER_LINEAR
    resized_img = cv2.resize(img, (target_width, new_h), interpolation=interpolation)

    dir_name = os.path.dirname(input_path)
    file_name = os.path.basename(input_path)
    name, ext = os.path.splitext(file_name)

    new_filename = f"{name}_1080p{ext}"
    new_path = os.path.join(dir_name, new_filename)

    cv2.imwrite(new_path, resized_img)
    print(f"✅ 预处理完成，使用新图片: {new_path}")

    return new_path


# -----------------------------
# DeepSeek 配置
# -----------------------------
deepseek_client = OpenAI(
    api_key=settings.DEEPSEEK_API_KEY,
    base_url=settings.DEEPSEEK_BASE_URL
)

# -----------------------------
# 初始化 OCR
# -----------------------------
ocr = PaddleOCR(
    use_doc_orientation_classify=False,
    use_doc_unwarping=False,
    use_textline_orientation=True,
    lang="ch"
)

# 初始化 LaMa 模型（如果可用）
lama_model = None
if LAMA_AVAILABLE:
    try:
        print("🤖 正在加载 AI 消除模型 (LaMa)...")
        lama_model = SimpleLama()
    except Exception as e:
        print(f"⚠️ LaMa 模型加载失败: {e}")


# -----------------------------
# 通用提取函数
# -----------------------------
def extract_all_items(result):
    items = []

    for elem in result:
        if isinstance(elem, dict):
            res = elem.get("res", elem)
            texts = res.get("rec_texts") or res.get("texts") or []
            scores = res.get("rec_scores") or []
            polys = res.get("rec_polys") or res.get("rec_boxes") or []

            for t, s, p in zip(texts, scores, polys):
                if isinstance(p, np.ndarray):
                    p = p.tolist()
                items.append({
                    "text": t,
                    "score": float(s) if s is not None else None,
                    "box": p
                })

        elif hasattr(elem, "texts") and hasattr(elem, "boxes"):
            texts = elem.texts
            boxes = elem.boxes
            scores = getattr(elem, "scores", [None] * len(texts))

            for t, s, b in zip(texts, scores, boxes):
                try:
                    box_list = b.tolist()
                except:
                    box_list = b
                items.append({
                    "text": t,
                    "score": float(s) if s is not None else None,
                    "box": box_list
                })

        elif isinstance(elem, (list, tuple)):
            for row in elem:
                if isinstance(row, (list, tuple)) and len(row) >= 2:
                    box = row[0]
                    pair = row[1]
                    if isinstance(pair, (list, tuple)) and len(pair) >= 2:
                        text = pair[0]
                        score = float(pair[1])
                    else:
                        text = pair
                        score = None

                    try:
                        box_list = [list(map(int, p)) for p in box]
                    except:
                        box_list = box

                    items.append({
                        "text": text,
                        "score": score,
                        "box": box_list
                    })

    return items


# -----------------------------
# 智能分割
# -----------------------------
def split_merged_items(items):
    result = []
    split_keywords = ["姓名", "性别", "民族", "公民身份号码"]
    protect_keywords = ["出生", "住址", "地址"]

    for item in items:
        text = item["text"]
        box = item["box"]
        score = item["score"]

        is_protected = any(kw in text for kw in protect_keywords)
        has_date = re.search(r'\d{4}.*?年.*?\d+.*?月.*?\d+.*?日', text)
        has_id = re.search(r'\d{15,18}', text)
        has_address = any(kw in text for kw in ['省', '市', '区', '县', '街', '路', '号', '室', '房'])

        if is_protected or has_date or has_id or has_address:
            result.append(item)
            continue

        needs_split = False

        if " " in text or "　" in text:
            found_keywords = sum(1 for kw in split_keywords if kw in text)
            if found_keywords >= 2:
                needs_split = True

        if needs_split:
            parts = split_text_by_rules(text, split_keywords)

            if len(parts) > 1:
                sub_items = calculate_sub_boxes(parts, box, score)
                result.extend(sub_items)
                print(f"🔧 分割: '{text}' → {parts}")
                continue

        result.append(item)

    return result


def split_text_by_rules(text, keywords):
    parts = []

    if " " in text or "　" in text:
        parts = re.split(r'[\s　]+', text)
        parts = [p.strip() for p in parts if p.strip()]

        if len(parts) > 1:
            return parts

    temp_parts = []
    remaining = text

    while remaining:
        matched = False

        for kw in sorted(keywords, key=len, reverse=True):
            if remaining.startswith(kw):
                temp_parts.append(kw)
                remaining = remaining[len(kw):]
                matched = True
                break

        if not matched:
            if temp_parts:
                next_kw_pos = len(remaining)
                for kw in keywords:
                    pos = remaining.find(kw)
                    if pos > 0:
                        next_kw_pos = min(next_kw_pos, pos)

                value = remaining[:next_kw_pos]
                if value:
                    temp_parts.append(value)
                    remaining = remaining[next_kw_pos:]
            else:
                first_kw_pos = len(remaining)
                for kw in keywords:
                    pos = remaining.find(kw)
                    if pos >= 0:
                        first_kw_pos = min(first_kw_pos, pos)

                if first_kw_pos < len(remaining):
                    temp_parts.append(remaining[:first_kw_pos])
                    remaining = remaining[first_kw_pos:]
                else:
                    temp_parts.append(remaining)
                    remaining = ""

    if len(temp_parts) > 1:
        return temp_parts

    return [text]


def calculate_sub_boxes(parts, original_box, score):
    sub_items = []

    x1, y1 = original_box[0]
    x2, y2 = original_box[1]
    x3, y3 = original_box[2] if len(original_box) > 2 else original_box[1]
    x4, y4 = original_box[3] if len(original_box) > 3 else original_box[0]

    total_width = x2 - x1
    total_chars = sum(len(p) for p in parts) + (len(parts) - 1)

    if total_chars == 0:
        return [{"text": "".join(parts), "score": score, "box": original_box}]

    char_width = total_width / total_chars
    current_x = x1

    for i, part in enumerate(parts):
        if not part:
            continue

        part_width = len(part) * char_width

        new_box = [
            [int(current_x), int(y1)],
            [int(current_x + part_width), int(y2)],
            [int(current_x + part_width), int(y3)],
            [int(current_x), int(y4)]
        ]

        sub_items.append({
            "text": part,
            "score": score,
            "box": new_box
        })

        current_x += part_width + char_width

    return sub_items


# -----------------------------
# 智能合并地址行
# -----------------------------
def get_global_metrics(items):
    """计算全图的标准行高和标准字符宽度"""
    heights = []
    char_widths = []

    for item in items:
        box = item['box']
        text = item['text']
        if not text:
            continue

        h = abs(box[2][1] - box[0][1])
        heights.append(h)

        w = abs(box[1][0] - box[0][0])
        if len(text) > 0:
            char_widths.append(w / len(text))

    std_h = np.median(heights) if heights else 20
    std_w = np.median(char_widths) if char_widths else 20

    return std_h, std_w


def merge_address_lines(items):
    """动态阈值版：智能合并地址行"""
    std_h, std_w = get_global_metrics(items)
    print(f"📏 图片排版基准: 标准行高={std_h:.1f}px, 标准字宽={std_w:.1f}px")

    merged_items = []
    i = 0
    n = len(items)

    while i < n:
        current_item = items[i]
        current_text = current_item['text']

        if '住址' in current_text or '地址' in current_text:
            lines_to_merge = [current_item]
            j = i + 1
            merge_count = 0

            current_box = current_item['box']
            current_y_top = current_box[0][1]
            current_y_bottom = current_box[2][1]
            current_x_left = current_box[0][0]

            while j < n and merge_count < 3:
                next_item = items[j]
                next_text = next_item['text']
                next_box = next_item['box']

                next_y_top = next_box[0][1]
                next_x_left = next_box[0][0]

                gap = next_y_top - current_y_bottom

                if gap > 1.5 * std_h:
                    print(f"   ❌ 垂直间距过大 ({gap:.1f}px > 1.5x标准高): 停止合并 '{next_text}'")
                    break

                indent = abs(next_x_left - current_x_left)

                if indent > 4.0 * std_w:
                    print(f"   ❌ 水平缩进过大 ({indent:.1f}px > 4x字宽): 停止合并 '{next_text}'")
                    break

                if len(next_text) > 40:
                    print(f"   ❌ 文本过长: 不合并 '{next_text}'")
                    break

                if re.search(r'\d{14,18}', next_text):
                    print(f"   ❌ 发现身份证号: 停止合并 '{next_text}'")
                    break

                if any(kw in next_text for kw in ['公民身份', '身份证', '签发机关', '有效期限']):
                    print(f"   ❌ 发现新字段: 停止合并 '{next_text}'")
                    break

                lines_to_merge.append(next_item)
                current_y_bottom = next_box[2][1]
                merge_count += 1
                j += 1

            if len(lines_to_merge) > 1:
                merged_text = ''.join([item['text'] for item in lines_to_merge])
                merged_box = merge_boxes_multi(lines_to_merge)
                valid_scores = [item['score'] for item in lines_to_merge if item.get('score')]
                merged_score = sum(valid_scores) / len(valid_scores) if valid_scores else 0.0

                merged_items.append({
                    'text': merged_text,
                    'score': merged_score,
                    'box': merged_box
                })

                debug_msg = ' + '.join([f"'{item['text']}'" for item in lines_to_merge])
                print(f"🔗 成功合并地址: {debug_msg} → '{merged_text}'")

                i = j
            else:
                merged_items.append(current_item)
                i += 1
        else:
            merged_items.append(current_item)
            i += 1

    return merged_items


def merge_boxes_multi(items):
    """合并多个box，返回外接矩形"""
    all_points = []
    for item in items:
        all_points.extend(item['box'])

    xs = [p[0] for p in all_points]
    ys = [p[1] for p in all_points]

    x_min, x_max = min(xs), max(xs)
    y_min, y_max = min(ys), max(ys)

    return [
        [x_min, y_min],
        [x_max, y_min],
        [x_max, y_max],
        [x_min, y_max]
    ]


# -----------------------------
# DeepSeek 翻译
# -----------------------------
FIELD_TRANSLATIONS = {
    '姓名': 'Name:',
    '性别': 'Gender:',
    '民族': 'Ethnicity:',
    '出生': 'Date of Birth:',
    '住址': 'Address:',
    '公民身份号码': 'ID Number:'
}

SINGLE_VALUE_MAP = {
    '男': 'Male',
    '女': 'Female',
    '汉': 'Han'
}

FIELD_KEYS = list(FIELD_TRANSLATIONS.keys())

# -----------------------------
# 身份证背面（国徽面）固定翻译映射
# -----------------------------
BACKSIDE_FIXED_MAP = {
    '中华人民共和国': "People's Republic of China",
    '居民身份证': 'Resident Identity Card',
    '签发机关': 'Issuing Authority',
    '有效期限': 'Validity Period',
}

# 背面「签发机关」box 右边界扩展像素，使 "Issuing Authority" 能单行放下（OCR 原 box 太窄）
BACKSIDE_ISSUING_AUTHORITY_BOX_EXTRA_WIDTH = 1000


def _expand_box_right(box, extra_width):
    """将四边形 box 的右边界向右扩展 extra_width 像素。box: [[x1,y1],[x2,y2],[x3,y3],[x4,y4]]"""
    if not box or len(box) < 4:
        return box
    # 深拷贝，避免改到原始数据
    new_box = [[float(p[0]), float(p[1])] for p in box]
    # 右上 [1]、右下 [2] 的 x 增加
    new_box[1][0] += extra_width
    new_box[2][0] += extra_width
    return new_box


def translate_all_items_backside(items, from_lang='zh', to_lang='en'):
    """
    身份证背面专用翻译：固定标签用硬编码映射，变量值（机关名、日期）走 DeepSeek。
    """
    translated_items = []

    for i, item in enumerate(items):
        text = item.get('text', '').strip()
        print(f"[{i + 1}/{len(items)}] ", end="")

        # 1. 固定映射命中 → 直接使用
        if text in BACKSIDE_FIXED_MAP:
            mapped = BACKSIDE_FIXED_MAP[text]
            print(f"✅ 背面固定映射: {text} → {mapped}")
            box = item.get("box")
            # 「签发机关」英文 "Issuing Authority" 比中文宽，直接拉宽 box 避免换行
            if text == '签发机关' and box and len(box) >= 4:
                box = _expand_box_right(box, BACKSIDE_ISSUING_AUTHORITY_BOX_EXTRA_WIDTH)
                print(f"   → 签发机关 box 右扩 {BACKSIDE_ISSUING_AUTHORITY_BOX_EXTRA_WIDTH}px")
            translated_items.append({
                "text": mapped,
                "original_text": text,
                "score": item.get("score"),
                "box": box,
                "fixed": True  # 标记为固定映射
            })
            continue

        # 2. 日期格式（纯数字+点+横线）→ 保持原文
        if re.match(r'^[\d.\-/]+$', text):
            print(f"✅ 日期保持原文: {text}")
            translated_items.append({
                "text": text,
                "original_text": text,
                "score": item.get("score"),
                "box": item.get("box")
            })
            continue

        # 3. 其余（如签发机关的具体名称）→ 调 DeepSeek 翻译
        try:
            translated = translate_basic(text, from_lang, to_lang)
            print(f"✅ 翻译: {text} → {translated}")
        except Exception as e:
            print(f"❌ 翻译出错: {e}")
            translated = text

        translated_items.append({
            "text": translated,
            "original_text": text,
            "score": item.get("score"),
            "box": item.get("box")
        })
        time.sleep(0.3)

    return translated_items


def preprocess_field_item_inline(text):
    """如果文本以字段开头，拆分并返回 (field_key, value)"""
    for key in sorted(FIELD_KEYS, key=len, reverse=True):
        if text.startswith(key):
            value = text[len(key):].strip()
            return key, value
    return None, text


def translate_basic(text, from_lang='zh', to_lang='en', max_tokens=100):
    """仅翻译值文本（value），不改变数字，保留简单map"""
    if not text or not text.strip():
        return text

    if re.match(r'^[\d\s\-X]+$', text):
        return text

    if text in SINGLE_VALUE_MAP:
        return SINGLE_VALUE_MAP[text]

    try:
        lang_map = {
            'zh': 'Chinese',
            'en': 'English',
            'ja': 'Japanese',
            'ko': 'Korean'
        }
        source_lang = lang_map.get(from_lang, 'Chinese')
        target_lang = lang_map.get(to_lang, 'English')

        response = deepseek_client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {
                    "role": "system",
                    "content": f"You are a professional translator. Translate from {source_lang} to {target_lang}. "
                               f"Rules:\n"
                               f"1. Keep ALL numbers unchanged\n"
                               f"2. Only return the translated text without explanation\n"
                               f"3. For names, use proper capitalization\n"
                               f"4. Be concise and accurate"
                },
                {
                    "role": "user",
                    "content": f"Translate: {text}"
                }
            ],
            temperature=0.2,
            max_tokens=max_tokens
        )

        translated = response.choices[0].message.content.strip().strip('"\'')
        return translated
    except Exception as e:
        print(f"❌ translate_basic 出错: {e}")
        return text


def translate_text_deepseek(text, from_lang='zh', to_lang='en'):
    """对单个 OCR 文本的翻译"""
    if not text or not text.strip():
        return text

    if re.match(r'^[\d\s\-X]+$', text):
        return text

    sorted_keys = sorted(FIELD_KEYS, key=len, reverse=True)
    pattern = f"({'|'.join(sorted_keys)})"
    parts = re.split(pattern, text)

    valid_parts = [p for p in parts if p.strip()]
    contained_keys = [p for p in valid_parts if p in FIELD_KEYS]

    if len(contained_keys) >= 1:
        translated_parts = []

        for idx, part in enumerate(parts):
            part = part.strip()
            if not part:
                continue

            if part in FIELD_KEYS:
                field_en = FIELD_TRANSLATIONS.get(part, part + ":")
                translated_parts.append(field_en)
            else:
                val_en = translate_basic(part, from_lang, to_lang)
                translated_parts.append(val_en)

        result = " ".join(translated_parts)

        if len(translated_parts) > 1:
            print(f"✅ 多字段内联翻译: {text} → {result}")
            return result

    if text in SINGLE_VALUE_MAP:
        return SINGLE_VALUE_MAP[text]

    try:
        translated = translate_basic(text, from_lang, to_lang)
        print(f"✅ 翻译: {text} → {translated}")
        return translated
    except Exception as e:
        print(f"❌ 翻译出错: {e}")
        return text


def translate_all_items(items, from_lang='zh', to_lang='en'):
    """遍历 items，处理字段+值合并和翻译"""
    translated_items = []
    i = 0
    n = len(items)

    while i < n:
        item = items[i]
        text = item.get('text', '').strip()
        print(f"[{i + 1}/{n}] ", end="")

        field_key, inline_value = preprocess_field_item_inline(text)
        if field_key and inline_value:
            translated_text = translate_text_deepseek(text, from_lang, to_lang)
            translated_items.append({
                "text": translated_text,
                "original_text": text,
                "score": item.get("score"),
                "box": item.get("box")
            })
            i += 1
            time.sleep(0.3)
            continue

        if text in FIELD_KEYS:
            next_idx = i + 1
            if next_idx < n:
                next_text = items[next_idx].get('text', '').strip()
                next_field_key, _ = preprocess_field_item_inline(next_text)
                if next_text and (next_text not in FIELD_KEYS) and not next_field_key:
                    value_en = translate_basic(next_text, from_lang, to_lang)
                    field_en = FIELD_TRANSLATIONS.get(text, text + ":")
                    combined = f"{field_en} {value_en}"
                    print(f"✅ 合并字段+下一行值: '{text}' + '{next_text}' → {combined}")

                    merged_box = merge_boxes_multi([item, items[next_idx]])
                    avg_score = None
                    try:
                        scores = [s for s in [item.get('score'), items[next_idx].get('score')] if s]
                        avg_score = sum(scores) / len(scores) if scores else None
                    except:
                        avg_score = item.get('score')

                    translated_items.append({
                        "text": combined,
                        "original_text": text + " " + next_text,
                        "score": avg_score,
                        "box": merged_box
                    })
                    i += 2
                    time.sleep(0.3)
                    continue

        translated_text = translate_text_deepseek(text, from_lang, to_lang)
        translated_items.append({
            "text": translated_text,
            "original_text": text,
            "score": item.get("score"),
            "box": item.get("box")
        })

        i += 1
        time.sleep(0.3)

    return translated_items


# -----------------------------
# 智能涂抹
# -----------------------------
def smart_inpaint(img, items):
    """使用 AI (LaMa) 智能消除文字，保留背景底纹"""
    print("\n🎨 正在生成遮罩并进行 AI 修复...")

    mask = np.zeros(img.shape[:2], dtype=np.uint8)

    for item in items:
        box = item.get("box")
        if not box or len(box) < 4:
            continue

        pts = np.array(box, np.int32)
        center = pts.mean(axis=0)
        expanded_pts = center + (pts - center) * 1.1
        expanded_pts = expanded_pts.astype(np.int32)

        cv2.fillPoly(mask, [expanded_pts.reshape((-1, 1, 2))], color=255)

    img_pil = Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
    mask_pil = Image.fromarray(mask)

    if LAMA_AVAILABLE and lama_model is not None:
        try:
            result_pil = lama_model(img_pil, mask_pil)
            inpainted_img = cv2.cvtColor(np.array(result_pil), cv2.COLOR_RGB2BGR)
            print("   ✅ AI 修复完成，底纹已重建")
            return inpainted_img
        except Exception as e:
            print(f"   ❌ AI 修复失败，回退到 OpenCV 算法: {e}")
            return cv2.inpaint(img, mask, 3, cv2.INPAINT_NS)
    else:
        return cv2.inpaint(img, mask, 3, cv2.INPAINT_NS)


# -----------------------------
# 自动换行
# -----------------------------
def should_wrap_text(item, text, bbox_width, draw, font):
    """判断是否应该自动换行"""
    field = item.get("field", "").lower()

    if "address" in field:
        return True

    if "issued" in field:
        return True

    words = text.split(" ")
    if len(words) <= 3:
        bbox = draw.textbbox((0, 0), text, font=font)
        width = bbox[2] - bbox[0]

        if width <= bbox_width * 1.2:
            return False

    bbox = draw.textbbox((0, 0), text, font=font)
    width = bbox[2] - bbox[0]

    return width > bbox_width


def wrap_text(text, font, max_width, draw):
    """按单词自动换行（英文最佳效果）"""
    words = text.split(" ")
    lines = []
    current = ""

    for word in words:
        test_line = (current + " " + word).strip()

        bbox = draw.textbbox((0, 0), test_line, font=font)
        w = bbox[2] - bbox[0]

        if w <= max_width:
            current = test_line
        else:
            if current:
                lines.append(current)
            current = word

    if current:
        lines.append(current)

    return lines


# -----------------------------
# 图片擦除 + 多行填充
# -----------------------------
def inpaint_and_fill(img_path: str, items: List[Dict], output_path: str = None, auto_wrap: bool = True) -> str:
    """
    智能擦除原文字区域，填充翻译后的文本（支持自动换行 + 底部对齐）
    
    Args:
        auto_wrap: 是否启用自动换行。正面(front)为True，背面(back)为False。
    """
    img = cv2.imread(img_path)
    if img is None:
        print("⚠️ 无法读取图片")
        return None

    img = smart_inpaint(img, items)

    print("✏️ 填充翻译文本（自动换行 + 底部对齐）...")

    img_pil = Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
    draw = ImageDraw.Draw(img_pil)

    # 尝试加载字体（兼容Windows和Linux）
    font_en_path = None
    font_zh_path = None
    
    # Windows字体路径
    windows_fonts = {
        "en": "C:/Windows/Fonts/arial.ttf",
        "zh": "C:/Windows/Fonts/simhei.ttf"
    }
    
    # Linux字体路径
    linux_fonts = {
        "en": [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
            "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf",
        ],
        "zh": [
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
            "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
            "/usr/share/fonts/truetype/arphic/uming.ttc",
        ]
    }
    
    # macOS字体路径
    mac_fonts = {
        "en": "/System/Library/Fonts/Helvetica.ttc",
        "zh": "/System/Library/Fonts/PingFang.ttc"
    }
    
    # 尝试加载英文字体
    if os.path.exists(windows_fonts["en"]):
        font_en_path = windows_fonts["en"]
    else:
        for path in linux_fonts["en"]:
            if os.path.exists(path):
                font_en_path = path
                break
        if not font_en_path and os.path.exists(mac_fonts["en"]):
            font_en_path = mac_fonts["en"]
    
    # 尝试加载中文字体
    if os.path.exists(windows_fonts["zh"]):
        font_zh_path = windows_fonts["zh"]
    else:
        for path in linux_fonts["zh"]:
            if os.path.exists(path):
                font_zh_path = path
                break
        if not font_zh_path and os.path.exists(mac_fonts["zh"]):
            font_zh_path = mac_fonts["zh"]

    for item in items:
        box = item.get("box")
        text = item.get("text", "")
        original_text = item.get("original_text", "")

        if not box or len(box) < 4 or not text:
            continue

        x1, y1 = box[0]
        x2, y2 = box[1]
        x3, y3 = box[2]
        x4, y4 = box[3]

        box_width = x2 - x1
        box_height = y3 - y1

        is_english = all(ord(c) < 128 or c.isspace() for c in text)
        font_path = font_en_path if is_english else font_zh_path

        font_size = 24
        min_font_size = 12
        best_font = None

        while font_size >= min_font_size:
            try:
                if font_path and os.path.exists(font_path):
                    test_font = ImageFont.truetype(font_path, font_size)
                else:
                    test_font = ImageFont.load_default()
            except Exception as e:
                test_font = ImageFont.load_default()

            # 兼容不同版本的Pillow
            try:
                bbox = draw.textbbox((0, 0), text, font=test_font)
                text_h = bbox[3] - bbox[1]
            except AttributeError:
                # 旧版本使用textsize
                try:
                    bbox = draw.textsize(text, font=test_font)
                    text_h = bbox[1]
                except:
                    text_h = font_size * 1.2  # 估算值

            if text_h <= box_height * 1.1:
                best_font = test_font
                break

            font_size -= 1

        if best_font is None:
            try:
                if font_path and os.path.exists(font_path):
                    best_font = ImageFont.truetype(font_path, min_font_size)
                else:
                    best_font = ImageFont.load_default()
            except:
                best_font = ImageFont.load_default()

        if is_english and auto_wrap:
            if should_wrap_text(item, text, box_width, draw, best_font):
                lines = wrap_text(text, best_font, box_width, draw)
            else:
                lines = [text]
        else:
            lines = [text]

        # 兼容不同版本的getbbox方法
        try:
            line_height = best_font.getbbox("A")[3] - best_font.getbbox("A")[1] + 6
        except AttributeError:
            try:
                # 旧版本使用getsize
                line_height = best_font.getsize("A")[1] + 6
            except:
                line_height = font_size + 6  # 估算值

        total_text_height = len(lines) * line_height

        y_start = y3 - total_text_height
        if y_start < y1:
            y_start = y1

        for i, line in enumerate(lines):
            y_line = y_start + i * line_height
            try:
                draw.text((int(x1), int(y_line)), line, font=best_font, fill=(0, 0, 0))
            except AttributeError:
                # 如果getmask2方法不存在，尝试不使用font参数
                try:
                    draw.text((int(x1), int(y_line)), line, fill=(0, 0, 0))
                except Exception as e:
                    print(f"⚠️ 绘制文本失败: {line}, 错误: {e}")
                    continue

    img_out = cv2.cvtColor(np.array(img_pil), cv2.COLOR_RGB2BGR)

    if output_path is None:
        output_path = os.path.splitext(img_path)[0] + ".translated.jpg"

    cv2.imwrite(output_path, img_out)
    print(f"✅ 翻译后的图片已保存到: {output_path}")

    return output_path


# -----------------------------
# 可视化函数
# -----------------------------
def draw_visualization(img_path: str, items: List[Dict], save_path: str = None) -> str:
    """生成OCR结果可视化图片"""
    img = cv2.imread(img_path)
    if img is None:
        print("⚠️ 无法读取图片，跳过可视化。")
        return None

    for it in items:
        box = it.get("box")
        if not box or len(box) < 4:
            continue
        pts = np.array(box, np.int32)
        cv2.polylines(img, [pts], isClosed=True, color=(0, 255, 0), thickness=2)

    img_pil = Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
    draw = ImageDraw.Draw(img_pil)

    # 尝试加载字体（兼容Windows和Linux）
    font = None
    font_paths = [
        "C:/Windows/Fonts/simhei.ttf",  # Windows
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",  # Linux常见路径
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",  # Linux
        "/System/Library/Fonts/Helvetica.ttc",  # macOS
    ]
    
    for path in font_paths:
        try:
            if os.path.exists(path):
                font = ImageFont.truetype(path, 18)
                break
        except:
            continue
    
    # 如果所有字体都加载失败，使用默认字体
    if font is None:
        try:
            font = ImageFont.load_default()
        except:
            # 如果默认字体也失败，创建一个简单的字体
            font = ImageFont.load_default()
    
    for it in items:
        box = it.get("box")
        text = it.get("text", "")
        if not box or len(box) < 4:
            continue
        x, y = box[0]
        try:
            # 使用兼容的文本绘制方法
            draw.text((x, y - 20), text, font=font, fill=(255, 0, 0))
        except AttributeError:
            # 如果getmask2方法不存在，使用备用方法
            try:
                # 尝试使用getmask方法（旧版本兼容）
                draw.text((x, y - 20), text, fill=(255, 0, 0))
            except Exception as e:
                print(f"⚠️ 绘制文本失败: {text}, 错误: {e}")
                continue

    img_out = cv2.cvtColor(np.array(img_pil), cv2.COLOR_RGB2BGR)
    if save_path is None:
        save_path = os.path.splitext(img_path)[0] + ".vis.jpg"
    cv2.imwrite(save_path, img_out)
    print(f"📊 可视化图片已保存到: {save_path}")
    return save_path


# -----------------------------
# 主处理函数
# -----------------------------
def process_image(
    input_path: str,
    output_dir: str = None,
    from_lang: str = 'zh',
    to_lang: str = 'en',
    enable_correction: bool = True,
    enable_visualization: bool = True,
    card_side: str = 'front'
) -> Dict[str, Any]:
    """
    完整的图片处理流程
    
    Args:
        input_path: 输入文件路径
        output_dir: 输出目录
        from_lang: 源语言
        to_lang: 目标语言
        enable_correction: 是否启用透视矫正
        enable_visualization: 是否生成可视化图片
        card_side: 证件面，'front'=正面（自动换行），'back'=背面（不换行）
    
    Returns:
        包含处理结果的字典
    """
    if output_dir is None:
        output_dir = settings.OUTPUT_DIR
    
    os.makedirs(output_dir, exist_ok=True)
    
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    
    print("=" * 60)
    print(f"🚀 开始处理图片 | card_side={card_side} | auto_wrap={card_side != 'back'}")
    print("=" * 60)

    # 步骤0: 图像预处理
    print("\n🔨 步骤0: 图像标准化处理...")
    img_path = preprocess_resize_image(input_path, target_width=settings.TARGET_IMAGE_WIDTH)

    # 步骤0.5: 透视矫正（可选）
    if enable_correction:
        print("\n📐 步骤0.5: 透视矫正...")
        img_path = auto_correct_perspective(img_path)

    # 步骤1: OCR识别
    print("\n📷 步骤1: OCR 识别中...")
    try:
        result = ocr.predict(img_path)  # PaddleOCR 3.x
    except (AttributeError, NotImplementedError):
        result = ocr.ocr(img_path)  # PaddleOCR 2.x 降级兼容

    # 步骤2: 提取文本
    print("📝 步骤2: 提取文本...")
    items = extract_all_items(result)
    print(f"   检测到 {len(items)} 个文本块")

    # 步骤3: 智能分割
    print("\n🔧 步骤3: 智能分割（保护日期和地址）...")
    items = split_merged_items(items)
    print(f"   分割后共 {len(items)} 个文本块")

    # 步骤4: 智能合并地址行
    print("\n🔗 步骤4: 智能合并地址行（多重判断）...")
    items = merge_address_lines(items)
    print(f"   合并后共 {len(items)} 个文本块")

    # 保存原始OCR结果
    out_json = os.path.join(output_dir, f"{base_name}_raw_ocr.json")
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False, indent=4)
    print(f"   原始OCR结果已保存: {out_json}")

    # 步骤5: 生成可视化（可选）
    vis_path = None
    if enable_visualization:
        print("\n📊 步骤5: 生成分割可视化...")
        vis_path = draw_visualization(img_path, items, 
                                     save_path=os.path.join(output_dir, f"{base_name}_vis.jpg"))

    # 步骤6: 翻译
    if card_side == 'back':
        print("\n🌐 步骤6: 身份证背面专用翻译（固定映射 + 变量翻译）...")
        translated_items = translate_all_items_backside(items, from_lang=from_lang, to_lang=to_lang)
    else:
        print("\n🌐 步骤6: DeepSeek 批量翻译（优化格式）...")
        translated_items = translate_all_items(items, from_lang=from_lang, to_lang=to_lang)

    # 仅背面：对「签发机关」项拉宽 box，使 "Issuing Authority" 单行
    if card_side == 'back':
        for it in translated_items:
            if it.get("original_text") != "签发机关":
                continue
            box = it.get("box")
            if box and len(box) >= 4:
                it["box"] = _expand_box_right(box, BACKSIDE_ISSUING_AUTHORITY_BOX_EXTRA_WIDTH)
                print(f"   → 签发机关 box 右扩 {BACKSIDE_ISSUING_AUTHORITY_BOX_EXTRA_WIDTH}px")

    trans_json = os.path.join(output_dir, f"{base_name}_translated.json")
    with open(trans_json, "w", encoding="utf-8") as f:
        json.dump(translated_items, f, ensure_ascii=False, indent=4)
    print(f"\n✅ 翻译结果已保存: {trans_json}")

    # 步骤7: 智能修复并填充翻译
    print("\n🎨 步骤7: 智能修复并填充翻译...")
    output_img = inpaint_and_fill(
        img_path, 
        translated_items,
        output_path=os.path.join(output_dir, f"{base_name}_translated.jpg")
    )

    print("\n" + "=" * 60)
    print("✅ 全部完成！")
    print("=" * 60)

    return {
        "input_path": input_path,
        "processed_image": img_path,
        "raw_ocr_json": out_json,
        "translated_json": trans_json,
        "visualization": vis_path,
        "final_output": output_img,
        "items_count": len(items)
    }

