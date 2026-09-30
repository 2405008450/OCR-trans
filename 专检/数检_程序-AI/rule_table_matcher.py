"""
数值归化与提取模块

功能：
  1. 从中文/英文文本中提取所有数值（含单位量级归化）
  2. 支持 "14.06万" → 140600, "340.368 million" → 340368000
  3. 避免把 "万" / "million" 单独替换后与前面数字拼接
  4. 英文数字单词只匹配独立单词，不匹配复合词内部
  5. 罗马数字匹配大写独立词（含单字母 I/V/X 等）
"""

import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

# ─────────────────────────────────────────
# 策略开关默认值
# ─────────────────────────────────────────

DEFAULT_STRATEGIES: Dict[str, bool] = {
    "chinese_upper": True,
    "chinese_trad": True,
    "month_name": True,
    "english_number": True,
    "roman": True,
}

# ─────────────────────────────────────────
# 映射表
# ─────────────────────────────────────────

_CHINESE_UPPER_DIGITS = {
    "零": 0, "壹": 1, "贰": 2, "叁": 3, "肆": 4,
    "伍": 5, "陆": 6, "柒": 7, "捌": 8, "玖": 9,
}

_CHINESE_UPPER_UNITS = {
    "拾": 10, "佰": 100, "仟": 1000,
    "万": 10000, "亿": 100000000,
}

_CHINESE_UPPER_ALL = {**_CHINESE_UPPER_DIGITS, **_CHINESE_UPPER_UNITS}

_CHINESE_TRAD_DIGITS = {
    "〇": 0, "零": 0, "一": 1, "二": 2, "三": 3, "四": 4,
    "五": 5, "六": 6, "七": 7, "八": 8, "九": 9,
}

_CHINESE_TRAD_UNITS = {
    "十": 10, "百": 100, "千": 1000,
    "万": 10000, "亿": 100000000,
}

_CHINESE_TRAD_ALL = {**_CHINESE_TRAD_DIGITS, **_CHINESE_TRAD_UNITS}

# 中文量级词（用于 "14.06万" 这种阿拉伯数字+中文量级的模式）
_CHINESE_SCALE_MAP = {
    "十": 10, "百": 100, "千": 1000, "万": 10000, "亿": 100000000,
    "拾": 10, "佰": 100, "仟": 1000,
}

_MONTH_MAP = {
    "january": 1, "jan": 1,
    "february": 2, "feb": 2,
    "march": 3, "mar": 3,
    "april": 4, "apr": 4,
    "may": 5,
    "june": 6, "jun": 6,
    "july": 7, "jul": 7,
    "august": 8, "aug": 8,
    "september": 9, "sep": 9, "sept": 9,
    "october": 10, "oct": 10,
    "november": 11, "nov": 11,
    "december": 12, "dec": 12,
}

# 英文量级词（用于 "340.368 million" 这种模式）
_ENGLISH_SCALE_WORDS = {
    "hundred": 100,
    "thousand": 1_000,
    "million": 1_000_000,
    "billion": 1_000_000_000,
    "trillion": 1_000_000_000_000,
}

_ENGLISH_CARDINAL = {
    "zero": 0, "one": 1, "two": 2, "three": 3, "four": 4,
    "five": 5, "six": 6, "seven": 7, "eight": 8, "nine": 9,
    "ten": 10, "eleven": 11, "twelve": 12, "thirteen": 13,
    "fourteen": 14, "fifteen": 15, "sixteen": 16, "seventeen": 17,
    "eighteen": 18, "nineteen": 19, "twenty": 20, "thirty": 30,
    "forty": 40, "fifty": 50, "sixty": 60, "seventy": 70,
    "eighty": 80, "ninety": 90,
}

_ENGLISH_ORDINAL = {
    "first": 1, "second": 2, "third": 3, "fourth": 4, "fifth": 5,
    "sixth": 6, "seventh": 7, "eighth": 8, "ninth": 9, "tenth": 10,
    "eleventh": 11, "twelfth": 12, "thirteenth": 13, "fourteenth": 14,
    "fifteenth": 15, "sixteenth": 16, "seventeenth": 17, "eighteenth": 18,
    "nineteenth": 19, "twentieth": 20, "thirtieth": 30, "fortieth": 40,
    "fiftieth": 50, "sixtieth": 60, "seventieth": 70, "eightieth": 80,
    "ninetieth": 90,
}

# 合并基数词和序数词（不含量级词，量级词单独处理）
_ENGLISH_NUMBER_WORDS = {**_ENGLISH_CARDINAL, **_ENGLISH_ORDINAL}

_ROMAN_VALUES = {"M": 1000, "D": 500, "C": 100, "L": 50, "X": 10, "V": 5, "I": 1}

# 罗马数字正则：匹配大写独立词（保持原始逻辑，含单字母）
_ROMAN_PATTERN = re.compile(
    r'\b(M{0,4}(?:CM|CD|D?C{0,3})(?:XC|XL|L?X{0,3})(?:IX|IV|V?I{0,3}))\b'
)

# 全角/Unicode 罗马数字（Ⅰ~Ⅻ 及 ⅰ~ⅻ）→ 整数
_FULLWIDTH_ROMAN_MAP = {
    "Ⅰ": 1, "Ⅱ": 2, "Ⅲ": 3, "Ⅳ": 4, "Ⅴ": 5, "Ⅵ": 6,
    "Ⅶ": 7, "Ⅷ": 8, "Ⅸ": 9, "Ⅹ": 10, "Ⅺ": 11, "Ⅻ": 12,
    "ⅰ": 1, "ⅱ": 2, "ⅲ": 3, "ⅳ": 4, "ⅴ": 5, "ⅵ": 6,
    "ⅶ": 7, "ⅷ": 8, "ⅸ": 9, "ⅹ": 10, "ⅺ": 11, "ⅻ": 12,
}
_FULLWIDTH_ROMAN_PATTERN = re.compile(
    "[" + "".join(re.escape(c) for c in _FULLWIDTH_ROMAN_MAP) + "]"
)


# ─────────────────────────────────────────
# 内部转换函数
# ─────────────────────────────────────────

def _chinese_to_int(s: str, table: dict) -> int:
    """将纯中文数字串转为整数（支持万/亿级）。正序处理。"""

    # 先按万/亿拆段，再处理每段内的千/百/十
    # 例：三亿二千五百万六千七百八十九
    # 策略：正序扫描，遇到高级单位（万/亿）时结算当前段

    def _parse_section(chars: list) -> int:
        """解析一个不含万/亿的段（最大到千位）。"""
        # 零/〇/壹零 等占位符直接跳过
        ZERO_CHARS = {"零", "〇"}
        sec = 0
        pending_num = None
        for ch in chars:
            if ch in ZERO_CHARS:
                continue  # 占位零，忽略
            v = table.get(ch, 0)
            if v >= 10:  # 十/百/千
                sec += (pending_num if pending_num is not None else 1) * v
                pending_num = None
            else:  # 数字字符
                pending_num = v
        if pending_num is not None:
            sec += pending_num
        return sec

    # 按亿/万拆分
    # 找出所有亿/万的位置
    BIG_UNITS = {}
    for ch, v in table.items():
        if v >= 10000:
            BIG_UNITS[ch] = v

    if not BIG_UNITS:
        return _parse_section(list(s))

    # 正序扫描，遇到大单位时结算
    result = 0
    section_chars: list = []
    for ch in s:
        if ch in BIG_UNITS:
            sec_val = _parse_section(section_chars)
            result += (sec_val if sec_val else 1) * BIG_UNITS[ch]
            section_chars = []
        else:
            section_chars.append(ch)
    # 最后剩余段
    result += _parse_section(section_chars)
    return result


def _roman_to_int(s: str) -> int:
    """罗马数字转整数。"""
    val = 0
    prev = 0
    for ch in reversed(s.upper()):
        cur = _ROMAN_VALUES.get(ch, 0)
        val += cur if cur >= prev else -cur
        prev = cur
    return val


def _format_number(value: float) -> str:
    """将数值格式化为字符串，整数不带小数点，浮点数去掉尾部多余的零。"""
    if value == int(value):
        return str(int(value))
    return f"{value:.10f}".rstrip("0").rstrip(".")


# ─────────────────────────────────────────
# 提取数值列表（核心功能）
# ─────────────────────────────────────────

def extract_numbers(text: str, strategies: Optional[Dict[str, bool]] = None) -> List[str]:
    """
    从文本中提取所有数值，返回字符串列表。

    处理逻辑：
      1. 先处理 "阿拉伯数字+中文量级" 模式（如 14.06万 → 140600）
      2. 再处理纯中文数字串（如 三百二十一 → 321）
      3. 处理 "阿拉伯数字+英文量级" 模式（如 340.368 million → 340368000）
      4. 处理英文数字单词（如 twenty → 20）
      5. 处理罗马数字（如 XIV → 14，含单字母 I → 1）
      6. 处理英文月份（如 January → 1）
      7. 最后提取所有剩余的纯阿拉伯数字

    Args:
        text:       输入文本
        strategies: 策略开关字典，缺省使用 DEFAULT_STRATEGIES

    Returns:
        提取到的数值字符串列表（按出现顺序）
    """
    if not text:
        return []

    s = strategies or DEFAULT_STRATEGIES

    # 用 (position, value_str) 收集所有找到的数值，最后按位置排序
    found: List[Tuple[int, str]] = []
    # 记录已被消费的字符区间，避免重复提取
    consumed: List[Tuple[int, int]] = []

    def _is_consumed(start: int, end: int) -> bool:
        """检查区间是否与已消费区间重叠。"""
        for cs, ce in consumed:
            if start < ce and end > cs:
                return True
        return False

    def _mark_consumed(start: int, end: int) -> None:
        consumed.append((start, end))

    # ─── 步骤1：阿拉伯数字 + 中文量级词 ───
    if s.get("chinese_upper") or s.get("chinese_trad"):
        scale_chars = "".join(_CHINESE_SCALE_MAP.keys())
        pattern_num_cn_scale = re.compile(
            r'(\d[\d,]*\.?\d*)\s*([' + scale_chars + r'])'
        )
        for m in pattern_num_cn_scale.finditer(text):
            num_str = m.group(1).replace(",", "")
            scale_char = m.group(2)
            scale_val = _CHINESE_SCALE_MAP.get(scale_char, 1)
            try:
                num_val = float(num_str) * scale_val
                found.append((m.start(), _format_number(num_val)))
                _mark_consumed(m.start(), m.end())
            except ValueError:
                pass

    # ─── 步骤2：纯中文数字串（合并两张表的字符集一次性分词，避免零桥接冲突）───
    if s.get("chinese_upper") or s.get("chinese_trad"):
        # 构建联合字符集（含零/〇）
        combined_chars: set = set()
        if s.get("chinese_upper"):
            combined_chars |= set(_CHINESE_UPPER_ALL.keys())
        if s.get("chinese_trad"):
            combined_chars |= set(_CHINESE_TRAD_ALL.keys())

        ZERO_CHARS = {"零", "〇"}
        n = len(text)
        i = 0
        while i < n:
            if text[i] in combined_chars:
                j = i
                while j < n:
                    if text[j] in combined_chars:
                        j += 1
                    elif text[j] in ZERO_CHARS and j + 1 < n and text[j + 1] in combined_chars:
                        j += 1
                    else:
                        break
                if not _is_consumed(i, j):
                    segment = text[i:j]
                    # 优先用 upper 表转换，失败或结果为0时用 trad 表
                    val = None
                    if s.get("chinese_upper") and all(
                            c in _CHINESE_UPPER_ALL or c in ZERO_CHARS for c in segment
                    ):
                        val = _chinese_to_int(segment, _CHINESE_UPPER_ALL)
                    if val is None and s.get("chinese_trad") and all(
                            c in _CHINESE_TRAD_ALL or c in ZERO_CHARS for c in segment
                    ):
                        val = _chinese_to_int(segment, _CHINESE_TRAD_ALL)
                    if val is not None:
                        found.append((i, str(val)))
                        _mark_consumed(i, j)
                i = j
            elif text[i] in ZERO_CHARS:
                if not _is_consumed(i, i + 1):
                    found.append((i, "0"))
                    _mark_consumed(i, i + 1)
                i += 1
            else:
                i += 1

    # ─── 步骤3：阿拉伯数字 + 英文量级词 ───
    if s.get("english_number"):
        scale_words = "|".join(_ENGLISH_SCALE_WORDS.keys())
        pattern_num_en_scale = re.compile(
            r'(\d[\d,]*\.?\d*)\s+(' + scale_words + r')\b',
            re.IGNORECASE
        )
        for m in pattern_num_en_scale.finditer(text):
            if _is_consumed(m.start(), m.end()):
                continue
            num_str = m.group(1).replace(",", "")
            scale_word = m.group(2).lower()
            scale_val = _ENGLISH_SCALE_WORDS.get(scale_word, 1)
            try:
                num_val = float(num_str) * scale_val
                found.append((m.start(), _format_number(num_val)))
                _mark_consumed(m.start(), m.end())
            except ValueError:
                pass

    # ─── 步骤4：英文数字单词（独立词，不含量级词） ───
    if s.get("english_number"):
        word_list = sorted(_ENGLISH_NUMBER_WORDS.keys(), key=len, reverse=True)
        word_pat = re.compile(
            r'\b(' + '|'.join(re.escape(w) for w in word_list) + r')\b',
            re.IGNORECASE
        )
        for m in word_pat.finditer(text):
            if _is_consumed(m.start(), m.end()):
                continue
            key = m.group(0).lower()
            if key in _ENGLISH_NUMBER_WORDS:
                found.append((m.start(), str(_ENGLISH_NUMBER_WORDS[key])))
                _mark_consumed(m.start(), m.end())

    # ─── 步骤5：罗马数字（保持原始逻辑，仅大写独立词） ───
    if s.get("roman"):
        def _replace_roman(m: re.Match) -> Optional[Tuple[int, str]]:
            s_val = m.group(1)
            if not s_val:
                return None
            val = _roman_to_int(s_val)
            return (m.start(), str(val)) if val > 0 else None

        for m in _ROMAN_PATTERN.finditer(text):
            if _is_consumed(m.start(), m.end()):
                continue
            s_val = m.group(1)
            if not s_val:
                continue
            val = _roman_to_int(s_val)
            if val > 0:
                found.append((m.start(), str(val)))
                _mark_consumed(m.start(), m.end())

        # 全角/Unicode 罗马数字（Ⅰ~Ⅻ）
        for m in _FULLWIDTH_ROMAN_PATTERN.finditer(text):
            if _is_consumed(m.start(), m.end()):
                continue
            val = _FULLWIDTH_ROMAN_MAP[m.group(0)]
            found.append((m.start(), str(val)))
            _mark_consumed(m.start(), m.end())

    # ─── 步骤6：英文月份 ───
    if s.get("month_name"):
        month_pat = re.compile(
            r'\b(January|February|March|April|May|June|July|August|'
            r'September|October|November|December|'
            r'Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\.?\b',
            re.IGNORECASE
        )
        for m in month_pat.finditer(text):
            if _is_consumed(m.start(), m.end()):
                continue
            key = m.group(1).lower()
            if key in _MONTH_MAP:
                found.append((m.start(), str(_MONTH_MAP[key])))
                _mark_consumed(m.start(), m.end())

    # ─── 步骤7：剩余的纯阿拉伯数字 ───
    num_pat = re.compile(r'-?\d[\d,]*\.?\d*')
    for m in num_pat.finditer(text):
        if _is_consumed(m.start(), m.end()):
            continue
        num_str = m.group(0).replace(",", "")
        try:
            val = float(num_str)
            found.append((m.start(), _format_number(val)))
            _mark_consumed(m.start(), m.end())
        except ValueError:
            pass

    # 按出现位置排序
    found.sort(key=lambda x: x[0])
    return [v for _, v in found]


def _extract_chinese_sequence(
        text: str,
        chars: set,
        table: dict,
        found: List[Tuple[int, str]],
        consumed: List[Tuple[int, int]],
        is_consumed_fn,
        mark_consumed_fn,
) -> None:
    """提取文本中连续的纯中文数字串。

    "零/〇" 作为透明连接符：只要其后紧跟本表字符，就纳入当前段；
    否则（孤立出现）单独提取为 0。
    """
    ZERO_CHARS = {"零", "〇"}
    n = len(text)
    i = 0
    while i < n:
        if text[i] in chars:
            j = i
            # 贪心扫描：本表字符 + 零作为桥接（零后面还有本表字符时才纳入）
            while j < n:
                if text[j] in chars:
                    j += 1
                elif text[j] in ZERO_CHARS and j + 1 < n and text[j + 1] in chars:
                    # 零后面还有本表字符，纳入作为桥接
                    j += 1
                else:
                    break
            if not is_consumed_fn(i, j):
                segment = text[i:j]
                val = _chinese_to_int(segment, table)
                found.append((i, str(val)))
                mark_consumed_fn(i, j)
            i = j
        elif text[i] in ZERO_CHARS:
            # 孤立的零（前后都不是本表字符），单独提取为 0
            if not is_consumed_fn(i, i + 1):
                found.append((i, "0"))
                mark_consumed_fn(i, i + 1)
            i += 1
        else:
            i += 1


# ─────────────────────────────────────────
# 文本归化接口（替换式）
# ─────────────────────────────────────────

def normalize(text: str, strategies: Optional[Dict[str, bool]] = None) -> str:
    """
    对文本执行数值归化，返回归化后的文本。

    Args:
        text:       输入文本
        strategies: 策略开关字典，缺省使用 DEFAULT_STRATEGIES

    Returns:
        归化后的文本（非数字部分保持不变）
    """
    if not text:
        return text

    s = strategies or DEFAULT_STRATEGIES

    # 1. 阿拉伯数字 + 中文量级词（必须在纯中文数字处理之前）
    if s.get("chinese_upper") or s.get("chinese_trad"):
        scale_chars = "".join(_CHINESE_SCALE_MAP.keys())
        pattern = re.compile(
            r'(\d[\d,]*\.?\d*)\s*([' + scale_chars + r'])'
        )

        def _replace_num_cn_scale(m: re.Match) -> str:
            num_str = m.group(1).replace(",", "")
            scale_char = m.group(2)
            scale_val = _CHINESE_SCALE_MAP.get(scale_char, 1)
            try:
                num_val = float(num_str) * scale_val
                return _format_number(num_val)
            except ValueError:
                return m.group(0)

        text = pattern.sub(_replace_num_cn_scale, text)

    # 2. 纯中文大写数字
    if s.get("chinese_upper"):
        text = _parse_chinese(text, _CHINESE_UPPER_ALL)

    # 3. 纯中文繁体数字
    if s.get("chinese_trad"):
        text = _parse_chinese(text, _CHINESE_TRAD_ALL)

    # 4. 阿拉伯数字 + 英文量级词
    if s.get("english_number"):
        scale_words = "|".join(_ENGLISH_SCALE_WORDS.keys())
        pattern = re.compile(
            r'(\d[\d,]*\.?\d*)\s+(' + scale_words + r')\b',
            re.IGNORECASE
        )

        def _replace_num_en_scale(m: re.Match) -> str:
            num_str = m.group(1).replace(",", "")
            scale_word = m.group(2).lower()
            scale_val = _ENGLISH_SCALE_WORDS.get(scale_word, 1)
            try:
                num_val = float(num_str) * scale_val
                return _format_number(num_val)
            except ValueError:
                return m.group(0)

        text = pattern.sub(_replace_num_en_scale, text)

    # 5. 英文月份
    if s.get("month_name"):
        month_pat = re.compile(
            r'\b(January|February|March|April|May|June|July|August|'
            r'September|October|November|December|'
            r'Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\.?\b',
            re.IGNORECASE
        )

        def _replace_month(m: re.Match) -> str:
            key = m.group(1).lower()
            return str(_MONTH_MAP.get(key, m.group(0)))

        text = month_pat.sub(_replace_month, text)

    # 6. 英文数字单词（不含量级词）
    if s.get("english_number"):
        word_list = sorted(_ENGLISH_NUMBER_WORDS.keys(), key=len, reverse=True)
        word_pat = re.compile(
            r'\b(' + '|'.join(re.escape(w) for w in word_list) + r')\b',
            re.IGNORECASE
        )
        text = word_pat.sub(
            lambda m: str(_ENGLISH_NUMBER_WORDS[m.group(0).lower()]),
            text
        )

    # 7. 罗马数字（保持原始逻辑：大写独立词，含单字母）
    if s.get("roman"):
        def _replace_roman(m: re.Match) -> str:
            s_val = m.group(1)
            if not s_val:
                return m.group(0)
            val = _roman_to_int(s_val)
            return str(val) if val > 0 else m.group(0)

        text = _ROMAN_PATTERN.sub(_replace_roman, text)

        # 全角/Unicode 罗马数字（Ⅰ~Ⅻ）
        text = _FULLWIDTH_ROMAN_PATTERN.sub(
            lambda m: str(_FULLWIDTH_ROMAN_MAP[m.group(0)]), text
        )

    return text


def _parse_chinese(text: str, table: dict) -> str:
    """将文本中连续的纯中文数字串替换为阿拉伯数字字符串。"""
    chars = set(table.keys())
    result = []
    i = 0
    while i < len(text):
        if text[i] in chars:
            j = i
            while j < len(text) and text[j] in chars:
                j += 1
            segment = text[i:j]
            result.append(str(_chinese_to_int(segment, table)))
            i = j
        else:
            result.append(text[i])
            i += 1
    return "".join(result)


# ─────────────────────────────────────────
# 数值对比工具
# ─────────────────────────────────────────

@dataclass
class CompareResult:
    """数值对比结果。"""
    cn_numbers: List[str] = field(default_factory=list)
    en_numbers: List[str] = field(default_factory=list)
    matched: bool = False
    mismatches: List[dict] = field(default_factory=list)


def compare_numbers(
        cn_text: str,
        en_text: str,
        strategies: Optional[Dict[str, bool]] = None,
) -> CompareResult:
    """
    分别从中文和英文文本中提取数值列表，进行逐项对比。

    Args:
        cn_text:    中文文本
        en_text:    英文文本
        strategies: 策略开关

    Returns:
        CompareResult 对象
    """
    cn_nums = extract_numbers(cn_text, strategies)
    en_nums = extract_numbers(en_text, strategies)

    result = CompareResult(cn_numbers=cn_nums, en_numbers=en_nums)

    max_len = max(len(cn_nums), len(en_nums))
    mismatches = []
    all_match = True

    for i in range(max_len):
        cn_val = cn_nums[i] if i < len(cn_nums) else "<缺失>"
        en_val = en_nums[i] if i < len(en_nums) else "<missing>"

        try:
            cn_f = float(cn_val) if cn_val != "<缺失>" else None
            en_f = float(en_val) if en_val != "<missing>" else None

            if cn_f is not None and en_f is not None:
                # 允许微小浮点误差
                if abs(cn_f - en_f) > 1e-6 * max(abs(cn_f), abs(en_f), 1):
                    all_match = False
                    mismatches.append({"index": i, "cn": cn_val, "en": en_val})
            else:
                all_match = False
                mismatches.append({"index": i, "cn": cn_val, "en": en_val})
        except (ValueError, TypeError):
            if cn_val != en_val:
                all_match = False
                mismatches.append({"index": i, "cn": cn_val, "en": en_val})

    result.matched = all_match and len(cn_nums) == len(en_nums)
    result.mismatches = mismatches
    return result


# ─────────────────────────────────────────
# 测试用例
# ─────────────────────────────────────────

if __name__ == "__main__":
    print("=" * 80)
    print("测试1：阿拉伯数字 + 中文量级词")
    print("=" * 80)

    cn1 = "全年重钙生产14.06万吨、销售12.86万吨，实现营收3.89亿元，贡献毛利约8,413万元。"
    en1 = ("TSP production was 140,600 tons and sales were 128,600 tons, "
           "achieving revenue of RMB 389 million and contributing a gross profit "
           "of approximately RMB 84.13 million.")

    cn_nums1 = extract_numbers(cn1)
    en_nums1 = extract_numbers(en1)
    print(f"  中文数值: {cn_nums1}")
    print(f"  英文数值: {en_nums1}")

    result1 = compare_numbers(cn1, en1)
    print(f"  匹配结果: {result1.matched}")
    if result1.mismatches:
        print(f"  不匹配项: {result1.mismatches}")

    print()
    print("=" * 80)
    print("测试2：340.368 million")
    print("=" * 80)

    cn2 = "2025年全国工业饲料总产量达34036.8万吨，同比增长8.2%"
    en2 = ("China's total industrial feed output in 2025 reached "
           "340.368 million tons, a year-on-year increase of 8.2%.")

    cn_nums2 = extract_numbers(cn2)
    en_nums2 = extract_numbers(en2)
    print(f"  中文数值: {cn_nums2}")
    print(f"  英文数值: {en_nums2}")

    result2 = compare_numbers(cn2, en2)
    print(f"  匹配结果: {result2.matched}")
    if result2.mismatches:
        print(f"  不匹配项: {result2.mismatches}")

    print()
    print("=" * 80)
    print("测试3：ISO9001:2008 不应被量级词干扰")
    print("=" * 80)

    cn3 = "通过了ISO9001:2008及FAMI-QS产品质量管理体系认证"
    en3 = ("passed the ISO9001:2008 and FAMI-QS product quality "
           "management system certifications")

    cn_nums3 = extract_numbers(cn3)
    en_nums3 = extract_numbers(en3)
    print(f"  中文数值: {cn_nums3}")
    print(f"  英文数值: {en_nums3}")

    print()
    print("=" * 80)
    print("测试4：罗马数字（保持原始逻辑，含单字母）")
    print("=" * 80)

    cn4 = " 零元（一）各板块经营情况一千一百零二十四   二十、补充资料  二十四  一百二十四 一千一百零二十四"
    en4 = "(I) Operating conditions of each segment  XX. Supplementary Materials  XVIII"

    cn_nums4 = extract_numbers(cn4)
    en_nums4 = extract_numbers(en4)
    print(f"  中文数值: {cn_nums4}")
    print(f"  英文数值: {en_nums4}")

    en4b = "Chapter XIV discusses the topic"
    en_nums4b = extract_numbers(en4b)
    print(f"  罗马数字 XIV: {en_nums4b}")

    en4c = "Section III, Part VI,RMBXX.X billion,XXX%."
    en_nums4c = extract_numbers(en4c)
    print(f"  罗马数字 III, VI: {en_nums4c}")

    print()
    print("=" * 80)
    print("测试5：normalize 函数（文本替换模式）")
    print("=" * 80)

    text5 = "产量达34036.8万吨，约340.368 million tons"
    print(f"  原文: {text5}")
    print(f"  归化: {normalize(text5)}")

    text5b = "Chapter III and Section IV"
    print(f"  原文: {text5b}")
    print(f"  归化: {normalize(text5b)}")

    print("测试6：英文数字单词独立匹配")
    print("=" * 80)

    text6 = "FIRST,tennis sixteen tons of material, with sixty workers.second"
    nums6 = extract_numbers(text6)
    print(f"  提取: {nums6}")

    text6b = "the sixteen items"
    nums6b = extract_numbers(text6b)
    print(f"  独立词 sixteen: {nums6b}")

    print()
    print("=" * 80)

    print()
    print("=" * 80)
    print("测试八：大写数字测试")
    print("=" * 80)

    text8 = "零元，壹仟贰佰叁拾肆元伍角陆分，壹拾贰万伍仟陆佰零玖元陆角"
    nums8 = extract_numbers(text8)
    print(f"  提取: {nums8}")

    print()
    print("=" * 80)

    print("测试7：完整段落对比")
    print("=" * 80)

    cn7 = """（一）各板块经营情况
1、重（富）过磷酸钙
2025年6月31日，公司继续发挥柔性生产线优势，根据市场需求动态调整重钙与富过磷酸钙的生产结构。\
全年重钙生产14.06万吨、销售12.86万吨，实现营收3.89亿元，贡献毛利约8,413万元。\
富过磷酸钙生产14.48万吨，销售7.34万吨，实现营收1.63亿元，贡献毛利7,804万元，产品市场认可度持续提升。利润提升8%   Ⅲ型"""

    en7 = """(I) Operating conditions of each segment
1. Triple (Enriched) superphosphate
In June 31, 2025, the Company continued to leverage the advantages of flexible production lines \
and dynamically adjusted the production structure of TSP and enriched superphosphate \
according to market demand. Throughout the year, TSP production was 140,600 tons and \
sales were 128,600 tons, achieving revenue of RMB 389 million and contributing a gross \
profit of approximately RMB 84.13 million. Enriched superphosphate production was \
144,800 tons and sales were 73,400 tons, achieving revenue of RMB 163 million and \
contributing a gross profit of RMB 78.04 million. Product market recognition continued \
to improve.Profit increased by 8%  Type III"""

    result7 = compare_numbers(cn7, en7)
    print(f"  中文数值: {result7.cn_numbers}")
    print(f"  英文数值: {result7.en_numbers}")
    print(f"  匹配结果: {result7.matched}")
    if result7.mismatches:
        print(f"  不匹配项:")
        for mm in result7.mismatches:
            print(f"    #{mm['index']}: 中文={mm['cn']} vs 英文={mm['en']}")

    # ─────────────────────────────────────────
    # 测试九：基于规则表示例对的预处理剔除
    # ─────────────────────────────────────────
    print()
    print("=" * 80)
    print("测试九：规则表示例对预处理 → 剔除非数值表达 → 提取数值对比")
    print("=" * 80)

    # 规则表示例对（直接内嵌，对应 测试/数检分类.xlsx Sheet1）
    # 格式：(类别, 规则说明, 中文示例, 英文示例)
    # 多个示例用 ；分隔，拆开后逐一匹配
    _RULE_EXAMPLES = [
        ("数字/序数词", "0-9用单词，10及以上用数字",       "三本书；15个人",              "three books; 15 people"),
        ("数字/序数词", "科学数据/单位用数字",              "7微克/立方米",                "7 μg/m3"),
        ("标题序号",   "一、二、三 译为 I, II, III",       "一、背景",                    "I. Background"),
        ("百分点",     "percentage point(s) 注意单复数",   "1个百分点；1.2个百分点",       "1 percentage point; 1.2 percentage points"),
        ("分数",       "用单词，形容词加hyphen",            "约占1/6；三分之二的减员",      "approximately one sixth; two-thirds reduction in staff"),
        ("日期",       "月日, 年 / 月年",                  "1998年6月4日；1998年6月",      "June 4, 1998; June 1998"),
        ("日期",       "年代前加the",                      "20世纪30年代",                "the 1930s"),
        ("日期",       "纪年与公元",                       "公元前221年；公元210年",        "221 BC; AD 210"),
        ("金额",       "USD/RMB + 空格 + 千分位",          "1,221,000美元",               "USD 1,221,000"),
        ("金额",       "小数点保留2位，超过则换单位",        "12.215百万美元 (不合规)",      "USD 12,215,000 (避免三位小数)"),
        ("单位",       "全称原则",                         "260.5千米",                   "260.5 kilometers"),
        ("单位",       "表格单位格式",                     "单位：亿元",                   "In RMB '00 Million"),
        ("经纬度",     "N/E，秒用两个单引号",              "北纬22°26′59″",               "22°26'59'' N"),
        ("含义数字",   "意译，不直译数字",                  '"9+2"城市群',                 "nine mainland cities and two special administrative regions in the Greater Bay Area"),
        ("期数",       "Phase + 大写罗马数字",             "项目二期",                    "Project Phase II"),
        ("上下标",     "PM2.5下标",                        "PM2.5",                       "PM2.5 (with subscript 2.5)"),
        ("句首数字",   "不能数字开头",                     "16,059家企业被纳入...",        "A total of 16,059 enterprises were included..."),
    ]

    def _strip_rule_examples(cn_text: str, en_text: str, rules: list) -> tuple:
        """
        从 cn_text / en_text 中剔除规则表示例片段。

        Returns:
            (cn_stripped, en_stripped, removed_pairs)
            removed_pairs: list of (category, rule, cn_fragment, en_fragment)
        """
        removed = []
        cn_out = cn_text
        en_out = en_text

        for category, rule_desc, cn_ex, en_ex in rules:
            # 拆分多示例（；或;）
            cn_parts = [p.strip() for p in re.split(r'[；;]', cn_ex) if p.strip()]
            en_parts = [p.strip() for p in re.split(r'[；;]', en_ex) if p.strip()]

            for cp, ep in zip(cn_parts, en_parts):
                cn_hit = cp in cn_out
                en_hit = ep in en_out
                if cn_hit or en_hit:
                    removed.append((category, rule_desc, cp if cn_hit else "—", ep if en_hit else "—"))
                if cn_hit:
                    cn_out = cn_out.replace(cp, " ")
                if en_hit:
                    en_out = en_out.replace(ep, " ")

        return cn_out, en_out, removed

    # ── 用例 A：含义数字（"9+2"城市群）──
    print()
    print("── 用例A：含义数字 ──")
    cn_a = '粤港澳大湾区"9+2"城市群共有16,059家企业被纳入统计范围，同比增长8.2%。'
    en_a = ('The Greater Bay Area, comprising nine mainland cities and two special administrative '
            'regions in the Greater Bay Area, recorded A total of 16,059 enterprises were included'
            ' in the statistics, a year-on-year increase of 8.2%.')
    cn_a2, en_a2, removed_a = _strip_rule_examples(cn_a, en_a, _RULE_EXAMPLES)
    print(f"  原始中文: {cn_a}")
    print(f"  原始英文: {en_a}")
    if removed_a:
        print("  【剔除片段】")
        for cat, rdesc, cf, ef in removed_a:
            print(f"    [{cat}] {rdesc}")
            print(f"      中文片段: {cf}")
            print(f"      英文片段: {ef}")
    print(f"  剔除后中文: {cn_a2.strip()}")
    print(f"  剔除后英文: {en_a2.strip()}")
    r_a = compare_numbers(cn_a2, en_a2)
    print(f"  数值对比 → 匹配: {r_a.matched}  中文={r_a.cn_numbers}  英文={r_a.en_numbers}")
    if r_a.mismatches:
        for mm in r_a.mismatches:
            print(f"    不匹配 #{mm['index']}: 中={mm['cn']} vs 英={mm['en']}")

    # ── 用例 B：日期（月日年格式）──
    print()
    print("── 用例B：日期格式 ──")
    cn_b = "本报告期为1998年6月4日至1998年6月30日，共计26天。"
    en_b = "The reporting period runs from June 4, 1998 to June 30, 1998, totaling 26 days."
    cn_b2, en_b2, removed_b = _strip_rule_examples(cn_b, en_b, _RULE_EXAMPLES)
    print(f"  原始中文: {cn_b}")
    print(f"  原始英文: {en_b}")
    if removed_b:
        print("  【剔除片段】")
        for cat, rdesc, cf, ef in removed_b:
            print(f"    [{cat}] {rdesc}")
            print(f"      中文片段: {cf}  英文片段: {ef}")
    print(f"  剔除后中文: {cn_b2.strip()}")
    print(f"  剔除后英文: {en_b2.strip()}")
    r_b = compare_numbers(cn_b2, en_b2)
    print(f"  数值对比 → 匹配: {r_b.matched}  中文={r_b.cn_numbers}  英文={r_b.en_numbers}")
    if r_b.mismatches:
        for mm in r_b.mismatches:
            print(f"    不匹配 #{mm['index']}: 中={mm['cn']} vs 英={mm['en']}")

    # ── 用例 C：金额（USD千分位）──
    print()
    print("── 用例C：金额格式 ──")
    cn_c = "本次交易总金额为1,221,000美元，折合人民币约8,547,000元。"
    en_c = "The total transaction amount is USD 1,221,000, equivalent to approximately RMB 8,547,000."
    cn_c2, en_c2, removed_c = _strip_rule_examples(cn_c, en_c, _RULE_EXAMPLES)
    print(f"  原始中文: {cn_c}")
    print(f"  原始英文: {en_c}")
    if removed_c:
        print("  【剔除片段】")
        for cat, rdesc, cf, ef in removed_c:
            print(f"    [{cat}] {rdesc}")
            print(f"      中文片段: {cf}  英文片段: {ef}")
    print(f"  剔除后中文: {cn_c2.strip()}")
    print(f"  剔除后英文: {en_c2.strip()}")
    r_c = compare_numbers(cn_c2, en_c2)
    print(f"  数值对比 → 匹配: {r_c.matched}  中文={r_c.cn_numbers}  英文={r_c.en_numbers}")
    if r_c.mismatches:
        for mm in r_c.mismatches:
            print(f"    不匹配 #{mm['index']}: 中={mm['cn']} vs 英={mm['en']}")

    # ── 用例 D：期数（Phase + 罗马数字）──
    print()
    print("── 用例D：期数 ──")
    cn_d = "项目二期总投资达12.5亿元，预计2026年竣工。"
    en_d = "Project Phase II has a total investment of RMB 1.25 billion, expected to be completed in 2026."
    cn_d2, en_d2, removed_d = _strip_rule_examples(cn_d, en_d, _RULE_EXAMPLES)
    print(f"  原始中文: {cn_d}")
    print(f"  原始英文: {en_d}")
    if removed_d:
        print("  【剔除片段】")
        for cat, rdesc, cf, ef in removed_d:
            print(f"    [{cat}] {rdesc}")
            print(f"      中文片段: {cf}  英文片段: {ef}")
    print(f"  剔除后中文: {cn_d2.strip()}")
    print(f"  剔除后英文: {en_d2.strip()}")
    r_d = compare_numbers(cn_d2, en_d2)
    print(f"  数值对比 → 匹配: {r_d.matched}  中文={r_d.cn_numbers}  英文={r_d.en_numbers}")
    if r_d.mismatches:
        for mm in r_d.mismatches:
            print(f"    不匹配 #{mm['index']}: 中={mm['cn']} vs 英={mm['en']}")

    # ── 用例 E：句首数字 ──
    print()
    print("── 用例E：句首数字 ──")
    cn_e = "16,059家企业被纳入统计，同比增长3.5%。"
    en_e = "A total of 16,059 enterprises were included in the statistics, up 3.5% year-on-year."
    cn_e2, en_e2, removed_e = _strip_rule_examples(cn_e, en_e, _RULE_EXAMPLES)
    print(f"  原始中文: {cn_e}")
    print(f"  原始英文: {en_e}")
    if removed_e:
        print("  【剔除片段】")
        for cat, rdesc, cf, ef in removed_e:
            print(f"    [{cat}] {rdesc}")
            print(f"      中文片段: {cf}  英文片段: {ef}")
    print(f"  剔除后中文: {cn_e2.strip()}")
    print(f"  剔除后英文: {en_e2.strip()}")
    r_e = compare_numbers(cn_e2, en_e2)
    print(f"  数值对比 → 匹配: {r_e.matched}  中文={r_e.cn_numbers}  英文={r_e.en_numbers}")
    if r_e.mismatches:
        for mm in r_e.mismatches:
            print(f"    不匹配 #{mm['index']}: 中={mm['cn']} vs 英={mm['en']}")