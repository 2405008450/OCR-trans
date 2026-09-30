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
    "〇": 0, "零": 0, "一": 1, "二": 2, "两": 2, "双": 2, "三": 3, "四": 4,
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
    # 带点缩写（关键补充）
    "jan.": "01", "feb.": "02", "mar.": "03", "apr.": "04",
    "jun.": "06", "jul.": "07", "aug.": "08", "sep.": "09",
    "oct.": "10", "nov.": "11", "dec.": "12",
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

# 小写罗马数字：与大写逻辑相同，直接匹配所有独立小写罗马词（含单字母 i）
_ROMAN_LOWER_PATTERN = re.compile(
    r'\b(m{0,4}(?:cm|cd|d?c{0,3})(?:xc|xl|l?x{0,3})(?:ix|iv|v?i{0,3}))\b'
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

# 带圈数字 ①~⑳
_CIRCLED_NUM_MAP = {
    "①": 1, "②": 2, "③": 3, "④": 4, "⑤": 5,
    "⑥": 6, "⑦": 7, "⑧": 8, "⑨": 9, "⑩": 10,
    "⑪": 11, "⑫": 12, "⑬": 13, "⑭": 14, "⑮": 15,
    "⑯": 16, "⑰": 17, "⑱": 18, "⑲": 19, "⑳": 20,
}
_CIRCLED_NUM_PATTERN = re.compile("[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]")


# ─────────────────────────────────────────
# 内部转换函数
# ─────────────────────────────────────────

def _chinese_to_int(s: str, table: dict) -> int:
    """将纯中文数字串转为整数（支持万/亿级）。正序处理。"""
    # 纯位数字串（每个字符都是 0-9 数字，无单位字符）→ 逐位组合
    # 例：二〇二六 → 2026，一九九八 → 1998
    DIGIT_ONLY = {k for k, v in table.items() if 0 <= v <= 9}
    ZERO_CHARS = {"零", "〇"}
    ALL_DIGIT = DIGIT_ONLY | ZERO_CHARS
    if all(c in ALL_DIGIT for c in s):
        result = 0
        for c in s:
            result = result * 10 + table.get(c, 0)
        return result

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
                # 保留原始精度，不去尾零（对比时用 float 比较有效位）
                found.append((m.start(), str(num_val)))
                _mark_consumed(m.start(), m.end())
            except ValueError:
                pass

    # ─── 步骤1b：中文逐字读年份（二〇二六年三月 → 2026-03，优先于步骤2）───
    if s.get("chinese_trad"):
        _cn_yr_digit = {"零": 0, "〇": 0, "一": 1, "二": 2, "三": 3, "四": 4,
                        "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
        _cn_yr_chars = "".join(_cn_yr_digit.keys())
        _cn_mon_map  = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6,
                        "七": 7, "八": 8, "九": 9, "十": 10, "十一": 11, "十二": 12}

        def _cn_year_to_int(s4: str) -> int:
            return sum(_cn_yr_digit.get(c, 0) * (10 ** (3 - i)) for i, c in enumerate(s4))

        # 带月（带日可选）
        for m in re.finditer(
            rf"([{_cn_yr_chars}]{{4}})年(十[一二]|[一二三四五六七八九十])月"
            r"(?:([一二三四五六七八九十]{1,2}|二十[一二三四五六七八九]?|三十[一]?)日)?",
            text
        ):
            year = _cn_year_to_int(m.group(1))
            if not (1000 <= year <= 2100):
                continue
            mon = _cn_mon_map.get(m.group(2), 0)
            day_str = m.group(3)
            if mon and day_str:
                day = _chinese_to_int(day_str, _CHINESE_TRAD_ALL) or 0
                found.append((m.start(), f"{year}-{str(mon).zfill(2)}-{str(day).zfill(2)}"))
            elif mon:
                found.append((m.start(), f"{year}-{str(mon).zfill(2)}"))
            else:
                found.append((m.start(), str(year)))
            _mark_consumed(m.start(), m.end())

        # 仅年份（二〇二六年）
        for m in re.finditer(rf"([{_cn_yr_chars}]{{4}})年", text):
            if _is_consumed(m.start(), m.end()):
                continue
            year = _cn_year_to_int(m.group(1))
            if 1000 <= year <= 2100:
                found.append((m.start(), str(year)))
                _mark_consumed(m.start(), m.end())

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
            r'(\d[\d,]*\.?\d*)\s*(' + scale_words + r')\b',
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
                # 保留原始精度，不去尾零（对比时用 float 比较有效位）
                found.append((m.start(), str(num_val)))
                _mark_consumed(m.start(), m.end())
            except ValueError:
                pass

    # ─── 步骤3b：英文数字单词 + 英文量级词（如 one hundred, ten thousand）───
    if s.get("english_number"):
        word_keys = "|".join(re.escape(w) for w in sorted(_ENGLISH_NUMBER_WORDS.keys(), key=len, reverse=True))
        scale_keys = "|".join(re.escape(w) for w in sorted(_ENGLISH_SCALE_WORDS.keys(), key=len, reverse=True))
        pattern_word_scale = re.compile(
            r'\b(' + word_keys + r')\s+(' + scale_keys + r')\b',
            re.IGNORECASE
        )
        for m in pattern_word_scale.finditer(text):
            if _is_consumed(m.start(), m.end()):
                continue
            base_word = m.group(1).lower()
            scale_word = m.group(2).lower()
            base_val = _ENGLISH_NUMBER_WORDS.get(base_word, 0)
            scale_val = _ENGLISH_SCALE_WORDS.get(scale_word, 1)
            num_val = base_val * scale_val
            if num_val > 0:
                found.append((m.start(), str(num_val)))
                _mark_consumed(m.start(), m.end())

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

    # ─── 步骤4b：带圈数字 ①~⑳ ───
    for m in _CIRCLED_NUM_PATTERN.finditer(text):
        if _is_consumed(m.start(), m.end()):
            continue
        val = _CIRCLED_NUM_MAP.get(m.group(0), 0)
        if val > 0:
            found.append((m.start(), str(val)))
            _mark_consumed(m.start(), m.end())

    # ─── 步骤5：罗马数字（大写独立词 + 小写序号/引用） ───
    if s.get("roman"):
        # C/D/M 单字母作序号极不常见（c=100,d=500,m=1000），排除误识别
        _single_roman_upper = set("IVXL")
        _single_roman_lower = set("ivxl")

        def _is_single_roman_context(m, txt):
            """
            单字母罗马数字需满足序号上下文：
              前面：行首 / '(' / 空白
              后面：'.' / ')' / ',' / 行尾 / 空白+大写字母
            排除缩写（i.e. / e.g.）：后面是 '.' 且 '.' 后紧跟字母
            """
            start, end = m.start(), m.end()
            before = txt[:start]
            after  = txt[end:]
            pre_ok  = (not before) or before[-1] in "( \t\n"
            post_ok = (not after)  or after[0] in ").," or \
                      (after and after[0] in " \t" and len(after) > 1 and after[1].isupper())
            # 排除缩写：后面是 '.' 且 '.' 后紧跟字母（如 i.e. / e.g.）
            if after and after[0] == "." and len(after) > 1 and after[1].isalpha():
                return False
            return pre_ok and post_ok

        # 5a. 大写罗马数字
        for m in _ROMAN_PATTERN.finditer(text):
            if _is_consumed(m.start(), m.end()):
                continue
            s_val = m.group(1)
            if not s_val:
                continue
            if len(s_val) == 1 and s_val in _single_roman_upper:
                if not _is_single_roman_context(m, text):
                    continue
            val = _roman_to_int(s_val)
            if val > 0:
                found.append((m.start(), str(val)))
                _mark_consumed(m.start(), m.end())

        # 5b. 全角/Unicode 罗马数字（Ⅰ~Ⅻ）
        for m in _FULLWIDTH_ROMAN_PATTERN.finditer(text):
            if _is_consumed(m.start(), m.end()):
                continue
            val = _FULLWIDTH_ROMAN_MAP[m.group(0)]
            found.append((m.start(), str(val)))
            _mark_consumed(m.start(), m.end())

        # 5c. 小写罗马数字
        for m in _ROMAN_LOWER_PATTERN.finditer(text):
            s_val = m.group(1)
            if not s_val:
                continue
            start, end = m.start(1), m.end(1)
            if _is_consumed(start, end):
                continue
            if len(s_val) == 1 and s_val in _single_roman_lower:
                if not _is_single_roman_context(m, text):
                    continue
            val = _roman_to_int(s_val)
            if val > 0:
                found.append((start, str(val)))
                _mark_consumed(start, end)

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
            float(num_str)  # 仅做合法性校验
            found.append((m.start(), num_str))  # 保留原始字符串，不去尾零
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
            r'(\d[\d,]*\.?\d*)\s*(' + scale_words + r')\b',
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
                # float 比较天然忽略尾零（14.70 == 14.7），允许微小浮点误差
                if abs(cn_f - en_f) > 1e-9 * max(abs(cn_f), abs(en_f), 1):
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

    cn1 = "v1.0.10,全年重钙生产14.06万吨、销售12.86万吨，实现营收3.89亿元，贡献毛利约8,413万元。14.77;56790.9万,三分之二、五成、八折"
    en1 = ("v1.0.10,TSP production was 140,600 tons and sales were 128,600 tons,"
           "achieving revenue of RMB389million and contributing a gross profit "
           "of approximately RMB 84.13 million.14.770;567.909 million,two thirds,half,twenty-five")

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

    cn4 = " 两个问题,零元（一）各板块经营情况一千一百零二十四   二十、补充资料  二十四  一百二十四 一千一百零二十四"
    en4 = "(I) Operating conditions of each segment  XX. Supplementary Materials  XVIII,two Hundred,Ten thousand,Ten"

    cn_nums4 = extract_numbers(cn4)
    en_nums4 = extract_numbers(en4)
    print(f"  中文数值: {cn_nums4}")
    print(f"  英文数值: {en_nums4}")

    en4b = "i,i.,vi.,Chapter XIV discusses the topic"
    en_nums4b = extract_numbers(en4b)
    print(f"  罗马数字 XIV: {en_nums4b}")

    en4c = """
序号形式         'ii. Chapter 3'                           
独立词          'ii'                                       
序号形式         'iii. Background'                          
括号序号         'iv) Section'                              
正文引用         'see section ii for details'               
正文引用         'type ii diabetes'                        
正文引用         'phase ii trial'                          
序号+大写        'ii.Chapter XIV'                          
缩写排除         'i.e. the result'                          
缩写+单i        'e.g. item i'                              
单独i排除        'item i on the list'                       
括号序号         '(ii) Operating conditions'                
正文多个         'Section iii, Part vi'                    
分号分隔         'i; ii; iii'                               
复合词排除        'vitamin b12'                              
型号排除         'ISO9001:2008'                             
单i序号         'i. Introduction'                          
括号单i         '(i) Background'                          
括号单i2        'i) first item'                            
逗号分隔         'i, ii, iii'                               
正文单i         'the item i is done'                       
缩写+ii        'i.e. see item ii'   """
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
    print("测试7：大写数字测试")
    print("=" * 80)

    text8 = "零元，壹仟贰佰叁拾肆元伍角陆分，壹拾贰万伍仟陆佰零玖元陆角"
    nums8 = extract_numbers(text8)
    print(f"  提取: {nums8}")

    print()
    print("=" * 80)
    print("测试8：日期 年份等测试")
    print("=" * 80)

    cn9 = "2025年6月31日,25年06月31日,2025.6.31,2025-06-31,2025/6/31,20250631"
    en9 = "June 31, 2025;Jun 31, 2025;June 31st, 2025;June 31 2025;06/31/2025;31 June 2025;The 31st of June, 2025;31-Jun-2025;31/06/2025;1.0.2025"

    cn_nums9 = extract_numbers(cn9)
    en_nums9 = extract_numbers(en9)
    print(f"  中文数值: {cn_nums9}")
    print(f"  英文数值: {en_nums9}")

    print()
    print("=" * 80)

    print("测试9：完整段落对比")
    print("=" * 80)

    cn7 = """①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳ vi. 双方，两岸，陆地，六月 ii. 二〇二六，两，双，陆地 两（一）各板块经营情况（021）-38969999
1、重（富）过磷酸钙
2025年6月31日，公司继续发挥柔性生产线优势，根据市场需求动态调整重钙与富过磷酸钙的生产结构。\
全年重钙生产14.06万吨、销售12.86万吨，实现营收3.89亿元，贡献毛利约8,413万元。\
富过磷酸钙生产14.48万吨，销售7.34万吨，实现营收1.63亿元，贡献毛利7,804万元，产品市场认可度持续提升。利润提升8%   Ⅲ型 one hundred,ten thousand"""

    en7 = """Phase I (I) Operating conditions of each segment·（021）-38969999
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