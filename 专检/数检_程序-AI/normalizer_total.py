"""
normalizer_total.py
====================
从中英文文本中提取数值，统一规范化后进行对比。

规范化原则：
  - 所有数值最终转为 float 字符串（整数去掉小数点）
  - 货币保留货币符号前缀，格式为 "USD/RMB/EUR <数值>"
  - 百分比保留数值（去掉 % / percent）
  - 百分点保留数值（去掉 percentage point(s)）
  - 罗马数字转阿拉伯数字
  - 英文数字词转阿拉伯数字
  - 中文数字转阿拉伯数字
  - 年份范围保留为 "YYYY-YYYY"
  - 日期保留为 "YYYY-MM-DD"
  - 经纬度保留为 "DEG°MIN'SEC'' DIR"
  - 期数保留为 "Phase N"（N 为阿拉伯数字）
  - 财年保留为 "FY YYYY"
  - 世纪保留为 "Nth century"
  - 年代保留为 "YYYYs"
  - 公元前/后保留为 "N BC" / "AD N"
  - 季度保留为 "QN"
  - 下标数字（PM2.5）保留原样
  - X+Y 结构保留为 "X+Y"
  - 分数保留为 "N/D"
"""

import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

# ─────────────────────────────────────────
# 策略开关
# ─────────────────────────────────────────

DEFAULT_STRATEGIES: Dict[str, bool] = {
    "chinese_upper": True,
    "chinese_trad": True,
    "chinese_unit": True,
    "month_name": True,
    "english_number": True,
    "roman": True,
    "decimal": True,
    "currency": True,
    "percentage": True,
    "unit": True,
    "date": True,
    "fraction": True,
    "ordinal": True,
    "year_range": True,
    "plus_expr": True,
    "subscript": True,
    "coordinate": True,
    "fiscal_year": True,
    "century": True,
    "decade": True,
    "bc_ad": True,
    "quarter": True,
}

# ─────────────────────────────────────────
# 月份映射
# ─────────────────────────────────────────

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

# ─────────────────────────────────────────
# 英文数字词
# ─────────────────────────────────────────

_ENGLISH_NUM = {
    "zero": 0, "one": 1, "two": 2, "three": 3, "four": 4,
    "five": 5, "six": 6, "seven": 7, "eight": 8, "nine": 9,
    "ten": 10, "eleven": 11, "twelve": 12, "thirteen": 13,
    "fourteen": 14, "fifteen": 15, "sixteen": 16,
    "seventeen": 17, "eighteen": 18, "nineteen": 19,
    "twenty": 20, "thirty": 30, "forty": 40, "fifty": 50,
    "sixty": 60, "seventy": 70, "eighty": 80, "ninety": 90,
    "hundred": 100, "thousand": 1000, "million": 1000000, "billion": 1000000000,
}

# 量级词（用于 "one hundred"、"ten thousand" 等组合）
_ENGLISH_SCALE_WORDS = {
    "hundred": 100,
    "thousand": 1000,
    "million": 1000000,
    "billion": 1000000000,
}

# 基数词（不含量级词本身，避免 "hundred hundred" 误匹配）
_ENGLISH_NUMBER_WORDS = {k: v for k, v in _ENGLISH_NUM.items()
                         if k not in _ENGLISH_SCALE_WORDS}

_ORDINAL_MAP = {
    "first": 1, "second": 2, "third": 3, "fourth": 4, "fifth": 5,
    "sixth": 6, "seventh": 7, "eighth": 8, "ninth": 9, "tenth": 10,
    "eleventh": 11, "twelfth": 12, "thirteenth": 13, "fourteenth": 14,
    "fifteenth": 15, "sixteenth": 16, "seventeenth": 17, "eighteenth": 18,
    "nineteenth": 19, "twentieth": 20, "thirtieth": 30, "fortieth": 40,
    "fiftieth": 50, "sixtieth": 60, "seventieth": 70, "eightieth": 80,
    "ninetieth": 90, "hundredth": 100,
}

_FRAC_NUM_MAP = {
    "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
    "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
}

_FRAC_DENOM_MAP = {
    "half": 2, "halves": 2,
    "third": 3, "thirds": 3,
    "fourth": 4, "fourths": 4, "quarter": 4, "quarters": 4,
    "fifth": 5, "fifths": 5,
    "sixth": 6, "sixths": 6,
    "seventh": 7, "sevenths": 7,
    "eighth": 8, "eighths": 8,
    "ninth": 9, "ninths": 9,
    "tenth": 10, "tenths": 10,
}

# ─────────────────────────────────────────
# 中文数字映射
# ─────────────────────────────────────────

_CN_DIGIT = {
    "零": 0, "〇": 0,
    "一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
    "六": 6, "七": 7, "八": 8, "九": 9,
    "壹": 1, "贰": 2, "叁": 3, "肆": 4, "伍": 5,
    "陆": 6, "柒": 7, "捌": 8, "玖": 9,
    "两": 2,  # 口语"两"表示2
    "双": 2,  # "双"表示2（一双、双倍）
}

_CN_UNIT_VAL = {
    "十": 10, "拾": 10,
    "百": 100, "佰": 100,
    "千": 1000, "仟": 1000,
    "万": 10000,
    "亿": 100000000,
}

_CN_UPPER_FRAC = {
    "角": 0.1, "分": 0.01, "厘": 0.001,
}

# 中文万/亿乘数
_CN_LARGE_UNIT = {"万": 10000, "亿": 100000000, "百": 100, "千": 1000}

# 中文序数词
_CN_ORDINAL = {
    "一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
    "六": 6, "七": 7, "八": 8, "九": 9, "十": 10,
}

# 中文季度
_CN_QUARTER = {"一": 1, "二": 2, "三": 3, "四": 4}

# ─────────────────────────────────────────
# 罗马数字
# ─────────────────────────────────────────

_ROMAN_VAL = {
    "I": 1, "V": 5, "X": 10, "L": 50,
    "C": 100, "D": 500, "M": 1000
}

_FULLWIDTH_ROMAN_MAP = {
    "Ⅰ": 1, "Ⅱ": 2, "Ⅲ": 3, "Ⅳ": 4, "Ⅴ": 5, "Ⅵ": 6,
    "Ⅶ": 7, "Ⅷ": 8, "Ⅸ": 9, "Ⅹ": 10, "Ⅺ": 11, "Ⅻ": 12,
    "ⅰ": 1, "ⅱ": 2, "ⅲ": 3, "ⅳ": 4, "ⅴ": 5, "ⅵ": 6,
    "ⅶ": 7, "ⅷ": 8, "ⅸ": 9, "ⅹ": 10, "ⅺ": 11, "ⅻ": 12,
}

# ─────────────────────────────────────────
# REGEX 注册中心
# ─────────────────────────────────────────

PATTERNS = {
    # ── 百分点（优先于百分比）──
    "percentage_point_en": re.compile(
        r"(\d[\d,]*(?:\.\d+)?)\s*percentage\s+points?", re.I
    ),
    # 中文百分点：N个百分点
    "percentage_point_cn": re.compile(
        r"(\d[\d,]*(?:\.\d+)?)\s*个百分点"
    ),

    # ── 百分比 ──
    "percent_en": re.compile(r"(\d[\d,]*(?:\.\d+)?)\s*percent\b", re.I),
    "percent_sym": re.compile(r"(\d[\d,]*(?:\.\d+)?)\s*%"),

    # ── 货币（USD/RMB/EUR + 数字 + 可选 million/billion/thousand）──
    "currency": re.compile(
        r"\b(USD|RMB|EUR)\s*(\d[\d,]*(?:\.\d+)?)"
        r"(?:\s*(million|billion|thousand))?\b",
        re.I
    ),
    # 中文货币：数字 + 可选量级（万/亿/百万/千万/百亿/千亿）+ 货币词
    "currency_cn": re.compile(
        r"(\d[\d,]*(?:\.\d+)?)\s*(百万|千万|百亿|千亿|万|亿)?\s*"
        r"(美元|欧元|元人民币|元|人民币)"
    ),

    # ── 千分位 ──
    "thousand": re.compile(r"\b\d{1,3}(?:,\d{3})+(?:\.\d+)?\b"),

    # ── 大数（5位+） ──
    "big_number": re.compile(r"(?<!\d)\d{5,}(?!\d)"),

    # ── 小数 ──
    "decimal": re.compile(r"(?<!\d)\d+\.\d+(?!\d)"),

    # ── 普通整数（1-4位） ──
    "integer": re.compile(r"(?<!\d)\d{1,4}(?!\d)"),

    # ── 完整日期 YYYY-MM-DD / YYYY/MM/DD ──
    "date_full": re.compile(r"\b(\d{4})[/\-](\d{1,2})[/\-](\d{1,2})\b"),

    # ── 年份范围 ──
    "year_range": re.compile(r"\b(\d{4})\s*[–—\-]\s*(\d{4})\b"),

    # ── 年代（英文：the 1930s）──
    "decade_en": re.compile(r"\b(?:the\s+)?(\d{4})s\b", re.I),
    # 中文年代：20世纪30年代
    "decade_cn": re.compile(r"(\d{2})世纪(\d{2})年代"),

    # ── 世纪（英文：the 21st century）──
    "century_en": re.compile(
        r"\b(?:the\s+)?(\d{1,2})(?:st|nd|rd|th)\s+century\b", re.I
    ),
    # 中文世纪：21世纪
    "century_cn": re.compile(r"(\d{1,2})世纪(?![\d年])"),

    # ── 公元前/公元（英文）──
    "bc_en": re.compile(r"\b(\d+)\s*BC\b", re.I),
    "ad_en": re.compile(r"\bAD\s*(\d+)\b", re.I),
    # 中文公元前/公元
    "bc_cn": re.compile(r"公元前(\d+)年"),
    "ad_cn": re.compile(r"公元(\d+)年"),

    # ── 财年 ──
    "fiscal_year_en": re.compile(r"\bfiscal\s+year\s+(\d{4})\b", re.I),
    "fiscal_year_cn": re.compile(r"(\d{4})\s*财年"),

    # ── 季度（英文）──
    "quarter_q": re.compile(r"\bQ([1-4])\b"),
    "quarter_word_en": re.compile(
        r"\bthe\s+(first|second|third|fourth)\s+quarter\b", re.I
    ),
    # 中文季度：第一季度
    "quarter_cn": re.compile(r"第([一二三四])\s*季度"),

    # ── 单位（英文）──
    "unit_en": re.compile(
        r"(\d[\d,]*(?:\.\d+)?)\s*"
        r"(μg/m[²2³3]?|mg/m[²2³3]?|μg|mg|kg(?!\w)|"
        r"square\s+kilometers?|cubic\s+meters?|"
        r"kilometers?|metres?|meters?|"
        r"m²|m[²2]|m[³3]|kWh)\b",
        re.I
    ),
    # 中文单位：N微克/立方米、N平方千米、N立方米、N千米
    "unit_cn": re.compile(
        r"(\d[\d,]*(?:\.\d+)?)\s*"
        r"(微克/立方米|微克/平方米|微克|毫克|千克|"
        r"平方千米|立方米|千米|公里|千瓦时)"
    ),

    # ── 中文万/亿单位数字（14.06万）──
    "cn_unit_num": re.compile(
        r"(\d[\d,]*(?:\.\d+)?)\s*([万亿百千])(?![元美欧人])"
    ),

    # ── 经纬度（英文格式）──
    "coordinate_en": re.compile(
        r"(\d{1,3})°(\d{1,2})[′'](\d{1,2}(?:\.\d+)?)[″\"'']{1,2}\s*([NSEW])\b"
    ),
    # 中文经纬度：北纬22°26′59″ / 东经113°45′42″
    "coordinate_cn": re.compile(
        r"(?:北纬|南纬|东经|西经)\s*(\d{1,3})°(\d{1,2})[′'′](\d{1,2}(?:\.\d+)?)[″\"″]"
    ),
    "coordinate_cn_dir": re.compile(
        r"(北纬|南纬|东经|西经)\s*(\d{1,3})°(\d{1,2})[′'′](\d{1,2}(?:\.\d+)?)[″\"″]"
    ),

    # ── 期数（英文：Phase II）──
    "phase_en": re.compile(
        r"\bPhase\s+(I{1,3}|IV|VI{0,3}|IX|X{0,3}|[1-9])\b"
    ),
    # 中文期数：二期、三期
    "phase_cn": re.compile(r"([一二三四五六七八九十])\s*期(?!间|末|初|内|间|望)"),

    # ── X+Y 结构 ──
    "plus_expr": re.compile(r'"?(\d+)\s*\+\s*(\d+)"?'),

    # ── 下标数字（PM2.5 / CO2 / H2O / NO2 / SO2）──
    "subscript": re.compile(
        r"\b(PM|CO|NO|SO|H|O|N|C|Fe|Ca|Na|K)(\d+(?:\.\d+)?)\b",
        re.I
    ),

    # ── 罗马数字（大写，独立词）──
    "roman": re.compile(
        r"\b(M{0,4}(?:CM|CD|D?C{0,3})(?:XC|XL|L?X{0,3})(?:IX|IV|V?I{0,3}))\b"
    ),

    # ── 罗马数字（小写，独立词）──
    "roman_lower": re.compile(
        r"\b(m{0,4}(?:cm|cd|d?c{0,3})(?:xc|xl|l?x{0,3})(?:ix|iv|v?i{0,3}))\b"
    ),

    # ── 全角/Unicode 罗马数字（Ⅰ~Ⅻ 及 ⅰ~ⅻ）──
    "roman_fullwidth": re.compile(
        "[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫⅰⅱⅲⅳⅴⅵⅶⅷⅸⅹⅺⅻ]"
    ),

    # ── 月份名 ──
    "month": re.compile(
        r"\b(January|February|March|April|May|June|July|August|"
        r"September|October|November|December|"
        r"Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\.?\b",
        re.I
    ),

    # ── 序数词（英文词形）──
    "ordinal_word": re.compile(
        r"\b(first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth|"
        r"eleventh|twelfth|thirteenth|fourteenth|fifteenth|sixteenth|"
        r"seventeenth|eighteenth|nineteenth|twentieth|thirtieth|fortieth|"
        r"fiftieth|sixtieth|seventieth|eightieth|ninetieth|hundredth)\b",
        re.I
    ),

    # ── 数字序数（1st/2nd/3rd/4th）──
    "ordinal_num": re.compile(r"\b(\d+)(?:st|nd|rd|th)\b", re.I),

    # ── 分数词（one sixth / two-thirds）──
    "fraction_word": re.compile(
        r"\b(one|two|three|four|five|six|seven|eight|nine|ten)\s*[-\s]"
        r"(half|halves|third|thirds|fourth|fourths|quarter|quarters|"
        r"fifth|fifths|sixth|sixths|seventh|sevenths|eighth|eighths|"
        r"ninth|ninths|tenth|tenths)\b",
        re.I
    ),
    # 中文分数：三分之二、四分之一
    "fraction_cn": re.compile(r"([一二三四五六七八九十百]+)分之([一二三四五六七八九十百]+)"),

    # ── 中文大写金额 ──
    "cn_upper_amount": re.compile(
        r"[壹贰叁肆伍陆柒捌玖拾佰仟万亿零]+"
        r"(?:元(?:[零壹贰叁肆伍陆柒捌玖拾佰仟万亿]*))?"
        r"(?:角[零壹贰叁肆伍陆柒捌玖]*)?"
        r"(?:分[零壹贰叁肆伍陆柒捌玖]*)?"
        r"(?:厘[零壹贰叁肆伍陆柒捌玖]*)?"
    ),

    # ── 中文数字串 ──
    "cn_number": re.compile(
        r"[零〇一二三四五六七八九十百千万亿壹贰叁肆伍陆柒捌玖拾佰仟两双]+"
    ),

    # ── 带圈数字 ①~⑳ ──
    "circled_number": re.compile(
        "[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]"
    ),

    # ── 字母序号 a) / b) / A. 等（严格上下文：前为行首/空白/括号，后为)/. 加空白或行尾）──
    "letter_label": re.compile(
        r"(?:^|(?<=[\s（(；;，,]))"   # 前置：行首或分隔符
        r"([A-Za-z])"                  # 单个字母
        r"(?:[)）]|\.(?=[^a-zA-Z.]|$))",  # 后置：) 或 ". "（点后不跟字母/点，排除 e.g. i.e.）
        re.MULTILINE
    ),
}

# ─────────────────────────────────────────
# 工具函数
# ─────────────────────────────────────────

def _roman_to_int(s: str) -> int:
    total, prev = 0, 0
    for c in reversed(s.upper()):
        v = _ROMAN_VAL.get(c, 0)
        total += v if v >= prev else -v
        prev = v
    return total


def _fmt(n) -> str:
    """将数值格式化为字符串，整数去掉小数点"""
    try:
        f = float(n)
        # 用相对误差判断是否为整数，避免浮点累积误差（如 79.79*1e8 = 7979000000.000001）
        i = round(f)
        if abs(f - i) <= abs(f) * 1e-9 + 1e-9:
            return str(i)
        return str(round(f, 10)).rstrip("0").rstrip(".")
    except Exception:
        return str(n)


def _strip_commas(s: str) -> str:
    return s.replace(",", "")


def _cn_to_int(s: str) -> Optional[int]:
    """中文数字字符串 → 整数，如 三十五 → 35"""
    if not s:
        return None
    total = 0
    cur = 0
    for c in s:
        if c in _CN_DIGIT:
            cur = _CN_DIGIT[c]
        elif c in _CN_UNIT_VAL:
            u = _CN_UNIT_VAL[c]
            if u >= 10000:
                total = (total + cur) * u
                cur = 0
            else:
                total += (cur if cur != 0 else 1) * u
                cur = 0
    total += cur
    return total if total > 0 else None


def _cn_upper_to_float(s: str) -> Optional[float]:
    """中文大写金额 → 浮点数，如 壹佰叁拾玖万伍仟捌佰玖拾柒元陆角柒分玖厘"""
    total = 0.0
    cur = 0
    frac = 0.0
    in_frac = False
    for c in s:
        if c in _CN_DIGIT:
            cur = _CN_DIGIT[c]
        elif c == "拾":
            total += (cur if cur != 0 else 1) * 10
            cur = 0
        elif c == "佰":
            total += (cur if cur != 0 else 1) * 100
            cur = 0
        elif c == "仟":
            total += (cur if cur != 0 else 1) * 1000
            cur = 0
        elif c == "万":
            total = (total + cur) * 10000
            cur = 0
        elif c == "亿":
            total = (total + cur) * 100000000
            cur = 0
        elif c == "元":
            total += cur
            cur = 0
            in_frac = True
        elif c == "角" and in_frac:
            frac += cur * 0.1
            cur = 0
        elif c == "分" and in_frac:
            frac += cur * 0.01
            cur = 0
        elif c == "厘" and in_frac:
            frac += cur * 0.001
            cur = 0
        elif c == "零":
            if not in_frac:
                total += cur
            cur = 0
    if not in_frac:
        total += cur
    result = round(total + frac, 4)
    return result if result > 0 else None


_DIR_MAP = {"北纬": "N", "南纬": "S", "东经": "E", "西经": "W"}
_CN_NUM_SIMPLE = {
    "一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
    "六": 6, "七": 7, "八": 8, "九": 9, "十": 10,
}

# ─────────────────────────────────────────
# 主提取函数
# ─────────────────────────────────────────

def extract_numbers(text: str, strategies=None) -> List[str]:
    """
    从文本中提取所有数值，统一规范化后返回字符串列表（按出现位置排序）。
    中英文均支持。
    """
    s = strategies or DEFAULT_STRATEGIES
    found: List[Tuple[int, str]] = []
    consumed: List[Tuple[int, int]] = []

    def used(a: int, b: int) -> bool:
        return any(a < e and b > st for st, e in consumed)

    def mark(a: int, b: int):
        consumed.append((a, b))

    def add(pos: int, val: str, end: int):
        if not used(pos, end):
            found.append((pos, val))
            mark(pos, end)

    # ══════════════════════════════════════
    # 最高优先级：单位说明行（单位：万元 / Unit: RMB10,000）
    # 提取为货币条目，并标记整段为已消费，避免后续重复提取
    # ══════════════════════════════════════

    # 中文：单位：[量级]元  →  RMB <面值>
    _cn_unit_decl_pat = re.compile(
        r"单位[：:]\s*(百万|千万|百亿|千亿|亿|千|百|万)?\s*(美元|欧元|元人民币|元|人民币)",
        re.I
    )
    _cn_unit_mul = {
        "百万": 1e6, "千万": 1e7, "百亿": 1e10, "千亿": 1e11,
        "亿": 1e8, "千": 1e3, "百": 1e2, "万": 1e4, "": 1,
    }
    _cn_unit_sym = {"美元": "USD", "欧元": "EUR", "元人民币": "RMB", "元": "RMB", "人民币": "RMB"}
    for m in _cn_unit_decl_pat.finditer(text):
        mul = _cn_unit_mul.get(m.group(1) or "", 1)
        sym = _cn_unit_sym.get(m.group(2), "RMB")
        add(m.start(), f"{sym} {_fmt(mul)}", m.end())

    # 英文：Unit: [USD/RMB/EUR][可选数字+量级]
    # 无数字时面值取 1（与中文"单位：元"→ RMB 1 对齐）
    _en_unit_decl_pat = re.compile(
        r"[Uu]nit[：:]\s*(USD|RMB|EUR)\s*(\d[\d,]*(?:\.\d+)?)?"
        r"(?:\s*(million|billion|thousand))?\b",
        re.I
    )
    _en_unit_mul = {"million": 1e6, "billion": 1e9, "thousand": 1e3}
    for m in _en_unit_decl_pat.finditer(text):
        sym = m.group(1).upper()
        num = float(_strip_commas(m.group(2))) if m.group(2) else 1.0
        mul = _en_unit_mul.get((m.group(3) or "").lower(), 1)
        add(m.start(), f"{sym} {_fmt(num * mul)}", m.end())

    # ══════════════════════════════════════
    # 高优先级：复合结构（先匹配，避免被子模式拆散）
    # ══════════════════════════════════════

    # ── 百分点 ──
    if s.get("percentage"):
        for m in PATTERNS["percentage_point_en"].finditer(text):
            add(m.start(), _fmt(_strip_commas(m.group(1))), m.end())
        for m in PATTERNS["percentage_point_cn"].finditer(text):
            add(m.start(), _fmt(_strip_commas(m.group(1))), m.end())

    # ── 百分比 ──
    if s.get("percentage"):
        for m in PATTERNS["percent_en"].finditer(text):
            add(m.start(), _fmt(_strip_commas(m.group(1))), m.end())
        for m in PATTERNS["percent_sym"].finditer(text):
            add(m.start(), _fmt(_strip_commas(m.group(1))), m.end())

    # ── 货币（英文：USD/RMB/EUR + 数字 + 可选 million/billion/thousand）──
    if s.get("currency"):
        _mul = {"million": 1e6, "billion": 1e9, "thousand": 1e3}
        for m in PATTERNS["currency"].finditer(text):
            symbol = m.group(1).upper()
            num = float(_strip_commas(m.group(2)))
            mul = _mul.get((m.group(3) or "").lower(), 1)
            add(m.start(), f"{symbol} {_fmt(num * mul)}", m.end())
        # 中文货币：统一映射到对应货币符号
        _cn_currency_sym = {
            "美元": "USD", "欧元": "EUR",
            "元人民币": "RMB", "元": "RMB", "人民币": "RMB",
        }
        _cn_mul = {"百万": 1e6, "千万": 1e7, "百亿": 1e10, "千亿": 1e11,
                   "万": 1e4, "亿": 1e8, "": 1}
        for m in PATTERNS["currency_cn"].finditer(text):
            num = float(_strip_commas(m.group(1)))
            mul = _cn_mul.get(m.group(2) or "", 1)
            sym = _cn_currency_sym.get(m.group(3), "RMB")
            add(m.start(), f"{sym} {_fmt(num * mul)}", m.end())

    # ── 经纬度 ──
    if s.get("coordinate"):
        for m in PATTERNS["coordinate_en"].finditer(text):
            add(m.start(),
                f"{m.group(1)}°{m.group(2)}'{m.group(3)}'' {m.group(4).upper()}",
                m.end())
        for m in PATTERNS["coordinate_cn_dir"].finditer(text):
            direction = _DIR_MAP.get(m.group(1), "?")
            add(m.start(),
                f"{m.group(2)}°{m.group(3)}'{m.group(4)}'' {direction}",
                m.end())

    # ── 财年 ──
    if s.get("fiscal_year"):
        for m in PATTERNS["fiscal_year_en"].finditer(text):
            add(m.start(), f"FY{m.group(1)}", m.end())
        for m in PATTERNS["fiscal_year_cn"].finditer(text):
            add(m.start(), f"FY{m.group(1)}", m.end())

    # ── 世纪 ──
    if s.get("century"):
        for m in PATTERNS["century_en"].finditer(text):
            add(m.start(), f"{m.group(1)}th century", m.end())
        for m in PATTERNS["century_cn"].finditer(text):
            add(m.start(), f"{m.group(1)}th century", m.end())

    # ── 年代 ──
    if s.get("decade"):
        for m in PATTERNS["decade_en"].finditer(text):
            add(m.start(), f"{m.group(1)}s", m.end())
        for m in PATTERNS["decade_cn"].finditer(text):
            # 20世纪30年代 → 1930s
            century = int(m.group(1))
            decade = int(m.group(2))
            year = (century - 1) * 100 + decade
            add(m.start(), f"{year}s", m.end())

    # ── 公元前/公元 ──
    if s.get("bc_ad"):
        for m in PATTERNS["bc_en"].finditer(text):
            add(m.start(), f"{m.group(1)} BC", m.end())
        for m in PATTERNS["ad_en"].finditer(text):
            add(m.start(), f"AD {m.group(1)}", m.end())
        for m in PATTERNS["bc_cn"].finditer(text):
            add(m.start(), f"{m.group(1)} BC", m.end())
        for m in PATTERNS["ad_cn"].finditer(text):
            add(m.start(), f"AD {m.group(1)}", m.end())

    # ── 季度 ──
    if s.get("quarter"):
        _qmap = {"first": 1, "second": 2, "third": 3, "fourth": 4}
        for m in PATTERNS["quarter_q"].finditer(text):
            add(m.start(), f"Q{m.group(1)}", m.end())
        for m in PATTERNS["quarter_word_en"].finditer(text):
            add(m.start(), f"Q{_qmap[m.group(1).lower()]}", m.end())
        for m in PATTERNS["quarter_cn"].finditer(text):
            n = _CN_NUM_SIMPLE.get(m.group(1), 0)
            if n:
                add(m.start(), f"Q{n}", m.end())

    # ── 年份范围（允许后跟"年"字）──
    if s.get("year_range"):
        for m in re.finditer(r"\b(\d{4})\s*[–—\-]\s*(\d{4})(?:年)?", text):
            add(m.start(), f"{m.group(1)}-{m.group(2)}", m.end())

    # ── 完整日期 ──
    if s.get("date"):
        for m in PATTERNS["date_full"].finditer(text):
            add(m.start(),
                f"{m.group(1)}-{m.group(2).zfill(2)}-{m.group(3).zfill(2)}",
                m.end())
        # 中文逐字读年份：二〇二六年三月 / 二〇二五年 → 2026-03 / 2026
        _cn_digit_char = "零〇一二三四五六七八九"
        _cn_digit_val  = {
            "零": 0, "〇": 0,
            "一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
            "六": 6, "七": 7, "八": 8, "九": 9,
        }
        def _cn_year_str_to_int(s4: str) -> int:
            return sum(_cn_digit_val.get(c, 0) * (10 ** (3 - i)) for i, c in enumerate(s4))
        _cn_month_map = {
            "一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6,
            "七": 7, "八": 8, "九": 9, "十": 10,
            "十一": 11, "十二": 12,
        }
        # 带月日
        for m in re.finditer(
            rf"([{_cn_digit_char}]{{4}})年(十[一二]|[一二三四五六七八九十])月(?:([一二三四五六七八九十]{{1,2}}|二十[一二三四五六七八九]?|三十[一]?)日)?",
            text
        ):
            year = _cn_year_str_to_int(m.group(1))
            if 1000 <= year <= 2100:
                mon = _cn_month_map.get(m.group(2), 0)
                day_str = m.group(3)
                if mon and day_str:
                    day = _cn_to_int(day_str) or 0
                    add(m.start(), f"{year}-{str(mon).zfill(2)}-{str(day).zfill(2)}", m.end())
                elif mon:
                    add(m.start(), f"{year}-{str(mon).zfill(2)}", m.end())
                else:
                    add(m.start(), str(year), m.end())
        # 仅年份（二〇二六年）
        for m in re.finditer(rf"([{_cn_digit_char}]{{4}})年", text):
            if not used(m.start(), m.end()):
                year = _cn_year_str_to_int(m.group(1))
                if 1000 <= year <= 2100:
                    add(m.start(), str(year), m.end())
        # 中文日期格式：YYYY年M月D日
        for m in re.finditer(r"(\d{4})年(\d{1,2})月(\d{1,2})日", text):
            add(m.start(),
                f"{m.group(1)}-{m.group(2).zfill(2)}-{m.group(3).zfill(2)}",
                m.end())
        # 英文日期格式：Month DD, YYYY（如 July 12, 2028）
        _all_months = {**{k: v for k, v in _MONTH_MAP.items()}}
        for m in re.finditer(
            r"\b(January|February|March|April|May|June|July|August|"
            r"September|October|November|December|"
            r"Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\.?\s+"
            r"(\d{1,2}),?\s+(\d{4})\b",
            text, re.I
        ):
            mon = _all_months.get(m.group(1).lower().rstrip("."), 0)
            if mon:
                add(m.start(),
                    f"{m.group(3)}-{str(mon).zfill(2)}-{m.group(2).zfill(2)}",
                    m.end())
        # 英文日期格式：DD Month YYYY（如 12 July 2028）
        for m in re.finditer(
            r"\b(\d{1,2})\s+"
            r"(January|February|March|April|May|June|July|August|"
            r"September|October|November|December|"
            r"Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\.?\s+"
            r"(\d{4})\b",
            text, re.I
        ):
            mon = _all_months.get(m.group(2).lower().rstrip("."), 0)
            if mon:
                add(m.start(),
                    f"{m.group(3)}-{str(mon).zfill(2)}-{m.group(1).zfill(2)}",
                    m.end())

    # ── X+Y 结构 ──
    if s.get("plus_expr"):
        for m in PATTERNS["plus_expr"].finditer(text):
            add(m.start(), f"{m.group(1)}+{m.group(2)}", m.end())

    # ── 中文大写金额 ──
    if s.get("chinese_upper"):
        for m in PATTERNS["cn_upper_amount"].finditer(text):
            val = _cn_upper_to_float(m.group(0))
            if val is not None:
                add(m.start(), _fmt(val), m.end())

    # ── 中文单位（长度/重量/面积等，优先于纯数量单位）──
    if s.get("unit"):
        _cn_unit_norm = {
            "微克/立方米": "μg/m3", "微克/平方米": "μg/m2",
            "微克": "μg", "毫克": "mg", "千克": "kg",
            "平方千米": "square kilometers", "立方米": "cubic meters",
            "千米": "kilometers", "公里": "kilometers",
            "千瓦时": "kWh",
        }
        for m in PATTERNS["unit_cn"].finditer(text):
            num = _strip_commas(m.group(1))
            unit = _cn_unit_norm.get(m.group(2), m.group(2))
            add(m.start(), f"{num} {unit}", m.end())

    # ── 中文万/亿单位数字（14.06万，复合单位优先）──
    if s.get("chinese_unit"):
        _compound_mul = {"百万": 1e6, "千万": 1e7, "百亿": 1e10, "千亿": 1e11}
        # 带中文单位的复合量级（如 1377亿立方米、14.06万平方千米）
        _cn_unit_norm = {
            "微克/立方米": "μg/m3", "微克/平方米": "μg/m2",
            "微克": "μg", "毫克": "mg", "千克": "kg",
            "平方千米": "square kilometers", "立方米": "cubic meters",
            "千米": "kilometers", "公里": "kilometers",
            "千瓦时": "kWh",
        }
        _cn_units_pat = "|".join(re.escape(u) for u in sorted(_cn_unit_norm, key=len, reverse=True))
        for m in re.finditer(
            r"(\d[\d,]*(?:\.\d+)?)\s*(百万|千万|百亿|千亿|[万亿])\s*(" + _cn_units_pat + r")",
            text
        ):
            num = float(_strip_commas(m.group(1)))
            mul = _compound_mul.get(m.group(2)) or _CN_LARGE_UNIT.get(m.group(2), 1)
            unit = _cn_unit_norm.get(m.group(3), m.group(3))
            add(m.start(), f"{_fmt(num * mul)} {unit}", m.end())
        for m in re.finditer(r"(\d[\d,]*(?:\.\d+)?)\s*(百万|千万|百亿|千亿)", text):
            mul = _compound_mul[m.group(2)]
            add(m.start(), _fmt(float(_strip_commas(m.group(1))) * mul), m.end())
        # 带千瓦时的亿/万（如 5亿千瓦时）
        for m in re.finditer(r"(\d[\d,]*(?:\.\d+)?)\s*([万亿])\s*(千瓦时)", text):
            num = float(_strip_commas(m.group(1)))
            mul = _CN_LARGE_UNIT.get(m.group(2), 1)
            add(m.start(), f"{_fmt(num * mul)} kWh", m.end())
        # 普通万/亿单位（排除后跟长度/单位词的情况）
        # 千后面跟米/克/瓦/升/焦等时不作为数量单位
        for m in re.finditer(
            r"(\d[\d,]*(?:\.\d+)?)\s*([万亿百千])(?![米克瓦升焦帕牛安伏欧赫兹])",
            text
        ):
            num = float(_strip_commas(m.group(1)))
            mul = _CN_LARGE_UNIT.get(m.group(2), 1)
            add(m.start(), _fmt(num * mul), m.end())

    # ── 英文单位（含 million/billion 前缀）──
    if s.get("unit"):
        # 先匹配带 million/billion/thousand 的单位（如 500 million kWh）
        _en_mul = {"million": 1e6, "billion": 1e9, "thousand": 1e3}
        for m in re.finditer(
            r"(\d[\d,]*(?:\.\d+)?)\s*(million|billion|thousand)\s+"
            r"(μg/m[²2³3]?|mg/m[²2³3]?|μg|mg|kg(?!\w)|"
            r"square\s+kilometers?|cubic\s+meters?|"
            r"kilometers?|metres?|meters?|m²|m[²2]|m[³3]|kWh)\b",
            text, re.I
        ):
            num = float(_strip_commas(m.group(1)))
            mul = _en_mul.get(m.group(2).lower(), 1)
            unit = m.group(3).strip()
            add(m.start(), f"{_fmt(num * mul)} {unit}", m.end())
        # 普通英文单位
        for m in PATTERNS["unit_en"].finditer(text):
            num = _strip_commas(m.group(1))
            unit = m.group(2).strip()
            add(m.start(), f"{num} {unit}", m.end())

    # ── 千分位数字（允许后跟中文字符）──
    for m in re.finditer(r"\d{1,3}(?:,\d{3})+(?:\.\d+)?", text):
        add(m.start(), _fmt(float(_strip_commas(m.group(0)))), m.end())

    # ── 下标数字（PM2.5 / CO2 / H2O，优先于小数）──
    if s.get("subscript"):
        for m in PATTERNS["subscript"].finditer(text):
            add(m.start(), f"{m.group(1)}{m.group(2)}", m.end())

    # ── 小数 ──
    if s.get("decimal"):
        for m in PATTERNS["decimal"].finditer(text):
            add(m.start(), m.group(0), m.end())

    # ── 大数（5位+） ──
    for m in PATTERNS["big_number"].finditer(text):
        add(m.start(), m.group(0), m.end())

    # ── 带圈数字 ①~⑳ ──
    _CIRCLED_NUM_MAP = {
        "①": 1, "②": 2, "③": 3, "④": 4, "⑤": 5,
        "⑥": 6, "⑦": 7, "⑧": 8, "⑨": 9, "⑩": 10,
        "⑪": 11, "⑫": 12, "⑬": 13, "⑭": 14, "⑮": 15,
        "⑯": 16, "⑰": 17, "⑱": 18, "⑲": 19, "⑳": 20,
    }
    for m in PATTERNS["circled_number"].finditer(text):
        val = _CIRCLED_NUM_MAP.get(m.group(0), 0)
        if val > 0:
            add(m.start(), str(val), m.end())

    # ── 罗马数字 ──
    if s.get("roman"):
        # C/D/M 单字母作序号极不常见（c=100,d=500,m=1000），排除误识别
        _single_roman_upper = set("IVXL")
        _single_roman_lower = set("ivxl")

        def _is_single_roman_context(m, txt):
            """
            单字母罗马数字需满足序号上下文：
              前面：行首 / '(' / '（' / 空白
              后面：'.' / ')' / '）' / ',' / 行尾 / 空白+字母或数字
            排除缩写（i.e. / e.g.）：后面是 '.' 且 '.' 后紧跟字母
            排除字母序号（c) 等）：后面是 ')' 或 '）'（已在调用处过滤）
            """
            start, end = m.start(), m.end()
            before = txt[:start]
            after  = txt[end:]
            pre_ok  = (not before) or before[-1] in "（( \t\n"
            post_ok = (not after) or after[0] in ").，," or \
                      (after and after[0] in " \t" and len(after) > 1 and after[1].isalpha())
            # 排除缩写：后面是 '.' 且 '.' 后紧跟字母（如 i.e. / e.g.）
            if after and after[0] == "." and len(after) > 1 and after[1].isalpha():
                return False
            return pre_ok and post_ok

        # 大写罗马数字
        for m in PATTERNS["roman"].finditer(text):
            s_val = m.group(1)
            if not s_val:
                continue
            # 单字母：需要序号上下文
            if len(s_val) == 1 and s_val in _single_roman_upper:
                if not _is_single_roman_context(m, text):
                    continue
            val = _roman_to_int(s_val)
            if val > 0:
                add(m.start(), str(val), m.end())

        # 小写罗马数字
        for m in PATTERNS["roman_lower"].finditer(text):
            s_val = m.group(1)
            if not s_val:
                continue
            if len(s_val) == 1 and s_val in _single_roman_lower:
                if not _is_single_roman_context(m, text):
                    continue
            val = _roman_to_int(s_val)
            if val > 0:
                add(m.start(), str(val), m.end())

        # 全角/Unicode（Ⅰ~Ⅻ / ⅰ~ⅻ）— 全角字符本身就是独立符号，无需上下文限制
        for m in PATTERNS["roman_fullwidth"].finditer(text):
            val = _FULLWIDTH_ROMAN_MAP.get(m.group(0), 0)
            if val > 0:
                add(m.start(), str(val), m.end())

    # ── 月份名 ──
    if s.get("month_name"):
        for m in PATTERNS["month"].finditer(text):
            key = m.group(1).lower().rstrip(".")
            if key in _MONTH_MAP:
                add(m.start(), str(_MONTH_MAP[key]), m.end())

    # ── 数字分数（N/N 格式，优先于单独数字）──
    if s.get("fraction"):
        for m in re.finditer(r"(?<!\d)(\d+)/(\d+)(?!\d)", text):
            add(m.start(), f"{m.group(1)}/{m.group(2)}", m.end())

    # ── 分数词（one sixth / two-thirds，优先于序数词和英文数字词）──
    if s.get("fraction"):
        for m in PATTERNS["fraction_word"].finditer(text):
            num = _FRAC_NUM_MAP.get(m.group(1).lower(), 1)
            denom = _FRAC_DENOM_MAP.get(m.group(2).lower(), 1)
            add(m.start(), f"{num}/{denom}", m.end())
        # 中文分数：三分之二 → 2/3
        for m in PATTERNS["fraction_cn"].finditer(text):
            denom = _cn_to_int(m.group(1))
            num = _cn_to_int(m.group(2))
            if denom and num:
                add(m.start(), f"{num}/{denom}", m.end())
        # 中文"一半" → 1/2
        for m in re.finditer(r"一半", text):
            add(m.start(), "1/2", m.end())

    # ── 序数词（英文词形）──
    if s.get("ordinal"):
        for m in PATTERNS["ordinal_word"].finditer(text):
            key = m.group(1).lower()
            if key in _ORDINAL_MAP:
                add(m.start(), str(_ORDINAL_MAP[key]), m.end())

    # ── 数字序数（1st/2nd）──
    if s.get("ordinal"):
        for m in PATTERNS["ordinal_num"].finditer(text):
            add(m.start(), m.group(1), m.end())

    # ── 英文数字词 ──
    if s.get("english_number"):
        # 先匹配 "X point Y Z..." 小数读法（如 three point one four → 3.14）
        _digit_words = {
            "zero": "0", "one": "1", "two": "2", "three": "3", "four": "4",
            "five": "5", "six": "6", "seven": "7", "eight": "8", "nine": "9",
        }
        _dw_pat = "|".join(_digit_words.keys())
        for m in re.finditer(
            rf"\b({_dw_pat})\s+point\s+((?:(?:{_dw_pat})\s*)+)\b",
            text, re.I
        ):
            int_part = str(_digit_words[m.group(1).lower()])
            frac_parts = re.findall(rf"\b({_dw_pat})\b", m.group(2), re.I)
            frac_str = "".join(_digit_words[w.lower()] for w in frac_parts)
            add(m.start(), f"{int_part}.{frac_str}", m.end())

        # 量级组合：word + scale（如 one hundred → 100，ten thousand → 10000）
        _word_keys = "|".join(re.escape(w) for w in sorted(
            _ENGLISH_NUMBER_WORDS.keys(), key=len, reverse=True))
        _scale_keys = "|".join(re.escape(w) for w in sorted(
            _ENGLISH_SCALE_WORDS.keys(), key=len, reverse=True))
        _pat_word_scale = re.compile(
            rf"\b({_word_keys})\s+({_scale_keys})\b", re.IGNORECASE
        )
        for m in _pat_word_scale.finditer(text):
            base_val = _ENGLISH_NUMBER_WORDS.get(m.group(1).lower(), 0)
            scale_val = _ENGLISH_SCALE_WORDS.get(m.group(2).lower(), 1)
            num_val = base_val * scale_val
            if num_val > 0:
                add(m.start(), str(num_val), m.end())

        # 普通英文数字词
        for w, v in sorted(_ENGLISH_NUM.items(), key=lambda x: -len(x[0])):
            for m in re.finditer(rf"\b{re.escape(w)}\b", text, re.I):
                add(m.start(), str(v), m.end())

    # ── 中文数字（兜底）──
    if s.get("chinese_trad"):
        for m in PATTERNS["cn_number"].finditer(text):
            if used(m.start(), m.end()):
                continue
            val = _cn_to_int(m.group(0))
            if val is not None and val > 0:
                add(m.start(), str(val), m.end())

    # ── 普通整数（1-4位，最后兜底）──
    for m in PATTERNS["integer"].finditer(text):
        add(m.start(), m.group(0), m.end())

    # ── 兜底：扫描未被消费的区间，提取其中单独的万/亿/千/百 ──
    # 所有策略执行完后，consumed 之外的字符里可能还有单独的量级词
    _standalone_scale = {"百": 100, "千": 1000, "万": 10000, "亿": 100000000}
    for m in re.finditer(r"[百千万亿]", text):
        pos = m.start()
        if not used(pos, pos + 1):
            mul = _standalone_scale[m.group(0)]
            add(pos, _fmt(mul), pos + 1)

    found.sort(key=lambda x: x[0])
    return [v for _, v in found]

# ─────────────────────────────────────────
# 对比
# ─────────────────────────────────────────

@dataclass
class CompareResult:
    cn_numbers: List[str]
    en_numbers: List[str]
    matched: bool = False
    mismatches: List[dict] = field(default_factory=list)


def compare_numbers(cn: str, en: str) -> CompareResult:
    cn_nums = extract_numbers(cn)
    en_nums = extract_numbers(en)
    res = CompareResult(cn_nums, en_nums)
    ok = True
    for i in range(max(len(cn_nums), len(en_nums))):
        a = cn_nums[i] if i < len(cn_nums) else None
        b = en_nums[i] if i < len(en_nums) else None
        if a != b:
            ok = False
            res.mismatches.append({"i": i, "cn": a, "en": b})
    res.matched = ok
    return res


def print_compare(label: str, cn: str, en: str):
    print("=" * 80)
    print(f"测试：{label}")
    r = compare_numbers(cn, en)
    print(f"中文数值: {r.cn_numbers}")
    print(f"英文数值: {r.en_numbers}")
    print(f"匹配结果: {r.matched}")
    if r.mismatches:
        print("不匹配项:")
        for m in r.mismatches:
            print(f"  #{m['i']}: 中文={m['cn']} vs 英文={m['en']}")
    print()

# ─────────────────────────────────────────
# 测试用例
# ─────────────────────────────────────────

if __name__ == "__main__":

    # ── 数字/序数词 ──
    print_compare("数字词：0-9用单词，10+用数字",
        "三本书；15个人",
        "three books; 15 people")

    print_compare("科学数据/单位",
        "7微克/立方米",
        "7 μg/m3")

    print_compare("句首数字（千分位）",
        "16,059家企业被纳入",
        "A total of 16,059 enterprises were included")

    print_compare("10-odd 结构",
        "十人出头",
        "10-odd people")

    print_compare("序数词拼写",
        "第三次会议",
        "the third meeting")

    print_compare("标题罗马数字",
        "第一章",
        "Chapter I")

    print_compare("年份范围 en dash",
        "2018—2025年",
        "2018–2025")

    print_compare("百万/十亿避免三位小数",
        "12.215百万美元",
        "USD 12,215,000")

    print_compare("千分位",
        "1234567元",
        "RMB 1,234,567")

    print_compare("小数位数",
        "3.1415926公里",
        "3.14 kilometers")

    # ── 标题序号 ──
    print_compare("标题序号 一二三 → I II III",
        "一、背景",
        "I. Background")

    print_compare("子标题阿拉伯数字",
        "（一）工作目标",
        "1. Objectives")

    print_compare("第X部分 → Part V",
        "第五部分",
        "Part V")

    # ── 百分点/百分比 ──
    print_compare("百分点单数",
        "1个百分点",
        "1 percentage point")

    print_compare("百分点复数",
        "1.2个百分点",
        "1.2 percentage points")

    print_compare("百分比 percent",
        "增长5%",
        "increased by 5 percent")

    print_compare("占比 account for",
        "占总量的40%",
        "accounted for 40 percent of the total")

    # ── 分数 ──
    print_compare("分数用单词",
        "约占1/6",
        "approximately one sixth")

    print_compare("分数形容词 hyphen",
        "三分之二的减员",
        "two-thirds reduction in staff")

    print_compare("一半人口",
        "一半人口",
        "one half of the population")

    print_compare("四分之一面积",
        "四分之一面积",
        "one fourth of the area")

    # ── 小数 ──
    print_compare("小数拼写（中文3.14 vs 英文 three point one four）",
        "3.14",
        "three point one four")

    # ── 日期 ──
    print_compare("完整日期",
        "1998-06-04",
        "1998-06-04")

    print_compare("月年",
        "1998-06-01",
        "1998-06-01")

    print_compare("年代",
        "20世纪30年代",
        "the 1930s")

    print_compare("公元前",
        "公元前221年",
        "221 BC")

    print_compare("公元",
        "公元210年",
        "AD 210")

    print_compare("世纪",
        "21世纪",
        "the 21st century")

    print_compare("季度",
        "第一季度",
        "the first quarter")

    print_compare("财年",
        "2025财年",
        "fiscal year 2025")

    # ── 金额 ──
    print_compare("USD千分位",
        "1,221,000美元",
        "USD 1,221,000")

    print_compare("人民币不加Yuan",
        "500万元人民币",
        "RMB 5,000,000")

    print_compare("大额避免小数",
        "15.67亿美元",
        "USD 1,567 million")

    print_compare("中文大写金额",
        "壹佰叁拾玖万伍仟捌佰玖拾柒元陆角柒分玖厘",
        "1,395,897.679")

    print_compare("金额格式统一",
        "总投资20亿元",
        "with a total investment of RMB 2,000,000,000")

    print_compare("欧元",
        "300万欧元",
        "EUR 3,000,000")

    # ── 单位 ──
    print_compare("千米全称",
        "260.5千米",
        "260.5 kilometers")

    print_compare("复合单位缩写",
        "7微克/平方米",
        "7 μg/m2")

    print_compare("kWh保留缩写",
        "5亿千瓦时",
        "500 million kWh")

    print_compare("平方千米",
        "50平方千米",
        "50 square kilometers")

    print_compare("立方米",
        "12立方米",
        "12 cubic meters")

    # ── 经纬度 ──
    print_compare("北纬",
        "北纬22°26′59″",
        "22°26'59'' N")

    print_compare("东经",
        "东经113°45′42″",
        "113°45'42'' E")

    # ── 含义数字 ──
    print_compare("9+2城市群",
        '"9+2"城市群',
        "nine mainland cities and two special administrative regions")

    # ── 期数 ──
    print_compare("项目二期",
        "项目二期",
        "Project Phase II")

    print_compare("三期工程",
        "建设期三期工程",
        "Phase III project")

    # ── 上下标 ──
    print_compare("PM2.5下标",
        "PM2.5",
        "PM2.5")

    # ── 完整段落对比 ──
    print_compare("完整段落对比",
        " 千 千万吨 万吨 万一 i. i) (i) 2027/2/24 ;2026/4/15 最近三个 单位：元	Unit: RMB ii. 二〇二六年三月四日，两，双，陆地2.单位：元 ;单位：万元 截至2025年6月31日，本公司资产总额为140,600万元，净利润128,600万元，"
        "营业收入38.969亿元，其中境外收入8,413万元，境内收入14,480万元，"
        "研发投入16,300万元，同比增长7,804万元，增幅约8%，员工总数3万人。",
        "w,ii Unit: RMB ;RMB 10000;As of June 31, 2025, the Company's total assets were RMB 1,406 million, "
        "net profit RMB 1,286 million, revenue RMB 3,896.9 million, "
        "of which overseas revenue RMB 841.3 million, domestic revenue RMB 1,448 million, "
        "R&D investment RMB 1,630 million, up RMB 780.4 million year on year, "
        "an increase of approximately 8 percent, with a total workforce of 30,000.")
