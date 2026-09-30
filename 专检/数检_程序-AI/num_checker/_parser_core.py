"""
_parser_core.py — FSM 扫描主逻辑
所有提取规则按优先级顺序执行，已消费区间不重复匹配。
"""
import re
from typing import List
from .symbolic_parser import Token, _Consumed, _fmt, _sc, _roman_to_int, _cn_to_int, _cn_upper_to_float
from .symbolic_parser import (
    _CN_DIGIT, _CN_UNIT, _CN_LARGE, _CN_SIMPLE, _CN_QUARTER,
    _EN_NUM, _EN_SCALE, _EN_ORDINAL, _FRAC_NUM, _FRAC_DEN,
    _MONTH, _ROMAN_VAL, _ROMAN_FW, _DIR_MAP,
    _CN_CURRENCY, _CN_MUL, _EN_MUL,
)

# ─────────────────────────────────────────
# 预编译：需要整体跳过的模式（化学式、化学名称、地名数字等）
# 这些区间会被预先标记为"已消费"，阻止后续规则提取其中的数字
# ─────────────────────────────────────────

# 化学式：LiFePO4、FePO4、H2SO4、Ca(H₂PO₄)₂·H₂O 等
# 不用 \b，改用前后不是字母/数字的边界（兼容中文上下文）
_CHEM_FORMULA = re.compile(
    r"(?<![A-Za-z\d])"
    r"(?:[A-Z][a-z]?|\d)+"
    r"(?:\((?:[A-Z][a-z]?|\d|[₀-₉·])+\)[₀-₉\d]*)?"
    r"(?:[A-Z][a-z]?|\d|[₀-₉])+"
    r"(?![A-Za-z\d])"
)

# 化学名称中的数字词（中文）：五氧化二磷、二水法、半水法、磷酸二氢钙 等
# 策略：中文数字后紧跟这些词时，整个词组标记为已消费
_CN_CHEM_SUFFIX = re.compile(
    r"[一二三四五六七八九十百千](?:"
    r"氧化[一二三四五六七八九十]?磷|"   # 五氧化二磷
    r"水法|"                             # 二水法、半水法
    r"水[-\-]|"                          # 二水-（复合工艺名如二水-半水法）
    r"氢[钙钠钾铵]|"                     # 二氢钙、一氢钙
    r"磷酸[钙钠钾铵]|"                   # 磷酸二钙
    r"倍过磷酸钙|"                       # 三倍过磷酸钙
    r"料过磷酸钙|"                       # 三料过磷酸钙
    r"方晶体|"                           # 六方晶体
    r"方体|"                             # 六方体
    r"价"                                # 正三价、正二价
    r")"
)

# 英文序数词在连字符复合词中：second-tier、first-class、third-party 等
_EN_ORDINAL_HYPHEN = re.compile(
    r"\b(first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth"
    r"|eleventh|twelfth|thirteenth|fourteenth|fifteenth|sixteenth"
    r"|seventeenth|eighteenth|nineteenth|twentieth|thirtieth|fortieth"
    r"|fiftieth|sixtieth|seventieth|eightieth|ninetieth|hundredth)"
    r"-\w+",
    re.I
)

# 英文数字词在复合词中（two-way、three-dimensional 等）
_EN_NUM_HYPHEN = re.compile(
    r"\b(one|two|three|four|five|six|seven|eight|nine|ten"
    r"|eleven|twelve|thirteen|fourteen|fifteen|sixteen"
    r"|seventeen|eighteen|nineteen|twenty|thirty|forty|fifty"
    r"|sixty|seventy|eighty|ninety)-\w+",
    re.I
)
# 但 two-thirds / three-quarters 是分数，不应跳过 → 排除分母词
_FRAC_DENOM_SET = {
    "half","halves","third","thirds","fourth","fourths","quarter","quarters",
    "fifth","fifths","sixth","sixths","seventh","sevenths","eighth","eighths",
    "ninth","ninths","tenth","tenths",
}

# 中文量词短语：一种、两类、三个、一批 等（数字+量词，无实际数量意义）
_CN_MEASURE_WORD = re.compile(
    r"[一两二三四五六七八九十]"
    r"(?:种|类|批|项|个|件|条|款|点|步|层(?!楼)|号(?!楼)|期(?!间|末|初|内|望))"
)

# 地名中的数字：四方地、七彩云南、第壹城 等
# 策略：中文数字后紧跟地名特征词时跳过
# 注意：第壹城 中的 壹 是楼盘编号，不跳过（由 cn_upper 规则处理）
_CN_PLACE_NUM = re.compile(
    r"[一二三四五六七八九十壹贰叁肆伍陆柒捌玖]"
    r"(?:方地|彩|星|枢|里(?!程)|坊|村(?!民)|镇(?!政)|街(?!道办)|路(?!程)|区(?!域)|园(?!区))"
)

# P205 写法的化学式（非下标）
_P205_PAT = re.compile(r"\bP\s*2\s*0\s*5\b")

# 化学式末尾孤立数字（LiFePO4 中的 4 被化学式预消费覆盖，但有时正则边界问题）
# 额外处理：字母紧跟数字且前面是化学元素符号
_CHEM_TRAILING_NUM = re.compile(r"(?<=[A-Za-z])(\d+)(?=[^A-Za-z\d]|$)")

# 英文描述性数字词黑名单：在特定搭配下不提取
# "two types of" / "one of the" / "one main" 等描述性用法
_EN_DESCRIPTIVE = re.compile(
    r"\b(one|two|three|four|five|six|seven|eight|nine|ten)\s+"
    r"(?:type|kind|sort|form|way|method|category|class|group|part|aspect|"
    r"of\s+the|of\s+a|main|major|key|important|common|typical|basic|primary)s?\b",
    re.I
)

# 中文"一般"、"一旦"、"一方面"等固定词组中的"一"
_CN_FIXED_PHRASES = re.compile(
    r"一(?:般|旦|方面|定|些|直|起|同|共|致|贯|再|味|概|律|向|贯|切|带|带一路|路)"
)


def extract(text: str) -> List[Token]:
    """主提取函数，返回按位置排序的 Token 列表"""
    tokens: List[Token] = []
    c = _Consumed()
    add = c.add

    # ══ 0. 预消费：整体跳过不应提取数字的区间 ══════════════════════

    # 化学式（LiFePO4、FePO4、H2SO4 等）
    # 排除货币符号、季度标记、常见金融/业务缩写，避免误消费
    _NOT_CHEM = re.compile(
        r'^(?:USD|RMB|EUR|CNY|HKD|GBP|JPY|'   # 货币
        r'Q[1-4]|H[12]|'                        # 季度/半年
        r'DCP|MCP|ESG|GDP|CPI|PPI|IPO|EPS|ROE|ROA|EBITDA|CAGR|'  # 常见缩写
        r'[A-Z]{2,6}\d{1,2}[A-Z]?'             # 股票代码类
        r')$', re.I
    )
    for m in _CHEM_FORMULA.finditer(text):
        raw = m.group(0)
        if re.search(r'[A-Za-z]', raw) and re.search(r'\d', raw) and not _NOT_CHEM.match(raw):
            c.mark(m.start(), m.end())

    # P205 写法
    for m in _P205_PAT.finditer(text):
        c.mark(m.start(), m.end())

    # 化学名称中的数字词（五氧化二磷、二水法 等）
    for m in _CN_CHEM_SUFFIX.finditer(text):
        c.mark(m.start(), m.end())

    # 英文序数词连字符复合词（second-tier 等），但排除分数词（two-thirds）
    for m in _EN_ORDINAL_HYPHEN.finditer(text):
        c.mark(m.start(), m.end())
    for m in _EN_NUM_HYPHEN.finditer(text):
        suffix = m.group(0).split("-", 1)[1].lower()
        if suffix not in _FRAC_DENOM_SET:
            c.mark(m.start(), m.end())

    # 中文量词短语（一种、两类 等）
    for m in _CN_MEASURE_WORD.finditer(text):
        c.mark(m.start(), m.end())

    # 地名数字（四方地、七彩 等）
    for m in _CN_PLACE_NUM.finditer(text):
        c.mark(m.start(), m.end())

    # 英文描述性数字词（two types of、one of the main 等）
    for m in _EN_DESCRIPTIVE.finditer(text):
        c.mark(m.start(), m.end())

    # 中文固定词组（一般、一旦 等）
    for m in _CN_FIXED_PHRASES.finditer(text):
        c.mark(m.start(), m.end())

    # ══ 1. 单位声明行 ══════════════════════════════════════════════
    # 中文：单位：[量级]元
    for m in re.finditer(r"单位[：:]\s*(百万|千万|百亿|千亿|亿|千|百|万)?\s*(美元|欧元|元人民币|元|人民币)", text):
        mul = _CN_MUL.get(m.group(1) or "", 1)
        sym = _CN_CURRENCY.get(m.group(2), "RMB")
        add(tokens, m.start(), f"{sym} {_fmt(mul)}", m.end(), m.group(0), "unit_decl")
    # 英文：Unit: [USD/RMB/EUR]数字[million/thousand]
    for m in re.finditer(r"[Uu]nit[：:]\s*(USD|RMB|EUR)?\s*(\d[\d,]*)(?:\s*(million|billion|thousand))?", text):
        sym = (m.group(1) or "RMB").upper()
        num = float(_sc(m.group(2)))
        mul = _EN_MUL.get((m.group(3) or "").lower(), 1)
        add(tokens, m.start(), f"{sym} {_fmt(num*mul)}", m.end(), m.group(0), "unit_decl")

    # ══ 1.5 地址编号预消费（No. X / #X / X# / X层/F）══════════════
    # 这类编号两侧对等，整体标记避免重复计数
    # "No. 1 Office Building" 中的 1 只算一次
    for m in re.finditer(r"\bNo\.\s*(\d+)(?:\s+\w+)*\b", text, re.I):
        # 只提取编号数字，整个 "No. 1 Office Building" 不重复计数
        add(tokens, m.start(), m.group(1), m.end(), m.group(0), "addr_no")
    for m in re.finditer(r"(\d+)[#＃]", text):
        add(tokens, m.start(), m.group(1), m.end(), m.group(0), "addr_no")
    for m in re.finditer(r"(\d+)[/／]F\b", text, re.I):
        add(tokens, m.start(), m.group(1), m.end(), m.group(0), "addr_floor")
    for m in re.finditer(r"(\d+)\s*层\b", text):
        add(tokens, m.start(), m.group(1), m.end(), m.group(0), "addr_floor")
    # 第壹城、第贰期 等（中文大写数字+城/期/栋）
    for m in re.finditer(r"第([壹贰叁肆伍陆柒捌玖拾])[城期栋号]", text):
        v = {"壹":1,"贰":2,"叁":3,"肆":4,"伍":5,"陆":6,"柒":7,"捌":8,"玖":9,"拾":10}.get(m.group(1), 0)
        if v: add(tokens, m.start(), str(v), m.end(), m.group(0), "addr_no")

    # ══ 2. 百分点 ══════════════════════════════════════════════════
    for m in re.finditer(r"(\d[\d,]*(?:\.\d+)?)\s*percentage\s+points?", text, re.I):
        add(tokens, m.start(), _fmt(_sc(m.group(1))), m.end(), m.group(0), "pct_point")
    for m in re.finditer(r"(\d[\d,]*(?:\.\d+)?)\s*个百分点", text):
        add(tokens, m.start(), _fmt(_sc(m.group(1))), m.end(), m.group(0), "pct_point")

    # ══ 3. 百分比 ══════════════════════════════════════════════════
    for m in re.finditer(r"(\d[\d,]*(?:\.\d+)?)\s*(?:%|percent\b)", text, re.I):
        add(tokens, m.start(), _fmt(_sc(m.group(1))), m.end(), m.group(0), "percent")

    # ══ 4. 货币 ════════════════════════════════════════════════════
    # 英文：USD/RMB/EUR + 数字 + 可选量级
    for m in re.finditer(r"\b(USD|RMB|EUR)\s*(\d[\d,]*(?:\.\d+)?)(?:\s*(million|billion|thousand))?\b", text, re.I):
        num = float(_sc(m.group(2)))
        mul = _EN_MUL.get((m.group(3) or "").lower(), 1)
        add(tokens, m.start(), f"{m.group(1).upper()} {_fmt(num*mul)}", m.end(), m.group(0), "currency")
    # 中文货币：数字 + 可选量级 + 货币词
    for m in re.finditer(r"(\d[\d,]*(?:\.\d+)?)\s*(百万|千万|百亿|千亿|万|亿)?\s*(美元|欧元|元人民币|元|人民币)", text):
        num = float(_sc(m.group(1)))
        mul = _CN_MUL.get(m.group(2) or "", 1)
        sym = _CN_CURRENCY.get(m.group(3), "RMB")
        add(tokens, m.start(), f"{sym} {_fmt(num*mul)}", m.end(), m.group(0), "currency")

    # ══ 5. 经纬度 ══════════════════════════════════════════════════
    for m in re.finditer(r"(\d{1,3})°(\d{1,2})[′'](\d{1,2}(?:\.\d+)?)[″\"]{1,2}\s*([NSEW])\b", text):
        add(tokens, m.start(), f"{m.group(1)}°{m.group(2)}'{m.group(3)}'' {m.group(4).upper()}", m.end(), m.group(0), "coord")
    for m in re.finditer(r"(北纬|南纬|东经|西经)\s*(\d{1,3})°(\d{1,2})[′'′](\d{1,2}(?:\.\d+)?)[″\"″]", text):
        d = _DIR_MAP.get(m.group(1), "?")
        add(tokens, m.start(), f"{m.group(2)}°{m.group(3)}'{m.group(4)}'' {d}", m.end(), m.group(0), "coord")

    # ══ 6. 完整日期 ════════════════════════════════════════════════
    # 中文年月（二〇二六年三月 → 2026-03）
    _cn_year_month = re.compile(
        r"([零〇一二三四五六七八九]{4})年"
        r"([一二三四五六七八九十]{1,2})月"
        r"(?:([一二三四五六七八九十]{1,2})日)?"
    )
    for m in _cn_year_month.finditer(text):
        year_str = m.group(1)
        # 逐字转换年份（二〇二六 → 2026）
        year = "".join(str(_CN_DIGIT.get(c, c)) for c in year_str)
        mon  = _cn_to_int(m.group(2))
        if not mon:
            continue
        if m.group(3):
            day = _cn_to_int(m.group(3))
            val = f"{year}-{str(mon).zfill(2)}-{str(day).zfill(2)}"
        else:
            val = f"{year}-{str(mon).zfill(2)}"
        add(tokens, m.start(), val, m.end(), m.group(0), "cn_date")
    # 数字年月（2026年3月）
    for m in re.finditer(r"(\d{4})年(\d{1,2})月(\d{1,2})日", text):
        add(tokens, m.start(), f"{m.group(1)}-{m.group(2).zfill(2)}-{m.group(3).zfill(2)}", m.end(), m.group(0), "date")
    for m in re.finditer(r"(\d{4})年(\d{1,2})月(?!\d)", text):
        add(tokens, m.start(), f"{m.group(1)}-{m.group(2).zfill(2)}", m.end(), m.group(0), "date")
    # Month DD, YYYY
    _mon_pat = r"(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\.?"
    for m in re.finditer(_mon_pat + r"\s+(\d{1,2}),?\s+(\d{4})\b", text, re.I):
        mon = _MONTH.get(m.group(1).lower().rstrip("."), 0)
        if mon:
            add(tokens, m.start(), f"{m.group(3)}-{str(mon).zfill(2)}-{m.group(2).zfill(2)}", m.end(), m.group(0), "date")
    # DD Month YYYY
    for m in re.finditer(r"\b(\d{1,2})\s+" + _mon_pat + r"\s+(\d{4})\b", text, re.I):
        mon = _MONTH.get(m.group(2).lower().rstrip("."), 0)
        if mon:
            add(tokens, m.start(), f"{m.group(3)}-{str(mon).zfill(2)}-{m.group(1).zfill(2)}", m.end(), m.group(0), "date")

    # ══ 7. 年份范围 ════════════════════════════════════════════════
    for m in re.finditer(r"\b(\d{4})\s*[–—\-]\s*(\d{4})(?:年)?\b", text):
        add(tokens, m.start(), f"{m.group(1)}-{m.group(2)}", m.end(), m.group(0), "year_range")

    # ══ 8. 年代 ════════════════════════════════════════════════════
    for m in re.finditer(r"\b(?:the\s+)?(\d{4})s\b", text, re.I):
        add(tokens, m.start(), f"{m.group(1)}s", m.end(), m.group(0), "decade")
    for m in re.finditer(r"(\d{2})世纪(\d{2})年代", text):
        year = (int(m.group(1))-1)*100 + int(m.group(2))
        add(tokens, m.start(), f"{year}s", m.end(), m.group(0), "decade")

    # ══ 9. 世纪 ════════════════════════════════════════════════════
    for m in re.finditer(r"\b(?:the\s+)?(\d{1,2})(?:st|nd|rd|th)\s+century\b", text, re.I):
        add(tokens, m.start(), f"{m.group(1)}th century", m.end(), m.group(0), "century")
    for m in re.finditer(r"(\d{1,2})世纪(?![\d年])", text):
        add(tokens, m.start(), f"{m.group(1)}th century", m.end(), m.group(0), "century")

    # ══ 10. 公元前/后 ══════════════════════════════════════════════
    for m in re.finditer(r"\b(\d+)\s*BC\b", text, re.I):
        add(tokens, m.start(), f"{m.group(1)} BC", m.end(), m.group(0), "bc_ad")
    for m in re.finditer(r"\bAD\s*(\d+)\b", text, re.I):
        add(tokens, m.start(), f"AD {m.group(1)}", m.end(), m.group(0), "bc_ad")
    for m in re.finditer(r"公元前(\d+)年", text):
        add(tokens, m.start(), f"{m.group(1)} BC", m.end(), m.group(0), "bc_ad")
    for m in re.finditer(r"公元(\d+)年", text):
        add(tokens, m.start(), f"AD {m.group(1)}", m.end(), m.group(0), "bc_ad")

    # ══ 11. 财年 ═══════════════════════════════════════════════════
    for m in re.finditer(r"\bfiscal\s+year\s+(\d{4})\b", text, re.I):
        add(tokens, m.start(), f"FY{m.group(1)}", m.end(), m.group(0), "fiscal")
    for m in re.finditer(r"(\d{4})\s*财年", text):
        add(tokens, m.start(), f"FY{m.group(1)}", m.end(), m.group(0), "fiscal")

    # ══ 12. 季度 ═══════════════════════════════════════════════════
    for m in re.finditer(r"\bQ([1-4])\b", text):
        add(tokens, m.start(), f"Q{m.group(1)}", m.end(), m.group(0), "quarter")
    _qmap = {"first":1,"second":2,"third":3,"fourth":4}
    for m in re.finditer(r"\bthe\s+(first|second|third|fourth)\s+quarter\b", text, re.I):
        add(tokens, m.start(), f"Q{_qmap[m.group(1).lower()]}", m.end(), m.group(0), "quarter")
    for m in re.finditer(r"第?([一二三四])\s*季度", text):
        n = _CN_QUARTER.get(m.group(1), 0)
        if n: add(tokens, m.start(), f"Q{n}", m.end(), m.group(0), "quarter")

    # ══ 13. 期数 ═══════════════════════════════════════════════════
    for m in re.finditer(r"\bPhase\s+(I{1,3}|IV|VI{0,3}|IX|X{0,3}|[1-9])\b", text):
        raw = m.group(1)
        n = int(raw) if raw.isdigit() else _roman_to_int(raw)
        if n: add(tokens, m.start(), f"Phase {n}", m.end(), m.group(0), "phase")
    for m in re.finditer(r"([一二三四五六七八九十])\s*期(?!间|末|初|内|望)", text):
        n = _CN_SIMPLE.get(m.group(1), 0)
        if n: add(tokens, m.start(), f"Phase {n}", m.end(), m.group(0), "phase")

    # ══ 14. X+Y 结构 ═══════════════════════════════════════════════
    for m in re.finditer(r'"?(\d+)\s*\+\s*(\d+)"?', text):
        add(tokens, m.start(), f"{m.group(1)}+{m.group(2)}", m.end(), m.group(0), "plus")

    # ══ 15. 下标化学式 ═════════════════════════════════════════════
    for m in re.finditer(r"\b(PM|CO|NO|SO|H|O|N|C|Fe|Ca|Na|K)(\d+(?:\.\d+)?)\b", text, re.I):
        add(tokens, m.start(), f"{m.group(1)}{m.group(2)}", m.end(), m.group(0), "subscript")

    # ══ 16. 中文大写金额 ═══════════════════════════════════════════
    # 单个大写数字（壹/贰/叁...）只在后跟量级词或货币词时才提取，
    # 避免"陆上"、"贰心"等普通汉字被误识别为数字。
    # 多字符串（壹佰叁拾元）直接提取。
    _upper_pat = re.compile(
        r"[壹贰叁肆伍陆柒捌玖拾佰仟万亿零]+"
        r"(?:元(?:[零壹贰叁肆伍陆柒捌玖拾佰仟万亿]*))?"
        r"(?:角[零壹贰叁肆伍陆柒捌玖]*)?"
        r"(?:分[零壹贰叁肆伍陆柒捌玖]*)?"
        r"(?:厘[零壹贰叁肆伍陆柒捌玖]*)?"
    )
    _UPPER_SCALE = frozenset("拾佰仟万亿元角分厘")
    for m in _upper_pat.finditer(text):
        raw = m.group(0)
        # 单字大写数字：后一个字必须是量级/货币词，否则是普通汉字
        if len(raw) == 1 and raw not in _UPPER_SCALE:
            after_char = text[m.end():m.end()+1]
            if after_char not in _UPPER_SCALE:
                continue
        val = _cn_upper_to_float(raw)
        if val is not None:
            add(tokens, m.start(), _fmt(val), m.end(), raw, "cn_upper")

    # ══ 17. 中文万/亿复合单位 ══════════════════════════════════════
    _cn_units = {"微克/立方米":"μg/m3","微克/平方米":"μg/m2","微克":"μg","毫克":"mg",
                 "千克":"kg","平方千米":"square kilometers","立方米":"cubic meters",
                 "千米":"kilometers","公里":"kilometers","千瓦时":"kWh"}
    _cu_pat = "|".join(re.escape(u) for u in sorted(_cn_units, key=len, reverse=True))
    # 数字 + 量级 + 中文单位
    for m in re.finditer(r"(\d[\d,]*(?:\.\d+)?)\s*(百万|千万|百亿|千亿|[万亿])\s*(" + _cu_pat + r")", text):
        num = float(_sc(m.group(1)))
        mul = _CN_MUL.get(m.group(2), 1)
        unit = _cn_units.get(m.group(3), m.group(3))
        add(tokens, m.start(), f"{_fmt(num*mul)} {unit}", m.end(), m.group(0), "cn_unit")
    # 数字 + 中文单位（无量级）
    for m in re.finditer(r"(\d[\d,]*(?:\.\d+)?)\s*(" + _cu_pat + r")", text):
        unit = _cn_units.get(m.group(2), m.group(2))
        add(tokens, m.start(), f"{_sc(m.group(1))} {unit}", m.end(), m.group(0), "cn_unit")
    # 数字 + 百万/千万/百亿/千亿（无单位）
    for m in re.finditer(r"(\d[\d,]*(?:\.\d+)?)\s*(百万|千万|百亿|千亿)", text):
        mul = _CN_MUL[m.group(2)]
        add(tokens, m.start(), _fmt(float(_sc(m.group(1)))*mul), m.end(), m.group(0), "cn_large")
    # 数字 + 万/亿（排除后跟物理单位词）
    for m in re.finditer(r"(\d[\d,]*(?:\.\d+)?)\s*([万亿])(?![米克瓦升焦帕牛安伏欧赫兹])", text):
        mul = _CN_LARGE.get(m.group(2), 1)
        add(tokens, m.start(), _fmt(float(_sc(m.group(1)))*mul), m.end(), m.group(0), "cn_large")

    # ══ 18. 英文单位 ═══════════════════════════════════════════════
    _en_units = r"(μg/m[²2³3]?|mg/m[²2³3]?|μg|mg|kg(?!\w)|square\s+kilometers?|cubic\s+meters?|kilometers?|metres?|meters?|m²|m[²2]|m[³3]|kWh)"
    # 带量级
    for m in re.finditer(r"(\d[\d,]*(?:\.\d+)?)\s*(million|billion|thousand)\s+" + _en_units, text, re.I):
        num = float(_sc(m.group(1)))
        mul = _EN_MUL.get(m.group(2).lower(), 1)
        add(tokens, m.start(), f"{_fmt(num*mul)} {m.group(3).strip()}", m.end(), m.group(0), "en_unit")
    # 普通
    for m in re.finditer(r"(\d[\d,]*(?:\.\d+)?)\s*" + _en_units, text, re.I):
        add(tokens, m.start(), f"{_sc(m.group(1))} {m.group(2).strip()}", m.end(), m.group(0), "en_unit")

    # ══ 19. 千分位数字 ═════════════════════════════════════════════
    for m in re.finditer(r"\d{1,3}(?:,\d{3})+(?:\.\d+)?", text):
        add(tokens, m.start(), _fmt(float(_sc(m.group(0)))), m.end(), m.group(0), "thousand")

    # ══ 20. 下标数字（PM2.5 已在15处理，这里兜底）══════════════════
    # 已由 subscript 覆盖，跳过

    # ══ 21. 小数 ═══════════════════════════════════════════════════
    for m in re.finditer(r"(?<!\d)\d+\.\d+(?!\d)", text):
        add(tokens, m.start(), m.group(0), m.end(), m.group(0), "decimal")

    # ══ 22. 大数（5位+）════════════════════════════════════════════
    for m in re.finditer(r"(?<!\d)\d{5,}(?!\d)", text):
        add(tokens, m.start(), m.group(0), m.end(), m.group(0), "bignum")

    # ══ 23. 罗马数字 ═══════════════════════════════════════════════
    _roman_pat = re.compile(r"\b(M{0,4}(?:CM|CD|D?C{0,3})(?:XC|XL|L?X{0,3})(?:IX|IV|V?I{0,3}))\b")
    _single_upper = set("IVXLCDM")
    def _roman_ctx_ok(m, txt):
        before = txt[:m.start()]
        after  = txt[m.end():]
        pre_ok  = (not before) or before[-1] in "( \t\n"
        post_ok = (not after)  or after[0] in ").," or (after and after[0]==" " and len(after)>1 and after[1].isupper())
        return pre_ok and post_ok
    for m in _roman_pat.finditer(text):
        s = m.group(1)
        if not s: continue
        if len(s)==1 and s in _single_upper and not _roman_ctx_ok(m, text): continue
        v = _roman_to_int(s)
        if v > 0: add(tokens, m.start(), str(v), m.end(), m.group(0), "roman")
    for m in re.finditer("[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫⅰⅱⅲⅳⅴⅵⅶⅷⅸⅹⅺⅻ]", text):
        v = _ROMAN_FW.get(m.group(0), 0)
        if v: add(tokens, m.start(), str(v), m.end(), m.group(0), "roman_fw")

    # ══ 23.5 分数词（必须在英文数字词之前，避免 Two 被先消费）══════
    _fn = "|".join(_FRAC_NUM.keys())
    _fd = "|".join(sorted(_FRAC_DEN.keys(), key=len, reverse=True))
    for m in re.finditer(r"\b(" + _fn + r")\s*[-\s]+(" + _fd + r")\b", text, re.I):
        n = _FRAC_NUM.get(m.group(1).lower(), 1)
        d = _FRAC_DEN.get(m.group(2).lower(), 1)
        add(tokens, m.start(), f"{n}/{d}", m.end(), m.group(0), "frac_word")

    # ══ 24. 英文数字词（one/two/three...）══════════════════════════
    # 先处理复合（twenty-one 等）
    _all_base = "|".join(sorted(_EN_NUM.keys(), key=len, reverse=True))
    _all_scale = "|".join(sorted(_EN_SCALE.keys(), key=len, reverse=True))
    for m in re.finditer(
        r"\b((?:(?:" + _all_base + r")\s*(?:" + _all_scale + r")?\s*)+)\b", text, re.I
    ):
        raw = m.group(0).strip().lower()
        words = re.split(r"[\s\-]+", raw)
        total, cur = 0, 0
        for w in words:
            if w in _EN_SCALE:
                s = _EN_SCALE[w]
                if s >= 1000:
                    total = (total + (cur or 1)) * s; cur = 0
                else:
                    cur = (cur or 1) * s
            elif w in _EN_NUM:
                cur += _EN_NUM[w]
        val = total + cur
        if val > 0:
            add(tokens, m.start(), str(val), m.end(), m.group(0), "en_word")

    # ══ 25. 序数词 ═════════════════════════════════════════════════
    _ord_words = "|".join(sorted(_EN_ORDINAL.keys(), key=len, reverse=True))
    for m in re.finditer(r"\b(" + _ord_words + r")\b", text, re.I):
        v = _EN_ORDINAL.get(m.group(1).lower(), 0)
        if v: add(tokens, m.start(), str(v), m.end(), m.group(0), "ordinal")
    for m in re.finditer(r"\b(\d+)(?:st|nd|rd|th)\b", text, re.I):
        add(tokens, m.start(), m.group(1), m.end(), m.group(0), "ordinal_num")

    # ══ 26. 分数词（数字分数 + 中文分数）══════════════════════════
    # 数字分数 N/N
    for m in re.finditer(r"(?<!\d)(\d+)/(\d+)(?!\d)", text):
        add(tokens, m.start(), f"{m.group(1)}/{m.group(2)}", m.end(), m.group(0), "frac_num")
    # 中文分数：三分之二
    for m in re.finditer(r"([一二三四五六七八九十百]+)分之([一二三四五六七八九十百]+)", text):
        d = _cn_to_int(m.group(1)); n = _cn_to_int(m.group(2))
        if d and n: add(tokens, m.start(), f"{n}/{d}", m.end(), m.group(0), "frac_cn")
    for m in re.finditer(r"一半", text):
        add(tokens, m.start(), "1/2", m.end(), m.group(0), "frac_cn")

    # ══ 27. 月份名 ═════════════════════════════════════════════════
    _mon_names = r"\b(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\.?\b"
    # 月份 + 年份（March 2026 → 2026-03）
    for m in re.finditer(_mon_names + r"\s+(\d{4})\b", text, re.I):
        k = m.group(1).lower().rstrip(".")
        v = _MONTH.get(k, 0)
        if v:
            add(tokens, m.start(), f"{m.group(2)}-{str(v).zfill(2)}", m.end(), m.group(0), "month_year")
    # 单独月份名
    for m in re.finditer(_mon_names, text, re.I):
        k = m.group(1).lower().rstrip(".")
        v = _MONTH.get(k, 0)
        if v: add(tokens, m.start(), str(v), m.end(), m.group(0), "month")

    # ══ 28. 中文数字串 ═════════════════════════════════════════════
    _cn_num_pat = re.compile(r"[零〇一二三四五六七八九十百千万亿壹贰叁肆伍陆柒捌玖拾佰仟]+")
    for m in _cn_num_pat.finditer(text):
        v = _cn_to_int(m.group(0))
        if v: add(tokens, m.start(), str(v), m.end(), m.group(0), "cn_num")

    # ══ 29. 普通整数（1-4位）══════════════════════════════════════
    for m in re.finditer(r"(?<!\d)\d{1,4}(?!\d)", text):
        add(tokens, m.start(), m.group(0), m.end(), m.group(0), "int")

    tokens.sort(key=lambda t: t.start)
    return tokens


def parse_values(text: str) -> List[str]:
    """只返回规范化数值字符串列表"""
    return [t.value for t in extract(text)]
