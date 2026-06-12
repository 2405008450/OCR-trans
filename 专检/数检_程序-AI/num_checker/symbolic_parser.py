"""
symbolic_parser.py — 模块A：符号化解析器（公共接口）
=====================================================
自研 FSM 扫描器，不依赖 normalizer_total。
具体规则实现在 _parser_core.py。
"""

import re
from dataclasses import dataclass
from typing import List, Tuple, Optional


# ─────────────────────────────────────────
# 数据结构
# ─────────────────────────────────────────

@dataclass
class Token:
    value: str    # 规范化数值字符串
    raw:   str    # 原始文本
    start: int
    end:   int
    tag:   str = ""  # 类型标签，便于调试


# ─────────────────────────────────────────
# 映射表（供 _parser_core 导入）
# ─────────────────────────────────────────

_CN_DIGIT = {
    "零":0,"〇":0,"一":1,"二":2,"三":3,"四":4,"五":5,
    "六":6,"七":7,"八":8,"九":9,
    "壹":1,"贰":2,"叁":3,"肆":4,"伍":5,
    "陆":6,"柒":7,"捌":8,"玖":9,
}
_CN_UNIT  = {"十":10,"拾":10,"百":100,"佰":100,"千":1000,"仟":1000,"万":10000,"亿":100000000}
_CN_LARGE = {"万":10000,"亿":100000000,"百":100,"千":1000}
_CN_SIMPLE= {"一":1,"二":2,"三":3,"四":4,"五":5,"六":6,"七":7,"八":8,"九":9,"十":10}
_CN_QUARTER={"一":1,"二":2,"三":3,"四":4}

_EN_NUM = {
    "zero":0,"one":1,"two":2,"three":3,"four":4,"five":5,"six":6,"seven":7,
    "eight":8,"nine":9,"ten":10,"eleven":11,"twelve":12,"thirteen":13,
    "fourteen":14,"fifteen":15,"sixteen":16,"seventeen":17,"eighteen":18,
    "nineteen":19,"twenty":20,"thirty":30,"forty":40,"fifty":50,
    "sixty":60,"seventy":70,"eighty":80,"ninety":90,
}
_EN_SCALE  = {"hundred":100,"thousand":1000,"million":1000000,"billion":1000000000}
_EN_ORDINAL= {
    "first":1,"second":2,"third":3,"fourth":4,"fifth":5,"sixth":6,"seventh":7,
    "eighth":8,"ninth":9,"tenth":10,"eleventh":11,"twelfth":12,"thirteenth":13,
    "fourteenth":14,"fifteenth":15,"sixteenth":16,"seventeenth":17,"eighteenth":18,
    "nineteenth":19,"twentieth":20,"thirtieth":30,"fortieth":40,"fiftieth":50,
    "sixtieth":60,"seventieth":70,"eightieth":80,"ninetieth":90,"hundredth":100,
}
_FRAC_NUM = {"one":1,"two":2,"three":3,"four":4,"five":5,"six":6,"seven":7,"eight":8,"nine":9,"ten":10}
_FRAC_DEN = {
    "half":2,"halves":2,"third":3,"thirds":3,"fourth":4,"fourths":4,
    "quarter":4,"quarters":4,"fifth":5,"fifths":5,"sixth":6,"sixths":6,
    "seventh":7,"sevenths":7,"eighth":8,"eighths":8,"ninth":9,"ninths":9,
    "tenth":10,"tenths":10,
}
_MONTH = {
    "january":1,"jan":1,"february":2,"feb":2,"march":3,"mar":3,
    "april":4,"apr":4,"may":5,"june":6,"jun":6,"july":7,"jul":7,
    "august":8,"aug":8,"september":9,"sep":9,"sept":9,
    "october":10,"oct":10,"november":11,"nov":11,"december":12,"dec":12,
}
_ROMAN_VAL = {"I":1,"V":5,"X":10,"L":50,"C":100,"D":500,"M":1000}
_ROMAN_FW  = {
    "Ⅰ":1,"Ⅱ":2,"Ⅲ":3,"Ⅳ":4,"Ⅴ":5,"Ⅵ":6,"Ⅶ":7,"Ⅷ":8,"Ⅸ":9,"Ⅹ":10,"Ⅺ":11,"Ⅻ":12,
    "ⅰ":1,"ⅱ":2,"ⅲ":3,"ⅳ":4,"ⅴ":5,"ⅵ":6,"ⅶ":7,"ⅷ":8,"ⅸ":9,"ⅹ":10,"ⅺ":11,"ⅻ":12,
}
_DIR_MAP     = {"北纬":"N","南纬":"S","东经":"E","西经":"W"}
_CN_CURRENCY = {"美元":"USD","欧元":"EUR","元人民币":"RMB","元":"RMB","人民币":"RMB"}
_CN_MUL = {"百万":1e6,"千万":1e7,"百亿":1e10,"千亿":1e11,"亿":1e8,"万":1e4,"千":1e3,"百":1e2}
_EN_MUL = {"million":1e6,"billion":1e9,"thousand":1e3}


# ─────────────────────────────────────────
# 工具函数（供 _parser_core 导入）
# ─────────────────────────────────────────

def _fmt(n) -> str:
    try:
        f = float(n)
        i = round(f)
        if abs(f - i) <= abs(f) * 1e-9 + 1e-9:
            return str(i)
        return str(round(f, 10)).rstrip("0").rstrip(".")
    except Exception:
        return str(n)

def _sc(s: str) -> str:
    return s.replace(",", "")

def _roman_to_int(s: str) -> int:
    total, prev = 0, 0
    for c in reversed(s.upper()):
        v = _ROMAN_VAL.get(c, 0)
        total += v if v >= prev else -v
        prev = v
    return total

def _cn_to_int(s: str) -> Optional[int]:
    total, cur = 0, 0
    for c in s:
        if c in _CN_DIGIT:
            cur = _CN_DIGIT[c]
        elif c in _CN_UNIT:
            u = _CN_UNIT[c]
            if u >= 10000:
                total = (total + cur) * u; cur = 0
            else:
                total += (cur if cur else 1) * u; cur = 0
    return (total + cur) or None

def _cn_upper_to_float(s: str) -> Optional[float]:
    total, cur, frac, in_frac = 0.0, 0, 0.0, False
    for c in s:
        if c in _CN_DIGIT: cur = _CN_DIGIT[c]
        elif c == "拾": total += (cur or 1)*10; cur=0
        elif c == "佰": total += (cur or 1)*100; cur=0
        elif c == "仟": total += (cur or 1)*1000; cur=0
        elif c == "万": total = (total+cur)*10000; cur=0
        elif c == "亿": total = (total+cur)*100000000; cur=0
        elif c == "元": total += cur; cur=0; in_frac=True
        elif c == "角" and in_frac: frac += cur*0.1; cur=0
        elif c == "分" and in_frac: frac += cur*0.01; cur=0
        elif c == "厘" and in_frac: frac += cur*0.001; cur=0
        elif c == "零": total += cur if not in_frac else 0; cur=0
    if not in_frac: total += cur
    r = round(total + frac, 4)
    return r if r > 0 else None


# ─────────────────────────────────────────
# 消费区间管理
# ─────────────────────────────────────────

class _Consumed:
    def __init__(self):
        self._spans: List[Tuple[int,int]] = []

    def used(self, a: int, b: int) -> bool:
        return any(a < e and b > s for s, e in self._spans)

    def mark(self, a: int, b: int):
        self._spans.append((a, b))

    def add(self, tokens: List[Token], pos: int, val: str, end: int,
            raw: str, tag: str):
        if not self.used(pos, end):
            tokens.append(Token(value=val, raw=raw, start=pos, end=end, tag=tag))
            self.mark(pos, end)


# ─────────────────────────────────────────
# 公共接口（延迟导入避免循环）
# ─────────────────────────────────────────

def extract(text: str) -> List[Token]:
    """提取所有数值 Token（含位置信息）"""
    from ._parser_core import extract as _extract
    return _extract(text)


def parse_values(text: str) -> List[str]:
    """只返回规范化数值字符串列表"""
    from ._parser_core import parse_values as _pv
    return _pv(text)
