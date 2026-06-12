"""
alignment_matrix.py — 对齐与校验矩阵
======================================
两侧数值列表清洗完毕后，建立对齐矩阵进行比对。

匹配策略（按优先级）：
  1. 精确字符串匹配
  2. 数值等价（1,000 == 1000，1.0 == 1）
  3. 货币 vs 纯数字跨类型等价（RMB 20000000000 == 20000000000）
  4. 未匹配 → MISSING / EXTRA
  5. 二次消除：MISSING+EXTRA 配对后再做等价比较，消除跨类型误报
"""

import re
from dataclasses import dataclass
from typing import List, Optional


@dataclass
class AlignError:
    error_type: str           # "MISSING" | "EXTRA" | "MISMATCH"
    src_value:  Optional[str]
    tgt_value:  Optional[str]
    message:    str
    position:   int = -1


# ─────────────────────────────────────────
# 规范化 & 等价判断
# ─────────────────────────────────────────

def _strip_currency(v: str) -> str:
    """去掉货币符号前缀，返回纯数字字符串"""
    m = re.match(r'^(?:USD|RMB|EUR)\s+(.+)$', v, re.I)
    return m.group(1) if m else v


def _to_float(s: str) -> Optional[float]:
    try:
        return float(s.replace(",", ""))
    except (ValueError, TypeError):
        return None


def _normalize(v: str) -> str:
    """规范化数值字符串，用于等价比较"""
    v = v.strip()
    # 特殊标记直接返回（大小写统一）
    if re.match(r'^(Q[1-4]|Phase \d+|FY\d{4}|\d{4}s|\d+th century|AD \d+|\d+ BC|\d+/\d+|\d+\+\d+)$', v, re.I):
        return v
    # 货币：保留符号，规范化数字
    m = re.match(r'^(USD|RMB|EUR)\s+(.+)$', v, re.I)
    if m:
        num = _normalize_num(m.group(2))
        return f"{m.group(1).upper()} {num}"
    return _normalize_num(v)


def _normalize_num(s: str) -> str:
    f = _to_float(s)
    if f is None:
        return s
    i = round(f)
    if abs(f - i) <= abs(f) * 1e-9 + 1e-9:
        return str(i)
    return str(round(f, 10)).rstrip("0").rstrip(".")


def _values_equal(a: str, b: str) -> bool:
    """判断两个规范化值是否等价（含跨货币/纯数字比较）"""
    if a == b or a.upper() == b.upper():
        return True
    # 跨类型：一侧有货币符号，另一侧是纯数字
    fa = _to_float(_strip_currency(a))
    fb = _to_float(_strip_currency(b))
    if fa is not None and fb is not None:
        return abs(fa - fb) <= max(abs(fa), abs(fb)) * 1e-9 + 1e-9
    return False


# ─────────────────────────────────────────
# 主对齐函数
# ─────────────────────────────────────────

def build_matrix(src_values: List[str], tgt_values: List[str]) -> List[AlignError]:
    src_norm = [_normalize(v) for v in src_values]
    tgt_norm = [_normalize(v) for v in tgt_values]

    src_used = [False] * len(src_norm)
    tgt_used = [False] * len(tgt_norm)

    # ── 预处理：年月等价（2026-03 vs 2026 + 3）──
    def _is_year_month(v: str):
        m = re.match(r'^(\d{4})-(\d{2})$', v)
        return (m.group(1), str(int(m.group(2)))) if m else None

    for j, tn in enumerate(tgt_norm):
        ym = _is_year_month(tn)
        if ym and not tgt_used[j]:
            y, mo = ym
            i1 = next((i for i, s in enumerate(src_norm) if s == y and not src_used[i]), None)
            i2 = next((i for i, s in enumerate(src_norm) if s == mo and not src_used[i]), None)
            if i1 is not None and i2 is not None:
                src_used[i1] = src_used[i2] = tgt_used[j] = True

    for i, sn in enumerate(src_norm):
        ym = _is_year_month(sn)
        if ym and not src_used[i]:
            y, mo = ym
            j1 = next((j for j, t in enumerate(tgt_norm) if t == y and not tgt_used[j]), None)
            j2 = next((j for j, t in enumerate(tgt_norm) if t == mo and not tgt_used[j]), None)
            if j1 is not None and j2 is not None:
                tgt_used[j1] = tgt_used[j2] = src_used[i] = True

    # ── 预处理：年份范围 vs 两个独立年份的等价消除 ──
    # 例：src=['2020','2024'] tgt=['2020-2024'] → 视为等价，全部标记已用
    def _is_year_range(v: str):
        m = re.match(r'^(\d{4})-(\d{4})$', v)
        return (m.group(1), m.group(2)) if m else None

    for j, tn in enumerate(tgt_norm):
        yr = _is_year_range(tn)
        if yr and not tgt_used[j]:
            y1, y2 = yr
            i1 = next((i for i, s in enumerate(src_norm) if s == y1 and not src_used[i]), None)
            i2 = next((i for i, s in enumerate(src_norm) if s == y2 and not src_used[i]), None)
            if i1 is not None and i2 is not None:
                src_used[i1] = src_used[i2] = tgt_used[j] = True

    for i, sn in enumerate(src_norm):
        yr = _is_year_range(sn)
        if yr and not src_used[i]:
            y1, y2 = yr
            j1 = next((j for j, t in enumerate(tgt_norm) if t == y1 and not tgt_used[j]), None)
            j2 = next((j for j, t in enumerate(tgt_norm) if t == y2 and not tgt_used[j]), None)
            if j1 is not None and j2 is not None:
                tgt_used[j1] = tgt_used[j2] = src_used[i] = True

    # 贪心匹配
    for i, s in enumerate(src_norm):
        for j, t in enumerate(tgt_norm):
            if not tgt_used[j] and _values_equal(s, t):
                src_used[i] = True
                tgt_used[j] = True
                break

    missing = [
        AlignError("MISSING", src_values[i], None,
                   f"译文缺失数值 [{src_values[i]}]（原文存在）", i)
        for i, used in enumerate(src_used) if not used
    ]
    extra = [
        AlignError("EXTRA", None, tgt_values[j],
                   f"译文多出数值 [{tgt_values[j]}]（原文不存在）", j)
        for j, used in enumerate(tgt_used) if not used
    ]

    # ── 二次消除：MISSING+EXTRA 配对等价比较 ──
    # 处理货币符号有无、单位差异等导致的跨类型误报
    cancelled_m, cancelled_e = set(), set()
    for mi, me in enumerate(missing):
        for ei, ee in enumerate(extra):
            if mi in cancelled_m or ei in cancelled_e:
                continue
            if _values_equal(_normalize(me.src_value), _normalize(ee.tgt_value)):
                cancelled_m.add(mi)
                cancelled_e.add(ei)

    errors = (
        [e for i, e in enumerate(missing) if i not in cancelled_m] +
        [e for i, e in enumerate(extra)   if i not in cancelled_e]
    )

    # MISMATCH：数量相同且无 MISSING/EXTRA 但顺序不同
    if len(src_values) == len(tgt_values) and not errors:
        for i in range(len(src_norm)):
            if not _values_equal(src_norm[i], tgt_norm[i]):
                errors.append(AlignError(
                    "MISMATCH", src_values[i], tgt_values[i],
                    f"数值不一致：原文 [{src_values[i]}] vs 译文 [{tgt_values[i]}]", i
                ))

    return errors


def format_errors(errors: List[AlignError], src: str = "", tgt: str = "") -> str:
    if not errors:
        return "✔ 数值校验通过"
    lines = [f"❗ 发现 {len(errors)} 处数值不一致："]
    for i, e in enumerate(errors, 1):
        lines.append(f"  [{i}] {e.error_type}: {e.message}")
    if src:
        lines.append(f"  原文: {src[:100]}{'...' if len(src)>100 else ''}")
    if tgt:
        lines.append(f"  译文: {tgt[:100]}{'...' if len(tgt)>100 else ''}")
    return "\n".join(lines)
