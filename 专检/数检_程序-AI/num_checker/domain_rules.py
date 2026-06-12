"""
domain_rules.py
===============
行业专属符号库 + 多义词黑名单。

可直接硬编码行业规则，RL判别器会将这里的规则视为"绝对真理"。
"""

# ─────────────────────────────────────────
# 行业专属数值映射（符号 → 标准数值字符串）
# 金融、医疗、工程等领域的特殊缩写
# ─────────────────────────────────────────
DOMAIN_SYMBOL_MAP: dict[str, str] = {
    # 金融 — 这里只放"替换原始数字"的符号，不放单位换算
    # basis point 是单位词，不映射为数值（50bp 的数值是 50，不是 0.01）
    "ppm":   "0.0001", # parts per million（极少数情况需要换算）
}

# ─────────────────────────────────────────
# 多义词黑名单
# 格式：{词: [(搭配词列表, 判定为"非数值"的理由), ...]}
# 搭配词：前后 N 个 token 内出现时，判定为非数值
# ─────────────────────────────────────────
AMBIGUOUS_WORDS: dict[str, list] = {
    # 动词搭配 → 非数值
    "double": [
        (["click", "clicked", "clicking", "tap", "tapped"], "固定搭配：double-click"),
        (["check", "checked", "blind", "standard"], "固定搭配：double-check/blind"),
        (["major", "role", "duty", "function", "purpose"], "比喻用法：double role"),
    ],
    "triple": [
        (["threat", "play", "jump", "axel"], "体育/比喻用法"),
    ],
    "single": [
        (["out", "minded", "handedly", "parent", "room", "bed"], "非数值用法"),
    ],
    "half": [
        (["hearted", "time", "way", "back", "life", "term"], "固定搭配非数值"),
    ],
    "quarter": [
        (["back", "final", "horse", "deck"], "非数值用法"),
    ],
    "one": [
        (["on", "sided", "way", "time", "off", "size", "stop"], "固定搭配非数值"),
    ],
    "two": [
        (["faced", "fold", "way", "sided", "step", "tone"], "固定搭配非数值"),
    ],
    "three": [
        (["dimensional", "fold", "way", "piece", "point"], "固定搭配非数值"),
    ],
    # 金融多义词
    "yield": [
        (["crop", "harvest", "produce", "result", "output", "sign", "give", "gave"], "非收益率用法"),
    ],
    "return": [
        (["home", "back", "trip", "flight", "ticket", "journey"], "非收益率用法"),
    ],
    "spread": [
        (["butter", "jam", "disease", "news", "wings", "arms"], "非利差用法"),
    ],
    "leverage": [
        (["advantage", "influence", "power", "position"], "比喻用法非财务杠杆"),
    ],
    "peer": [
        (["review", "pressure", "group", "learning", "support"], "同辈/同行非数值"),
    ],
}

# ─────────────────────────────────────────
# 数值型多义词映射（词 → 数值，仅在数值语境下生效）
# ─────────────────────────────────────────
NUMERIC_AMBIGUOUS: dict[str, float] = {
    "double":  2.0,
    "twice":   2.0,
    "triple":  3.0,
    "treble":  3.0,
    "quadruple": 4.0,
    "half":    0.5,
    "quarter": 0.25,
}

# ─────────────────────────────────────────
# 数值语境触发词（出现这些词时，多义词更可能是数值）
# ─────────────────────────────────────────
NUMERIC_CONTEXT_TRIGGERS: list[str] = [
    "increase", "decrease", "grow", "grew", "rise", "rose", "fall", "fell",
    "times", "fold", "rate", "ratio", "factor", "multiplier",
    "compared", "versus", "vs", "than", "more", "less",
    "revenue", "profit", "loss", "income", "cost", "price", "value",
    "percent", "%", "growth", "decline", "surge", "drop",
]
