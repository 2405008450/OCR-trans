"""
rl_discriminator.py — 模块B：RL 上下文判别器
==============================================
用确定性规则+权重表模拟 Q-learning 决策：

  State  = (当前词, 前N词上下文, 后N词上下文)
  Action = "数值" | "非数值"
  Reward = 语法/语义特征打分

黑名单搭配享有最高优先级（绝对真理，-10分直接否决）。
"""

import re
from typing import List, Tuple

from .domain_rules import AMBIGUOUS_WORDS, NUMERIC_AMBIGUOUS, NUMERIC_CONTEXT_TRIGGERS

# 决策阈值：Q值 > 此值才判定为数值
_THRESHOLD = 1.0


def is_numeric_word(word: str, context: str, window: int = 5) -> Tuple[bool, float, str]:
    """
    判断多义词 word 在 context 中是否表示数值。
    返回 (is_numeric, q_score, reason)
    """
    word_lower = word.lower()
    if word_lower not in NUMERIC_AMBIGUOUS:
        return True, 5.0, "非多义词"

    tokens = re.findall(r'\b\w+\b', context.lower())
    try:
        idx = next(i for i, t in enumerate(tokens) if t == word_lower)
    except StopIteration:
        return True, 0.5, "上下文未找到词，保守通过"

    before = tokens[max(0, idx-window): idx]
    after  = tokens[idx+1: idx+1+window]
    surr   = before + after
    score  = 0.0
    reasons = []

    # ── 黑名单（绝对真理）──────────────────────────────────────────
    if word_lower in AMBIGUOUS_WORDS:
        for collocations, reason in AMBIGUOUS_WORDS[word_lower]:
            if any(c in surr for c in collocations):
                return False, -10.0, f"黑名单: {reason}"

    # ── 正奖励 ─────────────────────────────────────────────────────
    if any(t in surr for t in NUMERIC_CONTEXT_TRIGGERS):
        score += 3.0; reasons.append("数值语境触发词")

    _units = ["percent","%","times","fold","million","billion","thousand",
              "万","亿","元","美元","欧元","kg","km","m²","m³","kWh"]
    if any(u in after for u in _units):
        score += 4.0; reasons.append("后跟单位")

    if any(re.match(r'^\d', t) for t in before):
        score += 3.0; reasons.append("前有数字")
    if any(re.match(r'^\d', t) for t in after):
        score += 3.0; reasons.append("后有数字")

    if len(tokens) <= 10:
        score += 2.0; reasons.append("短文本/表格")

    if any(c in surr for c in ["than","vs","versus","compared","relative","against"]):
        score += 2.5; reasons.append("比较语境")

    _finance = ["revenue","profit","loss","income","cost","price","return","yield",
                "rate","ratio","margin","growth","earnings","ebitda","eps","roe","roa"]
    if any(f in surr for f in _finance):
        score += 2.0; reasons.append("金融领域")

    # ── 负奖励 ─────────────────────────────────────────────────────
    _verbs = ["click","tap","press","hit","check","play","act","serve",
              "function","work","use","make","take","get","go","come"]
    if before and before[-1] in _verbs:
        score -= 3.0; reasons.append("前面是动词")

    if any(n in before for n in ["not","no","never","without","lack"]):
        score -= 2.0; reasons.append("否定语境")

    is_num = score >= _THRESHOLD
    return is_num, score, " | ".join(reasons) or "无特征"


def filter_ambiguous_tokens(text: str, values: List[str]) -> List[str]:
    """
    对已提取的数值列表做多义词过滤：
    - 若 RL 判定某多义词为非数值，从列表中移除其对应数值
    - 不主动补入新数值（避免误报）
    """
    result = list(values)
    for word, num_val in NUMERIC_AMBIGUOUS.items():
        if not re.search(r'\b' + re.escape(word) + r'\b', text, re.I):
            continue
        str_val = str(int(num_val)) if num_val == int(num_val) else str(num_val)
        is_num, _, _ = is_numeric_word(word, text)
        if not is_num and str_val in result:
            result.remove(str_val)
    return result
