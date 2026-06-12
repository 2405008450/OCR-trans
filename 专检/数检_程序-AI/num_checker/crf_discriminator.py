"""
crf_discriminator.py — CRF 序列标注判别器
==========================================
用条件随机场（CRF）替换静态权重表，实现真正从数据中学习的判别器。

架构设计：
  - 特征提取：滑动窗口 token 级特征（词形、词性、上下文、领域词汇）
  - 模型：线性链 CRF（sklearn-crfsuite），模型文件 < 1MB，无需 GPU
  - 训练数据：由 PPO 奖励信号自动生成（对齐矩阵作为环境反馈）
  - 黑名单：domain_rules 中的规则仍作为硬约束（绝对真理，不可被模型覆盖）

PPO 奖励信号设计：
  - 正奖励 +1：模型判定为"数值"，且对齐矩阵验证通过（两侧匹配）
  - 负奖励 -1：模型判定为"数值"，但对齐矩阵报 EXTRA（误报）
  - 负奖励 -1：模型判定为"非数值"，但对齐矩阵报 MISSING（漏报）
  - 零奖励  0：模型判定为"非数值"，对齐矩阵也无该值（正确过滤）

使用方式：
  # 首次使用（无模型文件）→ 自动回退到规则判别器
  discriminator = CRFDiscriminator()

  # 用标注数据训练
  discriminator.train(labeled_data)

  # 保存/加载模型
  discriminator.save("num_checker/crf_model.pkl")
  discriminator.load("num_checker/crf_model.pkl")

  # 判别
  is_num, score, reason = discriminator.predict(word, context)
"""

import re
import os
import pickle
from typing import List, Tuple, Dict, Optional

from .domain_rules import AMBIGUOUS_WORDS, NUMERIC_AMBIGUOUS, NUMERIC_CONTEXT_TRIGGERS

# ─────────────────────────────────────────
# 特征工程
# ─────────────────────────────────────────

# 领域词汇集（用于特征提取）
_FINANCE_WORDS = frozenset([
    "revenue","profit","loss","income","cost","price","return","yield",
    "rate","ratio","margin","growth","earnings","ebitda","eps","roe","roa",
    "dividend","equity","debt","asset","liability","cash","flow","capital",
    "interest","tax","depreciation","amortization","operating","net","gross",
])
_UNIT_WORDS = frozenset([
    "percent","%","times","fold","million","billion","thousand","trillion",
    "万","亿","元","美元","欧元","kg","km","m²","m³","kwh","ton","tonne",
    "meter","metre","liter","litre","gram","kilogram","hectare","acre",
])
_COMPARE_WORDS = frozenset([
    "than","vs","versus","compared","relative","against","over","above","below",
    "exceed","surpass","outperform","underperform","increase","decrease",
])
_VERB_BEFORE = frozenset([
    "click","tap","press","hit","check","play","act","serve","function",
    "work","use","make","take","get","go","come","do","did","does",
])
_NEGATION = frozenset(["not","no","never","without","lack","absent","neither","nor"])


def _tokenize(text: str) -> List[str]:
    """
    中英文混合 tokenization。
    - 英文：按词边界切分
    - 中文：逐字切分（每个汉字作为一个 token）
    - 数字：整体保留
    """
    tokens = []
    # 匹配：连续汉字 | 英文单词 | 数字串 | 单个标点
    for m in re.finditer(r'[\u4e00-\u9fff]|[A-Za-z]+|\d+(?:\.\d+)?|[^\s]', text):
        tokens.append(m.group(0).lower())
    return tokens


def _stem(word: str) -> str:
    """
    轻量词干化：处理常见英文变形，让 CRF 特征对词形变化鲁棒。
    不依赖 NLTK，只处理翻译数值场景中的高频变形。
    """
    w = word.lower()
    # 动词过去式/进行时 → 原形
    for suffix, replacement in [
        ("bled", "ble"), ("pled", "ple"), ("ied", "y"),
        ("led",  "le"),  ("ned",  "ne"),  ("red",  "re"),
        ("sed",  "se"),  ("ted",  "te"),  ("ved",  "ve"),
        ("ling", "le"),  ("pling","ple"), ("bling","ble"),
        ("ing",  ""),    ("ed",   ""),    ("s",    ""),
    ]:
        if w.endswith(suffix) and len(w) - len(suffix) >= 3:
            return w[:-len(suffix)] + replacement
    return w


def _word_features(tokens: List[str], idx: int, window: int = 3) -> Dict[str, any]:
    """
    为位置 idx 的 token 提取 CRF 特征字典。
    特征设计遵循 CRF 最佳实践：当前词 + 上下文窗口 + 全局特征。
    """
    word = tokens[idx].lower()
    before = tokens[max(0, idx-window): idx]
    after  = tokens[idx+1: min(len(tokens), idx+1+window)]
    surr   = before + after

    f: Dict[str, any] = {}

    # ── 当前词特征 ──
    f["word"]          = word
    f["stem"]          = _stem(word)          # 词干，对词形变化鲁棒
    f["is_digit"]      = word.isdigit()
    f["has_digit"]     = any(c.isdigit() for c in word)
    f["is_upper"]      = word.isupper()
    f["prefix2"]       = word[:2]
    f["suffix2"]       = word[-2:]
    f["len_bucket"]    = min(len(word), 10)  # 词长分桶

    # ── 上下文窗口特征 ──
    for i, w in enumerate(before[-3:], 1):
        f[f"prev{i}"]       = w.lower()
        f[f"prev{i}_digit"] = w.isdigit()
    for i, w in enumerate(after[:3], 1):
        f[f"next{i}"]       = w.lower()
        f[f"next{i}_digit"] = w.isdigit()

    # ── 领域特征（布尔值，CRF 友好）──
    f["has_finance_ctx"]  = any(w in _FINANCE_WORDS  for w in surr)
    f["has_unit_after"]   = any(w in _UNIT_WORDS     for w in after)
    f["has_compare_ctx"]  = any(w in _COMPARE_WORDS  for w in surr)
    f["has_verb_before"]  = bool(before) and before[-1] in _VERB_BEFORE
    f["has_negation"]     = any(w in _NEGATION       for w in before)
    f["has_num_trigger"]  = any(w in NUMERIC_CONTEXT_TRIGGERS for w in surr)
    f["prev_is_digit"]    = bool(before) and any(c.isdigit() for c in before[-1])
    f["next_is_digit"]    = bool(after)  and any(c.isdigit() for c in after[0])
    f["is_short_text"]    = len(tokens) <= 10  # 表格单元格

    # ── 位置特征 ──
    f["is_first"] = idx == 0
    f["is_last"]  = idx == len(tokens) - 1

    return f


def _sentence_features(tokens: List[str]) -> List[Dict]:
    """为整个句子的每个 token 提取特征列表"""
    return [_word_features(tokens, i) for i in range(len(tokens))]


def _sentence_labels(tokens: List[str], numeric_indices: List[int]) -> List[str]:
    """生成 BIO 标签序列（B-NUM / O）"""
    labels = ["O"] * len(tokens)
    for i in numeric_indices:
        if 0 <= i < len(tokens):
            labels[i] = "B-NUM"
    return labels


# ─────────────────────────────────────────
# PPO 奖励信号 → 训练样本生成器
# ─────────────────────────────────────────

class PPORewardSignal:
    """
    用对齐矩阵的结果作为 PPO 奖励信号，自动生成 CRF 训练样本。

    工作流程：
      1. 对一批句对运行提取器（不经过判别器）
      2. 对齐矩阵计算 MISSING / EXTRA
      3. 根据错误类型生成正/负样本标签
      4. 累积到训练缓冲区
    """

    def __init__(self):
        self.buffer: List[Tuple[List[str], List[str]]] = []  # (tokens, labels)
        self.rewards: List[float] = []

    def observe(self,
                text: str,
                extracted_values: List[str],
                missing_values: List[str],
                extra_values: List[str]) -> float:
        """
        观察一次提取结果，计算奖励并生成训练样本。

        参数：
          text             — 原始文本
          extracted_values — 提取器输出的数值列表
          missing_values   — 对齐矩阵报告的 MISSING（漏报）
          extra_values     — 对齐矩阵报告的 EXTRA（误报）

        返回：本次奖励值
        """
        tokens = _tokenize(text)
        if not tokens:
            return 0.0

        # 计算奖励
        reward = 0.0
        reward -= len(missing_values) * 1.0   # 漏报惩罚
        reward -= len(extra_values)  * 1.0    # 误报惩罚
        reward += len([v for v in extracted_values
                       if v not in extra_values]) * 0.5  # 正确提取奖励

        # 生成标签：找出哪些 token 位置对应数值
        numeric_indices = []
        for i, tok in enumerate(tokens):
            # 简单启发：token 是数字，或者是已知多义词且在 extracted_values 中
            if tok.isdigit():
                numeric_indices.append(i)
            elif tok in NUMERIC_AMBIGUOUS and str(int(NUMERIC_AMBIGUOUS[tok])) in extracted_values:
                numeric_indices.append(i)

        labels = _sentence_labels(tokens, numeric_indices)
        self.buffer.append((tokens, labels))
        self.rewards.append(reward)
        return reward

    def get_training_data(self):
        """返回 (X_features, y_labels) 供 CRF 训练"""
        X = [_sentence_features(toks) for toks, _ in self.buffer]
        y = [labels for _, labels in self.buffer]
        return X, y

    def clear(self):
        self.buffer.clear()
        self.rewards.clear()


# ─────────────────────────────────────────
# CRF 判别器主类
# ─────────────────────────────────────────

class CRFDiscriminator:
    """
    CRF 序列标注判别器。

    - 有模型文件时：用 CRF 预测
    - 无模型文件时：自动回退到规则判别器（兼容现有流程）
    - 黑名单规则始终作为硬约束（不可被模型覆盖）
    """

    def __init__(self, model_path: Optional[str] = None):
        self.model = None
        self.model_path = model_path or os.path.join(
            os.path.dirname(__file__), "crf_model.pkl"
        )
        self._try_load()

    def _try_load(self):
        if os.path.exists(self.model_path):
            try:
                self.load(self.model_path)
                print(f"[CRF] 已加载模型: {self.model_path}")
            except Exception as e:
                print(f"[CRF] 模型加载失败，回退到规则判别器: {e}")

    def train(self, X: List[List[Dict]], y: List[List[str]],
              max_iterations: int = 100, c1: float = 0.1, c2: float = 0.1):
        """
        训练 CRF 模型。

        参数：
          X              — 特征序列列表（每个元素是一个句子的特征字典列表）
          y              — 标签序列列表（BIO 格式）
          max_iterations — 最大迭代次数
          c1, c2         — L1/L2 正则化系数
        """
        try:
            import sklearn_crfsuite
        except ImportError:
            raise ImportError(
                "请安装 sklearn-crfsuite: pip install sklearn-crfsuite"
            )

        self.model = sklearn_crfsuite.CRF(
            algorithm="lbfgs",
            c1=c1,
            c2=c2,
            max_iterations=max_iterations,
            all_possible_transitions=True,
        )
        self.model.fit(X, y)
        print(f"[CRF] 训练完成，样本数: {len(X)}")

    def train_from_pairs(self, pairs: List[Tuple[str, str, List[str]]],
                         **kwargs):
        """
        从标注句对训练。

        参数：
          pairs — [(text, label_str, numeric_tokens), ...]
                  label_str: "NUM" 或 "O"
                  numeric_tokens: 该句中应被标注为数值的 token 列表
        """
        X, y = [], []
        for text, _, numeric_tokens in pairs:
            tokens = _tokenize(text)
            if not tokens:
                continue
            numeric_set = set(t.lower() for t in numeric_tokens)
            numeric_indices = [i for i, t in enumerate(tokens) if t in numeric_set]
            X.append(_sentence_features(tokens))
            y.append(_sentence_labels(tokens, numeric_indices))
        self.train(X, y, **kwargs)

    def save(self, path: Optional[str] = None):
        path = path or self.model_path
        with open(path, "wb") as f:
            pickle.dump(self.model, f)
        print(f"[CRF] 模型已保存: {path}")

    def load(self, path: Optional[str] = None):
        path = path or self.model_path
        with open(path, "rb") as f:
            self.model = pickle.load(f)

    def predict(self, word: str, context: str,
                window: int = 5) -> Tuple[bool, float, str]:
        """
        判断 word 在 context 中是否表示数值。
        返回 (is_numeric, confidence, reason)

        优先级：
          1. 黑名单硬约束（绝对真理）
          2. CRF 模型预测（有模型时）
          3. 规则回退（无模型时）
        """
        word_lower = word.lower()

        # ── 1. 黑名单硬约束 ──
        if word_lower in AMBIGUOUS_WORDS or _stem(word_lower) in AMBIGUOUS_WORDS:
            check_key = word_lower if word_lower in AMBIGUOUS_WORDS else _stem(word_lower)
            tokens_ctx = _tokenize(context)
            try:
                idx = next(
                    (i for i, t in enumerate(tokens_ctx) if t == word_lower),
                    next((i for i, t in enumerate(tokens_ctx) if _stem(t) == _stem(word_lower)), None)
                )
                if idx is not None:
                    before = tokens_ctx[max(0, idx-window): idx]
                    after  = tokens_ctx[idx+1: idx+1+window]
                    surr   = before + after
                else:
                    surr = []
            except StopIteration:
                surr = []
            for collocations, reason in AMBIGUOUS_WORDS.get(check_key, []):
                if any(c in surr for c in collocations):
                    return False, -1.0, f"黑名单: {reason}"

        # ── 2. CRF 模型预测 ──
        if self.model is not None:
            tokens = _tokenize(context)
            word_stem = _stem(word_lower)
            try:
                # 先精确匹配，再词干匹配
                idx = next(
                    (i for i, t in enumerate(tokens) if t == word_lower),
                    next((i for i, t in enumerate(tokens) if _stem(t) == word_stem), None)
                )
                if idx is None:
                    return True, 0.5, "上下文未找到词，保守通过"
            except StopIteration:
                return True, 0.5, "上下文未找到词，保守通过"

            features = _sentence_features(tokens)
            # 预测整个序列，取目标词的标签
            pred_labels = self.model.predict([features])[0]
            label = pred_labels[idx]

            # 获取置信度（边际概率）
            try:
                marginals = self.model.predict_marginals([features])[0]
                confidence = marginals[idx].get("B-NUM", 0.0)
            except Exception:
                confidence = 1.0 if label == "B-NUM" else 0.0

            is_num = label == "B-NUM"
            # 置信度过低时回退到规则判别器，避免低质量预测
            if confidence < 0.15:
                rule_result, rule_score, rule_reason = _rule_fallback(word_lower, context, window)
                return rule_result, confidence, f"低置信度回退规则: {rule_reason}"
            return is_num, confidence, f"CRF预测: {label} (conf={confidence:.2f})"

        # ── 3. 规则回退 ──
        return _rule_fallback(word_lower, context, window)

    def filter_tokens(self, text: str, values: List[str]) -> List[str]:
        """
        对已提取的数值列表做多义词过滤（替换 rl_discriminator.filter_ambiguous_tokens）
        """
        result = list(values)
        for word, num_val in NUMERIC_AMBIGUOUS.items():
            if not re.search(r'\b' + re.escape(word) + r'\b', text, re.I):
                continue
            str_val = str(int(num_val)) if num_val == int(num_val) else str(num_val)
            is_num, _, _ = self.predict(word, text)
            if not is_num and str_val in result:
                result.remove(str_val)
        return result


# ─────────────────────────────────────────
# 规则回退（无模型时使用，与原 rl_discriminator 等价）
# ─────────────────────────────────────────

def _rule_fallback(word: str, context: str, window: int = 5) -> Tuple[bool, float, str]:
    """原有规则判别器，作为 CRF 不可用时的回退"""
    tokens = _tokenize(context)
    try:
        idx = next(i for i, t in enumerate(tokens) if t == word)
    except StopIteration:
        return True, 0.5, "上下文未找到词，保守通过"

    before = tokens[max(0, idx-window): idx]
    after  = tokens[idx+1: idx+1+window]
    surr   = before + after
    score  = 0.0
    reasons = []

    if any(t in surr for t in NUMERIC_CONTEXT_TRIGGERS):
        score += 3.0; reasons.append("数值语境触发词")
    if any(u in after for u in _UNIT_WORDS):
        score += 4.0; reasons.append("后跟单位")
    if any(re.match(r'^\d', t) for t in before):
        score += 3.0; reasons.append("前有数字")
    if any(re.match(r'^\d', t) for t in after):
        score += 3.0; reasons.append("后有数字")
    if len(tokens) <= 10:
        score += 2.0; reasons.append("短文本/表格")
    if any(c in surr for c in _COMPARE_WORDS):
        score += 2.5; reasons.append("比较语境")
    if any(f in surr for f in _FINANCE_WORDS):
        score += 2.0; reasons.append("金融领域")
    if before and before[-1] in _VERB_BEFORE:
        score -= 3.0; reasons.append("前面是动词")
    if any(n in before for n in _NEGATION):
        score -= 2.0; reasons.append("否定语境")

    is_num = score >= 1.0
    return is_num, score / 10.0, " | ".join(reasons) or "无特征"


# ─────────────────────────────────────────
# 全局单例（供 checker.py 使用）
# ─────────────────────────────────────────

_default_discriminator: Optional[CRFDiscriminator] = None


def get_discriminator() -> CRFDiscriminator:
    global _default_discriminator
    if _default_discriminator is None:
        _default_discriminator = CRFDiscriminator()
    return _default_discriminator


def filter_ambiguous_tokens(text: str, values: List[str]) -> List[str]:
    """公共接口，供 checker.py 调用（替换 rl_discriminator 中的同名函数）"""
    return get_discriminator().filter_tokens(text, values)
