"""
train_crf.py — CRF 模型训练入口
=================================
数据来源与评估策略：

  训练数据：
    1. 种子数据（手工标注）的 80% → 训练集
    2. PPO 数据中"对齐矩阵无错误"的句对 → 训练集正样本
       （有 MISSING/EXTRA 的句对不用于训练，因为无法区分是提取器问题还是真实翻译错误）

  评估数据：
    - 种子数据的 20%（从未参与训练）→ 测试集
    - 评估结果才是真实泛化能力的体现

  PPO 数据的作用：
    - 扩充"正常数值语境"的训练样本，让 CRF 见过更多真实句子
    - 不用于评估（无人工标注，无法验证）
"""

import argparse
import random
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from num_checker.crf_discriminator import (
    CRFDiscriminator, _sentence_features, _sentence_labels, _tokenize, _stem,
)

# ─────────────────────────────────────────
# 种子数据（手工标注，是唯一可信的评估基准）
# 格式：(text, is_numeric, target_word)
# ─────────────────────────────────────────

SEED_DATA = [
    # ── double / doubled / doubling ──
    ("Revenue doubled to 20 billion",                True,  "doubled"),
    ("Revenue is doubling every year",               True,  "doubling"),
    ("The company achieved double the revenue",      True,  "double"),
    ("Sales grew double digits",                     True,  "double"),
    ("Users need to double-click to confirm",        False, "double"),
    ("Please double-check the figures",              False, "double"),
    ("He plays a double role in the film",           False, "double"),
    ("The drug has a double-blind trial design",     False, "double"),
    # ── triple / tripled / tripling ──
    ("Profits tripled year-on-year",                 True,  "tripled"),
    ("Output is tripling this quarter",              True,  "tripling"),
    ("Triple the investment in R&D",                 True,  "triple"),
    ("He scored a triple in the game",               False, "triple"),
    ("A triple threat performer",                    False, "triple"),
    # ── half ──
    ("Revenue fell by half",                         True,  "half"),
    ("Cut costs by half compared to last year",      True,  "half"),
    ("A half-hearted attempt at reform",             False, "half"),
    ("The half-time score was 2-1",                  False, "half"),
    ("Half-way through the project",                 False, "half"),
    # ── quarter / quarters ──
    ("Earnings declined by a quarter",               True,  "quarter"),
    ("Q3 revenue was a quarter higher",              True,  "quarter"),
    ("Profits fell by three quarters",               True,  "quarters"),
    ("The quarterback threw a touchdown",            False, "quarterback"),
    ("Quarter-final match results",                  False, "quarter"),
    # ── one / two / three ──
    ("One million units were sold",                  True,  "one"),
    ("Two billion dollars in revenue",               True,  "two"),
    ("Three times the market average",               True,  "three"),
    ("It is one of the main products",               False, "one"),
    ("Two types of mineral nutrients",               False, "two"),
    ("Three-dimensional analysis",                   False, "three"),
    ("A one-way street",                             False, "one"),
    ("A two-sided market",                           False, "two"),
    # ── yield / return / spread ──
    ("Bond yield increased by 50 basis points",      True,  "yield"),
    ("The crop yield was poor this season",          False, "yield"),
    ("Return on equity reached 15%",                 True,  "return"),
    ("He returned home after the trip",              False, "returned"),
    ("Credit spread widened by 30 bps",              True,  "spread"),
    ("The disease spread rapidly",                   False, "spread"),
    # ── 中文多义词 ──
    ("收入翻倍增长至200亿",                           True,  "翻倍"),
    ("用户需要双击确认",                              False, "双"),
    ("两类矿物质营养元素",                            False, "两"),
    ("三季度营收增长",                                True,  "三"),
    ("三料过磷酸钙",                                  False, "三"),
    ("五氧化二磷",                                    False, "五"),
    ("二水法生产工艺",                                False, "二"),
    ("一般为20%至30%",                               False, "一"),
    ("第壹城1号楼",                                   True,  "壹"),
    ("第壹期工程已完工",                              True,  "壹"),
]


# ─────────────────────────────────────────
# 样本构建
# ─────────────────────────────────────────

def _build_sample(text: str, target_word: str, is_numeric: bool):
    tokens = _tokenize(text)
    target = target_word.lower()
    target_stem = _stem(target)
    numeric_indices = []
    for i, tok in enumerate(tokens):
        tok_stem = _stem(tok)
        tok_match = (tok == target or tok_stem == target_stem or
                     tok.startswith(target) or
                     target.startswith(tok[:max(4, len(target) - 2)]))
        if tok_match and is_numeric:
            numeric_indices.append(i)
    # 阿拉伯数字始终标为数值
    for i, tok in enumerate(tokens):
        if tok.isdigit() and i not in numeric_indices:
            numeric_indices.append(i)
    return _sentence_features(tokens), _sentence_labels(tokens, numeric_indices)


def split_seed(test_ratio: float = 0.2, seed: int = 42):
    """
    将种子数据按 test_ratio 分割为训练集和测试集。
    分层采样：保证正/负样本比例在两个集合中一致。
    """
    random.seed(seed)
    pos = [s for s in SEED_DATA if s[1]]
    neg = [s for s in SEED_DATA if not s[1]]

    def _split(lst):
        n_test = max(1, round(len(lst) * test_ratio))
        test_idx = set(random.sample(range(len(lst)), n_test))
        train = [x for i, x in enumerate(lst) if i not in test_idx]
        test  = [x for i, x in enumerate(lst) if i in test_idx]
        return train, test

    train_pos, test_pos = _split(pos)
    train_neg, test_neg = _split(neg)
    return train_pos + train_neg, test_pos + test_neg


def build_from_seed(data):
    X, y = [], []
    for text, is_numeric, target in data:
        feat, lab = _build_sample(text, target, is_numeric)
        if feat:
            X.append(feat)
            y.append(lab)
    return X, y


# ─────────────────────────────────────────
# PPO 数据生成（只用无错误句对）
# ─────────────────────────────────────────

def augment_from_excel(alignment_path: str):
    """
    从对照 Excel 生成 PPO 训练数据。

    只保留"对齐矩阵无错误"的句对：
    - 无错误 → 提取器在该句上表现正确 → 可以作为正样本
    - 有错误 → 不知道是提取器问题还是真实翻译错误 → 丢弃，不污染训练集

    返回 (X_train, y_train, stats)
    """
    import pandas as pd
    from num_checker._parser_core import parse_values
    from num_checker.alignment_matrix import build_matrix

    df = pd.read_excel(alignment_path)
    X, y = [], []
    stats = {"total": 0, "used": 0, "skipped_error": 0, "skipped_empty": 0}

    for _, row in df.iterrows():
        src = str(row.get("原文", "")).strip()
        tgt = str(row.get("译文", "")).strip()
        stats["total"] += 1

        if not src or not tgt or src == "nan" or tgt == "nan":
            stats["skipped_empty"] += 1
            continue

        src_vals = parse_values(src)
        tgt_vals = parse_values(tgt)
        errors   = build_matrix(src_vals, tgt_vals)

        if errors:
            # 有对齐错误 → 不确定是提取器问题还是翻译错误 → 跳过
            stats["skipped_error"] += 1
            continue

        # 无错误 → 提取结果可信 → 用提取到的数值位置作为标注
        tokens = _tokenize(tgt)
        if not tokens:
            stats["skipped_empty"] += 1
            continue

        # 找出 tgt 中所有数值 token 的位置
        numeric_indices = []
        for i, tok in enumerate(tokens):
            if tok.isdigit():
                numeric_indices.append(i)

        X.append(_sentence_features(tokens))
        y.append(_sentence_labels(tokens, numeric_indices))
        stats["used"] += 1

    return X, y, stats


# ─────────────────────────────────────────
# 评估（只在测试集上，从未参与训练）
# ─────────────────────────────────────────

def evaluate(discriminator: CRFDiscriminator, test_data: list, label: str = "测试集"):
    correct, total = 0, 0
    errors = []
    for text, is_numeric, target in test_data:
        pred, conf, reason = discriminator.predict(target, text)
        total += 1
        if pred == is_numeric:
            correct += 1
        else:
            errors.append((text[:60], target, is_numeric, pred, conf, reason))

    acc = correct / total * 100 if total else 0
    print(f"\n📊 [{label}] 准确率: {correct}/{total} = {acc:.1f}%")
    if errors:
        print(f"   错误样本（{len(errors)} 条）：")
        for text, word, expected, got, conf, reason in errors:
            exp_str = "数值" if expected else "非数值"
            got_str = "数值" if got else "非数值"
            print(f"   [{word}] 期望={exp_str} 预测={got_str} conf={conf:.2f}")
            print(f"     {text}")
    return acc


# ─────────────────────────────────────────
# 主流程
# ─────────────────────────────────────────

def run_training(alignment_excel=None, model_output=None,
                 max_iter=200, eval_only=False, test_ratio=0.2):

    model_path = model_output or os.path.join(os.path.dirname(__file__), "crf_model.pkl")
    discriminator = CRFDiscriminator(model_path=model_path)

    # ── 分割种子数据 ──
    train_seed, test_seed = split_seed(test_ratio=test_ratio)
    print(f"📦 种子数据: 训练集 {len(train_seed)} 条 / 测试集 {len(test_seed)} 条")
    print(f"   （测试集从未参与训练，是唯一可信的准确率来源）")

    if eval_only:
        if discriminator.model is None:
            print("❌ 无模型文件，请先训练")
            return
        evaluate(discriminator, test_seed, "测试集（种子）")
        return

    # ── 构建训练数据 ──
    X, y = build_from_seed(train_seed)

    if alignment_excel:
        print(f"\n📂 从对照文件生成 PPO 数据: {alignment_excel}")
        X_aug, y_aug, stats = augment_from_excel(alignment_excel)
        print(f"   总行数: {stats['total']}")
        print(f"   ✅ 使用（无对齐错误）: {stats['used']} 条")
        print(f"   ⏭  跳过（有对齐错误，无法确认是提取器问题还是翻译错误）: {stats['skipped_error']} 条")
        print(f"   ⏭  跳过（空行）: {stats['skipped_empty']} 条")
        X += X_aug
        y += y_aug

    print(f"\n🚀 开始训练 CRF（训练样本: {len(X)} 条，最大迭代: {max_iter}）...")
    try:
        # 样本少时加大正则化防止过拟合
        c1 = 0.5 if len(X) < 100 else 0.1
        c2 = 0.5 if len(X) < 100 else 0.1
        discriminator.train(X, y, max_iterations=max_iter, c1=c1, c2=c2)
    except ImportError as e:
        print(f"❌ {e}")
        return

    # ── 评估（训练集 vs 测试集，对比看是否过拟合）──
    train_acc = evaluate(discriminator, train_seed, "训练集（种子）")
    test_acc  = evaluate(discriminator, test_seed,  "测试集（种子，未见过）")

    if train_acc - test_acc > 15:
        print(f"\n⚠️  训练集({train_acc:.1f}%) 与测试集({test_acc:.1f}%) 差距较大，"
              f"可能过拟合，建议增加种子数据或调大正则化参数")
    else:
        print(f"\n✅ 训练/测试差距正常（{train_acc:.1f}% / {test_acc:.1f}%），模型泛化良好")

    discriminator.save(model_path)
    print(f"💾 模型已保存: {model_path}")


def main():
    parser = argparse.ArgumentParser(description="CRF 判别器训练")
    parser.add_argument("--input",  "-i", help="对照 Excel 路径")
    parser.add_argument("--output", "-o", default=None, help="模型输出路径")
    parser.add_argument("--eval",   action="store_true", help="只评估不训练")
    parser.add_argument("--iter",   type=int, default=200)
    parser.add_argument("--test-ratio", type=float, default=0.2, help="测试集比例")
    args = parser.parse_args()
    run_training(
        alignment_excel=args.input,
        model_output=args.output,
        max_iter=args.iter,
        eval_only=args.eval,
        test_ratio=args.test_ratio,
    )


if __name__ == "__main__":
    # ══════════════════════════════════════════════════════════════
    # 直接填写路径运行，不需要命令行参数
    # ══════════════════════════════════════════════════════════════

    ALIGNMENT_EXCEL = r"D:\project\数检_程序-AI\测试文件\bilingual_pairs (2).xlsx"
    MODEL_OUTPUT    = None    # None = 默认 num_checker/crf_model.pkl
    MAX_ITER        = 200
    EVAL_ONLY       = False   # True = 只评估不训练
    TEST_RATIO      = 0.2     # 种子数据中留出 20% 作为测试集

    if ALIGNMENT_EXCEL is not None or MODEL_OUTPUT is not None or EVAL_ONLY:
        run_training(
            alignment_excel=ALIGNMENT_EXCEL,
            model_output=MODEL_OUTPUT,
            max_iter=MAX_ITER,
            eval_only=EVAL_ONLY,
            test_ratio=TEST_RATIO,
        )
    else:
        main()
