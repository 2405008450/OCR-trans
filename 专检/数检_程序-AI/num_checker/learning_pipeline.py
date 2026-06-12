"""
learning_pipeline.py -- 自学习闭环主流程
=======================================

职责：
  1. 将检查结果沉淀到统一学习仓库
  2. 为人工审核预填候选 training_signal
  3. 导出 rejected 翻译错题本
  4. 基于审核通过样本训练候选 CRF 并做门禁晋级
"""

import argparse
import hashlib
import json
import os
import random
import sys
from datetime import datetime, timezone
from typing import Any, Dict, Iterable, List, Optional, Tuple

if __package__ in (None, ""):
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from num_checker.crf_discriminator import CRFDiscriminator, _stem
    from num_checker.domain_rules import NUMERIC_AMBIGUOUS
    from num_checker.symbolic_parser import extract as extract_tokens
    from num_checker.train_crf import _build_sample, build_from_seed, split_seed
else:
    from .crf_discriminator import CRFDiscriminator, _stem
    from .domain_rules import NUMERIC_AMBIGUOUS
    from .symbolic_parser import extract as extract_tokens
    from .train_crf import _build_sample, build_from_seed, split_seed


FeedbackSample = Tuple[str, bool, str]
_CN_AMBIGUOUS_CHARS = set("零〇一二两三四五六七八九十壹贰叁肆伍陆柒捌玖拾半双俩")


def _utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def default_store_path() -> str:
    return os.path.join(os.path.dirname(__file__), "learning_data", "learning_store.json")


def default_candidate_model_path() -> str:
    return os.path.join(os.path.dirname(__file__), "learning_data", "candidate_crf_model.pkl")


def default_active_model_path() -> str:
    return os.path.join(os.path.dirname(__file__), "crf_model.pkl")


def default_model_meta_path() -> str:
    return os.path.join(os.path.dirname(__file__), "learning_data", "model_meta.json")


def default_translation_notebook_path() -> str:
    return os.path.join(os.path.dirname(__file__), "learning_data", "translation_error_notebook.xlsx")


def _ensure_parent(path: str):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)


def _default_training_signal() -> Dict[str, Any]:
    return {
        "text": "",
        "target_word": "",
        "is_numeric": None,
        "side": "",
        "normalized_value": "",
        "error_type": "",
        "prefill_reason": "",
    }


def _default_model_meta() -> Dict[str, Any]:
    return {
        "version": 1,
        "active_model_version": "v1.0.0",
        "last_promoted_at": "",
        "history": [],
    }


def _load_model_meta(model_meta_path: Optional[str] = None) -> Dict[str, Any]:
    path = model_meta_path or default_model_meta_path()
    if not os.path.exists(path):
        return _default_model_meta()
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, dict):
        return _default_model_meta()
    base = _default_model_meta()
    base.update(data)
    if not isinstance(base.get("history"), list):
        base["history"] = []
    return base


def _save_model_meta(data: Dict[str, Any], model_meta_path: Optional[str] = None):
    path = model_meta_path or default_model_meta_path()
    _ensure_parent(path)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def get_active_model_version(model_meta_path: Optional[str] = None) -> str:
    path = model_meta_path or default_model_meta_path()
    meta = _load_model_meta(path)
    if not os.path.exists(path):
        _save_model_meta(meta, path)
    return str(meta.get("active_model_version") or "v1.0.0")


def _bump_version(version: str) -> str:
    if version.startswith("v"):
        body = version[1:]
        parts = body.split(".")
        if len(parts) == 3 and all(p.isdigit() for p in parts):
            major, minor, patch = [int(p) for p in parts]
            return f"v{major}.{minor}.{patch + 1}"
    return version + ".1"


def _serialize_errors(errors: Iterable[Any]) -> List[Dict[str, Any]]:
    result = []
    for e in errors:
        result.append({
            "error_type": getattr(e, "error_type", ""),
            "src_value": getattr(e, "src_value", None),
            "tgt_value": getattr(e, "tgt_value", None),
            "message": getattr(e, "message", ""),
            "position": getattr(e, "position", -1),
        })
    return result


def _candidate_score(raw: str, side: str, error_type: str) -> int:
    raw_lower = raw.lower()
    score = 0
    if side == "src" and error_type == "MISSING":
        score += 4
    if side == "tgt" and error_type == "EXTRA":
        score += 4
    if any(ch.isdigit() for ch in raw):
        score -= 2
    else:
        score += 2
    if raw_lower in NUMERIC_AMBIGUOUS or _stem(raw_lower) in NUMERIC_AMBIGUOUS:
        score += 8
    if len(raw) == 1 and raw in _CN_AMBIGUOUS_CHARS:
        score += 6
    return score


def _make_candidate(text: str, raw: str, normalized_value: str, side: str, error_type: str, reason: str) -> Dict[str, Any]:
    return {
        "text": text,
        "target_word": raw,
        "is_numeric": None,
        "side": side,
        "normalized_value": normalized_value,
        "error_type": error_type,
        "prefill_reason": reason,
        "score": _candidate_score(raw, side, error_type),
    }


def infer_candidate_signals(src_text: str, tgt_text: str, errors: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    根据对齐错误自动预填候选 training_signal。

    策略：
      - MISSING 优先从原文 token 中找对应值
      - EXTRA 优先从译文 token 中找对应值
      - 优先选择非纯数字、多义词、单字中文数字等更适合 CRF 学习的目标词
    """
    src_tokens = extract_tokens(src_text)
    tgt_tokens = extract_tokens(tgt_text)
    candidates: List[Dict[str, Any]] = []

    for err in errors:
        error_type = str(err.get("error_type") or "")
        if error_type == "MISSING" and err.get("src_value") is not None:
            value = str(err["src_value"])
            for token in src_tokens:
                if token.value == value:
                    candidates.append(_make_candidate(
                        src_text, token.raw, value, "src", error_type, "根据原文未匹配数值自动预填"
                    ))
        if error_type == "EXTRA" and err.get("tgt_value") is not None:
            value = str(err["tgt_value"])
            for token in tgt_tokens:
                if token.value == value:
                    candidates.append(_make_candidate(
                        tgt_text, token.raw, value, "tgt", error_type, "根据译文多出数值自动预填"
                    ))

    dedup: Dict[Tuple[str, str, str, str], Dict[str, Any]] = {}
    for candidate in candidates:
        key = (
            candidate["side"],
            candidate["target_word"],
            candidate["normalized_value"],
            candidate["error_type"],
        )
        current = dedup.get(key)
        if current is None or candidate["score"] > current["score"]:
            dedup[key] = candidate

    ranked = sorted(
        dedup.values(),
        key=lambda x: (-int(x.get("score", 0)), x["side"], x["target_word"]),
    )
    for item in ranked:
        item.pop("score", None)
    return ranked


def _make_case_id(src_text: str, tgt_text: str) -> str:
    payload = f"{src_text}\n---\n{tgt_text}".encode("utf-8")
    return hashlib.sha1(payload).hexdigest()


def _upgrade_case_schema(case: Dict[str, Any]):
    case.setdefault("review_status", "pending")
    case.setdefault("review_label", "")
    case.setdefault("review_notes", "")
    case.setdefault("reviewed_at", "")
    case.setdefault("model_version", "")
    case.setdefault("last_model_version", case.get("model_version", ""))
    if not case.get("model_version") and case.get("last_model_version"):
        case["model_version"] = case["last_model_version"]
    case.setdefault("training_signal", _default_training_signal())
    case.setdefault("candidate_signals", [])
    if not isinstance(case.get("training_signal"), dict):
        case["training_signal"] = _default_training_signal()
    else:
        base = _default_training_signal()
        base.update(case["training_signal"])
        case["training_signal"] = base
    if not isinstance(case.get("candidate_signals"), list):
        case["candidate_signals"] = []


def _load_store(store_path: Optional[str] = None) -> Dict[str, Any]:
    path = store_path or default_store_path()
    if not os.path.exists(path):
        return {"version": 2, "updated_at": _utc_now(), "cases": []}
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if "cases" not in data or not isinstance(data["cases"], list):
        data["cases"] = []
    data["version"] = 2
    for case in data["cases"]:
        if isinstance(case, dict):
            _upgrade_case_schema(case)
    return data


def _save_store(data: Dict[str, Any], store_path: Optional[str] = None):
    path = store_path or default_store_path()
    _ensure_parent(path)
    data["updated_at"] = _utc_now()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_store(store_path: Optional[str] = None) -> Dict[str, Any]:
    return _load_store(store_path)


def save_store(data: Dict[str, Any], store_path: Optional[str] = None):
    _save_store(data, store_path)


def queue_check_results(results: Iterable[Any],
                        store_path: Optional[str] = None,
                        source: str = "run",
                        errors_only: bool = True,
                        model_version: Optional[str] = None,
                        model_meta_path: Optional[str] = None) -> Dict[str, int]:
    """
    将检查结果写入学习仓库，并预填候选 training_signal。
    """
    current_model_version = model_version or get_active_model_version(model_meta_path)
    data = _load_store(store_path)
    cases = data["cases"]
    index = {case["case_id"]: case for case in cases if "case_id" in case}
    added = 0
    updated = 0

    for result in results:
        if errors_only and getattr(result, "is_correct", False):
            continue

        src_text = getattr(result, "src_text", "")
        tgt_text = getattr(result, "tgt_text", "")
        errors = _serialize_errors(getattr(result, "errors", []))
        candidates = infer_candidate_signals(src_text, tgt_text, errors)
        prefilled_signal = candidates[0] if candidates else _default_training_signal()
        case_id = _make_case_id(src_text, tgt_text)
        now = _utc_now()

        payload = {
            "case_id": case_id,
            "source": source,
            "src_text": src_text,
            "tgt_text": tgt_text,
            "src_values": list(getattr(result, "src_values", [])),
            "tgt_values": list(getattr(result, "tgt_values", [])),
            "errors": errors,
            "is_correct": bool(getattr(result, "is_correct", False)),
            "summary": getattr(result, "summary", ""),
            "review_status": "pending",
            "review_label": "",
            "review_notes": "",
            "reviewed_at": "",
            "training_signal": prefilled_signal,
            "candidate_signals": candidates,
            "model_version": current_model_version,
            "last_model_version": current_model_version,
            "seen_count": 1,
            "first_seen_at": now,
            "last_seen_at": now,
        }

        if case_id in index:
            current = index[case_id]
            _upgrade_case_schema(current)
            current["source"] = source
            current["src_values"] = payload["src_values"]
            current["tgt_values"] = payload["tgt_values"]
            current["errors"] = payload["errors"]
            current["is_correct"] = payload["is_correct"]
            current["summary"] = payload["summary"]
            current["candidate_signals"] = candidates
            current["last_seen_at"] = now
            if not current.get("model_version"):
                current["model_version"] = current_model_version
            current["last_model_version"] = current_model_version
            current["seen_count"] = int(current.get("seen_count", 0)) + 1
            signal = current.get("training_signal") or {}
            if not str(signal.get("target_word") or "").strip() and candidates:
                current["training_signal"] = prefilled_signal
            updated += 1
            continue

        cases.append(payload)
        index[case_id] = payload
        added += 1

    _save_store(data, store_path)
    return {"added": added, "updated": updated, "total_cases": len(cases)}


def summarize_store(store_path: Optional[str] = None) -> Dict[str, int]:
    data = _load_store(store_path)
    summary = {
        "total": 0,
        "pending": 0,
        "approved": 0,
        "rejected": 0,
        "trainable": 0,
        "translation_error": 0,
    }
    for case in data["cases"]:
        summary["total"] += 1
        status = str(case.get("review_status", "pending")).lower()
        if status not in ("pending", "approved", "rejected"):
            status = "pending"
        summary[status] += 1
        if str(case.get("review_label", "")).lower() == "translation_error":
            summary["translation_error"] += 1
        if _case_to_feedback_sample(case) is not None:
            summary["trainable"] += 1
    return summary


def _case_to_feedback_sample(case: Dict[str, Any]) -> Optional[FeedbackSample]:
    if str(case.get("review_status", "")).lower() != "approved":
        return None
    signal = case.get("training_signal") or {}
    text = str(signal.get("text") or "").strip()
    target_word = str(signal.get("target_word") or "").strip()
    is_numeric = signal.get("is_numeric")
    if not text or not target_word or not isinstance(is_numeric, bool):
        return None
    return (text, is_numeric, target_word)


def load_feedback_samples(store_path: Optional[str] = None) -> List[FeedbackSample]:
    data = _load_store(store_path)
    samples = []
    for case in data["cases"]:
        sample = _case_to_feedback_sample(case)
        if sample is not None:
            samples.append(sample)
    return samples


def _split_samples(samples: List[FeedbackSample],
                   test_ratio: float = 0.2,
                   seed: int = 42) -> Tuple[List[FeedbackSample], List[FeedbackSample]]:
    if not samples:
        return [], []

    random.seed(seed)
    pos = [s for s in samples if s[1]]
    neg = [s for s in samples if not s[1]]

    def _split(group: List[FeedbackSample]) -> Tuple[List[FeedbackSample], List[FeedbackSample]]:
        if len(group) <= 1:
            return group[:], []
        n_test = max(1, round(len(group) * test_ratio))
        n_test = min(n_test, len(group) - 1)
        test_idx = set(random.sample(range(len(group)), n_test))
        train = [x for i, x in enumerate(group) if i not in test_idx]
        test = [x for i, x in enumerate(group) if i in test_idx]
        return train, test

    train_pos, test_pos = _split(pos)
    train_neg, test_neg = _split(neg)
    return train_pos + train_neg, test_pos + test_neg


def _rebalance_feedback_samples(seed_samples: list,
                                feedback_samples: List[FeedbackSample],
                                seed_min_share: float = 0.5,
                                seed: int = 42) -> Tuple[List[FeedbackSample], int]:
    """
    限制反馈样本在训练集中的占比，避免近期错误案例过度主导模型。
    """
    if not feedback_samples or seed_min_share <= 0:
        return feedback_samples, len(feedback_samples)
    seed_count = len(seed_samples)
    if seed_count <= 0 or seed_min_share >= 1:
        return [], 0
    max_feedback = int(seed_count * (1 - seed_min_share) / seed_min_share)
    if max_feedback <= 0:
        return [], 0
    if len(feedback_samples) <= max_feedback:
        return feedback_samples, len(feedback_samples)
    random.seed(seed)
    return random.sample(feedback_samples, max_feedback), max_feedback


def _build_feedback_train_data(samples: List[FeedbackSample]) -> Tuple[List[List[dict]], List[List[str]]]:
    X, y = [], []
    for text, is_numeric, target_word in samples:
        feat, lab = _build_sample(text, target_word, is_numeric)
        if feat:
            X.append(feat)
            y.append(lab)
    return X, y


def _evaluate_accuracy(discriminator: CRFDiscriminator, dataset: List[FeedbackSample]) -> float:
    if not dataset:
        return 0.0
    correct = 0
    for text, is_numeric, target_word in dataset:
        pred, _, _ = discriminator.predict(target_word, text)
        if pred == is_numeric:
            correct += 1
    return correct / len(dataset) * 100


def export_translation_errors(store_path: Optional[str] = None,
                              output_path: Optional[str] = None) -> Dict[str, Any]:
    """
    将 rejected / translation_error 样本导出为业务侧翻译错题本。
    """
    store = _load_store(store_path)
    rows: List[Dict[str, Any]] = []
    for case in store["cases"]:
        status = str(case.get("review_status", "")).lower()
        label = str(case.get("review_label", "")).lower()
        if status != "rejected":
            continue
        if label and label != "translation_error":
            continue
        rows.append({
            "case_id": case.get("case_id", ""),
            "model_version": case.get("model_version", ""),
            "last_model_version": case.get("last_model_version", ""),
            "原文": case.get("src_text", ""),
            "译文": case.get("tgt_text", ""),
            "原文数值": "，".join(case.get("src_values", [])),
            "译文数值": "，".join(case.get("tgt_values", [])),
            "冲突摘要": case.get("summary", ""),
            "审核标签": case.get("review_label", ""),
            "审核备注": case.get("review_notes", ""),
            "首次发现": case.get("first_seen_at", ""),
            "最后发现": case.get("last_seen_at", ""),
        })

    output = output_path or default_translation_notebook_path()
    _ensure_parent(output)
    try:
        import pandas as pd
        pd.DataFrame(rows).to_excel(output, index=False)
    except Exception:
        fallback = os.path.splitext(output)[0] + ".json"
        with open(fallback, "w", encoding="utf-8") as f:
            json.dump(rows, f, ensure_ascii=False, indent=2)
        output = fallback

    return {"count": len(rows), "output_path": output}


def train_and_promote(store_path: Optional[str] = None,
                      active_model_path: Optional[str] = None,
                      candidate_model_path: Optional[str] = None,
                      model_meta_path: Optional[str] = None,
                      max_iter: int = 200,
                      test_ratio: float = 0.2,
                      min_improvement: float = 0.0,
                      seed_min_share: float = 0.5) -> Dict[str, Any]:
    """
    用审核通过的反馈样本训练候选模型，并基于门禁决定是否升级正式模型。
    """
    store = store_path or default_store_path()
    active_path = active_model_path or default_active_model_path()
    candidate_path = candidate_model_path or default_candidate_model_path()
    meta_path = model_meta_path or default_model_meta_path()

    feedback_samples = load_feedback_samples(store)
    if not feedback_samples:
        return {
            "trained": False,
            "promoted": False,
            "reason": "没有已审核且可训练的反馈样本",
        }

    feedback_train_raw, feedback_eval = _split_samples(feedback_samples, test_ratio=test_ratio)
    seed_train, seed_eval = split_seed(test_ratio=test_ratio)
    feedback_train, feedback_used = _rebalance_feedback_samples(
        seed_train, feedback_train_raw, seed_min_share=seed_min_share
    )

    X_seed, y_seed = build_from_seed(seed_train)
    X_feedback, y_feedback = _build_feedback_train_data(feedback_train)
    X = X_seed + X_feedback
    y = y_seed + y_feedback
    if not X:
        return {
            "trained": False,
            "promoted": False,
            "reason": "训练数据为空",
        }

    eval_set: List[FeedbackSample] = [(text, is_numeric, word) for text, is_numeric, word in seed_eval]
    eval_set.extend(feedback_eval)

    baseline = CRFDiscriminator(model_path=active_path)
    baseline_version = get_active_model_version(meta_path)
    baseline_seed_acc = _evaluate_accuracy(baseline, seed_eval)
    baseline_eval_acc = _evaluate_accuracy(baseline, eval_set)

    candidate = CRFDiscriminator(model_path=candidate_path)
    candidate.model = None
    c1 = 0.5 if len(X) < 100 else 0.1
    c2 = 0.5 if len(X) < 100 else 0.1
    candidate.train(X, y, max_iterations=max_iter, c1=c1, c2=c2)

    candidate_seed_acc = _evaluate_accuracy(candidate, seed_eval)
    candidate_eval_acc = _evaluate_accuracy(candidate, eval_set)

    _ensure_parent(candidate_path)
    candidate.save(candidate_path)

    promoted = (
        candidate_seed_acc + 1e-9 >= baseline_seed_acc and
        candidate_eval_acc + 1e-9 >= baseline_eval_acc + min_improvement
    )

    promoted_version = baseline_version
    if promoted:
        promoted_version = _bump_version(baseline_version)
        candidate.save(active_path)
        meta = _load_model_meta(meta_path)
        meta["active_model_version"] = promoted_version
        meta["last_promoted_at"] = _utc_now()
        meta["history"].append({
            "from_version": baseline_version,
            "to_version": promoted_version,
            "promoted_at": meta["last_promoted_at"],
            "baseline_seed_acc": round(baseline_seed_acc, 2),
            "candidate_seed_acc": round(candidate_seed_acc, 2),
            "baseline_eval_acc": round(baseline_eval_acc, 2),
            "candidate_eval_acc": round(candidate_eval_acc, 2),
        })
        _save_model_meta(meta, meta_path)

    return {
        "trained": True,
        "promoted": promoted,
        "active_model_path": active_path,
        "candidate_model_path": candidate_path,
        "baseline_version": baseline_version,
        "candidate_version": promoted_version if promoted else baseline_version + "-candidate",
        "feedback_total": len(feedback_samples),
        "feedback_train_raw": len(feedback_train_raw),
        "feedback_train_used": feedback_used,
        "feedback_eval": len(feedback_eval),
        "train_samples": len(X),
        "seed_train_samples": len(X_seed),
        "feedback_train_samples": len(X_feedback),
        "baseline_seed_acc": round(baseline_seed_acc, 2),
        "candidate_seed_acc": round(candidate_seed_acc, 2),
        "baseline_eval_acc": round(baseline_eval_acc, 2),
        "candidate_eval_acc": round(candidate_eval_acc, 2),
        "seed_min_share": seed_min_share,
    }


def _print_summary(summary: Dict[str, int], store_path: str):
    print(f"学习仓库: {store_path}")
    print(f"   总案例: {summary['total']}")
    print(f"   待审核: {summary['pending']}")
    print(f"   已通过: {summary['approved']}")
    print(f"   已拒绝: {summary['rejected']}")
    print(f"   可训练: {summary['trainable']}")
    print(f"   翻译错误: {summary['translation_error']}")


def _print_training_result(result: Dict[str, Any]):
    if not result.get("trained"):
        print(f"[!] 未执行训练: {result.get('reason', '未知原因')}")
        return

    print("候选模型训练完成")
    print(f"   基线版本: {result['baseline_version']}")
    print(f"   反馈样本总量: {result['feedback_total']} 条")
    print(f"   反馈训练原始量: {result['feedback_train_raw']} 条")
    print(f"   反馈训练实际使用: {result['feedback_train_used']} 条")
    print(f"   反馈评估量: {result['feedback_eval']} 条")
    print(f"   seed / feedback 训练样本: {result['seed_train_samples']} / {result['feedback_train_samples']}")
    print(f"   seed 最低占比约束: {result['seed_min_share']}")
    print(f"   总训练样本: {result['train_samples']}")
    print(f"   基线种子准确率: {result['baseline_seed_acc']}%")
    print(f"   候选种子准确率: {result['candidate_seed_acc']}%")
    print(f"   基线综合准确率: {result['baseline_eval_acc']}%")
    print(f"   候选综合准确率: {result['candidate_eval_acc']}%")
    print(f"   候选模型: {result['candidate_model_path']}")
    if result["promoted"]:
        print(f"候选模型通过门禁，已升级正式模型: {result['active_model_path']}")
        print(f"   新版本: {result['candidate_version']}")
    else:
        print("候选模型未通过门禁，正式模型保持不变")


def main():
    parser = argparse.ArgumentParser(description="自学习反馈仓库与增量训练入口")
    parser.add_argument("--store", default=default_store_path(), help="学习仓库 JSON 路径")
    parser.add_argument("--train", action="store_true", help="基于已审核样本训练候选模型")
    parser.add_argument("--export-translation-errors", action="store_true", help="导出 rejected 翻译错题本")
    parser.add_argument("--translation-notebook", default=default_translation_notebook_path(), help="翻译错题本输出路径")
    parser.add_argument("--active-model", default=default_active_model_path(), help="正式模型路径")
    parser.add_argument("--candidate-model", default=default_candidate_model_path(), help="候选模型路径")
    parser.add_argument("--model-meta", default=default_model_meta_path(), help="模型版本元数据路径")
    parser.add_argument("--iter", type=int, default=200, help="CRF 最大迭代次数")
    parser.add_argument("--test-ratio", type=float, default=0.2, help="反馈样本评估比例")
    parser.add_argument("--min-improvement", type=float, default=0.0, help="候选模型综合评估所需最小提升")
    parser.add_argument("--seed-min-share", type=float, default=0.5, help="训练集中 seed 数据最低占比")
    args = parser.parse_args()

    summary = summarize_store(args.store)
    _print_summary(summary, args.store)
    print(f"当前模型版本: {get_active_model_version(args.model_meta)}")

    if args.export_translation_errors:
        exported = export_translation_errors(args.store, args.translation_notebook)
        print(f"翻译错题本已导出: {exported['output_path']}")
        print(f"   导出条数: {exported['count']}")

    if not args.train:
        return

    result = train_and_promote(
        store_path=args.store,
        active_model_path=args.active_model,
        candidate_model_path=args.candidate_model,
        model_meta_path=args.model_meta,
        max_iter=args.iter,
        test_ratio=args.test_ratio,
        min_improvement=args.min_improvement,
        seed_min_share=args.seed_min_share,
    )
    _print_training_result(result)


if __name__ == "__main__":
    main()
