"""
reviewer.py -- 交互式审核工具
=============================

目标：
  - 避免人工直接编辑 JSON
  - 审核人员只需做有限选择
  - 自动回写 review_status / review_label / training_signal
"""

import argparse
import os
import sys
from typing import Any, Dict, List, Optional

if __package__ in (None, ""):
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from num_checker.learning_pipeline import (
        _utc_now,
        default_store_path,
        load_store,
        save_store,
    )
else:
    from .learning_pipeline import (
        _utc_now,
        default_store_path,
        load_store,
        save_store,
    )


def _pick_candidate(case: Dict[str, Any]) -> Dict[str, Any]:
    candidates = case.get("candidate_signals") or []
    if candidates:
        return candidates[0]
    signal = case.get("training_signal") or {}
    if signal:
        return signal
    return {
        "text": case.get("tgt_text", "") or case.get("src_text", ""),
        "target_word": "",
        "is_numeric": None,
        "side": "",
        "normalized_value": "",
        "error_type": "",
        "prefill_reason": "",
    }


def _preview_errors(case: Dict[str, Any]) -> str:
    parts: List[str] = []
    for err in case.get("errors", []):
        error_type = err.get("error_type", "")
        value = err.get("src_value") or err.get("tgt_value") or ""
        parts.append(f"{error_type}:{value}")
    return " | ".join(parts) if parts else "无"


def _print_case(case: Dict[str, Any]):
    candidate = _pick_candidate(case)
    candidates = case.get("candidate_signals") or []
    print("\n" + "=" * 72)
    print(f"case_id: {case.get('case_id', '')}")
    print(f"模型版本: {case.get('model_version', '')} -> 最近复现: {case.get('last_model_version', '')}")
    print(f"原文: {case.get('src_text', '')}")
    print(f"译文: {case.get('tgt_text', '')}")
    print(f"原文数值: {case.get('src_values', [])}")
    print(f"译文数值: {case.get('tgt_values', [])}")
    print(f"冲突点: {_preview_errors(case)}")
    print(f"建议目标词: {candidate.get('target_word', '')}")
    print(f"建议上下文侧: {candidate.get('side', '')}")
    print(f"建议上下文: {candidate.get('text', '')}")
    print(f"建议原因: {candidate.get('prefill_reason', '')}")
    if len(candidates) > 1:
        print("候选列表:")
        for idx, item in enumerate(candidates, start=1):
            print(f"  {idx}. [{item.get('side', '')}] {item.get('target_word', '')} -> {item.get('normalized_value', '')}")
    print("=" * 72)


def _approve_case(case: Dict[str, Any], is_numeric: bool, note: str = "", candidate: Optional[Dict[str, Any]] = None):
    candidate = dict(candidate or _pick_candidate(case))
    candidate["is_numeric"] = is_numeric
    case["training_signal"] = candidate
    case["review_status"] = "approved"
    case["review_label"] = "model_feedback"
    case["review_notes"] = note
    case["reviewed_at"] = _utc_now()


def _reject_case(case: Dict[str, Any], note: str = "", label: str = "translation_error"):
    case["review_status"] = "rejected"
    case["review_label"] = label
    case["review_notes"] = note
    case["reviewed_at"] = _utc_now()


def _input_note() -> str:
    note = input("备注(可留空): ").strip()
    return note


def _select_candidate(case: Dict[str, Any]) -> Dict[str, Any]:
    candidates = case.get("candidate_signals") or []
    if len(candidates) <= 1:
        return _pick_candidate(case)
    raw = input(f"候选编号(1-{len(candidates)}，回车默认1): ").strip()
    if not raw:
        return candidates[0]
    try:
        idx = int(raw)
    except ValueError:
        return candidates[0]
    if 1 <= idx <= len(candidates):
        return candidates[idx - 1]
    return candidates[0]


def review_next_pending(store_path: Optional[str] = None) -> bool:
    data = load_store(store_path)
    for case in data["cases"]:
        if str(case.get("review_status", "pending")).lower() != "pending":
            continue

        _print_case(case)
        selected_candidate = _select_candidate(case)
        print("操作:")
        print("  y = 模型误判，目标词应判为数字")
        print("  n = 模型误判，目标词应判为非数字")
        print("  t = 翻译真实错误，拒绝进入训练")
        print("  r = 非翻译类拒绝样本")
        print("  s = 跳过")
        print("  q = 退出")
        choice = input("请选择(y/n/t/r/s/q): ").strip().lower()

        if choice == "y":
            _approve_case(case, True, _input_note(), selected_candidate)
        elif choice == "n":
            _approve_case(case, False, _input_note(), selected_candidate)
        elif choice == "t":
            _reject_case(case, _input_note(), label="translation_error")
        elif choice == "r":
            _reject_case(case, _input_note(), label="other_rejected")
        elif choice == "s":
            return True
        elif choice == "q":
            save_store(data, store_path)
            return False
        else:
            print("输入无效，本条未修改。")
            return True

        save_store(data, store_path)
        print("已回写审核结果。")
        return True

    print("没有待审核案例。")
    return False


def main():
    parser = argparse.ArgumentParser(description="学习仓库交互式审核工具")
    parser.add_argument("--store", default=default_store_path(), help="学习仓库 JSON 路径")
    parser.add_argument("--loop", action="store_true", help="连续审核直到退出")
    args = parser.parse_args()

    if args.loop:
        while review_next_pending(args.store):
            pass
        return

    review_next_pending(args.store)


if __name__ == "__main__":
    main()
