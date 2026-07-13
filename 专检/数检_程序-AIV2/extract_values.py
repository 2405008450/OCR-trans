"""
数值检查数据接口层

所有公开接口均返回 JSON 可序列化的 Python 原生类型（dict / list），
不依赖 DataFrame，供上层按需使用。

接口分三组：
  A. 提取类  — extract_numbers(text) -> List[str]
  B. 对比类  — compare_numbers(src, tgt) -> List[Dict]
               check_row(src, tgt)       -> Dict
  C. 流程类  — run_number_check(path)    -> List[Dict]   规则检查结果
               merge_ai_results(rows, ai_map) -> List[Dict]  合并AI结果
"""
import json
from typing import List, Dict, Optional

from normalizer_total import extract_numbers as extract_numbers, DEFAULT_STRATEGIES


# ─────────────────────────────────────────
# A. 提取类（由 normalizer_total.extract_numbers 提供）
# ─────────────────────────────────────────


# ─────────────────────────────────────────
# B. 对比类
# ─────────────────────────────────────────

def compare_numbers(src_list: List[str], tgt_list: List[str]) -> List[Dict]:
    """对比原文/译文数值列表，返回差异项列表。
    每项: {"type": "MISSING"|"EXTRA"|"ORDER", "msg": str}
    """
    src_used = [False] * len(src_list)
    tgt_used = [False] * len(tgt_list)
    diffs = []

    for i, s in enumerate(src_list):
        for j, t in enumerate(tgt_list):
            if not tgt_used[j] and s == t:
                src_used[i] = True
                tgt_used[j] = True
                break

    for i, s in enumerate(src_list):
        if not src_used[i]:
            diffs.append({"type": "MISSING", "msg": f"译文缺失数值 {s}（原文存在）"})

    for j, t in enumerate(tgt_list):
        if not tgt_used[j]:
            diffs.append({"type": "EXTRA", "msg": f"译文多出数值 {t}（原文不存在）"})

    if len(src_list) == len(tgt_list) and not diffs:
        for i in range(len(src_list)):
            if src_list[i] != tgt_list[i]:
                diffs.append({"type": "ORDER", "msg": f"数值顺序变化：{src_list[i]} → {tgt_list[i]}"})

    return diffs


def check_row(src: str, tgt: str, strategies: Optional[Dict[str, bool]] = None) -> Dict:
    """对单行原文/译文做完整数值检查。

    返回 JSON：
    {
        "原文":     str,
        "译文":     str,
        "原文数值": ["123", ...],
        "译文数值": ["123", ...],
        "是否错误": "❗错误" | "✔正确",
        "错误类型": "MISSING；EXTRA",   # 分号分隔
        "错误原因": "译文缺失...",       # 分号分隔
    }
    """
    src_nums = extract_numbers(src, strategies)
    tgt_nums = extract_numbers(tgt, strategies)
    diffs    = compare_numbers(src_nums, tgt_nums)

    return {
        "原文":     src,
        "译文":     tgt,
        "原文数值": src_nums,
        "译文数值": tgt_nums,
        "是否错误": "❗错误" if diffs else "✔正确",
        "错误类型": "；".join(d["type"] for d in diffs),
        "错误原因": "；".join(d["msg"]  for d in diffs),
    }


# ─────────────────────────────────────────
# C. 流程类
# ─────────────────────────────────────────

def run_number_check(alignment_path: str,
                     strategies: Optional[Dict[str, bool]] = None) -> List[Dict]:
    """
    读取对齐 Excel → 逐行规则检查 → 返回 List[Dict]（JSON 可序列化）。

    strategies: 归化策略开关，None 表示不启用归化。
    每项结构见 check_row() 返回值。
    """
    import pandas as pd
    df = pd.read_excel(alignment_path)
    return [
        check_row(str(row.get("原文", "")), str(row.get("译文", "")), strategies)
        for _, row in df.iterrows()
    ]


def check_text_pairs(pairs: List[tuple],
                     strategies: Optional[Dict[str, bool]] = None) -> List[Dict]:
    """
    对 [(src, tgt), ...] 文本对列表做规则检查，返回与 run_number_check 相同结构的 List[Dict]。
    用于页眉、页脚等非 Excel 来源的文本对检查。
    """
    return [check_row(str(src), str(tgt), strategies) for src, tgt, *_ in pairs]


def merge_ai_results(rows: List[Dict], ai_map: Dict[int, Dict]) -> List[Dict]:
    """
    将 AI 检查结果合并到规则检查结果中，返回 List[Dict]（JSON 可序列化）。

    参数：
      rows   — run_number_check() 的返回值
      ai_map — {行索引: {"is_correct": bool, "errors": [{"错误编号","原文上下文","译文上下文","原文数值","译文数值","替换锚点","译文修改建议值","错误类型","修改理由","违反的规则"}]}}

    每项新增字段：
      "AI是否正确": "✔正确" | "❗错误"
      "AI错误数量": int
      "AI错误类型": str   逗号分隔
      "AI错误详情": str   JSON字符串
      "一致性":     "✅一致" | "❌不一致"
      "差异类型":   "一致" | "漏检（FN）" | "误报（FP）"
    """
    result = []
    for idx, row in enumerate(rows):
        ai     = ai_map.get(idx, {"is_correct": True, "errors": [], "source_issues": []})
        errors = ai.get("errors", [])
        source_issues = ai.get("source_issues", [])
        ai_ok  = ai.get("is_correct", True)
        ai_pending = ai.get("_pending", False)
        rule_ok = row.get("是否错误", "✔正确") != "❗错误"

        if rule_ok == ai_ok:
            consistency = "✅一致"
            diff_type   = "一致"
        else:
            consistency = "❌不一致"
            diff_type   = "漏检（FN）" if rule_ok and not ai_ok else "误报（FP）"

        merged = dict(row)
        ai_change_details = "；".join(
            f"'{e.get('替换锚点', '') or e.get('译文数值', '')}' → '{e.get('译文修改建议值', '')}'"
            for e in errors
            if e.get("替换锚点") or e.get("译文数值") or e.get("译文修改建议值")
        )
        merged.update({
            "AI是否正确": "⚠️待确认" if ai_pending else ("✔正确" if ai_ok else "❗错误"),
            "AI错误数量": len(errors),
            "AI错误类型": ",".join({e.get("错误类型", "") for e in errors}),
            "AI错误详情": "；".join(e.get("修改理由", "") for e in errors if e.get("修改理由")),
            "AI修改详情": ai_change_details,
            "原文问题数量": len(source_issues),
            "原文问题详情": json.dumps(source_issues, ensure_ascii=False) if source_issues else "",
            "一致性":     consistency,
            "差异类型":   diff_type,
        })
        result.append(merged)

    return result
