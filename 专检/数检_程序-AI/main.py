import json
import os
from pathlib import Path
from report_generator import generate_combined_report
from extract_values import merge_ai_results, run_number_check, check_text_pairs
from program_check import _llm_check_block
from normalizer import DEFAULT_STRATEGIES as _DEFAULT_STRATEGIES_BASIC
from normalizer_total import DEFAULT_STRATEGIES as _DEFAULT_STRATEGIES_TOTAL


# =============================================================
# 格式检测
# =============================================================
def _detect_format(path: str) -> str:
    """根据文件扩展名返回格式标识：'docx' | 'xlsx' | 'pptx' | 'pdf' | 'unknown'"""
    if not path:
        return "unknown"
    ext = Path(path).suffix.lower()
    return {
        ".docx": "docx",
        ".doc":  "docx",
        ".xlsx": "xlsx",
        ".xls":  "xlsx",
        ".pptx": "pptx",
        ".ppt":  "pptx",
        ".pdf":  "pdf",
    }.get(ext, "unknown")


# =============================================================
# 原文/译文直接提取对照对
# =============================================================

def _build_pairs_from_docx(src_file: str, tgt_file: str) -> list:
    """
    从原文和译文 DOCX 直接提取文本对，返回 [(src_text, tgt_text), ...]。

    对齐策略：
    - 总段落数一致 → 逐行对比 source 标签，发现标签不同时报告位置并抛出异常
    - 总段落数不一致 → 同上，先逐行找分叉点再报错
    """
    from full_content import extract_docx_in_order

    _LABEL = {
        "body": "正文", "table": "表格", "header": "页眉",
        "footer": "页脚", "footnote": "脚注", "chart": "图表",
    }

    print("📖 提取原文...")
    src_segs = extract_docx_in_order(src_file)
    print("📖 提取译文...")
    tgt_segs = extract_docx_in_order(tgt_file)
    print(f"  原文片段: {len(src_segs)}  译文片段: {len(tgt_segs)}")

    # 逐行对比 source 标签，找第一个不一致的位置
    min_len = min(len(src_segs), len(tgt_segs))
    for i in range(min_len):
        s, t = src_segs[i], tgt_segs[i]
        if s.source != t.source:
            src_lbl = _LABEL.get(s.source, s.source)
            tgt_lbl = _LABEL.get(t.source, t.source)
            preview = s.text.strip()[:20]
            raise ValueError(
                f"原文与译文结构不符，请上传对照文件（alignment_path）。\n"
                f"  第 {i + 1} 行：原文【{src_lbl}】vs 译文【{tgt_lbl}】\n"
                f"  原文内容: 「{preview}」"
            )

    # source 标签全部一致但数量不同（一侧更长）
    if len(src_segs) != len(tgt_segs):
        longer = "原文" if len(src_segs) > len(tgt_segs) else "译文"
        raise ValueError(
            f"原文与译文片段数量不一致，请上传对照文件（alignment_path）。\n"
            f"  原文片段: {len(src_segs)}  译文片段: {len(tgt_segs)}\n"
            f"  {longer}在第 {min_len + 1} 行起多出内容"
        )

    return [(s.text, t.text, t.para_index, t.row_context) for s, t in zip(src_segs, tgt_segs)]



def _build_pairs(src_file: str, tgt_file: str) -> list:
    """
    多格式通用版本：从原文和译文文件提取文本对，返回 [(src_text, tgt_text), ...]。
    支持 .docx / .xlsx / .pdf / .pptx，DOCX 走原有逻辑，其余格式走 extract_any。
    """
    fmt = _detect_format(src_file)

    if fmt == "docx":
        return _build_pairs_from_docx(src_file, tgt_file)

    from extract_any import extract, align_pdf_segments

    print("📖 提取原文...")
    src_segs = extract(src_file)
    print("📖 提取译文...")
    tgt_segs = extract(tgt_file)
    print(f"  原文片段: {len(src_segs)}  译文片段: {len(tgt_segs)}")

    if fmt in ("pdf", "pptx"):
        # 按页分组对齐
        raw_pairs = align_pdf_segments(src_segs, tgt_segs)
        return [
            (s.text if s else "", t.text if t else "",
             t.para_index if t else -1, t.row_context if t else "")
            for s, t in raw_pairs
        ]
    else:
        # xlsx 等：按顺序位置对齐
        max_len = max(len(src_segs), len(tgt_segs))
        return [
            (src_segs[i].text if i < len(src_segs) else "",
             tgt_segs[i].text if i < len(tgt_segs) else "",
             tgt_segs[i].para_index if i < len(tgt_segs) else -1,
             tgt_segs[i].row_context if i < len(tgt_segs) else "")
            for i in range(max_len)
        ]


# =============================================================
# 修订写入（统一入口）
# =============================================================

def apply_revisions_from_ai_map(
        final_rows: list,
        ai_map: dict,
        docx: object,          # 已打开的 Document 对象
        revision_manager: object,
        region: str = "body",
        doc_path: str = None,
) -> list:
    """
    从 ai_map 收集所有待修订任务并执行，返回 task 列表（供调用方汇总打印）。
    不负责 doc.save()，由调用方统一保存。
    """
    from replace_revision import replace_and_revise_in_docx
    from clean_replace_duplicates import clean_suggestion

    tasks = []
    skipped = []

    for idx, _ in enumerate(final_rows):
        ai = ai_map.get(idx, {})
        for err_no, err in enumerate(ai.get("errors", []), 1):
            # is_source_consistent=true → 译文忠实还原了原文，不修订
            isc = err.get("is_source_consistent")
            if isc is True or (isinstance(isc, str) and isc.lower() == "true"):
                skipped.append((idx, err_no, "原译一致，疑似原文问题", err)); continue
            old_val = err.get("替换锚点", "").strip() or err.get("译文数值", "").strip()
            new_val = err.get("译文修改建议值", "").strip()
            context = err.get("译文上下文", "").strip()
            anchor  = err.get("替换锚点", "").strip()
            reason  = err.get("修改理由", "") or err.get("错误类型", "")

            # 从 final_rows 取段落索引和行上下文（模式B时有值，模式A时为 -1/""）
            row_data    = final_rows[idx] if idx < len(final_rows) else {}
            para_index  = row_data.get("para_index", -1)
            row_context = row_data.get("row_context", "")
            # 表格行：用行上下文增强 context 区分度
            if row_context and not context:
                context = row_context
            elif row_context and row_context not in context:
                context = row_context  # 行上下文比单元格上下文更有区分度，优先使用
            if not old_val:
                skipped.append((idx, err_no, "锚点/译文数值为空", err)); continue
            if not new_val:
                skipped.append((idx, err_no, "修改建议值为空", err)); continue
            if old_val == new_val:
                skipped.append((idx, err_no, f"old==new: '{old_val}'", err)); continue

            # 去重：检测 new_val 与上下文前后是否有重叠，剪裁后写回
            cleaned_new_val = clean_suggestion(context, old_val, new_val)
            if cleaned_new_val != new_val:
                print(f"  [去重] idx={idx} '{new_val}' → '{cleaned_new_val}'")
            if not cleaned_new_val or cleaned_new_val == old_val:
                skipped.append((idx, err_no, f"去重后为空或等于old: '{cleaned_new_val}'", err)); continue
            new_val = cleaned_new_val

            # 取前后句译文，用于夹逼定位重复文本
            prev_tgt = (final_rows[idx - 1].get("译文", "") if idx > 0 else "")
            next_tgt = (final_rows[idx + 1].get("译文", "") if idx + 1 < len(final_rows) else "")

            tasks.append((idx, err_no, old_val, new_val, context, anchor, reason, para_index, prev_tgt, next_tgt))

    if skipped:
        print(f"  [过滤] {len(skipped)} 条（字段缺失或无变化）")

    # 长文本优先：避免短锚点先替换后，长锚点找不到原文的冲突
    tasks.sort(key=lambda t: len(t[2]), reverse=True)

    # 计算每个 old_val 的出现序号（按任务列表顺序，即文档顺序的近似）
    # occurrence_counter[old_val] = 下一次出现时应取的序号
    from collections import defaultdict
    _occ_counter: dict = defaultdict(int)
    tasks_with_occ = []
    for task in tasks:
        idx, err_no, old_val, new_val, context, anchor, reason, para_index, prev_tgt, next_tgt = task
        occ = _occ_counter[old_val]
        _occ_counter[old_val] += 1
        tasks_with_occ.append((idx, err_no, old_val, new_val, context, anchor, reason, para_index, occ, prev_tgt, next_tgt))

    success, failed = 0, 0
    for i, (idx, err_no, old_val, new_val, context, anchor, reason, para_index, occ, prev_tgt, next_tgt) in enumerate(tasks_with_occ, 1):
        ok, strategy = replace_and_revise_in_docx(
            doc=docx, old_value=old_val, new_value=new_val, reason=reason,
            revision_manager=revision_manager, context=context, anchor_text=anchor,
            region=region, doc_path=doc_path, para_cache=None,
            target_para_idx=para_index,
            occurrence_index=occ,
            prev_tgt=prev_tgt,
            next_tgt=next_tgt,
        )
        if ok:
            success += 1
            if "批注兜底" in strategy:
                print(f"  📌 [{i:>2}/{len(tasks)}] '{old_val}' → '{new_val}'  ({strategy})")
            else:
                print(f"  ✅ [{i:>2}/{len(tasks)}] '{old_val}' → '{new_val}'  ({strategy})")
        else:
            failed += 1
            print(f"  ⚠️  [{i:>2}/{len(tasks)}] 未找到: '{old_val}'  ({strategy})")

    return tasks, success, failed


# =============================================================
# Excel 修订写入
# =============================================================

def apply_revisions_excel(
        final_rows: list,
        ai_map: dict,
        replacer,           # ExcelReplacer 实例
) -> tuple:
    """从 ai_map 收集修订任务并写入 Excel（批注形式），返回 (tasks, success, failed)。"""
    from clean_replace_duplicates import clean_suggestion
    tasks, skipped = [], []

    for idx, _ in enumerate(final_rows):
        ai = ai_map.get(idx, {})
        for err_no, err in enumerate(ai.get("errors", []), 1):
            isc = err.get("is_source_consistent")
            if isc is True or (isinstance(isc, str) and isc.lower() == "true"):
                skipped.append((idx, err_no, err)); continue
            old_val = err.get("替换锚点", "").strip() or err.get("译文数值", "").strip()
            new_val = err.get("译文修改建议值", "").strip()
            context = err.get("译文上下文", "").strip()
            reason  = err.get("修改理由", "") or err.get("错误类型", "")
            if not old_val or not new_val or old_val == new_val:
                skipped.append((idx, err_no, err)); continue
            cleaned = clean_suggestion(context, old_val, new_val)
            if cleaned != new_val:
                print(f"  [去重] idx={idx} '{new_val}' → '{cleaned}'")
            if not cleaned or cleaned == old_val:
                skipped.append((idx, err_no, err)); continue
            prev_tgt = (final_rows[idx - 1].get("译文", "") if idx > 0 else "")
            next_tgt = (final_rows[idx + 1].get("译文", "") if idx + 1 < len(final_rows) else "")
            tasks.append((idx, err_no, old_val, cleaned, context, reason, prev_tgt, next_tgt))

    if skipped:
        print(f"  [过滤] {len(skipped)} 条（字段缺失或无变化或原译一致）")

    # 长文本优先：避免短锚点先替换后，长锚点找不到原文的冲突
    tasks.sort(key=lambda t: len(t[2]), reverse=True)

    success, failed = 0, 0
    for i, (idx, err_no, old_val, new_val, context, reason, prev_tgt, next_tgt) in enumerate(tasks, 1):
        ok = replacer.replace_and_annotate(
            old_text=old_val, new_text=new_val,
            reason=reason, context=context, highlight=True,
            prev_tgt=prev_tgt, next_tgt=next_tgt,
        )
        if ok:
            success += 1
            print(f"  ✅ [{i:>2}/{len(tasks)}] '{old_val}' → '{new_val}'")
        else:
            failed += 1
            print(f"  ⚠️  [{i:>2}/{len(tasks)}] 未找到: '{old_val}'")

    return tasks, success, failed


# =============================================================
# PPTX 修订写入
# =============================================================

def apply_revisions_pptx(
        final_rows: list,
        ai_map: dict,
        replacer,           # PPTXReplacer 实例
) -> tuple:
    """从 ai_map 收集修订任务并写入 PPTX（批注形式），返回 (tasks, success, failed)。"""
    from clean_replace_duplicates import clean_suggestion
    tasks, skipped = [], []

    for idx, _ in enumerate(final_rows):
        ai = ai_map.get(idx, {})
        for err_no, err in enumerate(ai.get("errors", []), 1):
            isc = err.get("is_source_consistent")
            if isc is True or (isinstance(isc, str) and isc.lower() == "true"):
                skipped.append((idx, err_no, err)); continue
            old_val = err.get("替换锚点", "").strip() or err.get("译文数值", "").strip()
            new_val = err.get("译文修改建议值", "").strip()
            context = err.get("译文上下文", "").strip()
            reason  = err.get("修改理由", "") or err.get("错误类型", "")
            if not old_val or not new_val or old_val == new_val:
                skipped.append((idx, err_no, err)); continue
            cleaned = clean_suggestion(context, old_val, new_val)
            if cleaned != new_val:
                print(f"  [去重] idx={idx} '{new_val}' → '{cleaned}'")
            if not cleaned or cleaned == old_val:
                skipped.append((idx, err_no, err)); continue
            tasks.append((idx, err_no, old_val, cleaned, context, reason))

    if skipped:
        print(f"  [过滤] {len(skipped)} 条（字段缺失或无变化或原译一致）")

    # 长文本优先：避免短锚点先替换后，长锚点找不到原文的冲突
    tasks.sort(key=lambda t: len(t[2]), reverse=True)

    success, failed = 0, 0
    for i, (idx, err_no, old_val, new_val, context, reason) in enumerate(tasks, 1):
        ok = replacer.replace_and_annotate(
            old_text=old_val, new_text=new_val,
            reason=reason, context=context, highlight=True,
        )
        if ok:
            success += 1
            print(f"  ✅ [{i:>2}/{len(tasks)}] '{old_val}' → '{new_val}'")
        else:
            failed += 1
            print(f"  ⚠️  [{i:>2}/{len(tasks)}] 未找到: '{old_val}'")

    return tasks, success, failed


# =============================================================
# PDF 修订写入
# =============================================================

def apply_revisions_pdf(
        final_rows: list,
        ai_map: dict,
        replacer,           # ImprovedPDFReplacer 实例
) -> tuple:
    """从 ai_map 收集修订任务并写入 PDF（批注 + 文本覆盖），返回 (tasks, success, failed)。"""
    from clean_replace_duplicates import clean_suggestion
    tasks, skipped = [], []

    for idx, _ in enumerate(final_rows):
        ai = ai_map.get(idx, {})
        for err_no, err in enumerate(ai.get("errors", []), 1):
            isc = err.get("is_source_consistent")
            if isc is True or (isinstance(isc, str) and isc.lower() == "true"):
                skipped.append((idx, err_no, err)); continue
            old_val = err.get("替换锚点", "").strip() or err.get("译文数值", "").strip()
            new_val = err.get("译文修改建议值", "").strip()
            context = err.get("译文上下文", "").strip()
            reason  = err.get("修改理由", "") or err.get("错误类型", "")
            if not old_val or not new_val or old_val == new_val:
                skipped.append((idx, err_no, err)); continue
            cleaned = clean_suggestion(context, old_val, new_val)
            if cleaned != new_val:
                print(f"  [去重] idx={idx} '{new_val}' → '{cleaned}'")
            if not cleaned or cleaned == old_val:
                skipped.append((idx, err_no, err)); continue
            prev_tgt = (final_rows[idx - 1].get("译文", "") if idx > 0 else "")
            next_tgt = (final_rows[idx + 1].get("译文", "") if idx + 1 < len(final_rows) else "")
            tasks.append((idx, err_no, old_val, cleaned, context, reason, prev_tgt, next_tgt))

    if skipped:
        print(f"  [过滤] {len(skipped)} 条（字段缺失或无变化或原译一致）")

    success, failed = 0, 0
    for i, (idx, err_no, old_val, new_val, context, reason, prev_tgt, next_tgt) in enumerate(tasks, 1):
        comment = f"[修改] '{old_val}' → '{new_val}'"
        if reason:
            comment += f"\n理由: {reason}"
        rep_count, _, _ = replacer.replace_and_annotate(
            search_text=old_val, new_text=new_val,
            comment=comment, context=context,
            prev_tgt=prev_tgt, next_tgt=next_tgt,
        )
        if rep_count:
            success += 1
            print(f"  ✅ [{i:>2}/{len(tasks)}] '{old_val}' → '{new_val}'")
        else:
            failed += 1
            print(f"  ⚠️  [{i:>2}/{len(tasks)}] 未找到: '{old_val}'")

    return tasks, success, failed


# =============================================================
# 单区域检查流程（规则 + AI）→ 返回 (final_rows, ai_map)
# =============================================================

def _check_region(label: str, pairs: list,
                  block_size: int, normalize_strategies: dict,
                  check_all: bool = False) -> tuple:
    """
    对 [(src, tgt), ...] 做规则检查 + AI 复核。

    check_all=True  → 全部行送 AI（页眉/页脚段落少）
    check_all=False → 只对规则错误行送 AI（正文）

    返回 (final_rows, ai_map)
    """
    rows = check_text_pairs(pairs, strategies=normalize_strategies)
    rule_err = sum(1 for r in rows if r["是否错误"] == "❗错误")
    print(f"  [{label}] 规则错误: {rule_err} / {len(rows)}")

    if check_all:
        candidates = list(enumerate(rows))
    else:
        candidates = [(i, r) for i, r in enumerate(rows) if r["是否错误"] == "❗错误"]

    blocks = [candidates[i:i + block_size] for i in range(0, len(candidates), block_size)]
    print(f"  [{label}] AI复核 {len(candidates)} 行，共 {len(blocks)} 块...")

    ai_map = {}
    for b_idx, block in enumerate(blocks):
        print(f"    Block {b_idx + 1}/{len(blocks)}")
        results = _llm_check_block(block)
        for pos, (idx, _) in enumerate(block):
            ai_map[idx] = results[pos]

    final_rows = merge_ai_results(rows, ai_map)
    ai_err = sum(1 for r in final_rows if r.get("AI是否正确") == "❗错误")
    print(f"  [{label}] AI错误: {ai_err} / {len(final_rows)}")
    return final_rows, ai_map


# =============================================================
# 对照数据提取（页眉/页脚）
# =============================================================

def _align_hf(src_list: list, tgt_list: list, label: str) -> list:
    """
    将原文和译文的页眉/页脚列表按位置配对，返回 [(src, tgt), ...]。
    长度不一致时短的一侧补空字符串，并打印警告。
    """
    if len(src_list) != len(tgt_list):
        print(f"  [{label}] ⚠️  原文 {len(src_list)} 段 vs 译文 {len(tgt_list)} 段，按位置对齐，短侧补空")
    max_len = max(len(src_list), len(tgt_list), 1)
    src_pad = src_list + [""] * (max_len - len(src_list))
    tgt_pad = tgt_list + [""] * (max_len - len(tgt_list))
    return list(zip(src_pad, tgt_pad))


def _save_alignment_excel(pairs: list, output_dir: str, stem: str):
    """将文本对保存为对照 Excel，方便后续复用。"""
    try:
        import pandas as pd
        path = os.path.join(output_dir, f"{stem}_alignment.xlsx")
        df = pd.DataFrame([{"原文": p[0], "译文": p[1]} for p in pairs])
        df.to_excel(path, index=False)
        print(f"  💾 对照Excel已保存: {path}")
    except Exception as e:
        print(f"  ⚠️  对照Excel保存失败: {e}")


# =============================================================
# 主流程
# =============================================================

def run(alignment_path: str = None,
        output_dir: str = "output",
        block_size: int = 20,
        normalize_strategies: dict = None,
        src_docx_path: str = None,
        tgt_docx_path: str = None,
        src_hf_path: str = None,
        docx_path: str = None,
        revised_docx_path: str = None,
        revision_author: str = "翻译校对",
        check_header: bool = None,
        check_footer: bool = None,
        force_mode_b: bool = False,
        ai_check_all: bool = False,
        use_total_normalizer: bool = False):
    """
    两种输入模式（二选一，也可由格式自动推断）：

      模式A — 上传对照文件（PDF 默认此模式，其他格式也可手动指定）:
        alignment_path : 已对齐的原文/译文 Excel（含"原文"/"译文"列）
        src_hf_path    : 原文 .docx 路径，用于页眉/页脚检查（可选）
        docx_path      : 译文 .docx 路径，用于页眉/页脚检查 + 修订写入（可选）

      模式B — 直接上传原文+译文（docx/xlsx/pptx 默认此模式）:
        src_docx_path  : 原文文件（支持 .docx / .xlsx / .pdf / .pptx）
        tgt_docx_path  : 译文文件（同时作为修订写入目标，可被 docx_path 覆盖）
        force_mode_b   : True 时强制走模式B直接提取，即使格式为 PDF

    公共参数：
      output_dir       : 报告输出目录
      docx_path        : 修订写入目标（模式B 时若不填则默认使用 tgt_docx_path）
      revised_docx_path: 修订输出路径；None 时自动备份到 backup/ 子文件夹
      check_header     : 是否检查页眉（None 时自动推断：仅 docx 格式开启）
      check_footer     : 是否检查页脚（None 时自动推断：仅 docx 格式开启）
      ai_check_all     : True 时正文全量送 AI（含规则认为正确的行），默认 False 仅送规则错误行
      use_total_normalizer : True 时使用 normalizer_total（更全面的规范化策略），
                             默认 False 使用 normalizer（基础策略）
    """
    # ── 策略选择 ──────────────────────────────────────────────────────
    if normalize_strategies is None:
        normalize_strategies = _DEFAULT_STRATEGIES_TOTAL if use_total_normalizer else _DEFAULT_STRATEGIES_BASIC
    _normalizer_label = "normalizer_total" if use_total_normalizer else "normalizer"
    print(f"  [归化器] 使用 {_normalizer_label}")

    # ── 模式推断 ──────────────────────────────────────────────────────
    # 优先级：alignment_path（模式A）> src+tgt（模式B）
    # 两者都传时：用 alignment_path 做检查，src/tgt 用于修订写入和页眉页脚
    # PDF 默认模式A；其他格式有 src+tgt 且无 alignment_path 时走模式B
    _src_fmt = _detect_format(src_docx_path) if src_docx_path else "unknown"
    if not alignment_path and src_docx_path and tgt_docx_path:
        if _src_fmt == "pdf" and not force_mode_b:
            raise ValueError(
                "PDF 格式默认使用模式A（对照文件），请提供 alignment_path。\n"
                "如需直接从 PDF 提取对照，请传入 force_mode_b=True。"
            )

    if not alignment_path and not (src_docx_path and tgt_docx_path):
        raise ValueError("请提供 alignment_path（对照Excel）或同时提供 src_docx_path + tgt_docx_path")

    # 模式B：tgt_docx_path 作为修订目标的默认值
    if not docx_path and tgt_docx_path:
        docx_path = tgt_docx_path

    # 自动生成输出 xlsx 路径
    os.makedirs(output_dir, exist_ok=True)
    if alignment_path:
        align_stem = Path(alignment_path).stem
    else:
        align_stem = Path(tgt_docx_path).stem
    output_path = os.path.join(output_dir, f"{align_stem}_output.xlsx")

    fmt = _detect_format(docx_path) if docx_path else "unknown"

    # check_header/check_footer 默认只对 docx 开启
    if check_header is None:
        check_header = (fmt == "docx")
    if check_footer is None:
        check_footer = (fmt == "docx")

    # ── 阶段1：正文检查 ──────────────────────────────────────────────
    print("\n" + "="*60)
    print("📄 正文检查")
    print("="*60)

    if alignment_path:
        # 模式A：从对照 Excel 读取
        print(f"  [模式A] 读取对照文件: {Path(alignment_path).name}")
        body_rows = run_number_check(alignment_path, strategies=normalize_strategies)

        # 检测对照文件原文/译文是否都有内容
        empty_src = sum(1 for r in body_rows if not str(r.get("原文", "")).strip())
        empty_tgt = sum(1 for r in body_rows if not str(r.get("译文", "")).strip())
        if empty_src or empty_tgt:
            lines = []
            if empty_src:
                lines.append(f"原文为空: {empty_src} 行")
            if empty_tgt:
                lines.append(f"译文为空: {empty_tgt} 行")
            print(f"  ⚠️  对照文件存在空行（{'，'.join(lines)}），可能影响检查准确性")
        # 保存对照 JSON
        _pairs_for_json = [{"原文": r["原文"], "译文": r["译文"]} for r in body_rows]
        _align_json_path = os.path.join(output_dir, f"{align_stem}_alignment.json")
        with open(_align_json_path, "w", encoding="utf-8") as _f:
            json.dump(_pairs_for_json, _f, ensure_ascii=False, indent=2)
        print(f"  💾 对照JSON: {_align_json_path}")
    else:
        # 模式B：从原文/译文文件直接提取（支持 docx/xlsx/pdf/pptx）
        print(f"  [模式B] 直接提取原文/译文对照")
        try:
            pairs = _build_pairs(src_docx_path, tgt_docx_path)
        except ValueError as e:
            raise ValueError(
                f"{e}\n\n"
                f"💡 提示：请先制作对照文件，然后通过 alignment_path 参数传入：\n"
                f"   run(alignment_path=r\"...\", docx_path=r\"{tgt_docx_path}\")"
            ) from None
        body_rows = check_text_pairs(pairs, strategies=normalize_strategies)
        # 把 para_index / row_context 写入每个 row，供写回阶段精确定位
        for i, row in enumerate(body_rows):
            if i < len(pairs) and len(pairs[i]) >= 4:
                row["para_index"]  = pairs[i][2]
                row["row_context"] = pairs[i][3]
        # 保存对照 Excel 和 JSON 供后续复用
        _save_alignment_excel(pairs, output_dir, align_stem)
        _align_json_path = os.path.join(output_dir, f"{align_stem}_alignment.json")
        with open(_align_json_path, "w", encoding="utf-8") as _f:
            json.dump([{"原文": p[0], "译文": p[1]} for p in pairs], _f, ensure_ascii=False, indent=2)
        print(f"  💾 对照JSON: {_align_json_path}")

    rule_err = sum(1 for r in body_rows if r["是否错误"] == "❗错误")
    print(f"  [正文] 规则错误: {rule_err} / {len(body_rows)}")

    candidates = (list(enumerate(body_rows)) if ai_check_all
                  else [(i, r) for i, r in enumerate(body_rows) if r["是否错误"] == "❗错误"])
    mode_label = "全量" if ai_check_all else "仅规则错误行"
    blocks = [candidates[i:i + block_size] for i in range(0, len(candidates), block_size)]
    print(f"  [正文] AI复核 {len(candidates)} 行（{mode_label}），共 {len(blocks)} 块...")

    body_ai_map = {}
    for b_idx, block in enumerate(blocks):
        print(f"    Block {b_idx + 1}/{len(blocks)}")
        results = _llm_check_block(block)
        for pos, (idx, _) in enumerate(block):
            body_ai_map[idx] = results[pos]

    body_final = merge_ai_results(body_rows, body_ai_map)
    body_ai_err = sum(1 for r in body_final if r.get("AI是否正确") == "❗错误")
    print(f"  [正文] AI错误: {body_ai_err} / {len(body_final)}")

    # ── 阶段2：页眉/页脚检查（仅 docx）────────────────────────────
    header_final, header_ai_map = [], {}
    footer_final, footer_ai_map = [], {}
    header_pairs, footer_pairs = [], []

    # 页眉/页脚检查需要：docx 格式 + 原文路径 + 译文路径
    # 模式B：src_docx_path / tgt_docx_path
    # 模式A：src_hf_path（新增）/ docx_path
    _hf_src = src_docx_path or src_hf_path
    _hf_tgt = tgt_docx_path or docx_path

    if fmt == "docx" and _hf_src and _hf_tgt and (check_header or check_footer):
        from header_extractor import extract_headers
        from footer_extractor import extract_footers

        if check_header:
            print("\n" + "="*60)
            print("📄 页眉检查")
            print("="*60)
            src_h = extract_headers(_hf_src)
            tgt_h = extract_headers(_hf_tgt)
            print(f"  [页眉] 原文 {len(src_h)} 段 / 译文 {len(tgt_h)} 段")
            header_pairs = _align_hf(src_h, tgt_h, "页眉")
            header_final, header_ai_map = _check_region(
                "页眉", header_pairs, block_size, normalize_strategies, check_all=True)

        if check_footer:
            print("\n" + "="*60)
            print("📄 页脚检查")
            print("="*60)
            src_f = extract_footers(_hf_src)
            tgt_f = extract_footers(_hf_tgt)
            print(f"  [页脚] 原文 {len(src_f)} 段 / 译文 {len(tgt_f)} 段")
            footer_pairs = _align_hf(src_f, tgt_f, "页脚")
            footer_final, footer_ai_map = _check_region(
                "页脚", footer_pairs, block_size, normalize_strategies, check_all=True)

    # ── 阶段3：保存 AI 原始结果 JSON（供调试定位）─────────────────────
    def _save_json(data, filename):
        path = os.path.join(output_dir, filename)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"  💾 AI原始JSON: {path}")

    def _ai_map_to_list(rows, ai_map):
        """将 ai_map 转为带 seq/原文/译文 的列表，方便定位行号"""
        out = []
        for idx, row in enumerate(rows):
            entry = ai_map.get(idx, {"is_correct": True, "errors": [], "source_issues": []})
            out.append({
                "seq": idx,
                "原文": row.get("原文", ""),
                "译文": row.get("译文", ""),
                **entry,
            })
        return out

    _save_json(_ai_map_to_list(body_rows, body_ai_map), "align_body.json")

    # 额外保存带完整 errors 数组的原始 ai_map，供 apply_from_result.py 使用
    def _ai_map_to_errors_list(rows, ai_map):
        out = []
        for idx, row in enumerate(rows):
            entry = ai_map.get(idx)
            if not entry:
                continue
            errors = entry.get("errors", [])
            if not errors:
                continue
            out.append({
                "seq": idx,
                "原文": row.get("原文", ""),
                "译文": row.get("译文", ""),
                "errors": errors,
            })
        return out

    _save_json(_ai_map_to_errors_list(body_rows, body_ai_map), "align_body_errors.json")

    # 额外保存扁平化的 errors 列表（_ERROR_SCHEMA 格式），供快速测试直接使用
    def _ai_map_to_flat_errors(rows, ai_map):
        """
        将 ai_map 中每条 error 展开为独立条目，字段与 _ERROR_SCHEMA 一一对应。
        输出的 JSON 可直接复制到 test_quick_replace_number.py 的 TEST_ERRORS 中使用。
        """
        out = []
        global_err_no = 1
        for idx, row in enumerate(rows):
            entry = ai_map.get(idx)
            if not entry:
                continue
            for err in entry.get("errors", []):
                out.append({
                    "错误编号": str(global_err_no),
                    "原文上下文": err.get("原文上下文", row.get("原文", "")),
                    "译文上下文": err.get("译文上下文", row.get("译文", "")),
                    "原文数值":   err.get("原文数值", ""),
                    "译文数值":   err.get("译文数值", ""),
                    "替换锚点":   err.get("替换锚点", ""),
                    "译文修改建议值": err.get("译文修改建议值", ""),
                    "is_source_consistent": err.get("is_source_consistent", False),
                    "错误类型":   err.get("错误类型", ""),
                    "修改理由":   err.get("修改理由", ""),
                    "违反的规则": err.get("违反的规则", ""),
                })
                global_err_no += 1
        return out

    _save_json(_ai_map_to_flat_errors(body_rows, body_ai_map), "align_body_flat_errors.json")
    print(f"  💾 扁平errors JSON（可直接用于快速测试）已保存")
    if header_final:
        _save_json(_ai_map_to_list(
            [{"原文": s, "译文": t} for s, t in header_pairs], header_ai_map
        ), "align_header.json")
    if footer_final:
        _save_json(_ai_map_to_list(
            [{"原文": s, "译文": t} for s, t in footer_pairs], footer_ai_map
        ), "align_footer.json")

    # ── 阶段4：Excel 报告（正文主 sheet + 页眉/页脚追加 sheet）────────
    print("\n" + "="*60)
    print("📊 生成报告")
    print("="*60)
    generate_combined_report(body_final, output_path)

    def _append_sheet(rows, sheet_name):
        if not rows:
            return
        try:
            import pandas as pd
            df = pd.DataFrame(rows)
            with pd.ExcelWriter(output_path, engine="openpyxl", mode="a",
                                if_sheet_exists="replace") as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"  ✅ sheet [{sheet_name}] 写入完成")
        except Exception as e:
            print(f"  ⚠️  sheet [{sheet_name}] 写入失败: {e}")

    _append_sheet(header_final, "页眉")
    _append_sheet(footer_final, "页脚")

    # ── 阶段5：合并所有 ai_map，统一写修订 ───────────────────────────
    if docx_path:
        print("\n" + "="*60)
        print(f"📝 写入修订（格式: {fmt.upper()}）")
        print("="*60)

        # 确定输出路径：优先用指定路径，否则自动备份到 backup/ 子目录
        if revised_docx_path:
            out_path = revised_docx_path
        else:
            from datetime import datetime as _dt
            _stem, _ext = os.path.splitext(os.path.basename(docx_path))
            _ts = _dt.now().strftime("%Y%m%d_%H%M%S")
            out_path = os.path.join(output_dir, f"{_stem}_{_ts}{_ext}")
            import shutil as _shutil
            _shutil.copy2(docx_path, out_path)
            print(f"✓ 输出文件: {out_path}")

        total_success, total_failed = 0, 0

        # ── DOCX：Track Changes 修订 ──────────────────────────────────
        if fmt == "docx":
            from docx import Document
            from revision import RevisionManager

            doc = Document(docx_path)
            rm = RevisionManager(doc, author=revision_author)

            # 正文
            print(f"\n── 正文 ({len(body_final)} 行) ──")
            total_errors = sum(len(v.get("errors", [])) for v in body_ai_map.values())
            print(f"  [诊断] ai_map 覆盖行数: {len(body_ai_map)}  |  errors 对象总数: {total_errors}")
            _, s, f = apply_revisions_from_ai_map(
                body_final, body_ai_map, doc, rm, region="body", doc_path=docx_path)
            total_success += s; total_failed += f

            # 页眉
            if header_final and header_ai_map:
                print(f"\n── 页眉 ({len(header_final)} 段) ──")
                _, s, f = apply_revisions_from_ai_map(
                    header_final, header_ai_map, doc, rm, region="header", doc_path=docx_path)
                total_success += s; total_failed += f

            # 页脚
            if footer_final and footer_ai_map:
                print(f"\n── 页脚 ({len(footer_final)} 段) ──")
                _, s, f = apply_revisions_from_ai_map(
                    footer_final, footer_ai_map, doc, rm, region="footer", doc_path=docx_path)
                total_success += s; total_failed += f

            doc.save(out_path)

            try:
                from replace_revision import flush_footnote_replacements
                flushed = flush_footnote_replacements(doc, out_path)
                if flushed:
                    print(f"  📎 脚注替换: {flushed} 处")
            except Exception:
                pass

        # ── XLSX：批注替换 ────────────────────────────────────────────
        elif fmt == "xlsx":
            from excel.excel_replacer import ExcelReplacer

            replacer = ExcelReplacer(out_path)
            print(f"\n── 正文 ({len(body_final)} 行) ──")
            _, s, f = apply_revisions_excel(body_final, body_ai_map, replacer)
            total_success += s; total_failed += f
            replacer.save(out_path)

        # ── PPTX：批注替换 ────────────────────────────────────────────
        elif fmt == "pptx":
            from ppt.pptx_replacer import PPTXReplacer

            replacer = PPTXReplacer(out_path)
            print(f"\n── 正文 ({len(body_final)} 行) ──")
            _, s, f = apply_revisions_pptx(body_final, body_ai_map, replacer)
            total_success += s; total_failed += f
            replacer.save(out_path)

        # ── PDF：批注 + 文本覆盖 ──────────────────────────────────────
        elif fmt == "pdf":
            from pdf.pdf_replacer_improved import ImprovedPDFReplacer

            replacer = ImprovedPDFReplacer(out_path)
            print(f"\n── 正文 ({len(body_final)} 行) ──")
            _, s, f = apply_revisions_pdf(body_final, body_ai_map, replacer)
            total_success += s; total_failed += f
            replacer.save(out_path)

        else:
            print(f"  ⚠️  不支持的文件格式: {fmt}，跳过修订写入")

        print(f"\n{'='*60}")
        print(f"✅ 修订完成: 成功 {total_success}  失败 {total_failed}")
        print(f"   输出文件: {out_path}")
        print(f"{'='*60}")

    return body_final, header_final, footer_final


if __name__ == "__main__":
    # 三个核心参数：
    #   src_docx_path  : 原文文件（必填）
    #   tgt_docx_path  : 译文文件（必填）
    #   alignment_path : 对照文件（必填）
    #   ai_check_all=False :是否开启ai全量检查；默认关闭。
    #   use_total_normalizer=False : 默认使用 normalizer（基础策略）；
    #                                True 时切换为 normalizer_total（支持货币/日期/分数/季度等更多规则）

    # ── 用法：传原文+译文+对照文件，对齐检查────────────
    run(
        src_docx_path=r"测试文件\01 27 XXX2025年度审计情况汇报 （可编辑版本）  - GT.pptx",
        tgt_docx_path=r"测试文件\预排版_01 27 XXX2025年度审计情况汇报 （可编辑版本） - GT.pptx",
        alignment_path=r"测试文件\bilingual_pairs (9).xlsx",
        output_dir=r"output",
        revision_author="数值检查",
        #ai_check_all=True,
        use_total_normalizer=True
    )

