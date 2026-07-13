"""
通用快速替换测试（支持 docx / xlsx / pptx / pdf）

用法：
    python quick_replace.py "译文.docx" "output/align_body_flat_errors.json"
    python quick_replace.py "译文.xlsx" "output/align_body_flat_errors.json"
    python quick_replace.py "译文.pptx" "output/align_body_flat_errors.json"
    python quick_replace.py "译文.pdf"  "output/align_body_flat_errors.json"
"""

import sys
import json
import difflib
from pathlib import Path

project_root = str(Path(__file__).resolve().parent)
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from backup_copy.backup_manager import ensure_backup_copy


# ── 格式检测 ──────────────────────────────────────────────────────────

def _detect_format(path: str) -> str:
    ext = Path(path).suffix.lower()
    return {".docx": "docx", ".doc": "docx",
            ".xlsx": "xlsx", ".xls": "xlsx",
            ".pptx": "pptx", ".ppt": "pptx",
            ".pdf":  "pdf"}.get(ext, "unknown")


# ── 重叠修正 ─────────────────────────────────────────────────────────

def _fix_overlap(context: str, anchor: str, suggestion: str) -> str:
    if not context or not anchor or not suggestion:
        return suggestion
    if anchor not in context:
        return suggestion

    start = context.find(anchor)
    end = start + len(anchor)
    actual = suggestion

    prefix = context[max(0, start - len(suggestion)):start]
    m = difflib.SequenceMatcher(None, prefix, actual).find_longest_match(
        0, len(prefix), 0, len(actual))
    if m.size > 0 and m.a + m.size == len(prefix) and m.b == 0:
        actual = actual[m.size:]

    after = context[end: end + len(actual)]
    m = difflib.SequenceMatcher(None, actual, after).find_longest_match(
        0, len(actual), 0, len(after))
    if m.size > 0 and m.a + m.size == len(actual) and m.b == 0:
        actual = actual[:m.a]

    return actual or suggestion


# ── DOCX ─────────────────────────────────────────────────────────────

def _run_docx(backup_path: str, tasks: list) -> list:
    from docx import Document
    from revision import RevisionManager
    from replace_revision import replace_and_revise_in_docx, flush_footnote_replacements
    from collections import defaultdict

    # 编号/目录静态化
    try:
        from numbering_to_static import convert_numbering_to_static, has_auto_numbering, convert_toc_to_static
        if has_auto_numbering(backup_path):
            print("🔢 检测到自动编号，转换为静态文本...")
            convert_numbering_to_static(backup_path)
        convert_toc_to_static(backup_path)
    except Exception as e:
        print(f"⚠ 编号/目录静态化: {e}")

    doc = Document(backup_path)
    doc._numbering_staticized = True
    rm = RevisionManager(doc, author="快速测试")

    # 计算每个 old_val 的出现序号（tasks 已按长锚点优先排序，此处顺序对应替换顺序）
    occ_counter: dict = defaultdict(int)

    results = []
    for error_no, old_val, new_val, reason, context, anchor, prev_tgt, next_tgt in tasks:
        occ = occ_counter[old_val]
        occ_counter[old_val] += 1
        try:
            ok, strategy = replace_and_revise_in_docx(
                doc, old_val, new_val, reason, rm,
                context=context, anchor_text=anchor,
                region="body", doc_path=backup_path,
                occurrence_index=occ,
                prev_tgt=prev_tgt,
                next_tgt=next_tgt,
            )
            results.append((error_no, "成功" if ok else "失败", strategy))
        except Exception as e:
            results.append((error_no, "异常", str(e)))

    doc.save(backup_path)
    flushed = flush_footnote_replacements(doc, backup_path)
    if flushed:
        print(f"  📎 脚注替换: {flushed} 处")

    return results


# ── XLSX ─────────────────────────────────────────────────────────────

def _run_xlsx(backup_path: str, tasks: list) -> list:
    from excel.excel_replacer import ExcelReplacer

    replacer = ExcelReplacer(backup_path)
    results = []
    for error_no, old_val, new_val, reason, context, anchor, prev_tgt, next_tgt in tasks:
        try:
            ok = replacer.replace_and_annotate(
                old_text=old_val, new_text=new_val,
                reason=reason, context=context, highlight=True,
            )
            results.append((error_no, "成功" if ok else "失败", ""))
        except Exception as e:
            results.append((error_no, "异常", str(e)))

    replacer.save(backup_path)
    return results


# ── PPTX ─────────────────────────────────────────────────────────────

def _run_pptx(backup_path: str, tasks: list) -> list:
    from ppt.pptx_replacer import PPTXReplacer

    replacer = PPTXReplacer(backup_path)
    results = []
    for error_no, old_val, new_val, reason, context, anchor, prev_tgt, next_tgt in tasks:
        try:
            ok = replacer.replace_and_annotate(
                old_text=old_val, new_text=new_val,
                reason=reason, context=context, highlight=True,
            )
            results.append((error_no, "成功" if ok else "失败", ""))
        except Exception as e:
            results.append((error_no, "异常", str(e)))

    replacer.save(backup_path)
    return results


# ── PDF ──────────────────────────────────────────────────────────────

def _run_pdf(backup_path: str, tasks: list) -> list:
    from pdf.pdf_replacer_improved import ImprovedPDFReplacer

    replacer = ImprovedPDFReplacer(backup_path)
    results = []
    for error_no, old_val, new_val, reason, context, anchor, prev_tgt, next_tgt in tasks:
        try:
            comment = f"[修改] '{old_val}' → '{new_val}'"
            if reason:
                comment += f"\n理由: {reason}"
            rep_count, _, _ = replacer.replace_and_annotate(
                search_text=old_val, new_text=new_val,
                comment=comment, context=context,
            )
            results.append((error_no, "成功" if rep_count else "失败", ""))
        except Exception as e:
            results.append((error_no, "异常", str(e)))

    replacer.save(backup_path)
    return results


# ── 主入口 ───────────────────────────────────────────────────────────

def quick_test(doc_path: str, errors_json: str = None, test_cases: list = None):
    fmt = _detect_format(doc_path)
    if fmt == "unknown":
        print(f"❌ 不支持的格式: {doc_path}")
        return

    # 加载测试数据
    if test_cases is None:
        if errors_json:
            with open(errors_json, encoding="utf-8") as f:
                test_cases = json.load(f)
            print(f"📥 加载: {errors_json}  ({len(test_cases)} 条)")
        else:
            print("❌ 未提供 errors_json 或 test_cases")
            return

    print("=" * 70)
    print(f"快速替换测试  格式: {fmt.upper()}  用例: {len(test_cases)} 条")
    print("=" * 70)

    # 备份
    backup_path = str(ensure_backup_copy(doc_path, suffix="quicktest"))
    print(f"✓ 备份: {backup_path}")

    # 构建任务列表（含重叠修正）
    tasks = []
    skipped = 0
    for seq_idx, err in enumerate(test_cases):
        old_val  = (err.get("替换锚点") or err.get("译文数值") or "").strip()
        new_val  = (err.get("译文修改建议值") or "").strip()
        context  = err.get("译文上下文", "")
        anchor   = err.get("替换锚点", "")
        reason   = err.get("修改理由", "") or err.get("错误类型", "")
        error_no = err.get("错误编号", "?")
        prev_tgt = err.get("prev_tgt", "")
        next_tgt = err.get("next_tgt", "")

        isc = err.get("is_source_consistent", False)
        if isc is True or (isinstance(isc, str) and isc.lower() == "true"):
            skipped += 1; continue
        if not old_val or not new_val or old_val == new_val:
            skipped += 1; continue

        cleaned = _fix_overlap(context, old_val, new_val)
        if cleaned != new_val:
            print(f"  [去重] '{new_val}' → '{cleaned}'")
        if not cleaned or cleaned == old_val:
            skipped += 1; continue

        # seq_idx 保留原始加载顺序，用于同 old_val 内部按文档顺序排
        tasks.append((seq_idx, error_no, old_val, cleaned, reason, context, anchor, prev_tgt, next_tgt))

    # 排序规则：
    #   主键：长锚点优先（避免短锚点先消耗文本）
    #   次键：同 old_val 时，按原始 seq_idx 升序（保证顺序定位正确）
    tasks.sort(key=lambda t: (-len(t[2]), t[0]))

    # 展开为不含 seq_idx 的格式（保持 runner 接口不变）
    tasks_flat = [(error_no, old_val, new_val, reason, context, anchor, prev_tgt, next_tgt)
                  for seq_idx, error_no, old_val, new_val, reason, context, anchor, prev_tgt, next_tgt in tasks]

    if skipped:
        print(f"  [过滤] {skipped} 条（缺失/无变化/原译一致）")
    print(f"  [执行] {len(tasks_flat)} 条\n")

    # 按格式分发
    runner = {"docx": _run_docx, "xlsx": _run_xlsx,
              "pptx": _run_pptx, "pdf":  _run_pdf}[fmt]
    results = runner(backup_path, tasks_flat)

    # 统计
    success = sum(1 for _, s, _ in results if s == "成功")
    fail    = sum(1 for _, s, _ in results if s == "失败")
    error   = sum(1 for _, s, _ in results if s == "异常")

    print("\n" + "=" * 70)
    print(f"✓ 成功: {success}  ✗ 失败: {fail}  ⚠ 异常: {error}  共: {len(results)}")
    if success + fail > 0:
        print(f"成功率: {success / (success + fail):.1%}")
    print(f"📄 输出: {backup_path}")

    print("\n详细结果:")
    for error_no, status, detail in results:
        sym = {"成功": "✓", "失败": "✗", "异常": "⚠"}.get(status, "?")
        print(f"  {sym} [{error_no}] {status}  {detail[:60]}")

    return results


if __name__ == "__main__":
    DOC_PATH = r"D:\project\数检_程序-AI\新致力公司产品手册设计（翻译版本）_translated.docx"
    ERRORS_JSON = r"D:\project\数检_程序-AI\output\align_body_flat_errors.json"
    output=quick_test(DOC_PATH, errors_json=ERRORS_JSON)
    print(output)
