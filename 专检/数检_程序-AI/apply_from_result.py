"""
apply_from_result.py
────────────────────
直接读取 AI 检查结果文件，对译文文档执行修订写入。
无需重新调用大模型。

支持两种输入：
  1. align_body.json  — output/ 目录下的 AI 原始结果 JSON（推荐，信息最全）
  2. output_checked.xlsx — 最终报告 Excel（仅含摘要，精度较低）

用法：
  python apply_from_result.py                          # 使用脚本底部默认路径
  python apply_from_result.py result.json  译文.docx   # 命令行指定
"""

import json
import sys
import difflib
from pathlib import Path

# ── 路径修正，确保能 import 项目模块 ──────────────────────────────────
project_root = str(Path(__file__).resolve().parent)
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from docx import Document
from revision import RevisionManager
from replace_revision import replace_and_revise_in_docx, flush_footnote_replacements
from backup_copy.backup_manager import ensure_backup_copy


# ═══════════════════════════════════════════════════════════════════════
# 工具函数
# ═══════════════════════════════════════════════════════════════════════

def _fix_suggestion_overlap(context: str, anchor: str, suggestion: str) -> str:
    """检测 suggestion 与上下文的前/后重叠，返回修正后的实际替换值。"""
    if not context or not anchor or not suggestion:
        return suggestion
    if anchor not in context:
        return suggestion

    start_idx = context.find(anchor)
    end_idx   = start_idx + len(anchor)
    actual    = suggestion

    # 前部去重
    prefix = context[max(0, start_idx - len(suggestion)):start_idx]
    m = difflib.SequenceMatcher(None, prefix, actual).find_longest_match(
        0, len(prefix), 0, len(actual))
    if m.size > 0 and m.a + m.size == len(prefix) and m.b == 0:
        actual = actual[m.size:]

    # 后部去重
    suffix = context[end_idx: end_idx + len(actual)]
    m = difflib.SequenceMatcher(None, actual, suffix).find_longest_match(
        0, len(actual), 0, len(suffix))
    if m.size > 0 and m.a + m.size == len(actual) and m.b == 0:
        actual = actual[:m.a]

    return actual


def _load_tasks_from_json(json_path: str) -> list:
    """
    从 align_body_errors.json 或旧版 align_body.json 读取任务列表。
    每条 task: (seq, old_val, new_val, context, anchor, reason)
    """
    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)

    tasks = []
    for entry in data:
        seq    = entry.get("seq", "?")
        errors = entry.get("errors", [])
        if not errors:
            continue
        for err in errors:
            isc = err.get("is_source_consistent")
            if isc is True or (isinstance(isc, str) and isc.lower() == "true"):
                continue
            old_val = (err.get("替换锚点") or err.get("译文数值") or "").strip()
            new_val = (err.get("译文修改建议值") or "").strip()
            context = (err.get("译文上下文") or "").strip()
            anchor  = (err.get("替换锚点") or "").strip()
            reason  = (err.get("修改理由") or err.get("错误类型") or "").strip()
            if not old_val or not new_val or old_val == new_val:
                continue
            tasks.append((seq, old_val, new_val, context, anchor, reason))

    return tasks


def _load_tasks_from_excel(xlsx_path: str) -> list:
    """
    从 output_checked.xlsx 读取任务列表（降级方案，精度较低）。
    解析 AI修改详情 列，格式：'old' → 'new'；'old2' → 'new2'
    """
    import re
    import pandas as pd

    df = pd.read_excel(xlsx_path)
    if "AI修改详情" not in df.columns:
        raise ValueError("Excel 中未找到 'AI修改详情' 列，请使用 align_body.json")

    tasks = []
    pattern = re.compile(r"'(.+?)'\s*→\s*'(.+?)'")

    for seq, row in df.iterrows():
        if str(row.get("AI是否正确", "")).strip() != "❗错误":
            continue
        detail = str(row.get("AI修改详情", ""))
        context = str(row.get("译文", ""))
        for m in pattern.finditer(detail):
            old_val, new_val = m.group(1).strip(), m.group(2).strip()
            if old_val and new_val and old_val != new_val:
                tasks.append((seq, old_val, new_val, context, old_val, ""))

    return tasks


# ═══════════════════════════════════════════════════════════════════════
# 主流程
# ═══════════════════════════════════════════════════════════════════════

def apply_from_result(
    result_path: str,
    doc_path: str,
    region: str = "body",
    revision_author: str = "数值检查",
    dry_run: bool = False,
):
    """
    参数：
      result_path    — align_body.json 或 output_checked.xlsx
      doc_path       — 译文 .docx 路径
      region         — 替换区域：body / header / footer
      revision_author— 修订作者名
      dry_run        — True 时只打印任务，不实际修改文档
    """
    result_path = str(result_path)
    doc_path    = str(doc_path)

    print("=" * 70)
    print("apply_from_result — 从检查结果直接写入修订")
    print("=" * 70)
    print(f"  结果文件: {result_path}")
    print(f"  译文文档: {doc_path}")
    print(f"  区域:     {region}")
    print(f"  dry_run:  {dry_run}")

    if not Path(result_path).exists():
        print(f"\n❌ 结果文件不存在: {result_path}")
        return
    if not Path(doc_path).exists():
        print(f"\n❌ 译文文档不存在: {doc_path}")
        return

    # ── 读取任务 ──────────────────────────────────────────────────────
    suffix = Path(result_path).suffix.lower()
    if suffix == ".json":
        # 自动降级：如果传入的是旧版 align_body.json，尝试同目录的 align_body_errors.json
        errors_path = Path(result_path).parent / (Path(result_path).stem.replace("align_body", "align_body_errors") + ".json")
        if not any(entry.get("errors") for entry in json.load(open(result_path, encoding="utf-8"))[:5]):
            if errors_path.exists():
                print(f"\n⚠️  检测到旧版 JSON（无 errors 字段），自动切换到: {errors_path.name}")
                result_path = str(errors_path)
        tasks = _load_tasks_from_json(result_path)
        print(f"\n📂 JSON 读取完成，原始任务: {len(tasks)} 条")
    elif suffix in (".xlsx", ".xls"):
        tasks = _load_tasks_from_excel(result_path)
        print(f"\n📂 Excel 读取完成，原始任务: {len(tasks)} 条")
    else:
        print(f"\n❌ 不支持的结果文件格式: {suffix}（仅支持 .json / .xlsx）")
        return

    if not tasks:
        print("\n✅ 没有需要修订的任务，退出。")
        return

    # ── 长文本优先，避免短锚点先替换导致长锚点找不到 ─────────────────
    tasks.sort(key=lambda t: len(t[1]), reverse=True)

    # ── 修正重叠 ──────────────────────────────────────────────────────
    cleaned = []
    for seq, old_val, new_val, context, anchor, reason in tasks:
        fixed = _fix_suggestion_overlap(context, old_val, new_val)
        if fixed != new_val:
            print(f"  ⚙ [seq={seq}] 建议值修正: '{new_val}' → '{fixed}'")
        if fixed and fixed != old_val:
            cleaned.append((seq, old_val, fixed, context, anchor, reason))
        else:
            print(f"  ⊘ [seq={seq}] 修正后无效，跳过: '{old_val}'")

    tasks = cleaned
    print(f"  有效任务: {len(tasks)} 条")

    # ── 计算 occurrence_index（每个 old_val 在任务序列中是第几次出现）──
    from collections import defaultdict as _defaultdict
    _occ_counter_ar: dict = _defaultdict(int)
    tasks_with_occ = []
    for task in tasks:
        seq, old_val, new_val, context, anchor, reason = task
        occ = _occ_counter_ar[old_val]
        _occ_counter_ar[old_val] += 1
        tasks_with_occ.append((seq, old_val, new_val, context, anchor, reason, occ))
    tasks = tasks_with_occ

    if dry_run:
        print("\n[dry_run] 任务列表（不执行实际替换）:")
        for i, (seq, old_val, new_val, *_) in enumerate(tasks, 1):
            print(f"  [{i:>3}] seq={seq}  '{old_val[:50]}' → '{new_val[:50]}'")
        return

    # ── 备份 + 静态化 ─────────────────────────────────────────────────
    print(f"\n📦 创建备份...")
    backup_path = ensure_backup_copy(doc_path, suffix="apply_result")
    print(f"  ✓ 备份: {backup_path}")

    try:
        from numbering_to_static import convert_numbering_to_static, has_auto_numbering, convert_toc_to_static
        if has_auto_numbering(backup_path):
            print(f"\n🔢 检测到自动编号，转换为静态文本...")
            if convert_numbering_to_static(backup_path):
                print(f"  ✓ 自动编号已静态化")
        print(f"\n📑 转换目录域为静态文本...")
        convert_toc_to_static(backup_path)
    except Exception as e:
        print(f"  ⚠ 编号/目录静态化异常: {e}")

    # ── 打开文档，执行替换 ────────────────────────────────────────────
    doc = Document(backup_path)
    doc._numbering_staticized = True
    rm  = RevisionManager(doc, author=revision_author)

    print(f"\n🔄 开始写入修订（共 {len(tasks)} 条）...")
    print("=" * 70)

    success, failed = 0, 0
    for i, (seq, old_val, new_val, context, anchor, reason, occ) in enumerate(tasks, 1):
        try:
            ok, strategy = replace_and_revise_in_docx(
                doc=doc,
                old_value=old_val,
                new_value=new_val,
                reason=reason,
                revision_manager=rm,
                context=context,
                anchor_text=anchor,
                region=region,
                doc_path=backup_path,
                occurrence_index=occ,
            )
            if ok:
                success += 1
                print(f"  ✅ [{i:>3}/{len(tasks)}] seq={seq}  '{old_val[:45]}' → '{new_val[:45]}'  ({strategy})")
            else:
                failed += 1
                print(f"  ⚠️  [{i:>3}/{len(tasks)}] seq={seq}  未找到: '{old_val[:60]}'  ({strategy})")
        except Exception as e:
            failed += 1
            print(f"  ❌ [{i:>3}/{len(tasks)}] seq={seq}  异常: {e}")

    # ── 保存 ──────────────────────────────────────────────────────────
    print(f"\n💾 保存文档...")
    doc.save(backup_path)

    footnote_count = flush_footnote_replacements(doc, backup_path)
    if footnote_count:
        print(f"  📎 脚注替换: {footnote_count} 处")

    # ── 汇总 ──────────────────────────────────────────────────────────
    print("\n" + "=" * 70)
    total = success + failed
    print(f"✅ 成功: {success} / {total}   失败: {failed} / {total}")
    if total:
        print(f"   成功率: {success / total:.1%}")
    print(f"📄 输出文档: {backup_path}")
    print("=" * 70)

    return backup_path


# ═══════════════════════════════════════════════════════════════════════
# 入口
# ═══════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    # ── 命令行模式 ────────────────────────────────────────────────────
    if len(sys.argv) >= 3:
        apply_from_result(
            result_path=sys.argv[1],
            doc_path=sys.argv[2],
            region=sys.argv[3] if len(sys.argv) > 3 else "body",
        )
        sys.exit(0)

    # ── 默认路径（修改这里即可）──────────────────────────────────────
    RESULT_PATH = r"测试文件/output/align_body_errors.json"   # main.py 运行后生成；也可用 output_checked.xlsx
    DOC_PATH    = r"测试文件/译文-含不可编辑_01 (2026-007)2025年年度报告(1).docx"

    apply_from_result(
        result_path=RESULT_PATH,
        doc_path=DOC_PATH,
        region="body",
        revision_author="数值检查",
        dry_run=False,   # 改为 True 可先预览任务列表，不实际修改文档
    )
