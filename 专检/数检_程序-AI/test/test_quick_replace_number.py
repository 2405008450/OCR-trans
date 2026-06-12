"""
快速替换测试

直接使用内嵌的错误数据进行测试，或从 align_body_flat_errors.json 加载。

用法：
    python test_quick_replace_number.py "译文.docx"
    python test_quick_replace_number.py "译文.docx" "output/align_body_flat_errors.json"
"""

from docx import Document
from pathlib import Path
import sys
import difflib
import json

# 项目根目录（脚本所在目录）
project_root = str(Path(__file__).resolve().parent)
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from revision import RevisionManager
from replace_revision import replace_and_revise_in_docx, flush_footnote_replacements
from backup_copy.backup_manager import ensure_backup_copy


def _fix_suggestion_overlap(context: str, anchor: str, suggestion: str) -> str:
    """检测 suggestion 与上下文的前/后重叠，返回修正后的实际替换值。"""
    if not context or not anchor or not suggestion:
        return suggestion
    if anchor not in context:
        return suggestion

    start_idx = context.find(anchor)
    end_idx = start_idx + len(anchor)
    actual_sug = suggestion

    # 前部去重
    prefix_window = context[max(0, start_idx - len(suggestion)):start_idx]
    s = difflib.SequenceMatcher(None, prefix_window, suggestion)
    match = s.find_longest_match(0, len(prefix_window), 0, len(suggestion))
    if match.size > 0 and (match.a + match.size == len(prefix_window)) and (match.b == 0):
        actual_sug = suggestion[match.size:]

    # 后部去重
    after_window = context[end_idx: end_idx + len(actual_sug)]
    s = difflib.SequenceMatcher(None, actual_sug, after_window)
    match = s.find_longest_match(0, len(actual_sug), 0, len(after_window))
    if match.size > 0 and (match.a + match.size == len(actual_sug)) and (match.b == 0):
        actual_sug = actual_sug[:match.a]

    return actual_sug


# 内嵌测试数据（可被命令行 JSON 文件覆盖）
TEST_ERRORS = []


def quick_test(doc_path: str, test_cases: list = None, errors_json: str = None):
    """
    快速测试替换功能。

    Args:
        doc_path:    译文文档路径
        test_cases:  直接传入 error 列表（优先级高于 errors_json）
        errors_json: align_body_flat_errors.json 路径；test_cases 为 None 时从此文件加载
    """
    if test_cases is None:
        if errors_json:
            with open(errors_json, encoding="utf-8") as _f:
                test_cases = json.load(_f)
            print(f"📥 从文件加载测试用例: {errors_json}  ({len(test_cases)} 条)")
        else:
            test_cases = TEST_ERRORS

    print("=" * 80)
    print("快速替换测试")
    print("=" * 80)

    if not Path(doc_path).exists():
        print(f"❌ 文档不存在: {doc_path}")
        return

    print(f"\n📂 文档: {doc_path}")
    print(f"📊 测试用例: {len(test_cases)} 个")

    # 创建备份
    print(f"\n📦 创建备份...")
    backup_path = ensure_backup_copy(doc_path, suffix="quicktest")
    print(f"✓ 备份: {backup_path}")

    # 自动编号静态化 + 目录静态化
    try:
        from numbering_to_static import convert_numbering_to_static, has_auto_numbering, convert_toc_to_static
        if has_auto_numbering(backup_path):
            print(f"\n🔢 检测到自动编号，正在转换为静态文本...")
            if convert_numbering_to_static(backup_path):
                print(f"✓ 自动编号已转为静态文本")
            else:
                print(f"⚠ 自动编号静态化失败，部分编号可能无法替换")
        print(f"\n📑 正在转换目录域为静态文本...")
        convert_toc_to_static(backup_path)
    except Exception as e:
        print(f"⚠ 编号/目录静态化异常: {e}")

    # 打开文档
    doc = Document(backup_path)
    doc._numbering_staticized = True
    revision_manager = RevisionManager(doc, author="快速测试")

    # 执行替换
    print(f"\n🔄 开始测试...")
    print("=" * 80)

    results = []

    for idx, error in enumerate(test_cases, 1):
        old_text = (error.get("替换锚点") or "").strip()
        new_text = (error.get("译文修改建议值") or "").strip()
        reason = error.get("修改理由", "")
        context = error.get("译文上下文", "")
        anchor = error.get("替换锚点", "")
        error_no = error.get("错误编号", idx)

        if not old_text or not new_text:
            results.append((error_no, "跳过", "缺少数据"))
            continue

        # 重叠检测：修正建议值
        fixed_text = _fix_suggestion_overlap(context, old_text, new_text)
        if fixed_text != new_text:
            print(f"  ⚙ 建议值修正: '{new_text}' → '{fixed_text}'")
            new_text = fixed_text
        if not new_text or new_text == old_text:
            results.append((error_no, "跳过", f"建议值无效或与原文一致，跳过"))
            continue

        print(f"\n[测试 {idx}/{len(test_cases)}] 错误编号: {error_no}")
        print(f"  查找: '{old_text[:60]}'")
        print(f"  替换: '{new_text[:60]}'")

        try:
            ok, strategy = replace_and_revise_in_docx(
                doc, old_text, new_text, reason, revision_manager,
                context=context, anchor_text=anchor, region="body",
                doc_path=str(backup_path)
            )

            if ok:
                results.append((error_no, "成功", strategy))
                if "批注兜底" in strategy:
                    print(f"  📌 批注: {strategy}")
                else:
                    print(f"  ✓ 成功: {strategy}")
            else:
                results.append((error_no, "失败", strategy))
                print(f"  ✗ 失败: {strategy}")

        except Exception as e:
            results.append((error_no, "异常", str(e)))
            print(f"  ✗ 异常: {e}")

    # 保存
    print(f"\n💾 保存文档...")
    doc.save(backup_path)

    # 执行脚注替换（必须在 doc.save() 之后）
    footnote_count = flush_footnote_replacements(doc, str(backup_path))
    if footnote_count > 0:
        print(f"✓ 脚注替换完成: {footnote_count} 处")

    # 统计
    print("\n" + "=" * 80)
    print("测试结果")
    print("=" * 80)

    success = sum(1 for _, status, _ in results if status == "成功")
    annotated = sum(1 for _, status, detail in results if status == "成功" and "批注兜底" in detail)
    fail = sum(1 for _, status, _ in results if status == "失败")
    skip = sum(1 for _, status, _ in results if status == "跳过")
    error_count = sum(1 for _, status, _ in results if status == "异常")

    print(f"\n✓ 成功: {success}（其中 📌 批注兜底: {annotated}）")
    print(f"✗ 失败: {fail}")
    print(f"⊘ 跳过: {skip}")
    print(f"⚠ 异常: {error_count}")
    print(f"━ 总计: {len(results)}")

    if success + fail > 0:
        print(f"\n成功率: {success / (success + fail):.1%}")

    print("\n详细结果:")
    for error_no, status, detail in results:
        if status == "成功" and "批注兜底" in detail:
            symbol = "📌"
        else:
            symbol = {"成功": "✓", "失败": "✗", "跳过": "⊘", "异常": "⚠"}.get(status, "?")
        print(f"  {symbol} 错误 {error_no}: {status} - {detail[:70]}")

    print(f"\n✅ 测试完成！")
    print(f"📄 结果文档: {backup_path}")

    return results


if __name__ == "__main__":
    DOC_PATH = r"D:\project\数检_程序-AI\测试文件\译文-含不可编辑_01 (2026-007)2025年年度报告(1).docx"
    ERRORS_JSON = r"D:\project\数检_程序-AI\测试文件\output\align_body_flat_errors.json"

    if len(sys.argv) > 1:
        DOC_PATH = sys.argv[1]
    if len(sys.argv) > 2:
        ERRORS_JSON = sys.argv[2]

    print(f"使用文档: {DOC_PATH}")
    if ERRORS_JSON:
        print(f"使用错误JSON: {ERRORS_JSON}")

    if not Path(DOC_PATH).exists():
        print(f"\n❌ 文档不存在: {DOC_PATH}")
        print("\n💡 提示:")
        print("  - 需要使用译文文档")
        print("  - 如果文档在其他位置，请提供完整路径:")
        print(f"    python {Path(__file__).name} \"你的译文文档.docx\" \"output/align_body_flat_errors.json\"")
        sys.exit(1)

    print()
    quick_test(DOC_PATH, errors_json=ERRORS_JSON)
