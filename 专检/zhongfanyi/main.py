import os
import sys
import argparse
from pathlib import Path

# 把「专检」目录加入 path，使 zhongfanyi 可作为包被导入（项目内大量 from zhongfanyi.llm...）
_project_root = Path(__file__).resolve().parent.parent  # 专检
if str(_project_root) not in sys.path:
    sys.path.insert(0, str(_project_root))

from docx import Document
from llm.llm_project.parsers.pdf.pdf_parser import parse_pdf
from llm.llm_project.llm_check.rule_generation import rule
from llm.llm_project.llm_check.check import Match
from llm.llm_project.parsers.word.body_extractor import extract_body_text
from llm.llm_project.parsers.word.footer_extractor import extract_footers
from llm.llm_project.parsers.word.header_extractor import extract_headers
from llm.llm_project.replace.fix_replace_docx import ensure_backup_copy

from llm.llm_project.parsers.json.clean_json import extract_and_parse
from llm.llm_project.utils.json_files import write_json_with_timestamp
from llm.llm_project.note.pizhu import CommentManager
from llm.llm_project.replace.replace_clean import replace_and_comment_in_docx
from llm.llm_project.parsers.txt.txt_parser import parse_txt
from llm.llm_project.utils.txt_files import write_txt_with_timestamp


def load_any_rule(file_path: str) -> str:
    """
    根据文件后缀名自动选择解析器
    """
    path = Path(file_path)
    if not path.exists():
        print(f"⚠️ 文件不存在: {file_path}")
        return ""

    suffix = path.suffix.lower()
    print(f"📂 检测到格式: {suffix}，正在解析...")

    try:
        if suffix == ".pdf":
            # 调用你之前的 PDF 解析器
            return parse_pdf(str(path))

        elif suffix in [".docx", ".doc"]:
            # 调用 Word 正文提取
            return extract_body_text(str(path))

        elif suffix == ".txt":
            # 调用 TXT 解析器
            return parse_txt(str(path))

        else:
            print(f"❌ 不支持的规则文件格式: {suffix}")
            return ""
    except Exception as e:
        print(f"❌ 解析文件 {path.name} 时出错: {e}")
        return ""


def extract_any_document(file_path):
    """
    自动识别 PDF 或 Word 并提取全文内容（页眉+正文+页脚）
    """
    path = Path(file_path)
    if not path.exists():
        return "", "", ""

    suffix = path.suffix.lower()

    if suffix == ".pdf":
        # 注意：PDF 通常没有明确的页眉/页脚分离逻辑，通常全部归为正文
        print(f"📄 正在解析 PDF 文档: {path.name}")
        full_text = parse_pdf(str(path))
        return full_text, "", ""  # 返回：正文, 页眉, 页脚

    elif suffix in [".docx", ".doc"]:
        print(f"📝 正在解析 Word 文档: {path.name}")
        body = extract_body_text(str(path))
        header = extract_headers(str(path))
        footer = extract_footers(str(path))
        return body, header, footer

    else:
        print(f"⚠️ 不支持的文档格式: {suffix}")
        return "", "", ""

def _zhongfanyi_rule_dir():
    """规则目录：相对本文件所在项目根（专检/zhongfanyi）"""
    return Path(__file__).resolve().parent / "llm" / "llm_project" / "rule"


def run_comparison(original_path, translated_path, use_ai_rule=False, output_base_dir=None, ai_rule_file_path=None):
    """
    第一阶段：提取文本并调用 AI/Matcher 进行对比，生成 JSON 报告

    Args:
        original_path: 原文文档路径
        translated_path: 译文文档路径
        use_ai_rule: 是否使用 AI 生成规则（默认 False，使用自定义或默认规则）
        output_base_dir: 可选，JSON 报告输出根目录；为 None 时使用项目内默认路径
        ai_rule_file_path: 可选，使用 AI 规则时从此文件加载规则（支持 pdf/docx/txt）
    """
    print("\n--- 阶段 1: 文本提取与 AI 对比 ---")
    # 1. 自动识别格式并提取 (支持 PDF 和 Word)
    original_body, original_header, original_footer = extract_any_document(original_path)
    translated_body, translated_header, translated_footer = extract_any_document(translated_path)
    # print('==================================原文内容=========================================')
    # print('页眉', original_header)
    # print('正文', original_body)
    # print('页脚', original_footer)
    # print('==================================译文内容=========================================')
    # print('页眉', translated_header)
    # print('正文', translated_body)
    # print('页脚', translated_footer)

    matcher = Match()  # 实例化你的对比对象

    rule_dir = _zhongfanyi_rule_dir()
    if output_base_dir:
        out_root = Path(output_base_dir)
        zhengwen_dir = out_root / "zhengwen" / "output_json"
        yemei_dir = out_root / "yemei" / "output_json"
        yejiao_dir = out_root / "yejiao" / "output_json"
        for d in (zhengwen_dir, yemei_dir, yejiao_dir):
            d.mkdir(parents=True, exist_ok=True)
        parts = [
            ("正文", original_body, translated_body, str(zhengwen_dir)),
            ("页眉", original_header, translated_header, str(yemei_dir)),
            ("页脚", original_footer, translated_footer, str(yejiao_dir)),
        ]
    else:
        parts = [
            ("正文", original_body, translated_body,
             r"C:\Users\Administrator\Desktop\中翻译通用规则项目\zhongfanyi\llm\llm_project\zhengwen\output_json"),
            ("页眉", original_header, translated_header,
             r"C:\Users\Administrator\Desktop\中翻译通用规则项目\zhongfanyi\llm\llm_project\yemei\output_json"),
            ("页脚", original_footer, translated_footer,
             r"C:\Users\Administrator\Desktop\中翻译通用规则项目\zhongfanyi\llm\llm_project\yejiao\output_json")
        ]

    report_paths = {}

    # 规则文件路径（优先相对项目根）
    custom_path = str(rule_dir / "自定义规则.txt")
    default_path = str(rule_dir / "默认规则.txt")
    rule_backup_dir = str(rule_dir)
    if not Path(custom_path).exists():
        custom_path = os.path.join(os.path.dirname(__file__), "llm", "llm_project", "rule", "自定义规则.txt")
    if not Path(default_path).exists():
        default_path = os.path.join(os.path.dirname(__file__), "llm", "llm_project", "rule", "默认规则.txt")

    # 根据用户选择决定是否使用 AI 生成规则
    if use_ai_rule:
        # print("\n" + "=" * 60)
        # print("🤖 用户选择使用 AI 生成规则")
        # print("=" * 60)

        try:
            rule_gen = rule()
            # 支持调用方传入规则文件路径，否则使用默认
            ai_rule_path = ai_rule_file_path or r"C:\Users\Administrator\Desktop\项目文档文件\中翻译中译英规则\1. 通用规则\银行稿件_翻译规则_通用_220307.pdf"
            if not ai_rule_path or not os.path.exists(ai_rule_path):
                raise FileNotFoundError(f"AI 规则文件不存在: {ai_rule_path}")
            #支持word txt pdf格式文件
            ai_rule_text = load_any_rule(ai_rule_path)
            print("正在调用 AI 生成规则...")
            ai_rule = rule_gen.compare_texts(ai_rule_text)

            # 保存到自定义规则文件
            with open(custom_path, 'w', encoding='utf-8') as f:
                f.write(ai_rule)
            print(f"✅ AI 生成的规则已保存到: {custom_path}")

            # 同时保存一份带时间戳的备份
            backup_name, backup_path = write_txt_with_timestamp(ai_rule, rule_backup_dir)
            print(f"📦 备份已保存到: {backup_path}")
            print("=" * 60 + "\n")

            # 使用刚生成的规则
            rule_text = ai_rule

        except Exception as e:
            print(f"❌ AI 生成规则失败: {e}")
            print("⚠️  将尝试使用自定义规则或默认规则...")
            rule_text = parse_txt(custom_path) or parse_txt(default_path)
    else:
        # 不使用 AI，直接读取现有规则
        print("\nℹ️  用户选择使用现有规则（不调用 AI）")
        rule_text = parse_txt(custom_path) or parse_txt(default_path)

    # 最终有效性检查
    if not rule_text:
        print("❌ 错误：无法加载任何有效规则。")
        print("   请确保以下文件之一存在且有效：")
        print(f"   - 自定义规则: {custom_path}")
        print(f"   - 默认规则: {default_path}")
        print("   或使用 --use-ai-rule 参数让 AI 生成规则")
        return None, None
    else:
        # 判断实际使用的规则类型
        if use_ai_rule:
            used_type = "AI 生成规则"
        else:
            used_type = "自定义规则" if parse_txt(custom_path) else "默认规则"

        print(f"✅ 规则加载成功，当前使用: {used_type}")
        if not use_ai_rule:
            print(f"   规则文件: {custom_path if parse_txt(custom_path) else default_path}")
        print(f"--- 规则内容预览 (前200字符) ---")
        print(rule_text[:200] + "..." if len(rule_text) > 200 else rule_text)

    for name, orig_txt, tran_txt, out_dir in parts:
        print(f"====== 正在检查{name} ===========")
        if orig_txt and tran_txt:
            res = matcher.compare_texts(orig_txt, tran_txt, rule_text)
        else:
            res = []
            print(f"⚠️ {name}原文或译文为空")

        # 写入 JSON
        _, path = write_json_with_timestamp(res, out_dir)
        report_paths[name] = path

    return report_paths, rule_text


def run_fix_phase(translated_path, report_paths):
    """
    阶段二：根据 JSON 报告对译文副本执行替换与批注，保存并返回结果路径与统计。

    Args:
        translated_path: 译文文档路径（会被备份后修改）
        report_paths: run_comparison 返回的 {"正文": path, "页眉": path, "页脚": path}

    Returns:
        (result_docx_path, stats_dict)，失败时 result_docx_path 为 None
    """
    backup_copy_path = ensure_backup_copy(translated_path)
    doc = Document(backup_copy_path)
    comment_manager = CommentManager(doc)
    comment_manager.create_initial_comment()

    def load_errors(label, path):
        if path and os.path.exists(path):
            data = extract_and_parse(path)
            print(f"✓ 已加载{label}报告: {len(data)} 条错误")
            return data
        return []

    body_errors = load_errors("正文", report_paths.get("正文"))
    header_errors = load_errors("页眉", report_paths.get("页眉"))
    footer_errors = load_errors("页脚", report_paths.get("页脚"))

    from zhongfanyi.llm.llm_project.replace.replace_clean import is_list_pattern
    import re

    def apply_all_fixes(errors, label, region):
        if not errors:
            return 0, 0, 0, []
        print(f"\n>>> 正在修复 {label} 部分...")
        s_count, f_count = 0, 0
        skip_count = 0
        change_records = []
        processed_numbering_formats = set()
        for idx, e in enumerate(errors, 1):
            old = (e.get("译文数值") or "").strip()
            new = (e.get("译文修改建议值") or "").strip()
            reason = str(e.get("修改理由") or "数值错误").strip()
            context = e.get("译文上下文", "")
            anchor = e.get("替换锚点", "")
            if not old or not new:
                skip_count += 1
                continue
            if is_list_pattern(old):
                numbering_type = None
                if re.match(r'^[ivxlcdm]+\.$', old.lower()):
                    numbering_type = 'lowerRoman'
                elif re.match(r'^[IVXLCDM]+\.$', old):
                    numbering_type = 'upperRoman'
                elif re.match(r'^[a-z]+\.$', old.lower()):
                    numbering_type = 'lowerLetter'
                elif re.match(r'^[A-Z]+\.$', old):
                    numbering_type = 'upperLetter'
                if numbering_type and numbering_type in processed_numbering_formats:
                    s_count += 1
                    change_records.append(f"'{old}' → '{new}'")
                    continue
            ok, strategy = replace_and_comment_in_docx(
                doc, old, new, reason, comment_manager,
                context=context, anchor_text=anchor, region=region
            )
            if ok:
                s_count += 1
                change_records.append(f"'{old}' → '{new}'")
                if "Word自动编号替换" in strategy:
                    if re.match(r'^[ivxlcdm]+\.$', old.lower()):
                        processed_numbering_formats.add('lowerRoman')
                    elif re.match(r'^[IVXLCDM]+\.$', old):
                        processed_numbering_formats.add('upperRoman')
                    elif re.match(r'^[a-z]+\.$', old.lower()):
                        processed_numbering_formats.add('lowerLetter')
                    elif re.match(r'^[A-Z]+\.$', old):
                        processed_numbering_formats.add('upperLetter')
            else:
                f_count += 1
        return s_count, f_count, skip_count, change_records

    b_s, b_f, b_skip, body_changes = apply_all_fixes(body_errors, "正文", region="body")
    h_s, h_f, h_skip, header_changes = apply_all_fixes(header_errors, "页眉", region="header")
    f_s, f_f, f_skip, footer_changes = apply_all_fixes(footer_errors, "页脚", region="footer")

    if h_s or h_f or h_skip:
        comment_manager.append_to_initial_comment(f"\n【页眉】共修改 {h_s} 处")
        for i, change in enumerate(header_changes, 1):
            comment_manager.append_to_initial_comment(f"  {i}. {change}")
    else:
        comment_manager.append_to_initial_comment("\n【页眉】无修改")
    if f_s or f_f or f_skip:
        comment_manager.append_to_initial_comment(f"\n【页脚】共修改 {f_s} 处")
        for i, change in enumerate(footer_changes, 1):
            comment_manager.append_to_initial_comment(f"  {i}. {change}")
    else:
        comment_manager.append_to_initial_comment("\n【页脚】无修改")

    total_s, total_f, total_skip = b_s + h_s + f_s, b_f + h_f + f_f, b_skip + h_skip + f_skip
    doc.save(backup_copy_path)
    print(f"🎉 修复阶段完成，结果保存至: {backup_copy_path}")
    stats = {"success": total_s, "failed": total_f, "skipped": total_skip}
    return backup_copy_path, stats


def run_full_pipeline(original_path, translated_path, output_base_dir, use_ai_rule=False, ai_rule_file_path=None):
    """
    完整流程：对比 + 修复。供 Web 或脚本调用。

    Args:
        original_path: 原文文档路径
        translated_path: 译文文档路径
        output_base_dir: 输出根目录（JSON 报告与最终 docx 均在此下）
        use_ai_rule: 是否使用 AI 生成规则
        ai_rule_file_path: 使用 AI 规则时的规则文件路径（可选，支持 pdf/docx/txt）

    Returns:
        (result_docx_path, report_paths, stats_dict)，任一步失败时 result_docx_path 为 None
    """
    report_paths, _ = run_comparison(
        original_path, translated_path,
        use_ai_rule=use_ai_rule,
        output_base_dir=output_base_dir,
        ai_rule_file_path=ai_rule_file_path,
    )
    if report_paths is None:
        return None, None, None
    result_path, stats = run_fix_phase(translated_path, report_paths)
    return result_path, report_paths, stats


def main():
    # 1) 配置默认路径
    DEFAULT_ORIGINAL = r"C:\Users\Administrator\Desktop\测试文件\原文-中翻译规则测试文件.docx"
    DEFAULT_TRANSLATED = r"C:\Users\Administrator\Desktop\测试文件\译文-中翻译规则测试文件.docx"

    # 2) 命令行参数
    parser = argparse.ArgumentParser(
        description="Word 自动对比、检测与修复工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  # 使用现有规则（自定义规则或默认规则）
  python llm/main_ddd.py

  # 使用 AI 生成规则
  python llm/main_ddd.py --use-ai-rule

  # 指定文档路径
  python llm/main_ddd.py -o "原文.docx" -t "译文.docx"

  # 组合使用
  python llm/main_ddd.py -o "原文.docx" -t "译文.docx" --use-ai-rule
        """
    )
    parser.add_argument("--original", "-o", default=DEFAULT_ORIGINAL,
                        help="原文文档路径")
    parser.add_argument("--translated", "-t", default=DEFAULT_TRANSLATED,
                        help="译文文档路径")
    parser.add_argument("--use-ai-rule", "-a", action="store_true",
                        help="使用 AI 生成翻译规则（默认使用自定义规则或默认规则）")
    args = parser.parse_args()

    # ⚙️ 不使用命令行时的配置（在这里修改）
    # 如果不想用命令行参数，可以直接修改下面的变量
    USE_AI_RULE_CONFIG = False  # 👈 改为 True 使用 AI 生成规则

    # 优先使用命令行参数，如果没有命令行参数则使用配置变量
    use_ai_rule = args.use_ai_rule or USE_AI_RULE_CONFIG

    if not os.path.exists(args.original) or not os.path.exists(args.translated):
        print("❌ 错误: 输入的 docx 文件路径不存在")
        return

    # 3) 执行对比并获取生成的 JSON 路径
    try:
        report_paths, final_rule_text = run_comparison(
            args.original,
            args.translated,
            use_ai_rule=use_ai_rule  # 使用合并后的配置
        )

        # 检查返回值
        if report_paths is None or final_rule_text is None:
            print("❌ 规则加载失败，程序终止")
            return

        print(f"\n✅ 规则内容已加载 (长度: {len(final_rule_text)} 字符)")
    except Exception as e:
        print(f"❌ 调用 API 失败: {e}")
        print("   请检查账户余额、API Key 有效性或网络环境是否正常。")
        return  # 退出程序
    # 4) 核心修复逻辑
    print("\n--- 阶段 2: 自动替换与批注 ---")

    # 创建备份
    backup_copy_path = ensure_backup_copy(args.translated)
    doc = Document(backup_copy_path)
    comment_manager = CommentManager(doc)
    comment_manager.create_initial_comment()

    def load_errors(label, path):
        if path and os.path.exists(path):
            data = extract_and_parse(path)
            print(f"✓ 已加载{label}报告: {len(data)} 条错误")
            return data
        return []

    # 加载刚刚生成的 JSON
    body_errors = load_errors("正文", report_paths.get("正文"))
    header_errors = load_errors("页眉", report_paths.get("页眉"))
    footer_errors = load_errors("页脚", report_paths.get("页脚"))
    #
    # body_result_path=r"C:\Users\Administrator\Desktop\中翻译通用规则项目\zhongfanyi\llm\llm_project\zhengwen\output_json\文本对比结果_20260306_154849.json"
    # header_result_path=r"C:\Users\Administrator\Desktop\中翻译通用规则项目\zhongfanyi\llm\llm_project\yemei\output_json\文本对比结果_20260306_154913.json"
    # footer_result_path=r"C:\Users\Administrator\Desktop\中翻译通用规则项目\zhongfanyi\llm\llm_project\yejiao\output_json\文本对比结果_20260306_154937.json"
    # # 2) 读取错误报告并解析
    # print("\n正在提取解析正文错误报告...")
    # body_errors = extract_and_parse(body_result_path)
    # print("正文错误报告", body_errors)
    # for err in body_errors:
    #     print(err)
    # print("正文错误解析个数：", len(body_errors))
    #
    # print("\n正在提取解析页眉错误报告...")
    # header_errors = extract_and_parse(header_result_path)
    # print("页眉错误报告", header_errors)
    # print("页眉错误解析个数：", len(header_errors))
    #
    # print("\n正在提取解析页脚错误报告...")
    # footer_errors = extract_and_parse(footer_result_path)
    # print("页脚错误报告", footer_errors)
    # print("页脚错误解析个数：", len(footer_errors))
    # print("正文",body_errors)
    # print("页眉",header_errors)
    # print("页脚",footer_errors)

    # 统一定义替换执行函数 (增加区域参数)
    def apply_all_fixes(errors, label, region):
        """
        执行替换修复

        Args:
            errors: 错误列表
            label: 标签（用于显示）
            region: 区域 ("body"=正文, "header"=页眉, "footer"=页脚)

        Returns:
            (成功数, 失败数, 跳过数, 修改记录列表)
        """
        if not errors: return 0, 0, 0, []
        print(f"\n>>> 正在修复 {label} 部分...")
        s_count, f_count = 0, 0
        skip_count = 0
        change_records = []  # 记录所有成功的修改

        # 跟踪已经批量处理的编号格式
        processed_numbering_formats = set()

        for idx, e in enumerate(errors, 1):
            old = (e.get("译文数值") or "").strip()
            new = (e.get("译文修改建议值") or "").strip()
            reason = str(e.get("修改理由") or "数值错误").strip()
            context = e.get("译文上下文", "")
            anchor = e.get("替换锚点", "")

            if not old or not new:
                print(f"  [{idx}] 跳过: 缺少【译文数值】或【译文修改建议值】字段")
                skip_count += 1
                continue

            # 检查是否为编号模式，且已经被批量处理过
            from zhongfanyi.llm.llm_project.replace.replace_clean import is_list_pattern
            if is_list_pattern(old):
                # 提取编号格式类型（如 lowerRoman, upperRoman 等）
                import re
                numbering_type = None
                # 使用更精确的正则表达式匹配
                if re.match(r'^[ivxlcdm]+\.$', old.lower()):
                    numbering_type = 'lowerRoman'
                elif re.match(r'^[IVXLCDM]+\.$', old):
                    numbering_type = 'upperRoman'
                elif re.match(r'^[a-z]+\.$', old.lower()):
                    numbering_type = 'lowerLetter'
                elif re.match(r'^[A-Z]+\.$', old):
                    numbering_type = 'upperLetter'

                if numbering_type:
                    if numbering_type in processed_numbering_formats:
                        # 这个格式已经被之前的操作批量处理过了
                        s_count += 1
                        change_records.append(f"'{old}' → '{new}'")
                        print(f"  [{idx}] 成功: '{old}' -> '{new}' (已由编号批量替换完成)")
                        print(f"    修改理由: {reason}")
                        print(f"    策略: Word自动编号替换 (批量操作已完成)")
                        continue

            ok, strategy = replace_and_comment_in_docx(
                doc, old, new, reason, comment_manager,
                context=context, anchor_text=anchor, region=region
            )
            if ok:
                s_count += 1
                change_records.append(f"'{old}' → '{new}'")
                print(f"  [{idx}] 成功: '{old}' -> '{new}'")
                print(f"    修改理由: {reason}")
                print(f"    策略: {strategy}")
                print(f"    操作: '{old}' → '{new}' (已替换并添加批注)")

                # 如果是Word自动编号替换成功，标记该格式已处理
                if "Word自动编号替换" in strategy:
                    import re
                    if re.match(r'^[ivxlcdm]+\.$', old.lower()):
                        processed_numbering_formats.add('lowerRoman')
                    elif re.match(r'^[IVXLCDM]+\.$', old):
                        processed_numbering_formats.add('upperRoman')
                    elif re.match(r'^[a-z]+\.$', old.lower()):
                        processed_numbering_formats.add('lowerLetter')
                    elif re.match(r'^[A-Z]+\.$', old):
                        processed_numbering_formats.add('upperLetter')
            else:
                f_count += 1
                print(f"  [{idx}] 失败: 未匹配到 '{old}'")
        print(f"\n--- {label} 修复统计 ---")
        print(f"成功: {s_count}")
        print(f"失败: {f_count}")
        print(f"跳过: {skip_count}")
        print(f"总计: {s_count + f_count + skip_count}")
        if s_count + f_count + skip_count > 0:
            success_r = s_count / (s_count + f_count)
        else:
            success_r = 0
        print(f"成功率: {success_r:.2%}")

        return s_count, f_count, skip_count, change_records

    # 执行三部分修复（指定对应区域）
    b_s, b_f, b_skip, body_changes = apply_all_fixes(body_errors, "正文", region="body")
    h_s, h_f, h_skip, header_changes = apply_all_fixes(header_errors, "页眉", region="header")
    f_s, f_f, f_skip, footer_changes = apply_all_fixes(footer_errors, "页脚", region="footer")

    # 在系统批注中记录页眉页脚的修改情况
    print("\n" + "=" * 60)
    print("页眉页脚修改汇总")
    print("=" * 60)

    if h_s == 0 and h_f == 0 and h_skip == 0:
        comment_manager.append_to_initial_comment("\n【页眉】无修改")
        print("  ℹ️  页眉：无修改")
    else:
        # 添加汇总信息
        comment_manager.append_to_initial_comment(f"\n【页眉】共修改 {h_s} 处")
        print(f"\n  ✓ 页眉：共修改 {h_s} 处")

        # 添加详细修改记录
        if header_changes:
            print("  详细修改:")
            for i, change in enumerate(header_changes, 1):
                # 添加到系统批注
                comment_manager.append_to_initial_comment(f"  {i}. {change}")
                # 打印到控制台
                print(f"    {i}. {change}")

    if f_s == 0 and f_f == 0 and f_skip == 0:
        comment_manager.append_to_initial_comment("\n【页脚】无修改")
        print("\n  ℹ️  页脚：无修改")
    else:
        # 添加汇总信息
        comment_manager.append_to_initial_comment(f"\n【页脚】共修改 {f_s} 处")
        print(f"\n  ✓ 页脚：共修改 {f_s} 处")

        # 添加详细修改记录
        if footer_changes:
            print("  详细修改:")
            for i, change in enumerate(footer_changes, 1):
                # 添加到系统批注
                comment_manager.append_to_initial_comment(f"  {i}. {change}")
                # 打印到控制台
                print(f"    {i}. {change}")

    print(f"\n--- 修复统计 ---")
    total_s = b_s + h_s + f_s
    total_f = b_f + h_f + f_f
    total_skip = b_skip + h_skip + f_skip
    total_count = total_s + total_f + total_skip
    print(f"成功: {total_s}")
    print(f"失败: {total_f}")
    print(f"跳过: {total_skip}")
    print(f"总计: {total_count}")
    if total_count > 0:
        success_rate = total_s / (total_s + total_f)
    else:
        success_rate = 0

    print(f"成功率: {success_rate:.2%}")

    # 保存最终结果
    doc.save(backup_copy_path)
    print(f"\n" + "=" * 40)

    print(f"🎉 全部流程处理完成！")
    print(f"最终结果保存至: {backup_copy_path}")
    print("=" * 40)


if __name__ == '__main__':
    main()