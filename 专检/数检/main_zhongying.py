from datetime import datetime
import os
import argparse
from docx import Document
from pathlib import Path
from parsers.pdf.pdf_parser import parse_pdf
from llm.check import Match
from parsers.word.body_extractor import extract_body_text
from parsers.word.footer_extractor import extract_footers
from parsers.word.header_extractor import extract_headers
from replace.fix_replace_docx import ensure_backup_copy
from parsers.json.clean_json import extract_and_parse
from utils.json_files import write_json_with_timestamp
from revise.revision import RevisionManager
from replace.replace_revision import replace_and_revise_in_docx
from parsers.txt.txt_parser import parse_txt
from utils.txt_files import write_txt_with_timestamp
from parsers.excel.excel_parser import parse_excel_with_pandas
from parsers.excel.excel装载 import ExcelReportGenerator
from parsers.pptx.pptx_parser import parse_pptx
from replace.excel.excel_replacer import ExcelReplacer
from replace.pptx.pptx_replacer import PPTXReplacer
from divide.text_splitter import split_text_pair, _count_chars
from llm.check_2 import run_bilingual_comparison


# ========== 模式开关 ==========
# True  = 单文件双语对照模式（一个文件内中英对照）
# False = 双文件模式（原文 + 译文 分别两个文件）
BILINGUAL_MODE = True
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

        elif suffix == ".xlsx":
            # 调用 xlsx 解析器
            return parse_excel_with_pandas(str(path))

        elif suffix == ".pptx":
            # 调用 pptx 解析器
            return parse_pptx(str(path))


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
        print(full_text)
        return full_text, "", ""  # 返回：正文, 页眉, 页脚

    elif suffix in ".xlsx":
        print(f"📝 正在解析 xlsx 文档: {path.name}")
        body = parse_excel_with_pandas(str(path))
        return body, "", ""

    elif suffix in ".pptx":
        print(f"📝 正在解析 pptx 文档: {path.name}")
        body = parse_pptx(str(path))
        return body, "", ""

    elif suffix in [".docx", ".doc"]:
        print(f"📝 正在解析 Word 文档: {path.name}")
        body = extract_body_text(str(path))
        header_raw = extract_headers(str(path))
        footer_raw = extract_footers(str(path))
        # extract_headers/extract_footers 返回 List[str]，需要拼接为单个字符串
        header = "\n".join(header_raw) if isinstance(header_raw, list) else (header_raw or "")
        footer = "\n".join(footer_raw) if isinstance(footer_raw, list) else (footer_raw or "")
        return body, header, footer

    else:
        print(f"⚠️ 不支持的文档格式: {suffix}")
        return "", "", ""


def _compare_with_split(matcher, orig_txt, tran_txt, name):
    """对一组原文/译文文本进行分块对比，合并结果。

    如果文本较短则直接对比；较长则自动分割为对齐的块，
    逐块调用 LLM，最后合并去重返回完整结果列表。

    Returns:
        list[dict] — 始终返回解析后的错误列表（可能为空列表）
    """
    from parsers.json.clean_json import parse_json_content

    pairs = split_text_pair(orig_txt, tran_txt)
    num_parts = len(pairs)

    if num_parts <= 1:
        # 文本不长，直接对比
        print(f"  📄 {name}文本较短（{_count_chars(orig_txt)} 字），直接对比")
        raw = None
        for attempt in range(3):
            try:
                raw = matcher.compare_texts(orig_txt, tran_txt)
                break
            except Exception as e:
                if attempt < 2:
                    wait = (attempt + 1) * 10
                    print(f"  ⚠️ 调用失败: {e}，{wait}秒后重试...")
                    import time
                    time.sleep(wait)
                else:
                    print(f"  ❌ 调用 API 失败（已重试 2 次）: {e}")
        # 解析为列表以保持返回格式一致
        if raw:
            parsed = parse_json_content(raw)
            return parsed if isinstance(parsed, list) else []
        return []

    print(f"  📄 {name}文本较长（原文 {_count_chars(orig_txt)} 字），分割为 {num_parts} 块进行对比")

    all_errors = []
    seen_keys = set()  # 用于去重（缓冲区重叠可能产生重复）

    for i, (orig_chunk, trans_chunk) in enumerate(pairs, 1):
        print(f"\n  --- {name} 第 {i}/{num_parts} 块 ---")
        print(f"      原文块: {_count_chars(orig_chunk)} 字, 译文块: {_count_chars(trans_chunk)} 字")

        if not orig_chunk.strip() or not trans_chunk.strip():
            print(f"      ⚠️ 块内容为空，跳过")
            continue

        # 带重试的 API 调用（最多重试 2 次，间隔递增）
        max_retries = 2
        chunk_result = None
        for attempt in range(max_retries + 1):
            try:
                chunk_result = matcher.compare_texts(orig_chunk, trans_chunk)
                break  # 成功则跳出重试循环
            except Exception as e:
                if attempt < max_retries:
                    wait = (attempt + 1) * 10  # 10s, 20s
                    print(f"      ⚠️ 第 {i} 块第 {attempt + 1} 次调用失败: {e}")
                    print(f"      ⏳ 等待 {wait} 秒后重试...")
                    import time
                    time.sleep(wait)
                else:
                    print(f"      ❌ 第 {i} 块调用 API 失败（已重试 {max_retries} 次）: {e}")
                    print(f"      ⚠️ 跳过该块，继续处理剩余块")
                    chunk_result = None

        # 解析本块结果并去重合并
        chunk_count = 0
        if chunk_result:
            parsed = parse_json_content(chunk_result)
            if isinstance(parsed, list):
                chunk_count = len(parsed)
                for item in parsed:
                    # 用 (译文数值, 译文上下文前50字) 作为去重键
                    dedup_key = (
                        (item.get("译文数值") or "").strip(),
                        (item.get("译文上下文") or "").strip()[:50]
                    )
                    if dedup_key not in seen_keys:
                        seen_keys.add(dedup_key)
                        all_errors.append(item)
                    else:
                        print(f"      ⚠️ 去重: '{dedup_key[0]}' (缓冲区重叠)")

        print(f"      ✓ 本块发现 {chunk_count} 个问题，累计 {len(all_errors)} 个")

    print(f"\n  📊 {name}合并完成: 共 {len(all_errors)} 个不重复问题")

    # 重新编号错误编号
    for idx, item in enumerate(all_errors, 1):
        item["错误编号"] = str(idx)

    return all_errors


def run_comparison(original_path, translated_path, use_ai_rule=False):
    """
    第一阶段：提取文本并调用 AI/Matcher 进行对比，生成 JSON 报告

    对于长文档，自动分割原文/译文为对齐的块，逐块调用 LLM 检查，
    最后合并去重为整个文档的完整 JSON 结果。

    Args:
        original_path: 原文文档路径
        translated_path: 译文文档路径
        use_ai_rule: 是否使用 AI 生成规则（默认 False，使用自定义或默认规则）
    """
    print("\n--- 阶段 1: 文本提取与 AI 对比 ---")

    # 1. 自动识别格式并提取 (支持 PDF 和 Word)
    try:
        original_body, original_header, original_footer = extract_any_document(original_path)
        translated_body, translated_header, translated_footer = extract_any_document(translated_path)
        print(original_body)
        print(translated_body)

    except Exception as e:
        print(f"❌ 解析失败: {e}")
        print("   不支持的格式，请检查文件是否损坏。")
        return  # 退出程序

    matcher = Match()  # 实例化你的对比对象

    parts = [
        ("正文", original_body, translated_body,
         r"D:\Users\Administrator\Desktop\项目文件\专检\数值检查1\llm\llm_project\zhengwen\output_json"),
        ("页眉", original_header, translated_header,
         r"D:\Users\Administrator\Desktop\项目文件\专检\数值检查1\llm\llm_project\yemei\output_json"),
        ("页脚", original_footer, translated_footer,
         r"D:\Users\Administrator\Desktop\项目文件\专检\数值检查1\llm\llm_project\yejiao\output_json")
    ]
    report_paths = {}
    for name, orig_txt, tran_txt, out_dir in parts:
        print(f"====== 正在检查{name} ===========")
        if orig_txt and tran_txt:
            try:
                res = _compare_with_split(matcher, orig_txt, tran_txt, name)
            except Exception as e:
                print(f"❌ 调用 API 失败: {e}")
                print("   请检查账户余额、API Key 有效性或网络环境是否正常。")
                return  # 退出程序
        else:
            res = []
            print(f"⚠️ {name}原文或译文为空")

        # 写入 JSON（合并后的完整文档结果）
        _, path = write_json_with_timestamp(res, out_dir)
        report_paths[name] = path

    return report_paths


def main():

    # 1) 配置默认路径
    DEFAULT_ORIGINAL = r"E:\Documents\a_test\中英专检测试.docx"
    DEFAULT_TRANSLATED = r"D:\Users\Administrator\Desktop\项目文件\专检\数值检查V2\测试文件\译文-含不可编辑_【交付翻译】20260319 艾美疫苗2025年度环境、社会及管治（ESG）报告初稿V2.12.docx"

    # --- 单文件双语对照模式 ---
    if BILINGUAL_MODE:
        if not os.path.exists(DEFAULT_ORIGINAL):
            print(f"❌ 错误: 文件不存在 {DEFAULT_ORIGINAL}")
            return
        print("📋 模式: 单文件双语对照检查")
        report_paths = run_bilingual_comparison(DEFAULT_ORIGINAL)
        if report_paths is None:
            print("❌ 报告生成失败，程序终止")
            return

        # 检查文件格式
        path = Path(DEFAULT_ORIGINAL)
        suffix = path.suffix.lower()

        # 加载错误报告
        def load_errors(label, path):
            if path and os.path.exists(path):
                data = extract_and_parse(path)
                print(f"✓ 已加载{label}报告: {len(data)} 条错误")
                return data
            return []

        body_errors = load_errors("正文", report_paths.get("正文"))
        header_errors = load_errors("页眉", report_paths.get("页眉"))
        footer_errors = load_errors("页脚", report_paths.get("页脚"))
        all_errors = body_errors + header_errors + footer_errors

        if not all_errors:
            print("\n✓ 未发现错误，无需修改")
            return

        # 生成 Excel 报告
        print("\n--- 生成Excel报告 ---")
        e = ExcelReportGenerator()
        body_file_path, _ = e.get_excel(body_errors, output_dir=str(Path(__file__).parent / "zhengwen" / "output_excel"))
        print(f"✓ 正文报告已存至: {body_file_path}")
        header_file_path, _ = e.get_excel(header_errors, output_dir=str(Path(__file__).parent / "yemei" / "output_excel"))
        print(f"✓ 页眉报告已存至: {header_file_path}")
        footer_file_path, _ = e.get_excel(footer_errors, output_dir=str(Path(__file__).parent / "yejiao" / "output_excel"))
        print(f"✓ 页脚报告已存至: {footer_file_path}")

        # 根据文件格式执行修订
        from backup_copy.backup_manager import ensure_backup_copy
        if suffix in [".docx", ".doc"]:
            print("\n--- 阶段 2: Word 自动替换与修订 ---")
            backup_copy_path = ensure_backup_copy(DEFAULT_ORIGINAL)
            doc = Document(backup_copy_path)
            revision_manager = RevisionManager(doc, author="翻译校对")
            print("✓ 修订模式已启用（Track Changes）")

            # 执行修复
            def apply_all_fixes(errors, label, region):
                if not errors: return 0, 0, 0, [], []
                print(f"\n>>> 正在修复 {label} 部分...")
                s_count, f_count, skip_count = 0, 0, 0
                change_records, failed_items = [], []

                for idx, e in enumerate(errors, 1):
                    old = (e.get("译文数值") or "").strip()
                    new = (e.get("译文修改建议值") or "").strip()
                    reason = str(e.get("修改理由") or "数值错误").strip()
                    context = e.get("译文上下文", "")
                    anchor = e.get("替换锚点", "")

                    if not old or not new:
                        skip_count += 1
                        continue

                    ok, strategy = replace_and_revise_in_docx(
                        doc, old, new, reason, revision_manager,
                        context=context, anchor_text=anchor, region=region
                    )
                    if ok:
                        s_count += 1
                        change_records.append(f"'{old}' → '{new}'")
                        print(f"  [{idx}] ✓ '{old}' -> '{new}' ({strategy})")
                    else:
                        f_count += 1
                        failed_items.append({
                            "译文数值": old,
                            "译文修改建议值": new,
                            "修改理由": reason,
                            "译文上下文": context
                        })
                        print(f"  [{idx}] ✗ 未匹配到 '{old}'")

                print(f"\n--- {label} 修复统计: 成功 {s_count}, 失败 {f_count}, 跳过 {skip_count} ---")
                return s_count, f_count, skip_count, change_records, failed_items

            b_s, b_f, b_skip, _, body_failed = apply_all_fixes(body_errors, "正文", "body")
            h_s, h_f, h_skip, _, header_failed = apply_all_fixes(header_errors, "页眉", "header")
            f_s, f_f, f_skip, _, footer_failed = apply_all_fixes(footer_errors, "页脚", "footer")

            # 保存失败项
            all_failed = body_failed + header_failed + footer_failed
            if all_failed:
                fail_dir = Path(__file__).parent / "fail"
                fail_dir.mkdir(parents=True, exist_ok=True)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                fail_report_path = fail_dir / f"{path.stem}_failed_word_{timestamp}.txt"
                with open(fail_report_path, 'w', encoding='utf-8') as f:
                    f.write("Word 修订失败报告\n" + "=" * 60 + "\n\n")
                    for i, item in enumerate(all_failed, 1):
                        f.write(f"{i}. 查找: {item['译文数值']}\n")
                        f.write(f"   建议: {item['译文修改建议值']}\n")
                        f.write(f"   理由: {item['修改理由']}\n\n")
                print(f"\n✓ 失败项详情已保存至: {fail_report_path}")

            doc.save(backup_copy_path)
            print(f"\n✅ Word 处理完成，结果保存至: {backup_copy_path}")

        else:
            print(f"\n⚠️ 文件格式 {suffix} 暂不支持自动修订，仅生成报告")

        return

    # --- 双文件模式 ---
    print("📋 模式: 双文件原文/译文对比")

    # 2) 命令行参数
    parser = argparse.ArgumentParser(
        description="Word 自动对比、检测与修复工具",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument("--original", "-o", default=DEFAULT_ORIGINAL,
                        help="原文文档路径")
    parser.add_argument("--translated", "-t", default=DEFAULT_TRANSLATED,
                        help="译文文档路径")
    args = parser.parse_args()

    if not os.path.exists(args.original) or not os.path.exists(args.translated):
        print("❌ 错误: 输入的文件路径不存在")
        return

    # 检查文件格式
    original_ext = Path(args.original).suffix.lower()
    translated_ext = Path(args.translated).suffix.lower()

    # 判断是否为 PDF 文件
    is_pdf_workflow = (original_ext == ".pdf" or translated_ext == ".pdf")

    # 3) 执行对比并获取生成的 JSON 路径
    try:
        report_paths = run_comparison(
            args.original,
            args.translated,
        )

        # 检查返回值
        if report_paths is None:
            print("❌ 报告生成失败，程序终止")
            return

        print(f"\n✅ 报告已生成")
    except Exception as e:
        print(f"❌ 调用 API 失败: {e}")
        print("   请检查账户余额、API Key 有效性或网络环境是否正常。")
        return  # 退出程序

    # 如果是 PDF 文件，使用批注功能
    if is_pdf_workflow:
        print("\n⚠️ 检测到 PDF 文件")
        print("对比结果已保存至:")
        for name, path in report_paths.items():
            print(f"  - {name}: {path}")

        # 如果译文是 PDF，尝试添加批注
        if translated_ext == ".pdf":
            print("\n--- 阶段 2: PDF 批注 ---")
            try:
                from llm.llm_project.replace.pdf_annotator import PDFAnnotator
                from llm.llm_project.backup_copy.backup_manager import ensure_backup_copy

                # 加载错误报告
                def load_errors(label, path):
                    if path and os.path.exists(path):
                        data = extract_and_parse(path)
                        print(f"✓ 已加载{label}报告: {len(data)} 条错误")
                        return data
                    return []

                body_errors = load_errors("正文", report_paths.get("正文"))
                header_errors = load_errors("页眉", report_paths.get("页眉"))
                footer_errors = load_errors("页脚", report_paths.get("页脚"))

                all_errors = body_errors + header_errors + footer_errors

                if not all_errors:
                    print("✓ 未发现错误，无需添加批注")
                    return
                print(f"正在为 {len(all_errors)} 个错误添加批注...")

                # 输出为Excel
                print("\n--- 生成Excel报告 ---")
                e = ExcelReportGenerator()

                # 正文excel
                body_target_folder = r"D:\Users\Administrator\Desktop\项目文件\专检\数值检查1\llm\llm_project\zhengwen\output_excel"
                body_file_path, body_data = e.get_excel(body_errors, output_dir=body_target_folder)
                print(f"✓ 正文报告已存至: {body_file_path}")

                # 页眉excel
                header_target_folder = r"D:\Users\Administrator\Desktop\项目文件\专检\数值检查1\llm\llm_project\yemei\output_excel"
                header_file_path, header_data = e.get_excel(header_errors, output_dir=header_target_folder)
                print(f"✓ 页眉报告已存至: {header_file_path}")

                # 页脚excel
                footer_target_folder = r"D:\Users\Administrator\Desktop\项目文件\专检\数值检查1\llm\llm_project\yejiao\output_excel"
                footer_file_path, footer_data = e.get_excel(footer_errors, output_dir=footer_target_folder)
                print(f"✓ 页脚报告已存至: {footer_file_path}")

                # 先备份原文件到 backup 文件夹
                print("\n📦 正在备份原始 PDF 文件...")
                backup_path = ensure_backup_copy(args.translated, suffix="annotated")
                print(f"✓ 备份完成，将对备份文件进行批注")

                # 对备份文件进行批注（不是原文件）
                output_path = backup_path

                # 打开备份文件进行批注
                with PDFAnnotator(backup_path) as annotator:
                    success_count = 0
                    fail_count = 0
                    failed_items = []  # 记录失败的项目
                    annotated_positions = set()  # 记录已批注的位置，避免重复

                    # 检查是否启用调试模式
                    debug_mode = os.environ.get('PDF_ANNOTATE_DEBUG', '').lower() == 'true'
                    if debug_mode:
                        print("\n[调试模式已启用]")

                    for idx, error in enumerate(all_errors, 1):
                        old_text = (error.get("译文数值") or "").strip()
                        new_text = (error.get("译文修改建议值") or "").strip()
                        reason = error.get("修改理由", "数值错误")
                        error_type = error.get("错误类型", "数值错误")
                        context = error.get("译文上下文", "")  # 获取上下文

                        if not old_text:
                            continue

                        # 根据错误类型选择颜色
                        if "数值错误" in error_type:
                            color = (1, 1, 0)  # 黄色
                        elif "层级错误" in error_type or "编号错误" in error_type:
                            color = (1, 0.5, 0)  # 橙色
                        elif "日期错误" in error_type:
                            color = (1, 0, 0)  # 红色
                        else:
                            color = (0.5, 0.5, 1)  # 浅蓝色

                        # 构建批注内容
                        comment = f"【{error_type}】\n"
                        comment += f"原文: {error.get('原文数值', 'N/A')}\n"
                        comment += f"建议修改为: {new_text}\n"
                        comment += f"理由: {reason}"

                        # 使用基于上下文的智能批注
                        result = annotator.highlight_and_comment_with_context(
                            search_text=old_text,
                            comment=comment,
                            context=context,  # 传入上下文
                            color=color,
                            debug=debug_mode  # 传入调试标志
                        )

                        # result 现在返回 (count, position_key)
                        if isinstance(result, tuple):
                            count, position_key = result
                        else:
                            count = result
                            position_key = None

                        if count > 0:
                            # 检查是否重复批注
                            if position_key and position_key in annotated_positions:
                                print(f"  [{idx}/{len(all_errors)}] ⚠ 跳过重复位置: {old_text}")
                                continue

                            if position_key:
                                annotated_positions.add(position_key)

                            success_count += 1
                            print(f"  [{idx}/{len(all_errors)}] ✓ 已标注: {old_text}")
                        else:
                            fail_count += 1
                            failed_items.append({
                                'text': old_text,
                                'suggestion': new_text,
                                'reason': reason,
                                'context': context  # 保存上下文用于调试
                            })
                            print(f"  [{idx}/{len(all_errors)}] ✗ 未找到: {old_text}")
                            if context:
                                print(f"      上下文: {context[:50]}...")  # 显示前50个字符

                    # 保存批注后的文件（覆盖备份文件）
                    annotator.save(str(output_path))

                print(f"\n--- PDF 批注统计 ---")
                print(f"成功: {success_count}")
                print(f"失败: {fail_count}")
                print(f"总计: {len(all_errors)}")

                # 如果有失败项，生成失败报告
                if failed_items:
                    print(f"\n⚠️ 以下 {len(failed_items)} 项未能自动标注（可能需要人工检查）:")
                    for item in failed_items[:5]:  # 只显示前5个
                        print(f"  - '{item['text']}' → '{item['suggestion']}'")
                    if len(failed_items) > 5:
                        print(f"  ... 还有 {len(failed_items) - 5} 项")

                    # 保存失败报告到 fail 文件夹
                    fail_dir = Path(__file__).parent / "llm" / "llm_project" / "fail"
                    fail_dir.mkdir(parents=True, exist_ok=True)
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    fail_report_path = fail_dir / f"{Path(args.translated).stem}_failed_annotations_{timestamp}.txt"
                    with open(fail_report_path, 'w', encoding='utf-8') as f:
                        f.write("未能自动标注的项目\n")
                        f.write("=" * 60 + "\n\n")
                        for i, item in enumerate(failed_items, 1):
                            f.write(f"{i}. 查找文本: {item['text']}\n")
                            f.write(f"   建议修改: {item['suggestion']}\n")
                            f.write(f"   修改理由: {item['reason']}\n")
                            if item.get('context'):
                                f.write(f"   上下文: {item['context']}\n")
                            f.write("\n")
                    print(f"\n✓ 失败项详情已保存至: {fail_report_path}")

                print(f"\n✅ 带批注的 PDF 已保存至: {output_path}")

                # --- 阶段 3: PDF 文本替换（改进版 - 保留格式和批注） ---
                print("\n--- 阶段 3: PDF 文本替换（改进版） ---")
                try:
                    from replace.pdf_replacer_improved import ImprovedPDFReplacer

                    # 创建替换版本的备份
                    replace_backup_path = ensure_backup_copy(args.translated, suffix="replaced")
                    print(f"✓ 已创建替换版本备份: {replace_backup_path}")
                    print("\n使用改进的替换算法:")
                    print("  - 利用批注定位精确查找文本位置")
                    print("  - 保留原文本的格式和样式")
                    print("  - 替换后添加绿色批注标记")
                    print()

                    with ImprovedPDFReplacer(replace_backup_path) as replacer:
                        replace_success = 0
                        replace_fail = 0
                        annot_success = 0
                        failed_replacements = []

                        for idx, error in enumerate(all_errors, 1):
                            old_text = (error.get("译文数值") or "").strip()
                            new_text = (error.get("译文修改建议值") or "").strip()
                            reason = error.get("修改理由", "数值错误")
                            error_type = error.get("错误类型", "数值错误")
                            context = error.get("译文上下文", "")

                            if not old_text or not new_text:
                                continue

                            # 根据错误类型选择颜色（绿色表示已修改）
                            if "数值错误" in error_type:
                                color = (0, 1, 0)  # 绿色
                            elif "层级错误" in error_type or "编号错误" in error_type:
                                color = (0, 0.8, 0)  # 深绿色
                            elif "日期错误" in error_type:
                                color = (0, 0.6, 0)  # 墨绿色
                            else:
                                color = (0, 1, 0.5)  # 青绿色

                            # 构建批注内容
                            comment = f"【已修改】{error_type}\n"
                            comment += f"原值: {old_text}\n"
                            comment += f"新值: {new_text}\n"
                            comment += f"理由: {reason}"

                            # 执行替换并添加批注（一步完成）
                            rep_count, ann_count, _ = replacer.replace_and_annotate(
                                search_text=old_text,
                                new_text=new_text,
                                comment=comment,
                                context=context,
                                color=color,
                                debug=False
                            )

                            if rep_count > 0:
                                replace_success += 1
                                if ann_count > 0:
                                    annot_success += 1
                                print(f"  [{idx}/{len(all_errors)}] ✓ '{old_text}' → '{new_text}' (已替换并标注)")
                            else:
                                replace_fail += 1
                                failed_replacements.append((old_text, new_text, error))
                                print(f"  [{idx}/{len(all_errors)}] ✗ '{old_text}' (未替换)")

                        # 保存
                        replacer.save(str(replace_backup_path))

                    print(f"\n--- PDF 替换统计 ---")
                    print(f"替换成功: {replace_success}")
                    print(f"批注成功: {annot_success}")
                    print(f"替换失败: {replace_fail}")
                    print(f"总计: {len(all_errors)}")

                    if replace_success > 0:
                        print(f"\n✅ 替换并批注后的 PDF 已保存至: {replace_backup_path}")
                        print(f"   - 文本已替换（保留原格式）")
                        print(f"   - 已添加绿色批注标记")

                    # 如果有替换失败的项目，提示用户
                    if failed_replacements:
                        print(f"\n⚠️ 注意: 有 {len(failed_replacements)} 个项目替换失败")
                        print("   原因: 可能是文本格式特殊或上下文匹配失败")
                        print("   建议: 查看批注版本，手动修改这些项目")

                        # 保存失败项到文件
                        fail_dir = Path("llm/llm_project/fail")
                        fail_dir.mkdir(parents=True, exist_ok=True)
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        fail_report_path = fail_dir / f"pdf_replacement_failed_{timestamp}.txt"

                        with open(fail_report_path, 'w', encoding='utf-8') as f:
                            f.write("PDF 替换失败项目\n")
                            f.write("=" * 80 + "\n\n")
                            for old, new, err in failed_replacements:
                                f.write(f"查找: {old}\n")
                                f.write(f"替换: {new}\n")
                                f.write(f"理由: {err.get('修改理由', '')}\n")
                                f.write(f"上下文: {err.get('译文上下文', '')}\n")
                                f.write("\n" + "-" * 80 + "\n\n")

                        print(f"   失败项详情已保存至: {fail_report_path}")

                except Exception as e:
                    print(f"\n❌ PDF 替换失败: {e}")
                    import traceback
                    traceback.print_exc()

            except ImportError:
                print("\n❌ 错误: 未安装 PyMuPDF 库")
                print("请运行: pip install PyMuPDF")
            except Exception as e:
                print(f"\n❌ PDF 批注失败: {e}")
                import traceback
                traceback.print_exc()

        return

    # ===== Excel 工作流 =====
    is_excel_workflow = (translated_ext == ".xlsx")
    if is_excel_workflow:
        print("\n--- 阶段 2: Excel 替换与批注 ---")
        try:
            from llm.llm_project.backup_copy.backup_manager import ensure_backup_copy as ensure_backup

            def load_errors(label, path):
                if path and os.path.exists(path):
                    data = extract_and_parse(path)
                    print(f"✓ 已加载{label}报告: {len(data)} 条错误")
                    return data
                return []

            body_errors = load_errors("正文", report_paths.get("正文"))
            header_errors = load_errors("页眉", report_paths.get("页眉"))
            footer_errors = load_errors("页脚", report_paths.get("页脚"))
            all_errors = body_errors + header_errors + footer_errors

            if not all_errors:
                print("✓ 未发现错误，无需修改")
                return

            # 输出为Excel报告
            print("\n--- 生成Excel报告 ---")
            eg = ExcelReportGenerator()
            body_target_folder = r"D:\Users\Administrator\Desktop\项目文件\专检\数值检查1\llm\llm_project\zhengwen\output_excel"
            body_file_path, _ = eg.get_excel(body_errors, output_dir=body_target_folder)
            print(f"✓ 正文报告已存至: {body_file_path}")

            header_target_folder = r"D:\Users\Administrator\Desktop\项目文件\专检\数值检查1\llm\llm_project\yemei\output_excel"
            header_file_path, _ = eg.get_excel(header_errors, output_dir=header_target_folder)
            print(f"✓ 页眉报告已存至: {header_file_path}")

            footer_target_folder = r"D:\Users\Administrator\Desktop\项目文件\专检\数值检查1\llm\llm_project\yejiao\output_excel"
            footer_file_path, _ = eg.get_excel(footer_errors, output_dir=footer_target_folder)
            print(f"✓ 页脚报告已存至: {footer_file_path}")

            # 备份译文
            print("\n📦 正在备份原始 Excel 文件...")
            backup_path = ensure_backup(args.translated)
            print(f"✓ 备份完成: {backup_path}")

            # 执行替换与批注
            print(f"\n正在为 {len(all_errors)} 个错误执行替换与批注...")
            replacer = ExcelReplacer(backup_path)
            success_count = 0
            fail_count = 0
            failed_items = []

            for idx, error in enumerate(all_errors, 1):
                old_text = (error.get("译文数值") or "").strip()
                new_text = (error.get("译文修改建议值") or "").strip()
                reason = error.get("修改理由", "数值错误")
                context = error.get("译文上下文", "")

                if not old_text or not new_text:
                    continue

                ok = replacer.replace_and_annotate(
                    old_text=old_text,
                    new_text=new_text,
                    reason=reason,
                    context=context,
                )

                if ok:
                    success_count += 1
                    print(f"  [{idx}/{len(all_errors)}] ✓ '{old_text}' → '{new_text}'")
                else:
                    fail_count += 1
                    failed_items.append({'text': old_text, 'suggestion': new_text, 'reason': reason})
                    print(f"  [{idx}/{len(all_errors)}] ✗ 未找到: '{old_text}'")

            replacer.save(str(backup_path))

            print(f"\n--- Excel 替换统计 ---")
            print(f"成功: {success_count}")
            print(f"失败: {fail_count}")
            print(f"总计: {len(all_errors)}")
            if success_count + fail_count > 0:
                print(f"成功率: {success_count / (success_count + fail_count):.2%}")

            if failed_items:
                fail_dir = Path(__file__).parent / "llm" / "llm_project" / "fail"
                fail_dir.mkdir(parents=True, exist_ok=True)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                fail_report_path = fail_dir / f"{Path(args.translated).stem}_failed_excel_{timestamp}.txt"
                with open(fail_report_path, 'w', encoding='utf-8') as f:
                    f.write("Excel 替换失败项目\n" + "=" * 60 + "\n\n")
                    for i, item in enumerate(failed_items, 1):
                        f.write(
                            f"{i}. 查找: {item['text']}\n   建议: {item['suggestion']}\n   理由: {item['reason']}\n\n")
                print(f"✓ 失败项详情已保存至: {fail_report_path}")

            print(f"\n✅ Excel 处理完成，结果保存至: {backup_path}")

        except ImportError as ie:
            print(f"\n❌ 缺少依赖: {ie}")
            print("请运行: pip install openpyxl")
        except Exception as e:
            print(f"\n❌ Excel 处理失败: {e}")
            import traceback
            traceback.print_exc()

        return

    # ===== PPTX 工作流 =====
    is_pptx_workflow = (translated_ext == ".pptx")
    if is_pptx_workflow:
        print("\n--- 阶段 2: PPTX 替换与批注 ---")
        try:
            from llm.llm_project.backup_copy.backup_manager import ensure_backup_copy as ensure_backup

            def load_errors(label, path):
                if path and os.path.exists(path):
                    data = extract_and_parse(path)
                    print(f"✓ 已加载{label}报告: {len(data)} 条错误")
                    return data
                return []

            body_errors = load_errors("正文", report_paths.get("正文"))
            header_errors = load_errors("页眉", report_paths.get("页眉"))
            footer_errors = load_errors("页脚", report_paths.get("页脚"))
            all_errors = body_errors + header_errors + footer_errors

            if not all_errors:
                print("✓ 未发现错误，无需修改")
                return

            # 输出为Excel报告
            print("\n--- 生成Excel报告 ---")
            eg = ExcelReportGenerator()
            body_target_folder = r"D:\Users\Administrator\Desktop\项目文件\专检\数值检查1\llm\llm_project\zhengwen\output_excel"
            body_file_path, _ = eg.get_excel(body_errors, output_dir=body_target_folder)
            print(f"✓ 正文报告已存至: {body_file_path}")

            header_target_folder = r"D:\Users\Administrator\Desktop\项目文件\专检\数值检查1\llm\llm_project\yemei\output_excel"
            header_file_path, _ = eg.get_excel(header_errors, output_dir=header_target_folder)
            print(f"✓ 页眉报告已存至: {header_file_path}")

            footer_target_folder = r"D:\Users\Administrator\Desktop\项目文件\专检\数值检查1\llm\llm_project\yejiao\output_excel"
            footer_file_path, _ = eg.get_excel(footer_errors, output_dir=footer_target_folder)
            print(f"✓ 页脚报告已存至: {footer_file_path}")

            # 备份译文
            print("\n📦 正在备份原始 PPTX 文件...")
            backup_path = ensure_backup(args.translated)
            print(f"✓ 备份完成: {backup_path}")

            # 执行替换与批注
            print(f"\n正在为 {len(all_errors)} 个错误执行替换与批注...")
            replacer = PPTXReplacer(backup_path)
            success_count = 0
            fail_count = 0
            failed_items = []

            for idx, error in enumerate(all_errors, 1):
                old_text = (error.get("译文数值") or "").strip()
                new_text = (error.get("译文修改建议值") or "").strip()
                reason = error.get("修改理由", "数值错误")
                context = error.get("译文上下文", "")

                if not old_text or not new_text:
                    continue

                ok = replacer.replace_and_annotate(
                    old_text=old_text,
                    new_text=new_text,
                    reason=reason,
                    context=context,
                )

                if ok:
                    success_count += 1
                    print(f"  [{idx}/{len(all_errors)}] ✓ '{old_text}' → '{new_text}'")
                else:
                    fail_count += 1
                    failed_items.append({'text': old_text, 'suggestion': new_text, 'reason': reason})
                    print(f"  [{idx}/{len(all_errors)}] ✗ 未找到: '{old_text}'")

            replacer.save(str(backup_path))

            print(f"\n--- PPTX 替换统计 ---")
            print(f"成功: {success_count}")
            print(f"失败: {fail_count}")
            print(f"总计: {len(all_errors)}")
            if success_count + fail_count > 0:
                print(f"成功率: {success_count / (success_count + fail_count):.2%}")

            if failed_items:
                fail_dir = Path(__file__).parent / "llm" / "llm_project" / "fail"
                fail_dir.mkdir(parents=True, exist_ok=True)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                fail_report_path = fail_dir / f"{Path(args.translated).stem}_failed_pptx_{timestamp}.txt"
                with open(fail_report_path, 'w', encoding='utf-8') as f:
                    f.write("PPTX 替换失败项目\n" + "=" * 60 + "\n\n")
                    for i, item in enumerate(failed_items, 1):
                        f.write(
                            f"{i}. 查找: {item['text']}\n   建议: {item['suggestion']}\n   理由: {item['reason']}\n\n")
                print(f"✓ 失败项详情已保存至: {fail_report_path}")

            print(f"\n✅ PPTX 处理完成，结果保存至: {backup_path}")

        except ImportError as ie:
            print(f"\n❌ 缺少依赖: {ie}")
            print("请运行: pip install python-pptx")
        except Exception as e:
            print(f"\n❌ PPTX 处理失败: {e}")
            import traceback
            traceback.print_exc()

        return

    # 4) 核心修复逻辑（仅适用于 Word 文档）
    print("\n--- 阶段 2: 自动替换与修订 ---")

    # 确保译文是 Word 文档
    if translated_ext not in [".docx", ".doc"]:
        print("❌ 错误: 阶段2仅支持 Word 文档格式的译文")
        return

    # 创建备份
    from backup_copy.backup_manager import ensure_backup_copy
    backup_copy_path = ensure_backup_copy(args.translated)
    doc = Document(backup_copy_path)
    revision_manager = RevisionManager(doc, author="翻译校对")
    print("✓ 修订模式已启用（Track Changes）")

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

    # 输出为Excel
    print("\n--- 生成Excel报告 ---")
    e = ExcelReportGenerator()

    # 正文excel
    body_target_folder = r"D:\Users\Administrator\Desktop\项目文件\专检\数值检查1\llm\llm_project\zhengwen\output_excel"
    body_file_path, body_data = e.get_excel(body_errors, output_dir=body_target_folder)
    print(f"✓ 正文报告已存至: {body_file_path}")

    # 页眉excel
    header_target_folder = r"D:\Users\Administrator\Desktop\项目文件\专检\数值检查1\llm\llm_project\yemei\output_excel"
    header_file_path, header_data = e.get_excel(header_errors, output_dir=header_target_folder)
    print(f"✓ 页眉报告已存至: {header_file_path}")

    # 页脚excel
    footer_target_folder = r"D:\Users\Administrator\Desktop\项目文件\专检\数值检查1\llm\llm_project\yejiao\output_excel"
    footer_file_path, footer_data = e.get_excel(footer_errors, output_dir=footer_target_folder)
    print(f"✓ 页脚报告已存至: {footer_file_path}")

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
        if not errors: return 0, 0, 0, [], []
        print(f"\n>>> 正在修复 {label} 部分...")
        s_count, f_count = 0, 0
        skip_count = 0
        change_records = []  # 记录所有成功的修改
        failed_items = []  # 收集失败的项目

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
            from replace.replace_clean import is_list_pattern
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

            ok, strategy = replace_and_revise_in_docx(
                doc, old, new, reason, revision_manager,
                context=context, anchor_text=anchor, region=region
            )
            if ok:
                s_count += 1
                change_records.append(f"'{old}' → '{new}'")
                print(f"  [{idx}] 成功: '{old}' -> '{new}'")
                print(f"    修改理由: {reason}")
                print(f"    策略: {strategy}")
                print(f"    操作: '{old}' → '{new}' (修订模式)")

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
                failed_items.append({
                    "原文数值": e.get("原文数值", ""),
                    "译文数值": old,
                    "译文修改建议值": new,
                    "修改理由": reason,
                    "原文上下文": e.get("原文上下文", ""),
                    "译文上下文": context,
                    "替换锚点": anchor
                })
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

        return s_count, f_count, skip_count, change_records, failed_items

    # 执行三部分修复（指定对应区域）
    b_s, b_f, b_skip, body_changes, body_failed = apply_all_fixes(body_errors, "正文", region="body")
    h_s, h_f, h_skip, header_changes, header_failed = apply_all_fixes(header_errors, "页眉", region="header")
    f_s, f_f, f_skip, footer_changes, footer_failed = apply_all_fixes(footer_errors, "页脚", region="footer")

    # 页眉页脚修改汇总
    print("\n" + "=" * 60)
    print("页眉页脚修改汇总")
    print("=" * 60)

    if h_s == 0 and h_f == 0 and h_skip == 0:
        print("  ℹ️  页眉：无修改")
    else:
        print(f"\n  ✓ 页眉：共修改 {h_s} 处")
        if header_changes:
            print("  详细修改:")
            for i, change in enumerate(header_changes, 1):
                print(f"    {i}. {change}")

    if f_s == 0 and f_f == 0 and f_skip == 0:
        print("\n  ℹ️  页脚：无修改")
    else:
        print(f"\n  ✓ 页脚：共修改 {f_s} 处")
        if footer_changes:
            print("  详细修改:")
            for i, change in enumerate(footer_changes, 1):
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

    # 调试：打印失败项数量
    print(f"\n[调试] 正文失败: {len(body_failed)}, 页眉失败: {len(header_failed)}, 页脚失败: {len(footer_failed)}")

    # 保存失败项到 fail/ 文件夹
    all_failed = body_failed + header_failed + footer_failed
    if all_failed:
        # 创建 fail 文件夹
        fail_dir = Path("llm/llm_project/fail")
        fail_dir.mkdir(parents=True, exist_ok=True)

        # 生成失败报告文件名
        translated_filename = Path(args.translated).stem
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        fail_report_path = fail_dir / f"{translated_filename}_failed_word_{timestamp}.txt"

        # 写入失败报告
        with open(fail_report_path, 'w', encoding='utf-8') as f:
            f.write("=" * 80 + "\n")
            f.write("Word 批注失败报告\n")
            f.write("=" * 80 + "\n")
            f.write(f"文件: {Path(args.translated).name}\n")
            f.write(f"时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"失败数量: {len(all_failed)}\n")
            f.write("=" * 80 + "\n\n")

            for i, item in enumerate(all_failed, 1):
                f.write(f"[{i}] 失败项\n")
                f.write(f"  原文数值: {item['原文数值']}\n")
                f.write(f"  译文数值: {item['译文数值']}\n")
                f.write(f"  修改建议: {item['译文修改建议值']}\n")
                f.write(f"  修改理由: {item['修改理由']}\n")
                f.write(f"  原文上下文: {item['原文上下文']}\n")
                f.write(f"  译文上下文: {item['译文上下文']}\n")
                f.write(f"  替换锚点: {item['替换锚点']}\n")
                f.write("\n" + "-" * 80 + "\n\n")

        print(f"\n⚠️ 以下 {len(all_failed)} 项未能自动标注（可能需要人工检查）:")
        for item in all_failed:
            print(f"  - '{item['译文数值']}' → '{item['译文修改建议值']}'")
        print(f"\n✓ 失败项详情已保存至: {fail_report_path.absolute()}")

    # 保存最终结果
    doc.save(backup_copy_path)
    print(f"\n" + "=" * 40)

    print(f"🎉 全部流程处理完成！")
    print(f"最终结果保存至: {backup_copy_path}")
    print("=" * 40)


if __name__ == '__main__':
    main()
