from datetime import datetime
import os
import argparse
import re
from docx import Document
from pathlib import Path
from parsers.pdf.pdf_parser import parse_pdf
from llm_check.rule_generation import rule
from llm_check.check import Match
from llm_check.check_2 import Match as BilingualMatch
from parsers.word.body_extractor import extract_body_text
from parsers.word.footer_extractor import extract_footers
from parsers.word.header_extractor import extract_headers
from parsers.json.clean_json import extract_and_parse
from utils.json_files import write_json_with_timestamp
from replace.word.replace_revision import replace_and_revise_in_docx, flush_footnote_replacements
from replace.word.revision import RevisionManager
from replace.word.replace_clean import is_list_pattern
from parsers.txt.txt_parser import parse_txt
from utils.txt_files import write_txt_with_timestamp
from parsers.excel.excel装载 import ExcelReportGenerator
from parsers.pptx.pptx_parser import parse_pptx
from divide.text_splitter import split_text_pair, split_bilingual_text, _count_chars

# 如果不想用命令行参数，可以直接修改下面的变量
USE_AI_RULE_CONFIG = False  # 👈 改为 True 使用 AI 生成规则
BILINGUAL_CONFIG = False  # 👈 改为 True 使用双语对照模式
# 项目根目录（main.py 所在目录）
BASE_DIR = Path(__file__).resolve().parent
# LLM 项目目录
LLM_DIR = BASE_DIR


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
    自动识别文档格式并提取全文内容（页眉+正文+页脚）
    支持：PDF、Word、Excel、PPTX
    """
    path = Path(file_path)
    if not path.exists():
        return "", "", ""

    suffix = path.suffix.lower()

    if suffix == ".pdf":
        print(f"📄 正在解析 PDF 文档: {path.name}")
        full_text = parse_pdf(str(path))
        return full_text, "", ""

    elif suffix in [".docx", ".doc"]:
        print(f"📝 正在解析 Word 文档: {path.name}")
        body = extract_body_text(str(path))
        header = extract_headers(str(path))
        footer = extract_footers(str(path))
        return body, header, footer

    elif suffix in [".xlsx", ".xls"]:
        print(f"📊 正在解析 Excel 文档: {path.name}")
        try:
            from parsers.excel.excel_parser import parse_excel_with_pandas
            full_text = parse_excel_with_pandas(str(path))
            return full_text, "", ""
        except Exception as e:
            print(f"❌ Excel 解析失败: {e}")
            return "", "", ""

    elif suffix == ".pptx":
        print(f"📊 正在解析 PPTX 文档: {path.name}")
        try:
            full_text = parse_pptx(str(path))
            return full_text or "", "", ""
        except Exception as e:
            print(f"❌ PPTX 解析失败: {e}")
            return "", "", ""

    else:
        print(f"⚠️ 不支持的文档格式: {suffix}")
        return "", "", ""


def _compare_with_split(matcher, orig_txt, tran_txt, rule_text, name):
    """对一组原文/译文文本进行分块对比，合并结果。

    如果文本较短则直接对比；较长则自动分割为对齐的块，
    逐块调用 LLM，最后合并去重返回完整结果列表。

    Returns:
        list[dict] — 始终返回解析后的错误列表（可能为空列表）
    """
    from parsers.json.clean_json import parse_json_content
    import time

    pairs = split_text_pair(orig_txt, tran_txt)
    num_parts = len(pairs)

    if num_parts <= 1:
        print(f"  📄 {name}文本较短（{_count_chars(orig_txt)} 字），直接对比")
        raw = None
        for attempt in range(3):
            try:
                raw = matcher.compare_texts(orig_txt, tran_txt, rule_text)
                break
            except Exception as e:
                if attempt < 2:
                    wait = (attempt + 1) * 10
                    print(f"  ⚠️ 调用失败: {e}，{wait}秒后重试...")
                    time.sleep(wait)
                else:
                    print(f"  ❌ 调用 API 失败（已重试 2 次）: {e}")
        if raw:
            parsed = parse_json_content(raw)
            return parsed if isinstance(parsed, list) else []
        return []

    print(f"  📄 {name}文本较长（原文 {_count_chars(orig_txt)} 字），分割为 {num_parts} 块进行对比")

    all_errors = []
    seen_keys = set()

    for i, (orig_chunk, trans_chunk) in enumerate(pairs, 1):
        print(f"\n  --- {name} 第 {i}/{num_parts} 块 ---")
        print(f"      原文块: {_count_chars(orig_chunk)} 字, 译文块: {_count_chars(trans_chunk)} 字")

        if not orig_chunk.strip() or not trans_chunk.strip():
            print(f"      ⚠️ 块内容为空，跳过")
            continue

        max_retries = 2
        chunk_result = None
        for attempt in range(max_retries + 1):
            try:
                chunk_result = matcher.compare_texts(orig_chunk, trans_chunk, rule_text)
                break
            except Exception as e:
                if attempt < max_retries:
                    wait = (attempt + 1) * 10
                    print(f"      ⚠️ 第 {i} 块第 {attempt + 1} 次调用失败: {e}")
                    print(f"      ⏳ 等待 {wait} 秒后重试...")
                    time.sleep(wait)
                else:
                    print(f"      ❌ 第 {i} 块调用 API 失败（已重试 {max_retries} 次）: {e}")
                    print(f"      ⚠️ 跳过该块，继续处理剩余块")
                    chunk_result = None

        chunk_count = 0
        if chunk_result:
            parsed = parse_json_content(chunk_result)
            if isinstance(parsed, list):
                chunk_count = len(parsed)
                for item in parsed:
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

    for idx, item in enumerate(all_errors, 1):
        item["错误编号"] = str(idx)

    return all_errors


def _compare_bilingual_with_split(matcher, text, rule_text, name):
    """对双语对照文本进行分块对比，合并结果。

    Returns:
        list[dict] — 始终返回解析后的错误列表（可能为空列表）
    """
    from parsers.json.clean_json import parse_json_content
    import time

    chunks = split_bilingual_text(text)
    num_parts = len(chunks)

    if num_parts <= 1:
        print(f"  📄 {name}文本较短（{len(text)} 字符），直接对比")
        raw = None
        for attempt in range(3):
            try:
                raw = matcher.compare_texts(text, rule_text)
                break
            except Exception as e:
                if attempt < 2:
                    wait = (attempt + 1) * 10
                    print(f"  ⚠️ 调用失败: {e}，{wait}秒后重试...")
                    time.sleep(wait)
                else:
                    print(f"  ❌ 调用 API 失败（已重试 2 次）: {e}")
        if raw:
            parsed = parse_json_content(raw)
            return parsed if isinstance(parsed, list) else []
        return []

    print(f"  📄 {name}文本较长（{len(text)} 字符），分割为 {num_parts} 块进行对比")

    all_errors = []
    seen_keys = set()

    for i, chunk in enumerate(chunks, 1):
        print(f"\n  --- {name} 第 {i}/{num_parts} 块 ---")
        print(f"      块大小: {len(chunk)} 字符")

        if not chunk.strip():
            print(f"      ⚠️ 块内容为空，跳过")
            continue

        max_retries = 2
        chunk_result = None
        for attempt in range(max_retries + 1):
            try:
                chunk_result = matcher.compare_texts(chunk, rule_text)
                break
            except Exception as e:
                if attempt < max_retries:
                    wait = (attempt + 1) * 10
                    print(f"      ⚠️ 第 {i} 块第 {attempt + 1} 次调用失败: {e}")
                    print(f"      ⏳ 等待 {wait} 秒后重试...")
                    time.sleep(wait)
                else:
                    print(f"      ❌ 第 {i} 块调用 API 失败（已重试 {max_retries} 次）: {e}")
                    print(f"      ⚠️ 跳过该块，继续处理剩余块")
                    chunk_result = None

        chunk_count = 0
        if chunk_result:
            parsed = parse_json_content(chunk_result)
            if isinstance(parsed, list):
                chunk_count = len(parsed)
                for item in parsed:
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

    for idx, item in enumerate(all_errors, 1):
        item["错误编号"] = str(idx)

    return all_errors


SECTION_DIRS = {
    "正文": "zhengwen",
    "页眉": "yemei",
    "页脚": "yejiao",
}


def _resolve_section_output_dir(output_base_dir, section_name, leaf_name):
    if output_base_dir:
        base_dir = Path(output_base_dir)
    else:
        base_dir = LLM_DIR
    target_dir = base_dir / SECTION_DIRS[section_name] / leaf_name
    target_dir.mkdir(parents=True, exist_ok=True)
    return target_dir


def _build_parts(
    original_body,
    original_header,
    original_footer,
    translated_body,
    translated_header,
    translated_footer,
    bilingual,
    output_base_dir=None,
):
    if bilingual:
        return [
            ("正文", original_body, None, str(_resolve_section_output_dir(output_base_dir, "正文", "output_json"))),
            ("页眉", original_header, None, str(_resolve_section_output_dir(output_base_dir, "页眉", "output_json"))),
            ("页脚", original_footer, None, str(_resolve_section_output_dir(output_base_dir, "页脚", "output_json"))),
        ]
    return [
        ("正文", original_body, translated_body, str(_resolve_section_output_dir(output_base_dir, "正文", "output_json"))),
        ("页眉", original_header, translated_header, str(_resolve_section_output_dir(output_base_dir, "页眉", "output_json"))),
        ("页脚", original_footer, translated_footer, str(_resolve_section_output_dir(output_base_dir, "页脚", "output_json"))),
    ]


def run_comparison(
    original_path,
    translated_path,
    use_ai_rule=False,
    bilingual=False,
    output_base_dir=None,
    ai_rule_file_path=None,
    session_rule_text=None,
):
    """
    第一阶段：提取文本并调用 AI/Matcher 进行对比，生成 JSON 报告

    Args:
        original_path: 原文文档路径（双语模式下为双语对照文件路径）
        translated_path: 译文文档路径（双语模式下不使用）
        use_ai_rule: 是否使用 AI 生成规则（默认 False，使用自定义或默认规则）
        bilingual: 是否为双语对照模式（单文件，原文译文在同一文档中）
        output_base_dir: 可选输出目录，供 Web 任务使用
        ai_rule_file_path: 可选规则文件路径，供 AI 生成规则
        session_rule_text: 可选的会话规则文本，仅本次任务使用
    """
    print("\n--- 阶段 1: 文本提取与 AI 对比 ---")
    print("📋 模式: 中英双语对照（单文件）" if bilingual else "📋 模式: 原文+译文分开对比")

    try:
        if bilingual:
            bilingual_body, bilingual_header, bilingual_footer = extract_any_document(original_path)
            original_body = bilingual_body
            original_header = bilingual_header
            original_footer = bilingual_footer
            translated_body = translated_header = translated_footer = None
        else:
            original_body, original_header, original_footer = extract_any_document(original_path)
            translated_body, translated_header, translated_footer = extract_any_document(translated_path)
    except Exception as e:
        print(f"❌ 解析失败: {e}")
        print("   不支持的格式，请检查文件是否损坏。")
        return None, None

    matcher = Match() if not bilingual else BilingualMatch()
    parts = _build_parts(
        original_body,
        original_header,
        original_footer,
        translated_body,
        translated_header,
        translated_footer,
        bilingual,
        output_base_dir=output_base_dir,
    )

    report_paths = {}
    custom_path = str(LLM_DIR / "rule" / "自定义规则.txt")
    default_path = str(LLM_DIR / "rule" / "默认规则.txt")
    rule_backup_dir = str(LLM_DIR / "rule")

    if session_rule_text and session_rule_text.strip():
        rule_text = session_rule_text.strip()
        print("\nℹ️  使用本次会话编辑的规则（不会修改磁盘文件）")
    elif use_ai_rule:
        try:
            rule_gen = rule()
            default_ai_rule_path = LLM_DIR / "rule" / "银行稿件_翻译规则_通用_220307.pdf"
            ai_rule_path = ai_rule_file_path or (str(default_ai_rule_path) if default_ai_rule_path.exists() else None)
            if not ai_rule_path or not os.path.exists(ai_rule_path):
                raise FileNotFoundError(f"AI 规则文件不存在: {ai_rule_path}")

            ai_rule_text = load_any_rule(ai_rule_path)
            print("正在调用 AI 生成规则...")
            ai_rule = rule_gen.compare_texts(ai_rule_text)

            with open(custom_path, 'w', encoding='utf-8') as f:
                f.write(ai_rule)
            print(f"✅ AI 生成的规则已保存到: {custom_path}")

            _, backup_path = write_txt_with_timestamp(ai_rule, rule_backup_dir)
            print(f"📦 备份已保存到: {backup_path}")
            rule_text = ai_rule
        except Exception as e:
            print(f"❌ AI 生成规则失败: {e}")
            print("⚠️  将尝试使用自定义规则或默认规则...")
            rule_text = parse_txt(custom_path) or parse_txt(default_path)
    else:
        print("\nℹ️  用户选择使用现有规则（不调用 AI）")
        rule_text = parse_txt(custom_path) or parse_txt(default_path)

    if not rule_text:
        print("❌ 错误：无法加载任何有效规则。")
        print("   请确保以下文件之一存在且有效：")
        print(f"   - 自定义规则: {custom_path}")
        print(f"   - 默认规则: {default_path}")
        print("   或使用 --use-ai-rule 参数让 AI 生成规则")
        return None, None

    if session_rule_text and session_rule_text.strip():
        used_type = "会话规则"
    elif use_ai_rule:
        used_type = "AI 生成规则"
    else:
        used_type = "自定义规则" if parse_txt(custom_path) else "默认规则"

    print(f"✅ 规则加载成功，当前使用: {used_type}")
    if not use_ai_rule and not (session_rule_text and session_rule_text.strip()):
        print(f"   规则文件: {custom_path if parse_txt(custom_path) else default_path}")
    print(f"--- 规则内容预览 (前200字符) ---")
    print(rule_text[:200] + "..." if len(rule_text) > 200 else rule_text)

    for name, orig_txt, tran_txt, out_dir in parts:
        print(f"====== 正在检查{name} ===========")
        if bilingual:
            if orig_txt:
                try:
                    res = _compare_bilingual_with_split(matcher, orig_txt, rule_text, name)
                except Exception as e:
                    print(f"❌ 调用 API 失败: {e}")
                    print("   请检查账户余额、API Key 有效性或网络环境是否正常。")
                    return None, None
            else:
                res = []
                print(f"⚠️ {name}内容为空")
        else:
            if orig_txt and tran_txt:
                try:
                    res = _compare_with_split(matcher, orig_txt, tran_txt, rule_text, name)
                except Exception as e:
                    print(f"❌ 调用 API 失败: {e}")
                    print("   请检查账户余额、API Key 有效性或网络环境是否正常。")
                    return None, None
            else:
                res = []
                print(f"⚠️ {name}原文或译文为空")

        _, path = write_json_with_timestamp(res, out_dir)
        report_paths[name] = path

    return report_paths, rule_text


def _load_report_errors(report_paths):
    section_errors = {}
    for section_name in SECTION_DIRS:
        path = report_paths.get(section_name)
        if path and os.path.exists(path):
            data = extract_and_parse(path)
            print(f"✓ 已加载{section_name}报告: {len(data)} 条错误")
            section_errors[section_name] = data
        else:
            section_errors[section_name] = []
    return section_errors


def _export_excel_reports(section_errors, output_base_dir=None):
    excel_paths = {}
    exporter = ExcelReportGenerator()
    for section_name, errors in section_errors.items():
        target_dir = _resolve_section_output_dir(output_base_dir, section_name, "output_excel")
        file_path, _ = exporter.get_excel(errors, output_dir=str(target_dir))
        excel_paths[section_name] = file_path
        if file_path:
            print(f"✓ {section_name}报告已导出 Excel: {file_path}")
    return excel_paths


def _apply_word_fixes(doc, revision_manager, errors, label, region, doc_path):
    if not errors:
        return {"success": 0, "failed": 0, "skipped": 0}

    print(f"\n>>> 正在修复 {label} 部分...")
    success_count = 0
    failed_count = 0
    skipped_count = 0
    processed_numbering_formats = set()

    for idx, error in enumerate(errors, 1):
        old = (error.get("译文数值") or "").strip()
        new = (error.get("译文修改建议值") or "").strip()
        reason = str(error.get("修改理由") or "数值错误").strip()
        context = error.get("译文上下文", "")
        anchor = error.get("替换锚点", "")

        if not old or not new:
            print(f"  [{idx}] 跳过: 缺少【译文数值】或【译文修改建议值】字段")
            skipped_count += 1
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
                success_count += 1
                print(f"  [{idx}] 成功: '{old}' -> '{new}' (已由编号批量替换完成)")
                continue

        ok, strategy = replace_and_revise_in_docx(
            doc,
            old,
            new,
            reason,
            revision_manager,
            context=context,
            anchor_text=anchor,
            region=region,
            doc_path=str(doc_path),
        )
        if ok:
            success_count += 1
            print(f"  [{idx}] 成功: '{old}' -> '{new}'")
            print(f"    修改理由: {reason}")
            print(f"    策略: {strategy}")
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
            failed_count += 1
            print(f"  [{idx}] 失败: 未匹配到 '{old}'")

    return {"success": success_count, "failed": failed_count, "skipped": skipped_count}


def _run_word_fix_phase(translated_path, section_errors):
    print("\n--- 阶段 2: 自动替换与批注 ---")
    from backup_copy.backup_manager import ensure_backup_copy

    backup_copy_path = ensure_backup_copy(translated_path)
    try:
        from replace.word.numbering_to_static import convert_numbering_to_static, has_auto_numbering
        if has_auto_numbering(str(backup_copy_path)):
            print("\n🔢 检测到自动编号，正在转换为静态文本...")
            ok = convert_numbering_to_static(str(backup_copy_path))
            if ok:
                print("✓ 自动编号已转为静态文本")
            else:
                print("⚠ 自动编号静态化失败，部分编号可能无法替换")
    except Exception as e:
        print(f"⚠ 编号静态化异常: {e}")

    doc = Document(backup_copy_path)
    doc._numbering_staticized = True
    revision_manager = RevisionManager(doc)

    body_stat = _apply_word_fixes(doc, revision_manager, section_errors.get("正文", []), "正文", "body", backup_copy_path)
    header_stat = _apply_word_fixes(doc, revision_manager, section_errors.get("页眉", []), "页眉", "header", backup_copy_path)
    footer_stat = _apply_word_fixes(doc, revision_manager, section_errors.get("页脚", []), "页脚", "footer", backup_copy_path)

    doc.save(backup_copy_path)
    footnote_count = flush_footnote_replacements(doc, str(backup_copy_path))
    if footnote_count > 0:
        print(f"✓ 脚注替换完成: {footnote_count} 处")

    stats = {
        "success": body_stat["success"] + header_stat["success"] + footer_stat["success"],
        "failed": body_stat["failed"] + header_stat["failed"] + footer_stat["failed"],
        "skipped": body_stat["skipped"] + header_stat["skipped"] + footer_stat["skipped"],
    }
    print(f"🎉 Word 修复完成，结果保存至: {backup_copy_path}")
    return backup_copy_path, stats


def _run_pdf_fix_phase(translated_path, section_errors):
    print("\n--- 阶段 2: PDF 批注与替换 ---")
    from backup_copy.backup_manager import ensure_backup_copy
    from replace.pdf_replacer_improved import ImprovedPDFReplacer

    all_errors = section_errors.get("正文", []) + section_errors.get("页眉", []) + section_errors.get("页脚", [])
    if not all_errors:
        print("✓ 未发现错误，无需添加批注")
        return None, {"success": 0, "failed": 0, "skipped": 0}

    backup_path = ensure_backup_copy(translated_path, suffix="annotated")
    debug_mode = os.environ.get('PDF_ANNOTATE_DEBUG', '').lower() == 'true'
    success_count = 0
    failed_count = 0
    skipped_count = 0

    with ImprovedPDFReplacer(backup_path) as replacer:
        for idx, error in enumerate(all_errors, 1):
            old_text = (error.get("译文数值") or "").strip()
            new_text = (error.get("译文修改建议值") or "").strip()
            reason = error.get("修改理由", "数值错误")
            error_type = error.get("错误类型", "数值错误")
            context = error.get("译文上下文", "")

            if not old_text or not new_text:
                skipped_count += 1
                print(f"  [{idx}/{len(all_errors)}] ⊘ 跳过: 缺少译文数值或修改建议")
                continue

            color = (0, 1, 0) if "数值错误" in error_type else (0, 0.8, 0)
            comment = (
                f"【{error_type}】已修改\n"
                f"原文: {error.get('原文数值', 'N/A')}\n"
                f"原译文: {old_text}\n"
                f"已改为: {new_text}\n"
                f"理由: {reason}"
            )
            repl_success, _, _ = replacer.replace_and_annotate(
                search_text=old_text,
                new_text=new_text,
                comment=comment,
                context=context,
                color=color,
                debug=debug_mode,
            )
            if repl_success > 0:
                success_count += 1
                print(f"  [{idx}/{len(all_errors)}] ✓ 已替换并批注: '{old_text}' → '{new_text}'")
            else:
                failed_count += 1
                print(f"  [{idx}/{len(all_errors)}] ✗ 未找到或替换失败: {old_text}")

        replacer.save(str(backup_path))

    return backup_path, {"success": success_count, "failed": failed_count, "skipped": skipped_count}


def _run_sheet_fix_phase(translated_path, section_errors, kind_name, replacer_cls):
    print(f"\n--- 阶段 2: {kind_name} 批注与替换 ---")
    from backup_copy.backup_manager import ensure_backup_copy

    all_errors = section_errors.get("正文", []) + section_errors.get("页眉", []) + section_errors.get("页脚", [])
    if not all_errors:
        print(f"✓ 未发现错误，无需处理 {kind_name}")
        return None, {"success": 0, "failed": 0, "skipped": 0}

    backup_path = ensure_backup_copy(translated_path, suffix="annotated")
    replacer = replacer_cls(str(backup_path))
    success_count = 0
    failed_count = 0
    skipped_count = 0

    for idx, error in enumerate(all_errors, 1):
        old_text = (error.get("译文数值") or "").strip()
        new_text = (error.get("译文修改建议值") or "").strip()
        reason = error.get("修改理由", "数值错误")
        context = error.get("译文上下文", "")

        if not old_text or not new_text:
            skipped_count += 1
            print(f"  [{idx}/{len(all_errors)}] ⊘ 跳过: 缺少译文数值或修改建议")
            continue

        ok = replacer.replace_and_annotate(
            old_text=old_text,
            new_text=new_text,
            reason=reason,
            context=context,
        )
        if ok:
            success_count += 1
            print(f"  [{idx}/{len(all_errors)}] ✓ 已替换并批注: '{old_text}' → '{new_text}'")
        else:
            failed_count += 1
            print(f"  [{idx}/{len(all_errors)}] ✗ 未找到或替换失败: {old_text}")

    replacer.save(str(backup_path))
    return backup_path, {"success": success_count, "failed": failed_count, "skipped": skipped_count}


def run_full_pipeline(
    original_path,
    translated_path,
    output_base_dir,
    use_ai_rule=False,
    bilingual=False,
    ai_rule_file_path=None,
    session_rule_text=None,
):
    """
    完整流程：对比 + 导出报告 + 修复/批注。供 Web 或脚本调用。

    Returns:
        (result_path, report_paths, excel_paths, stats_dict)
    """
    if bilingual:
        translated_path = original_path

    report_paths, _ = run_comparison(
        original_path,
        translated_path,
        use_ai_rule=use_ai_rule,
        bilingual=bilingual,
        output_base_dir=output_base_dir,
        ai_rule_file_path=ai_rule_file_path,
        session_rule_text=session_rule_text,
    )
    if report_paths is None:
        return None, None, None, None

    section_errors = _load_report_errors(report_paths)
    excel_paths = _export_excel_reports(section_errors, output_base_dir=output_base_dir)

    translated_ext = Path(translated_path).suffix.lower()
    original_ext = Path(original_path).suffix.lower()
    is_pdf_workflow = (original_ext == ".pdf" or translated_ext == ".pdf")
    is_excel_workflow = translated_ext in [".xlsx", ".xls"]
    is_pptx_workflow = translated_ext == ".pptx"

    if is_pdf_workflow:
        result_path, stats = _run_pdf_fix_phase(translated_path, section_errors)
    elif is_excel_workflow:
        from replace.excel.excel_replacer import ExcelReplacer
        result_path, stats = _run_sheet_fix_phase(translated_path, section_errors, "Excel", ExcelReplacer)
    elif is_pptx_workflow:
        from replace.pptx.pptx_replacer import PPTXReplacer
        result_path, stats = _run_sheet_fix_phase(translated_path, section_errors, "PPTX", PPTXReplacer)
    else:
        result_path, stats = _run_word_fix_phase(translated_path, section_errors)

    return result_path, report_paths, excel_paths, stats


def main():
    # 1) 配置默认路径
    DEFAULT_ORIGINAL = str(
        BASE_DIR / "测试文件" / "原文-B260328387-关于消费者权益保护2025年工作情况与2026年工作计划的议案 - 副本(1).docx")
    DEFAULT_TRANSLATED = str(
        BASE_DIR / "测试文件" / "译文-B260328387-关于消费者权益保护2025年工作情况与2026年工作计划的议案.docx")

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

  # 中英双语对照模式（单文件包含原文和译文）
  python llm/main_ddd.py -o "双语对照.docx" --bilingual

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
    parser.add_argument("--bilingual", "-b", action="store_true",
                        help="中英双语对照模式（单文件包含原文和译文，无需分开上传）")
    args = parser.parse_args()

    # ⚙️ 不使用命令行时的配置（在这里修改）


    # 优先使用命令行参数，如果没有命令行参数则使用配置变量
    use_ai_rule = args.use_ai_rule or USE_AI_RULE_CONFIG
    bilingual = args.bilingual or BILINGUAL_CONFIG

    if bilingual:
        # 双语模式只需要一个文件
        if not os.path.exists(args.original):
            print("❌ 错误: 双语对照文件路径不存在")
            return
        print(f"📋 双语对照模式: {args.original}")
    else:
        if not os.path.exists(args.original) or not os.path.exists(args.translated):
            print("❌ 错误: 输入的 docx 文件路径不存在")
            return

    # 检查文件格式
    original_ext = Path(args.original).suffix.lower()
    translated_ext = Path(args.translated).suffix.lower() if not bilingual else original_ext

    # 双语模式下，译文路径就是原文路径（同一个文件）
    if bilingual:
        args.translated = args.original

    # 判断工作流类型
    is_pdf_workflow = (original_ext == ".pdf" or translated_ext == ".pdf")
    is_excel_workflow = (translated_ext in [".xlsx", ".xls"])
    is_pptx_workflow = (translated_ext == ".pptx")

    # 如果是 PDF 文件，使用批注功能
    if is_pdf_workflow:
        print("\n⚠️ 检测到 PDF 文件")
        # 3) 执行对比并获取生成的 JSON 路径

        report_paths, _ = run_comparison(
            args.original,
            args.translated,
            bilingual=bilingual,
        )

        # 检查返回值
        if report_paths is None:
            print("❌ 报告生成失败，程序终止")
            return

        print(f"\n✅ 报告已生成")
        print("对比结果已保存至:")
        for name, path in report_paths.items():
            print(f"  - {name}: {path}")

        # 如果译文是 PDF，尝试添加批注
        if translated_ext == ".pdf":
            print("\n--- 阶段 2: PDF 批注与替换 ---")
            try:
                from llm.llm_project.replace.pdf_replacer_improved import ImprovedPDFReplacer
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
                body_target_folder = str(LLM_DIR / "zhengwen" / "output_excel")
                body_file_path, body_data = e.get_excel(body_errors, output_dir=body_target_folder)
                print(f"✓ 正文报告已存至: {body_file_path}")

                # 页眉excel
                header_target_folder = str(LLM_DIR / "yemei" / "output_excel")
                header_file_path, header_data = e.get_excel(header_errors, output_dir=header_target_folder)
                print(f"✓ 页眉报告已存至: {header_file_path}")

                # 页脚excel
                footer_target_folder = str(LLM_DIR / "yejiao" / "output_excel")
                footer_file_path, footer_data = e.get_excel(footer_errors, output_dir=footer_target_folder)
                print(f"✓ 页脚报告已存至: {footer_file_path}")

                # 先备份原文件到 backup 文件夹
                print("\n📦 正在备份原始 PDF 文件...")
                backup_path = ensure_backup_copy(args.translated, suffix="annotated")
                print(f"✓ 备份完成，将对备份文件进行批注")

                # 对备份文件进行批注（不是原文件）
                output_path = backup_path

                # 打开备份文件进行批注和替换（使用改进的替换器）
                with ImprovedPDFReplacer(backup_path) as replacer:
                    success_count = 0
                    fail_count = 0
                    skip_count = 0  # 记录跳过的数量
                    replace_count = 0  # 记录替换成功的数量
                    annot_count = 0  # 记录批注成功的数量
                    failed_items = []  # 记录失败的项目
                    processed_positions = set()  # 记录已处理的位置，避免重复

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

                        if not old_text or not new_text:
                            # 跳过缺少数据的错误（始终输出信息）
                            skip_count += 1
                            skip_reason = []
                            if not old_text:
                                skip_reason.append("缺少译文数值")
                            if not new_text:
                                skip_reason.append("缺少修改建议")
                            print(f"  [{idx}/{len(all_errors)}] ⊘ 跳过: {', '.join(skip_reason)}")
                            if debug_mode:
                                print(f"      错误类型: {error_type}")
                                print(f"      译文数值: '{old_text}'")
                                print(f"      修改建议: '{new_text}'")
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
                        comment = f"【{error_type}】已修改\n"
                        comment += f"原文: {error.get('原文数值', 'N/A')}\n"
                        comment += f"原译文: {old_text}\n"
                        comment += f"已改为: {new_text}\n"
                        comment += f"理由: {reason}"

                        # 使用改进的替换+批注功能
                        result = replacer.replace_and_annotate(
                            search_text=old_text,
                            new_text=new_text,
                            comment=comment,
                            context=context,
                            color=color,
                            debug=debug_mode
                        )

                        # result 返回 (替换成功数, 批注成功数, 位置标识符)
                        if isinstance(result, tuple) and len(result) == 3:
                            repl_success, annot_success, position_key = result
                        else:
                            repl_success = 0
                            annot_success = 0
                            position_key = None

                        if repl_success > 0:
                            if position_key:
                                processed_positions.add(position_key)

                            success_count += 1
                            replace_count += repl_success
                            annot_count += annot_success

                            status = "✓ 已替换"
                            if annot_success > 0:
                                status += "并批注"
                            print(f"  [{idx}/{len(all_errors)}] {status}: '{old_text}' → '{new_text}'")
                        else:
                            fail_count += 1
                            failed_items.append({
                                'text': old_text,
                                'suggestion': new_text,
                                'reason': reason,
                                'context': context,
                                'error_type': error_type
                            })
                            print(f"  [{idx}/{len(all_errors)}] ✗ 未找到或替换失败: {old_text}")
                            if context and debug_mode:
                                print(f"      上下文: {context[:50]}...")

                    # 保存批注和替换后的文件（覆盖备份文件）
                    replacer.save(str(output_path))

                print(f"\n--- PDF 批注与替换统计 ---")
                print(f"成功处理: {success_count}")
                print(f"  - 替换成功: {replace_count}")
                print(f"  - 批注成功: {annot_count}")
                print(f"失败: {fail_count}")
                print(f"跳过: {skip_count}")
                print(f"总计: {len(all_errors)}")
                if success_count + fail_count > 0:
                    success_rate = success_count / (success_count + fail_count)
                    print(f"成功率: {success_rate:.1%}")

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

                print(f"\n✅ 带批注和替换的 PDF 已保存至: {output_path}")

            except ImportError:
                print("\n❌ 错误: 未安装 PyMuPDF 库")
                print("请运行: pip install PyMuPDF")
            except Exception as e:
                print(f"\n❌ PDF 批注失败: {e}")
                import traceback
                traceback.print_exc()

        return

    # ========== Excel 工作流 ==========
    if is_excel_workflow:
        print("\n⚠️ 检测到 Excel 文件")

        report_paths, _ = run_comparison(args.original, args.translated, bilingual=bilingual)
        if report_paths is None:
            print("❌ 报告生成失败，程序终止")
            return

        print(f"\n✅ 报告已生成")
        for name, path in report_paths.items():
            print(f"  - {name}: {path}")

        print("\n--- 阶段 2: Excel 批注与替换 ---")
        try:
            from replace.excel.excel_replacer import ExcelReplacer
            from backup_copy.backup_manager import ensure_backup_copy

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

            # 生成 Excel 报告
            print("\n--- 生成Excel报告 ---")
            e_report = ExcelReportGenerator()
            body_target_folder = str(LLM_DIR / "zhengwen" / "output_excel")
            body_file_path, _ = e_report.get_excel(body_errors, output_dir=body_target_folder)
            print(f"✓ 正文报告已存至: {body_file_path}")
            header_target_folder = str(LLM_DIR / "yemei" / "output_excel")
            header_file_path, _ = e_report.get_excel(header_errors, output_dir=header_target_folder)
            print(f"✓ 页眉报告已存至: {header_file_path}")
            footer_target_folder = str(LLM_DIR / "yejiao" / "output_excel")
            footer_file_path, _ = e_report.get_excel(footer_errors, output_dir=footer_target_folder)
            print(f"✓ 页脚报告已存至: {footer_file_path}")

            # 备份
            print("\n📦 正在备份原始 Excel 文件...")
            backup_path = ensure_backup_copy(args.translated, suffix="annotated")
            print(f"✓ 备份完成，将对备份文件进行批注")

            # 使用 ExcelReplacer 替换并批注
            replacer = ExcelReplacer(str(backup_path))
            success_count = 0
            fail_count = 0
            skip_count = 0
            failed_items = []

            for idx, error in enumerate(all_errors, 1):
                old_text = (error.get("译文数值") or "").strip()
                new_text = (error.get("译文修改建议值") or "").strip()
                reason = error.get("修改理由", "数值错误")
                error_type = error.get("错误类型", "数值错误")
                context = error.get("译文上下文", "")

                if not old_text or not new_text:
                    skip_count += 1
                    skip_reason = []
                    if not old_text:
                        skip_reason.append("缺少译文数值")
                    if not new_text:
                        skip_reason.append("缺少修改建议")
                    print(f"  [{idx}/{len(all_errors)}] ⊘ 跳过: {', '.join(skip_reason)}")
                    continue

                ok = replacer.replace_and_annotate(
                    old_text=old_text,
                    new_text=new_text,
                    reason=reason,
                    context=context,
                )
                if ok:
                    success_count += 1
                    print(f"  [{idx}/{len(all_errors)}] ✓ 已替换并批注: '{old_text}' → '{new_text}'")
                else:
                    fail_count += 1
                    failed_items.append({'text': old_text, 'suggestion': new_text, 'reason': reason})
                    print(f"  [{idx}/{len(all_errors)}] ✗ 未找到或替换失败: {old_text}")

            replacer.save(str(backup_path))

            print(f"\n--- Excel 批注与替换统计 ---")
            print(f"成功处理: {success_count}")
            print(f"失败: {fail_count}")
            print(f"跳过: {skip_count}")
            print(f"总计: {len(all_errors)}")
            if success_count + fail_count > 0:
                print(f"成功率: {success_count / (success_count + fail_count):.1%}")

            if failed_items:
                print(f"\n⚠️ 以下 {len(failed_items)} 项未能自动标注（可能需要人工检查）:")
                for item in failed_items[:5]:
                    print(f"  - '{item['text']}' → '{item['suggestion']}'")
                if len(failed_items) > 5:
                    print(f"  ... 还有 {len(failed_items) - 5} 项")

                fail_dir = Path(__file__).parent / "llm" / "llm_project" / "fail"
                fail_dir.mkdir(parents=True, exist_ok=True)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                fail_report_path = fail_dir / f"{Path(args.translated).stem}_failed_excel_{timestamp}.txt"
                with open(fail_report_path, 'w', encoding='utf-8') as f:
                    f.write("未能自动标注的项目 (Excel)\n")
                    f.write("=" * 60 + "\n\n")
                    for i, item in enumerate(failed_items, 1):
                        f.write(f"{i}. 查找文本: {item['text']}\n")
                        f.write(f"   建议修改: {item['suggestion']}\n")
                        f.write(f"   修改理由: {item['reason']}\n\n")
                print(f"\n✓ 失败项详情已保存至: {fail_report_path}")

            print(f"\n✅ 带批注和替换的 Excel 已保存至: {backup_path}")

        except Exception as e:
            print(f"\n❌ Excel 批注失败: {e}")
            import traceback
            traceback.print_exc()

        return

    # ========== PPTX 工作流 ==========
    if is_pptx_workflow:
        print("\n⚠️ 检测到 PPTX 文件")

        report_paths, _ = run_comparison(args.original, args.translated, bilingual=bilingual)
        if report_paths is None:
            print("❌ 报告生成失败，程序终止")
            return

        print(f"\n✅ 报告已生成")
        for name, path in report_paths.items():
            print(f"  - {name}: {path}")

        print("\n--- 阶段 2: PPTX 批注与替换 ---")
        try:
            from replace.pptx.pptx_replacer import PPTXReplacer
            from backup_copy.backup_manager import ensure_backup_copy

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

            # 生成 Excel 报告
            print("\n--- 生成Excel报告 ---")
            e_report = ExcelReportGenerator()
            body_target_folder = str(LLM_DIR / "zhengwen" / "output_excel")
            body_file_path, _ = e_report.get_excel(body_errors, output_dir=body_target_folder)
            print(f"✓ 正文报告已存至: {body_file_path}")
            header_target_folder = str(LLM_DIR / "yemei" / "output_excel")
            header_file_path, _ = e_report.get_excel(header_errors, output_dir=header_target_folder)
            print(f"✓ 页眉报告已存至: {header_file_path}")
            footer_target_folder = str(LLM_DIR / "yejiao" / "output_excel")
            footer_file_path, _ = e_report.get_excel(footer_errors, output_dir=footer_target_folder)
            print(f"✓ 页脚报告已存至: {footer_file_path}")

            # 备份
            print("\n📦 正在备份原始 PPTX 文件...")
            backup_path = ensure_backup_copy(args.translated, suffix="annotated")
            print(f"✓ 备份完成，将对备份文件进行批注")

            # 使用 PPTXReplacer 替换并批注
            replacer = PPTXReplacer(str(backup_path))
            success_count = 0
            fail_count = 0
            skip_count = 0
            failed_items = []

            for idx, error in enumerate(all_errors, 1):
                old_text = (error.get("译文数值") or "").strip()
                new_text = (error.get("译文修改建议值") or "").strip()
                reason = error.get("修改理由", "数值错误")
                error_type = error.get("错误类型", "数值错误")
                context = error.get("译文上下文", "")

                if not old_text or not new_text:
                    skip_count += 1
                    skip_reason = []
                    if not old_text:
                        skip_reason.append("缺少译文数值")
                    if not new_text:
                        skip_reason.append("缺少修改建议")
                    print(f"  [{idx}/{len(all_errors)}] ⊘ 跳过: {', '.join(skip_reason)}")
                    continue

                ok = replacer.replace_and_annotate(
                    old_text=old_text,
                    new_text=new_text,
                    reason=reason,
                    context=context,
                )
                if ok:
                    success_count += 1
                    print(f"  [{idx}/{len(all_errors)}] ✓ 已替换并批注: '{old_text}' → '{new_text}'")
                else:
                    fail_count += 1
                    failed_items.append({'text': old_text, 'suggestion': new_text, 'reason': reason})
                    print(f"  [{idx}/{len(all_errors)}] ✗ 未找到或替换失败: {old_text}")

            replacer.save(str(backup_path))

            print(f"\n--- PPTX 批注与替换统计 ---")
            print(f"成功处理: {success_count}")
            print(f"失败: {fail_count}")
            print(f"跳过: {skip_count}")
            print(f"总计: {len(all_errors)}")
            if success_count + fail_count > 0:
                print(f"成功率: {success_count / (success_count + fail_count):.1%}")

            if failed_items:
                print(f"\n⚠️ 以下 {len(failed_items)} 项未能自动标注（可能需要人工检查）:")
                for item in failed_items[:5]:
                    print(f"  - '{item['text']}' → '{item['suggestion']}'")
                if len(failed_items) > 5:
                    print(f"  ... 还有 {len(failed_items) - 5} 项")

                fail_dir = Path(__file__).parent / "llm" / "llm_project" / "fail"
                fail_dir.mkdir(parents=True, exist_ok=True)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                fail_report_path = fail_dir / f"{Path(args.translated).stem}_failed_pptx_{timestamp}.txt"
                with open(fail_report_path, 'w', encoding='utf-8') as f:
                    f.write("未能自动标注的项目 (PPTX)\n")
                    f.write("=" * 60 + "\n\n")
                    for i, item in enumerate(failed_items, 1):
                        f.write(f"{i}. 查找文本: {item['text']}\n")
                        f.write(f"   建议修改: {item['suggestion']}\n")
                        f.write(f"   修改理由: {item['reason']}\n\n")
                print(f"\n✓ 失败项详情已保存至: {fail_report_path}")

            print(f"\n✅ 带批注和替换的 PPTX 已保存至: {backup_path}")

        except Exception as e:
            print(f"\n❌ PPTX 批注失败: {e}")
            import traceback
            traceback.print_exc()

        return

    # 3) 执行对比并获取生成的 JSON 路径
    report_paths, final_rule_text = run_comparison(
        args.original,
        args.translated,
        use_ai_rule=use_ai_rule,
        bilingual=bilingual,
    )
    # 检查返回值
    if report_paths is None or final_rule_text is None:
        print("❌ 规则加载失败，程序终止")
        return

    print(f"\n✅ 规则内容已加载 (长度: {len(final_rule_text)} 字符)")

    # 4) 核心修复逻辑
    print("\n--- 阶段 2: 自动替换与批注 ---")

    # 确保译文是 Word 文档
    if translated_ext not in [".docx", ".doc"]:
        print("❌ 错误: 阶段2仅支持 Word/PDF/Excel/PPTX 文档格式的译文")
        return
    # 创建备份
    from backup_copy.backup_manager import ensure_backup_copy
    backup_copy_path = ensure_backup_copy(args.translated)
    
    # 自动编号静态化：把所有自动编号转成静态文本，避免改一个编号后面全乱
    try:
        from replace.word.numbering_to_static import convert_numbering_to_static, has_auto_numbering
        if has_auto_numbering(str(backup_copy_path)):
            print("\n🔢 检测到自动编号，正在转换为静态文本...")
            ok = convert_numbering_to_static(str(backup_copy_path))
            if ok:
                print("✓ 自动编号已转为静态文本")
            else:
                print("⚠ 自动编号静态化失败，部分编号可能无法替换")
    except Exception as e:
        print(f"⚠ 编号静态化异常: {e}")
    
    doc = Document(backup_copy_path)
    doc._numbering_staticized = True
    revision_manager = RevisionManager(doc)

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
    body_target_folder = str(LLM_DIR / "zhengwen" / "output_excel")
    body_file_path, body_data = e.get_excel(body_errors, output_dir=body_target_folder)
    print(f"✓ 正文报告已存至: {body_file_path}")

    # 页眉excel
    header_target_folder = str(LLM_DIR / "yemei" / "output_excel")
    header_file_path, header_data = e.get_excel(header_errors, output_dir=header_target_folder)
    print(f"✓ 页眉报告已存至: {header_file_path}")

    # 页脚excel
    footer_target_folder = str(LLM_DIR / "yejiao" / "output_excel")
    footer_file_path, footer_data = e.get_excel(footer_errors, output_dir=footer_target_folder)
    print(f"✓ 页脚报告已存至: {footer_file_path}")

    #
    # body_result_path=str(LLM_DIR / "zhengwen" / "output_json" / "文本对比结果_20260306_154849.json")
    # header_result_path=str(LLM_DIR / "yemei" / "output_json" / "文本对比结果_20260306_154913.json")
    # footer_result_path=str(LLM_DIR / "yejiao" / "output_json" / "文本对比结果_20260306_154937.json")
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
            from replace.word.replace_clean import is_list_pattern
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
                context=context, anchor_text=anchor, region=region,
                doc_path=str(backup_copy_path)
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

    # 在系统批注中记录页眉页脚的修改情况
    print("\n" + "=" * 60)
    print("页眉页脚修改汇总")
    print("=" * 60)

    if h_s == 0 and h_f == 0 and h_skip == 0:
        print("  ℹ️  页眉：无修改")
    else:
        # 添加汇总信息
        print(f"\n  ✓ 页眉：共修改 {h_s} 处")

        # 添加详细修改记录
        if header_changes:
            print("  详细修改:")
            for i, change in enumerate(header_changes, 1):
                print(f"    {i}. {change}")

    if f_s == 0 and f_f == 0 and f_skip == 0:
        print("\n  ℹ️  页脚：无修改")
    else:
        # 添加汇总信息
        print(f"\n  ✓ 页脚：共修改 {f_s} 处")

        # 添加详细修改记录
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
        fail_dir = LLM_DIR / "fail"
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
    
    # 执行脚注替换（必须在doc.save()之后，否则会被覆盖）
    footnote_count = flush_footnote_replacements(doc, str(backup_copy_path))
    if footnote_count > 0:
        print(f"\n✓ 脚注替换完成: {footnote_count} 处")
    print(f"\n" + "=" * 40)

    print(f"🎉 全部流程处理完成！")
    print(f"最终结果保存至: {backup_copy_path}")
    print("=" * 40)


if __name__ == '__main__':
    main()
