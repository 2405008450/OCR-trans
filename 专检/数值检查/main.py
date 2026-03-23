import os
import sys
import time
import argparse
from docx import Document

from llm.llm_project.llm_check.check import Match
from llm.llm_project.parsers.body_extractor import extract_body_text
from llm.llm_project.parsers.footer_extractor import extract_footers
from llm.llm_project.parsers.header_extractor import extract_headers
from llm.llm_project.replace.fix_replace_docx import ensure_backup_copy
from llm.llm_project.replace.fix_replace_json import replace_and_comment_in_docx, CommentManager
from llm.utils.clean_json import load_json_file
from llm.utils.json_files import write_json_with_timestamp


def log(msg: str) -> None:
    """带时间戳的打印，flush=True 保证立即输出不缓冲。"""
    print(f"[{time.strftime('%H:%M:%S')}] {msg}", flush=True)


def run_comparison(original_path, translated_path, route: str, model_name: str):
    """
    第一阶段：提取文本并调用 AI/Matcher 进行对比，生成 JSON 报告
    """
    log("=== 阶段 1：文本提取 ===")

    log("正在提取原文正文...")
    original_body = extract_body_text(original_path)
    log(f"  原文正文长度: {len(original_body)} 字符")

    log("正在提取译文正文...")
    translated_body = extract_body_text(translated_path)
    log(f"  译文正文长度: {len(translated_body)} 字符")

    log("正在提取页眉...")
    original_header = extract_headers(original_path)
    translated_header = extract_headers(translated_path)
    log(f"  原文页眉长度: {len(original_header)} 字符 / 译文页眉长度: {len(translated_header)} 字符")

    log("正在提取页脚...")
    original_footer = extract_footers(original_path)
    translated_footer = extract_footers(translated_path)
    log(f"  原文页脚长度: {len(original_footer)} 字符 / 译文页脚长度: {len(translated_footer)} 字符")

    print()
    print("==== 原文内容预览 ====")
    print(f"[页眉] {original_header[:200] or '(空)'}")
    print(f"[正文] {original_body[:300] or '(空)'}{'...' if len(original_body) > 300 else ''}")
    print(f"[页脚] {original_footer[:200] or '(空)'}")
    print()
    print("==== 译文内容预览 ====")
    print(f"[页眉] {translated_header[:200] or '(空)'}")
    print(f"[正文] {translated_body[:300] or '(空)'}{'...' if len(translated_body) > 300 else ''}")
    print(f"[页脚] {translated_footer[:200] or '(空)'}")
    print()

    log(f"=== 阶段 2：LLM 对比  路线={route}  模型={model_name} ===")

    def task_logger(msg: str) -> None:
        print(f"  [llm] {msg}", flush=True)

    matcher = Match(model_name=model_name, task_logger=task_logger)

    parts = [
        ("正文", original_body, translated_body,
         r"E:\fastapi-llm-demo\专检\数值检查\llm\llm_project\zhengwen\output_json"),
        ("页眉", original_header, translated_header,
         r"E:\fastapi-llm-demo\专检\数值检查\llm\llm_project\yemei\output_json"),
        ("页脚", original_footer, translated_footer,
         r"E:\fastapi-llm-demo\专检\数值检查\llm\llm_project\yejiao\output_json"),
    ]

    report_paths = {}

    for name, orig_txt, tran_txt, out_dir in parts:
        print()
        log(f"====== 正在检查{name} (原文 {len(orig_txt)} 字符 / 译文 {len(tran_txt)} 字符) ======")
        if orig_txt and tran_txt:
            t0 = time.time()
            res = matcher.compare_texts(orig_txt, tran_txt)
            elapsed = time.time() - t0
            error_count = len(res) if isinstance(res, list) else "?"
            log(f"  {name} 检查完成，耗时 {elapsed:.1f}s，发现 {error_count} 条问题")
        else:
            res = []
            log(f"  ⚠️  {name} 原文或译文为空，跳过")

        _, path = write_json_with_timestamp(res, out_dir)
        report_paths[name] = path
        log(f"  报告已写入: {path}")

    return report_paths


def main():
    # 1) 配置默认路径
    DEFAULT_ORIGINAL = r"D:\记忆异常\测试\文本框测试原文.docx"
    DEFAULT_TRANSLATED = r"D:\记忆异常\测试\文本框测试译文.docx"

    # 2) 命令行参数
    parser = argparse.ArgumentParser(description="Word 自动对比、检测与修复工具")
    parser.add_argument("--original", "-o", default=DEFAULT_ORIGINAL)
    parser.add_argument("--translated", "-t", default=DEFAULT_TRANSLATED)
    parser.add_argument(
        "--route", "-r",
        default=os.getenv("GEMINI_ROUTE", "openrouter"),
        choices=["google", "openrouter"],
        help="LLM 路线: google=直连 Google 官方 API, openrouter=OpenRouter 中转（默认）",
    )
    parser.add_argument(
        "--model", "-m",
        default="gemini-3-flash-preview",
        help="模型名称，默认 gemini-3-flash-preview",
    )
    args = parser.parse_args()

    if not os.path.exists(args.original) or not os.path.exists(args.translated):
        log("❌ 错误: 输入的 docx 文件路径不存在")
        return

    # 将路线写入环境变量，供 check.py 内部 os.getenv 读取
    os.environ["GEMINI_ROUTE"] = args.route
    log(f"启动配置  路线={args.route}  模型={args.model}")
    log(f"原文: {args.original}")
    log(f"译文: {args.translated}")

    t_total = time.time()

    # 3) 执行对比并获取生成的 JSON 路径
    report_paths = run_comparison(args.original, args.translated, args.route, args.model)

    # 4) 核心修复逻辑
    print()
    log("=== 阶段 3：自动替换与批注 ===")

    log("正在创建译文备份...")
    backup_copy_path = ensure_backup_copy(args.translated)
    log(f"备份路径: {backup_copy_path}")

    doc = Document(backup_copy_path)
    comment_manager = CommentManager(doc)
    comment_manager.create_initial_comment()

    def load_errors(label, path):
        if path and os.path.exists(path):
            data = load_json_file(path)
            log(f"已加载 {label} 报告: {len(data)} 条错误")
            if data:
                for i, item in enumerate(data[:3], 1):
                    log(f"  [{i}] 错误类型={item.get('错误类型','')}  "
                        f"译文数值={item.get('译文数值','')}  "
                        f"建议={item.get('译文修改建议值','')}")
                if len(data) > 3:
                    log(f"  ... 共 {len(data)} 条（仅预览前 3 条）")
            return data
        log(f"⚠️  {label} 报告文件不存在: {path}")
        return []

    body_errors   = load_errors("正文", report_paths.get("正文"))
    header_errors = load_errors("页眉", report_paths.get("页眉"))
    footer_errors = load_errors("页脚", report_paths.get("页脚"))

    # 统一定义替换执行函数
    def apply_all_fixes(errors, label):
        if not errors:
            log(f"{label} 无需修复")
            return 0, 0, 0
        log(f">>> 正在修复 {label}（共 {len(errors)} 条）...")
        s_count = f_count = skip_count = 0
        for idx, e in enumerate(errors, 1):
            old    = (e.get("译文数值") or "").strip()
            new    = (e.get("译文修改建议值") or "").strip()
            reason = str(e.get("修改理由") or "数值错误").strip()
            context = e.get("译文上下文", "")
            anchor  = e.get("替换锚点", "")

            if not old or not new:
                log(f"  [{idx}/{len(errors)}] 跳过: 缺少【译文数值】或【译文修改建议值】")
                skip_count += 1
                continue

            ok, strategy = replace_and_comment_in_docx(
                doc, old, new, reason, comment_manager,
                context=context, anchor_text=anchor,
            )
            if ok:
                s_count += 1
                log(f"  [{idx}/{len(errors)}] ✓ '{old}' → '{new}'  策略={strategy}  理由={reason}")
            else:
                f_count += 1
                log(f"  [{idx}/{len(errors)}] ✗ 未匹配到 '{old}'")

        total = s_count + f_count + skip_count
        rate  = s_count / (s_count + f_count) if (s_count + f_count) else 0
        log(f"--- {label} 修复统计: 成功={s_count} 失败={f_count} 跳过={skip_count} 成功率={rate:.0%} ---")
        return s_count, f_count, skip_count

    b_s, b_f, b_skip = apply_all_fixes(body_errors,   "正文")
    h_s, h_f, h_skip = apply_all_fixes(header_errors, "页眉")
    f_s, f_f, f_skip = apply_all_fixes(footer_errors, "页脚")

    total_s     = b_s + h_s + f_s
    total_f     = b_f + h_f + f_f
    total_skip  = b_skip + h_skip + f_skip
    total_count = total_s + total_f + total_skip
    success_rate = total_s / (total_s + total_f) if (total_s + total_f) else 0

    log("正在保存文档...")
    doc.save(backup_copy_path)

    elapsed = time.time() - t_total
    print()
    print("=" * 50)
    log(f"全部流程处理完成！总耗时 {elapsed:.1f}s")
    log(f"成功={total_s}  失败={total_f}  跳过={total_skip}  总计={total_count}  成功率={success_rate:.0%}")
    log(f"最终结果保存至: {backup_copy_path}")
    print("=" * 50)


if __name__ == '__main__':
    main()