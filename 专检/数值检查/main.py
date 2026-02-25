import os
import sys
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


# 假设 Match 类在这里
# from your_matcher_module import Match

def run_comparison(original_path, translated_path):
    """
    第一阶段：提取文本并调用 AI/Matcher 进行对比，生成 JSON 报告
    """
    print("\n--- 阶段 1: 文本提取与 AI 对比 ---")
    # 1. 提取文本
    orig_doc = Document(original_path)
    tran_doc = Document(translated_path)

    # 这里的 extract_body_text 等函数需要根据你实际的导入情况调用
    original_body = extract_body_text(original_path)
    translated_body = extract_body_text(translated_path)

    original_header = extract_headers(original_path)
    translated_header = extract_headers(translated_path)

    original_footer = extract_footers(original_path)
    translated_footer = extract_footers(translated_path)
    print('==================================原文内容=========================================')
    print('页眉',original_header)
    print('正文',original_body)
    print('页脚', original_footer)
    print('==================================译文内容=========================================')
    print('页眉', translated_header)
    print('正文', translated_body)
    print('页脚', translated_footer)

    matcher = Match()  # 实例化你的对比对象

    results = {}
    parts = [
        ("正文", original_body, translated_body,
         r"C:\Users\Administrator\Desktop\数值检查\llm\llm_project\zhengwen\output_json"),
        ("页眉", original_header, translated_header,
         r"C:\Users\Administrator\Desktop\数值检查\llm\llm_project\yemei\output_json"),
        ("页脚", original_footer, translated_footer,
         r"C:\Users\Administrator\Desktop\数值检查\llm\llm_project\yejiao\output_json")
    ]

    report_paths = {}

    for name, orig_txt, tran_txt, out_dir in parts:
        print(f"====== 正在检查{name} ===========")
        if orig_txt and tran_txt:
            res = matcher.compare_texts(orig_txt, tran_txt)
        else:
            res = []
            print(f"⚠️ {name}原文或译文为空")

        # 写入 JSON
        _, path = write_json_with_timestamp(res, out_dir)
        report_paths[name] = path

    return report_paths


def main():
    # 1) 配置默认路径
    DEFAULT_ORIGINAL = r"C:\Users\Administrator\Desktop\用完就扔\专检项目\TP251222006，香港资翻译，中译英（字数1.7w）\原文\RX-96 LAT Report Vol 19 - Zongtian Contract (pages 4-30).docx"
    DEFAULT_TRANSLATED = r"C:\Users\Administrator\Desktop\用完就扔\专检项目\TP251222006，香港资翻译，中译英（字数1.7w）\译文\RX-96 LAT Report Vol 19 - Zongtian Contract (pages 4-30).docx"

    # 2) 命令行参数
    parser = argparse.ArgumentParser(description="Word 自动对比、检测与修复工具")
    parser.add_argument("--original", "-o", default=DEFAULT_ORIGINAL)
    parser.add_argument("--translated", "-t", default=DEFAULT_TRANSLATED)
    args = parser.parse_args()

    if not os.path.exists(args.original) or not os.path.exists(args.translated):
        print("❌ 错误: 输入的 docx 文件路径不存在")
        return

    # 3) 执行对比并获取生成的 JSON 路径
    # 这一步代替了之前手动指定 JSON 的过程
    report_paths = run_comparison(args.original, args.translated)

    # 4) 核心修复逻辑
    print("\n--- 阶段 2: 自动替换与批注 ---")

    # 创建备份
    backup_copy_path = ensure_backup_copy(args.translated)
    doc = Document(backup_copy_path)
    comment_manager = CommentManager(doc)
    comment_manager.create_initial_comment()

    def load_errors(label, path):
        if path and os.path.exists(path):
            data = load_json_file(path)
            print(f"✓ 已加载{label}报告: {len(data)} 条错误")
            return data
        return []


    # 加载刚刚生成的 JSON
    body_errors = load_errors("正文", report_paths.get("正文"))
    header_errors = load_errors("页眉", report_paths.get("页眉"))
    footer_errors = load_errors("页脚", report_paths.get("页脚"))
    # body_result_path=r"C:\Users\Administrator\Desktop\project\llm\llm_project\zhengwen\output_json\文本对比结果_20260213_105307.json"
    # header_result_path=r"C:\Users\Administrator\Desktop\project\llm\llm_project\yemei\output_json\文本对比结果_20260213_105402.json"
    # footer_result_path=r"C:\Users\Administrator\Desktop\project\llm\llm_project\yejiao\output_json\文本对比结果_20260213_105433.json"
    # # 2) 读取错误报告并解析
    # print("\n正在提取解析正文错误报告...")
    # body_errors = load_json_file(body_result_path)
    # print("正文错误报告", body_errors)
    # for err in body_errors:
    #     print(err)
    # print("正文错误解析个数：", len(body_errors))
    #
    # print("\n正在提取解析页眉错误报告...")
    # header_errors = load_json_file(header_result_path)
    # print("页眉错误报告", header_errors)
    # print("页眉错误解析个数：", len(header_errors))
    #
    # print("\n正在提取解析页脚错误报告...")
    # footer_errors = load_json_file(footer_result_path)
    # print("页脚错误报告", footer_errors)
    # print("页脚错误解析个数：", len(footer_errors))
    print("正文",body_errors)
    print("页眉",header_errors)
    print("页脚",footer_errors)


    # 统一定义替换执行函数 (逻辑保持不变)
    def apply_all_fixes(errors, label):
        if not errors: return 0, 0, 0
        print(f"\n>>> 正在修复 {label} 部分...")
        s_count, f_count = 0, 0
        skip_count = 0
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

            ok, strategy = replace_and_comment_in_docx(
                doc, old, new, reason, comment_manager,
                context=context, anchor_text=anchor
            )
            if ok:
                s_count += 1
                print(f"  [{idx}] 成功: '{old}' -> '{new}'")
                print(f"    修改理由: {reason}")
                print(f"    策略: {strategy}")
                print(f"    操作: '{old}' → '{new}' (已替换并添加批注)")
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

        return s_count, f_count, skip_count

    # 执行三部分修复
    b_s, b_f, b_skip = apply_all_fixes(body_errors, "正文")
    h_s, h_f, h_skip = apply_all_fixes(header_errors, "页眉")
    f_s, f_f, f_skip = apply_all_fixes(footer_errors, "页脚")

    print(f"\n--- 修复统计 ---")
    total_s=b_s+h_s+f_s
    total_f=b_f+h_f+f_f
    total_skip=b_skip+h_skip+f_skip
    total_count=total_s + total_f + total_skip
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