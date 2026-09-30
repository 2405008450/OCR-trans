"""
run.py — num_checker 独立运行入口
===================================
用法示例：

  # 从对照 Excel 检查
  python -m num_checker.run --input 测试文件/对照.xlsx --output output/result.xlsx

  # 直接传入句对（调试用）
  python -m num_checker.run --demo
"""

import argparse
import sys
import os

from num_checker.excel_to_json import load_pairs_from_excel

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from num_checker.checker import check_excel, check_pairs, save_report, print_summary
from num_checker.learning_pipeline import queue_check_results, default_store_path


# ─────────────────────────────────────────
# Demo 数据（覆盖文档中提到的典型场景）
# ─────────────────────────────────────────
DEMO_PAIRS = load_pairs_from_excel(
    r"D:\project\数检_程序-AI\测试文件\output\译文-含不可编辑_01 (2026-007)2025年年度报告(1)_alignment.xlsx"
)
print(DEMO_PAIRS[:3])
# DEMO_PAIRS = [
#     # 正确
#     ("本季度营收增长3.3.3倍", "Revenue grew 3.3.3 times this quarter"),
#     # 季度错误（三季度 vs Q2）
#     ("持续关注三季度金融投资", "Continuous attention to Q2 financial investment"),
#     # 数量级错误（80万吨 vs 80,000 tons）
#     ("年产量达80万吨", "Annual output reached 80,000 tons"),
#     # 多义词正确过滤（double-click 不是数值2）
#     ("用户需要双击确认", "Users need to double-click to confirm"),
#     # 多义词数值语境（double the revenue -> 2倍）
#     ("收入翻倍增长至200亿", "Revenue doubled to RMB20 billion"),
#     # 货币等价
#     ("总资产1.5亿元", "Total assets of RMB 150,000,000"),
#     # 百分比错误
#     ("同比增长15%", "Year-on-year growth of 50%"),
#     # 年份范围
#     ("2020-2024年间", "During 2020-2024"),
#     # 年份错误
#     ("2023年第一季度", "Q2 2024"),
#     # 分数
#     ("三分之二的股权", "Two-thirds of the equity"),
#     # basis point 行业符号
#     ("利率上调50个基点", "Interest rate raised by 100 basis points"),
# ]


def main():
    parser = argparse.ArgumentParser(description="纯程序数值检查器（两步走架构）")
    parser.add_argument("--input",  "-i", help="对照 Excel 路径（含原文/译文列）")
    parser.add_argument("--output", "-o", help="输出报告路径（.xlsx）", default="output/num_check_result.xlsx")
    parser.add_argument("--demo",   action="store_true", help="运行内置 Demo 数据")
    parser.add_argument("--no-rl",  action="store_true", help="禁用 RL 多义词过滤")
    parser.add_argument("--all",    action="store_true", help="报告中包含正确行")
    parser.add_argument("--queue-feedback", action="store_true", help="将检查结果写入学习仓库")
    parser.add_argument("--queue-all", action="store_true", help="入队全部案例，默认仅入队错误案例")
    parser.add_argument("--review-store", default=default_store_path(), help="学习仓库 JSON 路径")
    args = parser.parse_args()

    use_rl = not args.no_rl

    if args.demo or not args.input:
        print("运行 Demo 数据...")
        results = check_pairs(DEMO_PAIRS, use_rl_filter=use_rl)
    else:
        print(f"读取对照文件: {args.input}")
        results = check_excel(args.input, use_rl_filter=use_rl)

    print_summary(results)
    save_report(results, args.output, include_correct=args.all)
    if args.queue_feedback:
        queued = queue_check_results(
            results,
            store_path=args.review_store,
            source=args.input or "demo",
            errors_only=not args.queue_all,
        )
        print(f"学习仓库已更新: {args.review_store}")
        print(f"   新增: {queued['added']}  |  更新: {queued['updated']}  |  总案例: {queued['total_cases']}")


if __name__ == "__main__":
    main()
