"""
dump_ai_errors_json.py — 提取 Excel 中「AI是否正确」为错误的全部行，输出为 JSON
用法：
  python dump_ai_errors_json.py --input 测试文件/output/result.xlsx
"""

import json
import argparse
import pandas as pd

# AI判断列名候选
AI_ERROR_COL_CANDIDATES = ["AI是否正确"]


def _find_col(df: pd.DataFrame, candidates: list):
    for c in candidates:
        if c in df.columns:
            return c
    return None


def collect_ai_errors_from_excel(excel_path: str) -> str:
    """
    读取检查结果 Excel，筛选「AI是否正确」为错误的行，
    用单一字段 ai_error_report 承接，返回 JSON 字符串。
    """
    df = pd.read_excel(excel_path)
    df.columns = [str(c).strip() for c in df.columns]

    ai_col = _find_col(df, AI_ERROR_COL_CANDIDATES)
    if ai_col is None:
        raise ValueError(f"未找到「AI是否正确」列，当前列名：{list(df.columns)}")

    # 筛选包含"错误"字样的行
    error_mask = df[ai_col].astype(str).str.contains("错误", na=False)
    error_df = df[error_mask].copy()

    ai_error_rows = []
    for row_idx, row in error_df.iterrows():
        record = {"excel_row": int(row_idx) + 2}  # +2 = 表头占1行 + 0-index
        for col in df.columns:
            val = row[col]
            if pd.isna(val):
                record[col] = None
            elif isinstance(val, (int, float, bool)):
                record[col] = val
            else:
                record[col] = str(val)
        ai_error_rows.append(record)

    payload = {
        "total_rows":      len(df),
        "ai_error_count":  len(ai_error_rows),
        "ai_error_report": ai_error_rows,  # 全量AI错误由此字段统一承接
    }
    return json.dumps(payload, ensure_ascii=False, indent=2)


def main():
    parser = argparse.ArgumentParser(description="提取 AI 判断错误行并输出 JSON")
    parser.add_argument("--input", "-i", required=True, help="检查结果 Excel 路径")
    args = parser.parse_args()

    result_json = collect_ai_errors_from_excel(args.input)
    print(result_json)


if __name__ == "__main__":
    excel_path = r"D:\project\数检_程序-AI\测试文件\译文-含不可编辑_01 (2026-007)2025年年度报告(1)_output.xlsx"

    #result_json = collect_errors_from_excel(excel_path)
    #print(result_json)
    output_data=collect_ai_errors_from_excel(excel_path)
    print(output_data)

