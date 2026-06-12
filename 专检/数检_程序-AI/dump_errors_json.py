"""
dump_errors_json.py — 输出全量检查数据为 JSON

两种输入方式：
  1. 传入 excel 路径（已有检查结果 Excel）
  2. 传入 program_check.run() 返回的 List[Dict]

用法（命令行）：
  python dump_errors_json.py --input 测试文件/output/result.xlsx

用法（代码调用）：
  from dump_errors_json import collect_all_from_rows
  from program_check import run

  final_rows = run(alignment_path="xxx.xlsx")
  json_str = collect_all_from_rows(final_rows)
  print(json_str)
"""

import json
import argparse
import pandas as pd
from typing import List, Dict, Optional


def _row_to_record(row_idx: int, row: dict) -> dict:
    """将单行 dict 转为可序列化 record，list 字段转逗号字符串。"""
    record = {"excel_row": row_idx + 2}  # +2 = 表头 + 0-index
    for k, v in row.items():
        if isinstance(v, list):
            record[k] = ", ".join(str(i) for i in v)
        elif v is None or (isinstance(v, float) and v != v):  # None / NaN
            record[k] = None
        else:
            record[k] = v
    return record


def collect_all_from_rows(rows: List[Dict]) -> str:
    """
    接收 program_check.run() 返回的 List[Dict]，输出全量数据 JSON。
    用单一字段 full_report 承接所有行。
    """
    records = [_row_to_record(i, r) for i, r in enumerate(rows)]
    payload = {
        "total_rows":   len(records),
        "error_count":  sum(1 for r in rows if r.get("是否错误") == "❗错误"),
        "ai_error_count": sum(1 for r in rows if r.get("AI是否正确") == "❗错误"),
        "full_report":  records,   # 全量数据统一由此字段承接
    }
    return json.dumps(payload, ensure_ascii=False, indent=2)


def collect_all_from_excel(excel_path: str) -> str:
    """
    读取已有检查结果 Excel，输出全量数据 JSON。
    用单一字段 full_report 承接所有行。
    """
    df = pd.read_excel(excel_path)
    df.columns = [str(c).strip() for c in df.columns]

    records = []
    for row_idx, row in df.iterrows():
        record = {"excel_row": int(row_idx) + 2}
        for col in df.columns:
            val = row[col]
            if isinstance(val, list):
                record[col] = ", ".join(str(i) for i in val)
            elif pd.isna(val):
                record[col] = None
            elif isinstance(val, (int, float, bool)):
                record[col] = val
            else:
                record[col] = str(val)
        records.append(record)

    payload = {
        "total_rows":     len(records),
        "error_count":    sum(1 for r in records if "错误" in str(r.get("是否错误", ""))),
        "ai_error_count": sum(1 for r in records if "错误" in str(r.get("AI是否正确", ""))),
        "full_report":    records,   # 全量数据统一由此字段承接
    }
    return json.dumps(payload, ensure_ascii=False, indent=2)


def collect_ai_errors_with_context(rows: List[Dict]) -> str:
    """
    筛选 AI 判断为错误的行，附带其前一句和后一句的原文/译文，
    用单一字段 ai_error_report 承接，返回 JSON 字符串。
    """
    # 找出所有 AI 错误行的索引
    error_indices = [i for i, r in enumerate(rows) if r.get("AI是否正确") == "❗错误"]

    def _ctx(row: Optional[Dict]) -> Optional[Dict]:
        if row is None:
            return None
        return {"原文": row.get("原文", ""), "译文": row.get("译文", "")}

    ai_error_rows = []
    for i in error_indices:
        row = rows[i]
        record = _row_to_record(i, row)
        record["前一句"] = _ctx(rows[i - 1] if i > 0 else None)
        record["后一句"] = _ctx(rows[i + 1] if i < len(rows) - 1 else None)
        ai_error_rows.append(record)

    payload = {
        "total_rows":      len(rows),
        "ai_error_count":  len(ai_error_rows),
        "ai_error_report": ai_error_rows,   # AI错误行（含上下文）统一由此字段承接
    }
    return json.dumps(payload, ensure_ascii=False, indent=2)


def main():
    parser = argparse.ArgumentParser(description="全量检查数据 JSON 输出")
    parser.add_argument("--input", "-i", help="检查结果 Excel 路径（不传则需代码调用传入 rows）")
    args = parser.parse_args()

    if args.input:
        result_json = collect_all_from_excel(args.input)
    else:
        parser.error("请通过 --input 传入 Excel 路径，或在代码中调用 collect_all_from_rows(rows)")

    print(result_json)


if __name__ == "__main__":
    # ── 方式1：从 Excel 文件读取 ──
    # excel_path = r"D:\project\数检_程序-AI\测试文件\译文-含不可编辑_01 (2026-007)2025年年度报告(1)_output.xlsx"
    # print(collect_all_from_excel(excel_path))

    # ── 方式2：直接接收 run() 返回值 ──
    from program_check import run
    final_rows = run(
        alignment_path=r"D:\project\数检_程序-AI\测试\bilingual_pairs (8).xlsx",
        output_path=r"D:\project\数检_程序-AI\测试文件\output_data1.xlsx",
    )
    print(collect_ai_errors_with_context(final_rows))
