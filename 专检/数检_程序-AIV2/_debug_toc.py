"""查看目录段落在提取结果里的 source 和内容"""
from full_content import extract_docx_in_order

path = r"D:\project\数检_程序-AI\测试文件\译文-含不可编辑_01 (2026-007)2025年年度报告(1).docx"
segs = extract_docx_in_order(path)

print("前60条片段（source + 文本）：")
for i, s in enumerate(segs[:60], 1):
    print(f"  [{i:>3}] source={s.source:<10} text={s.text[:60]}")
