"""临时调试脚本：打印表格的行列结构，帮助判断实际 XML 布局"""
import sys
import io
from zipfile import ZipFile
from lxml import etree

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NAMESPACES = {"w": W_NS}

doc_path = r"C:\Users\H\Desktop\测试项目\中翻译\测试文件\中英对照-含不可编辑_20260401 【翻译同步】迦南智能2025年度可持续发展报告初稿V2.9_Bilingual_corrected.docx"

with ZipFile(doc_path, "r") as zf:
    with zf.open("word/document.xml") as f:
        tree = etree.parse(f)

body = tree.find(f"{{{W_NS}}}body")

tbl_idx = 0
for elem in body.iterchildren():
    if not elem.tag.endswith("}tbl"):
        continue
    tbl_idx += 1

    # 只打印前10个表格
    if tbl_idx > 10:
        break

    rows = [r for r in elem.iterchildren() if r.tag.endswith("}tr")]
    print(f"\n=== 表格 {tbl_idx}：共 {len(rows)} 行 ===")

    for r_idx, row in enumerate(rows[:5]):  # 每个表格只看前5行
        cells = [c for c in row.iterchildren() if c.tag.endswith("}tc")]
        print(f"  行 {r_idx+1}：{len(cells)} 列")
        for c_idx, cell in enumerate(cells[:8]):  # 每行只看前8列
            # 收集单元格内所有段落文本
            paras = [p for p in cell.iterchildren() if p.tag.endswith("}p")]
            texts = []
            for p in paras:
                t_nodes = p.findall(f".//{{{W_NS}}}t", NAMESPACES)
                t = "".join(n.text or "" for n in t_nodes).strip()
                if t:
                    texts.append(t)
            print(f"    列{c_idx+1}({len(paras)}段落): {' | '.join(texts[:3])}")
