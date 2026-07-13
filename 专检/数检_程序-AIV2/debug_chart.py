import sys
sys.path.insert(0, r'D:\project\数检_程序-AI')
from zipfile import ZipFile
from lxml import etree
from full_content import scan_docx, SOURCE_LABELS
from collections import Counter

path = r'D:\project\数检_程序-AI\测试文件\雅本化学2025ESG报告文字稿-20260409.docx'

# 1. 扫描结果统计
segs = scan_docx(path)
counts = Counter(s.source for s in segs)
print('=== 片段统计 ===')
for src, cnt in counts.items():
    print(f'  {SOURCE_LABELS.get(src, src)}: {cnt}')

# 2. 找"重要性议题"附近的片段
print('\n=== 含"重要性议题"的片段 ===')
for i, s in enumerate(segs):
    if '重要性议题' in s.text:
        print(f'  [{i+1:>4}] [{SOURCE_LABELS.get(s.source, s.source)}] {s.text[:80]}')

# 3. 查看 docx 内部 XML 文件列表（chart 相关）
print('\n=== docx 内 chart/drawing 相关文件 ===')
with ZipFile(path) as zf:
    names = zf.namelist()
    chart_files = [n for n in names if 'chart' in n.lower() or 'drawing' in n.lower()]
    for n in chart_files[:20]:
        print(f'  {n}')
    print(f'  ... 共 {len(chart_files)} 个')
