"""调试：打印 TOC 域覆盖段落的子元素 tag 结构"""
from zipfile import ZipFile
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
DOC_PATH = r"C:\Users\H\Desktop\测试项目\中翻译\测试文件\译文-含不可编辑_01 (2026-007)2025年年度报告(1)_corrected.docx"

with ZipFile(DOC_PATH) as zf:
    with zf.open("word/document.xml") as f:
        tree = etree.parse(f)

root = tree.getroot()
body = root.find(f"{{{W}}}body")
children = list(body)

W_r = f"{{{W}}}r"
W_hyperlink = f"{{{W}}}hyperlink"
W_fldChar = f"{{{W}}}fldChar"
W_instrText = f"{{{W}}}instrText"
W_t = f"{{{W}}}t"
fldCharType = f"{{{W}}}fldCharType"

for i in range(21, min(30, len(children))):
    p = children[i]
    tag = p.tag.split("}")[-1]
    print(f"\n===== 段落 {i} ({tag}) =====")
    for child in p:
        ctag = child.tag.split("}")[-1]
        # 打印直接子元素
        sub_info = []
        for sub in child:
            stag = sub.tag.split("}")[-1]
            # 对 r 里的元素再展开一层
            sub_sub = [s.tag.split("}")[-1] for s in sub]
            if sub_sub:
                sub_info.append(f"{stag}[{','.join(sub_sub)}]")
            else:
                t = sub.text or ""
                sub_info.append(f"{stag}({repr(t[:30])})")
        print(f"  <{ctag}> {' | '.join(sub_info)}")
