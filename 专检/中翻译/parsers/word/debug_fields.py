"""
调试：打印文档中所有域代码（fldChar）的 instrText 内容，
帮助确认目录域的实际关键字。
"""
from zipfile import ZipFile
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

DOC_PATH = r"C:\Users\H\Desktop\测试项目\中翻译\测试文件\译文-含不可编辑_01 (2026-007)2025年年度报告(1)_corrected.docx"


def dump_fields(doc_path: str):
    with ZipFile(doc_path, "r") as zf:
        with zf.open("word/document.xml") as f:
            tree = etree.parse(f)

    root = tree.getroot()
    W_r = f"{{{W}}}r"
    W_fldChar = f"{{{W}}}fldChar"
    W_instrText = f"{{{W}}}instrText"
    W_t = f"{{{W}}}t"
    fldCharType = f"{{{W}}}fldCharType"

    field_no = 0
    for p in root.iter(f"{{{W}}}p"):
        children = list(p)
        i = 0
        while i < len(children):
            elem = children[i]
            if elem.tag != W_r:
                i += 1
                continue
            fc = elem.find(W_fldChar)
            if fc is None or fc.get(fldCharType) != "begin":
                i += 1
                continue

            # 收集整个域
            instr_parts = []
            cached_parts = []
            separate_found = False
            j = i + 1
            depth = 1
            while j < len(children):
                r = children[j]
                if r.tag == W_r:
                    fc2 = r.find(W_fldChar)
                    if fc2 is not None:
                        ft = fc2.get(fldCharType)
                        if ft == "begin":
                            depth += 1
                        elif ft == "separate" and depth == 1:
                            separate_found = True
                        elif ft == "end":
                            depth -= 1
                            if depth == 0:
                                break
                    instr = r.find(W_instrText)
                    if instr is not None and instr.text and not separate_found:
                        instr_parts.append(instr.text)
                    t = r.find(W_t)
                    if t is not None and t.text and separate_found:
                        cached_parts.append(t.text)
                j += 1

            field_no += 1
            instr_text = "".join(instr_parts).strip()
            cached_text = "".join(cached_parts).strip()[:80]
            print(f"[域 {field_no}] instrText: {repr(instr_text)}")
            print(f"         缓存值: {repr(cached_text)}")
            print()
            i = j + 1


if __name__ == "__main__":
    dump_fields(DOC_PATH)
