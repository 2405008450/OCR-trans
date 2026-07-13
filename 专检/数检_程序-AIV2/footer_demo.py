# footer_extractor.py
import re
from zipfile import ZipFile
from lxml import etree
from typing import List

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
V_NS = "urn:schemas-microsoft-com:vml"
WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NAMESPACES = {"w": W_NS, "r": R_NS, "v": V_NS, "wps": WPS_NS, "mc": MC_NS}


def _is_in_fallback(element) -> bool:
    """检查元素是否在 mc:Fallback 内"""
    parent = element.getparent()
    while parent is not None:
        if parent.tag == f"{{{MC_NS}}}Fallback":
            return True
        parent = parent.getparent()
    return False


def _get_xml_text(element) -> str:
    """
    提取元素中的文本，跳过：
    - mc:Fallback 内的文本（避免重复）
    - 域代码缓存值（fldChar separate ~ end 之间的 w:t，如页码 PAGE 域）
    """
    W_fldChar = f"{{{W_NS}}}fldChar"
    W_r = f"{{{W_NS}}}r"
    W_t = f"{{{W_NS}}}t"

    # 先收集所有"域缓存值 run"的 id（在 fldChar separate 和 end 之间的 run）
    field_cache_run_ids = set()
    for p in element.iter(f"{{{W_NS}}}p"):
        in_field = False
        in_cache = False
        for child in p:
            if child.tag != W_r:
                continue
            fld = child.find(W_fldChar)
            if fld is not None:
                ftype = fld.get(f"{{{W_NS}}}fldCharType")
                if ftype == "begin":
                    in_field = True
                    in_cache = False
                elif ftype == "separate":
                    in_cache = True
                elif ftype == "end":
                    in_field = False
                    in_cache = False
            elif in_cache:
                field_cache_run_ids.add(id(child))

    texts = []
    for t in element.iter(W_t):
        if _is_in_fallback(t):
            continue
        # 跳过域缓存值（父 run 在 fldChar separate~end 之间）
        parent_r = t.getparent()
        if parent_r is not None and id(parent_r) in field_cache_run_ids:
            continue
        if t.text:
            texts.append(t.text)
    return "".join(texts).strip()


def _is_textbox_paragraph(p_element) -> bool:
    """
    检查段落是否在文本框内（VML 或 Drawing）
    """
    parent = p_element.getparent()
    while parent is not None:
        tag = parent.tag
        # 检查是否在文本框容器内
        if tag.endswith("}txbxContent") or tag.endswith("}textbox"):
            return True
        parent = parent.getparent()
    return False


def extract_footer_structure(doc_path: str) -> dict:
    """
    提取页脚结构用于对比：
    - static_text: 静态文本内容
    - field_instructions: 域指令列表（如 ['PAGE \\* MERGEFORMAT', 'NUMPAGES']）
    返回 dict，可直接与另一个文件的结果对比
    """
    W_instrText = f"{{{W_NS}}}instrText"
    W_fldChar = f"{{{W_NS}}}fldChar"
    W_r = f"{{{W_NS}}}r"
    W_t = f"{{{W_NS}}}t"

    static_texts = []
    field_instructions = []
    seen_texts = set()
    seen_instrs = set()

    with ZipFile(doc_path, "r") as zf:
        pattern = r"word/footer\d*\.xml"
        matching = [f for f in zf.namelist() if re.match(pattern, f)]
        for filename in sorted(matching):
            with zf.open(filename) as f:
                tree = etree.parse(f)

            # 收集域指令（跳过 Fallback 避免重复）
            for instr in tree.iter(W_instrText):
                if _is_in_fallback(instr):
                    continue
                val = (instr.text or "").strip()
                if val and val not in seen_instrs:
                    field_instructions.append(val)
                    seen_instrs.add(val)

            # 收集静态文本（复用现有逻辑，跳过文本框和缓存值）
            for p in tree.findall(".//w:p", NAMESPACES):
                if _is_textbox_paragraph(p):
                    continue
                text = _get_xml_text(p)
                if text and text not in seen_texts:
                    static_texts.append(text)
                    seen_texts.add(text)

    return {
        "static_text": static_texts,
        "field_instructions": field_instructions,
    }


def compare_footer_structure(original_path: str, translated_path: str) -> dict:
    """
    对比两个文档的页脚结构，返回对比结果
    """
    orig = extract_footer_structure(original_path)
    trans = extract_footer_structure(translated_path)

    return {
        "static_text_match": orig["static_text"] == trans["static_text"],
        "field_instructions_match": orig["field_instructions"] == trans["field_instructions"],
        "original": orig,
        "translated": trans,
    }


def extract_page_numbers(doc_path: str) -> List[str]:
    """
    提取页脚中 PAGE 域的缓存值（即页码数字字符串）。
    只读 mc:Choice 内的文本框段落，跳过 mc:Fallback 避免重复。
    返回去重后的页码列表，如 ['1', '2', '3']
    """
    W_fldChar = f"{{{W_NS}}}fldChar"
    W_r = f"{{{W_NS}}}r"
    W_t = f"{{{W_NS}}}t"
    W_instrText = f"{{{W_NS}}}instrText"

    results: List[str] = []
    seen = set()

    with ZipFile(doc_path, "r") as zf:
        pattern = r"word/footer\d*\.xml"
        matching = [f for f in zf.namelist() if re.match(pattern, f)]
        for filename in sorted(matching):
            with zf.open(filename) as f:
                tree = etree.parse(f)

            for p in tree.iter(f"{{{W_NS}}}p"):
                # 只处理文本框内的段落（页码通常在文本框里）
                if not _is_textbox_paragraph(p):
                    continue
                # 跳过 Fallback（避免重复）
                if _is_in_fallback(p):
                    continue

                # 检查段落是否含 PAGE 域
                instr_texts = [
                    (r.find(W_instrText).text or "")
                    for r in p.iter(W_r)
                    if r.find(W_instrText) is not None
                ]
                if not any("PAGE" in t for t in instr_texts):
                    continue

                # 收集 separate~end 之间的缓存值
                in_cache = False
                for r in p:
                    if r.tag != W_r:
                        continue
                    fld = r.find(W_fldChar)
                    if fld is not None:
                        ftype = fld.get(f"{{{W_NS}}}fldCharType")
                        if ftype == "separate":
                            in_cache = True
                        elif ftype == "end":
                            in_cache = False
                    elif in_cache:
                        t_el = r.find(W_t)
                        if t_el is not None and t_el.text:
                            val = t_el.text.strip()
                            if val and val not in seen:
                                results.append(val)
                                seen.add(val)
    return results


def extract_footers(doc_path: str) -> List[str]:
    """
    提取 docx 的所有页脚段落（按 footer1.xml, footer2.xml... 顺序）
    输出：List[str]，每个元素是一段纯文本（不加任何额外标识）

    修复：避免文本框双重表示导致的重复问题
    - 跳过文本框内的段落（VML/Drawing 双重表示）
    - 使用 set 去重（保持顺序）
    """
    results: List[str] = []
    seen = set()

    with ZipFile(doc_path, "r") as zf:
        pattern = r"word/footer\d*\.xml"
        matching = [f for f in zf.namelist() if re.match(pattern, f)]
        # print(f"找到的页脚文件: {matching}")
        for filename in sorted(matching):
            with zf.open(filename) as f:
                tree = etree.parse(f)

                # 查找所有段落
                for p in tree.findall(".//w:p", NAMESPACES):
                    # 跳过文本框内的段落
                    if _is_textbox_paragraph(p):
                        continue

                    text = _get_xml_text(p)
                    if text and text not in seen:
                        results.append(text)
                        seen.add(text)
    return results


if __name__ == "__main__":
    # 示例：python footer_extractor.py your.docx
    # import sys, os
    #
    # if len(sys.argv) < 2:
    #     print("用法: python footer_extractor.py <docx_path>")
    #     raise SystemExit(1)
    #
    # docx_path = sys.argv[1]
    # if not os.path.exists(docx_path):
    #     print(f"文件不存在: {docx_path}")
    #     raise SystemExit(1)
    #
    # footers = extract_footers(docx_path)
    # print("\n".join(footers))
    # 示例文件路径
    original_path = r"C:\Users\H\Desktop\测试项目\中翻译\测试文件\原文-B260328127-Y-中国银行开源软件管理指引.docx"
    translated_path = r"C:\Users\H\Desktop\测试项目\中翻译\测试文件\译文-B260328127-Y-中国银行开源软件管理指引.docx"  # trans_docx_path = translated_path

    # 提取文件中的文本
    original_doc_path = original_path
    result = compare_footer_structure(original_path, translated_path)
    print("静态文本一致:", result["static_text_match"])
    print("域指令一致:  ", result["field_instructions_match"])
    print("原文结构:", result["original"])
    print("译文结构:", result["translated"])
