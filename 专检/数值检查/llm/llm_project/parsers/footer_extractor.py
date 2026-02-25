# footer_extractor.py
import re
from zipfile import ZipFile
from lxml import etree
from typing import List

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NAMESPACES = {"w": W_NS, "r": R_NS}


def _get_xml_text(element) -> str:
    texts = []
    for t in element.iter(f"{{{W_NS}}}t"):
        if t.text:
            texts.append(t.text)
    return "".join(texts).strip()


def extract_footers(doc_path: str) -> List[str]:
    """
    提取 docx 的所有页脚段落（按 footer1.xml, footer2.xml... 顺序）
    输出：List[str]，每个元素是一段纯文本（不加任何额外标识）
    """
    results: List[str] = []
    with ZipFile(doc_path, "r") as zf:
        pattern = r"word/footer\d*\.xml"
        matching = [f for f in zf.namelist() if re.match(pattern, f)]
        for filename in sorted(matching):
            with zf.open(filename) as f:
                tree = etree.parse(f)
                for p in tree.findall(".//w:p", NAMESPACES):
                    text = _get_xml_text(p)
                    if text:
                        results.append(text)
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
    original_path = r"C:\Users\Administrator\Desktop\project\效果\TP251107025，扬杰科技，中译英（字数3w）\原文-1. 公司章程.docx"
    translated_path = r"C:\Users\Administrator\Desktop\project\效果\TP251222006，香港资翻译，中译英（字数1.7w）\译文-RX-96 LAT Report Vol 19 - Zongtian Contract (pages 4-30).docx"
    trans_docx_path = translated_path
    error_docx_path = r"C:\Users\Administrator\Desktop\project\效果\TP251107025，扬杰科技，中译英（字数3w）\中翻译\文本对比结果.docx"

    # 提取文件中的文本
    original_doc_path = original_path
    original_text = extract_footers(original_doc_path)
    for footer in original_text:
        print(footer)

    translated_doc_path = translated_path
    translated_text = extract_footers(translated_doc_path)
    for footer in translated_text:
        print(footer)

    # print(original_text)
    # print("================================")
    # print(translated_text)
