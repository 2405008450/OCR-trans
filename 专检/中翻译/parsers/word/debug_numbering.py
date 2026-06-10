# debug_numbering.py
# 诊断脚本：dump 文档的编号系统 XML，找出（一）变(1)的根因
import io
import sys
from zipfile import ZipFile
from lxml import etree

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NAMESPACES = {"w": W_NS}


def dump_numbering(doc_path: str):
    print(f"=== 诊断文档: {doc_path} ===\n")

    with ZipFile(doc_path, "r") as zf:
        # 1. dump numbering.xml 中所有 abstractNum 和 num
        if "word/numbering.xml" in zf.namelist():
            with zf.open("word/numbering.xml") as f:
                tree = etree.parse(f)

            print("【1】abstractNum 定义：")
            for abstract_num in tree.findall(".//w:abstractNum", NAMESPACES):
                abs_id = abstract_num.get(f"{{{W_NS}}}abstractNumId")

                # 检查 styleLink / numStyleLink
                style_link = abstract_num.find(".//w:styleLink", NAMESPACES)
                num_style_link = abstract_num.find(".//w:numStyleLink", NAMESPACES)

                print(f"\n  abstractNumId={abs_id}")
                if style_link is not None:
                    print(f"    ★ styleLink = {style_link.get(f'{{{W_NS}}}val')}")
                if num_style_link is not None:
                    print(f"    ★ numStyleLink = {num_style_link.get(f'{{{W_NS}}}val')}")

                for lvl in abstract_num.findall(".//w:lvl", NAMESPACES):
                    ilvl = lvl.get(f"{{{W_NS}}}ilvl")
                    num_fmt = lvl.find(".//w:numFmt", NAMESPACES)
                    lvl_text = lvl.find(".//w:lvlText", NAMESPACES)
                    start = lvl.find(".//w:start", NAMESPACES)

                    fmt_val = num_fmt.get(f"{{{W_NS}}}val") if num_fmt is not None else "(无)"
                    text_val = lvl_text.get(f"{{{W_NS}}}val") if lvl_text is not None else "(无)"
                    start_val = start.get(f"{{{W_NS}}}val") if start is not None else "(无)"

                    print(f"    ilvl={ilvl}: numFmt={fmt_val}, lvlText='{text_val}', start={start_val}")

            print("\n【2】num 定义（含 lvlOverride）：")
            for num in tree.findall(".//w:num", NAMESPACES):
                num_id = num.get(f"{{{W_NS}}}numId")
                abs_elem = num.find(".//w:abstractNumId", NAMESPACES)
                abs_id = abs_elem.get(f"{{{W_NS}}}val") if abs_elem is not None else "(无)"
                print(f"\n  numId={num_id} -> abstractNumId={abs_id}")

                for override in num.findall(f"{{{W_NS}}}lvlOverride"):
                    ilvl = override.get(f"{{{W_NS}}}ilvl")
                    print(f"    lvlOverride ilvl={ilvl}:")

                    lvl = override.find(f"{{{W_NS}}}lvl")
                    if lvl is not None:
                        num_fmt = lvl.find(".//w:numFmt", NAMESPACES)
                        lvl_text = lvl.find(".//w:lvlText", NAMESPACES)
                        start = lvl.find(".//w:start", NAMESPACES)
                        if num_fmt is not None:
                            print(f"      numFmt={num_fmt.get(f'{{{W_NS}}}val')}")
                        if lvl_text is not None:
                            print(f"      lvlText='{lvl_text.get(f'{{{W_NS}}}val')}'")
                        if start is not None:
                            print(f"      start={start.get(f'{{{W_NS}}}val')}")

                    start_override = override.find(f"{{{W_NS}}}startOverride")
                    if start_override is not None:
                        print(f"      startOverride={start_override.get(f'{{{W_NS}}}val')}")
        else:
            print("【注意】文档中没有 word/numbering.xml")

        # 2. dump styles.xml 中关联编号的样式
        if "word/styles.xml" in zf.namelist():
            with zf.open("word/styles.xml") as f:
                tree = etree.parse(f)

            print("\n\n【3】styles.xml 中关联编号的样式：")
            for style in tree.findall(".//w:style", NAMESPACES):
                style_id = style.get(f"{{{W_NS}}}styleId")
                style_type = style.get(f"{{{W_NS}}}type")
                name_elem = style.find(".//w:name", NAMESPACES)
                style_name = name_elem.get(f"{{{W_NS}}}val") if name_elem is not None else "(无名)"

                ppr = style.find(".//w:pPr", NAMESPACES)
                if ppr is None:
                    continue
                num_pr = ppr.find(".//w:numPr", NAMESPACES)
                if num_pr is None:
                    continue

                num_id_elem = num_pr.find(".//w:numId", NAMESPACES)
                ilvl_elem = num_pr.find(".//w:ilvl", NAMESPACES)
                num_id = num_id_elem.get(f"{{{W_NS}}}val") if num_id_elem is not None else "(无)"
                ilvl = ilvl_elem.get(f"{{{W_NS}}}val") if ilvl_elem is not None else "0"

                print(f"  styleId={style_id} name='{style_name}' type={style_type} -> numId={num_id}, ilvl={ilvl}")

        # 3. dump 所有带编号的段落（不限 pStyle）
        print("\n\n【4】所有带 numPr 的段落（含页眉页脚）：")

        # 收集需要检查的 XML 文件
        xml_files = ["word/document.xml"]
        for name in zf.namelist():
            if (name.startswith("word/header") or name.startswith("word/footer")) and name.endswith(".xml"):
                xml_files.append(name)

        for xml_name in xml_files:
            with zf.open(xml_name) as f:
                tree = etree.parse(f)

            root = tree.getroot()
            count = 0
            for p in root.iter(f"{{{W_NS}}}p"):
                ppr = p.find(f"{{{W_NS}}}pPr")
                if ppr is None:
                    continue

                num_pr = ppr.find(f"{{{W_NS}}}numPr")
                if num_pr is None:
                    continue

                nid_elem = num_pr.find(f"{{{W_NS}}}numId")
                ilvl_elem = num_pr.find(f"{{{W_NS}}}ilvl")
                if nid_elem is None:
                    continue

                nid = nid_elem.get(f"{{{W_NS}}}val")
                ilvl = ilvl_elem.get(f"{{{W_NS}}}val") if ilvl_elem is not None else "0"

                if nid == "0":
                    continue

                pstyle = ppr.find(f"{{{W_NS}}}pStyle")
                style_val = pstyle.get(f"{{{W_NS}}}val") if pstyle is not None else "(无)"

                texts = [t.text for t in p.iter(f"{{{W_NS}}}t") if t.text]
                p_text = "".join(texts)[:80]

                # 检查是否在文本框内
                parent = p.getparent()
                in_textbox = False
                while parent is not None:
                    tag = parent.tag
                    if "txbxContent" in tag or "textbox" in tag:
                        in_textbox = True
                        break
                    parent = parent.getparent()

                location = "文本框内" if in_textbox else "顶层"
                print(f"  [{xml_name}] numId={nid}, ilvl={ilvl}, pStyle={style_val}, 位置={location} | {p_text}")
                count += 1

            if count == 0:
                print(f"  [{xml_name}] 无编号段落")


if __name__ == "__main__":
    doc_path = r"C:\Users\H\Desktop\测试项目\中翻译\测试文件\译文-B260328127-Y-中国银行开源软件管理指引.docx"
    dump_numbering(doc_path)
