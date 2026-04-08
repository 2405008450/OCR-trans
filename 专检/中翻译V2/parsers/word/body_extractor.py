# body_extractor.py
import os
import re
from typing import Dict, Optional, List
from zipfile import ZipFile

from docx import Document
from lxml import etree

# XML命名空间
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
V_NS = "urn:schemas-microsoft-com:vml"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
# 新增数学命名空间
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
# 更新 NAMESPACES 字典
NAMESPACES = {"w": W_NS, "wp": WP_NS, "a": A_NS, "wps": WPS_NS, "v": V_NS, "r": R_NS, "m": M_NS}


# ---------------- 编号系统 ----------------
class NumberingSystem:
    """处理 Word 自动编号系统（含 lvlOverride 和样式编号回退）"""

    def __init__(self, doc_path: str):
        self.doc_path = doc_path
        self.numbering_map = {}  # {numId: {ilvl: format_info}}
        self.abstract_num_map = {}  # {abstractNumId: {ilvl: format_info}}
        self.level_counters = {}  # {(numId, ilvl): current_count}
        self.style_num_map = {}  # {styleId: (numId, ilvl)}  样式关联的编号
        self._load_numbering()
        self._load_style_numbering()

    def _load_numbering(self):
        try:
            with ZipFile(self.doc_path, "r") as zf:
                if "word/numbering.xml" not in zf.namelist():
                    return

                with zf.open("word/numbering.xml") as f:
                    tree = etree.parse(f)

                    for abstract_num in tree.findall(".//w:abstractNum", NAMESPACES):
                        abstract_num_id = abstract_num.get(f"{{{W_NS}}}abstractNumId")
                        self.abstract_num_map[abstract_num_id] = {}

                        for lvl in abstract_num.findall(".//w:lvl", NAMESPACES):
                            ilvl = lvl.get(f"{{{W_NS}}}ilvl")

                            num_fmt = lvl.find(".//w:numFmt", NAMESPACES)
                            lvl_text = lvl.find(".//w:lvlText", NAMESPACES)
                            start = lvl.find(".//w:start", NAMESPACES)

                            fmt_val = num_fmt.get(f"{{{W_NS}}}val") if num_fmt is not None else "decimal"
                            text_val = lvl_text.get(f"{{{W_NS}}}val") if lvl_text is not None else "%1."
                            start_val = int(start.get(f"{{{W_NS}}}val", "1")) if start is not None else 1

                            self.abstract_num_map[abstract_num_id][ilvl] = {
                                "format": fmt_val,
                                "text": text_val,
                                "start": start_val,
                            }

                    for num in tree.findall(".//w:num", NAMESPACES):
                        num_id = num.get(f"{{{W_NS}}}numId")
                        abstract_num_id_elem = num.find(".//w:abstractNumId", NAMESPACES)
                        if abstract_num_id_elem is None:
                            continue
                        abstract_num_id = abstract_num_id_elem.get(f"{{{W_NS}}}val")
                        if abstract_num_id in self.abstract_num_map:
                            # 先用 abstractNum 的定义做基础
                            level_map = {}
                            for k, v in self.abstract_num_map[abstract_num_id].items():
                                level_map[k] = v.copy()

                            # 再用 lvlOverride 覆盖（用户在 Word 里修改编号格式后会产生这个）
                            for override in num.findall(f"{{{W_NS}}}lvlOverride"):
                                ilvl = override.get(f"{{{W_NS}}}ilvl")
                                if ilvl is None:
                                    continue

                                # lvlOverride 里可能有完整的 lvl 定义
                                lvl = override.find(f"{{{W_NS}}}lvl")
                                if lvl is not None:
                                    num_fmt = lvl.find(".//w:numFmt", NAMESPACES)
                                    lvl_text = lvl.find(".//w:lvlText", NAMESPACES)
                                    start = lvl.find(".//w:start", NAMESPACES)

                                    override_info = level_map.get(ilvl, {}).copy()
                                    if num_fmt is not None:
                                        override_info["format"] = num_fmt.get(f"{{{W_NS}}}val")
                                    if lvl_text is not None:
                                        override_info["text"] = lvl_text.get(f"{{{W_NS}}}val")
                                    if start is not None:
                                        override_info["start"] = int(start.get(f"{{{W_NS}}}val", "1"))
                                    level_map[ilvl] = override_info

                                # lvlOverride 也可能只有 startOverride（只改起始值）
                                start_override = override.find(f"{{{W_NS}}}startOverride")
                                if start_override is not None and ilvl in level_map:
                                    level_map[ilvl]["start"] = int(
                                        start_override.get(f"{{{W_NS}}}val", "1")
                                    )

                            self.numbering_map[num_id] = level_map

        except Exception as e:
            print(f"加载编号系统失败: {e}")

    def _load_style_numbering(self):
        """从 word/styles.xml 加载样式中关联的编号定义"""
        try:
            with ZipFile(self.doc_path, "r") as zf:
                if "word/styles.xml" not in zf.namelist():
                    return

                with zf.open("word/styles.xml") as f:
                    tree = etree.parse(f)

                    for style in tree.findall(".//w:style", NAMESPACES):
                        style_id = style.get(f"{{{W_NS}}}styleId")
                        if not style_id:
                            continue

                        # 查找样式中的 pPr/numPr
                        ppr = style.find(".//w:pPr", NAMESPACES)
                        if ppr is None:
                            continue

                        num_pr = ppr.find(".//w:numPr", NAMESPACES)
                        if num_pr is None:
                            continue

                        num_id_elem = num_pr.find(".//w:numId", NAMESPACES)
                        ilvl_elem = num_pr.find(".//w:ilvl", NAMESPACES)

                        num_id = num_id_elem.get(f"{{{W_NS}}}val") if num_id_elem is not None else None
                        ilvl = ilvl_elem.get(f"{{{W_NS}}}val") if ilvl_elem is not None else "0"

                        if num_id:
                            self.style_num_map[style_id] = (num_id, ilvl)

        except Exception as e:
            print(f"加载样式编号失败: {e}")

    def _format_number(self, num: int, fmt: str) -> str:
        if fmt == "decimal":
            return str(num)
        if fmt == "upperRoman":
            return self._to_roman(num).upper()
        if fmt == "lowerRoman":
            return self._to_roman(num).lower()
        if fmt == "upperLetter":
            return self._to_letter(num).upper()
        if fmt == "lowerLetter":
            return self._to_letter(num).lower()
        if fmt in ("chineseCountingThousand", "chineseCounting", "ideographTraditional"):
            return self._to_chinese(num)
        if fmt in ("japaneseCounting", "japaneseDigitalTenThousand"):
            return self._to_chinese(num)  # 日式计数在此场景下等同中文数字
        if fmt == "bullet":
            return "•"
        return str(num)

    @staticmethod
    def _to_roman(num: int) -> str:
        val_map = [
            (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
            (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
            (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
        ]
        result = ""
        for value, letter in val_map:
            while num >= value:
                result += letter
                num -= value
        return result

    @staticmethod
    def _to_letter(num: int) -> str:
        result = ""
        while num > 0:
            num -= 1
            result = chr(65 + num % 26) + result
            num //= 26
        return result

    @staticmethod
    def _to_chinese(num: int) -> str:
        chinese_nums = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"]
        units = ["", "十", "百", "千", "万"]
        if num == 0:
            return chinese_nums[0]
        result = ""
        unit_idx = 0
        while num > 0:
            digit = num % 10
            if digit != 0:
                result = chinese_nums[digit] + units[unit_idx] + result
            elif result and result[0] != "零":
                result = chinese_nums[0] + result
            num //= 10
            unit_idx += 1
        if result.startswith("一十"):
            result = result[1:]
        return result.rstrip("零")

    def get_paragraph_number(self, p_element) -> Optional[str]:
        try:
            num_pr = p_element.find(".//w:numPr", NAMESPACES)

            # 如果段落本身没有 numPr，尝试从样式定义中查找
            if num_pr is None:
                ppr = p_element.find(".//w:pPr", NAMESPACES)
                if ppr is not None:
                    pstyle = ppr.find(".//w:pStyle", NAMESPACES)
                    if pstyle is not None:
                        style_id = pstyle.get(f"{{{W_NS}}}val")
                        if style_id and style_id in self.style_num_map:
                            num_id, ilvl = self.style_num_map[style_id]
                            # 构造虚拟的 numId/ilvl 继续走后面的逻辑
                            return self._resolve_number(num_id, ilvl)
                return None

            num_id_elem = num_pr.find(".//w:numId", NAMESPACES)
            ilvl_elem = num_pr.find(".//w:ilvl", NAMESPACES)
            if num_id_elem is None or ilvl_elem is None:
                return None

            num_id = num_id_elem.get(f"{{{W_NS}}}val")
            ilvl = ilvl_elem.get(f"{{{W_NS}}}val")

            # numId=0 表示"取消编号"，不应输出编号
            if num_id == "0":
                return None

            return self._resolve_number(num_id, ilvl)

        except Exception as e:
            print(f"解析段落编号失败: {e}")
            return None

    def _resolve_number(self, num_id: str, ilvl: str) -> Optional[str]:
        """根据 numId 和 ilvl 计算编号文本，支持多级占位符和子级别重置"""
        if num_id not in self.numbering_map or ilvl not in self.numbering_map[num_id]:
            return None

        ilvl_int = int(ilvl)
        level_info = self.numbering_map[num_id][ilvl]
        counter_key = (num_id, ilvl)

        if counter_key not in self.level_counters:
            self.level_counters[counter_key] = level_info["start"]
        else:
            self.level_counters[counter_key] += 1

        # 当高级别递增时，重置所有更低级别的计数器
        for other_ilvl_str, other_info in self.numbering_map[num_id].items():
            other_ilvl_int = int(other_ilvl_str)
            if other_ilvl_int > ilvl_int:
                other_key = (num_id, other_ilvl_str)
                # 删除计数器，下次遇到时会从 start 重新开始
                if other_key in self.level_counters:
                    del self.level_counters[other_key]

        # 多级占位符替换：lvlText 可能包含 %1, %2, %3 等多个级别的引用
        text_template = level_info["text"]
        for lvl_idx in range(ilvl_int + 1):
            placeholder = f"%{lvl_idx + 1}"
            if placeholder not in text_template:
                continue
            lvl_str = str(lvl_idx)
            lvl_key = (num_id, lvl_str)
            if lvl_str in self.numbering_map[num_id] and lvl_key in self.level_counters:
                lvl_info = self.numbering_map[num_id][lvl_str]
                lvl_num = self.level_counters[lvl_key]
                formatted = self._format_number(lvl_num, lvl_info["format"])
                text_template = text_template.replace(placeholder, formatted)

        return text_template


# ---------------- 辅助内容加载：脚注/尾注/批注 ----------------
class DocAnchorsLoader:
    """只预加载正文需要用到的锚点内容：脚注/尾注/批注（剥离页眉/页脚）"""

    def __init__(self, doc_path: str):
        self.doc_path = doc_path
        self.footnotes: Dict[str, str] = {}
        self.endnotes: Dict[str, str] = {}
        self.comments: Dict[str, str] = {}
        self._load_all()

    def _load_xml_map(self, zip_file: ZipFile, filename: str, tag_name: str, id_attr: str = "id") -> Dict[str, str]:
        data_map: Dict[str, str] = {}
        if filename not in zip_file.namelist():
            return data_map

        try:
            with zip_file.open(filename) as f:
                tree = etree.parse(f)
                for elem in tree.findall(f".//w:{tag_name}", NAMESPACES):
                    eid = elem.get(f"{{{W_NS}}}{id_attr}")

                    elem_type = elem.get(f"{{{W_NS}}}type")
                    if elem_type in ("separator", "continuationSeparator"):
                        continue

                    # 使用 _get_xml_text 函数来正确提取所有文本，包括符号
                    full_text = _get_xml_text(elem)
                    if full_text and eid:
                        data_map[eid] = full_text
        except Exception as e:
            print(f"加载 {filename} 失败: {e}")

        return data_map

    def _load_all(self):
        with ZipFile(self.doc_path, "r") as zf:
            self.footnotes = self._load_xml_map(zf, "word/footnotes.xml", "footnote")
            self.endnotes = self._load_xml_map(zf, "word/endnotes.xml", "endnote")
            self.comments = self._load_xml_map(zf, "word/comments.xml", "comment")


def _get_xml_text(element) -> str:
    """
    终极增强版：深度递归抓取所有文本和属性值
    解决：乘号 (×)、除号、分式符号等在嵌套标签中丢失的问题
    """
    # Symbol字体字符映射表（常用数学符号）
    SYMBOL_CHAR_MAP = {
        'F0B4': '×',  # 乘号
        'F0B8': '÷',  # 除号
        'F0B1': '±',  # 加减号
        'F0B3': '≥',  # 大于等于
        'F0A3': '≤',  # 小于等于
        'F0B9': '≠',  # 不等于
        'F0BB': '≈',  # 约等于
    }

    texts = []
    for node in element.iter():
        # 1. 标准文本和数学文本
        if node.tag in (f"{{{W_NS}}}t", f"{{{M_NS}}}t"):
            if node.text:
                texts.append(node.text)
            # 重要：也要检查 tail 文本（标签后的文本）
            if node.tail:
                texts.append(node.tail)

        # 2. Symbol字体符号 (w:sym) - 关键修复点！
        elif node.tag == f"{{{W_NS}}}sym":
            char_code = node.get(f"{{{W_NS}}}char")
            if char_code:
                # 转换为大写以匹配映射表
                char_code_upper = char_code.upper()
                # 如果在映射表中，使用映射的字符；否则尝试直接转换
                if char_code_upper in SYMBOL_CHAR_MAP:
                    texts.append(SYMBOL_CHAR_MAP[char_code_upper])
                else:
                    # 尝试将十六进制转为Unicode字符
                    try:
                        texts.append(chr(int(char_code, 16)))
                    except (ValueError, OverflowError):
                        # 如果转换失败，保留原始代码作为占位符
                        texts.append(f"[{char_code}]")

        # 3. 数学符号 (关键点：很多符号存放在 m:char 的 w:val 属性中)
        elif node.tag == f"{{{M_NS}}}char":
            # 尝试抓取所有可能的 val 属性，这是乘号最常藏身的地方
            val = node.get(f"{{{M_NS}}}val") or node.get(f"{{{W_NS}}}val")
            if val:
                texts.append(val)

        # 4. 容错处理：某些特殊的公式操作符 (如分数的符号)
        # 如果节点没有 text，但它是一个 m: 标签且有 val 属性，也抓出来
        elif node.tag.startswith(f"{{{M_NS}}}"):
            val = node.get(f"{{{W_NS}}}val") or node.get(f"{{{M_NS}}}val")
            if val and len(val) == 1:  # 通常符号长度为 1
                texts.append(val)

        # 5. 图片替代文本
        elif node.tag == f"{{{WP_NS}}}docPr":
            alt = node.get("descr") or node.get("title")
            if alt:
                texts.append(f"[图片描述: {alt}]")

    return "".join(texts).strip()


def _process_anchored_content(p_element, loader: DocAnchorsLoader) -> List[str]:
    extras: List[str] = []

    # 1. 查找脚注引用 <w:footnoteReference w:id="1"/>
    for ref in p_element.findall(".//w:footnoteReference", NAMESPACES):
        fid = ref.get(f"{{{W_NS}}}id")
        if fid and fid in loader.footnotes:
            extras.append(loader.footnotes[fid])

    # 2. 查找尾注引用 <w:endnoteReference w:id="1"/>
    for ref in p_element.findall(".//w:endnoteReference", NAMESPACES):
        eid = ref.get(f"{{{W_NS}}}id")
        if eid and eid in loader.endnotes:
            extras.append(loader.endnotes[eid])

    # 3. 查找批注引用 <w:commentReference w:id="1"/>
    for ref in p_element.findall(".//w:commentReference", NAMESPACES):
        cid = ref.get(f"{{{W_NS}}}id")
        if cid and cid in loader.comments:
            extras.append(loader.comments[cid])

    # 4. 文本框：wps:txbxContent / v:textbox
    for txbx in p_element.iter(f"{{{WPS_NS}}}txbxContent"):
        text = _get_xml_text(txbx)
        if text:
            extras.append(text)

    for v_txbx in p_element.iter(f"{{{V_NS}}}textbox"):
        text = _get_xml_text(v_txbx)
        if text:
            extras.append(text)

    return extras


def extract_body_text(doc_path: str) -> str:
    """
    只提取正文（线性顺序）：段落 + 表格
    并把脚注/尾注/批注/文本框插入到对应锚点段落后
    不输出页眉页脚
    """
    if not os.path.exists(doc_path):
        raise FileNotFoundError(f"文件不存在: {doc_path}")

    loader = DocAnchorsLoader(doc_path)
    numbering_system = NumberingSystem(doc_path)

    doc = Document(doc_path)
    body_element = doc.element.body

    output_lines: List[str] = []

    for child in body_element.iterchildren():
        tag_name = child.tag

        # 段落
        if tag_name.endswith("p"):
            number_text = numbering_system.get_paragraph_number(child)
            text = _get_xml_text(child)
            extras = _process_anchored_content(child, loader)

            if number_text and text:
                full_text = f"{number_text} {text}"
            elif number_text:
                full_text = number_text
            elif text:
                full_text = text
            else:
                full_text = ""

            if full_text.strip():
                output_lines.append(full_text)

            for extra in extras:
                output_lines.append(extra)

        # 表格
        elif tag_name.endswith("tbl"):
            for row in child.iter(f"{{{W_NS}}}tr"):
                row_texts: List[str] = []
                for cell in row.iter(f"{{{W_NS}}}tc"):
                    cell_content: List[str] = []
                    for cell_p in cell.iter(f"{{{W_NS}}}p"):
                        cell_number = numbering_system.get_paragraph_number(cell_p)
                        p_text = _get_xml_text(cell_p)

                        if cell_number and p_text:
                            full_cell_p_text = f"{cell_number} {p_text}"
                        elif cell_number:
                            full_cell_p_text = cell_number
                        elif p_text:
                            full_cell_p_text = p_text
                        else:
                            full_cell_p_text = ""

                        if full_cell_p_text.strip():
                            cell_content.append(full_cell_p_text)

                        cell_extras = _process_anchored_content(cell_p, loader)
                        cell_content.extend(cell_extras)

                    row_texts.append("\t".join(cell_content))

                if any(row_texts):
                    output_lines.append("\t".join(row_texts))

    return "\n".join(output_lines)


if __name__ == "__main__":
    import io
    import sys

    # 设置标准输出为UTF-8编码
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

    # 示例文件路径
    original_path = r"C:\Users\H\Desktop\测试项目\中翻译\测试文件\中英对照-含不可编辑_20260401 【翻译同步】迦南智能2025年度可持续发展报告初稿V2.9_Bilingual_corrected.docx"
    translated_path = r"C:\Users\H\Desktop\测试项目\中翻译\测试文件\原文-B260328127-Y-中国银行开源软件管理指引.docx"  # trans_docx_path = translated_path

    # 提取文件中的文本
    original_doc_path = original_path
    original_text = extract_body_text(original_doc_path)
    print(original_text)

