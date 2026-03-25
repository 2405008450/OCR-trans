# body_extractor.py
import os
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
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NAMESPACES = {"w": W_NS, "wp": WP_NS, "a": A_NS, "wps": WPS_NS, "v": V_NS, "r": R_NS, "m": M_NS, "mc": MC_NS}


# ---------------- 编号系统 ----------------
class NumberingSystem:
    """处理 Word 自动编号系统（沿用你原逻辑）"""

    def __init__(self, doc_path: str):
        self.doc_path = doc_path
        self.numbering_map = {}        # {numId: {ilvl: format_info}}
        self.abstract_num_map = {}     # {abstractNumId: {ilvl: format_info}}
        self.level_counters = {}       # {(numId, ilvl): current_count}
        self._load_numbering()

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
                            self.numbering_map[num_id] = self.abstract_num_map[abstract_num_id].copy()

        except Exception as e:
            print(f"加载编号系统失败: {e}")

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
        if fmt == "chineseCountingThousand":
            return self._to_chinese(num)
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
            if num_pr is None:
                return None

            num_id_elem = num_pr.find(".//w:numId", NAMESPACES)
            ilvl_elem = num_pr.find(".//w:ilvl", NAMESPACES)
            if num_id_elem is None or ilvl_elem is None:
                return None

            num_id = num_id_elem.get(f"{{{W_NS}}}val")
            ilvl = ilvl_elem.get(f"{{{W_NS}}}val")

            if num_id not in self.numbering_map or ilvl not in self.numbering_map[num_id]:
                return None

            level_info = self.numbering_map[num_id][ilvl]
            counter_key = (num_id, ilvl)

            if counter_key not in self.level_counters:
                self.level_counters[counter_key] = level_info["start"]
            else:
                self.level_counters[counter_key] += 1

            current_num = self.level_counters[counter_key]
            formatted_num = self._format_number(current_num, level_info["format"])
            text_template = level_info["text"]

            return text_template.replace(f"%{int(ilvl) + 1}", formatted_num)

        except Exception as e:
            print(f"解析段落编号失败: {e}")
            return None


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

                    texts = [t.text for t in elem.iter(f"{{{W_NS}}}t") if t.text]
                    full_text = "".join(texts).strip()
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


def _extract_math_formula(math_element) -> str:
    """
    解析 Office MathML 公式为可读文本格式
    支持：分数、上下标、根号、括号、矩阵、积分等
    """
    if math_element is None:
        return ""
    
    tag = math_element.tag
    
    # 分数 <m:f>
    if tag == f"{{{M_NS}}}f":
        num_elem = math_element.find(".//m:num", NAMESPACES)
        den_elem = math_element.find(".//m:den", NAMESPACES)
        num = _extract_math_formula(num_elem) if num_elem is not None else ""
        den = _extract_math_formula(den_elem) if den_elem is not None else ""
        return f"({num}/{den})"
    
    # 上标 <m:sSup>
    elif tag == f"{{{M_NS}}}sSup":
        base_elem = math_element.find(".//m:e", NAMESPACES)
        sup_elem = math_element.find(".//m:sup", NAMESPACES)
        base = _extract_math_formula(base_elem) if base_elem is not None else ""
        sup = _extract_math_formula(sup_elem) if sup_elem is not None else ""
        return f"{base}^{sup}"
    
    # 下标 <m:sSub>
    elif tag == f"{{{M_NS}}}sSub":
        base_elem = math_element.find(".//m:e", NAMESPACES)
        sub_elem = math_element.find(".//m:sub", NAMESPACES)
        base = _extract_math_formula(base_elem) if base_elem is not None else ""
        sub = _extract_math_formula(sub_elem) if sub_elem is not None else ""
        return f"{base}_{sub}"
    
    # 上下标 <m:sSubSup>
    elif tag == f"{{{M_NS}}}sSubSup":
        base_elem = math_element.find(".//m:e", NAMESPACES)
        sub_elem = math_element.find(".//m:sub", NAMESPACES)
        sup_elem = math_element.find(".//m:sup", NAMESPACES)
        base = _extract_math_formula(base_elem) if base_elem is not None else ""
        sub = _extract_math_formula(sub_elem) if sub_elem is not None else ""
        sup = _extract_math_formula(sup_elem) if sup_elem is not None else ""
        return f"{base}_{sub}^{sup}"
    
    # 根号 <m:rad>
    elif tag == f"{{{M_NS}}}rad":
        deg_elem = math_element.find(".//m:deg", NAMESPACES)
        base_elem = math_element.find(".//m:e", NAMESPACES)
        base = _extract_math_formula(base_elem) if base_elem is not None else ""
        if deg_elem is not None:
            deg = _extract_math_formula(deg_elem)
            return f"√[{deg}]({base})"
        return f"√({base})"
    
    # 括号 <m:d>
    elif tag == f"{{{M_NS}}}d":
        # 获取括号字符
        dpr = math_element.find(".//m:dPr", NAMESPACES)
        beg_chr = "("
        end_chr = ")"
        if dpr is not None:
            beg = dpr.find(".//m:begChr", NAMESPACES)
            end = dpr.find(".//m:endChr", NAMESPACES)
            if beg is not None:
                beg_chr = beg.get(f"{{{M_NS}}}val", "(")
            if end is not None:
                end_chr = end.get(f"{{{M_NS}}}val", ")")
        
        content_parts = []
        for e_elem in math_element.findall(".//m:e", NAMESPACES):
            content_parts.append(_extract_math_formula(e_elem))
        content = "".join(content_parts)
        return f"{beg_chr}{content}{end_chr}"
    
    # 矩阵 <m:m>
    elif tag == f"{{{M_NS}}}m":
        rows = []
        for mr in math_element.findall(".//m:mr", NAMESPACES):
            cols = []
            for me in mr.findall(".//m:e", NAMESPACES):
                cols.append(_extract_math_formula(me))
            rows.append(" ".join(cols))
        return "[" + "; ".join(rows) + "]"
    
    # 函数应用 <m:func>
    elif tag == f"{{{M_NS}}}func":
        fname_elem = math_element.find(".//m:fName", NAMESPACES)
        arg_elem = math_element.find(".//m:e", NAMESPACES)
        fname = _extract_math_formula(fname_elem) if fname_elem is not None else ""
        arg = _extract_math_formula(arg_elem) if arg_elem is not None else ""
        return f"{fname}({arg})"
    
    # 积分/求和等 <m:nary>
    elif tag == f"{{{M_NS}}}nary":
        chr_elem = math_element.find(".//m:naryPr/m:chr", NAMESPACES)
        sub_elem = math_element.find(".//m:sub", NAMESPACES)
        sup_elem = math_element.find(".//m:sup", NAMESPACES)
        e_elem = math_element.find(".//m:e", NAMESPACES)
        
        symbol = "∫"
        if chr_elem is not None:
            symbol = chr_elem.get(f"{{{M_NS}}}val", "∫")
        
        sub = _extract_math_formula(sub_elem) if sub_elem is not None else ""
        sup = _extract_math_formula(sup_elem) if sup_elem is not None else ""
        expr = _extract_math_formula(e_elem) if e_elem is not None else ""
        
        if sub and sup:
            return f"{symbol}[{sub}→{sup}]({expr})"
        elif sub:
            return f"{symbol}[{sub}]({expr})"
        elif sup:
            return f"{symbol}^[{sup}]({expr})"
        return f"{symbol}({expr})"
    
    # 递归处理子元素
    parts = []
    for child in math_element:
        child_text = _extract_math_formula(child)
        if child_text:
            parts.append(child_text)
    
    # 提取当前节点的文本
    if math_element.text:
        parts.insert(0, math_element.text)
    if math_element.tail:
        parts.append(math_element.tail)
    
    return "".join(parts)


def _get_xml_text(element) -> str:
    """
    终极增强版：深度递归抓取所有文本和属性值
    解决：乘号 (×)、除号、分式符号等在嵌套标签中丢失的问题
    同时避免 AlternateContent 的 Fallback 导致的重复提取
    支持完整的数学公式解析
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

    def _is_in_fallback(node):
        """检查节点是否在 mc:Fallback 内"""
        parent = node.getparent()
        while parent is not None:
            if parent.tag == f"{{{MC_NS}}}Fallback":
                return True
            parent = parent.getparent()
        return False

    texts = []
    processed_math = set()  # 避免重复处理数学公式
    
    for node in element.iter():
        # 跳过 Fallback 内的所有内容
        if _is_in_fallback(node):
            continue
        
        # 检查是否已经作为数学公式的一部分被处理
        node_id = id(node)
        if node_id in processed_math:
            continue
        
        # 0. 数学公式容器 <m:oMath> - 优先处理
        if node.tag == f"{{{M_NS}}}oMath":
            formula_text = _extract_math_formula(node)
            if formula_text:
                texts.append(f"[公式: {formula_text}]")
            # 标记所有子节点为已处理
            for child in node.iter():
                processed_math.add(id(child))
            continue
            
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


def _clean_duplicate_text(text: str) -> str:
    """
    清理文本中的重复部分（处理 VML/Drawing 双重表示）
    
    只处理完全自重复的情况：如 "编号1编号1" → "编号1"
    不处理部分重复，避免误伤正常文本
    """
    if not text:
        return text
    
    # 只检查完全自重复（文本长度为偶数且前半部分等于后半部分）
    length = len(text)
    if length % 2 == 0:
        mid = length // 2
        if text[:mid] == text[mid:]:
            return text[:mid]
    
    return text


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

def _process_anchored_content(p_element, loader: DocAnchorsLoader) -> List[str]:
    extras: List[str] = []
    seen_textbox_texts = set()  # 用于去重文本框内容

    # 1. 查找脚注引用 <w:footnoteReference w:id="1"/>
    for ref in p_element.findall(".//w:footnoteReference", NAMESPACES):
        fid = ref.get(f"{{{W_NS}}}id")
        if fid and fid in loader.footnotes:
            extras.append(loader.footnotes[fid])
            print(f"[调试] 找到脚注引用 ID={fid}: {loader.footnotes[fid]}...")

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

    # 4. 文本框：wps:txbxContent / v:textbox（跳过文本框内的段落以避免重复）
    for txbx in p_element.iter(f"{{{WPS_NS}}}txbxContent"):
        for txbx_p in txbx.findall(".//w:p", NAMESPACES):
            if not _is_textbox_paragraph(txbx_p):
                text = _get_xml_text(txbx_p)
                if text:
                    # 清理可能的重复文本
                    cleaned_text = _clean_duplicate_text(text)
                    print(f"[调试-WPS文本框] 原文: {text[:50]}... -> 清理后: {cleaned_text[:50]}...")
                    if cleaned_text and cleaned_text not in seen_textbox_texts:
                        extras.append(cleaned_text)
                        seen_textbox_texts.add(cleaned_text)

    for v_txbx in p_element.iter(f"{{{V_NS}}}textbox"):
        for txbx_p in v_txbx.findall(".//w:p", NAMESPACES):
            if not _is_textbox_paragraph(txbx_p):
                text = _get_xml_text(txbx_p)
                if text:
                    # 清理可能的重复文本
                    cleaned_text = _clean_duplicate_text(text)
                    print(f"[调试-VML文本框] 原文: {text[:50]}... -> 清理后: {cleaned_text[:50]}...")
                    if cleaned_text and cleaned_text not in seen_textbox_texts:
                        extras.append(cleaned_text)
                        seen_textbox_texts.add(cleaned_text)

    return extras


def extract_body_text(doc_path: str) -> str:
    """
    只提取正文（线性顺序）：段落 + 表格
    并把脚注/尾注/批注/文本框插入到对应锚点段落后
    不输出页眉页脚
    
    修复：避免文本框双重表示导致的重复问题
    """
    if not os.path.exists(doc_path):
        raise FileNotFoundError(f"文件不存在: {doc_path}")

    loader = DocAnchorsLoader(doc_path)
    numbering_system = NumberingSystem(doc_path)

    doc = Document(doc_path)
    body_element = doc.element.body

    output_lines: List[str] = []
    seen_extras = set()  # 用于去重锚点内容

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

            # 去重锚点内容（脚注/尾注/批注/文本框）
            for extra in extras:
                if extra not in seen_extras:
                    output_lines.append(extra)
                    seen_extras.add(extra)

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
                        # 去重表格单元格中的锚点内容
                        for extra in cell_extras:
                            if extra not in seen_extras:
                                cell_content.append(extra)
                                seen_extras.add(extra)

                    row_texts.append("\t".join(cell_content))

                if any(row_texts):
                    output_lines.append("\t".join(row_texts))

    return "\n".join(output_lines)


if __name__ == "__main__":
    # 示例：python body_extractor.py your.docx
    # import sys
    #
    # if len(sys.argv) < 2:
    #     print("用法: python body_extractor.py <docx_path>")
    #     raise SystemExit(1)
    # body_result_path=r"C:\Users\Administrator\Desktop\zhongfanyi\llm\llm_project\zhengwen\output_json\文本对比结果_20260304_162055.json"
    # body_errors = extract_and_parse(body_result_path)
    # for err in body_errors:
    #     print(err)
    # print(body_errors)

    # 示例文件路径
    original_path = r"C:\Users\Administrator\Desktop\多语种标点检查\测试文件\修订测试-B260327527-比对更新～关于上海国利货币经纪有限公司2025年度经营情况报告和2026年度经营计划的议案-修订 -SY.docx"
    translated_path = r"C:\Users\Administrator\Desktop\项目文件\专检\数值检查1\测试文件\译文-数值检查测试文件1.docx"    # trans_docx_path = translated_path
    # error_docx_path = r"C:\Users\Administrator\Desktop\project\效果\TP251107025，扬杰科技，中译英（字数3w）\中翻译\文本对比结果.docx"
    #
    # # 提取文件中的文本
    original_doc_path = original_path
    original_text = extract_body_text(original_doc_path)
    #
    translated_doc_path = translated_path
    translated_text = extract_body_text(translated_doc_path)
    print(original_text)
    # print("================================")
    print(translated_text)
