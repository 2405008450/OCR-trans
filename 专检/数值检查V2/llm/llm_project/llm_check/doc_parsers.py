"""
按文档实际阅读顺序提取所有文字内容（包含自动编号、页眉页脚、脚注、文本框）
实现逻辑：线性遍历XML节点，解析自动章节序号、正文、表格混合排序，并把脚注/文本框等内容穿插在对应的锚点段落后。
核心改进：
1. 页眉页脚完整解析（支持编号、脚注、文本框）
2. 文本框全局检测（覆盖所有容器类型）
3. 脚注引用健壮匹配（兼容不同 XML 格式）
"""

import os
from docx import Document
from lxml import etree
from zipfile import ZipFile
import warnings
from typing import Dict, Optional, List, Tuple
import re

warnings.filterwarnings("ignore")

# ================= 配置区域 =================
SOURCE_DIR = r"C:\Users\Administrator\Desktop\project\效果\TP251117023，北京中翻译，中译英（字数2w）"
CHINESE_FILE = r"C:\Users\Administrator\Desktop\project\效果\TP251117023，北京中翻译，中译英（字数2w）\测试译文-清洁版-B251124195-附件1：中国银行股份有限公司模型风险管理政策（2025年修订）-.docx"

# XML命名空间
W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
WPS_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
V_NS = 'urn:schemas-microsoft-com:vml'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

NAMESPACES = {
    'w': W_NS,
    'wp': WP_NS,
    'a': A_NS,
    'wps': WPS_NS,
    'v': V_NS,
    'r': R_NS
}


# ================= 编号系统类 =================

class NumberingSystem:
    """处理 Word 自动编号系统"""

    def __init__(self, doc_path: str):
        self.doc_path = doc_path
        self.numbering_map = {}  # {numId: {ilvl: format_info}}
        self.abstract_num_map = {}  # {abstractNumId: {ilvl: format_info}}
        self.level_counters = {}  # {(numId, ilvl): current_count}
        self._load_numbering()

    def _load_numbering(self):
        """加载 numbering.xml 文件"""
        try:
            with ZipFile(self.doc_path, 'r') as zf:
                if 'word/numbering.xml' not in zf.namelist():
                    return

                with zf.open('word/numbering.xml') as f:
                    tree = etree.parse(f)

                    # 1. 加载抽象编号定义 (abstractNum)
                    for abstract_num in tree.findall('.//w:abstractNum', NAMESPACES):
                        abstract_num_id = abstract_num.get(f'{{{W_NS}}}abstractNumId')
                        self.abstract_num_map[abstract_num_id] = {}

                        for lvl in abstract_num.findall('.//w:lvl', NAMESPACES):
                            ilvl = lvl.get(f'{{{W_NS}}}ilvl')

                            num_fmt = lvl.find('.//w:numFmt', NAMESPACES)
                            lvl_text = lvl.find('.//w:lvlText', NAMESPACES)
                            start = lvl.find('.//w:start', NAMESPACES)

                            fmt_val = num_fmt.get(f'{{{W_NS}}}val') if num_fmt is not None else 'decimal'
                            text_val = lvl_text.get(f'{{{W_NS}}}val') if lvl_text is not None else '%1.'
                            start_val = int(start.get(f'{{{W_NS}}}val', '1')) if start is not None else 1

                            self.abstract_num_map[abstract_num_id][ilvl] = {
                                'format': fmt_val,
                                'text': text_val,
                                'start': start_val
                            }

                    # 2. 加载编号实例 (num)
                    for num in tree.findall('.//w:num', NAMESPACES):
                        num_id = num.get(f'{{{W_NS}}}numId')
                        abstract_num_id_elem = num.find('.//w:abstractNumId', NAMESPACES)

                        if abstract_num_id_elem is not None:
                            abstract_num_id = abstract_num_id_elem.get(f'{{{W_NS}}}val')
                            if abstract_num_id in self.abstract_num_map:
                                self.numbering_map[num_id] = self.abstract_num_map[abstract_num_id].copy()

        except Exception as e:
            print(f"⚠️ 加载编号系统失败: {e}")

    def _format_number(self, num: int, fmt: str) -> str:
        """将数字转换为指定格式"""
        if fmt == 'decimal':
            return str(num)
        elif fmt == 'upperRoman':
            return self._to_roman(num).upper()
        elif fmt == 'lowerRoman':
            return self._to_roman(num).lower()
        elif fmt == 'upperLetter':
            return self._to_letter(num).upper()
        elif fmt == 'lowerLetter':
            return self._to_letter(num).lower()
        elif fmt == 'chineseCountingThousand':
            return self._to_chinese(num)
        elif fmt == 'bullet':
            return '•'
        else:
            return str(num)

    @staticmethod
    def _to_roman(num: int) -> str:
        """转换为罗马数字"""
        val_map = [
            (1000, 'M'), (900, 'CM'), (500, 'D'), (400, 'CD'),
            (100, 'C'), (90, 'XC'), (50, 'L'), (40, 'XL'),
            (10, 'X'), (9, 'IX'), (5, 'V'), (4, 'IV'), (1, 'I')
        ]
        result = ''
        for value, letter in val_map:
            while num >= value:
                result += letter
                num -= value
        return result

    @staticmethod
    def _to_letter(num: int) -> str:
        """转换为字母 (A, B, C... Z, AA, AB...)"""
        result = ''
        while num > 0:
            num -= 1
            result = chr(65 + num % 26) + result
            num //= 26
        return result

    @staticmethod
    def _to_chinese(num: int) -> str:
        """转换为中文数字"""
        chinese_nums = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九']
        units = ['', '十', '百', '千', '万']

        if num == 0:
            return chinese_nums[0]

        result = ''
        unit_idx = 0

        while num > 0:
            digit = num % 10
            if digit != 0:
                result = chinese_nums[digit] + units[unit_idx] + result
            elif result and result[0] != '零':
                result = chinese_nums[0] + result
            num //= 10
            unit_idx += 1

        if result.startswith('一十'):
            result = result[1:]

        return result.rstrip('零')

    def get_paragraph_number(self, p_element) -> Optional[str]:
        """
        从段落元素中提取编号文本
        返回格式化后的编号字符串，如 "1.", "a)", "(1)" 等
        """
        try:
            num_pr = p_element.find('.//w:numPr', NAMESPACES)
            if num_pr is None:
                return None

            num_id_elem = num_pr.find('.//w:numId', NAMESPACES)
            ilvl_elem = num_pr.find('.//w:ilvl', NAMESPACES)

            if num_id_elem is None or ilvl_elem is None:
                return None

            num_id = num_id_elem.get(f'{{{W_NS}}}val')
            ilvl = ilvl_elem.get(f'{{{W_NS}}}val')

            if num_id not in self.numbering_map or ilvl not in self.numbering_map[num_id]:
                return None

            level_info = self.numbering_map[num_id][ilvl]

            counter_key = (num_id, ilvl)
            if counter_key not in self.level_counters:
                self.level_counters[counter_key] = level_info['start']
            else:
                self.level_counters[counter_key] += 1

            current_num = self.level_counters[counter_key]
            formatted_num = self._format_number(current_num, level_info['format'])

            text_template = level_info['text']
            result = text_template.replace(f'%{int(ilvl) + 1}', formatted_num)

            return result

        except Exception as e:
            print(f"⚠️ 解析段落编号失败: {e}")
            return None

    def reset_counters(self):
        """✨ 重置计数器（用于页眉页脚独立编号）"""
        self.level_counters.clear()


# ================= 辅助类与函数 =================

class DocContentLoader:
    """预加载辅助内容的类（脚注、尾注、批注、页眉、页脚）"""

    def __init__(self, doc_path: str):
        self.doc_path = doc_path
        self.footnotes = {}
        self.endnotes = {}
        self.comments = {}
        self.headers = []  # ✨ 改为存储结构化数据
        self.footers = []  # ✨ 改为存储结构化数据
        self._load_all()

    def _load_xml_map(self, zip_file: ZipFile, filename: str, tag_name: str, id_attr: str = 'id') -> Dict[str, str]:
        """
        通用加载函数：将XML文件解析为 {id: text} 的字典
        ✨ 增强：支持多种 ID 属性格式
        """
        data_map = {}
        if filename not in zip_file.namelist():
            return data_map

        try:
            with zip_file.open(filename) as f:
                tree = etree.parse(f)
                for elem in tree.findall(f'.//w:{tag_name}', NAMESPACES):
                    # ✨ 尝试多种 ID 属性格式
                    eid = (elem.get(f'{{{W_NS}}}{id_attr}') or
                           elem.get(id_attr) or
                           elem.get('id'))

                    elem_type = elem.get(f'{{{W_NS}}}type')
                    if elem_type in ('separator', 'continuationSeparator'):
                        continue

                    texts = [t.text for t in elem.iter(f'{{{W_NS}}}t') if t.text]
                    full_text = "".join(texts).strip()
                    if full_text and eid:
                        data_map[eid] = full_text
        except Exception as e:
            print(f"⚠️ 加载 {filename} 失败: {e}")
        return data_map

    def _load_header_footer_structured(self, zip_file: ZipFile, pattern: str,
                                       numbering_system: 'NumberingSystem') -> List[Tuple[str, List[str]]]:
        """
        ✨ 新增：结构化加载页眉/页脚（支持编号、脚注、文本框）
        返回: [(段落文本, [关联内容列表]), ...]
        """
        content_list = []
        matching_files = [f for f in zip_file.namelist() if re.match(pattern.replace('*', r'\d*'), f)]

        for filename in sorted(matching_files):
            try:
                with zip_file.open(filename) as f:
                    tree = etree.parse(f)

                    # ✨ 重置编号计数器（页眉页脚独立编号）
                    numbering_system.reset_counters()

                    for p in tree.findall('.//w:p', NAMESPACES):
                        # 获取编号
                        number_text = numbering_system.get_paragraph_number(p)

                        # 获取段落文本
                        text = get_xml_text(p)

                        # ✨ 获取关联内容（脚注、文本框等）
                        extras = self._extract_anchored_content_from_element(p)

                        # 组合编号和文本
                        if number_text and text:
                            full_text = f"{number_text} {text}"
                        elif number_text:
                            full_text = number_text
                        elif text:
                            full_text = text
                        else:
                            full_text = ""

                        if full_text.strip() or extras:
                            content_list.append((full_text, extras))

            except Exception as e:
                print(f"⚠️ 加载 {filename} 失败: {e}")

        return content_list

    def _extract_anchored_content_from_element(self, element) -> List[str]:
        """
        ✨ 新增：从任意元素中提取关联内容（脚注、文本框等）
        """
        extras = []

        # 1. 脚注引用（✨ 增强匹配逻辑）
        for ref in element.findall('.//w:footnoteReference', NAMESPACES):
            fid = ref.get(f'{{{W_NS}}}id') or ref.get('id') or ref.get(f'{{{W_NS}}}w:id')
            if fid and fid in self.footnotes:
                extras.append(self.footnotes[fid])
                print(f"✅ 找到脚注引用 ID={fid}: {self.footnotes[fid][:50]}...")

        # 2. 尾注引用
        for ref in element.findall('.//w:endnoteReference', NAMESPACES):
            eid = ref.get(f'{{{W_NS}}}id') or ref.get('id')
            if eid and eid in self.endnotes:
                extras.append(self.endnotes[eid])

        # 3. 批注引用
        for ref in element.findall('.//w:commentReference', NAMESPACES):
            cid = ref.get(f'{{{W_NS}}}id') or ref.get('id')
            if cid and cid in self.comments:
                extras.append(self.comments[cid])

        # 4. 文本框（✨ 扩展检测范围）
        # 4.1 Word 2010+ 文本框
        for txbx in element.iter(f'{{{WPS_NS}}}txbxContent'):
            text = get_xml_text(txbx)
            if text:
                extras.append(text)
                print(f"✅ 找到 wps:txbxContent 文本框: {text[:50]}...")

        # 4.2 兼容模式文本框
        for v_txbx in element.iter(f'{{{V_NS}}}textbox'):
            text = get_xml_text(v_txbx)
            if text:
                extras.append(text)
                print(f"✅ 找到 v:textbox 文本框: {text[:50]}...")

        # ✨ 4.3 检测 <w:txbxContent>（另一种文本框格式）
        for w_txbx in element.iter(f'{{{W_NS}}}txbxContent'):
            text = get_xml_text(w_txbx)
            if text:
                extras.append(text)
                print(f"✅ 找到 w:txbxContent 文本框: {text[:50]}...")

        return extras

    def _load_all(self):
        """✨ 改进：加载所有辅助内容，页眉页脚使用结构化方法"""
        # ✨ 需要临时创建编号系统实例
        temp_numbering = NumberingSystem(self.doc_path)

        with ZipFile(self.doc_path, 'r') as zf:
            # 加载脚注和尾注
            self.footnotes = self._load_xml_map(zf, 'word/footnotes.xml', 'footnote')
            self.endnotes = self._load_xml_map(zf, 'word/endnotes.xml', 'endnote')
            self.comments = self._load_xml_map(zf, 'word/comments.xml', 'comment')

            # ✨ 加载页眉和页脚（结构化）
            self.headers = self._load_header_footer_structured(zf, r'word/header\d*\.xml', temp_numbering)
            self.footers = self._load_header_footer_structured(zf, r'word/footer\d*\.xml', temp_numbering)

            # 调试输出
            print(f"📊 加载统计:")
            print(f"  - 脚注: {len(self.footnotes)} 个")
            print(f"  - 尾注: {len(self.endnotes)} 个")
            print(f"  - 页眉段落: {len(self.headers)} 个")
            print(f"  - 页脚段落: {len(self.footers)} 个")
            if self.footnotes:
                print(f"  - 脚注ID列表: {list(self.footnotes.keys())}")


def get_xml_text(element) -> str:
    """从任意XML元素及其子元素中提取纯文本"""
    texts = []
    for t in element.iter(f"{{{W_NS}}}t"):
        if t.text:
            texts.append(t.text)
    return "".join(texts)


def process_anchored_content(p_element, loader: DocContentLoader) -> List[str]:
    """
    ✨ 重构：直接调用 loader 的统一方法
    """
    return loader._extract_anchored_content_from_element(p_element)


def extract_doc_text(doc_path: str) -> str:
    """
    对外调用入口：传入 docx 路径，返回提取后的全文字符串
    ✨ 改进：页眉页脚支持完整解析
    """
    if not os.path.exists(doc_path):
        raise FileNotFoundError(f"文件不存在: {doc_path}")

    # 1) 预加载辅助内容
    loader = DocContentLoader(doc_path)

    # 2) 初始化编号系统
    numbering_system = NumberingSystem(doc_path)

    # 3) 读取主文档 XML
    doc = Document(doc_path)
    body_element = doc.element.body

    output_lines = []

    # ✨ --- 页眉内容（结构化输出） ---
    if loader.headers:
        output_lines.append("=== 页眉内容 ===")
        for main_text, extras in loader.headers:
            if main_text.strip():
                output_lines.append(main_text)
            for extra in extras:
                output_lines.append(extra)
        output_lines.append("")

    # --- 正文 ---
    # ✨ 重置编号计数器（正文独立编号）
    output_lines.append("=== 正文内容 ===")
    numbering_system.reset_counters()

    for child in body_element.iterchildren():
        tag_name = child.tag

        # 段落
        if tag_name.endswith('p'):
            number_text = numbering_system.get_paragraph_number(child)
            text = get_xml_text(child)
            extras = process_anchored_content(child, loader)

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
        elif tag_name.endswith('tbl'):
            for row in child.iter(f'{{{W_NS}}}tr'):
                row_texts = []
                for cell in row.iter(f'{{{W_NS}}}tc'):
                    cell_content = []
                    for cell_p in cell.iter(f'{{{W_NS}}}p'):
                        cell_number = numbering_system.get_paragraph_number(cell_p)
                        p_text = get_xml_text(cell_p)

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

                        # ✨ 单元格内的关联内容
                        cell_extras = process_anchored_content(cell_p, loader)
                        cell_content.extend(cell_extras)

                    full_cell_text = "\t".join(cell_content)
                    row_texts.append(full_cell_text)

                if any(row_texts):
                    output_lines.append("\t".join(row_texts))

    # ✨ --- 页脚内容（结构化输出） ---
    if loader.footers:
        output_lines.append("")
        output_lines.append("=== 页脚内容 ===")
        for main_text, extras in loader.footers:
            if main_text.strip():
                output_lines.append(main_text)
            for extra in extras:
                output_lines.append(extra)

    return "\n".join(output_lines)


def main():
    """主函数：演示完整的文档提取流程"""
    doc_path = os.path.join(SOURCE_DIR, CHINESE_FILE)
    if not os.path.exists(doc_path):
        print(f"❌ 文件不存在: {doc_path}")
        return

    print(f"📄 正在分析文档: {CHINESE_FILE}\n")

    try:
        full_text = extract_doc_text(doc_path)

        print("\n" + "=" * 50)
        print(full_text)
        print("=" * 50)

        # 保存到文件
        output_file = os.path.join(SOURCE_DIR, "extracted_output.txt")
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(full_text)
        print(f"\n✅ 提取结果已保存至: {output_file}")

    except Exception as e:
        print(f"❌ 提取失败: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()