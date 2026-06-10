# -*- coding: utf-8 -*-
"""
合并提取Word文档中所有文字
包括：正文、表格、脚注、尾注、文本框、批注、图表(Chart)等
图表策略：优先提取数据标签实际显示内容，无标签时回退到原始数值
"""
import os
import re
from typing import Dict, Optional, List, Tuple, Set
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
# 图表相关命名空间
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
C15_NS = "http://schemas.microsoft.com/office/drawing/2012/chart"

NAMESPACES = {
    "w": W_NS, "wp": WP_NS, "a": A_NS, "wps": WPS_NS,
    "v": V_NS, "r": R_NS, "m": M_NS, "c": C_NS, "c15": C15_NS,
}

# 文本框容器标签（用于跳过判断）
_TEXTBOX_TAGS = {
    f"{{{WPS_NS}}}txbxContent",
    f"{{{V_NS}}}textbox",
}


# ============================================================
#  编号系统（原有代码，无修改）
# ============================================================
class NumberingSystem:
    """处理 Word 自动编号系统（含 lvlOverride 和样式编号回退）"""

    def __init__(self, doc_path: str):
        self.doc_path = doc_path
        self.numbering_map = {}
        self.abstract_num_map = {}
        self.level_counters = {}
        self.style_num_map = {}
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
                                "format": fmt_val, "text": text_val, "start": start_val,
                            }

                    for num in tree.findall(".//w:num", NAMESPACES):
                        num_id = num.get(f"{{{W_NS}}}numId")
                        abstract_num_id_elem = num.find(".//w:abstractNumId", NAMESPACES)
                        if abstract_num_id_elem is None:
                            continue
                        abstract_num_id = abstract_num_id_elem.get(f"{{{W_NS}}}val")
                        if abstract_num_id in self.abstract_num_map:
                            level_map = {}
                            for k, v in self.abstract_num_map[abstract_num_id].items():
                                level_map[k] = v.copy()

                            for override in num.findall(f"{{{W_NS}}}lvlOverride"):
                                ilvl = override.get(f"{{{W_NS}}}ilvl")
                                if ilvl is None:
                                    continue
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

                                start_override = override.find(f"{{{W_NS}}}startOverride")
                                if start_override is not None and ilvl in level_map:
                                    level_map[ilvl]["start"] = int(
                                        start_override.get(f"{{{W_NS}}}val", "1")
                                    )

                            self.numbering_map[num_id] = level_map

        except Exception as e:
            print(f"加载编号系统失败: {e}")

    def _load_style_numbering(self):
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
                ppr = p_element.find(".//w:pPr", NAMESPACES)
                if ppr is not None:
                    pstyle = ppr.find(".//w:pStyle", NAMESPACES)
                    if pstyle is not None:
                        style_id = pstyle.get(f"{{{W_NS}}}val")
                        if style_id and style_id in self.style_num_map:
                            num_id, ilvl = self.style_num_map[style_id]
                            return self._resolve_number(num_id, ilvl)
                return None

            num_id_elem = num_pr.find(".//w:numId", NAMESPACES)
            ilvl_elem = num_pr.find(".//w:ilvl", NAMESPACES)
            if num_id_elem is None or ilvl_elem is None:
                return None
            num_id = num_id_elem.get(f"{{{W_NS}}}val")
            ilvl = ilvl_elem.get(f"{{{W_NS}}}val")
            if num_id == "0":
                return None
            return self._resolve_number(num_id, ilvl)
        except Exception as e:
            print(f"解析段落编号失败: {e}")
            return None

    def _resolve_number(self, num_id: str, ilvl: str) -> Optional[str]:
        if num_id not in self.numbering_map or ilvl not in self.numbering_map[num_id]:
            return None
        ilvl_int = int(ilvl)
        level_info = self.numbering_map[num_id][ilvl]
        counter_key = (num_id, ilvl)
        if counter_key not in self.level_counters:
            self.level_counters[counter_key] = level_info["start"]
        else:
            self.level_counters[counter_key] += 1
        for other_ilvl_str, other_info in self.numbering_map[num_id].items():
            other_ilvl_int = int(other_ilvl_str)
            if other_ilvl_int > ilvl_int:
                other_key = (num_id, other_ilvl_str)
                if other_key in self.level_counters:
                    del self.level_counters[other_key]
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


# ============================================================
#  字符样式加粗/斜体加载（处理 rStyle 继承链）
# ============================================================
class StyleFormatLoader:
    """
    加载 styles.xml 中每个字符样式的加粗/斜体属性。
    支持 basedOn 继承链（最多 10 层防止循环）。
    """

    def __init__(self, doc_path: str):
        # style_id -> {"bold": bool|None, "italic": bool|None, "basedOn": str|None}
        self._raw: Dict[str, Dict] = {}
        # 缓存最终解析结果
        self._bold_cache: Dict[str, bool] = {}
        self._italic_cache: Dict[str, bool] = {}
        self._load(doc_path)

    def _load(self, doc_path: str) -> None:
        try:
            with ZipFile(doc_path, "r") as zf:
                if "word/styles.xml" not in zf.namelist():
                    return
                with zf.open("word/styles.xml") as f:
                    tree = etree.parse(f)
                for style in tree.findall(".//w:style", NAMESPACES):
                    sid = style.get(f"{{{W_NS}}}styleId")
                    if not sid:
                        continue
                    rpr = style.find(f"{{{W_NS}}}rPr", NAMESPACES)
                    bold = italic = None
                    if rpr is not None:
                        b = rpr.find(f"{{{W_NS}}}b")
                        if b is not None:
                            val = b.get(f"{{{W_NS}}}val")
                            bold = val not in ["0", "false", "off"]
                        i = rpr.find(f"{{{W_NS}}}i")
                        if i is not None:
                            val = i.get(f"{{{W_NS}}}val")
                            italic = val not in ["0", "false", "off"]
                    based = style.find(f"{{{W_NS}}}basedOn")
                    based_id = based.get(f"{{{W_NS}}}val") if based is not None else None
                    self._raw[sid] = {"bold": bold, "italic": italic, "basedOn": based_id}
        except Exception as e:
            print(f"加载字符样式失败: {e}")

    def _resolve(self, style_id: str, attr: str, visited: Set[str]) -> bool:
        if style_id in visited or style_id not in self._raw:
            return False
        visited.add(style_id)
        info = self._raw[style_id]
        if info[attr] is not None:
            return info[attr]
        if info["basedOn"]:
            return self._resolve(info["basedOn"], attr, visited)
        return False

    def is_bold(self, style_id: str) -> bool:
        if style_id not in self._bold_cache:
            self._bold_cache[style_id] = self._resolve(style_id, "bold", set())
        return self._bold_cache[style_id]

    def is_italic(self, style_id: str) -> bool:
        if style_id not in self._italic_cache:
            self._italic_cache[style_id] = self._resolve(style_id, "italic", set())
        return self._italic_cache[style_id]


# ============================================================
#  辅助内容加载：脚注/尾注/批注（原有代码，无修改）
# ============================================================
class DocAnchorsLoader:
    def __init__(self, doc_path: str):
        self.doc_path = doc_path
        self.footnotes: Dict[str, str] = {}
        self.endnotes: Dict[str, str] = {}
        self.comments: Dict[str, str] = {}
        self._load_all()

    def _load_xml_map(self, zip_file: ZipFile, filename: str, tag_name: str,
                      id_attr: str = "id") -> Dict[str, str]:
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
                    full_text = _get_xml_text(elem, skip_containers=False)
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


# ============================================================
#  图表(Chart)提取系统（新增）
# ============================================================
class ChartExtractor:
    """
    从 .docx 内嵌的 chart XML 中提取所有可见文字。
    策略：
      1. 始终提取图表标题、轴标题
      2. 优先提取数据标签（datalabelsRange 缓存 > 内联 dLbl 文字）
      3. 无数据标签时才回退到原始数值
    """

    SEP = "\t"  # 值之间的分隔符

    def __init__(self, doc_path: str):
        self.doc_path = doc_path
        self.rid_to_chart_path: Dict[str, str] = {}
        self.chart_texts: Dict[str, List[str]] = {}
        self._load_chart_rels()
        self._parse_all_charts()

    # ---------- 加载关系 ----------
    def _load_chart_rels(self) -> None:
        """从 document.xml.rels 中找到所有 rId → charts/chartN.xml 的映射"""
        rels_path = "word/_rels/document.xml.rels"
        try:
            with ZipFile(self.doc_path, "r") as zf:
                if rels_path not in zf.namelist():
                    return
                with zf.open(rels_path) as f:
                    tree = etree.parse(f)
                    for rel in tree.getroot():
                        rel_type = rel.get("Type", "")
                        target = rel.get("Target", "")
                        rid = rel.get("Id", "")
                        if "chart" in rel_type.lower() and rid:
                            chart_path = target if target.startswith("word/") else f"word/{target}"
                            self.rid_to_chart_path[rid] = chart_path
        except Exception as e:
            print(f"加载图表关系失败: {e}")

    # ---------- 解析所有图表 ----------
    def _parse_all_charts(self) -> None:
        try:
            with ZipFile(self.doc_path, "r") as zf:
                # 通过 rels 找到的图表
                for rid, chart_path in self.rid_to_chart_path.items():
                    if chart_path in zf.namelist():
                        with zf.open(chart_path) as f:
                            tree = etree.parse(f)
                            self.chart_texts[chart_path] = self._extract_chart_texts(tree.getroot())

                # 兜底扫描：直接查找 word/charts/ 目录
                for name in zf.namelist():
                    if re.match(r"word/charts/chart\d*\.xml$", name) and name not in self.chart_texts:
                        with zf.open(name) as f:
                            tree = etree.parse(f)
                            self.chart_texts[name] = self._extract_chart_texts(tree.getroot())
        except Exception as e:
            print(f"解析图表文件失败: {e}")

    # ---------- 单个图表的文字提取 ----------
    def _extract_chart_texts(self, root) -> List[str]:
        texts: List[str] = []

        # 1. 图表标题
        title = self._get_chart_title(root)
        if title:
            texts.append(f"[图表标题] {title}")

        # 2. 轴标题
        texts.extend(self._get_axis_texts(root))

        # 3. 系列数据（数据标签优先）
        texts.extend(self._get_series_with_labels(root))

        # 4. 兜底：chart XML 中所有 a:t，去重后补充遗漏的文字
        fallback = self._get_all_drawingml_text(root)
        existing: Set[str] = set()
        for t in texts:
            for word in re.split(r'[\[\]\s]+', t):
                if word:
                    existing.add(word)
        for ft in fallback:
            if ft in existing or ft == "[CELLRANGE]":
                continue
            try:
                float(ft)
                continue  # 跳过纯数值，避免把坐标值混入
            except ValueError:
                pass
            texts.append(ft)
            existing.add(ft)

        return texts

    # ---------- 图表标题 ----------
    def _get_chart_title(self, root) -> Optional[str]:
        chart_elem = root.find(f"{{{C_NS}}}chart")
        if chart_elem is None:
            return None
        title_elem = chart_elem.find(f"{{{C_NS}}}title")
        if title_elem is None:
            return None

        # 优先：c:tx/c:rich 富文本
        tx = title_elem.find(f"{{{C_NS}}}tx")
        if tx is not None:
            rich = tx.find(f"{{{C_NS}}}rich")
            if rich is not None:
                parts: List[str] = []
                for para in rich.findall(f"{{{A_NS}}}p"):
                    para_parts = [at.text for run in para.findall(f"{{{A_NS}}}r")
                                  for at in run.findall(f"{{{A_NS}}}t") if at.text]
                    if para_parts:
                        parts.append("".join(para_parts))
                if parts:
                    return "\n".join(parts).strip()

            # 备选：c:tx/c:strRef
            str_ref = tx.find(f"{{{C_NS}}}strRef")
            if str_ref is not None:
                cache = str_ref.find(f"{{{C_NS}}}strCache")
                if cache is not None:
                    for pt in cache.findall(f"{{{C_NS}}}pt"):
                        v = pt.find(f"{{{C_NS}}}v")
                        if v is not None and v.text:
                            return v.text.strip()

        # 兜底：title 下所有 a:t
        parts = [at.text for at in title_elem.iter(f"{{{A_NS}}}t") if at.text]
        return "".join(parts).strip() if parts else None

    # ---------- 轴标题 ----------
    def _get_axis_texts(self, root) -> List[str]:
        results: List[str] = []
        axis_tags = [
            (f"{{{C_NS}}}catAx", "分类轴"), (f"{{{C_NS}}}valAx", "值轴"),
            (f"{{{C_NS}}}dateAx", "日期轴"), (f"{{{C_NS}}}serAx", "系列轴"),
        ]
        for axis_tag, axis_name in axis_tags:
            for axis in root.iter(axis_tag):
                title_elem = axis.find(f"{{{C_NS}}}title")
                if title_elem is not None:
                    parts = [at.text.strip() for at in title_elem.iter(f"{{{A_NS}}}t") if at.text]
                    if parts:
                        results.append(f"[{axis_name}标题] {''.join(parts)}")
        return results

    # ---------- 系列数据（数据标签优先）----------
    def _get_series_with_labels(self, root) -> List[str]:
        results: List[str] = []
        chart_type_tags = [
            "barChart", "bar3DChart", "lineChart", "line3DChart",
            "pieChart", "pie3DChart", "doughnutChart",
            "areaChart", "area3DChart", "scatterChart", "bubbleChart",
            "radarChart", "surfaceChart", "surface3DChart",
            "stockChart", "ofPieChart",
        ]
        for chart_type in chart_type_tags:
            for chart_elem in root.iter(f"{{{C_NS}}}{chart_type}"):
                for ser in chart_elem.iter(f"{{{C_NS}}}ser"):
                    results.extend(self._parse_series_smart(ser))
        return results

    def _parse_series_smart(self, ser_elem) -> List[str]:
        """智能解析：数据标签优先，无标签时回退到原始数值"""
        parts: List[str] = []
        sep = self.SEP

        # 系列名（始终提取）
        tx = ser_elem.find(f"{{{C_NS}}}tx")
        if tx is not None:
            name = self._get_str_or_ref(tx)
            if name:
                parts.append(f"  [系列名] {name}")

        # 分类名（始终提取）
        cat = ser_elem.find(f"{{{C_NS}}}cat")
        if cat is not None:
            cat_vals = self._get_values(cat)
            if cat_vals:
                parts.append(f"  [分类] {sep.join(cat_vals)}")

        # ---- 数据标签优先逻辑 ----
        range_labels = self._extract_datalabels_range_cache(ser_elem)
        inline_labels = self._extract_inline_data_labels(ser_elem)
        has_labels = False

        # 优先级1：datalabelsRange 缓存（解析 [CELLRANGE]）
        if range_labels:
            non_empty = [lb for lb in range_labels if lb]
            if non_empty:
                has_labels = True
                parts.append(f"  [数据标签] {sep.join(range_labels)}")

        # 优先级2：内联 dLbl 文字（过滤掉 [CELLRANGE] 占位符）
        if not has_labels and inline_labels:
            meaningful = [lb for lb in inline_labels if lb and lb != "[CELLRANGE]"]
            if meaningful:
                has_labels = True
                parts.append(f"  [数据标签] {sep.join(meaningful)}")

        # 优先级3：无标签时回退到原始数值
        if not has_labels:
            for tag_name, label in [("val", "数据"), ("xVal", "X值"),
                                    ("yVal", "Y值"), ("bubbleSize", "气泡大小")]:
                elem = ser_elem.find(f"{{{C_NS}}}{tag_name}")
                if elem is not None:
                    vals = self._get_values(elem)
                    if vals:
                        parts.append(f"  [{label}] {sep.join(vals)}")

        return parts

    # ---------- datalabelsRange 缓存提取 ----------
    def _extract_datalabels_range_cache(self, ser_elem) -> List[str]:
        """
        解析 c15:datalabelsRange/dlblRangeCache，
        这是 [CELLRANGE] 占位符对应的真实文字。
        """
        # 方式1：精确命名空间
        for dlbl_range in ser_elem.iter(f"{{{C15_NS}}}datalabelsRange"):
            cache = dlbl_range.find(f"{{{C15_NS}}}dlblRangeCache")
            if cache is not None:
                result = self._extract_cache_values(cache)
                if result:
                    return result

        # 方式2：通配符 localname 兜底（兼容不同 Word 版本）
        for ext in ser_elem.iter(f"{{{C_NS}}}ext"):
            for child in ext.iter():
                local = etree.QName(child.tag).localname if isinstance(child.tag, str) else ""
                if local == "datalabelsRange":
                    for cache_child in child.iter():
                        cache_local = etree.QName(cache_child.tag).localname if isinstance(cache_child.tag, str) else ""
                        if cache_local == "dlblRangeCache":
                            result = self._extract_cache_values(cache_child)
                            if result:
                                return result
        return []

    def _extract_cache_values(self, cache_elem) -> List[str]:
        """按 idx 顺序提取缓存中的值"""
        pt_map: Dict[int, str] = {}
        for pt in cache_elem.iter():
            local = etree.QName(pt.tag).localname if isinstance(pt.tag, str) else ""
            if local != "pt":
                continue
            idx_str = pt.get("idx", "") or pt.get(f"{{{C_NS}}}idx", "")
            try:
                idx = int(idx_str) if idx_str else len(pt_map)
            except ValueError:
                idx = len(pt_map)

            for v_child in pt:
                if etree.QName(v_child.tag).localname == "v" and v_child.text:
                    pt_map[idx] = v_child.text.strip()
                    break

        if not pt_map:
            return []
        return [pt_map.get(i, "") for i in range(max(pt_map.keys()) + 1)]

    # ---------- 内联 dLbl 文字提取 ----------
    def _extract_inline_data_labels(self, ser_elem) -> List[str]:
        """从 c:dLbl 中提取逐点数据标签文字"""
        label_map: Dict[int, str] = {}
        for dlbl in ser_elem.iter(f"{{{C_NS}}}dLbl"):
            idx_elem = dlbl.find(f"{{{C_NS}}}idx")
            if idx_elem is None:
                continue
            try:
                idx = int(idx_elem.get("val", idx_elem.get(f"{{{C_NS}}}val", "-1")))
            except (ValueError, TypeError):
                continue
            parts = [at.text.strip() for at in dlbl.iter(f"{{{A_NS}}}t") if at.text and at.text.strip()]
            if parts:
                label_map[idx] = "".join(parts)

        if not label_map:
            return []
        return [label_map.get(i, "") for i in range(max(label_map.keys()) + 1)]

    # ---------- 基础提取方法 ----------
    def _get_str_or_ref(self, elem) -> str:
        v = elem.find(f"{{{C_NS}}}v")
        if v is not None and v.text:
            return v.text.strip()
        str_ref = elem.find(f"{{{C_NS}}}strRef")
        if str_ref is not None:
            cache = str_ref.find(f"{{{C_NS}}}strCache")
            if cache is not None:
                for pt in cache.findall(f"{{{C_NS}}}pt"):
                    pv = pt.find(f"{{{C_NS}}}v")
                    if pv is not None and pv.text:
                        return pv.text.strip()
        parts = [at.text.strip() for at in elem.iter(f"{{{A_NS}}}t") if at.text]
        return "".join(parts)

    def _get_values(self, elem) -> List[str]:
        values: List[str] = []
        # strCache
        for cache in elem.iter(f"{{{C_NS}}}strCache"):
            for pt in cache.findall(f"{{{C_NS}}}pt"):
                v = pt.find(f"{{{C_NS}}}v")
                if v is not None and v.text:
                    values.append(v.text.strip())
        # numCache
        if not values:
            for cache in elem.iter(f"{{{C_NS}}}numCache"):
                for pt in cache.findall(f"{{{C_NS}}}pt"):
                    v = pt.find(f"{{{C_NS}}}v")
                    if v is not None and v.text:
                        values.append(v.text.strip())
        # multiLvlStrCache
        if not values:
            for cache in elem.iter(f"{{{C_NS}}}multiLvlStrCache"):
                for lvl in cache.findall(f"{{{C_NS}}}lvl"):
                    for pt in lvl.findall(f"{{{C_NS}}}pt"):
                        v = pt.find(f"{{{C_NS}}}v")
                        if v is not None and v.text:
                            values.append(v.text.strip())
        # strLit / numLit
        if not values:
            for lit_tag in (f"{{{C_NS}}}strLit", f"{{{C_NS}}}numLit"):
                for lit in elem.iter(lit_tag):
                    for pt in lit.findall(f"{{{C_NS}}}pt"):
                        v = pt.find(f"{{{C_NS}}}v")
                        if v is not None and v.text:
                            values.append(v.text.strip())
        return values

    def _get_all_drawingml_text(self, root) -> List[str]:
        texts: List[str] = []
        seen: Set[str] = set()
        for at in root.iter(f"{{{A_NS}}}t"):
            if at.text and at.text.strip():
                t = at.text.strip()
                if t not in seen:
                    texts.append(t)
                    seen.add(t)
        return texts

    # ---------- 外部接口 ----------
    def get_chart_text_by_rid(self, rid: str) -> List[str]:
        chart_path = self.rid_to_chart_path.get(rid, "")
        return self.chart_texts.get(chart_path, [])

    def get_all_chart_texts(self) -> Dict[str, List[str]]:
        return self.chart_texts


# ============================================================
#  文本提取（在原有基础上增加跳过图表/文本框子树的能力）
# ============================================================
def _get_xml_text(element, *, skip_containers: bool = True,
                  style_loader: "StyleFormatLoader | None" = None) -> str:
    """
    增强版文本提取，支持加粗/斜体标签。

    参数:
        skip_containers: True 时跳过文本框和图表 drawing 内部的节点，
                         避免在正文段落中重复提取。
                         脚注/尾注等独立 XML 中应传 False。
    """
    SYMBOL_CHAR_MAP = {
        'F0B4': '×', 'F0B8': '÷', 'F0B1': '±', 'F0B3': '≥',
        'F0A3': '≤', 'F0B9': '≠', 'F0BB': '≈',
        '\uf052': '☑', '\uf0a3': '☑', '': '☑', '': '☑',
        '\uf0a1': '☐', '': '☐', '\u25a1': '☐', '\u25fb': '☐',
        'F052': '☑', 'F0A1': '☐'
    }
    W14_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"

    # 预先收集需要跳过的子树根节点
    skip_ids: Set[int] = set()
    if skip_containers:
        # 跳过文本框
        for tag in _TEXTBOX_TAGS:
            for container in element.iter(tag):
                skip_ids.add(id(container))
        # 跳过包含图表引用的 drawing
        for drawing in element.iter(f"{{{W_NS}}}drawing"):
            if drawing.find(f".//{{{C_NS}}}chart") is not None:
                skip_ids.add(id(drawing))

    def _in_skip_subtree(node) -> bool:
        """判断节点是否在需要跳过的子树内"""
        if not skip_ids:
            return False
        if id(node) in skip_ids:
            return True
        parent = node.getparent()
        while parent is not None:
            if id(parent) in skip_ids:
                return True
            parent = parent.getparent()
        return False

    full_texts: List[str] = []

    for node in element.iter():

        # --- 1. 处理 Run (w:r) ---
        if node.tag == f"{{{W_NS}}}r":
            if skip_containers and _in_skip_subtree(node):
                continue

            rPr = node.find(f"{{{W_NS}}}rPr")
            is_bold = False
            is_italic = False

            if rPr is not None:
                # 先检查 run 自身的 w:b
                b_elem = rPr.find(f"{{{W_NS}}}b")
                if b_elem is not None:
                    val = b_elem.get(f"{{{W_NS}}}val")
                    is_bold = val not in ["0", "false", "off"]

                # 先检查 run 自身的 w:i
                i_elem = rPr.find(f"{{{W_NS}}}i")
                if i_elem is not None:
                    val = i_elem.get(f"{{{W_NS}}}val")
                    is_italic = val not in ["0", "false", "off"]

                # 若 run 自身没有显式设置，查 rStyle 继承的样式
                if style_loader is not None and (not is_bold or not is_italic):
                    r_style_elem = rPr.find(f"{{{W_NS}}}rStyle")
                    if r_style_elem is not None:
                        r_style_id = r_style_elem.get(f"{{{W_NS}}}val", "")
                        if r_style_id:
                            if not is_bold and b_elem is None:
                                is_bold = style_loader.is_bold(r_style_id)
                            if not is_italic and i_elem is None:
                                is_italic = style_loader.is_italic(r_style_id)

            run_content: List[str] = []
            for child in node.iterchildren():
                if child.tag == f"{{{W_NS}}}t":
                    if child.text:
                        t = child.text
                        for raw, char in SYMBOL_CHAR_MAP.items():
                            if len(raw) > 2:
                                t = t.replace(raw, char)
                        run_content.append(t)
                elif child.tag == f"{{{W_NS}}}sym":
                    char_code = child.get(f"{{{W_NS}}}char", "").upper()
                    if char_code in SYMBOL_CHAR_MAP:
                        run_content.append(SYMBOL_CHAR_MAP[char_code])
                    elif char_code:
                        try:
                            char_val = chr(int(char_code, 16))
                            run_content.append(SYMBOL_CHAR_MAP.get(char_val, char_val))
                        except (ValueError, OverflowError):
                            run_content.append(f"[{char_code}]")

            if run_content:
                text_segment = "".join(run_content)
                if text_segment.strip():
                    if is_bold:
                        text_segment = f"<bold>{text_segment}</bold>"
                    if is_italic:
                        text_segment = f"<italic>{text_segment}</italic>"
                full_texts.append(text_segment)

        # --- 2. 数学公式文本 (m:t) ---
        elif node.tag == f"{{{M_NS}}}t":
            if skip_containers and _in_skip_subtree(node):
                continue
            if node.text:
                full_texts.append(node.text)

        # --- 3. 数学特殊字符 (m:char) ---
        elif node.tag == f"{{{M_NS}}}char":
            if skip_containers and _in_skip_subtree(node):
                continue
            val = node.get(f"{{{M_NS}}}val") or node.get(f"{{{W_NS}}}val")
            if val:
                full_texts.append(val)

        # --- 4. 复选框 (SDT) ---
        elif node.tag.endswith("sdt"):
            if skip_containers and _in_skip_subtree(node):
                continue
            checkbox = node.find(f".//{{{W14_NS}}}checkbox")
            if checkbox is not None:
                checked = checkbox.find(f".//{{{W14_NS}}}checked")
                val = checked.get(f"{{{W14_NS}}}val") if checked is not None else "0"
                full_texts.append("☑" if val in ["1", "true"] else "☐")

        # --- 5. 图片描述 ---
        elif node.tag == f"{{{WP_NS}}}docPr":
            if skip_containers and _in_skip_subtree(node):
                continue
            alt = node.get("descr") or node.get("title")
            if alt:
                full_texts.append(f"[图片描述: {alt}]")

    return "".join(full_texts).strip()


# ============================================================
#  锚定内容处理（修复文本框重复）
# ============================================================
def _process_anchored_content(p_element, loader: DocAnchorsLoader,
                              style_loader: "StyleFormatLoader | None" = None) -> List[str]:
    """
    提取段落中的锚定内容：文本框、脚注、尾注、批注。
    文本框去重：只提取 wps:txbxContent，v:textbox 中若已包含则跳过。
    """
    extras: List[str] = []

    # 1. 脚注
    for ref in p_element.findall(".//w:footnoteReference", NAMESPACES):
        fid = ref.get(f"{{{W_NS}}}id")
        if fid and fid in loader.footnotes:
            extras.append(loader.footnotes[fid])

    # 2. 尾注
    for ref in p_element.findall(".//w:endnoteReference", NAMESPACES):
        eid = ref.get(f"{{{W_NS}}}id")
        if eid and eid in loader.endnotes:
            extras.append(loader.endnotes[eid])

    # 3. 批注
    for ref in p_element.findall(".//w:commentReference", NAMESPACES):
        cid = ref.get(f"{{{W_NS}}}id")
        if cid and cid in loader.comments:
            extras.append(loader.comments[cid])

    # 4. 文本框（去重处理）
    processed_txbx_ids: Set[int] = set()

    # 优先提取 wps:txbxContent（新版格式，最内层）
    for txbx in p_element.iter(f"{{{WPS_NS}}}txbxContent"):
        processed_txbx_ids.add(id(txbx))
        text = _get_xml_text(txbx, skip_containers=False, style_loader=style_loader)
        if text:
            extras.append(text)

    # 再提取 v:textbox（旧版格式），跳过已被上面处理过的
    for v_txbx in p_element.iter(f"{{{V_NS}}}textbox"):
        # 检查内部是否有已处理的 wps:txbxContent
        already_done = any(
            id(child) in processed_txbx_ids
            for child in v_txbx.iter(f"{{{WPS_NS}}}txbxContent")
        )
        if already_done:
            continue  # 已通过 wps:txbxContent 提取过
        text = _get_xml_text(v_txbx, skip_containers=False, style_loader=style_loader)
        if text:
            extras.append(text)

    return extras


# ============================================================
#  图表引用提取（新增）
# ============================================================
def _extract_chart_rids_from_element(element) -> List[str]:
    """从元素中找到所有嵌入图表的 rId"""
    rids: List[str] = []
    for chart_ref in element.iter(f"{{{C_NS}}}chart"):
        rid = chart_ref.get(f"{{{R_NS}}}id")
        if rid:
            rids.append(rid)
    return rids


# ============================================================
#  主函数
# ============================================================
def extract_body_text(doc_path: str) -> str:
    """
    提取正文（线性顺序）：段落 + 表格 + 图表
    脚注/尾注/批注/文本框插入到对应锚点段落后
    图表内容插入到包含图表引用的段落后
    """
    if not os.path.exists(doc_path):
        raise FileNotFoundError(f"文件不存在: {doc_path}")

    loader = DocAnchorsLoader(doc_path)
    numbering_system = NumberingSystem(doc_path)
    chart_extractor = ChartExtractor(doc_path)
    style_loader = StyleFormatLoader(doc_path)

    doc = Document(doc_path)
    body_element = doc.element.body

    output_lines: List[str] = []
    inserted_chart_paths: Set[str] = set()  # 记录已插入的图表，用于兜底

    def _process_paragraph(p_elem) -> List[str]:
        """处理单个段落：正文文字 + 锚定内容 + 图表"""
        lines: List[str] = []

        number_text = numbering_system.get_paragraph_number(p_elem)
        # 正文提取时跳过文本框和图表内部节点
        text = _get_xml_text(p_elem, skip_containers=True, style_loader=style_loader)
        extras = _process_anchored_content(p_elem, loader, style_loader=style_loader)

        if number_text and text:
            full_text = f"{number_text} {text}"
        elif number_text:
            full_text = number_text
        elif text:
            full_text = text
        else:
            full_text = ""

        if full_text.strip():
            lines.append(full_text)

        for extra in extras:
            lines.append(extra)

        # 提取段落中嵌入的图表
        chart_rids = _extract_chart_rids_from_element(p_elem)
        for rid in chart_rids:
            chart_path = chart_extractor.rid_to_chart_path.get(rid, "")
            if chart_path:
                inserted_chart_paths.add(chart_path)
            chart_texts = chart_extractor.get_chart_text_by_rid(rid)
            if chart_texts:
                lines.append("[图表内容开始]")
                lines.extend(chart_texts)
                lines.append("[图表内容结束]")

        return lines

    for child in body_element.iterchildren():
        tag_name = child.tag

        # 段落
        if tag_name.endswith("p"):
            output_lines.extend(_process_paragraph(child))

        # 表格
        elif tag_name.endswith("tbl"):
            for row in child.iter(f"{{{W_NS}}}tr"):
                row_texts: List[str] = []
                for cell in row.iter(f"{{{W_NS}}}tc"):
                    cell_content: List[str] = []
                    for cell_p in cell.iter(f"{{{W_NS}}}p"):
                        cell_content.extend(_process_paragraph(cell_p))
                    row_texts.append("\t".join(cell_content))

                if any(row_texts):
                    output_lines.append("\t".join(row_texts))

    # 兜底：未被任何段落引用的图表
    all_chart_texts = chart_extractor.get_all_chart_texts()
    for chart_path, texts in all_chart_texts.items():
        if chart_path not in inserted_chart_paths and texts:
            output_lines.append("[未关联图表内容开始]")
            output_lines.extend(texts)
            output_lines.append("[未关联图表内容结束]")

    result = "\n".join(output_lines)
    # 合并相邻相同格式的 run（Word 内部可能把同格式文字拆成多个 run）
    result = re.sub(r'</bold><bold>', '', result)
    result = re.sub(r'</italic><italic>', '', result)
    return result


# ============================================================
#  测试
# ============================================================


if __name__ == "__main__":
    import io
    import sys

    # 设置标准输出为UTF-8编码
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

    # 示例文件路径
    original_path = r"C:\Users\H\Desktop\测试项目\中翻译\测试文件\Proposal on the 15th Five-Year Development Plan of ICBC.docx"
    translated_path = r"C:\Users\H\Desktop\测试项目\中翻译\测试文件\Proposal on the 15th Five-Year Development Plan of ICBC.docx"# trans_docx_path = translated_path

    # 提取文件中的文本
    original_doc_path = original_path
    original_text = extract_body_text(original_doc_path)
    print(original_text)

