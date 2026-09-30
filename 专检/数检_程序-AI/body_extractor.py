# -*- coding: utf-8 -*-
"""
合并提取Word文档中所有文字
包括：正文、表格、脚注、尾注、页眉、页脚、文本框、批注、图表(Chart)等
图表提取策略：优先提取数据标签的实际显示内容，而非硬提取原始数值
"""
import os
import re
from typing import Dict, Optional, List, Tuple, Set
from zipfile import ZipFile

from docx import Document
from lxml import etree

# ---------------- XML命名空间 ----------------
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
V_NS = "urn:schemas-microsoft-com:vml"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
C15_NS = "http://schemas.microsoft.com/office/drawing/2012/chart"
C16R_NS = "http://schemas.microsoft.com/office/drawing/2017/03/chart"

NAMESPACES = {
    "w": W_NS, "wp": WP_NS, "a": A_NS,
    "wps": WPS_NS, "v": V_NS, "r": R_NS, "m": M_NS,
    "c": C_NS, "rel": REL_NS,
    "c15": C15_NS, "c16r": C16R_NS,
}

# 需要跳过的容器标签集合（文本框和图表，由专门的函数处理）
_SKIP_CONTAINER_TAGS = {
    f"{{{WPS_NS}}}txbxContent",  # 新版文本框内容
    f"{{{V_NS}}}textbox",        # 旧版文本框
}


# ---------------- XML文本提取 ----------------
def _get_xml_text(element, *, skip_chart_drawing: bool = True,
                  skip_textbox: bool = True) -> str:
    """
    从 XML 元素中提取所有可见文字。

    参数:
        skip_chart_drawing: 跳过图表 drawing 内的 a:t，避免图表文字重复
        skip_textbox: 跳过文本框内部节点，避免文本框文字重复
                      （文本框内容由 _process_anchored_content 专门处理）
    """
    SYMBOL_CHAR_MAP = {
        'F0B4': '×', 'F0B8': '÷', 'F0B1': '±', 'F0B3': '≥',
        'F0A3': '≤', 'F0B9': '≠', 'F0BB': '≈',
        '\uf052': '☑', '\uf0a3': '☑', '': '☑', '': '☑',
        '\uf0a1': '☐', '': '☐', '\u25a1': '☐', '\u25fb': '☐',
        'F052': '☑', 'F0A1': '☐',
    }
    W14_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"

    # 预先收集需要跳过的子树根节点 id
    skip_subtree_ids: Set[int] = set()

    if skip_chart_drawing:
        for drawing in element.iter(f"{{{W_NS}}}drawing"):
            if drawing.find(f".//{{{C_NS}}}chart") is not None:
                skip_subtree_ids.add(id(drawing))

    if skip_textbox:
        # 收集所有文本框容器，后续跳过其内部所有节点
        for tag in _SKIP_CONTAINER_TAGS:
            for container in element.iter(tag):
                skip_subtree_ids.add(id(container))

    def _should_skip(node) -> bool:
        """判断节点是否在需要跳过的子树内部"""
        if not skip_subtree_ids:
            return False
        # 检查自身是否就是被标记的容器
        if id(node) in skip_subtree_ids:
            return True
        # 向上遍历检查祖先
        parent = node.getparent()
        while parent is not None:
            if id(parent) in skip_subtree_ids:
                return True
            parent = parent.getparent()
        return False

    texts: List[str] = []
    for node in element.iter():
        # --- A. 处理标准文本 w:t / m:t ---
        if node.tag in (f"{{{W_NS}}}t", f"{{{M_NS}}}t"):
            # 检查是否在文本框或图表子树内
            if _should_skip(node):
                continue
            if node.text:
                t = node.text
                for raw, char in SYMBOL_CHAR_MAP.items():
                    if len(raw) > 2:
                        t = t.replace(raw, char)
                texts.append(t)
            if node.tail:
                texts.append(node.tail)

        # --- B. DrawingML 文本 a:t ---
        elif node.tag == f"{{{A_NS}}}t":
            if _should_skip(node):
                continue
            if node.text:
                texts.append(node.text)
            if node.tail:
                texts.append(node.tail)

        # --- C. 图表/图片描述 ---
        elif node.tag == f"{{{WP_NS}}}docPr":
            if _should_skip(node):
                continue
            alt = node.get("descr") or node.get("title")
            if alt:
                texts.append(f"[图表描述: {alt}]")

        # --- D. Symbol 符号标签 ---
        elif node.tag == f"{{{W_NS}}}sym":
            if _should_skip(node):
                continue
            char_code = node.get(f"{{{W_NS}}}char", "").upper()
            if char_code in SYMBOL_CHAR_MAP:
                texts.append(SYMBOL_CHAR_MAP[char_code])
            elif char_code:
                try:
                    char_val = chr(int(char_code, 16))
                    texts.append(SYMBOL_CHAR_MAP.get(char_val, char_val))
                except (ValueError, OverflowError):
                    texts.append(f"[{char_code}]")

        # --- E. 数学特殊字符 ---
        elif node.tag == f"{{{M_NS}}}char":
            if _should_skip(node):
                continue
            val = node.get(f"{{{M_NS}}}val") or node.get(f"{{{W_NS}}}val")
            if val:
                texts.append(val)

        # --- F. 结构化控件 Checkbox ---
        elif node.tag.endswith("sdt"):
            if _should_skip(node):
                continue
            checkbox = node.find(f".//{{{W14_NS}}}checkbox")
            if checkbox is not None:
                checked = checkbox.find(f".//{{{W14_NS}}}checked")
                val = checked.get(f"{{{W14_NS}}}val") if checked is not None else "0"
                texts.append("☑" if val in ["1", "true"] else "☐")

    return "".join(texts).strip()


# ============================================================
#  图表(Chart)文字提取系统
# ============================================================
class ChartExtractor:
    """
    核心策略：
    1. 优先提取数据标签的实际显示内容（解析 [CELLRANGE] 为真实文字）
    2. 如果没有数据标签，才回退到系列名+分类名+数据值
    3. 始终提取：图表标题、轴标题
    """

    SEP = "\t"

    def __init__(self, doc_path: str):
        self.doc_path = doc_path
        self.rid_to_chart_path: Dict[str, str] = {}
        self.chart_texts: Dict[str, List[str]] = {}
        self._load_chart_rels()
        self._parse_all_charts()

    def _load_chart_rels(self) -> None:
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

    def _parse_all_charts(self) -> None:
        try:
            with ZipFile(self.doc_path, "r") as zf:
                for rid, chart_path in self.rid_to_chart_path.items():
                    if chart_path in zf.namelist():
                        with zf.open(chart_path) as f:
                            tree = etree.parse(f)
                            texts = self._extract_chart_texts(tree.getroot())
                            self.chart_texts[chart_path] = texts
                for name in zf.namelist():
                    if re.match(r"word/charts/chart\d*\.xml$", name) and name not in self.chart_texts:
                        with zf.open(name) as f:
                            tree = etree.parse(f)
                            texts = self._extract_chart_texts(tree.getroot())
                            self.chart_texts[name] = texts
        except Exception as e:
            print(f"解析图表文件失败: {e}")

    def _extract_datalabels_range_cache(self, ser_elem) -> List[str]:
        """从 c15:datalabelsRange/dlblRangeCache 中提取 [CELLRANGE] 对应的真实文字"""
        labels: List[str] = []

        # 方式1: 精确命名空间
        for dlbl_range in ser_elem.iter(f"{{{C15_NS}}}datalabelsRange"):
            cache = dlbl_range.find(f"{{{C15_NS}}}dlblRangeCache")
            if cache is not None:
                labels = self._extract_cache_values(cache)
                if labels:
                    return labels

        # 方式2: 通配符 localname 兜底
        for ext in ser_elem.iter(f"{{{C_NS}}}ext"):
            for child in ext.iter():
                tag_local = etree.QName(child.tag).localname if isinstance(child.tag, str) else ""
                if tag_local == "datalabelsRange":
                    for cache_child in child.iter():
                        cache_local = etree.QName(cache_child.tag).localname if isinstance(cache_child.tag, str) else ""
                        if cache_local == "dlblRangeCache":
                            labels = self._extract_cache_values(cache_child)
                            if labels:
                                return labels
        return labels

    def _extract_cache_values(self, cache_elem) -> List[str]:
        """从缓存元素中按 idx 顺序提取所有值"""
        pt_map: Dict[int, str] = {}
        for pt in cache_elem.iter():
            tag_local = etree.QName(pt.tag).localname if isinstance(pt.tag, str) else ""
            if tag_local == "pt":
                idx_str = pt.get("idx", "")
                if not idx_str:
                    idx_str = pt.get(f"{{{C_NS}}}idx", "")
                try:
                    idx = int(idx_str) if idx_str else -1
                except ValueError:
                    idx = -1
                v_elem = None
                for v_child in pt:
                    v_local = etree.QName(v_child.tag).localname if isinstance(v_child.tag, str) else ""
                    if v_local == "v":
                        v_elem = v_child
                        break
                if v_elem is not None and v_elem.text:
                    if idx >= 0:
                        pt_map[idx] = v_elem.text.strip()
                    else:
                        pt_map[len(pt_map)] = v_elem.text.strip()
        if not pt_map:
            return []
        max_idx = max(pt_map.keys())
        return [pt_map.get(i, "") for i in range(max_idx + 1)]

    def _extract_chart_texts(self, root) -> List[str]:
        texts: List[str] = []

        title = self._get_chart_title(root)
        if title:
            texts.append(f"[图表标题] {title}")

        axis_texts = self._get_axis_texts(root)
        texts.extend(axis_texts)

        series_texts = self._get_series_with_labels(root)
        texts.extend(series_texts)

        # 兜底
        fallback = self._get_all_drawingml_text(root)
        existing_set: Set[str] = set()
        for t in texts:
            for word in re.split(r'[\[\]\s]+', t):
                if word:
                    existing_set.add(word)
        for ft in fallback:
            if ft in existing_set or ft == "[CELLRANGE]":
                continue
            try:
                float(ft)
                continue
            except ValueError:
                pass
            texts.append(ft)
            existing_set.add(ft)

        return texts

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
                    series_lines = self._parse_series_smart(ser)
                    if series_lines:
                        results.extend(series_lines)
        return results

    def _parse_series_smart(self, ser_elem) -> List[str]:
        parts: List[str] = []
        sep = self.SEP

        # 系列名称
        tx = ser_elem.find(f"{{{C_NS}}}tx")
        if tx is not None:
            series_name = self._get_str_or_ref(tx)
            if series_name:
                parts.append(f"  [系列名] {series_name}")

        # 分类名称
        cat = ser_elem.find(f"{{{C_NS}}}cat")
        cat_values: List[str] = []
        if cat is not None:
            cat_values = self._get_values(cat)
            if cat_values:
                parts.append(f"  [分类] {sep.join(cat_values)}")

        # 数据标签优先
        range_labels = self._extract_datalabels_range_cache(ser_elem)
        inline_labels = self._extract_inline_data_labels(ser_elem)

        has_meaningful_labels = False

        if range_labels:
            non_empty = [lb for lb in range_labels if lb]
            if non_empty:
                has_meaningful_labels = True
                parts.append(f"  [数据标签] {sep.join(range_labels)}")

        if inline_labels and not has_meaningful_labels:
            meaningful = [lb for lb in inline_labels if lb and lb != "[CELLRANGE]"]
            if meaningful:
                has_meaningful_labels = True
                parts.append(f"  [数据标签] {sep.join(meaningful)}")

        # 没有数据标签时才回退到原始数值
        if not has_meaningful_labels:
            val = ser_elem.find(f"{{{C_NS}}}val")
            if val is not None:
                data_values = self._get_values(val)
                if data_values:
                    parts.append(f"  [数据] {sep.join(data_values)}")

            xval = ser_elem.find(f"{{{C_NS}}}xVal")
            if xval is not None:
                x_values = self._get_values(xval)
                if x_values:
                    parts.append(f"  [X值] {sep.join(x_values)}")

            yval = ser_elem.find(f"{{{C_NS}}}yVal")
            if yval is not None:
                y_values = self._get_values(yval)
                if y_values:
                    parts.append(f"  [Y值] {sep.join(y_values)}")

            bubble = ser_elem.find(f"{{{C_NS}}}bubbleSize")
            if bubble is not None:
                b_values = self._get_values(bubble)
                if b_values:
                    parts.append(f"  [气泡大小] {sep.join(b_values)}")

        return parts

    def _extract_inline_data_labels(self, ser_elem) -> List[str]:
        label_map: Dict[int, str] = {}
        for dlbl in ser_elem.iter(f"{{{C_NS}}}dLbl"):
            idx_elem = dlbl.find(f"{{{C_NS}}}idx")
            if idx_elem is None:
                continue
            try:
                idx = int(idx_elem.get("val", idx_elem.get(f"{{{C_NS}}}val", "-1")))
            except (ValueError, TypeError):
                continue
            label_parts: List[str] = []
            for at in dlbl.iter(f"{{{A_NS}}}t"):
                if at.text and at.text.strip():
                    label_parts.append(at.text.strip())
            if label_parts:
                label_map[idx] = "".join(label_parts)
        if not label_map:
            return []
        max_idx = max(label_map.keys())
        return [label_map.get(i, "") for i in range(max_idx + 1)]

    def _get_chart_title(self, root) -> Optional[str]:
        chart_elem = root.find(f"{{{C_NS}}}chart")
        if chart_elem is None:
            return None
        title_elem = chart_elem.find(f"{{{C_NS}}}title")
        if title_elem is None:
            return None
        tx = title_elem.find(f"{{{C_NS}}}tx")
        if tx is not None:
            rich = tx.find(f"{{{C_NS}}}rich")
            if rich is not None:
                parts: List[str] = []
                for para in rich.findall(f"{{{A_NS}}}p"):
                    para_parts: List[str] = []
                    for run in para.findall(f"{{{A_NS}}}r"):
                        for at in run.findall(f"{{{A_NS}}}t"):
                            if at.text:
                                para_parts.append(at.text)
                    if para_parts:
                        parts.append("".join(para_parts))
                if parts:
                    return "\n".join(parts).strip()
            str_ref = tx.find(f"{{{C_NS}}}strRef")
            if str_ref is not None:
                cache = str_ref.find(f"{{{C_NS}}}strCache")
                if cache is not None:
                    for pt in cache.findall(f"{{{C_NS}}}pt"):
                        v = pt.find(f"{{{C_NS}}}v")
                        if v is not None and v.text:
                            return v.text.strip()
        parts = []
        for at in title_elem.iter(f"{{{A_NS}}}t"):
            if at.text:
                parts.append(at.text)
        if parts:
            return "".join(parts).strip()
        return None

    def _get_axis_texts(self, root) -> List[str]:
        results: List[str] = []
        axis_tags = [
            (f"{{{C_NS}}}catAx", "分类轴"),
            (f"{{{C_NS}}}valAx", "值轴"),
            (f"{{{C_NS}}}dateAx", "日期轴"),
            (f"{{{C_NS}}}serAx", "系列轴"),
        ]
        for axis_tag, axis_name in axis_tags:
            for axis in root.iter(axis_tag):
                title_elem = axis.find(f"{{{C_NS}}}title")
                if title_elem is not None:
                    parts: List[str] = []
                    for at in title_elem.iter(f"{{{A_NS}}}t"):
                        if at.text:
                            parts.append(at.text.strip())
                    if parts:
                        results.append(f"[{axis_name}标题] {''.join(parts)}")
        return results

    def _get_str_or_ref(self, elem) -> str:
        str_val = elem.find(f"{{{C_NS}}}v")
        if str_val is not None and str_val.text:
            return str_val.text.strip()
        str_ref = elem.find(f"{{{C_NS}}}strRef")
        if str_ref is not None:
            cache = str_ref.find(f"{{{C_NS}}}strCache")
            if cache is not None:
                for pt in cache.findall(f"{{{C_NS}}}pt"):
                    v = pt.find(f"{{{C_NS}}}v")
                    if v is not None and v.text:
                        return v.text.strip()
        parts: List[str] = []
        for at in elem.iter(f"{{{A_NS}}}t"):
            if at.text:
                parts.append(at.text.strip())
        if parts:
            return "".join(parts)
        return ""

    def _get_values(self, elem) -> List[str]:
        values: List[str] = []
        for cache in elem.iter(f"{{{C_NS}}}strCache"):
            for pt in cache.findall(f"{{{C_NS}}}pt"):
                v = pt.find(f"{{{C_NS}}}v")
                if v is not None and v.text:
                    values.append(v.text.strip())
        if not values:
            for cache in elem.iter(f"{{{C_NS}}}numCache"):
                for pt in cache.findall(f"{{{C_NS}}}pt"):
                    v = pt.find(f"{{{C_NS}}}v")
                    if v is not None and v.text:
                        values.append(v.text.strip())
        if not values:
            for cache in elem.iter(f"{{{C_NS}}}multiLvlStrCache"):
                for lvl in cache.findall(f"{{{C_NS}}}lvl"):
                    for pt in lvl.findall(f"{{{C_NS}}}pt"):
                        v = pt.find(f"{{{C_NS}}}v")
                        if v is not None and v.text:
                            values.append(v.text.strip())
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

    def get_chart_text_by_rid(self, rid: str) -> List[str]:
        chart_path = self.rid_to_chart_path.get(rid, "")
        return self.chart_texts.get(chart_path, [])

    def get_all_chart_texts(self) -> Dict[str, List[str]]:
        return self.chart_texts


# ---------------- 页眉页脚 ----------------
def _extract_header_footer(doc_path: str) -> Tuple[List[str], List[str]]:
    headers: List[str] = []
    footers: List[str] = []
    try:
        with ZipFile(doc_path, "r") as zf:
            for name in zf.namelist():
                if name.startswith("word/header"):
                    with zf.open(name) as f:
                        tree = etree.parse(f)
                        text = _get_xml_text(tree.getroot(), skip_chart_drawing=False,
                                             skip_textbox=False)
                        if text.strip():
                            headers.append(text)
                elif name.startswith("word/footer"):
                    with zf.open(name) as f:
                        tree = etree.parse(f)
                        text = _get_xml_text(tree.getroot(), skip_chart_drawing=False,
                                             skip_textbox=False)
                        if text.strip():
                            footers.append(text)
    except Exception as e:
        print(f"提取页眉页脚失败: {e}")
    return headers, footers


# ---------------- 锚点加载（脚注/尾注/批注） ----------------
class DocAnchorsLoader:
    def __init__(self, doc_path: str):
        self.doc_path = doc_path
        self.footnotes: Dict[str, str] = {}
        self.endnotes: Dict[str, str] = {}
        self.comments: Dict[str, str] = {}
        self._load_all()

    def _load_xml_map(self, zf: ZipFile, filename: str, tag: str) -> Dict[str, str]:
        data: Dict[str, str] = {}
        if filename not in zf.namelist():
            return data
        with zf.open(filename) as f:
            tree = etree.parse(f)
            for elem in tree.findall(f".//w:{tag}", NAMESPACES):
                eid = elem.get(f"{{{W_NS}}}id")
                elem_type = elem.get(f"{{{W_NS}}}type")
                if elem_type in ("separator", "continuationSeparator"):
                    continue
                text = _get_xml_text(elem, skip_chart_drawing=False, skip_textbox=False)
                if eid and text:
                    data[eid] = text
        return data

    def _load_all(self) -> None:
        try:
            with ZipFile(self.doc_path, "r") as zf:
                self.footnotes = self._load_xml_map(zf, "word/footnotes.xml", "footnote")
                self.endnotes = self._load_xml_map(zf, "word/endnotes.xml", "endnote")
                self.comments = self._load_xml_map(zf, "word/comments.xml", "comment")
        except Exception as e:
            print(f"加载锚点内容失败: {e}")


# ---------------- 锚点插入（文本框提取） ----------------
def _process_anchored_content(p_element, loader: DocAnchorsLoader) -> List[str]:
    """
    提取段落中的锚定内容：文本框、脚注、尾注、批注。
    文本框去重策略：只提取 wps:txbxContent（最内层），跳过 v:textbox 的重复。
    """
    extras: List[str] = []

    # 收集所有 wps:txbxContent 的 id，用于后续 v:textbox 去重
    txbx_content_ids: Set[int] = set()

    # 优先提取 wps:txbxContent（新版文本框格式）
    for txbx in p_element.iter(f"{{{WPS_NS}}}txbxContent"):
        txbx_content_ids.add(id(txbx))
        # 对文本框内部不再跳过文本框（因为这里就是文本框本身）
        t = _get_xml_text(txbx, skip_chart_drawing=False, skip_textbox=False)
        if t:
            extras.append(t)

    # 再提取 v:textbox（旧版文本框格式），但跳过已被 wps:txbxContent 覆盖的
    for vtextbox in p_element.iter(f"{{{V_NS}}}textbox"):
        # 检查这个 v:textbox 内部是否包含已处理过的 wps:txbxContent
        has_processed_child = False
        for child_txbx in vtextbox.iter(f"{{{WPS_NS}}}txbxContent"):
            if id(child_txbx) in txbx_content_ids:
                has_processed_child = True
                break

        if has_processed_child:
            continue  # 已经通过 wps:txbxContent 提取过了，跳过

        # 纯 v:textbox（没有内嵌 wps:txbxContent 的旧格式）
        t = _get_xml_text(vtextbox, skip_chart_drawing=False, skip_textbox=False)
        if t:
            extras.append(t)

    # 脚注
    for ref in p_element.findall(".//w:footnoteReference", NAMESPACES):
        fid = ref.get(f"{{{W_NS}}}id")
        if fid and fid in loader.footnotes:
            extras.append(loader.footnotes[fid])

    # 尾注
    for ref in p_element.findall(".//w:endnoteReference", NAMESPACES):
        eid = ref.get(f"{{{W_NS}}}id")
        if eid and eid in loader.endnotes:
            extras.append(loader.endnotes[eid])

    # 批注
    for ref in p_element.findall(".//w:commentReference", NAMESPACES):
        cid = ref.get(f"{{{W_NS}}}id")
        if cid and cid in loader.comments:
            extras.append(loader.comments[cid])

    return extras


# ---------------- 图表引用提取 ----------------
def _extract_chart_rids_from_paragraph(p_element) -> List[str]:
    rids: List[str] = []
    for chart_ref in p_element.iter(f"{{{C_NS}}}chart"):
        rid = chart_ref.get(f"{{{R_NS}}}id")
        if rid:
            rids.append(rid)
    return rids


# ---------------- 编号系统 ----------------
class NumberingSystem:
    """处理 Word 自动编号系统（含 lvlOverride 和样式编号回退）"""

    def __init__(self, doc_path: str):
        self.doc_path = doc_path
        self.numbering_map: Dict[str, Dict[str, Dict]] = {}
        self.abstract_num_map: Dict[str, Dict[str, Dict]] = {}
        self.level_counters: Dict[Tuple[str, str], int] = {}
        self.style_num_map: Dict[str, Tuple[str, str]] = {}
        self._load_numbering()
        self._load_style_numbering()

    def _load_numbering(self) -> None:
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

    def _load_style_numbering(self) -> None:
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
        match fmt:
            case "decimal":
                return str(num)
            case "upperRoman":
                return self._to_roman(num).upper()
            case "lowerRoman":
                return self._to_roman(num).lower()
            case "upperLetter":
                return self._to_letter(num).upper()
            case "lowerLetter":
                return self._to_letter(num).lower()
            case "chineseCountingThousand" | "chineseCounting" | "ideographTraditional":
                return self._to_chinese(num)
            case "japaneseCounting" | "japaneseDigitalTenThousand":
                return self._to_chinese(num)
            case "bullet":
                return "•"
            case _:
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
        for other_ilvl_str in self.numbering_map[num_id]:
            if int(other_ilvl_str) > ilvl_int:
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


# ---------------- 主函数 ----------------
def extract_body_text(doc_path: str) -> str:
    if not os.path.exists(doc_path):
        raise FileNotFoundError(f"文件不存在: {doc_path}")

    headers, footers = _extract_header_footer(doc_path)
    loader = DocAnchorsLoader(doc_path)
    numbering = NumberingSystem(doc_path)
    chart_extractor = ChartExtractor(doc_path)

    doc = Document(doc_path)
    body = doc.element.body

    body_lines: List[str] = []
    inserted_chart_paths: Set[str] = set()

    def _process_paragraph(p_elem) -> List[str]:
        lines: List[str] = []
        num = numbering.get_paragraph_number(p_elem)
        # 正文段落：跳过图表和文本框内部节点（由专门函数处理）
        text = _get_xml_text(p_elem, skip_chart_drawing=True, skip_textbox=True)
        extras = _process_anchored_content(p_elem, loader)

        if num and text:
            lines.append(f"{num} {text}")
        elif text:
            lines.append(text)

        for e in extras:
            lines.append(e)

        # 图表
        chart_rids = _extract_chart_rids_from_paragraph(p_elem)
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

    for child in body.iterchildren():
        if child.tag.endswith("p"):
            body_lines.extend(_process_paragraph(child))
        elif child.tag.endswith("tbl"):
            for row in child.iter(f"{{{W_NS}}}tr"):
                row_text: List[str] = []
                for cell in row.iter(f"{{{W_NS}}}tc"):
                    cell_content: List[str] = []
                    for p in cell.iter(f"{{{W_NS}}}p"):
                        cell_content.extend(_process_paragraph(p))
                    row_text.append("\t".join(cell_content))
                body_lines.append("\t".join(row_text))

    # 兜底：未被段落引用的图表
    all_chart_texts = chart_extractor.get_all_chart_texts()
    for chart_path, texts in all_chart_texts.items():
        if chart_path not in inserted_chart_paths and texts:
            body_lines.append("[未关联图表内容开始]")
            body_lines.extend(texts)
            body_lines.append("[未关联图表内容结束]")

    result: List[str] = []
    if headers:
        result.append("[HEADER]")
        result.extend(headers)
    result.append("\n[BODY]")
    result.extend(body_lines)
    if footers:
        result.append("\n[FOOTER]")
        result.extend(footers)

    return "\n".join(result)


# ---------------- 测试 ----------------
if __name__ == "__main__":
    path = r"C:\Users\H\Desktop\测试项目\中翻译\测试文件\雅本化学2025ESG报告文字稿-20260409.docx"
    try:
        result = extract_body_text(path)
        print(result)
    except FileNotFoundError as e:
        print(f"错误: {e}")
    except Exception as e:
        print(f"提取失败: {e}")