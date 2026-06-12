from zipfile import ZipFile
from lxml import etree
from pathlib import Path
from dataclasses import dataclass, field


# ──────────────────── 命名空间 ────────────────────
NS: dict[str, str] = {
    "w":   "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc":  "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "wp":  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c":   "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "v":   "urn:schemas-microsoft-com:vml",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

# Relationship 类型常量
REL_HEADER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
REL_FOOTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
REL_FOOTNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
REL_ENDNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
REL_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"


@dataclass
class TextSegment:
    """一个文本片段，带有来源标记"""
    source: str        # 来源标识: body / header / footer / footnote / endnote / chart / textbox
    text: str          # 文本内容
    xml_path: str = "" # 所在 ZIP 内的 XML 路径，用于写回
    para_index: int = -1  # 在译文 segment 序列中的顺序位置（模式B时用于精确定位）
    row_context: str = "" # 表格行上下文：同行所有单元格 tab 拼接（表格单元格时填充）


def _qn(tag: str) -> str:
    """
    将 'w:p' 这样的前缀标签转换为 Clark notation '{namespace}p'。
    方便与 element.tag 直接比较。
    """
    prefix, local = tag.split(":")
    return f"{{{NS[prefix]}}}{local}"


def _parse_relationships(zf: ZipFile, rels_path: str) -> dict[str, str]:
    """
    解析 .rels 文件，返回 {rId: target_path} 映射。
    """
    if rels_path not in zf.namelist():
        return {}
    with zf.open(rels_path) as f:
        tree = etree.parse(f)
    result: dict[str, str] = {}
    for rel in tree.getroot():
        rid = rel.get("Id", "")
        target = rel.get("Target", "")
        result[rid] = target
    return result


def _parse_rels_by_type(zf: ZipFile, rels_path: str) -> dict[str, list[tuple[str, str]]]:
    """
    解析 .rels 文件，返回 {type: [(rId, target_path), ...]} 映射。
    """
    if rels_path not in zf.namelist():
        return {}
    with zf.open(rels_path) as f:
        tree = etree.parse(f)
    result: dict[str, list[tuple[str, str]]] = {}
    for rel in tree.getroot():
        rtype = rel.get("Type", "")
        rid = rel.get("Id", "")
        target = rel.get("Target", "")
        result.setdefault(rtype, []).append((rid, target))
    return result


def _extract_paragraph_text(para_elem: etree._Element) -> str:
    """
    从一个 <w:p> 段落元素中提取拼接后的纯文本。
    跳过浮动文本框（<wps:txbx>）中的内容，避免文本框锚定位置不同
    导致原文/译文提取顺序错位。
    """
    # 先收集所有文本框内的 <w:t> 节点，提取时跳过
    txbx_wt_ids: set[int] = set()
    for txbx in para_elem.iter(_qn("wps:txbx")):
        for wt in txbx.iter(_qn("w:t")):
            txbx_wt_ids.add(id(wt))

    parts: list[str] = []
    for wt in para_elem.iter(_qn("w:t")):
        if id(wt) not in txbx_wt_ids and wt.text:
            parts.append(wt.text)

    return "".join(parts)


def _extract_xml_paragraphs(zf: ZipFile, xml_path: str) -> list[str]:
    """
    从指定 XML 文件中按段落顺序提取文本。
    """
    if xml_path not in zf.namelist():
        return []
    with zf.open(xml_path) as f:
        tree = etree.parse(f)
    paragraphs: list[str] = []
    for para in tree.iter(_qn("w:p")):
        text = _extract_paragraph_text(para)
        if text.strip():  # 跳过空段落
            paragraphs.append(text)
    return paragraphs


def _extract_chart_text(zf: ZipFile, chart_path: str) -> list[str]:
    """
    从图表 XML 中提取需要翻译的文字内容：
    - <a:t> 坐标轴标题、图例等（跳过 <a:fld> 内的占位符如 [CELLRANGE]）
    - <c:v> 中的字符串值（跳过纯数字，那些是数据值不需要翻译）
    去重并保持顺序。
    """
    if chart_path not in zf.namelist():
        return []
    with zf.open(chart_path) as f:
        tree = etree.parse(f)

    seen: set[str] = set()
    texts: list[str] = []

    def add(t: str) -> None:
        t = t.strip()
        if t and t not in seen:
            seen.add(t)
            texts.append(t)

    # a:t（跳过 a:fld 内的占位符）
    fld_tag = _qn("a:fld")
    for at in tree.iter(_qn("a:t")):
        if at.getparent() is not None and at.getparent().tag == fld_tag:
            continue
        if at.text:
            add(at.text)

    # c:v：只取非纯数字的字符串（议题名、系列名、分类名等）
    for cv in tree.iter(_qn("c:v")):
        if cv.text and cv.text.strip():
            val = cv.text.strip()
            # 跳过纯数字（整数或小数）
            try:
                float(val)
            except ValueError:
                add(val)

    return texts


def _load_footnotes(zf: ZipFile) -> dict[str, str]:
    """
    加载脚注，返回 {footnote_id: text} 映射。
    id=0 (分隔符) 和 id=1 (续延分隔符) 通常是系统占位，会被过滤。
    """
    path = "word/footnotes.xml"
    if path not in zf.namelist():
        return {}
    with zf.open(path) as f:
        tree = etree.parse(f)
    notes: dict[str, str] = {}
    for fn in tree.iter(_qn("w:footnote")):
        fid = fn.get(_qn("w:id"), "")
        if fid in ("0", "1", "-1"):
            continue  # 跳过系统占位脚注
        parts = []
        for wt in fn.iter(_qn("w:t")):
            if wt.text:
                parts.append(wt.text)
        text = "".join(parts).strip()
        if text:
            notes[fid] = text
    return notes


def _load_endnotes(zf: ZipFile) -> dict[str, str]:
    """
    加载尾注，返回 {endnote_id: text} 映射。
    """
    path = "word/endnotes.xml"
    if path not in zf.namelist():
        return {}
    with zf.open(path) as f:
        tree = etree.parse(f)
    notes: dict[str, str] = {}
    for en in tree.iter(_qn("w:endnote")):
        eid = en.get(_qn("w:id"), "")
        if eid in ("0", "1", "-1"):
            continue
        parts = []
        for wt in en.iter(_qn("w:t")):
            if wt.text:
                parts.append(wt.text)
        text = "".join(parts).strip()
        if text:
            notes[eid] = text
    return notes


def _get_section_header_footer_rids(sect_pr: etree._Element) -> tuple[list[str], list[str]]:
    """
    从 <w:sectPr> 中提取页眉和页脚的 r:id 列表。
    """
    header_rids: list[str] = []
    footer_rids: list[str] = []
    for child in sect_pr:
        tag = child.tag
        rid = child.get(_qn("r:id"), "")
        if tag == _qn("w:headerReference") and rid:
            header_rids.append(rid)
        elif tag == _qn("w:footerReference") and rid:
            footer_rids.append(rid)
    return header_rids, footer_rids


def extract_docx_in_order(docx_path: str | Path) -> list[TextSegment]:
    """
    按原文阅读顺序提取 .docx 全部文本内容。

    顺序逻辑：
      - 每个"节"(section) 开始时，先输出该节的页眉
      - 然后逐段落输出正文，遇到脚注/尾注引用时内联插入
      - 遇到图表引用时内联插入图表文本
      - 节结束时，输出该节的页脚

    Args:
        docx_path: .docx 文件路径

    Returns:
        按阅读顺序排列的 TextSegment 列表
    """
    docx_path = Path(docx_path)

    if not docx_path.exists():
        raise FileNotFoundError(f"文件不存在: {docx_path}")
    if docx_path.suffix.lower() != ".docx":
        raise ValueError(f"不是 .docx 文件: {docx_path}")

    segments: list[TextSegment] = []

    try:
        with ZipFile(docx_path) as zf:
            # ── 加载 relationship 映射 ──
            rels = _parse_relationships(zf, "word/_rels/document.xml.rels")

            # ── 预加载脚注和尾注 ──
            footnotes = _load_footnotes(zf)
            endnotes = _load_endnotes(zf)

            # ── 解析主文档 ──
            with zf.open("word/document.xml") as f:
                tree = etree.parse(f)

            body = tree.find(_qn("w:body"))
            if body is None:
                return segments

            # 收集所有"节"的信息
            # Word 文档的节由 <w:sectPr> 分隔：
            #   - 段落内的 <w:pPr>/<w:sectPr> 表示该段落是一个节的最后一段
            #   - <w:body> 末尾的 <w:sectPr> 是最后一个节的属性
            #
            # 策略：先遍历所有 body 的直接子元素，分成多个"节"，
            #        然后按节依次输出：页眉 → 正文段落 → 页脚

            sections: list[tuple[list[etree._Element], etree._Element | None]] = []
            current_section_elements: list[etree._Element] = []

            for child in body:
                tag = child.tag

                if tag == _qn("w:sectPr"):
                    # body 末尾的 sectPr（最后一节的属性）
                    sections.append((current_section_elements, child))
                    current_section_elements = []
                elif tag == _qn("w:p"):
                    # 检查段落内是否有 sectPr（节分隔）
                    ppr = child.find(_qn("w:pPr"))
                    sect_in_para = None
                    if ppr is not None:
                        sect_in_para = ppr.find(_qn("w:sectPr"))

                    current_section_elements.append(child)

                    if sect_in_para is not None:
                        # 这个段落是当前节的最后一段
                        sections.append((current_section_elements, sect_in_para))
                        current_section_elements = []
                elif tag == _qn("w:tbl"):
                    # 表格也是 body 的直接子元素
                    current_section_elements.append(child)
                else:
                    current_section_elements.append(child)

            # 如果还有剩余元素没有 sectPr（不太常见，但做兜底）
            if current_section_elements:
                sections.append((current_section_elements, None))

            # ── 按节顺序输出 ──
            _para_idx = 0  # 全局顺序计数器，贯穿所有节

            for sect_elements, sect_pr in sections:

                # 1) 输出页眉
                if sect_pr is not None:
                    header_rids, footer_rids = _get_section_header_footer_rids(sect_pr)
                    for rid in header_rids:
                        target = rels.get(rid, "")
                        if target:
                            header_path = f"word/{target}"
                            for text in _extract_xml_paragraphs(zf, header_path):
                                segments.append(TextSegment(source="header", text=text,
                                                            xml_path=header_path, para_index=_para_idx))
                                _para_idx += 1
                else:
                    header_rids, footer_rids = [], []

                # 2) 输出正文段落（含内联脚注/尾注/图表）
                for elem in sect_elements:
                    if elem.tag == _qn("w:p"):
                        # 提取段落正文
                        para_text = _extract_paragraph_text(elem)
                        if para_text.strip():
                            segments.append(TextSegment(source="body", text=para_text,
                                                        xml_path="word/document.xml", para_index=_para_idx))
                            _para_idx += 1

                        # 检查段落中的脚注引用，内联插入脚注内容
                        for fn_ref in elem.iter(_qn("w:footnoteReference")):
                            fid = fn_ref.get(_qn("w:id"), "")
                            if fid in footnotes:
                                segments.append(TextSegment(
                                    source="footnote",
                                    text=f"[脚注{fid}] {footnotes[fid]}",
                                    xml_path="word/footnotes.xml",
                                    para_index=_para_idx,
                                ))
                                _para_idx += 1

                        # 检查尾注引用
                        for en_ref in elem.iter(_qn("w:endnoteReference")):
                            eid = en_ref.get(_qn("w:id"), "")
                            if eid in endnotes:
                                segments.append(TextSegment(
                                    source="endnote",
                                    text=f"[尾注{eid}] {endnotes[eid]}",
                                    xml_path="word/endnotes.xml",
                                    para_index=_para_idx,
                                ))
                                _para_idx += 1

                        # 检查图表引用（嵌在 drawing 中）
                        for chart_ref in elem.iter(_qn("c:chart")):
                            chart_rid = chart_ref.get(_qn("r:id"), "")
                            if chart_rid and chart_rid in rels:
                                chart_target = rels[chart_rid]
                                chart_path = f"word/{chart_target}"
                                chart_texts = _extract_chart_text(zf, chart_path)
                                for ct in chart_texts:
                                    segments.append(TextSegment(source="chart", text=ct,
                                                                xml_path=chart_path, para_index=_para_idx))
                                    _para_idx += 1

                    elif elem.tag == _qn("w:tbl"):
                        # 表格：先按行收集所有单元格文本，构造行上下文，再逐单元格输出
                        for row in elem.iter(_qn("w:tr")):
                            # 收集整行所有非空单元格文本，用于构造 row_context
                            row_cell_texts: list[str] = []
                            for cell in row.iter(_qn("w:tc")):
                                cell_parts: list[str] = []
                                for wt in cell.iter(_qn("w:t")):
                                    if wt.text:
                                        cell_parts.append(wt.text)
                                row_cell_texts.append("".join(cell_parts))

                            row_context = "\t".join(t for t in row_cell_texts if t.strip())

                            for cell_text in row_cell_texts:
                                if cell_text.strip():
                                    segments.append(TextSegment(
                                        source="table",
                                        text=cell_text,
                                        xml_path="word/document.xml",
                                        para_index=_para_idx,
                                        row_context=row_context,
                                    ))
                                    _para_idx += 1

                # 3) 输出页脚
                if sect_pr is not None:
                    for rid in footer_rids:
                        target = rels.get(rid, "")
                        if target:
                            footer_path = f"word/{target}"
                            for text in _extract_xml_paragraphs(zf, footer_path):
                                segments.append(TextSegment(source="footer", text=text,
                                                            xml_path=footer_path, para_index=_para_idx))
                                _para_idx += 1

    except KeyError as e:
        raise ValueError(f"文件结构异常，缺少关键文件: {e}")
    except etree.XMLSyntaxError as e:
        raise ValueError(f"XML 解析失败: {e}")

    return segments


# ──────────────────── 来源标签的显示样式 ────────────────────
SOURCE_LABELS: dict[str, str] = {
    "body":     "正文",
    "header":   "页眉",
    "footer":   "页脚",
    "footnote": "脚注",
    "endnote":  "尾注",
    "chart":    "图表",
    "table":    "表格",
    "textbox":  "文本框",
}


if __name__ == "__main__":
    path = r"C:\Users\H\Desktop\测试项目\中翻译\测试文件\雅本化学2025ESG报告文字稿-20260409.docx"

    try:
        segments = extract_docx_in_order(path)
        print(len(segments))
        print("=" * 70)
        print(f"  📖 按原文顺序提取完成，共 {len(segments)} 个文本片段")
        print("=" * 70)

        # 统计各来源数量
        from collections import Counter
        source_counts = Counter(seg.source for seg in segments)
        for src, label in SOURCE_LABELS.items():
            count = source_counts.get(src, 0)
            if count > 0:
                print(f"    {label}: {count} 条")
        print("=" * 70)

        # 按顺序逐条输出，标注来源
        for i, seg in enumerate(segments, start=1):
            label = SOURCE_LABELS.get(seg.source, seg.source)
            print(f"[{i:>5d}] 【{label}】{seg.text}")

        # 纯文本拼接预览
        full_text = "\n".join(seg.text for seg in segments)
        print(f"\n{'=' * 70}")
        print(f"  📝 纯文本共 {len(full_text)} 字符，预览前 800 字符：")
        print(f"{'=' * 70}")
        # print(full_text[:800])

    except (FileNotFoundError, ValueError) as e:
        print(f"❌ 错误: {e}")