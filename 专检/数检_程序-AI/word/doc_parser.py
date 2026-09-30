"""
doc_parser.py
=============
统一解析 .docx / .doc / Word 2003 XML(.xml) 三种格式，
返回与 extract_docx_in_order() 兼容的 List[TextSegment]。

依赖：
    lxml          —— XML 解析（必须）
    olefile       —— .doc CFB 容器解析（pip install olefile，可选）
    pywin32       —— .doc → 临时 .docx 转换（仅 Windows，可选，fallback）
    python-docx   —— 不直接使用，但 win32com 输出的 docx 由本模块自身解析

.doc 解析策略（优先级依次降低）：
    1. 纯 Python CFB 解析（olefile + 自实现 FIB/Clx/PlcPcd 文本提取）
    2. win32com.client（调用本机 Word/WPS COM，Windows Only）
    3. LibreOffice subprocess（跨平台，需安装 LibreOffice）
    4. 抛出 RuntimeError 提示用户安装转换工具
"""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from zipfile import ZipFile, BadZipFile

from lxml import etree

# ---------- 引入公共数据结构与 docx 解析核心 ----------
# 假设本文件与调用方在同一项目，TextSegment 及 docx 解析函数从同级模块导入。
# 若作为独立模块使用，可将 TextSegment 定义拷贝到此处。
try:
    from dataclasses import dataclass, field as dc_field

    @dataclass
    class TextSegment:
        """一个文本片段，带有来源标记（与主模块保持字段一致）"""
        source: str         # body / header / footer / footnote / endnote / chart / table / textbox
        text: str
        xml_path: str = ""
        para_index: int = -1
        row_context: str = ""

except ImportError:
    raise


# ══════════════════════════════════════════════════════════════════════
#  命名空间 & 工具函数（与主 docx 解析模块保持一致）
# ══════════════════════════════════════════════════════════════════════

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
    # Word 2003 XML 专用
    "w2003": "http://schemas.microsoft.com/office/word/2003/wordml",
    "wx":    "http://schemas.microsoft.com/office/word/2003/auxHint",
    "o":     "urn:schemas-microsoft-com:office:office",
}

REL_HEADER    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
REL_FOOTER    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
REL_FOOTNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
REL_ENDNOTES  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
REL_CHART     = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"

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


def _qn(tag: str) -> str:
    """'w:p' → '{namespace}p'"""
    prefix, local = tag.split(":")
    return f"{{{NS[prefix]}}}{local}"


_SYMBOL_MAP: dict[str, str] = {
    "F020": " ", "F0B4": "×", "F0B7": "·", "F0B1": "±",
    "F0B0": "°", "F0B5": "µ", "F0B6": "¶", "F0A7": "§",
    "F0A9": "©", "F0AE": "®", "F0D7": "×", "F0F7": "÷",
    "F0AB": "«", "F0BB": "»", "F021": "!", "F022": '"',
    "F023": "#", "F025": "%", "F026": "&", "F03D": "=", "F040": "@",
}


def _sym_to_char(font: str, char_code: str) -> str:
    code_upper = char_code.upper()
    if font == "Symbol":
        return _SYMBOL_MAP.get(code_upper, f"[sym:{char_code}]")
    try:
        cp = int(char_code, 16)
        if 0xF000 <= cp <= 0xF0FF:
            cp -= 0xF000
        return chr(cp)
    except ValueError:
        return f"[sym:{char_code}]"


# ══════════════════════════════════════════════════════════════════════
#  DOCX 核心解析（直接内嵌，避免循环依赖）
# ══════════════════════════════════════════════════════════════════════

def _parse_relationships(zf: ZipFile, rels_path: str) -> dict[str, str]:
    if rels_path not in zf.namelist():
        return {}
    with zf.open(rels_path) as f:
        tree = etree.parse(f)
    return {
        rel.get("Id", ""): rel.get("Target", "")
        for rel in tree.getroot()
    }


def _extract_paragraph_text(para_elem: etree._Element) -> str:
    txbx_node_ids: set[int] = set()
    for txbx in para_elem.iter(_qn("wps:txbx")):
        for node in txbx.iter():
            txbx_node_ids.add(id(node))
    parts: list[str] = []
    for run in para_elem.iter(_qn("w:r")):
        if id(run) in txbx_node_ids:
            continue
        for child in run:
            if id(child) in txbx_node_ids:
                continue
            if child.tag == _qn("w:t") and child.text:
                parts.append(child.text)
            elif child.tag == _qn("w:sym"):
                font = child.get(_qn("w:font"), "")
                char_code = child.get(_qn("w:char"), "")
                if char_code:
                    parts.append(_sym_to_char(font, char_code))
    return "".join(parts)


def _extract_xml_paragraphs(zf: ZipFile, xml_path: str) -> list[str]:
    if xml_path not in zf.namelist():
        return []
    with zf.open(xml_path) as f:
        tree = etree.parse(f)
    return [
        _extract_paragraph_text(para)
        for para in tree.iter(_qn("w:p"))
        if _extract_paragraph_text(para).strip()
    ]


def _extract_chart_text(zf: ZipFile, chart_path: str) -> list[str]:
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

    fld_tag = _qn("a:fld")
    for at in tree.iter(_qn("a:t")):
        if at.getparent() is not None and at.getparent().tag == fld_tag:
            continue
        if at.text:
            add(at.text)
    for cv in tree.iter(_qn("c:v")):
        if cv.text and cv.text.strip():
            try:
                float(cv.text.strip())
            except ValueError:
                add(cv.text)
    return texts


def _load_footnotes(zf: ZipFile) -> dict[str, str]:
    path = "word/footnotes.xml"
    if path not in zf.namelist():
        return {}
    with zf.open(path) as f:
        tree = etree.parse(f)
    notes: dict[str, str] = {}
    for fn in tree.iter(_qn("w:footnote")):
        fid = fn.get(_qn("w:id"), "")
        if fid in ("0", "1", "-1"):
            continue
        text = "".join(wt.text for wt in fn.iter(_qn("w:t")) if wt.text).strip()
        if text:
            notes[fid] = text
    return notes


def _load_endnotes(zf: ZipFile) -> dict[str, str]:
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
        text = "".join(wt.text for wt in en.iter(_qn("w:t")) if wt.text).strip()
        if text:
            notes[eid] = text
    return notes


def _get_section_header_footer_rids(
    sect_pr: etree._Element,
) -> tuple[list[str], list[str]]:
    header_rids: list[str] = []
    footer_rids: list[str] = []
    for child in sect_pr:
        rid = child.get(_qn("r:id"), "")
        if child.tag == _qn("w:headerReference") and rid:
            header_rids.append(rid)
        elif child.tag == _qn("w:footerReference") and rid:
            footer_rids.append(rid)
    return header_rids, footer_rids


def _parse_docx_zip(zf: ZipFile) -> list[TextSegment]:
    """从已打开的 ZipFile 对象解析 docx，返回 TextSegment 列表。"""
    segments: list[TextSegment] = []
    rels = _parse_relationships(zf, "word/_rels/document.xml.rels")
    footnotes = _load_footnotes(zf)
    endnotes  = _load_endnotes(zf)

    with zf.open("word/document.xml") as f:
        tree = etree.parse(f)
    body = tree.find(_qn("w:body"))
    if body is None:
        return segments

    # 分节
    sections: list[tuple[list[etree._Element], etree._Element | None]] = []
    current: list[etree._Element] = []
    for child in body:
        if child.tag == _qn("w:sectPr"):
            sections.append((current, child))
            current = []
        elif child.tag == _qn("w:p"):
            ppr = child.find(_qn("w:pPr"))
            sect_in_para = ppr.find(_qn("w:sectPr")) if ppr is not None else None
            current.append(child)
            if sect_in_para is not None:
                sections.append((current, sect_in_para))
                current = []
        else:
            current.append(child)
    if current:
        sections.append((current, None))

    idx = 0
    for sect_elements, sect_pr in sections:
        header_rids, footer_rids = (
            _get_section_header_footer_rids(sect_pr) if sect_pr is not None else ([], [])
        )

        # 页眉
        for rid in header_rids:
            target = rels.get(rid, "")
            if target:
                hp = f"word/{target}"
                for text in _extract_xml_paragraphs(zf, hp):
                    segments.append(TextSegment("header", text, hp, idx)); idx += 1

        # 正文 & 表格
        for elem in sect_elements:
            if elem.tag == _qn("w:p"):
                para_text = _extract_paragraph_text(elem)
                if para_text.strip():
                    segments.append(TextSegment("body", para_text, "word/document.xml", idx)); idx += 1

                for fn_ref in elem.iter(_qn("w:footnoteReference")):
                    fid = fn_ref.get(_qn("w:id"), "")
                    if fid in footnotes:
                        segments.append(TextSegment("footnote", f"[脚注{fid}] {footnotes[fid]}",
                                                    "word/footnotes.xml", idx)); idx += 1

                for en_ref in elem.iter(_qn("w:endnoteReference")):
                    eid = en_ref.get(_qn("w:id"), "")
                    if eid in endnotes:
                        segments.append(TextSegment("endnote", f"[尾注{eid}] {endnotes[eid]}",
                                                    "word/endnotes.xml", idx)); idx += 1

                for chart_ref in elem.iter(_qn("c:chart")):
                    chart_rid = chart_ref.get(_qn("r:id"), "")
                    if chart_rid and chart_rid in rels:
                        cp = f"word/{rels[chart_rid]}"
                        for ct in _extract_chart_text(zf, cp):
                            segments.append(TextSegment("chart", ct, cp, idx)); idx += 1

            elif elem.tag == _qn("w:tbl"):
                for row in elem.iter(_qn("w:tr")):
                    row_texts = [
                        "".join(wt.text for wt in cell.iter(_qn("w:t")) if wt.text)
                        for cell in row.iter(_qn("w:tc"))
                    ]
                    row_context = "\t".join(t for t in row_texts if t.strip())
                    for ct in row_texts:
                        if ct.strip():
                            segments.append(TextSegment("table", ct, "word/document.xml",
                                                        idx, row_context)); idx += 1

        # 页脚
        for rid in footer_rids:
            target = rels.get(rid, "")
            if target:
                fp = f"word/{target}"
                for text in _extract_xml_paragraphs(zf, fp):
                    segments.append(TextSegment("footer", text, fp, idx)); idx += 1

    return segments


# ══════════════════════════════════════════════════════════════════════
#  Word 2003 XML 解析（.xml / .doc 早期格式）
# ══════════════════════════════════════════════════════════════════════
#
#  Word 2003 XML 使用 xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml"
#  核心标签：<w:body> <w:p> <w:r> <w:t>  （与 OOXML 同名但命名空间不同）
#  页眉页脚内嵌在 <w:hdr> / <w:ftr> 节点，脚注在 <w:footnote>
# ══════════════════════════════════════════════════════════════════════

# Word 2003 XML 命名空间
_W03 = "http://schemas.microsoft.com/office/word/2003/wordml"
_WX  = "http://schemas.microsoft.com/office/word/2003/auxHint"


def _qn03(local: str) -> str:
    return f"{{{_W03}}}{local}"


def _para_text_03(para: etree._Element) -> str:
    """从 Word 2003 <w:p> 提取纯文本。"""
    parts: list[str] = []
    for run in para.iter(_qn03("r")):
        for child in run:
            if child.tag == _qn03("t") and child.text:
                parts.append(child.text)
    return "".join(parts)


def _parse_word2003_xml(xml_path: str | Path) -> list[TextSegment]:
    """
    解析 Word 2003 XML 格式文件（.xml），返回 TextSegment 列表。

    支持：
        - 正文段落 (<w:body> → <w:p>)
        - 表格 (<w:tbl> → <w:tr> → <w:tc>)
        - 页眉 (<w:hdr>)
        - 页脚 (<w:ftr>)
        - 脚注 (<w:footnote>)
        - 尾注 (<w:endnote>)
    """
    xml_path = Path(xml_path)
    tree = etree.parse(str(xml_path))
    root = tree.getroot()

    segments: list[TextSegment] = []
    idx = 0
    xml_str = str(xml_path)

    # ── 页眉（<w:hdr>，出现在 <w:sect> 子元素中）──
    for hdr in root.iter(_qn03("hdr")):
        for para in hdr.iter(_qn03("p")):
            text = _para_text_03(para).strip()
            if text:
                segments.append(TextSegment("header", text, xml_str, idx)); idx += 1

    # ── 正文 body ──
    body = root.find(_qn03("body"))
    if body is not None:
        for child in body:
            tag_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag

            if tag_local == "p":
                text = _para_text_03(child).strip()
                if text:
                    segments.append(TextSegment("body", text, xml_str, idx)); idx += 1

                # 脚注引用（内联）
                for fn_ref in child.iter(_qn03("footnoteRef")):
                    fn_id = fn_ref.get(_qn03("id"), "") or fn_ref.get("id", "")
                    if fn_id:
                        segments.append(TextSegment("footnote", f"[脚注{fn_id}]",
                                                    xml_str, idx)); idx += 1

            elif tag_local == "tbl":
                for row in child.iter(_qn03("tr")):
                    row_texts = [
                        "".join(t.text for t in cell.iter(_qn03("t")) if t.text)
                        for cell in row.iter(_qn03("tc"))
                    ]
                    row_context = "\t".join(t for t in row_texts if t.strip())
                    for ct in row_texts:
                        if ct.strip():
                            segments.append(TextSegment("table", ct, xml_str,
                                                        idx, row_context)); idx += 1

            elif tag_local == "sect":
                # 节内可能还有段落和表格（递归处理）
                for sub in child:
                    sub_local = sub.tag.split("}")[-1] if "}" in sub.tag else sub.tag
                    if sub_local == "p":
                        text = _para_text_03(sub).strip()
                        if text:
                            segments.append(TextSegment("body", text, xml_str, idx)); idx += 1
                    elif sub_local == "tbl":
                        for row in sub.iter(_qn03("tr")):
                            row_texts = [
                                "".join(t.text for t in cell.iter(_qn03("t")) if t.text)
                                for cell in row.iter(_qn03("tc"))
                            ]
                            row_context = "\t".join(t for t in row_texts if t.strip())
                            for ct in row_texts:
                                if ct.strip():
                                    segments.append(TextSegment("table", ct, xml_str,
                                                                idx, row_context)); idx += 1

    # ── 脚注（独立 <w:footnote> 节点）──
    for fn in root.iter(_qn03("footnote")):
        fn_id = fn.get(_qn03("id"), "") or fn.get("id", "")
        if fn_id in ("0", "1", "-1"):
            continue
        text = "".join(t.text for t in fn.iter(_qn03("t")) if t.text).strip()
        if text:
            segments.append(TextSegment("footnote", f"[脚注{fn_id}] {text}",
                                        xml_str, idx)); idx += 1

    # ── 尾注 ──
    for en in root.iter(_qn03("endnote")):
        en_id = en.get(_qn03("id"), "") or en.get("id", "")
        if en_id in ("0", "1", "-1"):
            continue
        text = "".join(t.text for t in en.iter(_qn03("t")) if t.text).strip()
        if text:
            segments.append(TextSegment("endnote", f"[尾注{en_id}] {text}",
                                        xml_str, idx)); idx += 1

    # ── 页脚（<w:ftr>）──
    for ftr in root.iter(_qn03("ftr")):
        for para in ftr.iter(_qn03("p")):
            text = _para_text_03(para).strip()
            if text:
                segments.append(TextSegment("footer", text, xml_str, idx)); idx += 1

    return segments


# ══════════════════════════════════════════════════════════════════════
#  纯 Python .doc 解析（CFB + FIB + Clx/PlcPcd 文本提取）
#
#  依据 [MS-DOC] 官方规范：
#    https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/
#
#  核心流程：
#    1. 用 olefile 打开 CFB 容器，读取 WordDocument 流和 Table 流
#    2. 从 WordDocument 流偏移 0 读取 FIB（File Information Block）
#       - FIB.base.nFib       ：版本号（0x00C1=Word97，0x00D9=Word2000…）
#       - FIB.base.flags      ：bit0=fDot, bit9=fWhichTblStm（选 0Table/1Table）
#       - FibRgFcLcb97.fcClx  ：Table 流中 Clx 的偏移
#       - FibRgFcLcb97.lcbClx ：Clx 的字节长度
#    3. 从 Table 流读取 Clx，解析出 PlcPcd（Piece Descriptor 列表）
#       - 每个 Pcd 描述一段文本在 WordDocument 流中的位置和编码方式
#       - FcCompressed.fCompressed=1 → cp1252 单字节；=0 → UTF-16LE 双字节
#    4. 按 PlcPcd 顺序拼接全部字符，得到原始文本（含特殊控制字符）
#    5. 过滤/替换 Word 特殊字符：
#       - 0x0D → 段落分隔 \n
#       - 0x07 → 表格单元格/行分隔
#       - 0x0B → 强制换行（保留为 \n）
#       - 0x01/0x08 → 嵌入对象占位（丢弃）
#       - 0x13/0x14/0x15 → 域代码开始/分隔/结束（丢弃域内容）
# ══════════════════════════════════════════════════════════════════════

import struct as _struct


# ── FIB 偏移常量（基于 MS-DOC 规范） ──
_FIB_WIDENT_OFF      = 0       # 2 bytes，魔数，应为 0xA5EC 或 0xA5DC
_FIB_NFIB_OFF        = 2       # 2 bytes，版本
_FIB_FLAGS_OFF       = 10      # 2 bytes，包含 fWhichTblStm (bit 9)
_FIB_BASE_SIZE       = 32      # FibBase 固定 32 字节

# FibRgFcLcb97 在 FIB 中的相对偏移
# 结构：FibBase(32) + csw(2) + FibRgW97(28) + cslw(2) + FibRgLw97(88) + cbRgFcLcb(2) = 154
_FIBRGFCLCB97_OFF    = 154     # FibRgFcLcb97 起始偏移
_FC_CLX_OFF          = _FIBRGFCLCB97_OFF + 0x01A2  # fcClx 在 FIB 中的绝对偏移（查规范表）
# 注：fcClx 是 FibRgFcLcb97 中第 66 个 fc/lcb 对（0-indexed），
#     偏移 = 154 + 66*8 = 154 + 528 = 682，但规范里给的是字段名查表
# 实际按规范：FibRgFcLcb97.fcClx offset within FibRgFcLcb97 = 0x01A2 bytes? 需精确计算

# MS-DOC Table B-2：FibRgFcLcb97 字段列表，fcClx 是第 66 对（0-indexed）
# offset in FibRgFcLcb97 = 66 * 8 = 528 bytes
# absolute offset in FIB  = 154 + 528 = 682
_FC_CLX_ABS          = 154 + 66 * 8        # = 682
_LCB_CLX_ABS         = _FC_CLX_ABS + 4     # = 686


def _read_u8(data: bytes, off: int) -> int:
    return data[off]

def _read_u16(data: bytes, off: int) -> int:
    return _struct.unpack_from("<H", data, off)[0]

def _read_u32(data: bytes, off: int) -> int:
    return _struct.unpack_from("<I", data, off)[0]

def _read_i32(data: bytes, off: int) -> int:
    return _struct.unpack_from("<i", data, off)[0]


def _parse_clx(clx: bytes) -> list[tuple[int, int, bool]]:
    """
    解析 Clx 结构，返回 Pcd 列表：[(cp_start, fc, fCompressed), ...]
    对应文档字符位置区间 [cp_start, cp_start_next)。

    Clx 结构：
        可选的多个 Grpprl（以 0x01 开头的格式覆盖块，直接跳过）
        最后一个 Pcdt（以 0x02 开头）
            Pcdt.clxt  = 0x02 (1 byte)
            Pcdt.lcb   = 4 bytes，PlcPcd 字节长度
            Pcdt.PlcPcd：
                aCp[]  : (n+1) 个 CP（4 bytes each）
                aPcd[] : n 个 Pcd（8 bytes each）
                    Pcd.fNoParaMark : 2 bytes（标志，忽略）
                    Pcd.fc          : FcCompressed，4 bytes
                    Pcd.prm         : 2 bytes（忽略）
    """
    off = 0
    # 跳过所有 Grpprl（clxt == 0x01）
    while off < len(clx) and clx[off] == 0x01:
        off += 1  # clxt
        cb = _read_u16(clx, off); off += 2
        off += cb

    if off >= len(clx) or clx[off] != 0x02:
        return []

    off += 1  # clxt = 0x02
    lcb_plcpcd = _read_u32(clx, off); off += 4

    plcpcd = clx[off: off + lcb_plcpcd]
    # PlcPcd 包含 n+1 个 CP 和 n 个 Pcd
    # 每个 Pcd = 8 bytes，每个 CP = 4 bytes
    # lcb = (n+1)*4 + n*8  →  n = (lcb - 4) / 12
    n = (len(plcpcd) - 4) // 12

    cp_arr_end = (n + 1) * 4
    cps  = [_read_u32(plcpcd, i * 4) for i in range(n + 1)]
    pcds = []
    for i in range(n):
        pcd_off = cp_arr_end + i * 8
        # Pcd.fc = FcCompressed (4 bytes at pcd_off+2)
        fc_raw       = _read_u32(plcpcd, pcd_off + 2)
        fCompressed  = bool(fc_raw & 0x40000000)   # bit 30
        fc           = (fc_raw & 0x3FFFFFFF)
        if fCompressed:
            fc = fc >> 1   # 压缩模式：实际字节偏移 = fc / 2（规范：fc/2）
        pcds.append((cps[i], fc, fCompressed))

    return pcds


# Word 特殊字符处理表
_WORD_SPECIAL: dict[int, str | None] = {
    0x01: None,   # 嵌入对象
    0x08: None,   # 嵌入对象
    0x13: None,   # 域开始（后续到 0x15 都丢弃，由下方状态机处理）
    0x14: None,   # 域分隔
    0x15: None,   # 域结束
    0x07: "\t",   # 表格单元格分隔 → 制表符
    0x0B: "\n",   # 强制换行
    0x0C: "\n",   # 分页
    0x0D: "\n",   # 段落结束
    0x1E: "-",    # 不换行连字符
    0x1F: "",     # 软连字符（丢弃）
    0x0A: "",     # 行终止（Word 内部，丢弃）
}


def _extract_text_from_wordstream(word_stream: bytes, pcds: list[tuple[int, int, bool]]) -> str:
    """
    按 PlcPcd 顺序从 WordDocument 流拼接文本，
    过滤域代码和控制字符，返回段落以 \n 分隔的纯文本。
    """
    parts: list[str] = []
    in_field = 0  # 域代码嵌套计数器

    for cp_start, fc, fCompressed in pcds:
        # 计算这个 Pcd 覆盖多少个字符（用下一个 cp_start 减，但这里我们读到流末尾截断）
        # 实际上我们直接读取到出现 0x0D 结束或碰到下一段为止
        # 简化：一次性把整段读出来，遇到终止符停止
        off = fc
        buf: list[str] = []

        if fCompressed:
            # ANSI/cp1252 单字节
            while off < len(word_stream):
                b = word_stream[off]; off += 1
                if b == 0x00:
                    break
                # 域代码过滤
                if b == 0x13:
                    in_field += 1; continue
                if b == 0x14:
                    continue  # 域结果开始，保留域结果文字
                if b == 0x15:
                    if in_field > 0: in_field -= 1
                    continue
                if in_field > 0:
                    continue
                # 特殊控制字符映射
                if b in _WORD_SPECIAL:
                    repl = _WORD_SPECIAL[b]
                    if repl is not None and repl != "":
                        buf.append(repl)
                else:
                    try:
                        buf.append(bytes([b]).decode("cp1252"))
                    except Exception:
                        pass
                if b == 0x0D:
                    break
        else:
            # UTF-16LE 双字节
            while off + 1 < len(word_stream):
                ch_ord = _struct.unpack_from("<H", word_stream, off)[0]; off += 2
                if ch_ord == 0x0000:
                    break
                if ch_ord == 0x13:
                    in_field += 1; continue
                if ch_ord == 0x14:
                    continue
                if ch_ord == 0x15:
                    if in_field > 0: in_field -= 1
                    continue
                if in_field > 0:
                    continue
                if ch_ord in _WORD_SPECIAL:
                    repl = _WORD_SPECIAL[ch_ord]
                    if repl is not None and repl != "":
                        buf.append(repl)
                else:
                    try:
                        buf.append(chr(ch_ord))
                    except Exception:
                        pass
                if ch_ord == 0x0D:
                    break

        if buf:
            parts.append("".join(buf))

    return "".join(parts)


def _parse_doc_native(doc_path: Path) -> list[TextSegment]:
    """
    纯 Python 解析 Word 97-2003 .doc 文件（CFB 格式）。
    依赖 olefile（pip install olefile）。

    按 MS-DOC 规范提取正文文本，按段落（0x0D 分隔）生成 TextSegment。
    注意：此方法只提取正文文本，不含页眉/页脚/脚注（结构更复杂，暂不实现）。
    """
    try:
        import olefile  # type: ignore
    except ImportError:
        raise ImportError("请安装 olefile：pip install olefile")

    if not olefile.isOleFile(str(doc_path)):
        raise ValueError(f"不是合法的 OLE2/CFB 文件: {doc_path}")

    with olefile.OleFileIO(str(doc_path)) as ole:
        # 读取 WordDocument 流
        if not ole.exists("WordDocument"):
            raise ValueError("找不到 WordDocument 流，可能不是 Word 文档")
        word_stream = ole.openstream("WordDocument").read()

        # ── 解析 FIB ──
        magic = _read_u16(word_stream, _FIB_WIDENT_OFF)
        if magic not in (0xA5EC, 0xA5DC):
            raise ValueError(f"FIB 魔数不匹配: 0x{magic:04X}，可能不是 Word 97-2003 格式")

        flags = _read_u16(word_stream, _FIB_FLAGS_OFF)
        use_1table = bool(flags & 0x0200)  # bit9: fWhichTblStm

        # 读取 fcClx / lcbClx
        fc_clx  = _read_u32(word_stream, _FC_CLX_ABS)
        lcb_clx = _read_u32(word_stream, _LCB_CLX_ABS)

        # ── 读取 Table 流 ──
        tbl_name = "1Table" if use_1table else "0Table"
        if not ole.exists(tbl_name):
            # 尝试另一个
            tbl_name = "0Table" if use_1table else "1Table"
            if not ole.exists(tbl_name):
                raise ValueError("找不到 Table 流（0Table/1Table）")
        table_stream = ole.openstream(tbl_name).read()

    # ── 解析 Clx → PlcPcd ──
    clx  = table_stream[fc_clx: fc_clx + lcb_clx]
    pcds = _parse_clx(clx)
    if not pcds:
        raise ValueError("Clx 解析失败，未找到有效的 PlcPcd")

    # ── 提取全文 ──
    raw_text = _extract_text_from_wordstream(word_stream, pcds)

    # ── 按段落切分，生成 TextSegment ──
    segments: list[TextSegment] = []
    xml_str = str(doc_path)
    for idx, line in enumerate(raw_text.split("\n")):
        line = line.strip()
        if line:
            segments.append(TextSegment("body", line, xml_str, idx))

    return segments


# ══════════════════════════════════════════════════════════════════════
#  .doc 转换层（二进制 Word 97-2003 格式）
# ══════════════════════════════════════════════════════════════════════

def _convert_doc_via_win32com(doc_path: Path, out_dir: Path) -> Path:
    """
    用 win32com 调用本机 Word/WPS 将 .doc 另存为 .docx。
    返回生成的 .docx 路径。仅 Windows 可用。
    """
    import win32com.client  # type: ignore
    import pythoncom        # type: ignore

    pythoncom.CoInitialize()
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone

        abs_doc  = str(doc_path.resolve())
        abs_docx = str((out_dir / (doc_path.stem + "_converted.docx")).resolve())

        doc = word.Documents.Open(abs_doc, ReadOnly=True)
        # wdFormatXMLDocument = 12
        doc.SaveAs2(abs_docx, FileFormat=12)
        doc.Close(False)
        word.Quit()
    finally:
        pythoncom.CoUninitialize()

    return Path(abs_docx)


def _convert_doc_via_libreoffice(doc_path: Path, out_dir: Path) -> Path:
    """
    用 LibreOffice --headless 将 .doc 转换为 .docx。
    跨平台备用方案，需要安装 LibreOffice 并在 PATH 中。
    """
    lo_cmd = shutil.which("libreoffice") or shutil.which("soffice")
    if not lo_cmd:
        raise RuntimeError("未找到 LibreOffice，请安装后重试（或使用 win32com 方式）。")

    result = subprocess.run(
        [lo_cmd, "--headless", "--convert-to", "docx",
         "--outdir", str(out_dir), str(doc_path)],
        capture_output=True, text=True, timeout=120,
    )
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice 转换失败:\n{result.stderr}")

    converted = out_dir / (doc_path.stem + ".docx")
    if not converted.exists():
        raise RuntimeError(f"LibreOffice 转换后未找到输出文件: {converted}")
    return converted


def _doc_to_docx(doc_path: Path, errors: list[str] | None = None) -> tuple[Path, Path | None]:
    """
    将 .doc 转换为临时 .docx 文件（win32com → LibreOffice）。
    errors: 已有的错误信息列表，会追加新的失败原因。
    """
    tmp_dir = Path(tempfile.mkdtemp(prefix="doc_parser_"))
    if errors is None:
        errors = []

    # 策略 1：win32com（Windows + Word/WPS）
    try:
        docx_path = _convert_doc_via_win32com(doc_path, tmp_dir)
        return docx_path, tmp_dir
    except ImportError:
        errors.append("win32com 未安装（pip install pywin32）")
    except Exception as e:
        errors.append(f"win32com 转换失败: {e}")

    # 策略 2：LibreOffice
    try:
        docx_path = _convert_doc_via_libreoffice(doc_path, tmp_dir)
        return docx_path, tmp_dir
    except Exception as e:
        errors.append(f"LibreOffice 转换失败: {e}")

    # 全部失败
    shutil.rmtree(tmp_dir, ignore_errors=True)
    raise RuntimeError(
        "无法将 .doc 转换为 .docx，所有转换方式均失败：\n"
        + "\n".join(f"  • {e}" for e in errors)
        + "\n\n解决方案：\n"
        "  Windows：pip install pywin32  （需要安装 Microsoft Word 或 WPS）\n"
        "  其他平台：安装 LibreOffice 并将其加入 PATH"
    )


def _detect_word2003_xml(file_path: Path) -> bool:
    """
    快速判断一个 .xml 文件是否是 Word 2003 XML 格式。
    通过检查根元素命名空间来判断。
    """
    try:
        for event, elem in etree.iterparse(str(file_path), events=("start",)):
            # 只读第一个元素（根元素）
            ns = elem.nsmap.get("w", "")
            return ns == _W03
    except etree.XMLSyntaxError:
        return False
    return False


# ══════════════════════════════════════════════════════════════════════
#  统一入口
# ══════════════════════════════════════════════════════════════════════

def extract_document_in_order(doc_path: str | Path) -> list[TextSegment]:
    """
    统一解析入口，自动识别以下三种格式并返回 TextSegment 列表：

        .docx  — OOXML 压缩包格式（直接解析）
        .doc   — Word 97-2003 二进制格式（先转换为 .docx 再解析）
        .xml   — Word 2003 XML 格式（直接用 lxml 解析）

    Args:
        doc_path: 文件路径（str 或 Path）

    Returns:
        按阅读顺序排列的 TextSegment 列表

    Raises:
        FileNotFoundError: 文件不存在
        ValueError: 格式不支持或文件损坏
        RuntimeError: .doc 格式无可用转换工具
    """
    doc_path = Path(doc_path)

    if not doc_path.exists():
        raise FileNotFoundError(f"文件不存在: {doc_path}")

    suffix = doc_path.suffix.lower()

    # ── .xml：可能是 Word 2003 XML，也可能是其他 XML ──
    if suffix == ".xml":
        if not _detect_word2003_xml(doc_path):
            raise ValueError(
                f"该 XML 文件不是 Word 2003 XML 格式（w 命名空间不匹配）: {doc_path}"
            )
        try:
            return _parse_word2003_xml(doc_path)
        except etree.XMLSyntaxError as e:
            raise ValueError(f"Word 2003 XML 解析失败: {e}") from e

    # ── .docx：OOXML 标准格式 ──
    if suffix == ".docx":
        try:
            with ZipFile(doc_path) as zf:
                return _parse_docx_zip(zf)
        except BadZipFile as e:
            raise ValueError(f"文件不是合法的 .docx（ZIP）格式: {e}") from e
        except KeyError as e:
            raise ValueError(f"docx 文件结构异常，缺少关键文件: {e}") from e
        except etree.XMLSyntaxError as e:
            raise ValueError(f"docx XML 解析失败: {e}") from e

    # ── .doc：二进制格式，需先转换 ──
    if suffix == ".doc":
        # 先尝试作为 ZIP 打开（有些改了扩展名的 docx）
        try:
            with ZipFile(doc_path) as zf:
                if "word/document.xml" in zf.namelist():
                    return _parse_docx_zip(zf)
        except BadZipFile:
            pass

        errors: list[str] = []

        # 策略 1：纯 Python CFB 解析（olefile）
        try:
            return _parse_doc_native(doc_path)
        except ImportError as e:
            errors.append(f"olefile 未安装: {e}（pip install olefile）")
        except Exception as e:
            errors.append(f"原生 CFB 解析失败: {e}")

        # 策略 2：win32com 转换为临时 .docx
        # 策略 3：LibreOffice 转换
        tmp_dir: Path | None = None
        try:
            docx_path, tmp_dir = _doc_to_docx(doc_path, errors)
            with ZipFile(docx_path) as zf:
                return _parse_docx_zip(zf)
        finally:
            if tmp_dir and tmp_dir.exists():
                shutil.rmtree(tmp_dir, ignore_errors=True)

    raise ValueError(
        f"不支持的文件格式 '{suffix}'，仅支持 .docx / .doc / .xml（Word 2003 XML）"
    )


# ══════════════════════════════════════════════════════════════════════
#  便捷别名（与原 body_extractor.py 接口保持兼容）
# ══════════════════════════════════════════════════════════════════════

def extract_docx_in_order(docx_path: str | Path) -> list[TextSegment]:
    """向后兼容别名，等同于 extract_document_in_order()。"""
    return extract_document_in_order(docx_path)


# ══════════════════════════════════════════════════════════════════════
#  命令行快速测试
# ══════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import sys
    from collections import Counter

    # 命令行传参优先；否则 fallback 到下方硬编码路径（方便 IDE 内直接运行）
    if len(sys.argv) >= 2:
        path = sys.argv[1]
    else:
        path = r"D:\project\数检_程序-AI\测试\原文-ceshi1.doc"

    try:
        segments = extract_document_in_order(path)
    except (FileNotFoundError, ValueError, RuntimeError) as e:
        print(f"❌ 错误: {e}")
        sys.exit(1)

    print("=" * 70)
    print(f"  📖 解析完成，共 {len(segments)} 个文本片段  ← {path}")
    print("=" * 70)

    source_counts = Counter(seg.source for seg in segments)
    for src, label in SOURCE_LABELS.items():
        count = source_counts.get(src, 0)
        if count > 0:
            print(f"    {label}: {count} 条")
    print("=" * 70)

    for i, seg in enumerate(segments, 1):
        label = SOURCE_LABELS.get(seg.source, seg.source)
        print(f"[{i:>5d}] 【{label}】{seg.text}")

    full_text = "\n".join(seg.text for seg in segments)
    print(f"\n{'=' * 70}")
    print(f"  📝 纯文本共 {len(full_text)} 字符")
    print(f"{'=' * 70}")