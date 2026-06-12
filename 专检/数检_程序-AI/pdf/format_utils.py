"""
format_utils.py
替换修订完成后，对 <bold>/<italic> 标签标记的文本片段应用加粗/斜体格式。

关键设计：
  - 替换修订后新文本在 <w:ins><w:r> 里，paragraph.runs 不包含它
  - 必须直接遍历段落 XML，同时收集普通 <w:r> 和 <w:ins> 里的 <w:r>
  - 拆分 run 时完整复制 rPr，确保接受修订后样式与原文一致
"""
import re
from copy import deepcopy
from typing import List, Dict, Tuple
from docx import Document
from docx.oxml.ns import qn
from docx.oxml.shared import OxmlElement

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

_TAG_RE = re.compile(r'</?(?:bold|italic)>')


def strip_format_tags(text: str) -> str:
    """剥离 <bold>/<italic> 标签，返回纯文本。"""
    if not text:
        return text
    return _TAG_RE.sub('', text)


def parse_format_segments(text: str) -> List[Dict]:
    """解析含标签的文本，返回片段列表 {"text", "bold", "italic"}，支持嵌套。"""
    if not text:
        return []
    segments = []
    token_re = re.compile(r'(<bold>|</bold>|<italic>|</italic>)')
    parts = token_re.split(text)
    bold_depth = 0
    italic_depth = 0
    for part in parts:
        if part == '<bold>':
            bold_depth += 1
        elif part == '</bold>':
            bold_depth = max(0, bold_depth - 1)
        elif part == '<italic>':
            italic_depth += 1
        elif part == '</italic>':
            italic_depth = max(0, italic_depth - 1)
        elif part:
            segments.append({"text": part, "bold": bold_depth > 0, "italic": italic_depth > 0})
    return segments


def _collect_runs_from_para_xml(p_elem) -> List[Tuple]:
    """收集段落内所有 <w:r>，含 <w:ins> 里的（替换修订后新文本在此）。"""
    result = []
    W_R = f'{{{W_NS}}}r'
    W_INS = f'{{{W_NS}}}ins'
    W_DEL = f'{{{W_NS}}}del'
    W_PPR = f'{{{W_NS}}}pPr'
    for child in p_elem:
        tag = child.tag
        if tag in (W_PPR, W_DEL):
            continue
        elif tag == W_R:
            result.append((child, p_elem))
        elif tag == W_INS:
            for r in child:
                if r.tag == W_R:
                    result.append((r, child))
    return result


def _get_run_text(r_elem) -> str:
    W_T = f'{{{W_NS}}}t'
    return ''.join(child.text or '' for child in r_elem if child.tag == W_T)


def _set_bold(r_elem, bold: bool):
    rPr = r_elem.find(qn('w:rPr'))
    if rPr is None:
        if not bold:
            return
        rPr = OxmlElement('w:rPr')
        r_elem.insert(0, rPr)
    for tag_str in ['w:b', 'w:bCs']:
        elem = rPr.find(qn(tag_str))
        if bold:
            if elem is None:
                elem = OxmlElement(tag_str)
                rPr.append(elem)
            elem.attrib.pop(qn('w:val'), None)
        else:
            if elem is not None:
                rPr.remove(elem)


def _set_italic(r_elem, italic: bool):
    rPr = r_elem.find(qn('w:rPr'))
    if rPr is None:
        if not italic:
            return
        rPr = OxmlElement('w:rPr')
        r_elem.insert(0, rPr)
    for tag_str in ['w:i', 'w:iCs']:
        elem = rPr.find(qn(tag_str))
        if italic:
            if elem is None:
                elem = OxmlElement(tag_str)
                rPr.append(elem)
            elem.attrib.pop(qn('w:val'), None)
        else:
            if elem is not None:
                rPr.remove(elem)


def _make_run_copy(ref_r_elem, text: str):
    new_r = OxmlElement('w:r')
    ref_rPr = ref_r_elem.find(qn('w:rPr'))
    if ref_rPr is not None:
        new_r.append(deepcopy(ref_rPr))
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    new_r.append(t)
    return new_r


def _split_run(p_elem, parent_elem, r_elem, prefix: str, target: str, suffix: str,
               bold: bool, italic: bool):
    idx = list(parent_elem).index(r_elem)
    parent_elem.remove(r_elem)
    insert_at = idx
    if prefix:
        parent_elem.insert(insert_at, _make_run_copy(r_elem, prefix))
        insert_at += 1
    tgt_r = _make_run_copy(r_elem, target)
    if bold:
        _set_bold(tgt_r, True)
    if italic:
        _set_italic(tgt_r, True)
    parent_elem.insert(insert_at, tgt_r)
    insert_at += 1
    if suffix:
        parent_elem.insert(insert_at, _make_run_copy(r_elem, suffix))


def _apply_format_to_paragraph(p_elem, segments: List[Dict]) -> bool:
    if not segments or not any(s['bold'] or s['italic'] for s in segments):
        return False
    target_text = ''.join(s['text'] for s in segments)
    if not target_text.strip():
        return False
    run_list = _collect_runs_from_para_xml(p_elem)
    if not run_list:
        return False
    full_text = ''.join(_get_run_text(r) for r, _ in run_list)
    start_pos = full_text.find(target_text)
    if start_pos == -1:
        return False

    abs_pos = start_pos
    for seg in segments:
        seg_text = seg['text']
        if not seg_text:
            continue
        seg_end = abs_pos + len(seg_text)
        if seg['bold'] or seg['italic']:
            run_list = _collect_runs_from_para_xml(p_elem)
            char_offset = 0
            affected = []
            for r_elem, parent_elem in run_list:
                r_text = _get_run_text(r_elem)
                r_start = char_offset
                r_end = char_offset + len(r_text)
                if r_end > abs_pos and r_start < seg_end:
                    affected.append((r_elem, parent_elem, r_start, r_end))
                char_offset = r_end
            for r_elem, parent_elem, r_start, r_end in affected:
                r_text = _get_run_text(r_elem)
                local_start = max(abs_pos, r_start) - r_start
                local_end = min(seg_end, r_end) - r_start
                prefix = r_text[:local_start]
                target = r_text[local_start:local_end]
                suffix = r_text[local_end:]
                if not target:
                    continue
                if not prefix and not suffix:
                    if seg['bold']:
                        _set_bold(r_elem, True)
                    if seg['italic']:
                        _set_italic(r_elem, True)
                else:
                    _split_run(p_elem, parent_elem, r_elem, prefix, target, suffix,
                               seg['bold'], seg['italic'])
        abs_pos = seg_end
    return True


def _iter_para_elems(doc: Document):
    W_P = f'{{{W_NS}}}p'
    for elem in doc.element.body.iter(W_P):
        yield elem
    for section in doc.sections:
        for hf in [
            section.header, section.first_page_header, section.even_page_header,
            section.footer, section.first_page_footer, section.even_page_footer,
        ]:
            try:
                if hf is None:
                    continue
                for elem in hf._element.iter(W_P):
                    yield elem
            except Exception:
                pass


def apply_format_pass(doc: Document, all_errors: list) -> int:
    """替换成功后立即调用，对含 <bold>/<italic> 标签的建议值应用格式。"""
    format_tasks = []
    for err in all_errors:
        suggestion = err.get('译文修改建议值') or ''
        if '<bold>' not in suggestion and '<italic>' not in suggestion:
            continue
        segments = parse_format_segments(suggestion)
        if not any(s['bold'] or s['italic'] for s in segments):
            continue
        pure_text = strip_format_tags(suggestion).strip()
        context = strip_format_tags(err.get('译文上下文') or '').strip()
        if not pure_text:
            continue
        format_tasks.append({'segments': segments, 'pure_text': pure_text, 'context': context})

    if not format_tasks:
        return 0

    all_para_elems = list(_iter_para_elems(doc))
    applied = 0

    for task in format_tasks:
        segments = task['segments']
        pure_text = task['pure_text']
        context = task['context']

        candidates = []
        for p_elem in all_para_elems:
            run_list = _collect_runs_from_para_xml(p_elem)
            full = ''.join(_get_run_text(r) for r, _ in run_list)
            if pure_text not in full:
                continue
            score = 2.0 if (context and context[:50] in full) else 1.0
            candidates.append((score, p_elem))

        if not candidates:
            print(f"    ⚠ 格式应用: 未找到段落含 '{pure_text[:40]}'")
            continue

        candidates.sort(key=lambda x: x[0], reverse=True)
        ok = _apply_format_to_paragraph(candidates[0][1], segments)
        if ok:
            applied += 1
            fmt_desc = ' | '.join(
                f"{'bold' if s['bold'] else ''}{'italic' if s['italic'] else ''}: '{s['text'][:20]}'"
                for s in segments if s['bold'] or s['italic']
            )
            print(f"    ✓ 格式已应用 [{fmt_desc}]")
        else:
            print(f"    ✗ 格式应用失败: '{pure_text[:50]}'")

    return applied
