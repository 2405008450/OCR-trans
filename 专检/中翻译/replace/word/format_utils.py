"""
format_utils.py
替换修订完成后，对 <bold>/<italic> 标签标记的文本片段应用加粗/斜体格式。

关键设计：
  - 替换修订后新文本在 <w:ins><w:r> 里，paragraph.runs 不包含它
  - 必须直接遍历段落 XML，同时收集普通 <w:r> 和 <w:ins> 里的 <w:r>
  - 拆分 run 时只复制字体/字号，不复制颜色等属性，避免颜色污染
"""
import re
from copy import deepcopy
from typing import List, Dict, Tuple
from docx import Document
from docx.oxml.ns import qn
from docx.oxml.shared import OxmlElement

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

# ------------------------------------------------------------------ #
# 1. 标签清洗
# ------------------------------------------------------------------ #

_TAG_RE = re.compile(r'</?(?:bold|italic)>')


def strip_format_tags(text: str) -> str:
    """剥离 <bold>/<italic> 标签，返回纯文本。"""
    if not text:
        return text
    return _TAG_RE.sub('', text)


# ------------------------------------------------------------------ #
# 2. 标签解析 → 片段列表
# ------------------------------------------------------------------ #

def parse_format_segments(text: str) -> List[Dict]:
    """
    解析含 <bold>/<italic> 标签的文本，返回片段列表。
    每个片段：{"text": str, "bold": bool, "italic": bool}
    支持嵌套 <bold><italic>text</italic></bold>
    """
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
            segments.append({
                "text": part,
                "bold": bold_depth > 0,
                "italic": italic_depth > 0,
            })

    return segments


# ------------------------------------------------------------------ #
# 3. XML 层面收集段落内所有 run（含 <w:ins> 里的）
# ------------------------------------------------------------------ #

def _collect_runs_from_para_xml(p_elem) -> List[Tuple]:
    """
    直接遍历段落 XML，收集所有有效 <w:r>，包括：
      - 直接子 <w:r>
      - <w:ins> 里的 <w:r>（替换修订后新文本在这里）

    返回：[(r_elem, parent_elem), ...]
    按文档顺序排列。
    """
    result = []
    W_R = f'{{{W_NS}}}r'
    W_INS = f'{{{W_NS}}}ins'
    W_DEL = f'{{{W_NS}}}del'
    W_PPR = f'{{{W_NS}}}pPr'

    for child in p_elem:
        tag = child.tag
        if tag == W_PPR or tag == W_DEL:
            # 跳过段落属性和删除标记
            continue
        elif tag == W_R:
            result.append((child, p_elem))
        elif tag == W_INS:
            # ins 里可能有多个 run
            for r in child:
                if r.tag == W_R:
                    result.append((r, child))

    return result


def _get_run_text(r_elem) -> str:
    """获取 run 的文本内容。"""
    W_T = f'{{{W_NS}}}t'
    parts = []
    for child in r_elem:
        if child.tag == W_T:
            parts.append(child.text or '')
    return ''.join(parts)


def _set_run_text(r_elem, text: str):
    """设置 run 的文本内容。"""
    W_T = f'{{{W_NS}}}t'
    for child in list(r_elem):
        if child.tag == W_T:
            r_elem.remove(child)
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    r_elem.append(t)


# ------------------------------------------------------------------ #
# 4. run 格式设置（只改 bold/italic，不动其他属性）
# ------------------------------------------------------------------ #

def _set_bold(r_elem, bold: bool):
    """
    设置/清除加粗，不影响其他 rPr 属性。
    bold=False 时写入 <w:b w:val="0"/> 显式覆盖，
    确保来自 rStyle 样式继承的加粗也能被关闭。
    """
    rPr = r_elem.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        r_elem.insert(0, rPr)

    for tag_str in ['w:b', 'w:bCs']:
        elem = rPr.find(qn(tag_str))
        if bold:
            if elem is None:
                elem = OxmlElement(tag_str)
                rPr.append(elem)
            elem.attrib.pop(qn('w:val'), None)  # 恢复加粗
        else:
            if elem is None:
                elem = OxmlElement(tag_str)
                rPr.append(elem)
            elem.set(qn('w:val'), '0')  # 显式关闭，覆盖 rStyle 继承


def _set_italic(r_elem, italic: bool):
    """
    设置/清除斜体，不影响其他 rPr 属性。
    italic=False 时写入 <w:i w:val="0"/> 显式覆盖样式继承。
    """
    rPr = r_elem.find(qn('w:rPr'))
    if rPr is None:
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
            if elem is None:
                elem = OxmlElement(tag_str)
                rPr.append(elem)
            elem.set(qn('w:val'), '0')  # 显式关闭，覆盖 rStyle 继承


# ------------------------------------------------------------------ #
# 5. 拆分 run 并应用格式
# ------------------------------------------------------------------ #

def _make_run_copy(ref_r_elem, text: str) -> object:
    """
    完整复制参考 run 的 rPr（字体/字号/颜色等全部保留），
    确保接受修订后样式与原译文完全一致。
    bold/italic 由调用方按需叠加设置。
    """
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
    """
    将一个 run 拆分为 prefix / target(格式化) / suffix 三段。
    parent_elem 是 r_elem 的直接父节点（可能是 p_elem 或 w:ins）。
    """
    idx = list(parent_elem).index(r_elem)
    parent_elem.remove(r_elem)
    insert_at = idx

    if prefix:
        pre_r = _make_run_copy(r_elem, prefix)
        parent_elem.insert(insert_at, pre_r)
        insert_at += 1

    tgt_r = _make_run_copy(r_elem, target)
    if bold:
        _set_bold(tgt_r, True)
    if italic:
        _set_italic(tgt_r, True)
    parent_elem.insert(insert_at, tgt_r)
    insert_at += 1

    if suffix:
        suf_r = _make_run_copy(r_elem, suffix)
        parent_elem.insert(insert_at, suf_r)


# ------------------------------------------------------------------ #
# 6. 段落级格式应用
# ------------------------------------------------------------------ #

def _apply_format_to_paragraph(p_elem, segments: List[Dict]) -> bool:
    """
    在段落 XML 中定位 segments 对应的文本，应用加粗/斜体。
    直接操作 lxml 元素，不依赖 paragraph.runs。

    Returns:
        bool: 是否成功应用
    """
    if not segments:
        return False

    has_format = any(s['bold'] or s['italic'] for s in segments)
    if not has_format:
        return False

    target_text = ''.join(s['text'] for s in segments)
    if not target_text.strip():
        return False

    # 收集所有 run（含 w:ins 里的）
    run_list = _collect_runs_from_para_xml(p_elem)
    if not run_list:
        return False

    full_text = ''.join(_get_run_text(r) for r, _ in run_list)
    start_pos = full_text.find(target_text)
    if start_pos == -1:
        return False

    # 按 segment 逐段处理
    abs_pos = start_pos

    for seg in segments:
        seg_text = seg['text']
        if not seg_text:
            continue
        seg_end = abs_pos + len(seg_text)

        if seg['bold'] or seg['italic']:
            # 重新收集（前面的拆分可能改变了 run_list）
            run_list = _collect_runs_from_para_xml(p_elem)
            full_text = ''.join(_get_run_text(r) for r, _ in run_list)

            # 找出覆盖 [abs_pos, seg_end) 的 run
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
                    # 整个 run 都在范围内，直接改格式
                    if seg['bold']:
                        _set_bold(r_elem, True)
                    if seg['italic']:
                        _set_italic(r_elem, True)
                else:
                    _split_run(p_elem, parent_elem, r_elem,
                               prefix, target, suffix,
                               seg['bold'], seg['italic'])

        abs_pos = seg_end

    return True


# ------------------------------------------------------------------ #
# 7. 遍历文档所有段落
# ------------------------------------------------------------------ #

def _iter_para_elems(doc: Document):
    """
    遍历文档所有段落的 lxml 元素（正文含表格 + 页眉页脚）。
    直接返回 p_elem，不封装为 Paragraph 对象。
    """
    W_P = f'{{{W_NS}}}p'

    # 正文（body 下所有 p，含表格内的）
    body = doc.element.body
    for elem in body.iter(W_P):
        yield elem

    # 页眉页脚
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


# ------------------------------------------------------------------ #
# 8. 主入口
# ------------------------------------------------------------------ #

def apply_format_pass(doc: Document, all_errors: list) -> int:
    """
    对含 <bold>/<italic> 标签的建议值在文档中应用格式。
    设计为替换成功后立即调用（单条或多条均可）。

    Args:
        doc:        python-docx Document 对象
        all_errors: 错误列表（含 "译文修改建议值" 字段）

    Returns:
        int: 成功应用格式的片段数
    """
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
        format_tasks.append({
            'segments': segments,
            'pure_text': pure_text,
            'context': context,
        })

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
            # 上下文打分：取上下文前50字符匹配
            score = 1.0
            if context:
                snippet = context[:50]
                score = 2.0 if snippet in full else 1.0
            candidates.append((score, p_elem))

        if not candidates:
            print(f"    ⚠ 格式应用: 未找到段落含 '{pure_text[:40]}'")
            continue

        candidates.sort(key=lambda x: x[0], reverse=True)
        best_p_elem = candidates[0][1]

        ok = _apply_format_to_paragraph(best_p_elem, segments)
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


# ------------------------------------------------------------------ #
# 9. 清除格式（去掉多余的 bold/italic）
# ------------------------------------------------------------------ #

def clear_format_pass(doc: Document, all_errors: list) -> int:
    """
    对 old_value 含 <bold>/<italic> 标签、但 new_value 不含标签的情况，
    在文档中找到替换后的新文本，清除对应 run 的加粗/斜体格式。

    应在 apply_format_pass 之后调用（或替换成功后立即调用单条版本）。

    Returns:
        int: 成功清除格式的段落数
    """
    tasks = []
    for err in all_errors:
        old_raw = err.get('译文数值') or ''
        new_raw = err.get('译文修改建议值') or ''
        # old 含格式标签，new 不含 → 需要清除格式
        old_has_fmt = ('<bold>' in old_raw or '<italic>' in old_raw)
        new_has_fmt = ('<bold>' in new_raw or '<italic>' in new_raw)
        if not old_has_fmt or new_has_fmt:
            continue

        old_bold = '<bold>' in old_raw
        old_italic = '<italic>' in old_raw
        # new_value 在文档里实际写入的是清洗后的纯文本
        pure_new = strip_format_tags(new_raw).strip()
        context = strip_format_tags(err.get('译文上下文') or '').strip()
        if not pure_new:
            continue
        tasks.append({
            'pure_new': pure_new,
            'context': context,
            'clear_bold': old_bold,
            'clear_italic': old_italic,
        })

    if not tasks:
        return 0

    all_para_elems = list(_iter_para_elems(doc))
    cleared = 0

    for task in tasks:
        pure_new = task['pure_new']
        context = task['context']
        clear_bold = task['clear_bold']
        clear_italic = task['clear_italic']

        candidates = []
        for p_elem in all_para_elems:
            run_list = _collect_runs_from_para_xml(p_elem)
            full = ''.join(_get_run_text(r) for r, _ in run_list)
            if pure_new not in full:
                continue
            score = 2.0 if (context and context[:50] in full) else 1.0
            candidates.append((score, p_elem))

        if not candidates:
            print(f"    ⚠ 格式清除: 未找到段落含 '{pure_new[:40]}'")
            continue

        candidates.sort(key=lambda x: x[0], reverse=True)
        best_p_elem = candidates[0][1]

        run_list = _collect_runs_from_para_xml(best_p_elem)
        full = ''.join(_get_run_text(r) for r, _ in run_list)
        start = full.find(pure_new)
        if start == -1:
            continue

        end = start + len(pure_new)
        char_offset = 0
        for r_elem, _ in run_list:
            r_text = _get_run_text(r_elem)
            r_start = char_offset
            r_end = char_offset + len(r_text)
            if r_end > start and r_start < end:
                if clear_bold:
                    _set_bold(r_elem, False)
                if clear_italic:
                    _set_italic(r_elem, False)
            char_offset = r_end

        cleared += 1
        print(f"    ✓ 格式已清除 [{'bold ' if clear_bold else ''}{'italic' if clear_italic else ''}]: '{pure_new[:40]}'")

    return cleared
