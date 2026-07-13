"""
脚注和尾注替换模块

直接修改 footnotes.xml 和 endnotes.xml 文件，支持修订模式（Track Changes）
"""

from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree
from copy import deepcopy
import tempfile
import os
import random
from datetime import datetime
from pathlib import Path


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"


def _w(tag):
    return f'{{{W_NS}}}{tag}'


def _get_run_text(run):
    """获取一个run中所有<w:t>的合并文本"""
    return ''.join(t.text or '' for t in run.findall(f'.//{_w("t")}'))


def _make_run_with_text(rPr_source, text):
    """创建一个带格式的run"""
    run = etree.Element(_w('r'))
    if rPr_source is not None:
        run.append(deepcopy(rPr_source))
    t = etree.SubElement(run, _w('t'))
    t.set(f'{{{XML_NS}}}space', 'preserve')
    t.text = text
    return run


def _make_del_run(rPr_source, text):
    """创建一个删除标记的run（<w:delText>）"""
    run = etree.Element(_w('r'))
    if rPr_source is not None:
        run.append(deepcopy(rPr_source))
    dt = etree.SubElement(run, _w('delText'))
    dt.set(f'{{{XML_NS}}}space', 'preserve')
    dt.text = text
    return run


def _find_runs_spanning_text(runs, old_text):
    """
    找到old_text跨越了哪些run。
    返回 (start_run_idx, start_offset, end_run_idx, end_offset) 或 None
    """
    run_texts = []
    cumulative = 0
    for run in runs:
        text = _get_run_text(run)
        run_texts.append((cumulative, text))
        cumulative += len(text)

    full_text = ''.join(text for _, text in run_texts)
    if old_text not in full_text:
        return None

    pos = full_text.index(old_text)
    end_pos = pos + len(old_text)

    start_run_idx = None
    start_offset = 0
    end_run_idx = None
    end_offset = 0

    for i, (cum_start, text) in enumerate(run_texts):
        cum_end = cum_start + len(text)
        if start_run_idx is None and cum_end > pos:
            start_run_idx = i
            start_offset = pos - cum_start
        if cum_end >= end_pos:
            end_run_idx = i
            end_offset = end_pos - cum_start
            break

    if start_run_idx is None or end_run_idx is None:
        return None
    return start_run_idx, start_offset, end_run_idx, end_offset


def _replace_with_revision(runs, start_idx, start_offset, end_idx, end_offset,
                           old_text, new_text, author, date_str, rev_id_counter):
    """
    对runs[start_idx:end_idx+1]执行修订替换，保留每个run的原始格式。

    策略：
    - 前缀文本（old_text之前）→ 保留原run格式
    - 被覆盖部分 → <w:del>（每个run保留各自格式）
    - 新文本 → <w:ins>（复制第一个run的格式）
    - 后缀文本（old_text之后）→ 保留最后一个run格式
    """
    affected_runs = runs[start_idx:end_idx + 1]
    if not affected_runs:
        return False, rev_id_counter

    first_run = affected_runs[0]
    parent = first_run.getparent()
    if parent is None:
        return False, rev_id_counter

    insert_pos = list(parent).index(first_run)
    first_rPr = first_run.find(_w('rPr'))

    # 计算前缀和后缀
    if start_idx == end_idx:
        run_text = _get_run_text(affected_runs[0])
        prefix_text = run_text[:start_offset]
        suffix_text = run_text[end_offset:]
    else:
        prefix_text = _get_run_text(affected_runs[0])[:start_offset]
        suffix_text = _get_run_text(affected_runs[-1])[end_offset:]

    # 移除所有affected runs
    for run in affected_runs:
        parent.remove(run)

    current_pos = insert_pos

    # 1. 前缀run（保留原格式）
    if prefix_text:
        parent.insert(current_pos, _make_run_with_text(first_rPr, prefix_text))
        current_pos += 1

    # 2. <w:del> - 每个run保留各自格式
    del_elem = etree.Element(_w('del'))
    del_elem.set(_w('id'), str(rev_id_counter))
    del_elem.set(_w('author'), author)
    del_elem.set(_w('date'), date_str)
    rev_id_counter += 1

    for i, run in enumerate(affected_runs):
        run_text = _get_run_text(run)
        if i == 0 and i == len(affected_runs) - 1:
            del_text = run_text[start_offset:end_offset]
        elif i == 0:
            del_text = run_text[start_offset:]
        elif i == len(affected_runs) - 1:
            del_text = run_text[:end_offset]
        else:
            del_text = run_text
        if del_text:
            del_elem.append(_make_del_run(run.find(_w('rPr')), del_text))

    parent.insert(current_pos, del_elem)
    current_pos += 1

    # 3. <w:ins> - 复制第一个run的格式
    ins_elem = etree.Element(_w('ins'))
    ins_elem.set(_w('id'), str(rev_id_counter))
    ins_elem.set(_w('author'), author)
    ins_elem.set(_w('date'), date_str)
    rev_id_counter += 1
    ins_elem.append(_make_run_with_text(first_rPr, new_text))

    parent.insert(current_pos, ins_elem)
    current_pos += 1

    # 4. 后缀run（保留最后一个run格式）
    if suffix_text:
        last_rPr = affected_runs[-1].find(_w('rPr'))
        parent.insert(current_pos, _make_run_with_text(last_rPr, suffix_text))

    return True, rev_id_counter


def replace_in_footnotes_xml(doc_path: str, old_text: str, new_text: str,
                             reason: str = "", author: str = "翻译校对") -> bool:
    """
    在脚注/尾注中替换文本，使用修订模式（Track Changes），保留原格式。
    """
    if not old_text or not new_text:
        return False

    rev_id_counter = random.randint(100000, 999999)
    date_str = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
    replaced = False

    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            with ZipFile(doc_path, 'r') as zf:
                zf.extractall(temp_dir)

            for xml_name in ['word/footnotes.xml', 'word/endnotes.xml']:
                xml_path = os.path.join(temp_dir, xml_name)
                if not os.path.exists(xml_path):
                    continue

                tree = etree.parse(xml_path)
                root = tree.getroot()
                tag = 'footnote' if 'footnote' in xml_name else 'endnote'
                notes = root.findall(f'.//{_w(tag)}')

                for note in notes:
                    # 收集段落级别的run
                    paragraphs = note.findall(f'.//{_w("p")}')
                    for para in paragraphs:
                        runs = para.findall(_w('r'))
                        if not runs:
                            continue

                        span = _find_runs_spanning_text(runs, old_text)
                        if span is None:
                            continue

                        s_idx, s_off, e_idx, e_off = span
                        ok, rev_id_counter = _replace_with_revision(
                            runs, s_idx, s_off, e_idx, e_off,
                            old_text, new_text, author, date_str, rev_id_counter
                        )
                        if ok:
                            replaced = True
                            label = "脚注" if 'footnote' in xml_name else "尾注"
                            print(f"    [{label}] 修订替换: '{old_text[:30]}...' -> '{new_text[:30]}...'")

                if replaced:
                    tree.write(xml_path, encoding='utf-8', xml_declaration=True)

            if replaced:
                with ZipFile(doc_path, 'w', compression=ZIP_DEFLATED) as zf:
                    for root_dir, dirs, files in os.walk(temp_dir):
                        for f in files:
                            fp = os.path.join(root_dir, f)
                            zf.write(fp, os.path.relpath(fp, temp_dir))

    except Exception as e:
        print(f"    [脚注替换] 失败: {e}")
        import traceback
        traceback.print_exc()
        return False

    return replaced


def check_text_in_footnotes(doc_path: str, search_text: str) -> bool:
    """检查文本是否在脚注/尾注中"""
    try:
        with ZipFile(doc_path, 'r') as zf:
            for xml_name in ['word/footnotes.xml', 'word/endnotes.xml']:
                if xml_name not in zf.namelist():
                    continue
                root = etree.fromstring(zf.read(xml_name))
                tag = 'footnote' if 'footnote' in xml_name else 'endnote'
                for note in root.findall(f'.//{_w(tag)}'):
                    texts = [t.text for t in note.findall(f'.//{_w("t")}') if t.text]
                    if search_text in ''.join(texts):
                        return True
    except Exception as e:
        print(f"    [脚注检查] 失败: {e}")
    return False
