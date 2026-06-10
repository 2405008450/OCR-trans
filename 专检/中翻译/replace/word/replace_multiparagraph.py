"""
replace_multiparagraph.py

跨段落替换模块：处理 old_value/new_value 含 \\n 的多段落替换情况。

策略（方案C）：
  1. 把 old_value / new_value 按 \\n 拆成行列表
  2. 用 difflib 做行级 diff，得到 equal / replace / insert / delete 四类操作
  3. 在文档中定位 old_value 第一行所在的段落集合
  4. 按 opcode 执行：
       equal   → 只做格式修正（去多余 bold/italic）
       replace → 单段落内文本替换（复用现有 RevisionManager）
       insert  → 在目标位置前/后插入新段落（复制邻近段落样式）
       delete  → 标记段落为删除（Track Changes）
"""

import difflib
import re
from copy import deepcopy
from typing import List, Tuple, Optional

from docx import Document
from docx.oxml.ns import qn
from docx.oxml.shared import OxmlElement
from lxml import etree

from replace.word.format_utils import strip_format_tags, _set_bold, _set_italic, _collect_runs_from_para_xml, _get_run_text
from replace.word.revision import RevisionManager
from replace.word.replace_clean import clean_text_thoroughly, iter_body_paragraphs, iter_header_paragraphs, iter_footer_paragraphs

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


# ─────────────────────────────────────────────
# 轻量包装：让 lxml 元素也能当 paragraph 用（只需 .element）
# ─────────────────────────────────────────────

class _FakePara:
    """轻量包装，让新插入的 lxml p 元素能作为 insert_anchor 使用。"""
    def __init__(self, p_elem):
        self._element = p_elem


# ─────────────────────────────────────────────
# 辅助：段落文本（含 w:ins 里的新文本）
# ─────────────────────────────────────────────

def _para_full_text(p_elem) -> str:
    """取段落所有 run 的文字（含修订插入的部分）。"""
    return "".join(_get_run_text(r) for r, _ in _collect_runs_from_para_xml(p_elem))


def _para_text_clean(p_elem) -> str:
    return clean_text_thoroughly(_para_full_text(p_elem))


# ─────────────────────────────────────────────
# 辅助：克隆段落样式（pPr）
# ─────────────────────────────────────────────

def _clone_pPr(src_p_elem):
    """深拷贝段落属性 pPr，供插入新段落时复用。"""
    pPr = src_p_elem.find(qn("w:pPr"))
    if pPr is not None:
        return deepcopy(pPr)
    return None


# ─────────────────────────────────────────────
# 辅助：在指定段落前插入新段落（修订标记）
# ─────────────────────────────────────────────

def _insert_paragraph_before(anchor_p_elem, text: str, ref_p_elem,
                              revision_manager: RevisionManager) -> None:
    """
    在 anchor_p_elem 之前插入一个新段落，内容为 text，
    段落格式从 ref_p_elem 复制，并用 w:ins 标记为新增。
    """
    new_p = OxmlElement("w:p")

    # 复制段落格式
    pPr = _clone_pPr(ref_p_elem)
    if pPr is not None:
        new_p.append(pPr)

    # 构建 w:ins 包裹的 run
    date = revision_manager._now_iso()
    rev_id = revision_manager._next_rev_id()

    ins = OxmlElement("w:ins")
    ins.set(qn("w:id"), rev_id)
    ins.set(qn("w:author"), revision_manager.author)
    ins.set(qn("w:date"), date)

    new_r = OxmlElement("w:r")

    # 复制第一个 run 的字符格式
    first_r = ref_p_elem.find(qn("w:r"))
    if first_r is not None:
        rPr = first_r.find(qn("w:rPr"))
        if rPr is not None:
            new_r.append(deepcopy(rPr))

    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    new_r.append(t)
    ins.append(new_r)
    new_p.append(ins)

    # 插入到 anchor 之前
    anchor_p_elem.addprevious(new_p)


# ─────────────────────────────────────────────
# 辅助：在指定段落后插入新段落（修订标记）
# ─────────────────────────────────────────────

def _insert_paragraph_after(anchor_p_elem, text: str, ref_p_elem,
                             revision_manager: RevisionManager):
    """在 anchor_p_elem 之后插入新段落，返回新段落元素供后续链式插入。"""
    new_p = OxmlElement("w:p")

    pPr = _clone_pPr(ref_p_elem)
    if pPr is not None:
        new_p.append(pPr)

    date = revision_manager._now_iso()
    rev_id = revision_manager._next_rev_id()

    ins = OxmlElement("w:ins")
    ins.set(qn("w:id"), rev_id)
    ins.set(qn("w:author"), revision_manager.author)
    ins.set(qn("w:date"), date)

    new_r = OxmlElement("w:r")
    first_r = ref_p_elem.find(qn("w:r"))
    if first_r is not None:
        rPr = first_r.find(qn("w:rPr"))
        if rPr is not None:
            new_r.append(deepcopy(rPr))

    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    new_r.append(t)
    ins.append(new_r)
    new_p.append(ins)

    anchor_p_elem.addnext(new_p)
    return new_p  # 返回新段落元素，供链式 anchor 更新


# ─────────────────────────────────────────────
# 辅助：将整个段落标记为删除（修订）
# ─────────────────────────────────────────────

def _mark_paragraph_deleted(p_elem, revision_manager: RevisionManager) -> None:
    """
    把段落内所有 run 的 w:t 改为 w:del/w:delText，标记为删除。
    同时在段落末尾添加 w:rPr/w:del 标记段落符号删除。
    """
    date = revision_manager._now_iso()

    for r_elem in list(p_elem.findall(qn("w:r"))):
        # 找出所有 w:t
        t_elems = r_elem.findall(qn("w:t"))
        if not t_elems:
            continue

        text = "".join(t.text or "" for t in t_elems)

        # 构建 w:del 包裹
        del_id = revision_manager._next_rev_id()
        del_elem = OxmlElement("w:del")
        del_elem.set(qn("w:id"), del_id)
        del_elem.set(qn("w:author"), revision_manager.author)
        del_elem.set(qn("w:date"), date)

        del_r = deepcopy(r_elem)
        for t in del_r.findall(qn("w:t")):
            dt = OxmlElement("w:delText")
            dt.set(qn("xml:space"), "preserve")
            dt.text = t.text or ""
            t.addprevious(dt)
            del_r.remove(t)

        del_elem.append(del_r)
        p_elem.replace(r_elem, del_elem)

    # 标记段落符号删除（pPr/rPr/del）
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_elem.insert(0, pPr)
    rPr_in_pPr = pPr.find(qn("w:rPr"))
    if rPr_in_pPr is None:
        rPr_in_pPr = OxmlElement("w:rPr")
        pPr.append(rPr_in_pPr)
    para_del = OxmlElement("w:del")
    para_del.set(qn("w:id"), revision_manager._next_rev_id())
    para_del.set(qn("w:author"), revision_manager.author)
    para_del.set(qn("w:date"), date)
    rPr_in_pPr.append(para_del)


# ─────────────────────────────────────────────
# 辅助：单段落文本替换（带格式标签处理）
# ─────────────────────────────────────────────

def _replace_in_single_para(p_elem, old_text: str, new_text: str,
                             reason: str, revision_manager: RevisionManager) -> bool:
    """在单个段落中替换 old_text → new_text，走修订模式。"""
    from docx.text.paragraph import Paragraph
    try:
        para = Paragraph(p_elem, None)
        return revision_manager.replace_in_paragraph(para, old_text, new_text, reason)
    except Exception as e:
        print(f"    [多段落] 单段落替换异常: {e}")
        return False


# ─────────────────────────────────────────────
# 辅助：格式修正（清除多余 bold/italic）
# ─────────────────────────────────────────────

def _fix_format_in_para(p_elem, old_line_raw: str, new_line_raw: str) -> None:
    """
    比较 old/new 行的格式标签，对段落内对应文本做格式修正：
    - old 有 <bold> 而 new 没有 → 清除加粗
    - new 有 <bold> 而 old 没有 → 添加加粗
    同理处理 italic。
    """
    old_bold = "<bold>" in old_line_raw
    new_bold = "<bold>" in new_line_raw
    old_italic = "<italic>" in old_line_raw
    new_italic = "<italic>" in new_line_raw

    pure_new = strip_format_tags(new_line_raw).strip()
    if not pure_new:
        return

    run_list = _collect_runs_from_para_xml(p_elem)
    full = "".join(_get_run_text(r) for r, _ in run_list)
    start = full.find(pure_new)
    if start == -1:
        return
    end = start + len(pure_new)

    char_offset = 0
    for r_elem, _ in run_list:
        r_text = _get_run_text(r_elem)
        r_start = char_offset
        r_end = char_offset + len(r_text)
        if r_end > start and r_start < end:
            if old_bold and not new_bold:
                _set_bold(r_elem, False)
            elif not old_bold and new_bold:
                _set_bold(r_elem, True)
            if old_italic and not new_italic:
                _set_italic(r_elem, False)
            elif not old_italic and new_italic:
                _set_italic(r_elem, True)
        char_offset = r_end


# ─────────────────────────────────────────────
# 核心：定位 old_lines 对应的段落列表
# ─────────────────────────────────────────────

def _locate_old_paragraphs(
        old_lines: List[str],
        paragraph_iterator,
        context: str = ""
) -> Optional[List]:
    """
    在文档中找到与 old_lines 对应的段落序列。

    策略：
      1. 用 old_lines[0]（第一行）在所有段落中搜索候选起始段落
      2. 从候选段落开始，依次比对后续段落是否匹配 old_lines[1], [2]...
      3. 多个候选时用上下文相似度打分取最优

    Returns:
        匹配到的段落列表（python-docx Paragraph 对象），未找到返回 None
    """
    all_paras = list(paragraph_iterator())
    if not all_paras or not old_lines:
        return None

    first_line = strip_format_tags(old_lines[0]).strip()
    first_clean = clean_text_thoroughly(first_line)
    context_clean = clean_text_thoroughly(strip_format_tags(context))

    candidates = []
    for i, para in enumerate(all_paras):
        pt = _para_text_clean(para._element)
        if not pt:
            continue
        # 第一行匹配（精确或包含）
        if first_clean and (first_clean == pt or first_clean in pt or pt in first_clean):
            # 检查后续行是否连续匹配
            matched = [para]
            ok = True
            for j, line in enumerate(old_lines[1:], 1):
                line_clean = clean_text_thoroughly(strip_format_tags(line).strip())
                if not line_clean:
                    # 空行：接受任何空段落，或直接跳过
                    if i + j < len(all_paras):
                        matched.append(all_paras[i + j])
                    continue
                if i + j >= len(all_paras):
                    ok = False
                    break
                next_pt = _para_text_clean(all_paras[i + j]._element)
                if not (line_clean == next_pt or line_clean in next_pt or next_pt in line_clean):
                    ok = False
                    break
                matched.append(all_paras[i + j])

            if ok and len(matched) >= 1:
                # 上下文打分
                score = 1.0
                if context_clean:
                    para_text = " ".join(_para_text_clean(p._element) for p in matched)
                    words_p = set(para_text.lower().split())
                    words_c = set(context_clean.lower().split())
                    inter = len(words_p & words_c)
                    union = len(words_p | words_c)
                    score = inter / union if union else 0.0
                candidates.append((score, i, matched))

    if not candidates:
        return None

    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][2]


# ─────────────────────────────────────────────
# 主函数
# ─────────────────────────────────────────────

def replace_multiparagraph(
        doc: Document,
        old_value: str,
        new_value: str,
        reason: str,
        revision_manager: RevisionManager,
        context: str = "",
        region: str = "body"
) -> Tuple[bool, str]:
    """
    跨段落替换主函数。

    Args:
        doc:              python-docx Document
        old_value:        含 \\n 的旧译文（可含 <bold>/<italic> 标签）
        new_value:        含 \\n 的新译文（可含 <bold>/<italic> 标签）
        reason:           修改理由
        revision_manager: RevisionManager 实例
        context:          译文上下文（用于定位）
        region:           "body" / "header" / "footer"

    Returns:
        (success: bool, strategy_desc: str)
    """
    # 选段落迭代器
    if region == "header":
        para_iter = lambda: iter_header_paragraphs(doc)
    elif region == "footer":
        para_iter = lambda: iter_footer_paragraphs(doc)
    else:
        para_iter = lambda: iter_body_paragraphs(doc)

    # 按行拆分（保留格式标签，供格式处理使用）
    old_lines_raw = old_value.split("\n")
    new_lines_raw = new_value.split("\n")

    # 纯文本行（用于段落匹配和文本替换）
    old_lines_pure = [strip_format_tags(l).strip() for l in old_lines_raw]
    new_lines_pure = [strip_format_tags(l).strip() for l in new_lines_raw]

    # 过滤掉全空行参与 diff（但保留位置信息）
    old_nonempty = [(i, l) for i, l in enumerate(old_lines_pure) if l]
    new_nonempty = [(i, l) for i, l in enumerate(new_lines_pure) if l]

    if not old_nonempty:
        return False, "old_value 无有效内容"

    # 定位 old_lines 对应的段落
    matched_paras = _locate_old_paragraphs(
        [l for _, l in old_nonempty],
        para_iter,
        context=context
    )

    if not matched_paras:
        print(f"    [多段落] 未找到匹配段落，第一行: '{old_lines_pure[0][:50]}'")
        return False, "未找到匹配段落"

    print(f"    [多段落] 找到 {len(matched_paras)} 个匹配段落")

    # 用 difflib 对非空行做行级 diff
    old_texts = [l for _, l in old_nonempty]
    new_texts = [l for _, l in new_nonempty]

    sm = difflib.SequenceMatcher(None, old_texts, new_texts)
    opcodes = sm.get_opcodes()

    # 建立 old_text → paragraph 的映射（按顺序）
    para_map = {old_texts[i]: matched_paras[i] for i in range(min(len(old_texts), len(matched_paras)))}

    # 插入操作需要知道"在哪个段落前/后插入"，用第一个匹配段落作为基准
    first_para = matched_paras[0]
    last_para = matched_paras[-1]

    # 追踪插入偏移（insert 操作按顺序追加在 last_para 后）
    insert_anchor = last_para  # 最后已处理段落，insert 追加在其后

    ops_done = []

    for op_idx, (tag, i1, i2, j1, j2) in enumerate(opcodes):
        old_chunk = old_texts[i1:i2]
        new_chunk = new_texts[j1:j2]
        old_chunk_raw = [old_lines_raw[old_nonempty[i1 + k][0]] for k in range(i2 - i1)]
        new_chunk_raw = [new_lines_raw[new_nonempty[j1 + k][0]] for k in range(j2 - j1)]

        # 找下一个 equal/replace 操作中的段落（用于 insert 定位）
        def _next_anchor_para(from_idx):
            """找从 from_idx 开始之后第一个 equal/replace 的已有段落"""
            for ni in range(from_idx + 1, len(opcodes)):
                nt, ni1, ni2, _, _ = opcodes[ni]
                if nt in ("equal", "replace") and ni2 > ni1:
                    p = para_map.get(old_texts[ni1])
                    if p is not None:
                        return p
            return None

        if tag == "equal":
            # 文本不变，只做格式修正
            for k, old_text in enumerate(old_chunk):
                para = para_map.get(old_text)
                if para is None:
                    continue
                old_raw = old_chunk_raw[k] if k < len(old_chunk_raw) else old_text
                new_raw = new_chunk_raw[k] if k < len(new_chunk_raw) else old_text
                _fix_format_in_para(para._element, old_raw, new_raw)
                insert_anchor = para
            ops_done.append(f"equal({len(old_chunk)}行)")

        elif tag == "replace":
            # 逐行替换
            trailing_inserts = []  # 收集 replace 尾部多出来的 new 行，最后反向插入
            for k in range(max(len(old_chunk), len(new_chunk))):
                if k < len(old_chunk) and k < len(new_chunk):
                    old_text = old_chunk[k]
                    new_text = new_chunk[k]
                    para = para_map.get(old_text)
                    if para is not None:
                        old_raw = old_chunk_raw[k]
                        new_raw = new_chunk_raw[k]
                        if old_text != new_text:
                            ok = _replace_in_single_para(
                                para._element, old_text, new_text, reason, revision_manager
                            )
                            print(f"    [多段落] replace '{old_text[:30]}' → '{new_text[:30]}': {'✓' if ok else '✗'}")
                        _fix_format_in_para(para._element, old_raw, new_raw)
                        insert_anchor = para
                elif k < len(old_chunk):
                    old_text = old_chunk[k]
                    para = para_map.get(old_text)
                    if para is not None:
                        _mark_paragraph_deleted(para._element, revision_manager)
                        print(f"    [多段落] delete '{old_text[:40]}'")
                else:
                    trailing_inserts.append(new_chunk[k])
            # 反向插入保证顺序正确
            for new_text in reversed(trailing_inserts):
                _insert_paragraph_after(
                    insert_anchor._element, new_text,
                    insert_anchor._element, revision_manager
                )
                print(f"    [多段落] insert(replace尾部) '{new_text[:40]}'")
            ops_done.append(f"replace({len(old_chunk)}→{len(new_chunk)}行)")

        elif tag == "insert":
            # 找后面是否有已知段落可以用 addprevious
            next_para = _next_anchor_para(op_idx)
            if next_para is not None:
                # 在下一个已知段落之前，正向插入（先插A再插B，B会在A后面因为addprevious是在同一节点前）
                # 正向 addprevious：A先插到X前，B再插到X前 → 最终顺序 A,B,X ✓
                for new_text in new_chunk:
                    _insert_paragraph_before(
                        next_para._element, new_text,
                        next_para._element, revision_manager
                    )
                    print(f"    [多段落] insert(before) '{new_text[:40]}'")
            else:
                # 没有后续段落，反向 addnext 到 insert_anchor 后
                for new_text in reversed(new_chunk):
                    _insert_paragraph_after(
                        insert_anchor._element, new_text,
                        insert_anchor._element, revision_manager
                    )
                    print(f"    [多段落] insert(after) '{new_text[:40]}'")
            ops_done.append(f"insert({len(new_chunk)}行)")

        elif tag == "delete":
            # 纯删除
            for old_text in old_chunk:
                para = para_map.get(old_text)
                if para is not None:
                    _mark_paragraph_deleted(para._element, revision_manager)
                    print(f"    [多段落] delete '{old_text[:40]}'")
            ops_done.append(f"delete({len(old_chunk)}行)")

    strategy = f"多段落替换 [{', '.join(ops_done)}] (region={region})"
    print(f"    [多段落] 完成: {strategy}")
    return True, strategy
