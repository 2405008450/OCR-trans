"""
基于修订（Track Changes）的替换模块（优化版）

核心改进：
1. 候选段落打分排序 — 替换得分最高的段落而非第一个碰到的
2. 上下文验证真正生效 — calculate_context_similarity 不再是死代码
3. 策略降级时验证门槛递增 — 匹配越宽松，上下文要求越严格
4. 唯一性安全网 — 多处出现且无上下文时拒绝替换
5. preprocess_special_cases 尊重 region 参数
"""
import re
from typing import Tuple, List, Optional
from docx import Document
from lxml import etree

from replace.word.revision import RevisionManager
from replace.word.replace_clean import (
    clean_text_thoroughly,
    _normalize_spaces,
    is_list_pattern,
    build_smart_pattern,
    extract_anchor_with_target,
    calculate_context_similarity,
    iter_all_paragraphs,
    iter_body_paragraphs,
    iter_header_paragraphs,
    iter_footer_paragraphs,
    is_fuzzy_match,
    get_alphanumeric_fingerprint,
    preprocess_special_cases,
)


# =========================
# 辅助：候选段落打分
# =========================

def _score_paragraph(full_text: str, old_value: str, context: str,
                     anchor_text: str, match_level: int) -> float:
    """对候选段落打分，分数越高越可能是正确的替换位置。

    Args:
        full_text: 段落完整文本
        old_value: 待替换的值
        context: LLM 返回的译文上下文
        anchor_text: LLM 返回的替换锚点
        match_level: 匹配层级 (1=精确, 2=清洗, 3=正则, 4=模糊, 5=宽松)

    Returns:
        0.0 ~ 1.0 的得分
    """
    score = 0.0
    full_clean = clean_text_thoroughly(full_text)

    # 维度1：上下文相似度（权重最高，0 ~ 0.5）
    if context:
        sim = calculate_context_similarity(full_clean, clean_text_thoroughly(context))
        score += sim * 0.5

    # 维度2：锚点命中（0 或 0.25）
    if anchor_text:
        anchor_clean = clean_text_thoroughly(anchor_text)
        if anchor_clean and anchor_clean in full_clean:
            score += 0.25

    # 维度3：匹配精确度（0.05 ~ 0.15）
    level_scores = {1: 0.15, 2: 0.12, 3: 0.10, 4: 0.07, 5: 0.05}
    score += level_scores.get(match_level, 0.05)

    # 维度4：old_value 在段落中的占比越高越可能是目标段落（0 ~ 0.1）
    if full_text:
        ratio = len(old_value) / len(full_text)
        score += min(ratio, 1.0) * 0.1

    return score


# =========================
# 1) 修订版替换核心函数
# =========================

def apply_revision(paragraph, runs, old_value, new_value, reason,
                   revision_manager: RevisionManager, match_type="正则", region="body"):
    """
    执行实际替换（修订模式）：在段落中找到 old_value 并用 Track Changes 替换。
    """
    if not runs:
        return False

    full_text = "".join(r.text or "" for r in runs)

    if old_value in full_text:
        ok = revision_manager.replace_in_paragraph(paragraph, old_value, new_value, reason=reason)
        return ok

    return False


def _try_match_in_paragraph(paragraph, old_value: str, old_value_clean: str,
                            pattern: str) -> Optional[int]:
    """尝试在段落中匹配 old_value，返回匹配层级（1-5），未匹配返回 None。

    层级越低越精确：
      1 = 精确包含
      2 = 清洗后包含
      3 = 正则匹配
      4 = 模糊匹配 / 指纹匹配
      5 = 无空格匹配
    """
    runs = list(paragraph.runs)
    if not runs:
        return None
    full_text = "".join(r.text or "" for r in runs)
    if not full_text.strip():
        return None

    # 层1：精确包含（单 run 或跨 run）
    if old_value in full_text:
        return 1

    full_text_clean = clean_text_thoroughly(full_text)

    # 层2：清洗后包含
    if old_value_clean and old_value_clean in full_text_clean:
        if old_value in full_text:
            return 2

    # 层3：正则匹配
    if pattern:
        try:
            if re.search(pattern, full_text_clean, flags=re.IGNORECASE | re.DOTALL):
                if old_value in full_text:
                    return 3
        except re.error:
            pass

    # 层4：模糊匹配 / 指纹匹配
    if is_fuzzy_match(full_text_clean, old_value_clean, threshold=0.85):
        if old_value in full_text:
            return 4

    fingerprint_old = get_alphanumeric_fingerprint(old_value)
    fingerprint_full = get_alphanumeric_fingerprint(full_text)
    if len(fingerprint_old) >= 3 and fingerprint_old in fingerprint_full:
        if old_value in full_text:
            return 4

    # 层5：无空格匹配
    full_no_space = full_text_clean.replace(' ', '')
    old_no_space = old_value_clean.replace(' ', '')
    if old_no_space and old_no_space in full_no_space:
        if old_value in full_text:
            return 5

    return None


def _execute_replace(paragraph, old_value: str, new_value: str, reason: str,
                     revision_manager: RevisionManager) -> bool:
    """在段落中执行实际替换，处理单 run 和跨 run 的情况。"""
    runs = list(paragraph.runs)
    if not runs:
        return False
    full_text = "".join(r.text or "" for r in runs)

    # 优先尝试跨 run 替换（最通用）
    if old_value in full_text:
        return revision_manager.replace_in_paragraph(paragraph, old_value, new_value, reason=reason)

    # 单 run 忽略空格的特殊情况
    old_no_space = old_value.replace(' ', '')
    for run in runs:
        run_text = run.text or ""
        run_no_space = run_text.replace(' ', '')
        if old_no_space and len(run_text.strip()) < 50 and old_no_space == run_no_space:
            revision_manager.replace_run_text(run, new_value, reason=reason)
            return True

    return False


# =========================
# 辅助：编号段落上下文定位替换
# =========================

def _replace_numbering_with_context(doc, old_value, new_value, reason,
                                     revision_manager, context, anchor_text,
                                     paragraph_iterator, region_desc):
    """对编号类值（如 iv. → v.）进行上下文验证后替换。

    编号通常不在段落 runs 文本中（Word 自动编号），所以这里的策略是：
    1. 从上下文中提取编号后面的文本关键词
    2. 遍历有编号的段落，用关键词匹配定位正确段落
    3. 通过修改该段落的编号起始值来修正序号

    对于编号出现在 runs 文本中的情况（手动编号），走常规替换逻辑。

    Returns:
        (bool, str) — (是否成功, 策略描述)
    """
    from docx.oxml.ns import qn as _qn

    context_clean = clean_text_thoroughly(context or "")
    anchor_clean = clean_text_thoroughly(anchor_text or "")

    # 从上下文中去掉编号前缀，提取纯内容部分
    def _strip_numbering_prefix(text):
        """去掉文本开头的编号前缀（如 iv., (1), a) 等）"""
        stripped = re.sub(r'^[ivxlcdm]+\.\s*', '', text.strip(), flags=re.IGNORECASE)
        stripped = re.sub(r'^\(\d+\)\s*', '', stripped)
        stripped = re.sub(r'^\d+[\.\)]\s*', '', stripped)
        stripped = re.sub(r'^\([a-z]\)\s*', '', stripped, flags=re.IGNORECASE)
        return stripped.strip()

    context_content = _strip_numbering_prefix(context_clean)

    # 解析编号值 → 整数
    def _numbering_to_int(s):
        """将各种编号格式解析为整数值"""
        s = s.strip().lower()
        # 罗马数字：iv. → 4
        m = re.match(r'^([ivxlcdm]+)\.$', s)
        if m:
            roman = m.group(1)
            roman_map = {'i': 1, 'v': 5, 'x': 10, 'l': 50, 'c': 100, 'd': 500, 'm': 1000}
            result = 0
            prev = 0
            for ch in reversed(roman):
                val = roman_map.get(ch, 0)
                if val < prev:
                    result -= val
                else:
                    result += val
                prev = val
            return result
        # 数字：(1) / 1. / 1)
        m = re.match(r'^\(?(\d+)[\.\)]$', s)
        if m:
            return int(m.group(1))
        # 字母：a. / (a) / a)
        m = re.match(r'^\(?([a-z])[\.\)]$', s)
        if m:
            return ord(m.group(1)) - ord('a') + 1
        return 0

    new_num_int = _numbering_to_int(new_value)
    if new_num_int <= 0:
        return False, ""

    # 收集候选段落（有自动编号或手动编号的段落）
    candidates = []
    for p in paragraph_iterator():
        para_text = (p.text or "").strip()
        if not para_text:
            continue

        # 检查段落是否有 Word 自动编号
        has_auto_numbering = False
        num_id_val = None
        ilvl_val = None
        try:
            pPr = p._element.pPr
            if pPr is not None:
                numPr = pPr.numPr
                if numPr is not None and numPr.numId is not None:
                    has_auto_numbering = True
                    num_id_val = numPr.numId.get(_qn('w:val'))
                    ilvl_elem = numPr.ilvl
                    ilvl_val = ilvl_elem.get(_qn('w:val')) if ilvl_elem is not None else '0'
        except:
            pass

        # 也检查手动编号（old_value 出现在 runs 文本中）
        full_text = "".join(r.text or "" for r in p.runs)
        has_manual_numbering = old_value in full_text

        if not has_auto_numbering and not has_manual_numbering:
            continue

        # 打分：核心是用去掉编号前缀的上下文内容匹配段落文本
        para_clean = clean_text_thoroughly(para_text)
        score = 0.0

        if context_content:
            context_content_lower = context_content.lower()
            para_clean_lower = para_clean.lower()
            if context_content_lower and context_content_lower in para_clean_lower:
                score += 0.6
            elif para_clean_lower and para_clean_lower in context_content_lower:
                score += 0.5
            else:
                sim = calculate_context_similarity(para_clean, context_content)
                score += sim * 0.4

        if context_clean:
            sim_full = calculate_context_similarity(para_clean, context_clean)
            score += sim_full * 0.1

        if anchor_clean:
            anchor_stripped = _strip_numbering_prefix(anchor_clean)
            if anchor_stripped and anchor_stripped.lower() in para_clean.lower():
                score += 0.25

        candidates.append((score, p, has_auto_numbering, has_manual_numbering,
                           para_text, num_id_val, ilvl_val))

    if not candidates:
        return False, ""

    candidates.sort(key=lambda x: x[0], reverse=True)
    best_score, best_para, best_auto, best_manual, best_text, best_num_id, best_ilvl = candidates[0]

    # 安全检查
    if len(candidates) > 1:
        second_score = candidates[1][0]
        if not context_clean and not anchor_clean:
            print(f"    [编号替换] '{old_value}' 出现 {len(candidates)} 处且无上下文，拒绝替换")
            return False, ""
        if best_score > 0 and second_score / best_score > 0.95:
            print(f"    [编号替换] 前两个候选得分接近 ({best_score:.3f} vs {second_score:.3f})，无法确定位置")
            return False, ""
        if best_score < 0.1:
            print(f"    [编号替换] 最佳候选得分过低 ({best_score:.3f})，上下文不匹配")
            return False, ""

    print(f"    [编号替换] 定位到段落: '{best_text[:60]}...' (得分:{best_score:.3f}, 候选:{len(candidates)})")

    # 执行替换
    if best_manual:
        ok = _execute_replace(best_para, old_value, new_value, reason, revision_manager)
        if ok:
            return True, f"编号上下文定位(手动编号) [{region_desc}] (得分:{best_score:.2f}, 候选:{len(candidates)})"

    if best_auto and best_num_id is not None:
        ilvl = best_ilvl or '0'
        try:
            numbering_part = doc.part.numbering_part
            numbering_xml = numbering_part.element

            abstract_num_id = None
            for num_elem in numbering_xml.findall(_qn('w:num')):
                if num_elem.get(_qn('w:numId')) == best_num_id:
                    abs_elem = num_elem.find(_qn('w:abstractNumId'))
                    if abs_elem is not None:
                        abstract_num_id = abs_elem.get(_qn('w:val'))
                    break

            if abstract_num_id is None:
                print(f"    [编号替换] 找不到 abstractNumId")
                return False, ""

            max_num_id = 0
            for num_elem in numbering_xml.findall(_qn('w:num')):
                nid = num_elem.get(_qn('w:numId'))
                if nid and nid.isdigit():
                    max_num_id = max(max_num_id, int(nid))
            new_num_id = str(max_num_id + 1)

            new_num = etree.SubElement(numbering_xml, _qn('w:num'))
            new_num.set(_qn('w:numId'), new_num_id)
            abs_ref = etree.SubElement(new_num, _qn('w:abstractNumId'))
            abs_ref.set(_qn('w:val'), abstract_num_id)

            lvl_override = etree.SubElement(new_num, _qn('w:lvlOverride'))
            lvl_override.set(_qn('w:ilvl'), ilvl)
            start_override = etree.SubElement(lvl_override, _qn('w:startOverride'))
            start_override.set(_qn('w:val'), str(new_num_int))

            pPr = best_para._element.pPr
            numPr = pPr.numPr
            numPr.numId.set(_qn('w:val'), new_num_id)

            return True, (f"编号上下文定位(自动编号 新numId={new_num_id} start={new_num_int}) "
                          f"[{region_desc}] (得分:{best_score:.2f}, 候选:{len(candidates)})")

        except Exception as e:
            print(f"    [编号替换] 修改自动编号失败: {e}")
            import traceback
            traceback.print_exc()

    return False, ""


# =========================
# 2) 主替换函数（修订版）
# =========================

def replace_and_revise_in_docx(
        doc: Document,
        old_value: str,
        new_value: str,
        reason: str,
        revision_manager: RevisionManager,
        context: str = "",
        anchor_text: str = "",
        region: str = "all"
) -> Tuple[bool, str]:
    """
    多策略执行替换（修订模式），带上下文验证和候选打分。

    核心改进：
    - 候选段落打分排序，替换最佳匹配而非首个匹配
    - 上下文相似度真正参与决策
    - 策略越宽松，验证门槛越高
    - 多处出现且无上下文时拒绝替换（安全网）

    Args:
        doc: Word 文档对象
        old_value: 原始值
        new_value: 新值
        reason: 修改理由
        revision_manager: 修订管理器
        context: 上下文文本
        anchor_text: 显式锚点文本
        region: 替换区域 ("all"/"body"/"header"/"footer")

    Returns:
        (是否成功, 匹配策略描述)
    """
    old_value_original = old_value or ""
    old_value = (old_value or "").strip()
    new_value = clean_text_thoroughly(new_value or "").strip()
    context = clean_text_thoroughly(context or "")

    if isinstance(reason, (list, tuple)):
        reason = " ".join([str(i) for i in reason if i]).strip()
    reason = reason or "数值/术语不一致"

    if not old_value or not new_value:
        return False, "数据缺失"

    # 根据 region 选择段落迭代器
    if region == "body":
        paragraph_iterator = lambda: iter_body_paragraphs(doc)
        region_desc = "正文"
    elif region == "header":
        paragraph_iterator = lambda: iter_header_paragraphs(doc)
        region_desc = "页眉"
    elif region == "footer":
        paragraph_iterator = lambda: iter_footer_paragraphs(doc)
        region_desc = "页脚"
    else:
        paragraph_iterator = lambda: iter_all_paragraphs(doc)
        region_desc = "全部"

    # ===== 策略0A：Word自动编号替换（带上下文验证） =====
    if is_list_pattern(old_value):
        def _detect_numbering_type(val):
            v = val.strip().lower()
            if re.match(r'^[ivxlcdm]+\.$', v): return 'roman'
            if re.match(r'^\(\d+\)$', v): return 'paren_digit'
            if re.match(r'^\d+[\.\)]$', v): return 'digit'
            if re.match(r'^\([a-z]\)$', v): return 'paren_letter'
            if re.match(r'^[a-z][\.\)]$', v): return 'letter'
            return 'unknown'

        old_type = _detect_numbering_type(old_value)
        new_type = _detect_numbering_type(new_value)
        is_same_format = (old_type == new_type and old_type != 'unknown')

        if is_same_format:
            # 同格式序号值错误（如 iv. → v.）：需要定位到正确段落单独修改
            ok, strategy = _replace_numbering_with_context(
                doc, old_value, new_value, reason, revision_manager,
                context, anchor_text, paragraph_iterator, region_desc
            )
            if ok:
                return True, strategy
            return False, f"编号替换失败: 同格式序号值错误但无法定位 [{region_desc}]"
        else:
            # 格式变更（如 i. → (1)）
            try:
                from llm.llm_project.replace.numbering_replacer import replace_numbering_in_docx
                success, message = replace_numbering_in_docx(
                    doc, old_value, new_value, context,
                    None, reason
                )
                if success:
                    return True, f"Word自动编号替换: {message}"
                else:
                    print(f"  提示: {old_value} 自动编号替换未成功，尝试手动编号定位")
            except Exception as e:
                print(f"  编号替换失败: {e}")

            ok, strategy = _replace_numbering_with_context(
                doc, old_value, new_value, reason, revision_manager,
                context, anchor_text, paragraph_iterator, region_desc
            )
            if ok:
                return True, strategy

            return False, f"编号替换失败: 未找到匹配的编号 [{region_desc}]"

    # ===== 策略0：预处理特殊情况（带上下文验证，尊重 region） =====
    found, matched_text, strategy = preprocess_special_cases(old_value_original, doc)
    if found:
        # 收集所有包含 old_value 的候选段落并打分（不限于 matched_text）
        preprocess_candidates = []
        for p in paragraph_iterator():
            full_text = "".join(r.text or "" for r in p.runs)
            if old_value in full_text or old_value_original in full_text:
                score = _score_paragraph(full_text, old_value, context, anchor_text, match_level=1)
                preprocess_candidates.append((score, p, full_text))

        if preprocess_candidates:
            preprocess_candidates.sort(key=lambda x: x[0], reverse=True)
            best_score, best_para, best_text = preprocess_candidates[0]

            # 多候选时的安全检查（与策略3-6一致）
            safe = True
            if len(preprocess_candidates) > 1:
                if not context and not anchor_text:
                    safe = False  # 多处出现且无上下文，跳过策略0
                else:
                    second_score = preprocess_candidates[1][0]
                    if best_score > 0 and second_score / best_score > 0.95:
                        safe = False  # 得分太接近，无法区分

            if safe:
                ok = _execute_replace(best_para, old_value, new_value, reason, revision_manager)
                if ok:
                    return True, f"{strategy} [{region_desc}]"
                ok = _execute_replace(best_para, old_value_original, new_value, reason, revision_manager)
                if ok:
                    return True, f"{strategy} [{region_desc}]"

    # ===== 策略1：显式锚点（带候选打分） =====
    if anchor_text:
        anchor_clean = clean_text_thoroughly(anchor_text)
        if anchor_clean:
            anchor_candidates = []
            for p in paragraph_iterator():
                full_text = "".join(r.text or "" for r in p.runs)
                if anchor_clean in clean_text_thoroughly(full_text) and old_value in full_text:
                    score = _score_paragraph(full_text, old_value, context, anchor_text, match_level=1)
                    anchor_candidates.append((score, p))

            if anchor_candidates:
                anchor_candidates.sort(key=lambda x: x[0], reverse=True)
                best_score, best_para = anchor_candidates[0]
                ok = _execute_replace(best_para, old_value, new_value, reason, revision_manager)
                if ok:
                    return True, f"锚点匹配 [{region_desc}]"

    # ===== 策略2：上下文锚点 =====
    if context:
        context_anchor = extract_anchor_with_target(context, old_value, window=60)
        if context_anchor:
            anchor_clean = clean_text_thoroughly(context_anchor)
            for p in paragraph_iterator():
                full_text = "".join(r.text or "" for r in p.runs)
                full_clean = clean_text_thoroughly(full_text)
                if anchor_clean in full_clean and old_value in full_text:
                    ok = _execute_replace(p, old_value, new_value, reason, revision_manager)
                    if ok:
                        return True, f"上下文锚点匹配 [{region_desc}]"

    # ===== 策略3-6：候选打分模式 =====
    old_value_clean = clean_text_thoroughly(old_value)

    strategies = [
        # (策略名, 正则模式, 最低上下文相似度门槛, 最大匹配层级)
        ("严格模式+上下文", build_smart_pattern(old_value, mode="strict"), 0.15, 2),
        ("严格模式", build_smart_pattern(old_value, mode="strict"), 0.0, 2),
        ("平衡模式", build_smart_pattern(old_value, mode="balanced"), 0.0, 3),
        ("宽松模式", build_smart_pattern(old_value, mode="balanced"), 0.0, 5),
    ]

    for strategy_name, pattern, min_similarity, max_level in strategies:
        # 策略3（严格+上下文）要求必须有上下文
        if "上下文" in strategy_name and not context:
            continue

        candidates = []  # [(score, paragraph, match_level, full_text)]

        for p in paragraph_iterator():
            match_level = _try_match_in_paragraph(p, old_value, old_value_clean, pattern)
            if match_level is None or match_level > max_level:
                continue

            full_text = "".join(r.text or "" for r in p.runs)
            score = _score_paragraph(full_text, old_value, context, anchor_text, match_level)
            candidates.append((score, p, match_level, full_text))

        if not candidates:
            continue

        # 按得分降序排序
        candidates.sort(key=lambda x: x[0], reverse=True)
        best_score, best_para, best_level, best_text = candidates[0]

        # ---- 安全检查 ----

        # 检查1：上下文相似度门槛
        if context and min_similarity > 0:
            sim = calculate_context_similarity(
                clean_text_thoroughly(best_text),
                clean_text_thoroughly(context)
            )
            if sim < min_similarity:
                continue

        # 检查2：唯一性安全网
        if len(candidates) > 1 and not context and not anchor_text:
            continue

        # 检查3：多处出现时，最佳和次佳得分差距太小 → 不确定
        if len(candidates) > 1:
            second_score = candidates[1][0]
            if best_score > 0 and second_score / best_score > 0.95:
                continue

        # ---- 执行替换 ----
        ok = _execute_replace(best_para, old_value, new_value, reason, revision_manager)
        if ok:
            detail = f"{strategy_name} [{region_desc}] (得分:{best_score:.2f}, 候选:{len(candidates)})"
            return True, detail

    return False, f"未找到匹配项 (搜索区域: {region_desc})"


# =========================
# 兼容旧接口
# =========================

def replace_and_add_revision_in_paragraph(
        paragraph, pattern, old_value, new_value, reason,
        revision_manager: RevisionManager,
        anchor_pattern=None, context_text=None, similarity_threshold=0.3, region="body"
) -> bool:
    """
    兼容旧接口：在段落中查找并以修订模式替换。
    当提供 context_text 时，会先验证段落与上下文的相似度，不匹配则跳过。
    """
    runs = list(paragraph.runs)
    if not runs:
        return False
    full_text = "".join(r.text or "" for r in runs)

    # 上下文验证
    if context_text:
        sim = calculate_context_similarity(
            clean_text_thoroughly(full_text),
            clean_text_thoroughly(context_text)
        )
        if sim < similarity_threshold:
            return False

    full_text_clean = clean_text_thoroughly(full_text)
    old_value_clean = clean_text_thoroughly(old_value)

    # 精确匹配
    if old_value in full_text:
        return revision_manager.replace_in_paragraph(paragraph, old_value, new_value, reason=reason)

    # 清洗后匹配
    if old_value_clean in full_text_clean and old_value in full_text:
        return revision_manager.replace_in_paragraph(paragraph, old_value, new_value, reason=reason)

    # 正则匹配
    if pattern:
        try:
            if re.search(pattern, full_text_clean, flags=re.IGNORECASE | re.DOTALL):
                if old_value in full_text:
                    return revision_manager.replace_in_paragraph(paragraph, old_value, new_value, reason=reason)
        except re.error:
            pass

    # 忽略空格
    old_no_space = old_value.replace(' ', '')
    for run in runs:
        run_text = run.text or ""
        run_no_space = run_text.replace(' ', '')
        if old_no_space and len(run_text.strip()) < 50 and old_no_space == run_no_space:
            revision_manager.replace_run_text(run, new_value, reason=reason)
            return True

    return False
