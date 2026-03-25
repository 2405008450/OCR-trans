from __future__ import annotations

from typing import Callable, Iterable, Tuple

from docx import Document

from llm.llm_project.revise.revision import RevisionManager
from llm.llm_project.replace.replace_clean import (
    calculate_context_similarity,
    clean_text_thoroughly,
    extract_anchor_with_target,
    is_list_pattern,
    iter_all_paragraphs,
    iter_body_paragraphs,
    iter_footer_paragraphs,
    iter_header_paragraphs,
    preprocess_special_cases,
)


def _get_iterator(doc: Document, region: str) -> tuple[Callable[[], Iterable], str]:
    if region == 'body':
        return lambda: iter_body_paragraphs(doc), '正文'
    if region == 'header':
        return lambda: iter_header_paragraphs(doc), '页眉'
    if region == 'footer':
        return lambda: iter_footer_paragraphs(doc), '页脚'
    return lambda: iter_all_paragraphs(doc), '全部'


def _replace_in_paragraph(paragraph, old_value: str, new_value: str, reason: str, revision_manager: RevisionManager) -> bool:
    runs = list(paragraph.runs)
    if not runs:
        return False
    full_text = ''.join(run.text or '' for run in runs)
    if old_value not in full_text:
        return False
    return revision_manager.replace_in_paragraph(paragraph, old_value, new_value, reason=reason)


def _find_best_context_match(paragraph_iterator: Callable[[], Iterable], old_value: str, context: str):
    context_clean = clean_text_thoroughly(context or '')
    best = None
    best_score = 0.0
    for paragraph in paragraph_iterator():
        full_text = ''.join(run.text or '' for run in paragraph.runs)
        if old_value not in full_text:
            continue
        score = calculate_context_similarity(clean_text_thoroughly(full_text), context_clean)
        if score > best_score:
            best = paragraph
            best_score = score
    return best, best_score


def replace_and_revise_in_docx(
    doc: Document,
    old_value: str,
    new_value: str,
    reason: str,
    revision_manager: RevisionManager,
    context: str = '',
    anchor_text: str = '',
    region: str = 'all',
) -> Tuple[bool, str]:
    old_value_original = (old_value or '').strip()
    new_value = clean_text_thoroughly(new_value or '').strip()
    context = clean_text_thoroughly(context or '')
    anchor_text = clean_text_thoroughly(anchor_text or '')
    reason = (reason or '数值或术语不一致').strip()

    if not old_value_original or not new_value:
        return False, '数据缺失'

    paragraph_iterator, region_desc = _get_iterator(doc, region)

    if is_list_pattern(old_value_original):
        try:
            from llm.llm_project.replace.numbering_replacer import replace_numbering_in_docx

            success, message = replace_numbering_in_docx(
                doc, old_value_original, new_value, context, None, reason
            )
            if success:
                return True, f'Word自动编号替换: {message}'
        except Exception:
            pass

    found, matched_text, special_strategy = preprocess_special_cases(old_value_original, doc)
    if found:
        for paragraph in paragraph_iterator():
            full_text = ''.join(run.text or '' for run in paragraph.runs)
            if full_text == matched_text:
                if _replace_in_paragraph(paragraph, old_value_original, new_value, reason, revision_manager):
                    return True, f'{special_strategy} [{region_desc}]'

    if anchor_text:
        for paragraph in paragraph_iterator():
            full_text = ''.join(run.text or '' for run in paragraph.runs)
            if old_value_original in full_text and anchor_text in clean_text_thoroughly(full_text):
                if _replace_in_paragraph(paragraph, old_value_original, new_value, reason, revision_manager):
                    return True, f'锚点匹配 [{region_desc}]'

    if context:
        context_anchor = extract_anchor_with_target(context, old_value_original, window=60)
        if context_anchor:
            context_anchor = clean_text_thoroughly(context_anchor)
            for paragraph in paragraph_iterator():
                full_text = ''.join(run.text or '' for run in paragraph.runs)
                if old_value_original in full_text and context_anchor in clean_text_thoroughly(full_text):
                    if _replace_in_paragraph(paragraph, old_value_original, new_value, reason, revision_manager):
                        return True, f'上下文锚点匹配 [{region_desc}]'

        best_paragraph, best_score = _find_best_context_match(paragraph_iterator, old_value_original, context)
        if best_paragraph is not None and best_score >= 0.15:
            if _replace_in_paragraph(best_paragraph, old_value_original, new_value, reason, revision_manager):
                return True, f'上下文定位 [{region_desc}] (得分:{best_score:.2f})'

    for paragraph in paragraph_iterator():
        if _replace_in_paragraph(paragraph, old_value_original, new_value, reason, revision_manager):
            return True, f'精确匹配 [{region_desc}]'

    return False, f'未找到匹配项 [{region_desc}]'
