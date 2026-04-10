"""
基于修订（Track Changes）的替换模块（优化版）

优化点：
1. _context_similarity — 包含关系优先 + Jaccard 兜底，短词组更准确
2. _try_match_in_paragraph — 去掉层2-5 的二次精确检查（死代码修复）
3. _execute_replace — 增加清洗后匹配 + 忽略空格 fallback
4. 策略0（preprocess_special_cases）删除 — 功能被策略3-6 完全覆盖，且有 region 泄漏问题
5. 策略1/2 锚点匹配 — 去掉 old_value in full_text 的过严限制
6. 脚注检测后移到策略7 — 避免正文和脚注都有同一文本时走错路径
7. 策略3-6 安全检查改用 _context_similarity
8. 编号静态化后空格归一化 — 中文括号编号与后续文本之间的空格差异不再导致匹配失败
9. 中文后缀剥离 — old_value 含中文后缀（年/月/日/万/亿等）时自动剥离后重试
"""
import re
from typing import Tuple, List, Optional
from docx import Document
from lxml import etree

from revise.revision import RevisionManager
from replace.replace_clean import (
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
# 辅助：编号静态化后的空格归一化
# =========================

# 匹配中文括号编号，如 （一）、（二）、(一)、(二) 等
_CN_PAREN_NUM_RE = re.compile(
    r'[（\(][一二三四五六七八九十百千万零\d]+[）\)]\s*'
)

def _normalize_numbering_spaces(text: str) -> str:
    """
    归一化中文括号编号后的空格：统一为零个空格。
    这样 '（二） ii.' 和 '（二）ii.' 都会变成 '（二）ii.'，
    clean 之后都变成 '(二)ii.'，保证匹配一致。
    """
    return _CN_PAREN_NUM_RE.sub(lambda m: m.group(0).rstrip(), text)


def _find_actual_text_in_paragraph(full_text: str, old_value: str) -> str:
    """
    在原始段落文本中，用宽松方式定位 old_value 对应的实际子串。
    核心思路：把 old_value 的每个"有意义字符"转成正则，字符之间允许可选空格，
    全角/半角括号互通，直接在原始文本上搜。
    """
    # 全角↔半角映射
    fw_to_hw = {'（': '(', '）': ')', '，': ',', '。': '.', '：': ':', '；': ';',
                '！': '!', '？': '?', '【': '[', '】': ']'}
    hw_to_fw = {v: k for k, v in fw_to_hw.items()}

    pieces = []
    for ch in old_value:
        if ch in (' ', '\t', '\n', '\r', '\u3000', '\u00a0'):
            # 空白 → 可选空格（0或多个）
            if pieces and pieces[-1] != r'\s*':
                pieces.append(r'\s*')
        elif ch in fw_to_hw:
            # 全角符号 → 匹配全角或半角
            hw = fw_to_hw[ch]
            pieces.append(f'[{re.escape(ch)}{re.escape(hw)}]')
        elif ch in hw_to_fw:
            # 半角符号 → 匹配半角或全角
            fw = hw_to_fw[ch]
            pieces.append(f'[{re.escape(ch)}{re.escape(fw)}]')
        else:
            pieces.append(re.escape(ch))

    pattern = ''.join(pieces)
    if not pattern:
        return ""

    m = re.search(pattern, full_text, re.IGNORECASE)
    return m.group(0) if m else ""


# 中文后缀（年/月/日/万/亿/元/条/章/节/款/项 等），在英文译文中不存在
_CN_SUFFIX_RE = re.compile(r'[年月日万亿元条章节款项个次份]+$')

def _strip_cn_suffix(value: str) -> str:
    """剥离 old_value 末尾的中文后缀，如 '2024年' → '2024'"""
    return _CN_SUFFIX_RE.sub('', value).strip()


# =========================
# 辅助：混合相似度
# =========================

def _context_similarity(para_text: str, context: str) -> float:
    """
    混合相似度：包含关系优先，Jaccard + 覆盖率兜底。
    对短词组（如 "Chief Compliance Officer"）比纯 Jaccard 更可靠。
    """
    if not para_text or not context:
        return 0.0

    p = para_text.lower().strip()
    c = context.lower().strip()

    # 上下文被段落包含 → 段落涵盖了完整上下文，高置信
    if c and c in p:
        return 0.9

    # 段落被上下文包含 → 需要根据长度比打折
    # 例如段落只有 "2024"(4字符) 而上下文有70字符，不能给高分
    if p and p in c:
        length_ratio = len(p) / len(c) if len(c) > 0 else 0
        # 长度比越高说明段落越接近上下文全文，越可信
        # 比如段落占上下文80%以上 → 0.85，占5% → 很低
        return 0.3 + 0.55 * length_ratio  # 范围 ~0.3 ~ 0.85

    # 部分包含：上下文前60字符在段落中
    c_head = c[:60]
    if c_head and c_head in p:
        return 0.75

    # Jaccard + 覆盖率
    words_p = set(p.split())
    words_c = set(c.split())
    if not words_p or not words_c:
        return 0.0
    inter = len(words_p & words_c)
    union = len(words_p | words_c)
    jaccard = inter / union if union else 0.0
    coverage = inter / len(words_c) if words_c else 0.0

    return max(jaccard, coverage * 0.8)


# =========================
# 辅助：候选段落打分
# =========================

def _score_paragraph(full_text: str, old_value: str, context: str,
                     anchor_text: str, match_level: int) -> float:
    """对候选段落打分，分数越高越可能是正确的替换位置。"""
    score = 0.0
    full_clean = clean_text_thoroughly(full_text)

    # 维度1：上下文相似度（权重最高，0 ~ 0.5）
    if context:
        sim = _context_similarity(full_clean, clean_text_thoroughly(context))
        score += sim * 0.5

    # 维度2：锚点命中（0 或 0.25）
    if anchor_text:
        anchor_clean = clean_text_thoroughly(anchor_text)
        anchor_norm = _normalize_numbering_spaces(anchor_clean)
        anchor_stripped = _strip_cn_suffix(anchor_clean)
        full_norm = _normalize_numbering_spaces(full_clean)
        if anchor_clean and (anchor_clean in full_clean
                             or anchor_norm in full_norm
                             or (anchor_stripped and anchor_stripped != anchor_clean
                                 and anchor_stripped in full_clean)):
            score += 0.25

    # 维度3：匹配精确度（0.05 ~ 0.15）
    level_scores = {1: 0.15, 2: 0.12, 3: 0.10, 4: 0.07, 5: 0.05}
    score += level_scores.get(match_level, 0.05)

    # 维度4：old_value 占段落比例（0 ~ 0.1）
    if full_text:
        ratio = len(old_value) / len(full_text)
        score += min(ratio, 1.0) * 0.1

    return score


# =========================
# 1) 修订版替换核心函数
# =========================

def apply_revision(paragraph, runs, old_value, new_value, reason,
                   revision_manager: RevisionManager, match_type="正则", region="body"):
    """执行实际替换（修订模式）。"""
    if not runs:
        return False
    full_text = "".join(r.text or "" for r in runs)
    if old_value in full_text:
        return revision_manager.replace_in_paragraph(paragraph, old_value, new_value, reason=reason)
    return False


def _try_match_in_paragraph(paragraph, old_value: str, old_value_clean: str,
                            pattern: str) -> Optional[int]:
    """尝试在段落中匹配 old_value，返回匹配层级（1-6），未匹配返回 None。

    层级越低越精确：
      1 = 精确包含
      2 = 清洗后包含（原文有多余空格/特殊字符）
      2.5 = 编号归一化后包含（中文括号编号空格差异）
      2.7 = 中文后缀剥离后包含（如 '2024年' → '2024'）
      3 = 正则匹配
      4 = 模糊匹配 / 指纹匹配
      5 = 无空格匹配

    各层不再要求 old_value in full_text，由 _execute_replace 负责实际替换。
    """
    runs = list(paragraph.runs)
    if not runs:
        return None
    full_text = "".join(r.text or "" for r in runs)
    if not full_text.strip():
        return None

    if old_value in full_text:
        return 1

    full_text_clean = clean_text_thoroughly(full_text)

    if old_value_clean and old_value_clean in full_text_clean:
        return 2

    # 编号归一化匹配（消除中文括号编号后的空格差异）
    full_norm = _normalize_numbering_spaces(full_text_clean)
    old_norm = _normalize_numbering_spaces(old_value_clean)
    if old_norm and old_norm != old_value_clean and old_norm in full_norm:
        return 2

    # 中文后缀剥离匹配（如 '2024年' 在英文文本中只有 '2024'）
    old_stripped = _strip_cn_suffix(old_value_clean)
    if old_stripped and old_stripped != old_value_clean and old_stripped in full_text_clean:
        return 2

    if pattern:
        try:
            if re.search(pattern, full_text_clean, flags=re.IGNORECASE | re.DOTALL):
                return 3
        except re.error:
            pass

    if is_fuzzy_match(full_text_clean, old_value_clean, threshold=0.85):
        return 4

    fingerprint_old = get_alphanumeric_fingerprint(old_value)
    fingerprint_full = get_alphanumeric_fingerprint(full_text)
    if len(fingerprint_old) >= 3 and fingerprint_old in fingerprint_full:
        return 4

    full_no_space = full_text_clean.replace(' ', '')
    old_no_space = old_value_clean.replace(' ', '')
    if old_no_space and old_no_space in full_no_space:
        return 5

    return None


def _execute_replace(paragraph, old_value: str, new_value: str, reason: str,
                     revision_manager: RevisionManager) -> bool:
    """在段落中执行实际替换，依次尝试精确、清洗后、无空格、编号归一化、中文后缀剥离等方式。"""
    runs = list(paragraph.runs)
    if not runs:
        return False
    full_text = "".join(r.text or "" for r in runs)

    # 精确匹配
    if old_value in full_text:
        return revision_manager.replace_in_paragraph(paragraph, old_value, new_value, reason=reason)

    # 清洗后匹配 — 用正则在原始文本中定位实际片段
    full_clean = clean_text_thoroughly(full_text)
    old_clean = clean_text_thoroughly(old_value)
    if old_clean and old_clean in full_clean:
        pattern = build_smart_pattern(old_value, mode="balanced")
        if pattern:
            m = re.search(pattern, full_text, re.IGNORECASE)
            if m:
                actual_old = m.group(0)
                return revision_manager.replace_in_paragraph(paragraph, actual_old, new_value, reason=reason)

    # 编号归一化匹配 — 消除中文括号编号后空格差异
    # 直接在原始文本上用宽松正则搜（全角半角互通、空格可选）
    actual_old = _find_actual_text_in_paragraph(full_text, old_value)
    if actual_old and actual_old != old_value:
        return revision_manager.replace_in_paragraph(paragraph, actual_old, new_value, reason=reason)

    # 中文后缀剥离匹配 — 如 '2024年' 在英文文本中只有 '2024'
    old_stripped = _strip_cn_suffix(old_value)
    if old_stripped and old_stripped != old_value and old_stripped in full_text:
        return revision_manager.replace_in_paragraph(paragraph, old_stripped, new_value, reason=reason)
    # 也尝试清洗后剥离
    old_clean_stripped = _strip_cn_suffix(old_clean)
    if old_clean_stripped and old_clean_stripped != old_clean and old_clean_stripped in full_clean:
        pattern = build_smart_pattern(old_stripped, mode="balanced")
        if pattern:
            m = re.search(pattern, full_text, re.IGNORECASE)
            if m:
                actual_old = m.group(0)
                return revision_manager.replace_in_paragraph(paragraph, actual_old, new_value, reason=reason)

    # 无空格匹配（单 run）
    old_no_space = old_value.replace(' ', '')
    for run in runs:
        run_text = run.text or ""
        if old_no_space and old_no_space == run_text.replace(' ', ''):
            revision_manager.replace_run_text(run, new_value, reason=reason)
            return True

    return False


# =========================
# 辅助：中文自动编号替换
# =========================

def _replace_cn_auto_numbering(doc, old_value, new_value, reason,
                                revision_manager, context, anchor_text,
                                paragraph_iterator, region_desc, doc_path=None):
    """替换 Word 自动生成的中文编号（如 第二章、第五条、（一）等）。"""
    from docx.oxml.ns import qn as _qn
    from zipfile import ZipFile

    W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    NS = {"w": W_NS}

    context_clean = clean_text_thoroughly(context or "")
    anchor_clean = clean_text_thoroughly(anchor_text or "")
    old_clean = old_value.strip()

    numbering_map = {}
    abstract_num_map = {}
    level_counters = {}

    def _to_chinese(num):
        cn = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"]
        units = ["", "十", "百", "千", "万"]
        if num == 0:
            return cn[0]
        result = ""
        ui = 0
        n = num
        while n > 0:
            d = n % 10
            if d != 0:
                result = cn[d] + units[ui] + result
            elif result and result[0] != "零":
                result = cn[0] + result
            n //= 10
            ui += 1
        if result.startswith("一十"):
            result = result[1:]
        return result.rstrip("零")

    def _format_number(num, fmt):
        if fmt == "decimal":
            return str(num)
        if fmt in ("chineseCountingThousand", "chineseCounting",
                   "japaneseCounting", "japaneseDigitalTenThousand",
                   "ideographTraditional"):
            return _to_chinese(num)
        if fmt == "upperRoman":
            from parsers.word.body_extractor import NumberingSystem
            return NumberingSystem._to_roman(num).upper()
        if fmt == "lowerRoman":
            from parsers.word.body_extractor import NumberingSystem
            return NumberingSystem._to_roman(num).lower()
        if fmt == "upperLetter":
            return chr(64 + num) if 1 <= num <= 26 else str(num)
        if fmt == "lowerLetter":
            return chr(96 + num) if 1 <= num <= 26 else str(num)
        return str(num)

    try:
        zip_path = doc_path
        if not zip_path:
            if hasattr(doc, '_part') and hasattr(doc._part, 'package'):
                try:
                    zip_path = doc._part.package._pkg_file
                except Exception:
                    pass
        if not zip_path:
            return False, ""

        with ZipFile(zip_path, "r") as zf:
            if "word/numbering.xml" not in zf.namelist():
                return False, ""
            with zf.open("word/numbering.xml") as f:
                tree = etree.parse(f)

            for abs_num in tree.findall(".//w:abstractNum", NS):
                abs_id = abs_num.get(f"{{{W_NS}}}abstractNumId")
                abstract_num_map[abs_id] = {}
                for lvl in abs_num.findall(".//w:lvl", NS):
                    ilvl = lvl.get(f"{{{W_NS}}}ilvl")
                    nf = lvl.find(".//w:numFmt", NS)
                    lt = lvl.find(".//w:lvlText", NS)
                    st = lvl.find(".//w:start", NS)
                    abstract_num_map[abs_id][ilvl] = {
                        "format": nf.get(f"{{{W_NS}}}val") if nf is not None else "decimal",
                        "text": lt.get(f"{{{W_NS}}}val") if lt is not None else "%1.",
                        "start": int(st.get(f"{{{W_NS}}}val", "1")) if st is not None else 1,
                    }

            for num in tree.findall(".//w:num", NS):
                nid = num.get(f"{{{W_NS}}}numId")
                abs_elem = num.find(".//w:abstractNumId", NS)
                if abs_elem is None:
                    continue
                abs_id = abs_elem.get(f"{{{W_NS}}}val")
                if abs_id not in abstract_num_map:
                    continue
                lm = {k: v.copy() for k, v in abstract_num_map[abs_id].items()}
                for override in num.findall(f"{{{W_NS}}}lvlOverride"):
                    ov_ilvl = override.get(f"{{{W_NS}}}ilvl")
                    if ov_ilvl is None:
                        continue
                    lvl = override.find(f"{{{W_NS}}}lvl")
                    if lvl is not None:
                        oi = lm.get(ov_ilvl, {}).copy()
                        nf = lvl.find(".//w:numFmt", NS)
                        lt = lvl.find(".//w:lvlText", NS)
                        st = lvl.find(".//w:start", NS)
                        if nf is not None:
                            oi["format"] = nf.get(f"{{{W_NS}}}val")
                        if lt is not None:
                            oi["text"] = lt.get(f"{{{W_NS}}}val")
                        if st is not None:
                            oi["start"] = int(st.get(f"{{{W_NS}}}val", "1"))
                        lm[ov_ilvl] = oi
                    so = override.find(f"{{{W_NS}}}startOverride")
                    if so is not None and ov_ilvl in lm:
                        lm[ov_ilvl]["start"] = int(so.get(f"{{{W_NS}}}val", "1"))
                numbering_map[nid] = lm
    except Exception as e:
        print(f"    [中文编号] 加载编号定义失败: {e}")
        return False, ""

    def _get_para_numbering_text(p_elem):
        pPr = p_elem.find(f"{{{W_NS}}}pPr")
        if pPr is None:
            return None, None, None
        numPr = pPr.find(f"{{{W_NS}}}numPr")
        if numPr is None:
            return None, None, None
        nid_elem = numPr.find(f"{{{W_NS}}}numId")
        ilvl_elem = numPr.find(f"{{{W_NS}}}ilvl")
        if nid_elem is None:
            return None, None, None
        nid = nid_elem.get(f"{{{W_NS}}}val")
        ilvl = ilvl_elem.get(f"{{{W_NS}}}val") if ilvl_elem is not None else "0"
        if nid == "0" or nid not in numbering_map or ilvl not in numbering_map[nid]:
            return None, None, None

        li = numbering_map[nid][ilvl]
        ck = (nid, ilvl)
        if ck not in level_counters:
            level_counters[ck] = li["start"]
        else:
            level_counters[ck] += 1

        ilvl_int = int(ilvl)
        for oi_str in numbering_map[nid]:
            if int(oi_str) > ilvl_int:
                ok = (nid, oi_str)
                if ok in level_counters:
                    del level_counters[ok]

        tmpl = li["text"]
        for li_idx in range(ilvl_int + 1):
            ph = f"%{li_idx + 1}"
            if ph not in tmpl:
                continue
            ls = str(li_idx)
            lk = (nid, ls)
            if ls in numbering_map[nid] and lk in level_counters:
                linfo = numbering_map[nid][ls]
                formatted = _format_number(level_counters[lk], linfo["format"])
                tmpl = tmpl.replace(ph, formatted)

        return tmpl.strip(), nid, ilvl

    candidates = []
    for p in paragraph_iterator():
        num_text, nid, ilvl = _get_para_numbering_text(p._element)
        if num_text is None:
            continue
        if num_text != old_clean and old_clean not in num_text:
            continue

        para_text = (p.text or "").strip()
        para_clean = clean_text_thoroughly(para_text)

        score = 0.5
        if context_clean:
            if para_clean and para_clean.lower() in context_clean.lower():
                score += 0.4
            elif context_clean.lower() in para_clean.lower():
                score += 0.3
            else:
                sim = calculate_context_similarity(para_clean, context_clean)
                score += sim * 0.3
        if anchor_clean:
            if para_clean and para_clean.lower() in anchor_clean.lower():
                score += 0.15

        candidates.append((score, p, num_text, nid, ilvl, para_text))

    if not candidates:
        return False, ""

    candidates.sort(key=lambda x: x[0], reverse=True)
    best_score, best_para, best_num_text, best_nid, best_ilvl, best_text = candidates[0]

    if len(candidates) > 1:
        second_score = candidates[1][0]
        if not context_clean and not anchor_clean:
            print(f"    [中文编号] '{old_value}' 匹配 {len(candidates)} 处且无上下文，拒绝替换")
            return False, ""
        if best_score > 0 and second_score / best_score > 0.95:
            print(f"    [中文编号] 前两个候选得分接近 ({best_score:.3f} vs {second_score:.3f})")
            return False, ""

    print(f"    [中文编号] 定位到段落: 编号='{best_num_text}' 文本='{best_text[:50]}' (得分:{best_score:.3f})")

    try:
        pPr = best_para._element.pPr
        if pPr is not None:
            numPr = pPr.find(f"{{{W_NS}}}numPr")
            if numPr is not None:
                pPr.remove(numPr)

        new_text_with_space = (new_value.strip() + "  ") if (best_num_text and best_num_text != old_clean) else (new_value.strip() + " ")
        revision_manager.insert_text_at_beginning(best_para, new_text_with_space, reason)

        return True, (f"中文自动编号替换 '{best_num_text}'→'{new_value}' "
                      f"[{region_desc}] (得分:{best_score:.2f}, 候选:{len(candidates)})")
    except Exception as e:
        print(f"    [中文编号] 替换执行失败: {e}")
        return False, ""


# =========================
# 辅助：编号段落上下文定位替换
# =========================

def _replace_numbering_with_context(doc, old_value, new_value, reason,
                                     revision_manager, context, anchor_text,
                                     paragraph_iterator, region_desc):
    """对编号类值（如 iv. → v.）进行上下文验证后替换。"""
    from docx.oxml.ns import qn as _qn

    context_clean = clean_text_thoroughly(context or "")
    anchor_clean = clean_text_thoroughly(anchor_text or "")

    def _strip_numbering_prefix(text):
        stripped = re.sub(r'^[ivxlcdm]+\.\s*', '', text.strip(), flags=re.IGNORECASE)
        stripped = re.sub(r'^\(\d+\)\s*', '', stripped)
        stripped = re.sub(r'^\d+[\.\)]\s*', '', stripped)
        stripped = re.sub(r'^\([a-z]\)\s*', '', stripped, flags=re.IGNORECASE)
        return stripped.strip()

    context_content = _strip_numbering_prefix(context_clean)

    def _numbering_to_int(s):
        s = s.strip().lower()
        m = re.match(r'^([ivxlcdm]+)\.$', s)
        if m:
            roman = m.group(1)
            roman_map = {'i': 1, 'v': 5, 'x': 10, 'l': 50, 'c': 100, 'd': 500, 'm': 1000}
            result, prev = 0, 0
            for ch in reversed(roman):
                val = roman_map.get(ch, 0)
                result += val if val >= prev else -val
                prev = val
            return result
        m = re.match(r'^\(?(\d+)[\.\)]$', s)
        if m:
            return int(m.group(1))
        m = re.match(r'^\(?([a-z])[\.\)]$', s)
        if m:
            return ord(m.group(1)) - ord('a') + 1
        return 0

    new_num_int = _numbering_to_int(new_value)
    if new_num_int <= 0:
        return False, ""

    candidates = []
    for p in paragraph_iterator():
        para_text = (p.text or "").strip()
        if not para_text:
            continue

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
        except Exception:
            pass

        full_text = "".join(r.text or "" for r in p.runs)
        has_manual_numbering = old_value in full_text

        if not has_auto_numbering and not has_manual_numbering:
            continue

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

    if len(candidates) > 1:
        second_score = candidates[1][0]
        if not context_clean and not anchor_clean:
            print(f"    [编号替换] '{old_value}' 出现 {len(candidates)} 处且无上下文，拒绝替换")
            return False, ""
        if best_score > 0 and second_score / best_score > 0.95:
            print(f"    [编号替换] 前两个候选得分接近 ({best_score:.3f} vs {second_score:.3f})")
            return False, ""
        if best_score < 0.1:
            print(f"    [编号替换] 最佳候选得分过低 ({best_score:.3f})")
            return False, ""

    print(f"    [编号替换] 定位到段落: '{best_text[:60]}' (得分:{best_score:.3f}, 候选:{len(candidates)})")

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
                return False, ""

            max_num_id = max(
                (int(e.get(_qn('w:numId'))) for e in numbering_xml.findall(_qn('w:num'))
                 if e.get(_qn('w:numId'), '').isdigit()),
                default=0
            )
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
            pPr.numPr.numId.set(_qn('w:val'), new_num_id)

            return True, (f"编号上下文定位(自动编号 新numId={new_num_id} start={new_num_int}) "
                          f"[{region_desc}] (得分:{best_score:.2f}, 候选:{len(candidates)})")
        except Exception as e:
            print(f"    [编号替换] 修改自动编号失败: {e}")

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
        region: str = "all",
        doc_path: str = None
) -> Tuple[bool, str]:
    """
    多策略执行替换（修订模式），带上下文验证和候选打分。

    策略顺序：
      0A  编号特判（仅未静态化时）
      0B  中文自动编号警告
      1   显式锚点匹配
      2   上下文锚点匹配
      3-6 候选打分（严格/平衡/宽松）
      7   脚注/尾注（正文未找到时才检查）
    """
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

    # ===== 策略0A：Word自动编号替换（仅未静态化时触发） =====
    if is_list_pattern(old_value) and not getattr(doc, '_numbering_staticized', False):
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
            ok, strategy = _replace_numbering_with_context(
                doc, old_value, new_value, reason, revision_manager,
                context, anchor_text, paragraph_iterator, region_desc
            )
            if ok:
                return True, strategy
            # 失败则继续走后续策略
        else:
            try:
                from replace.numbering_replacer import replace_numbering_in_docx
                success, message = replace_numbering_in_docx(
                    doc, old_value, new_value, context, None, reason
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
            # 失败则继续走后续策略

    # ===== 策略0B：中文自动编号警告 =====
    _cn_num_pattern = re.compile(
        r'^(第[一二三四五六七八九十百千万零]+[章条节篇部款项]|'
        r'[（\(][一二三四五六七八九十百千万零]+[）\)]|'
        r'[一二三四五六七八九十百千万零]+[、．.])'
    )
    if _cn_num_pattern.match(old_value.strip()):
        if not getattr(doc, '_numbering_staticized', False):
            print(f"    [警告] '{old_value}' 疑似自动编号，但文档未静态化，可能无法替换")

    # ===== 策略1：显式锚点（锚点命中即进入候选，不要求精确包含 old_value） =====
    if anchor_text:
        anchor_clean = clean_text_thoroughly(anchor_text)
        anchor_norm = _normalize_numbering_spaces(anchor_clean)
        anchor_stripped = _strip_cn_suffix(anchor_clean)
        context_clean = clean_text_thoroughly(context) if context else ""
        if anchor_clean:
            anchor_candidates = []
            for p in paragraph_iterator():
                full_text = "".join(r.text or "" for r in p.runs)
                full_clean = clean_text_thoroughly(full_text)
                full_norm = _normalize_numbering_spaces(full_clean)
                # 尝试原始清洗匹配 + 编号归一化匹配 + 中文后缀剥离匹配
                if (anchor_clean in full_clean
                        or anchor_norm in full_norm
                        or (anchor_stripped and anchor_stripped != anchor_clean
                            and anchor_stripped in full_clean)):
                    score = _score_paragraph(full_text, old_value, context, anchor_text, match_level=1)
                    # 上下文验证：有上下文时计算相似度，用于过滤和加权
                    ctx_sim = 0.0
                    if context_clean:
                        ctx_sim = _context_similarity(full_clean, context_clean)
                        # 短锚点（<15字符）且上下文相似度极低 → 很可能是误命中，跳过
                        if ctx_sim < 0.1 and len(anchor_clean) < 15:
                            continue
                    anchor_candidates.append((score, ctx_sim, p, full_text))

            if anchor_candidates:
                # 多候选时：短锚点优先按上下文相似度排序，长锚点按综合得分排序
                if len(anchor_candidates) > 1 and len(anchor_clean) < 20 and context_clean:
                    anchor_candidates.sort(key=lambda x: (x[1], x[0]), reverse=True)
                else:
                    anchor_candidates.sort(key=lambda x: x[0], reverse=True)
                # 遍历候选列表，尝试替换直到成功（避免最高分段落替换失败就放弃）
                for _, _, cand_para, cand_text in anchor_candidates:
                    ok = _execute_replace(cand_para, old_value, new_value, reason, revision_manager)
                    if ok:
                        return True, f"锚点匹配 [{region_desc}]"

    # ===== 策略2：上下文锚点（从 context 提取子串定位段落） =====
    if context:
        context_anchor = extract_anchor_with_target(context, old_value, window=60)
        if context_anchor:
            anchor_clean = clean_text_thoroughly(context_anchor)
            anchor_norm = _normalize_numbering_spaces(anchor_clean)
            ctx_candidates = []
            for p in paragraph_iterator():
                full_text = "".join(r.text or "" for r in p.runs)
                full_clean = clean_text_thoroughly(full_text)
                full_norm = _normalize_numbering_spaces(full_clean)
                if anchor_clean in full_clean or anchor_norm in full_norm:
                    score = _score_paragraph(full_text, old_value, context, anchor_text or "", match_level=1)
                    ctx_candidates.append((score, p))
            # 按得分排序，依次尝试替换
            ctx_candidates.sort(key=lambda x: x[0], reverse=True)
            for _, cand_para in ctx_candidates:
                ok = _execute_replace(cand_para, old_value, new_value, reason, revision_manager)
                if ok:
                    return True, f"上下文锚点匹配 [{region_desc}]"

    # ===== 策略3-6：候选打分模式 =====
    old_value_clean = clean_text_thoroughly(old_value)

    strategies = [
        # (策略名, 正则模式, 最低上下文相似度门槛, 最大匹配层级)
        ("严格模式+上下文", build_smart_pattern(old_value, mode="strict"), 0.15, 2),
        ("严格模式",        build_smart_pattern(old_value, mode="strict"), 0.0,  2),
        ("平衡模式",        build_smart_pattern(old_value, mode="balanced"), 0.0, 3),
        ("宽松模式",        build_smart_pattern(old_value, mode="balanced"), 0.0, 5),
    ]

    for strategy_name, pattern, min_similarity, max_level in strategies:
        if "上下文" in strategy_name and not context:
            continue

        candidates = []

        for p in paragraph_iterator():
            match_level = _try_match_in_paragraph(p, old_value, old_value_clean, pattern)
            if match_level is None or match_level > max_level:
                continue
            full_text = "".join(r.text or "" for r in p.runs)
            score = _score_paragraph(full_text, old_value, context, anchor_text, match_level)
            candidates.append((score, p, match_level, full_text))

        if not candidates:
            continue

        candidates.sort(key=lambda x: x[0], reverse=True)
        best_score, best_para, best_level, best_text = candidates[0]

        # 检查1：上下文相似度门槛（使用改进的 _context_similarity）
        if context and min_similarity > 0:
            sim = _context_similarity(
                clean_text_thoroughly(best_text),
                clean_text_thoroughly(context)
            )
            if sim < min_similarity:
                continue

        # 检查2：唯一性安全网（多处出现且无上下文/锚点时拒绝）
        if len(candidates) > 1 and not context and not anchor_text:
            continue

        # 检查3：最佳和次佳得分差距太小 → 不确定
        if len(candidates) > 1:
            second_score = candidates[1][0]
            if best_score > 0 and second_score / best_score > 0.95:
                continue

        ok = _execute_replace(best_para, old_value, new_value, reason, revision_manager)
        if ok:
            return True, f"{strategy_name} [{region_desc}] (得分:{best_score:.2f}, 候选:{len(candidates)})"

    # ===== 策略7：脚注/尾注（正文未找到时才检查，避免走错路径） =====
    if region in ("body", "all"):
        try:
            from replace.footnote_replacer import check_text_in_footnotes

            footnote_doc_path = doc_path
            if not footnote_doc_path:
                if hasattr(doc, '_part') and hasattr(doc._part, 'package'):
                    try:
                        footnote_doc_path = doc._part.package._pkg_file
                    except Exception:
                        pass

            if footnote_doc_path and check_text_in_footnotes(footnote_doc_path, old_value):
                print(f"    [脚注检测] 文本在脚注/尾注中，将在保存后替换")
                if not hasattr(doc, '_pending_footnote_replacements'):
                    doc._pending_footnote_replacements = []
                doc._pending_footnote_replacements.append(
                    (footnote_doc_path, old_value, new_value, reason)
                )
                return True, f"脚注/尾注替换（待保存后执行） [{region}]"
        except Exception:
            pass

    return False, f"未找到匹配项 (搜索区域: {region_desc})"


# =========================
# 兼容旧接口
# =========================

def replace_and_add_revision_in_paragraph(
        paragraph, pattern, old_value, new_value, reason,
        revision_manager: RevisionManager,
        anchor_pattern=None, context_text=None, similarity_threshold=0.3, region="body"
) -> bool:
    """兼容旧接口：在段落中查找并以修订模式替换。"""
    runs = list(paragraph.runs)
    if not runs:
        return False
    full_text = "".join(r.text or "" for r in runs)

    if context_text:
        sim = _context_similarity(
            clean_text_thoroughly(full_text),
            clean_text_thoroughly(context_text)
        )
        if sim < similarity_threshold:
            return False

    return _execute_replace(paragraph, old_value, new_value, reason, revision_manager)


def flush_footnote_replacements(doc, save_path: str) -> int:
    """
    在 doc.save() 之后执行所有挂起的脚注替换任务。

    用法：
        doc.save(path)
        flush_footnote_replacements(doc, path)
    """
    pending = getattr(doc, '_pending_footnote_replacements', [])
    if not pending:
        return 0

    from replace.footnote_replacer import replace_in_footnotes_xml

    count = 0
    for footnote_doc_path, old_value, new_value, reason in pending:
        try:
            target_path = save_path or footnote_doc_path
            if replace_in_footnotes_xml(target_path, old_value, new_value, reason):
                count += 1
                print(f"    [脚注] 已替换: '{old_value[:30]}...'")
        except Exception as e:
            print(f"    [脚注] 替换失败: {e}")

    doc._pending_footnote_replacements = []
    return count
