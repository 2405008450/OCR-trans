"""
长文本分块模块

将原文/译文按对齐的块切分，用于逐块调用 LLM 检查。
支持两种模式：
  - 原文+译文分开（split_text_pair）：按段落索引比例对齐
  - 中英双语对照单文件（split_bilingual_text）：按段落对分组

原文+译文分开模式的分块流程（4步）：
  1. auto_num_parts：只看原文字数，除以 DEFAULT_CHUNK_SIZE 向上取整
  2. compute_split_ratios：原文按段落累计字数，二分查找分割点，转为比例
  3. split_text：原文按字数均分切；译文按原文的比例映射段落索引切
  4. split_text_pair：原文第i块和译文第i块 zip 配对
"""
import re
import math
import bisect
from typing import List, Tuple, Optional


# ========================= 配置 =========================
# 原文+译文分开模式：每块目标字数
DEFAULT_CHUNK_SIZE = 15000
# 缓冲区扩展字符数（段落级前后扩展）
BUFFER_CHARS = 500
# 双语对照模式：每块最大字符数
BILINGUAL_CHUNK_SIZE = 15000


# ========================= 通用工具 =========================

def _count_chars(text: str) -> int:
    """统计文本的有效字数。

    中文：每个汉字算1字。
    英文：按空格分词，每个单词算1字。
    数字：连续数字串算1字。
    """
    if not text:
        return 0
    count = 0
    i = 0
    while i < len(text):
        ch = text[i]
        if '\u4e00' <= ch <= '\u9fff':
            count += 1
            i += 1
        elif ch.isdigit():
            while i < len(text) and (text[i].isdigit() or text[i] in '.,'):
                i += 1
            count += 1
        elif ch.isalpha():
            while i < len(text) and text[i].isalpha():
                i += 1
            count += 1
        else:
            i += 1
    return count


# =========================================================
# 原文+译文分开模式
# =========================================================

def auto_num_parts(text: str, chunk_size: int = DEFAULT_CHUNK_SIZE) -> int:
    """第1步：根据原文字数自动计算分块数。"""
    total = _count_chars(text)
    if total <= chunk_size:
        return 1
    return math.ceil(total / chunk_size)


def _paragraphs_and_cumulative_chars(text: str) -> Tuple[List[str], List[int]]:
    """将文本按换行拆成段落，计算每段的累计字数。

    Returns:
        (paragraphs, cumulative_chars)
        cumulative_chars[i] = 前 i+1 段的总字数
    """
    paragraphs = text.split('\n')
    cumulative = []
    total = 0
    for p in paragraphs:
        total += _count_chars(p)
        cumulative.append(total)
    return paragraphs, cumulative


def compute_split_ratios(text: str, num_parts: int) -> List[float]:
    """第2步：根据原文段落累计字数，用二分查找算出分割比例。

    比如 num_parts=3 → 需要2个分割点 → 返回 [0.33, 0.66] 之类的比例。

    Returns:
        长度为 num_parts-1 的比例列表，每个值在 0~1 之间
    """
    if num_parts <= 1:
        return []

    paragraphs, cumulative = _paragraphs_and_cumulative_chars(text)
    total_chars = cumulative[-1] if cumulative else 0
    num_paras = len(paragraphs)

    if total_chars == 0 or num_paras == 0:
        return []

    ratios = []
    for k in range(1, num_parts):
        target_chars = total_chars * k / num_parts
        # 二分查找：找到累计字数最接近 target_chars 的段落索引
        idx = bisect.bisect_left(cumulative, target_chars)
        idx = min(idx, num_paras - 1)
        ratio = idx / num_paras
        ratios.append(ratio)

    return ratios


def _buffer_end(paragraphs: List[str], start_idx: int, direction: int,
                buffer_chars: int = BUFFER_CHARS) -> int:
    """从 start_idx 向 direction 方向扩展约 buffer_chars 字的缓冲区。

    Args:
        paragraphs: 段落列表
        start_idx: 起始段落索引
        direction: 1=向后扩展, -1=向前扩展
        buffer_chars: 缓冲区目标字数

    Returns:
        扩展后的段落索引（包含）
    """
    accumulated = 0
    idx = start_idx
    while 0 <= idx < len(paragraphs) and accumulated < buffer_chars:
        accumulated += _count_chars(paragraphs[idx])
        idx += direction
    # 回退一步（idx 已经越过了）
    return idx - direction


def split_text(text: str, num_parts: int,
               split_ratios: Optional[List[float]] = None,
               buffer_chars: int = BUFFER_CHARS) -> List[str]:
    """第3步：将文本切成 num_parts 块，在段落边界切分，带缓冲区重叠。

    Args:
        text: 要切分的文本
        num_parts: 目标块数
        split_ratios: 分割比例列表（长度 num_parts-1）。
                      None 时按自身字数均分（用于原文）；
                      传入时按比例映射段落索引（用于译文）。
        buffer_chars: 缓冲区字数

    Returns:
        切分后的文本块列表
    """
    if num_parts <= 1 or not text:
        return [text] if text else []

    paragraphs = text.split('\n')
    num_paras = len(paragraphs)

    if num_paras <= num_parts:
        # 段落数不够分，直接返回整个文本
        return [text]

    # 计算分割点（段落索引）
    if split_ratios is None:
        # 原文模式：按自身字数均分
        split_ratios = compute_split_ratios(text, num_parts)

    # 比例 → 段落索引
    split_indices = []
    for ratio in split_ratios:
        idx = int(round(ratio * num_paras))
        idx = max(1, min(idx, num_paras - 1))
        split_indices.append(idx)

    # 去重并排序
    split_indices = sorted(set(split_indices))

    # 构建块边界：[0, split1, split2, ..., num_paras]
    boundaries = [0] + split_indices + [num_paras]

    chunks = []
    for i in range(len(boundaries) - 1):
        chunk_start = boundaries[i]
        chunk_end = boundaries[i + 1]

        # 向前扩展缓冲区（不是第一块时）
        if i > 0:
            buf_start = _buffer_end(paragraphs, chunk_start - 1, -1, buffer_chars)
            buf_start = max(0, buf_start)
        else:
            buf_start = chunk_start

        # 向后扩展缓冲区（不是最后一块时）
        if i < len(boundaries) - 2:
            buf_end = _buffer_end(paragraphs, chunk_end, 1, buffer_chars)
            buf_end = min(num_paras, buf_end + 1)
        else:
            buf_end = chunk_end

        chunk_text = '\n'.join(paragraphs[buf_start:buf_end])
        chunks.append(chunk_text)

    return chunks


def split_text_pair(orig_text: str, trans_text: str,
                    chunk_size: int = DEFAULT_CHUNK_SIZE,
                    buffer_chars: int = BUFFER_CHARS) -> List[Tuple[str, str]]:
    """第4步：将原文和译文按段落索引比例对齐切块，zip 配对。

    流程：
    1. 按原文字数算分块数
    2. 按原文段落累计字数算分割比例
    3. 原文按自身字数均分切；译文按原文比例映射切
    4. zip 配对

    Returns:
        [(orig_chunk_1, trans_chunk_1), ...]
    """
    if not orig_text or not trans_text:
        return [(orig_text or "", trans_text or "")]

    # 算分块数
    num_parts = auto_num_parts(orig_text, chunk_size)

    if num_parts <= 1:
        return [(orig_text, trans_text)]

    # 算原文的分割比例
    ratios = compute_split_ratios(orig_text, num_parts)

    if not ratios:
        return [(orig_text, trans_text)]

    # 原文：按自身字数均分切（split_ratios=None → 内部自己算）
    orig_chunks = split_text(orig_text, num_parts, split_ratios=None, buffer_chars=buffer_chars)

    # 译文：按原文的比例映射切
    trans_chunks = split_text(trans_text, num_parts, split_ratios=ratios, buffer_chars=buffer_chars)

    # 对齐块数（理论上应该相等，防御性处理）
    min_len = min(len(orig_chunks), len(trans_chunks))
    pairs = list(zip(orig_chunks[:min_len], trans_chunks[:min_len]))

    # 如果有剩余，追加到最后一块
    if len(orig_chunks) > min_len:
        extra_orig = '\n'.join(orig_chunks[min_len:])
        last_o, last_t = pairs[-1]
        pairs[-1] = (last_o + '\n' + extra_orig, last_t)
    if len(trans_chunks) > min_len:
        extra_trans = '\n'.join(trans_chunks[min_len:])
        last_o, last_t = pairs[-1]
        pairs[-1] = (last_o, last_t + '\n' + extra_trans)

    return pairs


# =========================================================
# 中英双语对照单文件模式
# =========================================================

def _is_chinese_start(s: str) -> bool:
    """判断段落是否以中文字符开头（跳过空白）"""
    for c in s:
        if c.isspace():
            continue
        return '\u4e00' <= c <= '\u9fff'
    return False


def _is_table_line(s: str) -> bool:
    """判断是否为表格行（含制表符或多个管道符）"""
    if '\t' in s:
        return True
    if s.count('|') >= 2:
        return True
    return False


def group_bilingual_paragraphs(paragraphs: List[str]) -> List[str]:
    """将段落列表按中英对照关系分组。

    规则：
    - 遇到以中文开头的段落 → 开新组
    - 英文段落、表格行 → 归入当前组
    - 一个组 = 一个完整的中英对照单元 [中文段, 英文段, 可能的表格行...]

    Returns:
        每个元素是一个组的完整文本（组内段落用 \\n 连接）
    """
    groups = []
    current = []

    for para in paragraphs:
        # 表格行无条件归入当前组
        if _is_table_line(para):
            current.append(para)
            continue

        if _is_chinese_start(para):
            # 遇到中文段：如果当前组非空，保存当前组，开新组
            if current:
                groups.append('\n'.join(current))
            current = [para]
        else:
            # 英文段：归入当前组
            current.append(para)

    if current:
        groups.append('\n'.join(current))

    return groups


def split_bilingual_text(text: str, max_chars: int = BILINGUAL_CHUNK_SIZE,
                         overlap_pairs: int = 2) -> List[str]:
    """将中英双语对照文本按段落对分块。

    分块流程：
    1. 按 \\n 拆段落，过滤空行
    2. 用 _is_chinese_start 判断首字符是否中文，识别中文段
    3. 用 _is_table_line 识别表格行（含 \\t 或多个 |）
    4. 分组：遇到中文段就开新组，英文段和表格行归入当前组，一组 = 一对中英文
    5. 按组累积字符数，到 max_chars 时在组边界切，重叠 overlap_pairs 组

    Args:
        text: 双语对照全文
        max_chars: 每块最大字符数
        overlap_pairs: 块之间重叠的段落对数

    Returns:
        [chunk_1, chunk_2, ...]
    """
    if not text or len(text) <= max_chars:
        return [text] if text else []

    # 第1步：按换行拆段落，过滤空行
    paragraphs = [line.strip() for line in text.split('\n') if line.strip()]

    if not paragraphs:
        return [text]

    # 第2-4步：按中英对照关系分组
    groups = group_bilingual_paragraphs(paragraphs)

    if not groups:
        return [text]

    # 第5步：按组累积字符数，在组边界切分
    chunks = []
    current_parts = []
    current_size = 0

    for group_text in groups:
        group_size = len(group_text)

        # 单个组就超限 → 单独成块（不拆开中英对）
        if group_size > max_chars:
            if current_parts:
                chunks.append('\n\n'.join(current_parts))
                current_parts = []
                current_size = 0
            chunks.append(group_text)
            continue

        # 加入后超限 → 先保存当前块
        if current_size + group_size + 2 > max_chars and current_parts:
            chunks.append('\n\n'.join(current_parts))

            # 重叠：回退 overlap_pairs 个组到下一块
            if overlap_pairs > 0 and len(current_parts) > overlap_pairs:
                current_parts = current_parts[-overlap_pairs:]
                current_size = sum(len(p) for p in current_parts) + 2 * (len(current_parts) - 1)
            else:
                current_parts = []
                current_size = 0

        current_parts.append(group_text)
        current_size += group_size + (2 if current_size > 0 else 0)

    if current_parts:
        chunks.append('\n\n'.join(current_parts))

    return chunks
