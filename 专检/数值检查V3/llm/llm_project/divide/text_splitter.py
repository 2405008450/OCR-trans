"""
文本级分割器：借鉴 partition.py 的思想，对提取后的纯文本按字数均分+缓冲区重叠，
确保原文/译文按相同比例分割以保持对齐。

不操作 docx 文件本身，只对字符串做分割。
"""
import re
import math


# --------------- 配置 ---------------
DEFAULT_CHUNK_SIZE = 15000   # 每块目标字数（中文字符 / 英文单词）
DEFAULT_BUFFER_CHARS = 500   # 缓冲区重叠字数
PARAGRAPH_SEP = "\n"         # 段落分隔符


def _count_chars(text: str) -> int:
    """统计有效字符数（中文按字，英文按单词近似）"""
    if not text:
        return 0
    # 简单策略：中文字符数 + 英文单词数
    chinese = len(re.findall(r'[\u4e00-\u9fff]', text))
    english_words = len(re.findall(r'[a-zA-Z]+', text))
    digits = len(re.findall(r'\d+', text))
    return chinese + english_words + digits


def _split_into_paragraphs(text: str):
    """将文本按段落拆分，返回 (段落列表, 每段累计字数列表)"""
    paragraphs = text.split(PARAGRAPH_SEP)
    cumulative = []
    total = 0
    for p in paragraphs:
        total += _count_chars(p)
        cumulative.append(total)
    return paragraphs, cumulative


def _find_split_index(cumulative, target_chars):
    """二分查找最接近 target_chars 的段落索引"""
    left, right = 0, len(cumulative) - 1
    best = 0
    while left <= right:
        mid = (left + right) // 2
        if cumulative[mid] < target_chars:
            best = mid
            left = mid + 1
        else:
            right = mid - 1
    # 比较相邻段落哪个更接近
    if best + 1 < len(cumulative):
        diff_before = target_chars - cumulative[best]
        diff_after = cumulative[best + 1] - target_chars
        if diff_after < diff_before:
            return best + 1
    return best


def compute_split_ratios(text: str, num_parts: int):
    """计算主文档的分割比例（段落位置占比），供另一文档对齐使用。

    Returns:
        ideal_ratios: list[float]  分割点的段落位置比例 (len = num_parts - 1)
    """
    paragraphs, cumulative = _split_into_paragraphs(text)
    total = cumulative[-1] if cumulative else 0
    if total == 0 or num_parts <= 1:
        return []

    target_per_part = total / num_parts
    ratios = []
    for i in range(1, num_parts):
        target = target_per_part * i
        idx = _find_split_index(cumulative, target)
        ratios.append(idx / len(paragraphs) if paragraphs else 0)
    return ratios


def split_text(text: str, num_parts: int, buffer_chars: int = DEFAULT_BUFFER_CHARS,
               split_ratios=None):
    """将文本分割为 num_parts 块，带缓冲区重叠。

    Args:
        text: 待分割文本
        num_parts: 分割份数
        buffer_chars: 缓冲区字数
        split_ratios: 可选，由主文档计算出的分割比例。提供时按此比例分割以保持对齐。

    Returns:
        chunks: list[str]  分割后的文本块
    """
    if num_parts <= 1 or not text:
        return [text] if text else [""]

    paragraphs, cumulative = _split_into_paragraphs(text)
    n = len(paragraphs)
    total = cumulative[-1] if cumulative else 0

    if total == 0:
        return [text]

    # 计算理想分割点（段落索引）
    if split_ratios is not None:
        ideal_splits = []
        for ratio in split_ratios:
            idx = max(0, min(int(ratio * n), n - 1))
            ideal_splits.append(idx)
    else:
        target_per_part = total / num_parts
        ideal_splits = []
        for i in range(1, num_parts):
            target = target_per_part * i
            idx = _find_split_index(cumulative, target)
            ideal_splits.append(idx)

    # 确保严格递增
    for i in range(1, len(ideal_splits)):
        if ideal_splits[i] <= ideal_splits[i - 1]:
            ideal_splits[i] = ideal_splits[i - 1] + 1
    for i in range(len(ideal_splits)):
        ideal_splits[i] = min(ideal_splits[i], n - 1)

    # 计算带缓冲的范围
    def _buffer_end(split_idx, direction):
        if direction == 'right':
            base = cumulative[split_idx] if split_idx < n else cumulative[-1]
            target = base + buffer_chars
            for j in range(split_idx + 1, n):
                if cumulative[j] >= target:
                    return j + 1
            return n
        else:  # left
            base = cumulative[split_idx] if split_idx < n else cumulative[-1]
            target = base - buffer_chars
            if target <= 0:
                return 0
            for j in range(split_idx - 1, -1, -1):
                if cumulative[j] <= target:
                    return j
            return 0

    split_ranges = []
    for part_idx in range(num_parts):
        if part_idx == 0:
            start = 0
            end = _buffer_end(ideal_splits[0], 'right') if ideal_splits else n
        elif part_idx == num_parts - 1:
            start = _buffer_end(ideal_splits[-1], 'left')
            end = n
        else:
            start = _buffer_end(ideal_splits[part_idx - 1], 'left')
            end = _buffer_end(ideal_splits[part_idx], 'right')
        start = max(0, min(start, n - 1))
        end = max(start + 1, min(end, n))
        split_ranges.append((start, end))

    # 拼接段落生成文本块
    chunks = []
    for s, e in split_ranges:
        chunk = PARAGRAPH_SEP.join(paragraphs[s:e])
        chunks.append(chunk)

    return chunks


def auto_num_parts(text: str, chunk_size: int = DEFAULT_CHUNK_SIZE) -> int:
    """根据文本总字数自动计算需要分割的份数"""
    total = _count_chars(text)
    if total <= chunk_size:
        return 1
    return math.ceil(total / chunk_size)


def split_text_pair(original_text: str, translated_text: str,
                    chunk_size: int = DEFAULT_CHUNK_SIZE,
                    buffer_chars: int = DEFAULT_BUFFER_CHARS):
    """分割原文/译文文本对，确保对齐。

    以原文为主文档计算分割比例，译文按相同比例分割。

    Returns:
        list[(orig_chunk, trans_chunk)]  对齐的文本块对
    """
    num_parts = auto_num_parts(original_text, chunk_size)
    if num_parts <= 1:
        return [(original_text, translated_text)]

    # 原文作为主文档，计算分割比例
    ratios = compute_split_ratios(original_text, num_parts)

    # 分割原文和译文
    orig_chunks = split_text(original_text, num_parts, buffer_chars)
    trans_chunks = split_text(translated_text, num_parts, buffer_chars, split_ratios=ratios)

    # 确保数量一致（安全兜底）
    while len(trans_chunks) < len(orig_chunks):
        trans_chunks.append("")
    while len(orig_chunks) < len(trans_chunks):
        orig_chunks.append("")

    return list(zip(orig_chunks, trans_chunks))

# --------------- 双语对照文本分块器 ---------------
# 适用于单文件中英对照场景：中文段开始一组，后面的英文段和表格行归入同一组，
# 只在英文段结束后（即一组结束时）才允许切分。
# 表格行（含 \t 或多个 |）无条件跟随当前组，不会成为切分点。
# 累积到接近 BILINGUAL_CHUNK_SIZE 字时在下一个组边界切，块之间重叠 2 个段落对。

BILINGUAL_CHUNK_SIZE = 100000  # 每块目标字数
BILINGUAL_OVERLAP_PAIRS = 2   # 块间重叠的段落对数


def _is_table_row(line: str) -> bool:
    """判断是否为表格行（含 \\t 或多个 |）"""
    if '\t' in line:
        return True
    if line.count('|') >= 2:
        return True
    return False


def _is_chinese_start(line: str) -> bool:
    """判断一行是否以中文字符为主（作为一组的开始标志）"""
    if not line.strip():
        return False
    chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', line))
    total_alpha = len(re.findall(r'[a-zA-Z]', line))
    # 中文字符数 > 0 且中文占比超过非空字符的一定比例
    return chinese_chars > 0 and chinese_chars >= total_alpha * 0.3


def group_bilingual_paragraphs(text: str):
    """将双语对照文本按 中文段→英文段(+表格行) 分组。

    规则：
    - 遇到中文段开始一组
    - 后面的英文段和表格行都归入这一组
    - 只在英文段后面（即一组结束时）才允许切分
    - 表格行无条件跟随当前组，不会成为切分点

    Returns:
        list[list[str]]  每个元素是一个段落对组（包含中文段+英文段+表格行）
    """
    lines = text.split('\n')
    groups = []
    current_group = []

    for line in lines:
        stripped = line.strip()

        # 空行跟随当前组
        if not stripped:
            if current_group:
                current_group.append(line)
            continue

        # 表格行无条件跟随当前组
        if _is_table_row(stripped):
            if not current_group:
                current_group = []
            current_group.append(line)
            continue

        # 遇到中文段：如果当前组已有内容，说明上一组结束了，切分
        if _is_chinese_start(stripped):
            if current_group:
                groups.append(current_group)
            current_group = [line]
        else:
            # 英文段或其他内容，归入当前组
            if not current_group:
                current_group = []
            current_group.append(line)

    # 最后一组
    if current_group:
        groups.append(current_group)

    return groups


def split_bilingual_text(text: str,
                         chunk_size: int = BILINGUAL_CHUNK_SIZE,
                         overlap_pairs: int = BILINGUAL_OVERLAP_PAIRS):
    """对单文件双语对照文本进行分块。

    按中英段落对分组后，累积到接近 chunk_size 字时在组边界切分，
    块之间重叠 overlap_pairs 个段落对以保持上下文连贯。

    Args:
        text: 双语对照全文
        chunk_size: 每块目标字数（默认 10000）
        overlap_pairs: 块间重叠的段落对数（默认 2）

    Returns:
        list[str]  分块后的文本列表
    """
    if not text or not text.strip():
        return [text] if text else [""]

    groups = group_bilingual_paragraphs(text)

    if not groups:
        return [text]

    # 计算每组的字数
    group_char_counts = []
    for g in groups:
        joined = '\n'.join(g)
        group_char_counts.append(_count_chars(joined))

    total_chars = sum(group_char_counts)

    # 文本不长，不需要分块
    if total_chars <= chunk_size:
        return [text]

    # 按组边界切分，累积到接近 chunk_size 时切
    chunks = []
    current_start = 0  # 当前块的起始组索引
    accumulated = 0

    for i, char_count in enumerate(group_char_counts):
        accumulated += char_count

        # 到达目标字数且不是最后一组 → 在此组边界切分
        if accumulated >= chunk_size and i < len(groups) - 1:
            # 当前块：从 current_start 到 i（含）
            chunk_groups = groups[current_start:i + 1]
            chunk_text = '\n'.join('\n'.join(g) for g in chunk_groups)
            chunks.append(chunk_text)

            # 下一块的起始位置：回退 overlap_pairs 个组作为重叠
            next_start = max(current_start + 1, i + 1 - overlap_pairs)
            current_start = next_start
            # 重新计算累积字数（包含重叠部分）
            accumulated = sum(group_char_counts[current_start:i + 1])

    # 最后一块
    if current_start < len(groups):
        chunk_groups = groups[current_start:]
        chunk_text = '\n'.join('\n'.join(g) for g in chunk_groups)
        chunks.append(chunk_text)

    return chunks

