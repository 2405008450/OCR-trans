"""
向量相似度句级对齐（Bitext Alignment）。

流程：规则切句 → gemini-embedding-001 → 带状 DP 配对 → 低置信度窗口交给 LLM。
输出行始终为原文/译文切片，不做生成式改写。
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, List, Optional, Sequence, Tuple

import numpy as np

from app.core.config import settings
from app.service.docx_structured_translation import normalize_text, split_sentence_spans
from app.service.gemini_service import (
    DEFAULT_EMBEDDING_DIMENSIONS,
    DEFAULT_EMBEDDING_MODEL,
    embed_texts,
)

LogFn = Callable[[str], None]
EmbedFn = Callable[[Sequence[str]], List[List[float]]]
LlmAlignFn = Callable[[str, str], List[dict]]

# DP 操作代价：空对齐略惩罚，避免整篇全 skip
_SKIP_PENALTY = 0.12
_MATCH_BONUS = 0.02
_MIN_ABS_SIM = 0.25


@dataclass(frozen=True)
class TextUnit:
    text: str
    start: int
    end: int


@dataclass
class AlignmentLink:
    src_start: int
    src_end: int  # exclusive
    tgt_start: int
    tgt_end: int  # exclusive
    score: float
    low_confidence: bool = False


def split_text_units(text: str) -> List[TextUnit]:
    spans = split_sentence_spans(text or "")
    units: List[TextUnit] = []
    for span in spans:
        chunk = (text or "")[span.start:span.end]
        cleaned = chunk.strip()
        if not cleaned:
            continue
        units.append(TextUnit(text=cleaned, start=span.start, end=span.end))
    if units:
        return units
    fallback = normalize_text(text or "")
    if not fallback:
        return []
    return [TextUnit(text=fallback, start=0, end=len(text or ""))]


def _mean_rows(matrix: np.ndarray, start: int, end: int) -> np.ndarray:
    if end <= start:
        return np.zeros(matrix.shape[1], dtype=np.float32)
    block = matrix[start:end]
    vec = block.mean(axis=0)
    norm = float(np.linalg.norm(vec))
    if norm > 0:
        vec = vec / norm
    return vec.astype(np.float32)


def _cosine(a: np.ndarray, b: np.ndarray) -> float:
    return float(np.dot(a, b))


def _pair_score(
    src_emb: np.ndarray,
    tgt_emb: np.ndarray,
    i: int,
    j: int,
    src_n: int,
    tgt_n: int,
    src_len: int,
    tgt_len: int,
) -> float:
    src_vec = _mean_rows(src_emb, i, i + src_len)
    tgt_vec = _mean_rows(tgt_emb, j, j + tgt_len)
    sim = _cosine(src_vec, tgt_vec)
    # 轻微奖励 1-1，抑制无意义跳过
    if src_len == 1 and tgt_len == 1:
        sim += _MATCH_BONUS
    return sim


def dp_align_embeddings(
    src_emb: np.ndarray,
    tgt_emb: np.ndarray,
    *,
    band_ratio: float = 0.12,
    min_band: int = 8,
    confidence_threshold: float = 0.55,
) -> List[AlignmentLink]:
    """
    带状 DP 对齐，支持 (1,1)/(1,2)/(2,1)/(1,0)/(0,1)。
    """
    m = int(src_emb.shape[0])
    n = int(tgt_emb.shape[0])
    if m == 0 and n == 0:
        return []
    if m == 0:
        return [
            AlignmentLink(0, 0, j, j + 1, 0.0, True)
            for j in range(n)
        ]
    if n == 0:
        return [
            AlignmentLink(i, i + 1, 0, 0, 0.0, True)
            for i in range(m)
        ]

    band = max(min_band, int(max(m, n) * band_ratio))
    neg = -1e9
    dp = np.full((m + 1, n + 1), neg, dtype=np.float32)
    back: List[List[Optional[Tuple[int, int, int, int, float]]]] = [
        [None] * (n + 1) for _ in range(m + 1)
    ]
    dp[0, 0] = 0.0

    for i in range(m + 1):
        j_lo = 0 if i == 0 else max(0, int(i * n / m) - band)
        j_hi = n if i == m else min(n, int(i * n / m) + band)
        for j in range(j_lo, j_hi + 1):
            if i == 0 and j == 0:
                continue
            best = neg
            best_op = None

            if i > 0 and j > 0:
                score = _pair_score(src_emb, tgt_emb, i - 1, j - 1, m, n, 1, 1)
                cand = dp[i - 1, j - 1] + score
                if cand > best:
                    best, best_op = cand, (i - 1, j - 1, 1, 1, score)

            if i > 0 and j > 1:
                score = _pair_score(src_emb, tgt_emb, i - 1, j - 2, m, n, 1, 2)
                cand = dp[i - 1, j - 2] + score * 0.98
                if cand > best:
                    best, best_op = cand, (i - 1, j - 2, 1, 2, score)

            if i > 1 and j > 0:
                score = _pair_score(src_emb, tgt_emb, i - 2, j - 1, m, n, 2, 1)
                cand = dp[i - 2, j - 1] + score * 0.98
                if cand > best:
                    best, best_op = cand, (i - 2, j - 1, 2, 1, score)

            if i > 0:
                cand = dp[i - 1, j] - _SKIP_PENALTY
                if cand > best:
                    best, best_op = cand, (i - 1, j, 1, 0, 0.0)

            if j > 0:
                cand = dp[i, j - 1] - _SKIP_PENALTY
                if cand > best:
                    best, best_op = cand, (i, j - 1, 0, 1, 0.0)

            if best_op is not None:
                dp[i, j] = best
                back[i][j] = best_op

    # 若终点不可达，放宽到全矩阵回溯最近可达点再补 skip
    i, j = m, n
    if back[i][j] is None and (i > 0 or j > 0):
        # 找 dp 最大的可达格子作为近似终点
        best_ij = (0, 0)
        best_val = dp[0, 0]
        for ii in range(m + 1):
            for jj in range(n + 1):
                if dp[ii, jj] > best_val:
                    best_val = dp[ii, jj]
                    best_ij = (ii, jj)
        i, j = best_ij

    links_rev: List[AlignmentLink] = []
    while i > 0 or j > 0:
        op = back[i][j]
        if op is None:
            if i > 0:
                links_rev.append(AlignmentLink(i - 1, i, j, j, 0.0, True))
                i -= 1
                continue
            if j > 0:
                links_rev.append(AlignmentLink(i, i, j - 1, j, 0.0, True))
                j -= 1
                continue
            break
        pi, pj, src_len, tgt_len, score = op
        low = (
            (src_len == 0 or tgt_len == 0)
            or score < confidence_threshold
            or score < _MIN_ABS_SIM
        )
        links_rev.append(
            AlignmentLink(
                src_start=pi,
                src_end=pi + src_len,
                tgt_start=pj,
                tgt_end=pj + tgt_len,
                score=float(score),
                low_confidence=low,
            )
        )
        i, j = pi, pj

    links = list(reversed(links_rev))

    # 补上未覆盖尾部（近似终点导致）
    covered_src = max((link.src_end for link in links), default=0)
    covered_tgt = max((link.tgt_end for link in links), default=0)
    while covered_src < m:
        links.append(AlignmentLink(covered_src, covered_src + 1, n, n, 0.0, True))
        covered_src += 1
    while covered_tgt < n:
        links.append(AlignmentLink(m, m, covered_tgt, covered_tgt + 1, 0.0, True))
        covered_tgt += 1
    return links


def links_to_rows(
    src_units: Sequence[TextUnit],
    tgt_units: Sequence[TextUnit],
    links: Sequence[AlignmentLink],
) -> List[dict]:
    rows: List[dict] = []
    for link in links:
        src_text = " ".join(
            unit.text for unit in src_units[link.src_start:link.src_end]
        ).strip()
        tgt_text = " ".join(
            unit.text for unit in tgt_units[link.tgt_start:link.tgt_end]
        ).strip()
        if not src_text and not tgt_text:
            continue
        rows.append(
            {
                "原文": src_text,
                "译文": tgt_text,
                "_score": link.score,
                "_low_confidence": link.low_confidence,
                "_src_range": (link.src_start, link.src_end),
                "_tgt_range": (link.tgt_start, link.tgt_end),
            }
        )
    return rows


def _merge_low_confidence_windows(
    rows: List[dict],
    *,
    max_units: int = 12,
) -> List[Tuple[int, int]]:
    """返回需要 LLM 重对齐的行区间 [start, end)。"""
    windows: List[Tuple[int, int]] = []
    i = 0
    n = len(rows)
    while i < n:
        if not rows[i].get("_low_confidence"):
            i += 1
            continue
        j = i
        unit_budget = 0
        while j < n and rows[j].get("_low_confidence"):
            src_range = rows[j].get("_src_range") or (0, 0)
            tgt_range = rows[j].get("_tgt_range") or (0, 0)
            unit_budget += max(src_range[1] - src_range[0], 1) + max(
                tgt_range[1] - tgt_range[0], 0
            )
            j += 1
            if unit_budget >= max_units:
                break
        windows.append((i, j))
        i = j
    return windows


def align_bitext_with_embeddings(
    original_text: str,
    translated_text: str,
    *,
    embed_fn: Optional[EmbedFn] = None,
    llm_align_fn: Optional[LlmAlignFn] = None,
    confidence_threshold: float = 0.55,
    use_llm_fallback: bool = True,
    log_fn: Optional[LogFn] = None,
) -> List[dict]:
    """
    返回 [{"原文","译文"}, ...]，文本均来自源切片（LLM 回退段除外，但仍要求模型原样引用）。
    """
    def _log(msg: str) -> None:
        if log_fn:
            log_fn(msg)

    src_units = split_text_units(original_text)
    tgt_units = split_text_units(translated_text)
    _log(f"[embedding-align] 切句完成：原文 {len(src_units)} 句，译文 {len(tgt_units)} 句")

    if not src_units and not tgt_units:
        return []

    if embed_fn is None:
        def embed_fn(texts: Sequence[str]) -> List[List[float]]:
            return embed_texts(
                texts,
                model=DEFAULT_EMBEDDING_MODEL,
                dimensions=DEFAULT_EMBEDDING_DIMENSIONS,
                task_type="SEMANTIC_SIMILARITY",
                log_callback=log_fn,
            )

    src_texts = [u.text for u in src_units]
    tgt_texts = [u.text for u in tgt_units]
    src_vectors = embed_fn(src_texts) if src_texts else []
    tgt_vectors = embed_fn(tgt_texts) if tgt_texts else []

    src_emb = (
        np.asarray(src_vectors, dtype=np.float32)
        if src_vectors
        else np.zeros((0, DEFAULT_EMBEDDING_DIMENSIONS), dtype=np.float32)
    )
    tgt_emb = (
        np.asarray(tgt_vectors, dtype=np.float32)
        if tgt_vectors
        else np.zeros((0, DEFAULT_EMBEDDING_DIMENSIONS), dtype=np.float32)
    )

    threshold = float(
        confidence_threshold
        if confidence_threshold is not None
        else settings.ALIGNMENT_EMBEDDING_CONFIDENCE
    )
    links = dp_align_embeddings(
        src_emb,
        tgt_emb,
        confidence_threshold=threshold,
    )
    rows = links_to_rows(src_units, tgt_units, links)
    high = sum(1 for row in rows if not row.get("_low_confidence"))
    low = len(rows) - high
    _log(f"[embedding-align] DP 对齐完成：{len(rows)} 行（高置信 {high}，低置信 {low}）")

    if use_llm_fallback and llm_align_fn and low > 0:
        windows = _merge_low_confidence_windows(rows)
        _log(f"[embedding-align] 低置信窗口数: {len(windows)}，调用 LLM 补齐")
        rebuilt: List[dict] = []
        cursor = 0
        for start, end in windows:
            rebuilt.extend(rows[cursor:start])
            src_bits = []
            tgt_bits = []
            for row in rows[start:end]:
                if row.get("原文"):
                    src_bits.append(row["原文"])
                if row.get("译文"):
                    tgt_bits.append(row["译文"])
            src_block = "\n".join(src_bits).strip()
            tgt_block = "\n".join(tgt_bits).strip()
            if src_block or tgt_block:
                try:
                    llm_rows = llm_align_fn(src_block, tgt_block) or []
                    if llm_rows:
                        for item in llm_rows:
                            rebuilt.append(
                                {
                                    "原文": item.get("原文", ""),
                                    "译文": item.get("译文", ""),
                                }
                            )
                    else:
                        rebuilt.extend(
                            {"原文": r.get("原文", ""), "译文": r.get("译文", "")}
                            for r in rows[start:end]
                            if r.get("原文") or r.get("译文")
                        )
                except Exception as exc:
                    _log(f"[embedding-align] LLM 回退失败，保留向量结果: {exc}")
                    rebuilt.extend(
                        {"原文": r.get("原文", ""), "译文": r.get("译文", "")}
                        for r in rows[start:end]
                        if r.get("原文") or r.get("译文")
                    )
            cursor = end
        rebuilt.extend(rows[cursor:])
        rows = rebuilt
    else:
        rows = [
            {"原文": r.get("原文", ""), "译文": r.get("译文", "")}
            for r in rows
            if r.get("原文") or r.get("译文")
        ]

    # 去掉空对
    return [
        {"原文": (r.get("原文") or "").strip(), "译文": (r.get("译文") or "").strip()}
        for r in rows
        if (r.get("原文") or "").strip() or (r.get("译文") or "").strip()
    ]
