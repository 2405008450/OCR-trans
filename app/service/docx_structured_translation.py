# -*- coding: utf-8 -*-
"""Word 句段级翻译与原格式回写。

该模块只处理可编辑的 DOCX 文本。它直接修改原 DOCX 包中的 OpenXML，
不会使用 python-docx 或 LibreOffice 重建文档，因此表格、图片、页面设置和
绝大多数段落/字符样式都保留在原位置。
"""

from __future__ import annotations

import hashlib
import json
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable, Sequence

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NS = {"w": W_NS, "a": A_NS}
W_P = f"{{{W_NS}}}p"
W_T = f"{{{W_NS}}}t"
W_TAB = f"{{{W_NS}}}tab"
W_BR = f"{{{W_NS}}}br"
W_CR = f"{{{W_NS}}}cr"
A_P = f"{{{A_NS}}}p"
A_T = f"{{{A_NS}}}t"
XML_SPACE = f"{{{XML_NS}}}space"

SENTENCE_ENDINGS = "。？！?!."
TRAILING_SENTENCE_CLOSERS = "\"'”’】》〉』）)]}"
COMMON_ABBREVIATIONS = {
    "mr", "mrs", "ms", "dr", "prof", "sr", "jr", "vs", "etc", "inc", "ltd",
    "co", "corp", "st", "ave", "blvd", "rd", "dept", "gov", "gen", "col", "lt",
    "sgt", "rev", "hon", "pres", "pp", "vol", "no", "fig", "ed", "eds", "trans",
    "approx", "e.g", "i.e", "cf", "al", "et",
}
STRICT_PRESERVE_SYMBOLS = frozenset(
    "□☐☑☒✓✔✗✘•○●◦▪▫■◆◇→←↑↓"
)
LINE_BREAK_RE = re.compile(r"⟦LB_(\d+)⟧")
INVISIBLE_RE = re.compile(
    r"[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F-\u009F\u200B-\u200D\uFEFF]"
)
SPACE_BEFORE_PUNCTUATION_RE = re.compile(
    r"\s+([。！？!?.,，、；;：:）)\]}])"
)
NUMBER_DOT_RE = re.compile(r"\d$")
SINGLE_LETTER_RE = re.compile(r"^[A-Z]$")
ROMAN_NUMERAL_RE = re.compile(r"^[ivxlcdm]+$", re.IGNORECASE)

MAX_GROUP_ITEMS = 15
MAX_GROUP_SOURCE_CHARS = 3000


class StructuredTranslationError(RuntimeError):
    """句段级翻译或回写失败。"""


class StructuredResponseError(StructuredTranslationError):
    """模型返回不符合句段协议。"""


@dataclass(frozen=True)
class SentenceSpan:
    start: int
    end: int


@dataclass(frozen=True)
class DocxSentence:
    sentence_id: str
    source_text: str
    source_layout_text: str
    part_name: str
    block_index: int
    sentence_index: int

    @property
    def source_hash(self) -> str:
        return hashlib.sha256(normalize_text(self.source_text).encode("utf-8")).hexdigest()


@dataclass
class _TextToken:
    display_text: str
    element: etree._Element | None
    start: int = 0
    end: int = 0

    @property
    def writable(self) -> bool:
        return self.element is not None


@dataclass
class _DocxBlock:
    part_name: str
    block_index: int
    tokens: list[_TextToken]
    sentences: list[DocxSentence] = field(default_factory=list)


@dataclass(frozen=True)
class DocxTranslationResult:
    output_path: Path
    source_sentences: list[DocxSentence]
    translations: dict[str, str]

    @property
    def translated_text(self) -> str:
        return "\n\n".join(
            self.translations.get(sentence.sentence_id, "")
            for sentence in self.source_sentences
        )


LLMMessageCallback = Callable[[list[dict[str, str]], bool], str]


def normalize_text(text: str) -> str:
    if not text:
        return ""
    value = INVISIBLE_RE.sub(" ", text)
    value = re.sub(r"\s+", " ", value)
    value = SPACE_BEFORE_PUNCTUATION_RE.sub(r"\1", value)
    return value.strip()


def normalize_layout_text(text: str) -> str:
    if not text:
        return ""
    value = text.replace("\r\n", "\n").replace("\r", "\n")
    value = INVISIBLE_RE.sub(" ", value)
    value = re.sub(r"[^\S\n]+", " ", value)
    value = re.sub(r"\n{3,}", "\n\n", value)
    return value.strip()


def split_sentence_spans(text: str) -> list[SentenceSpan]:
    scan_text = INVISIBLE_RE.sub(" ", text or "")
    if not normalize_text(scan_text):
        return []

    spans: list[SentenceSpan] = []
    start: int | None = None
    index = 0
    while index < len(scan_text):
        char = scan_text[index]
        if start is None:
            if char.isspace():
                index += 1
                continue
            start = index

        if char in {"\n", "\r"}:
            newline_start = index
            newline_count = 0
            while index < len(scan_text) and scan_text[index] in {"\n", "\r"}:
                newline_count += 1
                if (
                    scan_text[index] == "\r"
                    and index + 1 < len(scan_text)
                    and scan_text[index + 1] == "\n"
                ):
                    index += 2
                else:
                    index += 1
            if newline_count >= 2:
                end = _trim_right(scan_text, newline_start)
                if end > start:
                    spans.append(SentenceSpan(start, end))
                start = None
            continue

        if char in SENTENCE_ENDINGS:
            if char == "." and not _is_sentence_ending_dot(scan_text, index):
                index += 1
                continue
            end = index + 1
            while end < len(scan_text) and scan_text[end] in SENTENCE_ENDINGS:
                end += 1
            while end < len(scan_text) and scan_text[end] in TRAILING_SENTENCE_CLOSERS:
                end += 1
            spans.append(SentenceSpan(start, end))
            start = None
            index = end
            continue
        index += 1

    if start is not None:
        end = _trim_right(scan_text, len(scan_text))
        if end > start:
            spans.append(SentenceSpan(start, end))
    return [span for span in spans if normalize_text(scan_text[span.start:span.end])]


def _trim_right(text: str, end: int) -> int:
    while end > 0 and text[end - 1].isspace():
        end -= 1
    return end


def _is_sentence_ending_dot(text: str, dot_index: int) -> bool:
    word_start = dot_index - 1
    while word_start >= 0 and (text[word_start].isalnum() or text[word_start] == "."):
        word_start -= 1
    word = text[word_start + 1:dot_index]
    if word and NUMBER_DOT_RE.search(word):
        return False
    word_lower = word.lower().rstrip(".")
    if word_lower in COMMON_ABBREVIATIONS:
        return False
    if SINGLE_LETTER_RE.match(word) or ROMAN_NUMERAL_RE.match(word):
        return False
    if dot_index + 1 < len(text):
        next_char = text[dot_index + 1]
        if next_char.isascii() and next_char.isalnum():
            return False
        if next_char.isspace():
            next_index = dot_index + 2
            while next_index < len(text) and text[next_index].isspace():
                next_index += 1
            if next_index >= len(text) or text[next_index].isupper():
                return True
            if text[next_index].islower():
                return False
    return True


def extract_docx_sentences(docx_path: str | Path) -> list[DocxSentence]:
    """提取 DOCX 中所有可编辑句段，不修改源文件。"""
    with zipfile.ZipFile(docx_path) as package:
        _, sentences = _parse_package_blocks(package)
    return sentences


def translate_docx_preserving_format(
    source_path: str | Path,
    output_path: str | Path,
    *,
    source_language: str,
    target_language: str,
    call_llm: LLMMessageCallback,
    translation_rules: str = "",
    bilingual: bool = False,
    retries: int = 2,
) -> DocxTranslationResult:
    """按句翻译 DOCX，并在原 OpenXML 上回写译文。"""
    source = Path(source_path)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(source) as package:
        blocks, sentences = _parse_package_blocks(package)
        if not sentences:
            raise StructuredTranslationError("Word 文档中未找到可编辑的文本句段。")
        translations = _translate_sentence_groups(
            sentences,
            source_language=source_language,
            target_language=target_language,
            call_llm=call_llm,
            translation_rules=translation_rules,
            retries=max(int(retries), 1),
        )
        replacements = {
            sentence.sentence_id: (
                f"{sentence.source_layout_text}\n{translations[sentence.sentence_id]}"
                if bilingual
                else translations[sentence.sentence_id]
            )
            for sentence in sentences
        }
        modified_parts = _apply_translations(blocks, replacements)
        _write_modified_package(package, output, modified_parts)

    return DocxTranslationResult(
        output_path=output,
        source_sentences=sentences,
        translations=translations,
    )


def _content_part_names(package: zipfile.ZipFile) -> list[str]:
    names = set(package.namelist())
    result = ["word/document.xml"] if "word/document.xml" in names else []
    result.extend(
        sorted(
            name
            for name in names
            if (
                name.startswith("word/header")
                or name.startswith("word/footer")
            )
            and name.endswith(".xml")
        )
    )
    result.extend(
        name
        for name in ("word/footnotes.xml", "word/endnotes.xml", "word/comments.xml")
        if name in names
    )
    return result


def _parse_package_blocks(
    package: zipfile.ZipFile,
) -> tuple[list[_DocxBlock], list[DocxSentence]]:
    blocks: list[_DocxBlock] = []
    sentences: list[DocxSentence] = []
    for part_name in _content_part_names(package):
        try:
            root = etree.fromstring(package.read(part_name))
        except etree.XMLSyntaxError as exc:
            raise StructuredTranslationError(f"Word 内容 XML 解析失败：{part_name}: {exc}") from exc

        paragraph_nodes = list(root.iter(W_P))
        drawing_paragraph_nodes = list(root.iter(A_P))
        part_block_index = 0
        for paragraph in [*paragraph_nodes, *drawing_paragraph_nodes]:
            tokens = _collect_paragraph_tokens(paragraph)
            if not normalize_text("".join(token.display_text for token in tokens)):
                continue
            block = _DocxBlock(part_name=part_name, block_index=part_block_index, tokens=tokens)
            block.sentences = _build_block_sentences(block)
            part_block_index += 1
            if block.sentences:
                blocks.append(block)
                sentences.extend(block.sentences)
    return blocks, sentences


def _collect_paragraph_tokens(paragraph: etree._Element) -> list[_TextToken]:
    text_tag = W_T if paragraph.tag == W_P else A_T
    tokens: list[_TextToken] = []

    def walk(node: etree._Element, *, root: bool = False) -> None:
        if not root and node.tag in {W_P, A_P}:
            return
        if node.tag == text_tag:
            tokens.append(_TextToken(node.text or "", node))
            return
        if node.tag == W_TAB:
            tokens.append(_TextToken("\t", None))
            return
        if node.tag in {W_BR, W_CR}:
            tokens.append(_TextToken("\n", None))
            return
        for child in node:
            walk(child)

    walk(paragraph, root=True)
    cursor = 0
    for token in tokens:
        token.start = cursor
        cursor += len(token.display_text)
        token.end = cursor
    return tokens


def _build_block_sentences(block: _DocxBlock) -> list[DocxSentence]:
    display_text = "".join(token.display_text for token in block.tokens)
    spans = split_sentence_spans(display_text)
    if not spans and normalize_text(display_text):
        start = len(display_text) - len(display_text.lstrip())
        end = len(display_text.rstrip())
        spans = [SentenceSpan(start, end)]

    result: list[DocxSentence] = []
    for sentence_index, span in enumerate(spans):
        layout_text = _span_text(block.tokens, span)
        source_text = normalize_text(layout_text)
        if not source_text:
            continue
        identity = f"{block.part_name}|{block.block_index}|{sentence_index}|{source_text}"
        short_hash = hashlib.sha256(identity.encode("utf-8")).hexdigest()[:16]
        result.append(
            DocxSentence(
                sentence_id=f"docx-{short_hash}",
                source_text=source_text,
                source_layout_text=normalize_layout_text(layout_text),
                part_name=block.part_name,
                block_index=block.block_index,
                sentence_index=sentence_index,
            )
        )
    return result


def _span_text(tokens: Sequence[_TextToken], span: SentenceSpan) -> str:
    pieces: list[str] = []
    for token in tokens:
        overlap_start = max(span.start, token.start)
        overlap_end = min(span.end, token.end)
        if overlap_end <= overlap_start:
            continue
        pieces.append(
            token.display_text[
                overlap_start - token.start:overlap_end - token.start
            ]
        )
    return "".join(pieces)


def _translate_sentence_groups(
    sentences: list[DocxSentence],
    *,
    source_language: str,
    target_language: str,
    call_llm: LLMMessageCallback,
    translation_rules: str,
    retries: int,
) -> dict[str, str]:
    grouped: dict[tuple[str, int], list[DocxSentence]] = {}
    for sentence in sentences:
        grouped.setdefault((sentence.part_name, sentence.block_index), []).append(sentence)

    results: dict[str, str] = {}
    for block_sentences in grouped.values():
        for group in _split_group(block_sentences):
            try:
                results.update(
                    _translate_one_group(
                        group,
                        source_language=source_language,
                        target_language=target_language,
                        call_llm=call_llm,
                        translation_rules=translation_rules,
                        retries=retries,
                    )
                )
            except StructuredResponseError:
                if len(group) == 1:
                    raise
                for sentence in group:
                    results.update(
                        _translate_one_group(
                            [sentence],
                            source_language=source_language,
                            target_language=target_language,
                            call_llm=call_llm,
                            translation_rules=translation_rules,
                            retries=retries,
                        )
                    )
    return results


def _split_group(sentences: Iterable[DocxSentence]) -> list[list[DocxSentence]]:
    groups: list[list[DocxSentence]] = []
    current: list[DocxSentence] = []
    current_chars = 0
    for sentence in sentences:
        if current and (
            len(current) >= MAX_GROUP_ITEMS
            or current_chars + len(sentence.source_text) > MAX_GROUP_SOURCE_CHARS
        ):
            groups.append(current)
            current = []
            current_chars = 0
        current.append(sentence)
        current_chars += len(sentence.source_text)
    if current:
        groups.append(current)
    return groups


def _translate_one_group(
    sentences: list[DocxSentence],
    *,
    source_language: str,
    target_language: str,
    call_llm: LLMMessageCallback,
    translation_rules: str,
    retries: int,
) -> dict[str, str]:
    last_error: Exception | None = None
    for attempt in range(retries):
        messages = _build_messages(
            sentences,
            source_language=source_language,
            target_language=target_language,
            translation_rules=translation_rules,
            retry_reason=str(last_error) if last_error else "",
        )
        try:
            raw = call_llm(messages, True)
            return _parse_response(raw, sentences)
        except Exception as exc:  # noqa: BLE001
            last_error = exc
    if isinstance(last_error, StructuredResponseError):
        raise last_error
    raise StructuredResponseError(str(last_error) if last_error else "模型未返回译文。")


def _build_messages(
    sentences: list[DocxSentence],
    *,
    source_language: str,
    target_language: str,
    translation_rules: str,
    retry_reason: str,
) -> list[dict[str, str]]:
    system_prompt = (
        f"你是专业的证件和文档翻译专家。请把句子从{source_language}翻译为{target_language}。"
        "输入中的句子来自同一个 Word 段落或表格区域，可结合上下文理解，但必须逐句返回。"
        "严格保留姓名拼写、证件号码、字段编号、MRZ、URL、邮箱、数字和特殊符号。"
        "⟦LB_n⟧ 是 Word 版式换行标记，必须原样保留且顺序、数量不变。"
        "只返回合法 JSON，不要返回 Markdown、解释或代码块。"
    )
    if translation_rules:
        system_prompt += f"\n\n用户自定义翻译规则：\n{translation_rules}"
    if retry_reason:
        system_prompt += f"\n\n这是纠错重试。上一轮失败原因：{retry_reason}"

    payload_sentences = []
    response_contract: dict[str, dict[str, str]] = {}
    for sentence in sentences:
        source_layout_text = _encode_line_breaks(sentence.source_layout_text)
        payload_sentences.append(
            {
                "sentence_id": sentence.sentence_id,
                "source_hash": sentence.source_hash,
                "source_text": source_layout_text,
            }
        )
        response_contract[sentence.sentence_id] = {
            "source_hash": sentence.source_hash,
            "target_text": f"<{target_language}译文>",
        }

    user_prompt = (
        "翻译输入中的每个句子。translations 的键集合必须和输入 sentence_id 完全一致；"
        "source_hash 必须原样带回，target_text 只填写最终译文。\n\n"
        f"输入：\n{json.dumps({'sentences': payload_sentences}, ensure_ascii=False, indent=2)}\n\n"
        f"输出格式：\n{json.dumps({'translations': response_contract}, ensure_ascii=False, indent=2)}"
    )
    return [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt},
    ]


def _encode_line_breaks(text: str) -> str:
    index = 0
    parts: list[str] = []
    for char in normalize_layout_text(text):
        if char == "\n":
            index += 1
            parts.append(f"⟦LB_{index}⟧")
        else:
            parts.append(char)
    return "".join(parts)


def _parse_response(raw_text: str, sentences: list[DocxSentence]) -> dict[str, str]:
    text = str(raw_text or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\s*```$", "", text)
    try:
        payload = json.loads(text)
    except json.JSONDecodeError:
        start = text.find("{")
        end = text.rfind("}")
        if start < 0 or end <= start:
            raise StructuredResponseError("模型返回的内容不是有效 JSON。") from None
        try:
            payload = json.loads(text[start:end + 1])
        except json.JSONDecodeError as exc:
            raise StructuredResponseError("模型返回的内容不是有效 JSON。") from exc

    translations = payload.get("translations") if isinstance(payload, dict) else None
    if not isinstance(translations, dict):
        raise StructuredResponseError("模型返回缺少 translations 对象。")

    expected = {sentence.sentence_id: sentence for sentence in sentences}
    if set(map(str, translations)) != set(expected):
        raise StructuredResponseError("模型返回的 sentence_id 集合与请求不一致。")

    result: dict[str, str] = {}
    for sentence_id, sentence in expected.items():
        item = translations.get(sentence_id)
        if not isinstance(item, dict):
            raise StructuredResponseError(f"{sentence_id} 的结果不是对象。")
        if str(item.get("source_hash") or "") != sentence.source_hash:
            raise StructuredResponseError(f"{sentence_id} 的 source_hash 不匹配。")
        target = LINE_BREAK_RE.sub("\n", str(item.get("target_text") or "").strip())
        _validate_translation(sentence, target)
        result[sentence_id] = target
    return result


def _validate_translation(sentence: DocxSentence, target: str) -> None:
    if not normalize_text(target):
        raise StructuredResponseError(f"{sentence.sentence_id} 返回了空译文。")
    source_symbols = [char for char in sentence.source_layout_text if char in STRICT_PRESERVE_SYMBOLS]
    target_symbols = [char for char in target if char in STRICT_PRESERVE_SYMBOLS]
    if source_symbols != target_symbols:
        raise StructuredResponseError(f"{sentence.sentence_id} 的特殊符号未按原文保留。")
    if sentence.source_layout_text.count("\n") != target.count("\n"):
        raise StructuredResponseError(f"{sentence.sentence_id} 的版式换行数量不一致。")


def _apply_translations(
    blocks: list[_DocxBlock],
    replacements: dict[str, str],
) -> dict[str, etree._Element]:
    roots: dict[str, etree._Element] = {}
    for block in blocks:
        display_text = "".join(token.display_text for token in block.tokens)
        spans = split_sentence_spans(display_text)
        if not spans and normalize_text(display_text):
            spans = [SentenceSpan(len(display_text) - len(display_text.lstrip()), len(display_text.rstrip()))]
        edits_by_element: dict[etree._Element, list[tuple[int, int, str]]] = {}
        for sentence, span in zip(block.sentences, spans, strict=False):
            replacement = replacements.get(sentence.sentence_id)
            if replacement is None:
                continue
            _queue_span_replacement(block.tokens, span, replacement, edits_by_element)
        for element, edits in edits_by_element.items():
            value = element.text or ""
            for start, end, replacement in sorted(edits, key=lambda item: item[0], reverse=True):
                value = f"{value[:start]}{replacement}{value[end:]}"
            element.text = value
            if value[:1].isspace() or value[-1:].isspace():
                element.set(XML_SPACE, "preserve")
            else:
                element.attrib.pop(XML_SPACE, None)
        if edits_by_element:
            first_element = next(
                (token.element for token in block.tokens if token.element is not None),
                None,
            )
            if first_element is not None:
                roots.setdefault(block.part_name, first_element.getroottree().getroot())
    return roots


def _queue_span_replacement(
    tokens: list[_TextToken],
    span: SentenceSpan,
    replacement: str,
    edits_by_element: dict[etree._Element, list[tuple[int, int, str]]],
) -> None:
    line_break_positions = [
        token.start
        for token in tokens
        if not token.writable
        and token.display_text == "\n"
        and span.start <= token.start < span.end
    ]
    replacement_lines = replacement.split("\n")
    if len(replacement_lines) != len(line_break_positions) + 1:
        raise StructuredTranslationError("译文换行数量与 Word 原文结构不一致。")

    region_boundaries = [span.start, *line_break_positions, span.end]
    for region_index, replacement_line in enumerate(replacement_lines):
        region_start = region_boundaries[region_index]
        region_end = region_boundaries[region_index + 1]
        writable: list[tuple[_TextToken, int, int]] = []
        for token in tokens:
            if not token.writable:
                continue
            start = max(region_start, token.start)
            end = min(region_end, token.end)
            if end > start:
                writable.append((token, start - token.start, end - token.start))
        if not writable:
            if replacement_line:
                raise StructuredTranslationError("Word 换行区域没有可写文本节点。")
            continue

        # 按源 run 覆盖长度比例分配译文，尽量保留原有行内样式分布。
        source_lengths = [max(end - start, 1) for _, start, end in writable]
        chunks = _split_replacement_by_weights(replacement_line, source_lengths)
        for (token, start, end), chunk in zip(writable, chunks, strict=False):
            edits_by_element.setdefault(token.element, []).append((start, end, chunk))


def _split_replacement_by_weights(text: str, weights: list[int]) -> list[str]:
    if len(weights) <= 1:
        return [text]
    total = max(sum(weights), 1)
    boundaries = [0]
    cumulative = 0
    for weight in weights[:-1]:
        cumulative += weight
        candidate = round(len(text) * cumulative / total)
        candidate = _nearest_word_boundary(text, candidate, boundaries[-1])
        boundaries.append(candidate)
    boundaries.append(len(text))
    return [text[boundaries[i]:boundaries[i + 1]] for i in range(len(weights))]


def _nearest_word_boundary(text: str, candidate: int, minimum: int) -> int:
    candidate = max(minimum, min(candidate, len(text)))
    if candidate in {minimum, len(text)} or text[candidate - 1:candidate + 1].isspace():
        return candidate
    for distance in range(1, 13):
        for position in (candidate + distance, candidate - distance):
            if position <= minimum or position >= len(text):
                continue
            if text[position - 1].isspace() or text[position].isspace():
                return position
    return candidate


def _write_modified_package(
    source_package: zipfile.ZipFile,
    output_path: Path,
    modified_parts: dict[str, etree._Element],
) -> None:
    with zipfile.ZipFile(output_path, "w") as target:
        for info in source_package.infolist():
            if info.filename in modified_parts:
                data = etree.tostring(
                    modified_parts[info.filename],
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=True,
                )
            else:
                data = source_package.read(info.filename)
            target.writestr(info, data)
