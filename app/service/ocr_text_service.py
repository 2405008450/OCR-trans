# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from typing import Any, Callable, Iterable, Optional

from bs4 import BeautifulSoup
from markdown_it import MarkdownIt

from pdf2docx import ocr_file


_PAGE_BREAK_RE = re.compile(r"\s*<page_break\s*/>\s*", flags=re.IGNORECASE)
_CODE_FENCE_RE = re.compile(r"^```(?:html|markdown)?\s*\n?|\n?```\s*$", flags=re.IGNORECASE)
_BLOCK_TAGS = (
    "address",
    "article",
    "blockquote",
    "div",
    "footer",
    "h1",
    "h2",
    "h3",
    "h4",
    "h5",
    "h6",
    "header",
    "li",
    "p",
    "section",
    "tr",
)


def ocr_markup_to_plain_text(raw_text: str) -> str:
    """把 OCR 的 HTML/Markdown 输出转换为只包含可见文字的纯文本。"""
    pages: list[str] = []
    for fragment in _PAGE_BREAK_RE.split(raw_text or ""):
        normalized = _CODE_FENCE_RE.sub("", fragment.strip())
        if not normalized:
            pages.append("")
            continue

        rendered = MarkdownIt("commonmark", {"html": True}).render(normalized)
        soup = BeautifulSoup(rendered, "html.parser")
        for tag in soup.find_all(["script", "style", "head"]):
            tag.decompose()
        for tag in soup.find_all("br"):
            tag.replace_with("\n")
        for tag in soup.find_all(["td", "th"]):
            tag.append("\t")
        for tag in soup.find_all(_BLOCK_TAGS):
            tag.append("\n")

        lines = []
        for line in soup.get_text().replace("\xa0", " ").splitlines():
            compact = re.sub(r"[ \t]+", " ", line).strip()
            if compact:
                lines.append(compact)
        pages.append("\n".join(lines))

    return "\n\n".join(page for page in pages if page).strip()


def extract_ocr_plain_text(
    *,
    file_path: str,
    model: str,
    gemini_route: str,
    page_numbers: Optional[Iterable[int]] = None,
    page_progress_callback: Optional[Callable[[int, int], None]] = None,
    status_callback: Optional[Callable[[str], None]] = None,
    continue_on_error: bool = True,
) -> dict[str, Any]:
    """复用 PDF2DOCX OCR，并返回可直接统计的逐页纯文本。"""
    payload = ocr_file(
        file_path=file_path,
        model=model,
        gemini_route=gemini_route,
        page_progress_callback=page_progress_callback,
        return_metadata=True,
        ocr_status_callback=status_callback,
        page_numbers=page_numbers,
        continue_on_error=continue_on_error,
    )
    if not isinstance(payload, dict):
        payload = {"text": str(payload or ""), "page_results": []}

    page_results: list[dict[str, Any]] = []
    for item in payload.get("page_results") or []:
        normalized = dict(item)
        normalized["raw_text"] = str(item.get("text") or "")
        normalized["text"] = "" if item.get("error") else ocr_markup_to_plain_text(normalized["raw_text"])
        page_results.append(normalized)

    if not page_results and not payload.get("failed_pages"):
        page_results = [
            {
                "page_number": 1,
                "raw_text": str(payload.get("text") or ""),
                "text": ocr_markup_to_plain_text(str(payload.get("text") or "")),
                "blank": False,
                "error": "",
            }
        ]

    result = dict(payload)
    result["raw_text"] = str(payload.get("text") or "")
    result["page_results"] = page_results
    result["text"] = "\n\n".join(
        str(item.get("text") or "") for item in page_results if not item.get("error")
    ).strip()
    return result
