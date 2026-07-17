# -*- coding: utf-8 -*-
"""英式 / 美式英语词汇转换服务。"""

from __future__ import annotations

import json
import re
from collections import Counter
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from typing import Any, Literal


TargetStyle = Literal["british", "american"]

REPO_ROOT = Path(__file__).resolve().parents[2]
DICTIONARY_PATH = REPO_ROOT / "data" / "english_variant" / "dictionary.json"
ALLOWED_EXTENSIONS = [".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt"]
WORD_BOUNDARY_LEFT = r"(?<![A-Za-z0-9])"
WORD_BOUNDARY_RIGHT = r"(?![A-Za-z0-9])"


def normalize_target_style(target_style: str | None) -> TargetStyle:
    normalized = str(target_style or "").strip().lower()
    if normalized not in {"british", "american"}:
        raise ValueError("target_style 只能是 british 或 american")
    return normalized  # type: ignore[return-value]


def _capitalize_first_letter(text: str) -> str:
    for index, char in enumerate(text):
        if "a" <= char <= "z":
            return text[:index] + char.upper() + text[index + 1 :]
        if "A" <= char <= "Z":
            return text
    return text


def _match_case(source_text: str, target_text: str) -> str:
    letters = [char for char in source_text if char.isalpha()]
    if letters and all(char.isupper() for char in letters):
        return target_text.upper()
    if source_text.istitle():
        return target_text.title()

    first_letter_seen = False
    first_is_upper = False
    remaining_are_lower = True
    for char in source_text:
        if not char.isalpha():
            continue
        if not first_letter_seen:
            first_letter_seen = True
            first_is_upper = char.isupper()
        elif not char.islower():
            remaining_are_lower = False
    if first_letter_seen and first_is_upper and remaining_are_lower:
        return _capitalize_first_letter(target_text)
    return target_text


@dataclass(frozen=True)
class DirectionRules:
    lookup: dict[str, str]
    canonical_sources: dict[str, str]
    ambiguous: dict[str, tuple[str, ...]]
    pattern: re.Pattern[str] | None


class EnglishVariantConverter:
    def __init__(self, payload: dict[str, Any]) -> None:
        if payload.get("schema_version") != 1:
            raise ValueError("不支持的英美词库 schema_version")
        self.payload = payload
        self.dictionary_version = str(payload.get("dictionary_version") or "")
        self.source_sha256 = str(payload.get("source_sha256") or "")
        directions = payload.get("directions") or {}
        self._to_american = self._compile_direction(directions.get("british_to_american") or {})
        self._to_british = self._compile_direction(directions.get("american_to_british") or {})

    @staticmethod
    def _compile_direction(direction: dict[str, Any]) -> DirectionRules:
        lookup: dict[str, str] = {}
        canonical_sources: dict[str, str] = {}
        targets: set[str] = set()
        for rule in direction.get("rules") or []:
            source = str(rule.get("source") or "").strip()
            target = str(rule.get("target") or "").strip()
            if not source or not target or source.casefold() == target.casefold():
                continue
            key = source.casefold()
            lookup[key] = target
            canonical_sources[key] = source
            targets.add(target.casefold())

        ambiguous: dict[str, tuple[str, ...]] = {}
        for item in direction.get("ambiguous") or []:
            source = str(item.get("source") or "").strip()
            candidates = tuple(
                str(candidate.get("target") or "").strip()
                for candidate in item.get("candidates") or []
                if str(candidate.get("target") or "").strip()
            )
            if source and candidates:
                ambiguous[source.casefold()] = candidates
                canonical_sources[source.casefold()] = source

        protected_targets = {target for target in targets if target not in lookup}
        candidates = sorted(
            set(lookup) | set(ambiguous) | protected_targets,
            key=lambda value: (-len(value), value),
        )
        pattern = None
        if candidates:
            alternatives = "|".join(re.escape(candidate) for candidate in candidates)
            pattern = re.compile(
                f"{WORD_BOUNDARY_LEFT}(?:{alternatives}){WORD_BOUNDARY_RIGHT}",
                flags=re.IGNORECASE,
            )
        return DirectionRules(lookup, canonical_sources, ambiguous, pattern)

    def convert(
        self,
        text: str,
        target_style: str,
        *,
        include_edits: bool = False,
    ) -> dict[str, Any]:
        if not isinstance(text, str):
            raise TypeError("text 必须是字符串")
        style = normalize_target_style(target_style)
        rules = self._to_british if style == "british" else self._to_american
        if not text or rules.pattern is None:
            return _empty_result(text, style, self.dictionary_version, self.source_sha256)

        replacement_counts: Counter[tuple[str, str]] = Counter()
        ambiguous_counts: Counter[str] = Counter()
        edits: list[dict[str, Any]] = []

        def replace_match(match: re.Match[str]) -> str:
            before = match.group(0)
            key = before.casefold()
            target = rules.lookup.get(key)
            if target is not None:
                after = _match_case(before, target)
                canonical_source = rules.canonical_sources.get(key, before)
                replacement_counts[(canonical_source, target)] += 1
                if include_edits:
                    edits.append(
                        {
                            "start": match.start(),
                            "end": match.end(),
                            "before": before,
                            "after": after,
                        }
                    )
                return after
            if key in rules.ambiguous:
                ambiguous_counts[key] += 1
            return before

        converted = rules.pattern.sub(replace_match, text)
        replacements = [
            {"source": source, "target": target, "count": count}
            for (source, target), count in sorted(
                replacement_counts.items(), key=lambda item: (-item[1], item[0][0].casefold())
            )
        ]
        ambiguous_hits = [
            {
                "term": rules.canonical_sources.get(key, key),
                "candidates": list(rules.ambiguous[key]),
                "count": count,
            }
            for key, count in sorted(
                ambiguous_counts.items(), key=lambda item: (-item[1], item[0])
            )
        ]
        result = {
            "converted_text": converted,
            "target_style": style,
            "replacement_count": sum(replacement_counts.values()),
            "distinct_rule_count": len(replacement_counts),
            "replacements": replacements,
            "ambiguous_hit_count": sum(ambiguous_counts.values()),
            "ambiguous_hits": ambiguous_hits,
            "dictionary_version": self.dictionary_version,
            "dictionary_sha256": self.source_sha256,
        }
        if include_edits:
            result["_edits"] = edits
        return result


def _empty_result(
    text: str,
    target_style: TargetStyle,
    dictionary_version: str,
    source_sha256: str,
) -> dict[str, Any]:
    return {
        "converted_text": text,
        "target_style": target_style,
        "replacement_count": 0,
        "distinct_rule_count": 0,
        "replacements": [],
        "ambiguous_hit_count": 0,
        "ambiguous_hits": [],
        "dictionary_version": dictionary_version,
        "dictionary_sha256": source_sha256,
    }


@lru_cache(maxsize=1)
def get_converter() -> EnglishVariantConverter:
    if not DICTIONARY_PATH.is_file():
        raise FileNotFoundError(f"英美词库不存在: {DICTIONARY_PATH}")
    payload = json.loads(DICTIONARY_PATH.read_text(encoding="utf-8"))
    return EnglishVariantConverter(payload)


def convert_text(text: str, target_style: str) -> dict[str, Any]:
    return get_converter().convert(text, target_style)


def get_english_variant_config() -> dict[str, Any]:
    converter = get_converter()
    stats = converter.payload.get("stats") or {}
    return {
        "allowed_extensions": ALLOWED_EXTENSIONS,
        "target_styles": {
            "british": {"label": "英式英语"},
            "american": {"label": "美式英语"},
        },
        "default_target_style": "british",
        "dictionary_version": converter.dictionary_version,
        "dictionary_sha256": converter.source_sha256,
        "stats": stats,
    }


__all__ = [
    "ALLOWED_EXTENSIONS",
    "DICTIONARY_PATH",
    "EnglishVariantConverter",
    "convert_text",
    "get_converter",
    "get_english_variant_config",
    "normalize_target_style",
]
