from __future__ import annotations

from typing import Any, Mapping, Optional

TASK_MODEL_FIELDS: dict[str, tuple[str, ...]] = {
    "alignment": ("model_name",),
    "business_licence": ("model",),
    "doc_translate": ("translation_model", "ocr_model"),
    "number_check": ("model_name",),
    "pdf2docx": ("model",),
    "zhongfanyi": ("model_name",),
}

_MODEL_ALIASES = {
    "google gemini 2.5 flash": "google/gemini-2.5-flash",
    "google gemini 2.5 pro": "google/gemini-2.5-pro",
    "google gemini-3-flash-preview": "google/gemini-3-flash-preview",
    "google: google/gemini-3.1-pro-preview": "google/gemini-3.1-pro-preview",
    "gemini-3-flash-preview": "google/gemini-3-flash-preview",
    "gemini-3.1-pro-preview": "google/gemini-3.1-pro-preview",
}

_MODEL_LABELS = {
    "google/gemini-2.5-flash": "快速版V2",
    "google/gemini-3-flash-preview": "快速版V2",
    "google/gemini-2.5-pro": "增强版V2",
    "google/gemini-3.1-pro-preview": "增强版V2",
    "gemini-3.1-flash-lite-preview": "极速版V2",
}


def _clean_model_name(value: Any) -> Optional[str]:
    if not isinstance(value, str):
        return None
    candidate = value.strip()
    return candidate or None



def canonicalize_model_name(value: Any) -> Optional[str]:
    candidate = _clean_model_name(value)
    if not candidate:
        return None
    lowered = candidate.lower()
    return _MODEL_ALIASES.get(lowered, candidate)



def build_task_model_info(
    task_type: Optional[str],
    params: Optional[Mapping[str, Any]],
    result: Optional[Mapping[str, Any]],
) -> Optional[dict[str, str]]:
    field_candidates = list(TASK_MODEL_FIELDS.get(task_type or "", ())) + ["translation_model", "model_name", "ocr_model", "model"]
    data_sources: tuple[tuple[str, Optional[Mapping[str, Any]]], ...] = (("params", params), ("result", result))

    raw_model = None
    source = None
    field_name = None
    for candidate_field in field_candidates:
        for source_name, source_payload in data_sources:
            if not isinstance(source_payload, Mapping):
                continue
            value = _clean_model_name(source_payload.get(candidate_field))
            if value:
                raw_model = value
                source = source_name
                field_name = candidate_field
                break
        if raw_model:
            break

    if not raw_model:
        return None

    model_code = canonicalize_model_name(raw_model) or raw_model
    model_label = _MODEL_LABELS.get(model_code, raw_model)
    return {
        "label": model_label,
        "code": model_code,
        "raw": raw_model,
        "source": f"{source}.{field_name}" if source and field_name else "",
    }
