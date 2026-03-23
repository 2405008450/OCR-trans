from __future__ import annotations

from typing import Callable, Dict, Optional

from google import genai
from google.genai import types
from openai import OpenAI

from app.core.config import settings

GeminiLogCallback = Optional[Callable[[str], None]]

GEMINI_ROUTE_GOOGLE = "google"
GEMINI_ROUTE_OPENROUTER = "openrouter"
DEFAULT_GEMINI_ROUTE = GEMINI_ROUTE_GOOGLE

GEMINI_ROUTE_OPTIONS: Dict[str, Dict[str, str]] = {
    GEMINI_ROUTE_GOOGLE: {
        "label": "线路1",
        "description": "默认主线路，直连 Google 官方 Gemini API。",
    },
    GEMINI_ROUTE_OPENROUTER: {
        "label": "线路2",
        "description": "备用线路，通过 OpenRouter 转发 Gemini 模型。",
    },
}


def get_gemini_routes() -> Dict[str, Dict[str, str]]:
    return GEMINI_ROUTE_OPTIONS


def normalize_gemini_route(route: Optional[str]) -> str:
    if route in GEMINI_ROUTE_OPTIONS:
        return route
    return settings.GEMINI_DEFAULT_ROUTE if settings.GEMINI_DEFAULT_ROUTE in GEMINI_ROUTE_OPTIONS else DEFAULT_GEMINI_ROUTE


def normalize_google_model(model: str) -> str:
    if model.startswith("google/"):
        return model.split("/", 1)[1]
    return model


def resolve_model_for_route(model: str, route: Optional[str]) -> str:
    normalized_route = normalize_gemini_route(route)
    if normalized_route == GEMINI_ROUTE_GOOGLE:
        return normalize_google_model(model)
    if "/" not in model:
        return f"google/{model}"
    return model


def ensure_gemini_route_configured(route: Optional[str]) -> str:
    normalized = normalize_gemini_route(route)
    if normalized == GEMINI_ROUTE_GOOGLE:
        if not settings.GOOGLE_API_KEY:
            raise ValueError("未配置 GOOGLE_API_KEY，无法使用 Google 官方 Gemini 线路")
        return normalized
    if not settings.OPENROUTER_API_KEY:
        raise ValueError("未配置 OPENROUTER_API_KEY，无法使用 OpenRouter Gemini 线路")
    return normalized


def _extract_google_text(response) -> str:
    text = getattr(response, "text", None)
    if text:
        return text

    parts = []
    for candidate in getattr(response, "candidates", []) or []:
        content = getattr(candidate, "content", None)
        for part in getattr(content, "parts", []) or []:
            part_text = getattr(part, "text", None)
            if part_text:
                parts.append(part_text)
    return "".join(parts)


def generate_text(
    *,
    system_prompt: str,
    user_prompt: str,
    model: str,
    route: Optional[str] = None,
    temperature: float = 0.1,
    max_output_tokens: int = 65536,
    timeout: float = 600.0,
    log_callback: GeminiLogCallback = None,
) -> str:
    normalized = ensure_gemini_route_configured(route)
    resolved_model = resolve_model_for_route(model, normalized)
    if log_callback:
        log_callback(f"[gemini] route={normalized}, model={resolved_model}")

    if normalized == GEMINI_ROUTE_GOOGLE:
        client = genai.Client(api_key=settings.GOOGLE_API_KEY)
        response = client.models.generate_content(
            model=resolved_model,
            contents=user_prompt,
            config=types.GenerateContentConfig(
                system_instruction=system_prompt,
                temperature=temperature,
                max_output_tokens=max_output_tokens,
            ),
        )
        return (_extract_google_text(response) or "").strip()

    client = OpenAI(base_url=settings.OPENROUTER_BASE_URL, api_key=settings.OPENROUTER_API_KEY, timeout=timeout)
    response = client.chat.completions.create(
        model=resolved_model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        temperature=temperature,
        max_tokens=max_output_tokens,
        extra_headers={"HTTP-Referer": "local-debug", "X-Title": "fastapi-llm-demo"},
    )
    return (response.choices[0].message.content or "").strip()


def generate_vision_html(
    *,
    system_prompt: str,
    image_bytes: bytes,
    mime_type: str,
    model: str,
    route: Optional[str] = None,
    user_prompt: str = "请严格根据图片内容执行OCR并输出HTML。",
    temperature: float = 0.0,
    max_output_tokens: int = 65536,
    timeout: float = 600.0,
    log_callback: GeminiLogCallback = None,
) -> str:
    normalized = ensure_gemini_route_configured(route)
    resolved_model = resolve_model_for_route(model, normalized)
    if log_callback:
        log_callback(f"[gemini-vision] route={normalized}, model={resolved_model}, mime={mime_type}")

    if normalized == GEMINI_ROUTE_GOOGLE:
        client = genai.Client(api_key=settings.GOOGLE_API_KEY)
        response = client.models.generate_content(
            model=resolved_model,
            contents=[
                types.Content(
                    role="user",
                    parts=[
                        types.Part.from_text(text=user_prompt),
                        types.Part.from_bytes(data=image_bytes, mime_type=mime_type),
                    ],
                )
            ],
            config=types.GenerateContentConfig(
                system_instruction=system_prompt,
                temperature=temperature,
                max_output_tokens=max_output_tokens,
            ),
        )
        return (_extract_google_text(response) or "").strip()

    client = OpenAI(base_url=settings.OPENROUTER_BASE_URL, api_key=settings.OPENROUTER_API_KEY, timeout=timeout)
    image_b64 = image_bytes.decode("utf-8") if mime_type == "text/plain-base64" else None
    if image_b64 is None:
        import base64

        image_b64 = base64.standard_b64encode(image_bytes).decode("utf-8")

    response = client.chat.completions.create(
        model=resolved_model,
        messages=[
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": user_prompt},
                    {"type": "image_url", "image_url": {"url": f"data:{mime_type};base64,{image_b64}"}},
                ],
            },
        ],
        temperature=temperature,
        max_tokens=max_output_tokens,
        extra_headers={"HTTP-Referer": "local-debug", "X-Title": "fastapi-llm-demo"},
    )
    return (response.choices[0].message.content or "").strip()
