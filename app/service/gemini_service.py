from __future__ import annotations

import base64
import random
import time
from typing import Callable, Dict, Optional

import requests.exceptions
import urllib3.exceptions
from google import genai
from google.genai import types
from google.genai.errors import ClientError
from openai import OpenAI

from app.core.config import settings

_RETRYABLE_NETWORK_ERRORS = (
    ConnectionError,
    TimeoutError,
    requests.exceptions.ConnectionError,
    requests.exceptions.Timeout,
    requests.exceptions.ChunkedEncodingError,
    urllib3.exceptions.ProtocolError,
    urllib3.exceptions.TimeoutError,
    OSError,
)

_VERTEX_RETRYABLE_ERROR_MARKERS = (
    "429",
    "500",
    "502",
    "503",
    "504",
    "resource_exhausted",
    "internal",
    "unavailable",
    "deadline_exceeded",
    "connection reset",
    "connection aborted",
    "timed out",
    "goaway",
)

GeminiLogCallback = Optional[Callable[[str], None]]

GEMINI_ROUTE_GOOGLE = "google"
GEMINI_ROUTE_OPENROUTER = "openrouter"

DEFAULT_GEMINI_ROUTE = (
    settings.GEMINI_DEFAULT_ROUTE
    if settings.GEMINI_DEFAULT_ROUTE in (GEMINI_ROUTE_GOOGLE, GEMINI_ROUTE_OPENROUTER)
    else GEMINI_ROUTE_GOOGLE
)

GEMINI_ROUTE_OPTIONS: Dict[str, Dict[str, str]] = {
    GEMINI_ROUTE_GOOGLE: {
        "label": "线路1 (推荐)",
        "description": "直连 Google Vertex AI，速度快；失败时会自动回退到 OpenRouter。",
    },
    GEMINI_ROUTE_OPENROUTER: {
        "label": "线路2",
        "description": "经 OpenRouter 中转，通常更宽松，适合作为备选线路。",
    },
}


def get_gemini_routes() -> Dict[str, Dict[str, str]]:
    return GEMINI_ROUTE_OPTIONS


def normalize_gemini_route(route: Optional[str]) -> str:
    if route in GEMINI_ROUTE_OPTIONS:
        return route
    if settings.GEMINI_DEFAULT_ROUTE in GEMINI_ROUTE_OPTIONS:
        return settings.GEMINI_DEFAULT_ROUTE
    return DEFAULT_GEMINI_ROUTE


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
        if not settings.VERTEX_PROJECT_ID:
            raise ValueError("未配置 VERTEX_PROJECT_ID，无法使用 Google Vertex AI 线路")
        return normalized
    if not settings.OPENROUTER_API_KEY:
        raise ValueError("未配置 OPENROUTER_API_KEY，无法使用 OpenRouter Gemini 线路")
    return normalized


def _get_vertex_client(timeout: float = 600.0) -> genai.Client:
    return genai.Client(
        vertexai=True,
        project=settings.VERTEX_PROJECT_ID,
        location=settings.VERTEX_LOCATION,
        http_options=types.HttpOptions(timeout=int(timeout * 1000)),
    )


def _is_retryable_vertex_client_error(exc: ClientError) -> bool:
    message = str(exc).lower()
    return any(marker in message for marker in _VERTEX_RETRYABLE_ERROR_MARKERS)


def _generate_with_retry(
    model: str,
    contents,
    config=None,
    max_retries: int = 2,
    timeout: float = 600.0,
    log_callback: GeminiLogCallback = None,
):
    client = _get_vertex_client(timeout=timeout)
    delay = 2.0
    for attempt in range(max_retries):
        try:
            return client.models.generate_content(
                model=model,
                contents=contents,
                config=config,
            )
        except ClientError as exc:
            if attempt == max_retries - 1 or not _is_retryable_vertex_client_error(exc):
                raise
            sleep_s = delay + random.uniform(0, 1.5)
            if log_callback:
                log_callback(
                    f"[vertex] 可重试错误 ({type(exc).__name__})，等待 {sleep_s:.1f}s 后重试... ({attempt + 1}/{max_retries})"
                )
            time.sleep(sleep_s)
            delay = min(delay * 2, 30)
            client = _get_vertex_client(timeout=timeout)
        except _RETRYABLE_NETWORK_ERRORS as exc:
            if attempt == max_retries - 1:
                raise
            sleep_s = delay + random.uniform(0, 2.0)
            if log_callback:
                log_callback(
                    f"[vertex] 网络异常 ({type(exc).__name__})，等待 {sleep_s:.1f}s 后重试... ({attempt + 1}/{max_retries})"
                )
            time.sleep(sleep_s)
            delay = min(delay * 2, 30)
            client = _get_vertex_client(timeout=timeout)


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


def _should_fallback_to_openrouter(route: str) -> bool:
    return route == GEMINI_ROUTE_GOOGLE and bool(settings.OPENROUTER_API_KEY)


def _generate_openrouter_text(
    *,
    system_prompt: str,
    user_prompt: str,
    model: str,
    temperature: float,
    max_output_tokens: int,
    timeout: float,
    log_callback: GeminiLogCallback = None,
) -> str:
    client = OpenAI(
        base_url=settings.OPENROUTER_BASE_URL,
        api_key=settings.OPENROUTER_API_KEY,
        timeout=timeout,
    )
    stream = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        temperature=temperature,
        max_tokens=max_output_tokens,
        extra_headers={"HTTP-Referer": "local-debug", "X-Title": "fastapi-llm-demo"},
        stream=True,
    )
    full_text = ""
    char_count = 0
    last_log_at = 0
    for chunk in stream:
        delta = (chunk.choices[0].delta.content or "") if chunk.choices else ""
        if not delta:
            continue
        full_text += delta
        char_count += len(delta)
        if log_callback and char_count - last_log_at >= 200:
            log_callback(f"[openrouter] 正在生成... 已收到约 {char_count} 字符")
            last_log_at = char_count
    if log_callback:
        log_callback(f"[openrouter] 生成完成，共 {char_count} 字符")
    return full_text.strip()


def _generate_openrouter_vision(
    *,
    system_prompt: str,
    user_prompt: str,
    image_bytes: bytes,
    mime_type: str,
    model: str,
    temperature: float,
    max_output_tokens: int,
    timeout: float,
) -> str:
    client = OpenAI(
        base_url=settings.OPENROUTER_BASE_URL,
        api_key=settings.OPENROUTER_API_KEY,
        timeout=timeout,
    )
    image_b64 = image_bytes.decode("utf-8") if mime_type == "text/plain-base64" else None
    if image_b64 is None:
        image_b64 = base64.standard_b64encode(image_bytes).decode("utf-8")

    response = client.chat.completions.create(
        model=model,
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
        try:
            response = _generate_with_retry(
                model=resolved_model,
                contents=user_prompt,
                config=types.GenerateContentConfig(
                    system_instruction=system_prompt,
                    temperature=temperature,
                    max_output_tokens=max_output_tokens,
                ),
                timeout=timeout,
                log_callback=log_callback,
            )
            return (_extract_google_text(response) or "").strip()
        except Exception as exc:
            if not _should_fallback_to_openrouter(normalized):
                raise
            fallback_model = resolve_model_for_route(model, GEMINI_ROUTE_OPENROUTER)
            if log_callback:
                log_callback(f"[gemini] Vertex 失败，自动回退 OpenRouter: {type(exc).__name__}: {exc}")
            return _generate_openrouter_text(
                system_prompt=system_prompt,
                user_prompt=user_prompt,
                model=fallback_model,
                temperature=temperature,
                max_output_tokens=max_output_tokens,
                timeout=timeout,
                log_callback=log_callback,
            )

    return _generate_openrouter_text(
        system_prompt=system_prompt,
        user_prompt=user_prompt,
        model=resolved_model,
        temperature=temperature,
        max_output_tokens=max_output_tokens,
        timeout=timeout,
        log_callback=log_callback,
    )


def generate_vision_html(
    *,
    system_prompt: str,
    image_bytes: bytes,
    mime_type: str,
    model: str,
    route: Optional[str] = None,
    user_prompt: str = "请严格根据图片内容执行 OCR 并输出 HTML。",
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
        try:
            response = _generate_with_retry(
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
                timeout=timeout,
                log_callback=log_callback,
            )
            return (_extract_google_text(response) or "").strip()
        except Exception as exc:
            if not _should_fallback_to_openrouter(normalized):
                raise
            fallback_model = resolve_model_for_route(model, GEMINI_ROUTE_OPENROUTER)
            if log_callback:
                log_callback(f"[gemini-vision] Vertex 失败，自动回退 OpenRouter: {type(exc).__name__}: {exc}")
            return _generate_openrouter_vision(
                system_prompt=system_prompt,
                user_prompt=user_prompt,
                image_bytes=image_bytes,
                mime_type=mime_type,
                model=fallback_model,
                temperature=temperature,
                max_output_tokens=max_output_tokens,
                timeout=timeout,
            )

    return _generate_openrouter_vision(
        system_prompt=system_prompt,
        user_prompt=user_prompt,
        image_bytes=image_bytes,
        mime_type=mime_type,
        model=resolved_model,
        temperature=temperature,
        max_output_tokens=max_output_tokens,
        timeout=timeout,
    )
