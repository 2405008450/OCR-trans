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

_GOOGLE_RETRYABLE_ERROR_MARKERS = (
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
GEMINI_ROUTE_AI_STUDIO = "google_ai_studio"

_ALL_GEMINI_ROUTES = (
    GEMINI_ROUTE_GOOGLE,
    GEMINI_ROUTE_OPENROUTER,
    GEMINI_ROUTE_AI_STUDIO,
)

DEFAULT_GEMINI_ROUTE = (
    settings.GEMINI_DEFAULT_ROUTE if settings.GEMINI_DEFAULT_ROUTE in _ALL_GEMINI_ROUTES else GEMINI_ROUTE_GOOGLE
)

GEMINI_ROUTE_OPTIONS: Dict[str, Dict[str, str]] = {
    GEMINI_ROUTE_GOOGLE: {
        "label": "\u7ebf\u8def1",
        "description": "\u8c37\u6b4c Vertex \u5b98\u65b9\u7ebf\u8def\uff0c\u901f\u5ea6\u5feb\uff0c\u9002\u5408\u5e38\u89c4\u4efb\u52a1\u3002",
    },
    GEMINI_ROUTE_OPENROUTER: {
        "label": "\u7ebf\u8def2",
        "description": "OpenRouter \u4e2d\u8f6c\u7ebf\u8def\uff0c\u9002\u5408\u5f53\u524d\u5df2\u6709\u7684\u517c\u5bb9\u8c03\u7528\u3002",
    },
    GEMINI_ROUTE_AI_STUDIO: {
        "label": "\u7ebf\u8def3",
        "description": "Google AI Studio \u5b98\u65b9\u7ebf\u8def\uff0c\u9002\u5408\u8865\u5145\u5927\u6587\u4ef6\u6216 Vertex \u4e0d\u4fbf\u5904\u7406\u7684\u573a\u666f\u3002",
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
    if normalized_route in (GEMINI_ROUTE_GOOGLE, GEMINI_ROUTE_AI_STUDIO):
        return normalize_google_model(model)
    if "/" not in model:
        return f"google/{model}"
    return model


def ensure_gemini_route_configured(route: Optional[str]) -> str:
    normalized = normalize_gemini_route(route)
    if normalized == GEMINI_ROUTE_GOOGLE:
        if not settings.VERTEX_PROJECT_ID:
            raise ValueError("??? VERTEX_PROJECT_ID????? Google Vertex AI ??")
        return normalized
    if normalized == GEMINI_ROUTE_AI_STUDIO:
        if not settings.AI_STUDIO_API_KEY:
            raise ValueError("??? AI_STUDIO_API_KEY ? GEMINI_API_KEY????? Google AI Studio ??")
        return normalized
    if not settings.OPENROUTER_API_KEY:
        raise ValueError("??? OPENROUTER_API_KEY????? OpenRouter Gemini ??")
    return normalized


def _get_vertex_client(timeout: float = 600.0) -> genai.Client:
    return genai.Client(
        vertexai=True,
        project=settings.VERTEX_PROJECT_ID,
        location=settings.VERTEX_LOCATION,
        http_options=types.HttpOptions(timeout=int(timeout * 1000)),
    )


def _get_ai_studio_client(timeout: float = 600.0) -> genai.Client:
    return genai.Client(
        api_key=settings.AI_STUDIO_API_KEY,
        http_options=types.HttpOptions(timeout=int(timeout * 1000)),
    )


def _is_retryable_google_client_error(exc: ClientError) -> bool:
    message = str(exc).lower()
    return any(marker in message for marker in _GOOGLE_RETRYABLE_ERROR_MARKERS)


def _generate_google_with_retry(
    *,
    route_name: str,
    client_factory,
    model: str,
    contents,
    config=None,
    max_retries: int = 2,
    timeout: float = 600.0,
    log_callback: GeminiLogCallback = None,
) -> str:
    client = client_factory(timeout=timeout)
    delay = 2.0
    for attempt in range(max_retries):
        try:
            response_stream = client.models.generate_content_stream(model=model, contents=contents, config=config)
            full_text = ""
            char_count = 0
            last_log_at = 0
            
            for chunk in response_stream:
                chunk_text = getattr(chunk, "text", None)
                if not chunk_text:
                    parts = []
                    for candidate in getattr(chunk, "candidates", []) or []:
                        content = getattr(candidate, "content", None)
                        if content:
                            for part in getattr(content, "parts", []) or []:
                                part_text = getattr(part, "text", None)
                                if part_text:
                                    parts.append(part_text)
                    chunk_text = "".join(parts)
                
                if chunk_text:
                    full_text += chunk_text
                    char_count += len(chunk_text)
                    
                    if log_callback and (char_count - last_log_at) >= 200:
                        log_callback(f"[{route_name}] 生成中... 已接收 {char_count} 字符")
                        last_log_at = char_count
                        
            if log_callback and char_count > 0:
                log_callback(f"[{route_name}] 生成完毕，共 {char_count} 字符")
                
            return full_text
            
        except Exception as exc:
            exc_name = type(exc).__name__
            is_network_err = (
                exc_name in ("TransportError", "ConnectionError", "TimeoutError", "ProtocolError", "OSError", "Timeout", "ConnectError", "ReadTimeout", "WriteTimeout", "ConnectionResetError") 
                or any(isinstance(exc, err) for err in _RETRYABLE_NETWORK_ERRORS)
            )
            
            if isinstance(exc, ClientError):
                if attempt == max_retries - 1 or not _is_retryable_google_client_error(exc):
                    raise
                sleep_s = delay + random.uniform(0, 1.5)
                if log_callback:
                    log_callback(f"[{route_name}] 客户端异常 ({exc_name})等待 {sleep_s:.1f}s 重试... ({attempt + 1}/{max_retries})")
                time.sleep(sleep_s)
                delay = min(delay * 2, 30)
                client = client_factory(timeout=timeout)
            elif is_network_err:
                if attempt == max_retries - 1:
                    raise
                sleep_s = delay + random.uniform(0, 2.0)
                if log_callback:
                    log_callback(f"[{route_name}] 网络异常 ({exc_name})等待 {sleep_s:.1f}s 重试... ({attempt + 1}/{max_retries})")
                time.sleep(sleep_s)
                delay = min(delay * 2, 30)
                client = client_factory(timeout=timeout)
            else:
                raise


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
            log_callback(f"[openrouter] ????... ???? {char_count} ??")
            last_log_at = char_count
    if log_callback:
        log_callback(f"[openrouter] ?????? {char_count} ??")
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
            text_result = _generate_google_with_retry(
                route_name="vertex",
                client_factory=_get_vertex_client,
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
            return (text_result or "").strip()
        except Exception as exc:
            if not _should_fallback_to_openrouter(normalized):
                raise
            fallback_model = resolve_model_for_route(model, GEMINI_ROUTE_OPENROUTER)
            if log_callback:
                log_callback(f"[gemini] Vertex ??????? OpenRouter: {type(exc).__name__}: {exc}")
            return _generate_openrouter_text(
                system_prompt=system_prompt,
                user_prompt=user_prompt,
                model=fallback_model,
                temperature=temperature,
                max_output_tokens=max_output_tokens,
                timeout=timeout,
                log_callback=log_callback,
            )

    if normalized == GEMINI_ROUTE_AI_STUDIO:
        text_result = _generate_google_with_retry(
            route_name="ai-studio",
            client_factory=_get_ai_studio_client,
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
        return (text_result or "").strip()

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
    user_prompt: str = "??????????? OCR ??? HTML?",
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
            text_result = _generate_google_with_retry(
                route_name="vertex",
                client_factory=_get_vertex_client,
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
            return (text_result or "").strip()
        except Exception as exc:
            if not _should_fallback_to_openrouter(normalized):
                raise
            fallback_model = resolve_model_for_route(model, GEMINI_ROUTE_OPENROUTER)
            if log_callback:
                log_callback(f"[gemini-vision] Vertex ??????? OpenRouter: {type(exc).__name__}: {exc}")
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

    if normalized == GEMINI_ROUTE_AI_STUDIO:
        text_result = _generate_google_with_retry(
            route_name="ai-studio",
            client_factory=_get_ai_studio_client,
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
        return (text_result or "").strip()

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
