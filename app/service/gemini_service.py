from __future__ import annotations

import random
import time
from typing import Callable, Dict, Optional

from google import genai
from google.genai import types
from google.genai.errors import ClientError
from openai import OpenAI

from app.core.config import settings

GeminiLogCallback = Optional[Callable[[str], None]]

GEMINI_ROUTE_GOOGLE = "google"
GEMINI_ROUTE_OPENROUTER = "openrouter"
# 读取 .env 中的 GEMINI_DEFAULT_ROUTE，fallback 到 google
DEFAULT_GEMINI_ROUTE = (
    settings.GEMINI_DEFAULT_ROUTE
    if settings.GEMINI_DEFAULT_ROUTE in (GEMINI_ROUTE_GOOGLE, GEMINI_ROUTE_OPENROUTER)
    else GEMINI_ROUTE_GOOGLE
)

GEMINI_ROUTE_OPTIONS: Dict[str, Dict[str, str]] = {
    GEMINI_ROUTE_GOOGLE: {
        "label": "线路1 (推荐)",
        "description": "主线路，直连官方 API，速度更快、更稳定。",
    },
    GEMINI_ROUTE_OPENROUTER: {
        "label": "线路2",
        "description": "备用线路，通过中转服务转发，适合主线路不可用时使用。",
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
        if not settings.VERTEX_PROJECT_ID:
            raise ValueError("未配置 VERTEX_PROJECT_ID，无法使用 Google Vertex AI 线路")
        return normalized
    if not settings.OPENROUTER_API_KEY:
        raise ValueError("未配置 OPENROUTER_API_KEY，无法使用 OpenRouter Gemini 线路")
    return normalized


def _get_vertex_client() -> genai.Client:
    return genai.Client(
        vertexai=True,
        project=settings.VERTEX_PROJECT_ID,
        location=settings.VERTEX_LOCATION,
    )


def _generate_with_retry(
    model: str,
    contents,
    config=None,
    max_retries: int = 6,
    log_callback: GeminiLogCallback = None,
):
    client = _get_vertex_client()
    delay = 2.0
    for attempt in range(max_retries):
        try:
            return client.models.generate_content(
                model=model,
                contents=contents,
                config=config,
            )
        except ClientError as e:
            msg = str(e)
            if "429" in msg or "RESOURCE_EXHAUSTED" in msg:
                if attempt == max_retries - 1:
                    raise
                sleep_s = delay + random.uniform(0, 1.5)
                if log_callback:
                    log_callback(f"[vertex] 429 限流，等待 {sleep_s:.1f}s 后重试... ({attempt + 1}/{max_retries})")
                time.sleep(sleep_s)
                delay *= 2
            else:
                raise


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
        response = _generate_with_retry(
            model=resolved_model,
            contents=user_prompt,
            config=types.GenerateContentConfig(
                system_instruction=system_prompt,
                temperature=temperature,
                max_output_tokens=max_output_tokens,
            ),
            log_callback=log_callback,
        )
        return (_extract_google_text(response) or "").strip()

    client = OpenAI(base_url=settings.OPENROUTER_BASE_URL, api_key=settings.OPENROUTER_API_KEY, timeout=timeout)
    stream = client.chat.completions.create(
        model=resolved_model,
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
    token_count = 0
    last_log_at = 0
    for chunk in stream:
        delta = (chunk.choices[0].delta.content or "") if chunk.choices else ""
        if delta:
            full_text += delta
            token_count += len(delta)
            if log_callback and token_count - last_log_at >= 200:
                log_callback(f"[stream] 正在生成... 已收到约 {token_count} 字符")
                last_log_at = token_count
    if log_callback:
        log_callback(f"[stream] 生成完成，共 {token_count} 字符")
    return full_text.strip()


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
            log_callback=log_callback,
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
