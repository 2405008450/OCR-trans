from __future__ import annotations

import base64
import math
import os
import random
import time
from typing import Callable, Dict, List, Optional, Sequence

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
GEMINI_ROUTE_AISTUDIO = "aistudio"
GEMINI_ROUTE_OPENROUTER = "openrouter"

_ALL_GEMINI_ROUTES = (
    GEMINI_ROUTE_GOOGLE,
    GEMINI_ROUTE_AISTUDIO,
    GEMINI_ROUTE_OPENROUTER,
)

DEFAULT_GEMINI_ROUTE = (
    settings.GEMINI_DEFAULT_ROUTE if settings.GEMINI_DEFAULT_ROUTE in _ALL_GEMINI_ROUTES else GEMINI_ROUTE_GOOGLE
)
DEFAULT_GEMINI_TIMEOUT_SECONDS = float(settings.GEMINI_TIMEOUT_SECONDS)
DEFAULT_EMBEDDING_MODEL = settings.GEMINI_EMBEDDING_MODEL or "gemini-embedding-001"
DEFAULT_EMBEDDING_DIMENSIONS = int(settings.GEMINI_EMBEDDING_DIMENSIONS or 768)
OPENROUTER_EMBEDDING_MODEL = (
    DEFAULT_EMBEDDING_MODEL
    if "/" in DEFAULT_EMBEDDING_MODEL
    else f"google/{DEFAULT_EMBEDDING_MODEL}"
)

GEMINI_ROUTE_OPTIONS: Dict[str, Dict[str, str]] = {
    GEMINI_ROUTE_GOOGLE: {
        "label": "\u7ebf\u8def1",
        "description": "\u8c37\u6b4c Vertex \u5b98\u65b9\u7ebf\u8def\uff0c\u901f\u5ea6\u5feb\uff0c\u9002\u5408\u5e38\u89c4\u4efb\u52a1\u3002",
    },
    GEMINI_ROUTE_AISTUDIO: {
        "label": "Google AI Studio",
        "description": "\u4f7f\u7528 GOOGLE_API_KEY \u76f4\u8fde Google AI Studio\uff0c\u9002\u5408\u6ca1\u6709 Vertex \u8ba4\u8bc1\u7684\u670d\u52a1\u5668\u3002",
    },
    GEMINI_ROUTE_OPENROUTER: {
        "label": "\u7ebf\u8def2",
        "description": "OpenRouter \u4e2d\u8f6c\u7ebf\u8def\uff0c\u9002\u5408\u5f53\u524d\u5df2\u6709\u7684\u517c\u5bb9\u8c03\u7528\u3002",
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
    if normalized_route in {GEMINI_ROUTE_GOOGLE, GEMINI_ROUTE_AISTUDIO}:
        return normalize_google_model(model)
    if "/" not in model:
        return f"google/{model}"
    return model


def ensure_gemini_route_configured(route: Optional[str]) -> str:
    normalized = normalize_gemini_route(route)
    if normalized == GEMINI_ROUTE_GOOGLE:
        if not settings.VERTEX_PROJECT_ID:
            raise ValueError("未配置 VERTEX_PROJECT_ID，无法使用 Google Vertex AI。")
        return normalized
    if normalized == GEMINI_ROUTE_AISTUDIO:
        if not settings.GOOGLE_API_KEY:
            raise ValueError("未配置 GOOGLE_API_KEY 或 GEMINI_API_KEY，无法使用 Google AI Studio。")
        return normalized
    if not settings.OPENROUTER_API_KEY:
        raise ValueError("未配置 OPENROUTER_API_KEY，无法使用 OpenRouter Gemini。")
    return normalized


def _get_vertex_client(timeout: float = DEFAULT_GEMINI_TIMEOUT_SECONDS) -> genai.Client:
    return genai.Client(
        vertexai=True,
        project=settings.VERTEX_PROJECT_ID,
        location=settings.VERTEX_LOCATION,
        http_options=types.HttpOptions(timeout=int(timeout * 1000)),
    )


def _get_aistudio_client(timeout: float = DEFAULT_GEMINI_TIMEOUT_SECONDS) -> genai.Client:
    return genai.Client(
        api_key=settings.GOOGLE_API_KEY,
        http_options=types.HttpOptions(timeout=int(timeout * 1000)),
    )


def _is_retryable_google_client_error(exc: ClientError) -> bool:
    message = str(exc).lower()
    return any(marker in message for marker in _GOOGLE_RETRYABLE_ERROR_MARKERS)


def _extract_google_response_text(response) -> str:
    try:
        response_text = getattr(response, "text", None)
    except Exception:
        response_text = None
    if response_text:
        return response_text

    parts = []
    for candidate in getattr(response, "candidates", []) or []:
        content = getattr(candidate, "content", None)
        if not content:
            continue
        for part in getattr(content, "parts", []) or []:
            part_text = getattr(part, "text", None)
            if part_text:
                parts.append(part_text)
    return "".join(parts)


def _generate_google_with_retry(
    *,
    route_name: str,
    client_factory,
    model: str,
    contents,
    config=None,
    max_retries: int = 3,
    timeout: float = DEFAULT_GEMINI_TIMEOUT_SECONDS,
    log_callback: GeminiLogCallback = None,
    stream: bool = True,
) -> str:
    client = client_factory(timeout=timeout)
    delay = 2.0
    for attempt in range(max_retries):
        try:
            if not stream:
                response = client.models.generate_content(model=model, contents=contents, config=config)
                full_text = _extract_google_response_text(response)
                if log_callback:
                    log_callback(f"[{route_name}] 生成完毕，共 {len(full_text)} 字符")
                return full_text

            response_stream = client.models.generate_content_stream(model=model, contents=contents, config=config)
            full_text = ""
            char_count = 0
            last_log_at = 0
            
            for chunk in response_stream:
                chunk_text = _extract_google_response_text(chunk)
                
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
                exc_name in ("TransportError", "ConnectionError", "TimeoutError", "ProtocolError", "OSError", "Timeout", "ConnectError", "ReadTimeout", "WriteTimeout", "ConnectionResetError", "ChunkedEncodingError", "RemoteProtocolError") 
                or any(isinstance(exc, err) for err in _RETRYABLE_NETWORK_ERRORS)
            )
            
            if isinstance(exc, ClientError):
                if attempt == max_retries - 1 or not _is_retryable_google_client_error(exc):
                    raise
                sleep_s = delay + random.uniform(0, 1.5)
                if log_callback:
                    log_callback(
                        f"[{route_name}] 客户端异常（{exc_name}），等待 {sleep_s:.1f}s 后重试... "
                        f"({attempt + 1}/{max_retries})"
                    )
                time.sleep(sleep_s)
                delay = min(delay * 2, 30)
                client = client_factory(timeout=timeout)
            elif is_network_err:
                if attempt == max_retries - 1:
                    raise
                sleep_s = delay + random.uniform(0, 2.0)
                if log_callback:
                    log_callback(
                        f"[{route_name}] 网络异常（{exc_name}），等待 {sleep_s:.1f}s 后重试... "
                        f"({attempt + 1}/{max_retries})"
                    )
                time.sleep(sleep_s)
                delay = min(delay * 2, 30)
                client = client_factory(timeout=timeout)
            else:
                raise


def _should_fallback_to_openrouter(route: str) -> bool:
    return (
        route in {GEMINI_ROUTE_GOOGLE, GEMINI_ROUTE_AISTUDIO}
        and settings.GEMINI_ENABLE_OPENROUTER_FALLBACK_ENABLED
        and bool(settings.OPENROUTER_API_KEY)
    )


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
        default_headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
        }
    )
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt},
    ]
    extra_headers = {"HTTP-Referer": "local-debug", "X-Title": "fastapi-llm-demo"}

    if log_callback:
        log_callback("[openrouter] 使用非流式生成，等待模型完整返回...")
    response = client.chat.completions.create(
        model=model,
        messages=messages,
        temperature=temperature,
        max_tokens=max_output_tokens,
        extra_headers=extra_headers,
        stream=False,
    )
    full_text = (response.choices[0].message.content or "").strip()
    if log_callback:
        log_callback(f"[openrouter] 非流式生成完毕，共 {len(full_text)} 字符")
    return full_text


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
        default_headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
        }
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


def _generate_openrouter_audio(
    *,
    system_prompt: str,
    user_prompt: str,
    audio_bytes: bytes,
    audio_format: str,
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
        default_headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
        },
    )
    audio_b64 = base64.standard_b64encode(audio_bytes).decode("ascii")
    if log_callback:
        log_callback("[openrouter-audio] 正在提交音频并等待模型分析...")
    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": user_prompt},
                    {
                        "type": "input_audio",
                        "input_audio": {
                            "data": audio_b64,
                            "format": audio_format,
                        },
                    },
                ],
            },
        ],
        temperature=temperature,
        max_tokens=max_output_tokens,
        extra_headers={"HTTP-Referer": "local-debug", "X-Title": "fastapi-llm-demo"},
        stream=False,
    )
    full_text = (response.choices[0].message.content or "").strip()
    if log_callback:
        log_callback(f"[openrouter-audio] 分析完成，共 {len(full_text)} 个字符")
    return full_text


def generate_text(
    *,
    system_prompt: str,
    user_prompt: str,
    model: str,
    route: Optional[str] = None,
    temperature: float = 0.1,
    max_output_tokens: int = 65536,
    timeout: float = DEFAULT_GEMINI_TIMEOUT_SECONDS,
    log_callback: GeminiLogCallback = None,
) -> str:
    normalized = ensure_gemini_route_configured(route)
    resolved_model = resolve_model_for_route(model, normalized)
    if log_callback:
        log_callback(f"[gemini] route={normalized}, model={resolved_model}")

    config_kwargs = {
        "temperature": temperature,
        "max_output_tokens": max_output_tokens,
    }
    if (system_prompt or "").strip():
        config_kwargs["system_instruction"] = system_prompt

    if normalized in {GEMINI_ROUTE_GOOGLE, GEMINI_ROUTE_AISTUDIO}:
        route_name = "vertex" if normalized == GEMINI_ROUTE_GOOGLE else "aistudio"
        client_factory = _get_vertex_client if normalized == GEMINI_ROUTE_GOOGLE else _get_aistudio_client
        try:
            text_result = _generate_google_with_retry(
                route_name=route_name,
                client_factory=client_factory,
                model=resolved_model,
                contents=user_prompt,
                config=types.GenerateContentConfig(**config_kwargs),
                timeout=timeout,
                log_callback=log_callback,
                stream=normalized != GEMINI_ROUTE_AISTUDIO,
            )
            return (text_result or "").strip()
        except Exception as exc:
            if not _should_fallback_to_openrouter(normalized):
                raise
            fallback_model = resolve_model_for_route(model, GEMINI_ROUTE_OPENROUTER)
            if log_callback:
                log_callback(f"[gemini] {route_name} 失败，回退到 OpenRouter: {type(exc).__name__}: {exc}")
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
    user_prompt: str = "请根据上传图片执行 OCR 并输出 HTML。",
    temperature: float = 0.0,
    max_output_tokens: int = 65536,
    timeout: float = DEFAULT_GEMINI_TIMEOUT_SECONDS,
    log_callback: GeminiLogCallback = None,
) -> str:
    normalized = ensure_gemini_route_configured(route)
    resolved_model = resolve_model_for_route(model, normalized)
    if log_callback:
        log_callback(f"[gemini-vision] route={normalized}, model={resolved_model}, mime={mime_type}")

    config_kwargs = {
        "temperature": temperature,
        "max_output_tokens": max_output_tokens,
    }
    if (system_prompt or "").strip():
        config_kwargs["system_instruction"] = system_prompt

    parts = [types.Part.from_bytes(data=image_bytes, mime_type=mime_type)]
    if (user_prompt or "").strip():
        parts.insert(0, types.Part.from_text(text=user_prompt))

    if normalized in {GEMINI_ROUTE_GOOGLE, GEMINI_ROUTE_AISTUDIO}:
        route_name = "vertex" if normalized == GEMINI_ROUTE_GOOGLE else "aistudio"
        client_factory = _get_vertex_client if normalized == GEMINI_ROUTE_GOOGLE else _get_aistudio_client
        try:
            text_result = _generate_google_with_retry(
                route_name=route_name,
                client_factory=client_factory,
                model=resolved_model,
                contents=[
                    types.Content(
                        role="user",
                        parts=parts,
                    )
                ],
                config=types.GenerateContentConfig(**config_kwargs),
                timeout=timeout,
                log_callback=log_callback,
                stream=normalized != GEMINI_ROUTE_AISTUDIO,
            )
            return (text_result or "").strip()
        except Exception as exc:
            if not _should_fallback_to_openrouter(normalized):
                raise
            fallback_model = resolve_model_for_route(model, GEMINI_ROUTE_OPENROUTER)
            if log_callback:
                log_callback(f"[gemini-vision] {route_name} 失败，回退到 OpenRouter: {type(exc).__name__}: {exc}")
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


def generate_audio_analysis(
    *,
    system_prompt: str,
    user_prompt: str,
    audio_bytes: bytes,
    mime_type: str,
    audio_format: str,
    model: str,
    route: Optional[str] = None,
    temperature: float = 0.1,
    max_output_tokens: int = 32768,
    timeout: float = DEFAULT_GEMINI_TIMEOUT_SECONDS,
    log_callback: GeminiLogCallback = None,
) -> str:
    normalized = ensure_gemini_route_configured(route)
    resolved_model = resolve_model_for_route(model, normalized)
    if log_callback:
        log_callback(
            f"[gemini-audio] route={normalized}, model={resolved_model}, "
            f"mime={mime_type}, bytes={len(audio_bytes)}"
        )

    config_kwargs = {
        "temperature": temperature,
        "max_output_tokens": max_output_tokens,
    }
    if (system_prompt or "").strip():
        config_kwargs["system_instruction"] = system_prompt

    parts = [
        types.Part.from_text(text=user_prompt),
        types.Part.from_bytes(data=audio_bytes, mime_type=mime_type),
    ]
    if normalized in {GEMINI_ROUTE_GOOGLE, GEMINI_ROUTE_AISTUDIO}:
        route_name = "vertex" if normalized == GEMINI_ROUTE_GOOGLE else "aistudio"
        client_factory = _get_vertex_client if normalized == GEMINI_ROUTE_GOOGLE else _get_aistudio_client
        try:
            text_result = _generate_google_with_retry(
                route_name=route_name,
                client_factory=client_factory,
                model=resolved_model,
                contents=[types.Content(role="user", parts=parts)],
                config=types.GenerateContentConfig(**config_kwargs),
                timeout=timeout,
                log_callback=log_callback,
                stream=normalized != GEMINI_ROUTE_AISTUDIO,
            )
            return (text_result or "").strip()
        except Exception as exc:
            if not _should_fallback_to_openrouter(normalized):
                raise
            fallback_model = resolve_model_for_route(model, GEMINI_ROUTE_OPENROUTER)
            if log_callback:
                log_callback(
                    f"[gemini-audio] {route_name} 失败，回退到 OpenRouter: "
                    f"{type(exc).__name__}: {exc}"
                )
            return _generate_openrouter_audio(
                system_prompt=system_prompt,
                user_prompt=user_prompt,
                audio_bytes=audio_bytes,
                audio_format=audio_format,
                model=fallback_model,
                temperature=temperature,
                max_output_tokens=max_output_tokens,
                timeout=timeout,
                log_callback=log_callback,
            )

    return _generate_openrouter_audio(
        system_prompt=system_prompt,
        user_prompt=user_prompt,
        audio_bytes=audio_bytes,
        audio_format=audio_format,
        model=resolved_model,
        temperature=temperature,
        max_output_tokens=max_output_tokens,
        timeout=timeout,
        log_callback=log_callback,
    )


def _l2_normalize(vector: Sequence[float]) -> List[float]:
    norm = math.sqrt(sum(float(v) * float(v) for v in vector))
    if norm <= 0:
        return [0.0] * len(vector)
    return [float(v) / norm for v in vector]


def _extract_google_embeddings(response) -> List[List[float]]:
    embeddings: List[List[float]] = []
    items = getattr(response, "embeddings", None)
    if items:
        for item in items:
            values = getattr(item, "values", None)
            if values is None and isinstance(item, dict):
                values = item.get("values")
            if values is not None:
                embeddings.append(_l2_normalize(values))
        return embeddings

    # Single embedding response shape
    embedding = getattr(response, "embedding", None)
    values = getattr(embedding, "values", None) if embedding is not None else None
    if values is not None:
        return [_l2_normalize(values)]
    return embeddings


def _embed_google_batch(
    *,
    client: genai.Client,
    model: str,
    texts: Sequence[str],
    task_type: str,
    dimensions: int,
) -> List[List[float]]:
    config_kwargs = {
        "task_type": task_type,
        "output_dimensionality": dimensions,
    }
    # Prefer batch contents when supported; fall back to one-by-one.
    try:
        response = client.models.embed_content(
            model=model,
            contents=list(texts),
            config=types.EmbedContentConfig(**config_kwargs),
        )
        vectors = _extract_google_embeddings(response)
        if len(vectors) == len(texts):
            return vectors
    except Exception:
        pass

    vectors = []
    for text in texts:
        response = client.models.embed_content(
            model=model,
            contents=text,
            config=types.EmbedContentConfig(**config_kwargs),
        )
        item_vectors = _extract_google_embeddings(response)
        if not item_vectors:
            raise RuntimeError("embedding response missing values")
        vectors.append(item_vectors[0])
    return vectors


def _embed_openrouter_batch(
    *,
    texts: Sequence[str],
    model: str,
    dimensions: int,
    timeout: float,
    log_callback: GeminiLogCallback = None,
) -> List[List[float]]:
    if not settings.OPENROUTER_API_KEY:
        raise ValueError("未配置 OPENROUTER_API_KEY，无法使用 OpenRouter Embedding。")
    client = OpenAI(
        base_url=settings.OPENROUTER_BASE_URL,
        api_key=settings.OPENROUTER_API_KEY,
        timeout=timeout,
    )
    if log_callback:
        log_callback(f"[embedding-openrouter] model={model}, batch={len(texts)}, dims={dimensions}")
    kwargs = {
        "model": model,
        "input": list(texts),
        "extra_headers": {"HTTP-Referer": "local-debug", "X-Title": "fastapi-llm-demo"},
    }
    if dimensions and dimensions > 0:
        kwargs["dimensions"] = dimensions
    response = client.embeddings.create(**kwargs)
    ordered = sorted(response.data, key=lambda item: item.index)
    return [_l2_normalize(item.embedding) for item in ordered]


def resolve_embedding_route(route: Optional[str] = None) -> str:
    """Prefer google/aistudio for embeddings; fall back to openrouter when needed."""
    preferred = normalize_gemini_route(route)
    if preferred in {GEMINI_ROUTE_GOOGLE, GEMINI_ROUTE_AISTUDIO}:
        try:
            return ensure_gemini_route_configured(preferred)
        except ValueError:
            pass
    if preferred == GEMINI_ROUTE_OPENROUTER:
        try:
            return ensure_gemini_route_configured(GEMINI_ROUTE_OPENROUTER)
        except ValueError:
            pass
    for candidate in (GEMINI_ROUTE_AISTUDIO, GEMINI_ROUTE_GOOGLE, GEMINI_ROUTE_OPENROUTER):
        try:
            return ensure_gemini_route_configured(candidate)
        except ValueError:
            continue
    raise ValueError("未配置可用的 Embedding 路线（需要 GOOGLE_API_KEY / Vertex / OPENROUTER_API_KEY）。")


def embed_texts(
    texts: Sequence[str],
    *,
    model: Optional[str] = None,
    route: Optional[str] = None,
    task_type: str = "SEMANTIC_SIMILARITY",
    dimensions: Optional[int] = None,
    batch_size: int = 64,
    timeout: float = DEFAULT_GEMINI_TIMEOUT_SECONDS,
    log_callback: GeminiLogCallback = None,
) -> List[List[float]]:
    """
    Embed texts with gemini-embedding-001 (or configured model).
    Returns L2-normalized vectors in the same order as inputs.
    """
    cleaned = [str(text or "").strip() or " " for text in texts]
    if not cleaned:
        return []

    dims = int(dimensions or DEFAULT_EMBEDDING_DIMENSIONS)
    normalized_route = resolve_embedding_route(route)
    raw_model = (model or DEFAULT_EMBEDDING_MODEL).strip() or DEFAULT_EMBEDDING_MODEL

    if log_callback:
        log_callback(
            f"[embedding] route={normalized_route}, model={raw_model}, "
            f"count={len(cleaned)}, dims={dims}, task_type={task_type}"
        )

    vectors: List[List[float]] = []
    for start in range(0, len(cleaned), max(1, batch_size)):
        batch = cleaned[start:start + max(1, batch_size)]
        if normalized_route == GEMINI_ROUTE_OPENROUTER:
            openrouter_model = raw_model if "/" in raw_model else f"google/{raw_model}"
            batch_vectors = _embed_openrouter_batch(
                texts=batch,
                model=openrouter_model,
                dimensions=dims,
                timeout=timeout,
                log_callback=log_callback,
            )
        else:
            google_model = normalize_google_model(raw_model)
            client_factory = (
                _get_vertex_client if normalized_route == GEMINI_ROUTE_GOOGLE else _get_aistudio_client
            )
            client = client_factory(timeout=timeout)
            batch_vectors = _embed_google_batch(
                client=client,
                model=google_model,
                texts=batch,
                task_type=task_type,
                dimensions=dims,
            )
        if len(batch_vectors) != len(batch):
            raise RuntimeError(
                f"embedding batch size mismatch: got {len(batch_vectors)} for {len(batch)} texts"
            )
        vectors.extend(batch_vectors)

    return vectors
