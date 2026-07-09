# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import os
import sys
import time
from pathlib import Path
from typing import Optional

import httpx
from dotenv import load_dotenv
from openai import OpenAI


PROXY_ENV_KEYS = (
    "HTTP_PROXY",
    "HTTPS_PROXY",
    "ALL_PROXY",
    "NO_PROXY",
    "http_proxy",
    "https_proxy",
    "all_proxy",
    "no_proxy",
)


def _redact(value: Optional[str]) -> str:
    text = str(value or "")
    if not text:
        return "<empty>"
    if len(text) <= 12:
        return "***"
    return f"{text[:8]}...{text[-4:]}"


def _print_env(base_url: str, api_key: str) -> None:
    print("=== runtime ===")
    print(f"python={sys.version.split()[0]}")
    print(f"executable={sys.executable}")
    print(f"cwd={Path.cwd()}")
    print()
    print("=== openrouter ===")
    print(f"OPENROUTER_BASE_URL={base_url}")
    print(f"OPENROUTER_API_KEY={_redact(api_key)}")
    print()
    print("=== proxy env seen by this process ===")
    for key in PROXY_ENV_KEYS:
        print(f"{key}={os.getenv(key) or '<empty>'}")
    print()


def _timed(label: str, func) -> bool:
    started = time.perf_counter()
    try:
        func()
    except Exception as exc:
        elapsed = time.perf_counter() - started
        print(f"[FAIL] {label} after {elapsed:.2f}s")
        print(f"       {type(exc).__name__}: {exc}")
        return False
    elapsed = time.perf_counter() - started
    print(f"[OK] {label} in {elapsed:.2f}s")
    return True


def _test_httpx_models(*, base_url: str, api_key: str, timeout: float, proxy: Optional[str]) -> None:
    headers = {"Authorization": f"Bearer {api_key}"} if api_key else {}
    url = f"{base_url.rstrip('/')}/models"
    if proxy:
        client = httpx.Client(proxy=proxy, timeout=timeout, trust_env=False)
    else:
        client = httpx.Client(timeout=timeout, trust_env=True)
    with client:
        response = client.get(url, headers=headers)
        print(f"       GET {url} -> {response.status_code}")
        print(f"       body={response.text[:240].replace(chr(10), ' ')}")
        response.raise_for_status()


def _test_openai_chat(
    *,
    base_url: str,
    api_key: str,
    model: str,
    timeout: float,
    proxy: Optional[str],
) -> None:
    http_client = None
    if proxy:
        http_client = httpx.Client(proxy=proxy, timeout=timeout, trust_env=False)
    try:
        client = OpenAI(
            base_url=base_url,
            api_key=api_key,
            timeout=timeout,
            http_client=http_client,
            default_headers={
                "HTTP-Referer": "local-proxy-test",
                "X-Title": "fastapi-llm-demo proxy test",
            },
        )
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": "You are a connectivity test. Reply with exactly: pong"},
                {"role": "user", "content": "ping"},
            ],
            temperature=0,
            max_tokens=16,
            stream=False,
        )
        text = (response.choices[0].message.content or "").strip()
        print(f"       model={model}")
        print(f"       reply={text[:120]}")
    finally:
        if http_client is not None:
            http_client.close()


def main() -> int:
    parser = argparse.ArgumentParser(description="Test OpenRouter connectivity and proxy behavior.")
    parser.add_argument("--env-file", default=".env", help="Path to .env file. Default: .env")
    parser.add_argument("--no-dotenv-override", action="store_true", help="Mimic app startup: do not override existing OS env vars.")
    parser.add_argument("--model", default=os.getenv("OPENROUTER_TEST_MODEL", "openai/gpt-4o-mini"))
    parser.add_argument("--timeout", type=float, default=60)
    parser.add_argument("--explicit-proxy", default="", help="Force a proxy for an extra explicit-proxy test, e.g. http://127.0.0.1:7897")
    args = parser.parse_args()

    env_path = Path(args.env_file)
    if env_path.exists():
        load_dotenv(env_path, override=not args.no_dotenv_override)
        print(f"loaded_env={env_path.resolve()} override={not args.no_dotenv_override}")
    else:
        print(f"loaded_env=<missing: {env_path}>")

    base_url = os.getenv("OPENROUTER_BASE_URL", "https://openrouter.ai/api/v1")
    api_key = os.getenv("OPENROUTER_API_KEY", "")
    if not api_key:
        print("ERROR: OPENROUTER_API_KEY is empty")
        return 2

    env_proxy = os.getenv("HTTPS_PROXY") or os.getenv("HTTP_PROXY") or os.getenv("ALL_PROXY")
    explicit_proxy = args.explicit_proxy.strip()
    _print_env(base_url, api_key)

    ok = True
    ok &= _timed(
        "httpx /models using environment proxy",
        lambda: _test_httpx_models(base_url=base_url, api_key=api_key, timeout=args.timeout, proxy=None),
    )
    ok &= _timed(
        "OpenAI SDK chat using environment proxy",
        lambda: _test_openai_chat(base_url=base_url, api_key=api_key, model=args.model, timeout=args.timeout, proxy=None),
    )

    proxy_to_force = explicit_proxy or env_proxy or ""
    if proxy_to_force:
        print()
        print(f"=== explicit proxy check: {proxy_to_force} ===")
        ok &= _timed(
            "httpx /models using explicit proxy",
            lambda: _test_httpx_models(base_url=base_url, api_key=api_key, timeout=args.timeout, proxy=proxy_to_force),
        )
        ok &= _timed(
            "OpenAI SDK chat using explicit proxy",
            lambda: _test_openai_chat(base_url=base_url, api_key=api_key, model=args.model, timeout=args.timeout, proxy=proxy_to_force),
        )

    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
