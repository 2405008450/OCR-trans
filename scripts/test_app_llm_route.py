# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import os
import sys
import time
from pathlib import Path
from typing import Optional


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


def _print_runtime(settings, route: str, model: str) -> None:
    print("=== runtime ===")
    print(f"python={sys.version.split()[0]}")
    print(f"executable={sys.executable}")
    print(f"cwd={Path.cwd()}")
    print()
    print("=== app settings ===")
    print(f"GEMINI_DEFAULT_ROUTE={settings.GEMINI_DEFAULT_ROUTE}")
    print(f"GEMINI_ENABLE_OPENROUTER_FALLBACK={settings.GEMINI_ENABLE_OPENROUTER_FALLBACK}")
    print(f"OPENROUTER_BASE_URL={settings.OPENROUTER_BASE_URL}")
    print(f"OPENROUTER_API_KEY={_redact(settings.OPENROUTER_API_KEY)}")
    print(f"test_route={route}")
    print(f"test_model={model}")
    print()
    print("=== proxy env seen by app code ===")
    for key in PROXY_ENV_KEYS:
        print(f"{key}={os.getenv(key) or '<empty>'}")
    print()


def _print_env_file_proxy_warnings(env_file: Path) -> None:
    if not env_file.exists():
        return
    counts: dict[str, int] = {}
    for raw_line in env_file.read_text(encoding="utf-8", errors="ignore").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key = line.split("=", 1)[0].strip()
        if key in PROXY_ENV_KEYS:
            counts[key] = counts.get(key, 0) + 1
    duplicated = [key for key, count in counts.items() if count > 1]
    if duplicated:
        print("=== .env warning ===")
        print(f"duplicate proxy keys: {', '.join(duplicated)}")
        print("python-dotenv normally keeps the last value, so remove old Docker proxy lines if they are still present.")
        print()


def main() -> int:
    parser = argparse.ArgumentParser(description="Test the app's real LLM route through app.service.gemini_service.generate_text.")
    parser.add_argument("--route", default="openrouter", help="Route to test. Default: openrouter")
    parser.add_argument("--model", default=os.getenv("OPENROUTER_TEST_MODEL", "openai/gpt-4o-mini"))
    parser.add_argument("--timeout", type=float, default=60)
    args = parser.parse_args()

    from app.core.config import _ENV_FILE, settings
    from app.service.gemini_service import generate_text

    _print_env_file_proxy_warnings(_ENV_FILE)
    _print_runtime(settings, args.route, args.model)
    if args.route == "openrouter" and not settings.OPENROUTER_API_KEY:
        print("ERROR: OPENROUTER_API_KEY is empty in app settings")
        return 2

    def _log(message: str) -> None:
        print(f"       {message}")

    started = time.perf_counter()
    try:
        text = generate_text(
            system_prompt="You are a connectivity test. Reply with exactly: pong",
            user_prompt="ping",
            model=args.model,
            route=args.route,
            temperature=0,
            max_output_tokens=16,
            timeout=args.timeout,
            log_callback=_log,
        )
    except Exception as exc:
        elapsed = time.perf_counter() - started
        print(f"[FAIL] app generate_text after {elapsed:.2f}s")
        print(f"       {type(exc).__name__}: {exc}")
        return 1

    elapsed = time.perf_counter() - started
    print(f"[OK] app generate_text in {elapsed:.2f}s")
    print(f"reply={text[:200]}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
