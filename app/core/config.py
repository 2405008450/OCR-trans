import json
import os
from pathlib import Path

try:
    from dotenv import dotenv_values, load_dotenv
except ImportError:
    load_dotenv = None
    dotenv_values = None

try:
    from pydantic_settings import BaseSettings, SettingsConfigDict
    _USE_SETTINGS_CONFIG_DICT = True
except ImportError:
    try:
        from pydantic import BaseSettings
        _USE_SETTINGS_CONFIG_DICT = False
    except ImportError:
        class BaseSettings:  # type: ignore[override]
            def __init__(self, **kwargs):
                for key, value in kwargs.items():
                    setattr(self, key, value)

        _USE_SETTINGS_CONFIG_DICT = False


_ROOT_DIR = Path(__file__).resolve().parents[2]
_ENV_FILE = _ROOT_DIR / ".env"
_PROXY_ENV_KEYS = {
    "HTTP_PROXY",
    "HTTPS_PROXY",
    "ALL_PROXY",
    "NO_PROXY",
    "http_proxy",
    "https_proxy",
    "all_proxy",
    "no_proxy",
}


def _apply_dotenv_proxy_overrides(env_file: Path) -> None:
    if not dotenv_values:
        return
    values = dotenv_values(env_file)
    for key, value in values.items():
        if key not in _PROXY_ENV_KEYS:
            continue
        if value is None:
            continue
        text = str(value).strip()
        if text:
            os.environ[key] = text
        else:
            os.environ.pop(key, None)

if load_dotenv and _ENV_FILE.exists():
    load_dotenv(_ENV_FILE, override=False)
    _apply_dotenv_proxy_overrides(_ENV_FILE)


class Settings(BaseSettings):
    GLM_API_KEY: str = os.getenv("GLM_API_KEY", "")
    DEEPSEEK_API_KEY: str = os.getenv("DEEPSEEK_API_KEY", "")
    DEEPSEEK_BASE_URL: str = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com")

    GOOGLE_API_KEY: str = os.getenv("GOOGLE_API_KEY") or os.getenv("GEMINI_API_KEY", "")
    GEMINI_DEFAULT_ROUTE: str = os.getenv("GEMINI_DEFAULT_ROUTE", "google")
    GEMINI_ENABLE_OPENROUTER_FALLBACK: str = os.getenv("GEMINI_ENABLE_OPENROUTER_FALLBACK", "False")
    GEMINI_TIMEOUT_SECONDS: float = float(os.getenv("GEMINI_TIMEOUT_SECONDS", "120"))

    VERTEX_PROJECT_ID: str = os.getenv("VERTEX_PROJECT_ID", "gen-lang-client-0128671098")
    VERTEX_LOCATION: str = os.getenv("VERTEX_LOCATION", "global")

    OPENROUTER_API_KEY: str = os.getenv("OPENROUTER_API_KEY") or os.getenv("OPENAI_API_KEY", "")
    OPENROUTER_BASE_URL: str = os.getenv("OPENROUTER_BASE_URL") or os.getenv("OPENAI_BASE_URL", "https://openrouter.ai/api/v1")

    HOST: str = os.getenv("HOST", "0.0.0.0")
    PORT: int = int(os.getenv("PORT", "8001"))
    DEBUG: str = os.getenv("DEBUG", "False")

    UPLOAD_DIR: str = os.getenv("UPLOAD_DIR", "uploads")
    OUTPUT_DIR: str = os.getenv("OUTPUT_DIR", "outputs")
    TEMP_IMAGES_DIR: str = os.getenv("TEMP_IMAGES_DIR", "temp_images")

    TARGET_IMAGE_WIDTH: int = int(os.getenv("TARGET_IMAGE_WIDTH", "1080"))
    ALLOWED_ORIGINS: str = os.getenv("ALLOWED_ORIGINS", "*")
    TASK_QUEUE_MAX_CONCURRENT_TASKS: int = int(os.getenv("TASK_QUEUE_MAX_CONCURRENT_TASKS", "2"))
    TASK_QUEUE_EXECUTOR_MAX_WORKERS: int = int(os.getenv("TASK_QUEUE_EXECUTOR_MAX_WORKERS", "4"))
    TASK_QUEUE_POLL_INTERVAL_SECONDS: float = float(os.getenv("TASK_QUEUE_POLL_INTERVAL_SECONDS", "0.5"))
    TASK_QUEUE_CANDIDATE_BATCH_SIZE: int = int(os.getenv("TASK_QUEUE_CANDIDATE_BATCH_SIZE", "20"))
    TASK_QUEUE_TYPE_LIMITS_JSON: str = os.getenv("TASK_QUEUE_TYPE_LIMITS_JSON", "")
    WORD_COUNT_ALLOWED_ROOTS_JSON: str = os.getenv("WORD_COUNT_ALLOWED_ROOTS_JSON", "")
    WORD_COUNT_UNC_MOUNT_MAP_JSON: str = os.getenv("WORD_COUNT_UNC_MOUNT_MAP_JSON", "")
    WORD_COUNT_UNC_AUTO_MOUNT_ROOTS_JSON: str = os.getenv(
        "WORD_COUNT_UNC_AUTO_MOUNT_ROOTS_JSON",
        '["/mnt","/media","/shares","/srv","/data","/app/data","/app/mnt","/app"]',
    )
    WORD_COUNT_ALLOW_LOCAL_PATHS: str = os.getenv("WORD_COUNT_ALLOW_LOCAL_PATHS", "False")
    WORD_COUNT_MAX_FILES: int = int(os.getenv("WORD_COUNT_MAX_FILES", "5000"))
    WORD_COUNT_MAX_FILE_MB: int = int(os.getenv("WORD_COUNT_MAX_FILE_MB", "200"))
    WORD_COUNT_FOLLOW_SYMLINKS: str = os.getenv("WORD_COUNT_FOLLOW_SYMLINKS", "False")

    if _USE_SETTINGS_CONFIG_DICT:
        model_config = SettingsConfigDict(
            env_file=str(_ENV_FILE) if _ENV_FILE.exists() else ".env",
            case_sensitive=True,
            extra="ignore",
        )
    else:
        class Config:
            env_file = str(_ENV_FILE) if _ENV_FILE.exists() else ".env"
            case_sensitive = True
            extra = "ignore"

    @property
    def DEBUG_ENABLED(self) -> bool:
        return str(self.DEBUG).strip().lower() in {"1", "true", "yes", "on", "debug"}

    @property
    def GEMINI_ENABLE_OPENROUTER_FALLBACK_ENABLED(self) -> bool:
        return str(self.GEMINI_ENABLE_OPENROUTER_FALLBACK).strip().lower() in {"1", "true", "yes", "on"}

    @property
    def TASK_QUEUE_TYPE_LIMITS(self) -> dict[str, int]:
        raw = (self.TASK_QUEUE_TYPE_LIMITS_JSON or "").strip()
        if not raw:
            return {}
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError:
            return {}
        if not isinstance(parsed, dict):
            return {}
        limits: dict[str, int] = {}
        for key, value in parsed.items():
            try:
                limit = int(value)
            except (TypeError, ValueError):
                continue
            if limit > 0:
                limits[str(key)] = limit
        return limits

    @property
    def WORD_COUNT_ALLOWED_ROOTS(self) -> list[str]:
        raw = (self.WORD_COUNT_ALLOWED_ROOTS_JSON or "").strip()
        if not raw:
            return [str(_ROOT_DIR)]
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError:
            return [str(_ROOT_DIR)]
        if isinstance(parsed, str):
            parsed = [parsed]
        if not isinstance(parsed, list):
            return [str(_ROOT_DIR)]
        roots = [str(item).strip() for item in parsed if str(item).strip()]
        return roots or [str(_ROOT_DIR)]

    @property
    def WORD_COUNT_UNC_MOUNT_MAP(self) -> dict[str, str]:
        raw = (self.WORD_COUNT_UNC_MOUNT_MAP_JSON or "").strip()
        if not raw:
            return {}
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError:
            return {}
        mappings: dict[str, str] = {}
        if isinstance(parsed, dict):
            items = parsed.items()
        elif isinstance(parsed, list):
            items = (
                (item.get("unc") or item.get("source") or item.get("from"), item.get("mount") or item.get("target") or item.get("to"))
                for item in parsed
                if isinstance(item, dict)
            )
        else:
            return {}
        for unc_path, mount_path in items:
            unc_text = str(unc_path or "").strip()
            mount_text = str(mount_path or "").strip()
            if unc_text and mount_text:
                mappings[unc_text] = mount_text
        return mappings

    @property
    def WORD_COUNT_UNC_AUTO_MOUNT_ROOTS(self) -> list[str]:
        raw = (self.WORD_COUNT_UNC_AUTO_MOUNT_ROOTS_JSON or "").strip()
        if not raw:
            return []
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError:
            return []
        if isinstance(parsed, str):
            parsed = [parsed]
        if not isinstance(parsed, list):
            return []
        return [str(item).strip() for item in parsed if str(item).strip()]

    @property
    def WORD_COUNT_FOLLOW_SYMLINKS_ENABLED(self) -> bool:
        return str(self.WORD_COUNT_FOLLOW_SYMLINKS).strip().lower() in {"1", "true", "yes", "on"}

    @property
    def WORD_COUNT_ALLOW_LOCAL_PATHS_ENABLED(self) -> bool:
        return str(self.WORD_COUNT_ALLOW_LOCAL_PATHS).strip().lower() in {"1", "true", "yes", "on"}


settings = Settings()
