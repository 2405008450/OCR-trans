import os
from pathlib import Path

try:
    from dotenv import load_dotenv
except ImportError:
    load_dotenv = None

try:
    from pydantic_settings import BaseSettings
except ImportError:
    try:
        from pydantic import BaseSettings
    except ImportError:
        class BaseSettings:  # type: ignore[override]
            def __init__(self, **kwargs):
                for key, value in kwargs.items():
                    setattr(self, key, value)


_ROOT_DIR = Path(__file__).resolve().parents[2]
_ENV_FILE = _ROOT_DIR / ".env"

if load_dotenv and _ENV_FILE.exists():
    load_dotenv(_ENV_FILE, override=False)


class Settings(BaseSettings):
    GLM_API_KEY: str = os.getenv("GLM_API_KEY", "")
    DEEPSEEK_API_KEY: str = os.getenv("DEEPSEEK_API_KEY", "sk-f2a71209abd64087a69147ab6a0bb2ec")
    DEEPSEEK_BASE_URL: str = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com")

    GOOGLE_API_KEY: str = os.getenv("GOOGLE_API_KEY") or os.getenv("GEMINI_API_KEY", "")
    AI_STUDIO_API_KEY: str = (
        os.getenv("AI_STUDIO_API_KEY")
        or os.getenv("GOOGLE_API_KEY")
        or os.getenv("GEMINI_API_KEY", "")
    )
    GEMINI_DEFAULT_ROUTE: str = os.getenv("GEMINI_DEFAULT_ROUTE", "google")
    GEMINI_ENABLE_OPENROUTER_FALLBACK: str = os.getenv("GEMINI_ENABLE_OPENROUTER_FALLBACK", "False")
    AI_STUDIO_USE_FILES_API: str = os.getenv("AI_STUDIO_USE_FILES_API", "True")
    AI_STUDIO_FILES_API_DELETE_REMOTE: str = os.getenv("AI_STUDIO_FILES_API_DELETE_REMOTE", "True")

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
    def AI_STUDIO_USE_FILES_API_ENABLED(self) -> bool:
        return str(self.AI_STUDIO_USE_FILES_API).strip().lower() in {"1", "true", "yes", "on"}

    @property
    def AI_STUDIO_FILES_API_DELETE_REMOTE_ENABLED(self) -> bool:
        return str(self.AI_STUDIO_FILES_API_DELETE_REMOTE).strip().lower() in {"1", "true", "yes", "on"}


settings = Settings()
