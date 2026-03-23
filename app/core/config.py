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
    DEEPSEEK_API_KEY: str = os.getenv("DEEPSEEK_API_KEY", "sk-f2a71209abd64087a69147ab6a0bb2ec")
    DEEPSEEK_BASE_URL: str = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com")

    GOOGLE_API_KEY: str = os.getenv("GOOGLE_API_KEY") or os.getenv("GEMINI_API_KEY", "")
    GEMINI_DEFAULT_ROUTE: str = os.getenv("GEMINI_DEFAULT_ROUTE", "google")

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

    @property
    def DEBUG_ENABLED(self) -> bool:
        return str(self.DEBUG).strip().lower() in {"1", "true", "yes", "on", "debug"}


settings = Settings()
