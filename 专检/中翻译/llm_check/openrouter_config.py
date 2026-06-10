import os
from pathlib import Path

from dotenv import load_dotenv
from openai import OpenAI


DEFAULT_OPENROUTER_BASE_URL = "https://openrouter.ai/api/v1"


def _load_project_env() -> None:
    """加载项目根目录 .env，兼容从模块目录直接运行脚本的情况。"""
    project_env = Path(__file__).resolve().parents[3] / ".env"
    if project_env.exists():
        load_dotenv(project_env, override=False)
    load_dotenv(override=False)


def resolve_openrouter_config() -> tuple[str | None, str]:
    _load_project_env()
    api_key = (
        os.getenv("OPENROUTER_API_KEY")
        or os.getenv("OPENAI_API_KEY")
        or os.getenv("API_KEY")
    )
    base_url = (
        os.getenv("OPENROUTER_BASE_URL")
        or os.getenv("OPENAI_BASE_URL")
        or os.getenv("BASE_URL")
        or DEFAULT_OPENROUTER_BASE_URL
    )

    if api_key:
        os.environ.setdefault("OPENROUTER_API_KEY", api_key)
        os.environ.setdefault("OPENAI_API_KEY", api_key)
        os.environ.setdefault("API_KEY", api_key)
    os.environ.setdefault("OPENROUTER_BASE_URL", base_url)
    os.environ.setdefault("OPENAI_BASE_URL", base_url)
    os.environ.setdefault("BASE_URL", base_url)
    return api_key, base_url


def create_openrouter_client() -> OpenAI:
    api_key, base_url = resolve_openrouter_config()
    if not api_key:
        raise RuntimeError("未配置 OPENROUTER_API_KEY，请按根目录 env.example 创建或更新 .env。")
    return OpenAI(api_key=api_key, base_url=base_url)


def get_model_name(default_model: str) -> str:
    return os.getenv("ZHONGFANYI_MODEL_NAME") or default_model
