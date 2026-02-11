import os
from typing import Optional

try:
    from pydantic_settings import BaseSettings
except ImportError:
    try:
        # 兼容旧版本的pydantic (v1.x)
        from pydantic import BaseSettings
    except ImportError:
        # 如果都导入失败，使用普通类
        class BaseSettings:
            def __init__(self, **kwargs):
                for key, value in kwargs.items():
                    setattr(self, key, value)

class Settings(BaseSettings):
    # DeepSeek API配置
    DEEPSEEK_API_KEY: str = os.getenv("DEEPSEEK_API_KEY", "sk-f2a71209abd64087a69147ab6a0bb2ec")
    DEEPSEEK_BASE_URL: str = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com")
    
    # 服务器配置
    HOST: str = os.getenv("HOST", "0.0.0.0")  # 0.0.0.0 允许所有网络接口访问
    PORT: int = int(os.getenv("PORT", "8001"))  # 端口号（默认8001，避免与8000冲突）
    DEBUG: bool = os.getenv("DEBUG", "False").lower() == "true"  # 调试模式
    
    # 文件路径配置
    UPLOAD_DIR: str = os.getenv("UPLOAD_DIR", "uploads")
    OUTPUT_DIR: str = os.getenv("OUTPUT_DIR", "outputs")
    TEMP_IMAGES_DIR: str = os.getenv("TEMP_IMAGES_DIR", "temp_images")
    
    # 图片处理配置
    TARGET_IMAGE_WIDTH: int = int(os.getenv("TARGET_IMAGE_WIDTH", "1080"))
    
    # CORS配置（云服务器可能需要）
    ALLOWED_ORIGINS: str = os.getenv("ALLOWED_ORIGINS", "*")  # 允许的源，用逗号分隔
    
    class Config:
        env_file = ".env"
        case_sensitive = True

settings = Settings()

