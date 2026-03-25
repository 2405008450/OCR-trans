"""配置管理"""

from dataclasses import dataclass
import os


@dataclass
class TranslatorConfig:
    """翻译系统配置"""
    glm_api_key: str             # 智谱 GLM API 密钥
    deepseek_api_key: str        # DeepSeek API 密钥
    
    @classmethod
    def from_env(cls) -> 'TranslatorConfig':
        """从环境变量加载配置"""
        return cls(
            glm_api_key=os.getenv('GLM_API_KEY', ''),
            deepseek_api_key=os.getenv('DEEPSEEK_API_KEY', '')
        )
