# exceptions.py

class CheckError(Exception):
    """项目所有自定义异常的基类

    Attributes:
        message -- 错误描述
        details -- 可选的详细错误信息（如原始报错堆栈或 API 响应）
    """

    def __init__(self, message: str, details: str = None):
        super().__init__(message)
        self.message = message
        self.details = details

    def __str__(self):
        if self.details:
            return f"{self.message} | 详情: {self.details}"
        return self.message


# ---- 文件/解析相关 ----
class DocumentParseError(CheckError):
    """文档解析失败（格式不支持、文件损坏、提取内容为空等）"""
    pass


class RuleLoadError(CheckError):
    """规则文件加载失败"""
    pass


# ---- API/LLM 相关 ----
class APIProcessError(CheckError):
    """API 请求层面的错误（如网络超时、鉴权失败、余额不足）"""
    pass


class RuleGenerationError(CheckError):
    """AI 生成规则失败（返回内容无效）"""
    pass


class ComparisonError(CheckError):
    """对比过程出错（JSON 格式错误等）"""
    pass


# ---- 配置/参数相关 ----
class ConfigurationError(CheckError):
    """配置项错误"""
    pass