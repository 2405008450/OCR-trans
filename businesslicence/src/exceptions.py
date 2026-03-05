"""Custom exceptions for image translation system.

This module defines the exception hierarchy for handling various
error conditions in the image translation pipeline.
"""


class ImageTranslationError(Exception):
    """Base exception class for image translation errors."""
    pass


class OCRError(ImageTranslationError):
    """Exception raised when OCR processing fails."""
    pass


class TranslationError(ImageTranslationError):
    """Exception raised when translation service fails."""
    pass


class ImageLoadError(ImageTranslationError):
    """Exception raised when image loading fails."""
    pass


class ImageSaveError(ImageTranslationError):
    """Exception raised when image saving fails."""
    pass


class ConfigError(ImageTranslationError):
    """Exception raised when configuration is invalid or missing.
    
    配置错误异常
    
    当配置无效或缺失时抛出的异常。
    """
    pass


class ConfigValidationError(ConfigError):
    """Exception raised when configuration validation fails.
    
    配置验证错误异常
    
    当配置验证失败时抛出的异常，包含详细的错误位置和原因信息。
    
    Attributes:
        field: 出错的配置字段路径 | Path to the configuration field that failed validation
        message: 错误描述信息 | Error description message
        expected: 期望的值或类型 | Expected value or type
        actual: 实际的值或类型 | Actual value or type
    """
    
    def __init__(
        self,
        field: str,
        message: str,
        expected: str = None,
        actual: str = None
    ):
        """初始化配置验证错误
        
        Initialize configuration validation error.
        
        Args:
            field: 出错的配置字段路径 | Path to the configuration field
            message: 错误描述信息 | Error description message
            expected: 期望的值或类型（可选）| Expected value or type (optional)
            actual: 实际的值或类型（可选）| Actual value or type (optional)
        """
        self.field = field
        self.message = message
        self.expected = expected
        self.actual = actual
        
        # 构建详细的错误信息 | Build detailed error message
        error_parts = [f"配置验证失败 | Configuration validation failed: {field}"]
        error_parts.append(f"  错误 | Error: {message}")
        
        if expected is not None:
            error_parts.append(f"  期望 | Expected: {expected}")
        
        if actual is not None:
            error_parts.append(f"  实际 | Actual: {actual}")
        
        full_message = "\n".join(error_parts)
        super().__init__(full_message)
    
    def __str__(self) -> str:
        """返回格式化的错误信息 | Return formatted error message."""
        return super().__str__()


class ConfigMigrationError(ConfigError):
    """Exception raised when configuration migration fails.
    
    配置迁移错误异常
    
    当配置迁移失败时抛出的异常。
    
    Attributes:
        source_format: 源配置格式 | Source configuration format
        target_format: 目标配置格式 | Target configuration format
        reason: 失败原因 | Reason for failure
    """
    
    def __init__(self, source_format: str, target_format: str, reason: str):
        """初始化配置迁移错误
        
        Initialize configuration migration error.
        
        Args:
            source_format: 源配置格式 | Source configuration format
            target_format: 目标配置格式 | Target configuration format
            reason: 失败原因 | Reason for failure
        """
        self.source_format = source_format
        self.target_format = target_format
        self.reason = reason
        
        message = (
            f"配置迁移失败 | Configuration migration failed: "
            f"{source_format} -> {target_format}\n"
            f"  原因 | Reason: {reason}"
        )
        super().__init__(message)


class PipelineError(ImageTranslationError):
    """Exception raised when the translation pipeline fails."""
    pass
