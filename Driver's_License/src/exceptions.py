"""异常类定义"""


class TranslatorError(Exception):
    """翻译系统基础异常"""
    pass


class OCRError(TranslatorError):
    """OCR 识别异常"""
    pass


class ImageLoadError(TranslatorError):
    """图片加载异常"""
    pass


class TranslationError(TranslatorError):
    """翻译异常"""
    pass


class DocumentGenerationError(TranslatorError):
    """文档生成异常"""
    pass


class TranslationPipelineError(TranslatorError):
    """翻译流程异常"""
    pass
