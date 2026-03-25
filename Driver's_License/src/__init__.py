"""驾驶证翻译文档生成系统"""

__version__ = "1.0.0"

from .models import TextBlock, LicenseField, ExtractedImage, LicenseData
from .exceptions import (
    TranslatorError,
    OCRError,
    ImageLoadError,
    TranslationError,
    DocumentGenerationError,
    TranslationPipelineError
)
from .config import TranslatorConfig
from .ocr_service import OCRService
from .translator_pipeline import TranslatorPipeline

__all__ = [
    'TextBlock',
    'LicenseField',
    'ExtractedImage',
    'LicenseData',
    'TranslatorError',
    'OCRError',
    'ImageLoadError',
    'TranslationError',
    'DocumentGenerationError',
    'TranslationPipelineError',
    'TranslatorConfig',
    'OCRService',
    'TranslatorPipeline',
]
