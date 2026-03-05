"""Cache module for image translation system.

This module provides caching functionality for OCR results
to improve performance by avoiding redundant processing.
"""

from src.cache.ocr_cache import OCRCache, ImageHasher

__all__ = ['OCRCache', 'ImageHasher']
