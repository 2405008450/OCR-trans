"""Paragraph merge module for field-aware text region merging.

This module provides components for intelligent paragraph merging,
particularly for long text fields like "经营范围" (Business Scope).

本模块提供字段感知的段落合并功能,特别针对"经营范围"等长文本字段。
"""

from src.paragraph_merge.spatial_analyzer import SpatialAnalyzer
from src.paragraph_merge.field_content_detector import FieldContentDetector
from src.paragraph_merge.region_merger import RegionMerger
from src.paragraph_merge.merge_strategy_selector import MergeStrategySelector, MergeStrategy
from src.paragraph_merge.field_aware_paragraph_merger import FieldAwareParagraphMerger

__all__ = [
    'SpatialAnalyzer', 
    'FieldContentDetector', 
    'RegionMerger',
    'MergeStrategySelector',
    'MergeStrategy',
    'FieldAwareParagraphMerger'
]
