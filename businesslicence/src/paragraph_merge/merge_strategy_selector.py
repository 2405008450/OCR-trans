"""Merge strategy selector for choosing appropriate merge strategies.

This module provides strategy selection capabilities for determining
the appropriate merge strategy based on field type and content.

本模块提供合并策略选择功能，根据字段类型和内容确定合适的合并策略。
"""

import logging
from enum import Enum
from typing import List, Optional
from src.models.data_models import TextRegion, MergeConfig
from src.config.config_manager import ConfigManager


logger = logging.getLogger(__name__)


class MergeStrategy(Enum):
    """合并策略枚举
    
    Merge strategy enumeration.
    
    Defines the available merge strategies for text regions.
    
    Attributes:
        STANDARD: 标准段落合并 | Standard paragraph merge
        LONG_TEXT_FIELD: 长文本字段合并 | Long text field merge
        NO_MERGE: 不合并 | No merge
    """
    STANDARD = "standard"
    LONG_TEXT_FIELD = "long_text_field"
    NO_MERGE = "no_merge"


class MergeStrategySelector:
    """合并策略选择器
    
    Merge strategy selector.
    
    Selects the appropriate merge strategy based on field type,
    configuration settings, and region characteristics.
    
    Attributes:
        config: Configuration manager containing merge parameters
        merge_config: Merge configuration with field definitions
    """
    
    def __init__(self, config: ConfigManager):
        """初始化策略选择器
        
        Initialize the merge strategy selector.
        
        Args:
            config: Configuration manager instance
        """
        self.config = config
        self.merge_config = self._load_merge_config()
    
    def _load_merge_config(self) -> MergeConfig:
        """从配置管理器加载合并配置
        
        Load merge configuration from config manager.
        
        Returns:
            MergeConfig instance with loaded parameters
        """
        # Get configuration values with defaults
        enabled = self.config.get('ocr.paragraph_merge.enabled', True)
        x_tolerance = self.config.get('ocr.paragraph_merge.x_tolerance', 10)
        y_gap_max = self.config.get('ocr.paragraph_merge.y_gap_max', 5)
        font_size_diff_threshold = self.config.get(
            'ocr.paragraph_merge.font_size_diff_threshold', 0.2
        )
        field_aware_merge_enabled = self.config.get(
            'ocr.field_aware_merge.enabled', True
        )
        known_field_labels = self.config.get(
            'ocr.field_aware_merge.known_field_labels',
            ["经营范围", "住所", "法定代表人", "注册资本", "成立日期", "营业期限"]
        )
        long_text_fields = self.config.get(
            'ocr.field_aware_merge.long_text_fields',
            ["经营范围", "住所"]
        )
        long_text_field_y_gap_max = self.config.get(
            'ocr.field_aware_merge.long_text_field_y_gap_max', 10
        )
        log_merge_decisions = self.config.get(
            'ocr.paragraph_merge.log_merge_decisions', True
        )
        log_level = self.config.get('ocr.paragraph_merge.log_level', 'INFO')
        
        return MergeConfig(
            enabled=enabled,
            x_tolerance=x_tolerance,
            y_gap_max=y_gap_max,
            font_size_diff_threshold=font_size_diff_threshold,
            field_aware_merge_enabled=field_aware_merge_enabled,
            known_field_labels=known_field_labels,
            long_text_fields=long_text_fields,
            long_text_field_y_gap_max=long_text_field_y_gap_max,
            log_merge_decisions=log_merge_decisions,
            log_level=log_level
        )
    
    def select_strategy(
        self,
        regions: List[TextRegion],
        field_label: Optional[str] = None
    ) -> MergeStrategy:
        """选择合并策略
        
        Select the appropriate merge strategy.
        
        The strategy is selected based on:
        1. Whether paragraph merging is enabled
        2. Whether field-aware merging is enabled
        3. The field label (if provided)
        4. The number of regions to merge
        
        Strategy selection logic:
        - If merging is disabled: NO_MERGE
        - If only one region: NO_MERGE
        - If field_label is provided and is a long text field: LONG_TEXT_FIELD
        - Otherwise: STANDARD
        
        Args:
            regions: List of text regions to merge
            field_label: Optional field label name
            
        Returns:
            Selected merge strategy
        """
        # Check if paragraph merging is enabled
        if not self.merge_config.enabled:
            if self.merge_config.log_merge_decisions:
                logger.debug("Paragraph merge disabled, selecting NO_MERGE strategy")
            return MergeStrategy.NO_MERGE
        
        # Check if we have regions to merge
        if not regions or len(regions) <= 1:
            if self.merge_config.log_merge_decisions:
                logger.debug(
                    f"Only {len(regions) if regions else 0} region(s), "
                    "selecting NO_MERGE strategy"
                )
            return MergeStrategy.NO_MERGE
        
        # Check if field-aware merging is enabled and we have a field label
        if (self.merge_config.field_aware_merge_enabled and 
            field_label and 
            field_label in self.merge_config.long_text_fields):
            if self.merge_config.log_merge_decisions:
                logger.info(
                    f"Field '{field_label}' is a long text field, "
                    "selecting LONG_TEXT_FIELD strategy"
                )
            return MergeStrategy.LONG_TEXT_FIELD
        
        # Default to standard merge strategy
        if self.merge_config.log_merge_decisions:
            if field_label:
                logger.debug(
                    f"Field '{field_label}' is not a long text field, "
                    "selecting STANDARD strategy"
                )
            else:
                logger.debug("No field label provided, selecting STANDARD strategy")
        
        return MergeStrategy.STANDARD
