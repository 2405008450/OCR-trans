"""Field-aware paragraph merger for intelligent text region merging.

This module provides the main paragraph merging coordinator that uses
field-aware strategies to merge text regions intelligently based on
their semantic meaning and spatial relationships.

本模块提供主要的段落合并协调器，使用字段感知策略根据文本区域的
语义和空间关系智能地合并文本区域。
"""

import logging
from typing import List, Optional
from src.models.data_models import TextRegion, MergeConfig
from src.config.config_manager import ConfigManager
from src.paragraph_merge.spatial_analyzer import SpatialAnalyzer
from src.paragraph_merge.field_content_detector import FieldContentDetector
from src.paragraph_merge.merge_strategy_selector import MergeStrategySelector, MergeStrategy
from src.paragraph_merge.region_merger import RegionMerger


logger = logging.getLogger(__name__)


class FieldAwareParagraphMerger:
    """字段感知的段落合并器
    
    Field-aware paragraph merger.
    
    Coordinates the entire paragraph merging process using field-aware
    strategies. This is the main entry point for paragraph merging.
    
    Attributes:
        config: Configuration manager containing merge parameters
        merge_config: Merge configuration with thresholds and tolerances
        spatial_analyzer: Analyzer for spatial relationships between regions
        field_detector: Detector for field labels and content blocks
        strategy_selector: Selector for choosing merge strategies
        region_merger: Merger for combining text regions
    """
    
    def __init__(self, config: ConfigManager):
        """初始化合并器
        
        Initialize the field-aware paragraph merger.
        
        Args:
            config: Configuration manager instance
        """
        self.config = config
        self.merge_config = self._load_merge_config()
        self.spatial_analyzer = SpatialAnalyzer(config)
        self.field_detector = FieldContentDetector(config)
        self.strategy_selector = MergeStrategySelector(config)
        self.region_merger = RegionMerger(config)
    
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
    
    def merge_regions(self, regions: List[TextRegion]) -> List[TextRegion]:
        """合并文本区域
        
        Merge text regions using field-aware strategies.
        
        This is the main entry point for paragraph merging. The method:
        1. Logs the initial region count
        2. Identifies field labels and their content blocks
        3. Selects appropriate merge strategies for each field
        4. Merges regions according to the selected strategies
        5. Handles errors and falls back to original regions if needed
        6. Logs the final region count and merge decisions
        
        Args:
            regions: Original text region list
            
        Returns:
            Merged text region list
        """
        # Log initial region count
        if self.merge_config.log_merge_decisions:
            logger.info(f"Starting paragraph merge with {len(regions)} regions")
        
        # Check if merging is enabled
        if not self.merge_config.enabled:
            if self.merge_config.log_merge_decisions:
                logger.info("Paragraph merge disabled, returning original regions")
            return regions
        
        # Check if we have regions to merge
        if not regions or len(regions) == 0:
            if self.merge_config.log_merge_decisions:
                logger.debug("No regions to merge")
            return regions
        
        try:
            # Identify field labels and content blocks
            merged_regions = []
            processed_indices = set()
            
            for i, region in enumerate(regions):
                # Skip if already processed
                if i in processed_indices:
                    continue
                
                # Check if this region is a field label
                field_label = self.field_detector.detect_field_label(region)
                
                if field_label:
                    # Mark this region as a field label
                    region.is_field_label = True
                    region.field_label_name = field_label
                    
                    # Find content blocks for this field
                    content_blocks = self.field_detector.find_field_content_blocks(
                        region, regions
                    )
                    
                    if self.merge_config.log_merge_decisions:
                        logger.info(
                            f"Found field label '{field_label}' with "
                            f"{len(content_blocks)} content block(s)"
                        )
                    
                    # Mark content blocks as belonging to this field
                    for content_block in content_blocks:
                        content_block.belongs_to_field = field_label
                    
                    # Select merge strategy for this field
                    strategy = self.strategy_selector.select_strategy(
                        content_blocks, field_label
                    )
                    
                    if self.merge_config.log_merge_decisions:
                        logger.debug(
                            f"Selected strategy '{strategy.value}' for field '{field_label}'"
                        )
                    
                    # Apply merge strategy
                    if strategy == MergeStrategy.LONG_TEXT_FIELD and content_blocks:
                        # Merge content blocks for long text field
                        merged_content = self.region_merger.merge_long_text_field(
                            region, content_blocks
                        )
                        
                        # Add field label and merged content as separate regions
                        merged_regions.append(region)
                        merged_regions.append(merged_content)
                        
                        # Mark content blocks as processed
                        for content_block in content_blocks:
                            content_idx = regions.index(content_block)
                            processed_indices.add(content_idx)
                    
                    elif strategy == MergeStrategy.STANDARD and content_blocks:
                        # Merge content blocks using standard strategy
                        merged_content = self.region_merger.merge_standard(content_blocks)
                        
                        # Add field label and merged content as separate regions
                        merged_regions.append(region)
                        merged_regions.append(merged_content)
                        
                        # Mark content blocks as processed
                        for content_block in content_blocks:
                            content_idx = regions.index(content_block)
                            processed_indices.add(content_idx)
                    
                    else:
                        # NO_MERGE or no content blocks - keep field label as is
                        merged_regions.append(region)
                        
                        # Add content blocks individually, but mark them as field content
                        for content_block in content_blocks:
                            # Mark as field content to prevent re-pairing by translation_pipeline
                            content_block.is_field_content = True
                            merged_regions.append(content_block)
                            content_idx = regions.index(content_block)
                            processed_indices.add(content_idx)
                    
                    # Mark field label as processed
                    processed_indices.add(i)
                
                else:
                    # Not a field label - check if it should be merged with adjacent regions
                    # For now, just add it as is (standard paragraph merging can be added later)
                    merged_regions.append(region)
                    processed_indices.add(i)
            
            # Log final region count
            if self.merge_config.log_merge_decisions:
                logger.info(
                    f"Paragraph merge complete: {len(regions)} regions -> "
                    f"{len(merged_regions)} regions"
                )
            
            return merged_regions
        
        except Exception as e:
            # Log error and return original regions
            logger.error(
                f"Error during paragraph merge: {e}. "
                "Falling back to original regions."
            )
            return regions
