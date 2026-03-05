"""Region merger for combining text regions into paragraphs.

This module provides region merging capabilities for combining multiple
text regions into single paragraphs, with support for different merge strategies.

本模块提供文本区域合并功能，用于将多个文本区域合并成单个段落，
支持不同的合并策略。
"""

import logging
from typing import List, Tuple, Optional
from src.models.data_models import TextRegion, MergeConfig
from src.config.config_manager import ConfigManager


logger = logging.getLogger(__name__)


class RegionMerger:
    """文本区域合并器
    
    Text region merger.
    
    Merges multiple text regions into single regions using different
    strategies based on the field type and content.
    
    Attributes:
        config: Configuration manager containing merge parameters
        merge_config: Merge configuration with thresholds and tolerances
    """
    
    def __init__(self, config: ConfigManager):
        """初始化合并器
        
        Initialize the region merger.
        
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
    
    def calculate_merged_bbox(
        self,
        regions: List[TextRegion]
    ) -> Tuple[int, int, int, int]:
        """计算合并后的边界框
        
        Calculate the merged bounding box for a list of regions.
        
        The merged bounding box is calculated as:
        - x1: minimum x1 of all regions (leftmost edge)
        - y1: minimum y1 of all regions (topmost edge)
        - x2: maximum x2 of all regions (rightmost edge)
        - y2: maximum y2 of all regions (bottommost edge)
        
        This ensures the merged bounding box contains all original regions.
        
        Args:
            regions: List of text regions to merge
            
        Returns:
            Merged bounding box as (x1, y1, x2, y2)
            
        Raises:
            ValueError: If regions list is empty or contains invalid regions
        """
        if not regions:
            raise ValueError("Cannot calculate merged bbox for empty regions list")
        
        # Validate all regions have valid bounding boxes
        for i, region in enumerate(regions):
            if not region or not region.bbox:
                raise ValueError(f"Region at index {i} has invalid bbox")
            
            if len(region.bbox) != 4:
                raise ValueError(
                    f"Region at index {i} has invalid bbox length: "
                    f"expected 4, got {len(region.bbox)}"
                )
        
        # Extract all bounding box coordinates
        x1_values = [region.bbox[0] for region in regions]
        y1_values = [region.bbox[1] for region in regions]
        x2_values = [region.bbox[2] for region in regions]
        y2_values = [region.bbox[3] for region in regions]
        
        # Calculate merged bounding box
        merged_x1 = min(x1_values)
        merged_y1 = min(y1_values)
        merged_x2 = max(x2_values)
        merged_y2 = max(y2_values)
        
        # Validate the merged bounding box
        if merged_x2 <= merged_x1 or merged_y2 <= merged_y1:
            raise ValueError(
                f"Invalid merged bounding box: ({merged_x1}, {merged_y1}, {merged_x2}, {merged_y2})"
            )
        
        if self.merge_config.log_merge_decisions:
            logger.debug(
                f"Calculated merged bbox for {len(regions)} regions: "
                f"({merged_x1}, {merged_y1}, {merged_x2}, {merged_y2})"
            )
        
        return (merged_x1, merged_y1, merged_x2, merged_y2)
    
    def preserve_text_formatting(
        self,
        regions: List[TextRegion]
    ) -> str:
        """保留文本格式进行合并
        
        Preserve text formatting when merging regions.
        
        This method merges the text content of multiple regions while
        preserving the original formatting, including:
        - Line breaks between regions
        - Whitespace within each region's text
        - Original text order (top to bottom)
        
        The regions are sorted by vertical position (y1 coordinate) before
        merging to ensure correct text order.
        
        Args:
            regions: List of text regions to merge
            
        Returns:
            Merged text with preserved formatting
            
        Raises:
            ValueError: If regions list is empty
        """
        if not regions:
            raise ValueError("Cannot preserve text formatting for empty regions list")
        
        # Sort regions by vertical position (top to bottom)
        sorted_regions = sorted(regions, key=lambda r: r.bbox[1])
        
        # Collect text from each region
        text_parts = []
        for region in sorted_regions:
            if region.text:
                # Preserve the original text including any whitespace
                text_parts.append(region.text)
        
        # Join text parts with newlines to preserve line breaks
        merged_text = '\n'.join(text_parts)
        
        if self.merge_config.log_merge_decisions:
            logger.debug(
                f"Preserved text formatting for {len(regions)} regions: "
                f"merged text length = {len(merged_text)} characters"
            )
        
        return merged_text
    
    def merge_standard(
        self,
        regions: List[TextRegion]
    ) -> TextRegion:
        """标准段落合并
        
        Standard paragraph merge.
        
        Merges multiple text regions into a single region using standard
        paragraph merging rules. This method:
        1. Calculates the merged bounding box
        2. Preserves text formatting (line breaks and whitespace)
        3. Calculates average confidence and font size
        4. Marks the result as paragraph-merged
        5. Stores the original regions for reference
        
        Args:
            regions: List of text regions to merge
            
        Returns:
            Merged text region
            
        Raises:
            ValueError: If regions list is empty
        """
        if not regions:
            raise ValueError("Cannot merge empty regions list")
        
        # Calculate merged bounding box
        merged_bbox = self.calculate_merged_bbox(regions)
        
        # Preserve text formatting
        merged_text = self.preserve_text_formatting(regions)
        
        # Calculate average confidence
        total_confidence = sum(region.confidence for region in regions)
        avg_confidence = total_confidence / len(regions)
        
        # Calculate average font size
        total_font_size = sum(region.font_size for region in regions)
        avg_font_size = int(total_font_size / len(regions))
        
        # Use the angle from the first region (assuming all regions have similar angles)
        angle = regions[0].angle if regions else 0.0
        
        # Create merged region
        merged_region = TextRegion(
            bbox=merged_bbox,
            text=merged_text,
            confidence=avg_confidence,
            font_size=avg_font_size,
            angle=angle,
            is_paragraph_merged=True,
            is_field_content=True,  # Mark as field content to prevent re-pairing
            original_regions=regions.copy(),
            merge_strategy="standard"
        )
        
        if self.merge_config.log_merge_decisions:
            logger.info(
                f"Standard merge: merged {len(regions)} regions into 1 region "
                f"with text: '{merged_text[:50]}...'"
            )
        
        return merged_region
    
    def merge_long_text_field(
        self,
        field_label_region: TextRegion,
        content_regions: List[TextRegion]
    ) -> TextRegion:
        """长文本字段合并
        
        Long text field merge.
        
        Merges content regions for a long text field (like "经营范围").
        This method is similar to merge_standard but:
        1. Does NOT include the field label in the merged region
        2. Only merges the content blocks
        3. Uses the "long_text_field" merge strategy
        4. Applies more relaxed merging rules (handled by caller)
        
        The field label is kept separate to maintain the field label
        separation functionality.
        
        Args:
            field_label_region: The field label region (not included in merge)
            content_regions: List of content regions to merge
            
        Returns:
            Merged text region containing only the field content
            
        Raises:
            ValueError: If content_regions list is empty
        """
        if not content_regions:
            raise ValueError("Cannot merge empty content regions list")
        
        # Calculate merged bounding box for content regions only
        merged_bbox = self.calculate_merged_bbox(content_regions)
        
        # Preserve text formatting for content regions
        merged_text = self.preserve_text_formatting(content_regions)
        
        # Calculate average confidence
        total_confidence = sum(region.confidence for region in content_regions)
        avg_confidence = total_confidence / len(content_regions)
        
        # Calculate average font size
        total_font_size = sum(region.font_size for region in content_regions)
        avg_font_size = int(total_font_size / len(content_regions))
        
        # Use the angle from the first content region
        angle = content_regions[0].angle if content_regions else 0.0
        
        # Get the field label name for reference
        field_label_name = field_label_region.field_label_name
        
        # Create merged region
        merged_region = TextRegion(
            bbox=merged_bbox,
            text=merged_text,
            confidence=avg_confidence,
            font_size=avg_font_size,
            angle=angle,
            is_paragraph_merged=True,
            is_field_content=True,  # Mark as field content to prevent re-pairing
            belongs_to_field=field_label_name,
            original_regions=content_regions.copy(),
            merge_strategy="long_text_field"
        )
        
        if self.merge_config.log_merge_decisions:
            logger.info(
                f"Long text field merge: merged {len(content_regions)} content regions "
                f"for field '{field_label_name}' into 1 region "
                f"with text: '{merged_text[:50]}...'"
            )
        
        return merged_region
