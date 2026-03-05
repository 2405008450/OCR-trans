"""Spatial analyzer for text region relationships.

This module provides spatial analysis capabilities for determining
relationships between text regions, such as vertical continuity,
horizontal alignment, and font size similarity.

本模块提供文本区域空间关系分析功能，用于判断区域之间的垂直连续性、
水平对齐和字体大小相似性等关系。
"""

import logging
from typing import Optional
from src.models.data_models import TextRegion, MergeConfig
from src.config.config_manager import ConfigManager


logger = logging.getLogger(__name__)


class SpatialAnalyzer:
    """文本区域空间关系分析器
    
    Text region spatial relationship analyzer.
    
    Analyzes spatial relationships between text regions to determine
    if they should be merged into paragraphs.
    
    Attributes:
        config: Configuration manager containing merge parameters
        merge_config: Merge configuration with thresholds and tolerances
    """
    
    def __init__(self, config: ConfigManager):
        """初始化空间分析器
        
        Initialize the spatial analyzer.
        
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
    
    def analyze_vertical_continuity(
        self, 
        region1: TextRegion, 
        region2: TextRegion,
        y_gap_threshold: Optional[int] = None
    ) -> bool:
        """分析两个区域是否垂直连续
        
        Analyze whether two regions are vertically continuous.
        
        Two regions are considered vertically continuous if:
        1. region2 is below region1 (region2.y1 >= region1.y2)
        2. The vertical gap between them is within the threshold
        
        Args:
            region1: First region (upper region)
            region2: Second region (lower region)
            y_gap_threshold: Optional custom vertical gap threshold in pixels.
                           If None, uses the default from merge_config.
        
        Returns:
            True if the two regions are vertically continuous
        """
        # Use custom threshold or default from config
        if y_gap_threshold is None:
            y_gap_threshold = self.merge_config.y_gap_max
        
        # Validate threshold
        if y_gap_threshold < 0:
            logger.warning(
                f"Invalid y_gap_threshold {y_gap_threshold}, using default {self.merge_config.y_gap_max}"
            )
            y_gap_threshold = self.merge_config.y_gap_max
        
        # Extract bounding box coordinates
        x1_1, y1_1, x2_1, y2_1 = region1.bbox
        x1_2, y1_2, x2_2, y2_2 = region2.bbox
        
        # Check if region2 is below region1
        if y1_2 < y2_1:
            # region2 starts before region1 ends, not vertically continuous
            if self.merge_config.log_merge_decisions:
                logger.debug(
                    f"Regions not vertically continuous: region2 (y1={y1_2}) "
                    f"starts before region1 ends (y2={y2_1})"
                )
            return False
        
        # Calculate vertical gap
        vertical_gap = y1_2 - y2_1
        
        # Check if gap is within threshold
        is_continuous = vertical_gap <= y_gap_threshold
        
        if self.merge_config.log_merge_decisions:
            if is_continuous:
                logger.debug(
                    f"Regions vertically continuous: gap={vertical_gap}px "
                    f"<= threshold={y_gap_threshold}px"
                )
            else:
                logger.debug(
                    f"Regions not vertically continuous: gap={vertical_gap}px "
                    f"> threshold={y_gap_threshold}px"
                )
        
        return is_continuous
    
    def analyze_horizontal_alignment(
        self, 
        region1: TextRegion, 
        region2: TextRegion,
        tolerance: Optional[int] = None,
        check_right_edge: bool = False
    ) -> bool:
        """分析两个区域是否水平对齐
        
        Analyze whether two regions are horizontally aligned.
        
        Two regions are considered horizontally aligned if their left edges
        (and optionally right edges) are within the specified tolerance.
        
        According to requirement 3.2, right edge alignment is optional because
        the last line of a paragraph may be shorter.
        
        Args:
            region1: First region
            region2: Second region
            tolerance: Optional custom alignment tolerance in pixels.
                      If None, uses the default from merge_config.
            check_right_edge: If True, also checks right edge alignment.
                            Default is False (only check left edge).
        
        Returns:
            True if the two regions are horizontally aligned
        """
        # Use custom tolerance or default from config
        if tolerance is None:
            tolerance = self.merge_config.x_tolerance
        
        # Validate tolerance
        if tolerance < 0:
            logger.warning(
                f"Invalid tolerance {tolerance}, using default {self.merge_config.x_tolerance}"
            )
            tolerance = self.merge_config.x_tolerance
        
        # Extract left edge coordinates
        x1_1 = region1.bbox[0]
        x1_2 = region2.bbox[0]
        
        # Calculate left edge horizontal difference
        left_diff = abs(x1_1 - x1_2)
        
        # Check if left edges are aligned
        left_aligned = left_diff <= tolerance
        
        # If only checking left edge, return result
        if not check_right_edge:
            if self.merge_config.log_merge_decisions:
                if left_aligned:
                    logger.debug(
                        f"Regions horizontally aligned (left edge): diff={left_diff}px "
                        f"<= tolerance={tolerance}px"
                    )
                else:
                    logger.debug(
                        f"Regions not horizontally aligned (left edge): diff={left_diff}px "
                        f"> tolerance={tolerance}px"
                    )
            return left_aligned
        
        # Extract right edge coordinates
        x2_1 = region1.bbox[2]
        x2_2 = region2.bbox[2]
        
        # Calculate right edge horizontal difference
        right_diff = abs(x2_1 - x2_2)
        
        # Check if right edges are aligned
        right_aligned = right_diff <= tolerance
        
        # Both edges must be aligned
        is_aligned = left_aligned and right_aligned
        
        if self.merge_config.log_merge_decisions:
            if is_aligned:
                logger.debug(
                    f"Regions horizontally aligned (both edges): "
                    f"left_diff={left_diff}px, right_diff={right_diff}px "
                    f"<= tolerance={tolerance}px"
                )
            else:
                logger.debug(
                    f"Regions not horizontally aligned: "
                    f"left_diff={left_diff}px (aligned={left_aligned}), "
                    f"right_diff={right_diff}px (aligned={right_aligned}), "
                    f"tolerance={tolerance}px"
                )
        
        return is_aligned
    
    def analyze_font_size_similarity(
        self,
        region1: TextRegion,
        region2: TextRegion,
        threshold: Optional[float] = None
    ) -> bool:
        """分析两个区域的字体大小是否相近
        
        Analyze whether two regions have similar font sizes.
        
        Two regions are considered to have similar font sizes if the
        relative difference is within the specified threshold.
        
        Args:
            region1: First region
            region2: Second region
            threshold: Optional custom font size difference threshold (ratio).
                      If None, uses the default from merge_config.
        
        Returns:
            True if the font sizes are similar
        """
        # Use custom threshold or default from config
        if threshold is None:
            threshold = self.merge_config.font_size_diff_threshold
        
        # Validate threshold
        if not isinstance(threshold, (int, float)) or not 0.0 <= threshold <= 1.0:
            logger.warning(
                f"Invalid threshold {threshold}, using default {self.merge_config.font_size_diff_threshold}"
            )
            threshold = self.merge_config.font_size_diff_threshold
        
        # Get font sizes
        font_size1 = region1.font_size
        font_size2 = region2.font_size
        
        # Handle zero or negative font sizes
        if font_size1 <= 0 or font_size2 <= 0:
            logger.warning(
                f"Invalid font sizes: region1={font_size1}, region2={font_size2}"
            )
            return False
        
        # Calculate relative difference
        max_size = max(font_size1, font_size2)
        min_size = min(font_size1, font_size2)
        relative_diff = (max_size - min_size) / max_size
        
        # Check if difference is within threshold
        is_similar = relative_diff <= threshold
        
        if self.merge_config.log_merge_decisions:
            if is_similar:
                logger.debug(
                    f"Font sizes similar: diff={relative_diff:.2%} "
                    f"<= threshold={threshold:.2%} "
                    f"(sizes: {font_size1}, {font_size2})"
                )
            else:
                logger.debug(
                    f"Font sizes not similar: diff={relative_diff:.2%} "
                    f"> threshold={threshold:.2%} "
                    f"(sizes: {font_size1}, {font_size2})"
                )
        
        return is_similar
