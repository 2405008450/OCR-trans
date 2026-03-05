"""Field content detector for identifying field labels and content blocks.

This module provides field content detection capabilities for identifying
field labels (like "经营范围") and their associated content blocks.

本模块提供字段内容检测功能，用于识别字段标签（如"经营范围"）及其对应的内容块。
"""

import logging
import re
from typing import Optional, List
from src.models.data_models import TextRegion, MergeConfig
from src.config.config_manager import ConfigManager


logger = logging.getLogger(__name__)


class FieldContentDetector:
    """字段内容检测器
    
    Field content detector.
    
    Detects field labels and identifies content blocks that belong to
    specific fields in business license documents.
    
    Attributes:
        config: Configuration manager containing field detection parameters
        merge_config: Merge configuration with field label lists
        known_field_labels: List of known field labels to detect
        long_text_fields: List of fields that typically contain long text
    """
    
    def __init__(self, config: ConfigManager):
        """初始化检测器
        
        Initialize the field content detector.
        
        Args:
            config: Configuration manager instance
        """
        self.config = config
        self.merge_config = self._load_merge_config()
        self.known_field_labels = self.merge_config.known_field_labels
        self.long_text_fields = self.merge_config.long_text_fields
    
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
    
    def detect_field_label(self, region: TextRegion) -> Optional[str]:
        """检测文本区域是否是字段标签
        
        Detect whether a text region is a field label.
        
        This method checks if the text content of a region matches any known
        field labels, handling common variations such as:
        - With or without trailing colon (e.g., "经营范围" vs "经营范围：")
        - With or without spaces (e.g., "经营范围" vs "经营 范围")
        - With or without punctuation
        
        Args:
            region: Text region to check
            
        Returns:
            Field label name if the region is a field label, None otherwise
        """
        if not region or not region.text:
            return None
        
        # Get the text content and normalize it
        text = region.text.strip()
        
        if not text:
            return None
        
        # Check each known field label
        for field_label in self.known_field_labels:
            # Normalize the field label for comparison
            normalized_label = self._normalize_field_label(field_label)
            normalized_text = self._normalize_field_label(text)
            
            # Check for exact match after normalization
            if normalized_text == normalized_label:
                if self.merge_config.log_merge_decisions:
                    logger.debug(
                        f"Detected field label '{field_label}' in region with text '{text}'"
                    )
                return field_label
            
            # Check if the text starts with the field label
            # This handles cases like "经营范围：批发零售..." where the label
            # and content are in the same region
            # We need to check the original text (before normalization) to detect separators
            if text.startswith(field_label):
                # Make sure it's followed by a separator (colon, space, etc.)
                # or is at the end of the text
                remaining = text[len(field_label):]
                if not remaining or remaining[0] in [':', '：', ' ', '\t', '\n']:
                    if self.merge_config.log_merge_decisions:
                        logger.debug(
                            f"Detected field label '{field_label}' at start of region with text '{text}'"
                        )
                    return field_label
        
        # No field label detected
        if self.merge_config.log_merge_decisions:
            logger.debug(f"No field label detected in region with text '{text}'")
        
        return None
    
    def _normalize_field_label(self, text: str) -> str:
        """规范化字段标签文本
        
        Normalize field label text for comparison.
        
        This method removes common variations to enable flexible matching:
        - Removes trailing colons (both Chinese and English)
        - Removes all whitespace
        - Converts to lowercase (for English text)
        
        Args:
            text: Text to normalize
            
        Returns:
            Normalized text
        """
        if not text:
            return ""
        
        # Remove leading/trailing whitespace
        normalized = text.strip()
        
        # Remove trailing colons (both Chinese and English)
        normalized = re.sub(r'[：:]+$', '', normalized)
        
        # Remove all whitespace (spaces, tabs, newlines)
        normalized = re.sub(r'\s+', '', normalized)
        
        # Convert to lowercase for case-insensitive comparison
        # (mainly for English text, doesn't affect Chinese)
        normalized = normalized.lower()
        
        return normalized
    
    def find_field_content_blocks(
        self,
        field_label_region: TextRegion,
        all_regions: List[TextRegion]
    ) -> List[TextRegion]:
        """查找字段标签对应的所有内容块
        
        Find all content blocks that belong to a field label.
        
        This method identifies text regions that are likely to be content
        for the given field label based on spatial relationships:
        - First checks for content on the same line as the field label (right side)
        - Then looks for content blocks below the field label
        - Content blocks should be horizontally aligned
        - Content blocks should be vertically continuous with each other
        
        Args:
            field_label_region: The field label region
            all_regions: All text regions in the document
            
        Returns:
            List of text regions that are content blocks for this field
        """
        if not field_label_region or not all_regions:
            return []
        
        # Import SpatialAnalyzer for spatial relationship analysis
        from src.paragraph_merge.spatial_analyzer import SpatialAnalyzer
        spatial_analyzer = SpatialAnalyzer(self.config)
        
        # Get the field label name to check if it's a long text field
        field_label = self.detect_field_label(field_label_region)
        
        # Determine the vertical gap threshold based on field type
        if field_label and self.is_long_text_field(field_label):
            # Use more relaxed threshold for long text fields
            y_gap_threshold = self.merge_config.long_text_field_y_gap_max
        else:
            # Use standard threshold
            y_gap_threshold = self.merge_config.y_gap_max
        
        # Get horizontal alignment tolerance
        # For long text fields, use more relaxed tolerance
        if field_label and self.is_long_text_field(field_label):
            x_tolerance = self.merge_config.x_tolerance * 10  # 10x tolerance for long text fields
            if self.merge_config.log_merge_decisions:
                logger.debug(
                    f"Using relaxed x_tolerance={x_tolerance}px for long text field '{field_label}'"
                )
        else:
            x_tolerance = self.merge_config.x_tolerance
        
        # Extract field label bounding box
        label_x1, label_y1, label_x2, label_y2 = field_label_region.bbox
        
        # Sort all regions by vertical position (top to bottom)
        sorted_regions = sorted(all_regions, key=lambda r: r.bbox[1])
        
        # Step 1: Check for content on the same line as the field label (right side)
        # This handles cases like "经营范围 品牌设计、营销、策划..."
        same_line_content = None
        for region in sorted_regions:
            # Skip the field label itself
            if region is field_label_region:
                continue
            
            # Skip if this region is also a field label
            if self.detect_field_label(region) is not None:
                continue
            
            # Get region bounding box
            region_x1, region_y1, region_x2, region_y2 = region.bbox
            
            # Check if region is on the same line (vertically overlapping)
            # Two regions are on the same line if their vertical ranges overlap
            vertical_overlap = min(label_y2, region_y2) - max(label_y1, region_y1)
            region_height = region_y2 - region_y1
            
            # Consider it same line if vertical overlap is at least 50% of region height
            if vertical_overlap >= region_height * 0.5:
                # Check if region is to the right of the label
                # Allow small gap or slight overlap
                horizontal_gap = region_x1 - label_x2
                
                # Region is on the same line if:
                # 1. It starts at or after the label ends (horizontal_gap >= 0)
                # 2. Or it slightly overlaps with the label (horizontal_gap < 0 but small)
                if -x_tolerance <= horizontal_gap <= x_tolerance * 2:
                    same_line_content = region
                    if self.merge_config.log_merge_decisions:
                        logger.info(
                            f"Found same-line content for field '{field_label}': "
                            f"text='{region.text[:30]}...', "
                            f"horizontal_gap={horizontal_gap}px, "
                            f"vertical_overlap={vertical_overlap}px"
                        )
                    break
        
        # Step 2: Find all candidate regions below the field label (or below same-line content)
        # Start from the field label or same-line content, whichever is lower
        reference_region = same_line_content if same_line_content else field_label_region
        reference_y2 = reference_region.bbox[3]
        
        if self.merge_config.log_merge_decisions:
            logger.debug(
                f"Looking for candidates below reference region for field '{field_label}': "
                f"reference_y2={reference_y2}"
            )
        
        candidates = []
        for region in sorted_regions:
            # Skip the field label itself
            if region is field_label_region:
                continue
            
            # Skip the same-line content (we'll add it separately)
            if same_line_content and region is same_line_content:
                continue
            
            # Skip if this region is also a field label
            if self.detect_field_label(region) is not None:
                continue
            
            # Get region bounding box
            region_x1, region_y1, region_x2, region_y2 = region.bbox
            
            # Check if region is below the reference region
            # Allow slight overlap (within y_gap_threshold)
            vertical_gap = region_y1 - reference_y2
            if vertical_gap < -y_gap_threshold:
                # Region starts too far above reference end, skip
                # (negative gap means overlap; we allow small overlap)
                continue
            
            # Check horizontal alignment (left edge)
            # For same-line content, use its left edge as reference
            # Otherwise, use label's left or right edge
            if same_line_content:
                # Align with same-line content's left edge
                reference_x1 = same_line_content.bbox[0]
                left_diff = abs(region_x1 - reference_x1)
            else:
                # Align with label's left edge or right edge (content area)
                left_diff = abs(region_x1 - label_x1)
                content_start_x = label_x2
                left_diff_from_content = abs(region_x1 - content_start_x)
                left_diff = min(left_diff, left_diff_from_content)
            
            # Region is aligned if left edge difference is within tolerance
            is_aligned = left_diff <= x_tolerance
            
            if self.merge_config.log_merge_decisions:
                logger.debug(
                    f"Candidate region for field '{field_label}': "
                    f"text='{region.text[:20]}...', "
                    f"y1={region_y1}, reference_y2={reference_y2}, "
                    f"left_diff={left_diff}px, x_tolerance={x_tolerance}px, "
                    f"is_aligned={is_aligned}"
                )
            
            if not is_aligned:
                # Not horizontally aligned, skip
                continue
            
            # This region is a candidate
            candidates.append(region)
        
        if self.merge_config.log_merge_decisions:
            logger.debug(
                f"Found {len(candidates)} candidate(s) for field '{field_label}'"
            )
        
        # Step 3: Build content blocks list starting with same-line content (if any)
        content_blocks = []
        
        if same_line_content:
            content_blocks.append(same_line_content)
            last_region = same_line_content
        else:
            last_region = field_label_region
        
        # Step 4: Find vertically continuous content blocks
        for candidate in candidates:
            # Calculate vertical distance from the last region
            last_y2 = last_region.bbox[3]
            candidate_y1 = candidate.bbox[1]
            vertical_distance = candidate_y1 - last_y2
            
            if self.merge_config.log_merge_decisions:
                logger.debug(
                    f"Checking candidate for field '{field_label}': "
                    f"text='{candidate.text[:20]}...', "
                    f"vertical_distance={vertical_distance}px, "
                    f"threshold={y_gap_threshold}px"
                )
            
            # Check if this candidate is within the threshold distance from the last region
            if vertical_distance <= y_gap_threshold:
                # This candidate is part of the content
                content_blocks.append(candidate)
                last_region = candidate
                
                if self.merge_config.log_merge_decisions:
                    logger.info(
                        f"Found content block for field '{field_label}': "
                        f"text='{candidate.text[:20]}...', "
                        f"vertical_distance={vertical_distance}px"
                    )
            else:
                # Gap is too large, stop looking for more content blocks
                # (we've reached the end of this field's content)
                if self.merge_config.log_merge_decisions:
                    logger.debug(
                        f"Gap too large for field '{field_label}': "
                        f"vertical_distance={vertical_distance}px > threshold={y_gap_threshold}px, "
                        f"stopping search"
                    )
                break
        
        if self.merge_config.log_merge_decisions and content_blocks:
            logger.info(
                f"Found {len(content_blocks)} content block(s) for field '{field_label}' "
                f"(including {1 if same_line_content else 0} same-line content)"
            )
        
        return content_blocks
    
    def is_long_text_field(self, field_label: str) -> bool:
        """判断是否是长文本字段
        
        Determine whether a field is a long text field.
        
        Long text fields (like "经营范围" and "住所") typically span multiple
        lines and require special handling during paragraph merging.
        
        Args:
            field_label: Field label name
            
        Returns:
            True if the field is a long text field, False otherwise
        """
        return field_label in self.long_text_fields
