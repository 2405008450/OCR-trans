"""Text renderer for image translation system.

This module provides the TextRenderer class for rendering translated text
onto images with proper font sizing, positioning, alignment, and visual effects
including anti-aliasing, stroke, shadow, and rotation support.
"""

import logging
import math
from enum import Enum
from typing import Tuple, List, Optional
from pathlib import Path

import numpy as np
import cv2
from PIL import Image, ImageDraw, ImageFont, ImageFilter

from src.config import ConfigManager
from src.models import TextRegion


logger = logging.getLogger(__name__)


class TextAlignment(Enum):
    """Text alignment options."""
    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"


class TextType(Enum):
    """Text type options for font size standardization."""
    MAIN_TITLE = "main_title"      # 主标题
    SUBTITLE = "subtitle"           # 副标题
    FIELD_NAME = "field_name"       # 字段名
    BODY_TEXT = "body_text"         # 正文
    FOOTER = "footer"               # 页脚


class FontSizeCalculator:
    """Calculator for standardized font sizes based on text type and bbox height.
    
    This class provides methods to calculate appropriate font sizes according to
    standardized scaling ratios for different text types.
    """
    
    # Scaling ratios for each text type (min, max)
    SCALE_RATIOS = {
        TextType.MAIN_TITLE: (0.9, 1.0),    # 主标题：0.9-1.0 倍 bbox_height
        TextType.SUBTITLE: (0.5, 0.6),       # 副标题：0.5-0.6 倍 bbox_height
        TextType.FIELD_NAME: (0.85, 0.85),   # 字段名：0.85 倍 bbox_height（固定）
        TextType.BODY_TEXT: (0.75, 0.8),     # 正文：0.75-0.8 倍 bbox_height
        TextType.FOOTER: (0.7, 0.7)          # 页脚：0.7 倍 bbox_height（固定）
    }
    
    # Font size limits
    MIN_FONT_SIZE = 12   # 最小字号 12px（从8px增加，避免印章文字过小）
    MAX_FONT_SIZE = 48  # 最大字号 48px
    
    @staticmethod
    def validate_bbox_height(bbox_height: float) -> None:
        """Validate that bbox_height is a positive number.
        
        Args:
            bbox_height: Bounding box height in pixels
            
        Raises:
            ValueError: If bbox_height is not a number or is not positive
        """
        if not isinstance(bbox_height, (int, float)):
            raise ValueError(f"bbox_height must be a number, got {type(bbox_height).__name__}")
        
        if bbox_height <= 0:
            raise ValueError(f"bbox_height must be positive, got {bbox_height}")
    
    @staticmethod
    def validate_text_type(text_type) -> None:
        """Validate that text_type is a valid TextType enum.
        
        Args:
            text_type: Text type to validate
            
        Raises:
            TypeError: If text_type is not a TextType enum
        """
        if not isinstance(text_type, TextType):
            raise TypeError(f"text_type must be a TextType enum, got {type(text_type).__name__}")
    
    @staticmethod
    def get_scale_ratio(text_type: TextType) -> float:
        """Get the scaling ratio for a given text type.
        
        For text types with a range, returns the middle value.
        For text types with a fixed value, returns that value.
        
        Args:
            text_type: Text type
            
        Returns:
            Scaling ratio as a float
        """
        min_ratio, max_ratio = FontSizeCalculator.SCALE_RATIOS[text_type]
        # Return the middle value of the range
        return (min_ratio + max_ratio) / 2.0
    
    @staticmethod
    def clamp_font_size(font_size: float) -> int:
        """Clamp font size to valid range and convert to integer.
        
        Args:
            font_size: Raw font size (may be float)
            
        Returns:
            Font size clamped to [MIN_FONT_SIZE, MAX_FONT_SIZE] and rounded to int
        """
        # Round to nearest integer
        font_size_int = round(font_size)
        
        # Clamp to valid range
        if font_size_int < FontSizeCalculator.MIN_FONT_SIZE:
            return FontSizeCalculator.MIN_FONT_SIZE
        elif font_size_int > FontSizeCalculator.MAX_FONT_SIZE:
            return FontSizeCalculator.MAX_FONT_SIZE
        else:
            return font_size_int
    
    @staticmethod
    def calculate_font_size(text_type: TextType, bbox_height: float) -> int:
        """Calculate standardized font size based on text type and bbox height.
        
        This method implements the font size standardization algorithm:
        1. Validate inputs
        2. Get scaling ratio for text type
        3. Calculate raw font size (bbox_height × scale_ratio)
        4. Clamp to valid range [8, 48] and convert to integer
        
        Args:
            text_type: Type of text (MAIN_TITLE, SUBTITLE, FIELD_NAME, BODY_TEXT, FOOTER)
            bbox_height: Height of bounding box in pixels
            
        Returns:
            Calculated font size in pixels (integer in range [8, 48])
            
        Raises:
            TypeError: If text_type is not a TextType enum
            ValueError: If bbox_height is not a positive number
            
        Examples:
            >>> calculate_font_size(TextType.MAIN_TITLE, 48)
            46  # 48 × 0.95 = 45.6 → 46
            
            >>> calculate_font_size(TextType.FIELD_NAME, 20)
            17  # 20 × 0.85 = 17.0 → 17
            
            >>> calculate_font_size(TextType.FOOTER, 12)
            8   # 12 × 0.7 = 8.4 → 8
        """
        # Step 1: Validate inputs
        FontSizeCalculator.validate_text_type(text_type)
        FontSizeCalculator.validate_bbox_height(bbox_height)
        
        # Step 2: Get scaling ratio
        scale_ratio = FontSizeCalculator.get_scale_ratio(text_type)
        
        # Step 3: Calculate raw font size
        raw_font_size = bbox_height * scale_ratio
        
        # Step 4: Clamp and convert to integer
        final_font_size = FontSizeCalculator.clamp_font_size(raw_font_size)
        
        return final_font_size


class TextRenderer:
    """Renders translated text onto images.
    
    The TextRenderer handles:
    - Font size calculation to fit text within regions
    - Text position calculation with alignment support
    - Anti-aliased text rendering
    - Text rotation to match original orientation
    - Multi-line text with consistent line spacing
    - Stroke and shadow effects
    - Multiple font selection with fallbacks
    
    Attributes:
        config: Configuration manager instance
        font_family: Primary font family name
        font_fallback: List of fallback font names
        enable_antialiasing: Whether to enable anti-aliasing
        enable_stroke: Whether to enable text stroke
        stroke_width: Width of text stroke in pixels
    """
    
    # Common font paths for different operating systems
    FONT_PATHS = {
        'Arial': [
            'C:/Windows/Fonts/arial.ttf',
            '/usr/share/fonts/truetype/msttcorefonts/Arial.ttf',
            '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
            '/System/Library/Fonts/Helvetica.ttc',
        ],
        'DejaVu Sans': [
            'C:/Windows/Fonts/DejaVuSans.ttf',
            '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
            '/usr/share/fonts/dejavu/DejaVuSans.ttf',
        ],
        'Liberation Sans': [
            'C:/Windows/Fonts/LiberationSans-Regular.ttf',
            '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf',
            '/usr/share/fonts/liberation/LiberationSans-Regular.ttf',
        ],
    }
    
    # Minimum and maximum font sizes
    MIN_FONT_SIZE = 12  # 最小字体12px（从8px增加，避免印章文字过小）
    MAX_FONT_SIZE = 200
    
    # Default line spacing ratio (relative to font size)
    DEFAULT_LINE_SPACING_RATIO = 1.15  # 从1.2降低到1.15，使行间距更紧凑
    
    def __init__(self, config: ConfigManager):
        """Initialize the text renderer.
        
        Args:
            config: Configuration manager instance
        """
        self.config = config
        self.font_family = config.get('rendering.font_family', 'Arial')
        self.font_fallback = config.get('rendering.font_fallback', ['DejaVu Sans', 'Liberation Sans'])
        self.enable_antialiasing = config.get('rendering.enable_antialiasing', True)
        self.enable_stroke = config.get('rendering.enable_stroke', False)
        self.stroke_width = config.get('rendering.stroke_width', 1)
        
        # Cache for loaded fonts
        self._font_cache = {}
        
        # Dynamic wrap threshold for field content (calculated from regions)
        self._field_content_wrap_threshold = 100  # 默认值100px
        
        # Find available font path
        self._font_path = self._find_font_path()
        
        logger.debug(f"TextRenderer initialized with font: {self._font_path}")
    
    def calculate_field_content_wrap_threshold(self, regions: List[TextRegion]) -> int:
        """计算字段内容的动态换行阈值。
        
        阈值 = 所有字段内容中最长文本框宽度的一半
        
        Args:
            regions: 所有文本区域列表
            
        Returns:
            动态计算的阈值（像素），如果没有字段内容则返回默认值100
        """
        if not regions:
            logger.debug("没有文本区域，使用默认阈值100px")
            return 100
        
        # 找出所有字段内容区域
        field_content_regions = [
            r for r in regions 
            if hasattr(r, 'is_field_content') and r.is_field_content
        ]
        
        if not field_content_regions:
            logger.debug("没有字段内容区域，使用默认阈值100px")
            return 100
        
        # 找出最长的字段内容宽度
        max_width = max(r.width for r in field_content_regions)
        
        # 阈值 = 最长宽度的一半
        threshold = int(max_width / 2)
        
        logger.info(
            f"动态计算字段内容换行阈值: "
            f"字段内容数量={len(field_content_regions)}, "
            f"最长宽度={max_width}px, "
            f"阈值={threshold}px (最长宽度的一半)"
        )
        
        return threshold
    
    def set_field_content_wrap_threshold_from_regions(self, regions: List[TextRegion]) -> None:
        """根据文本区域列表设置字段内容换行阈值。
        
        应该在渲染之前调用此方法，以便根据实际文档动态调整阈值。
        
        Args:
            regions: 所有文本区域列表
        """
        self._field_content_wrap_threshold = self.calculate_field_content_wrap_threshold(regions)
        logger.info(f"已设置字段内容换行阈值为: {self._field_content_wrap_threshold}px")
    
    def get_field_content_wrap_threshold(self) -> int:
        """获取当前的字段内容换行阈值。
        
        Returns:
            当前阈值（像素）
        """
        return self._field_content_wrap_threshold
    
    def _find_font_path(self) -> Optional[str]:
        """Find an available font file path.
        
        Searches for the primary font family first, then fallbacks.
        
        Returns:
            Path to an available font file, or None if not found
        """
        # Try primary font
        fonts_to_try = [self.font_family] + self.font_fallback
        
        for font_name in fonts_to_try:
            if font_name in self.FONT_PATHS:
                for path in self.FONT_PATHS[font_name]:
                    if Path(path).exists():
                        logger.debug(f"Found font: {path}")
                        return path
        
        # Try to find any TrueType font
        common_paths = [
            'C:/Windows/Fonts',
            '/usr/share/fonts/truetype',
            '/usr/share/fonts',
            '/System/Library/Fonts',
        ]
        
        for base_path in common_paths:
            base = Path(base_path)
            if base.exists():
                for font_file in base.rglob('*.ttf'):
                    logger.debug(f"Using fallback font: {font_file}")
                    return str(font_file)
        
        logger.warning("No TrueType font found, using default")
        return None
    
    def _get_font(self, size: int) -> ImageFont.FreeTypeFont:
        """Get a font object of the specified size.
        
        Uses caching to avoid reloading fonts.
        
        Args:
            size: Font size in pixels
            
        Returns:
            PIL ImageFont object
        """
        if size in self._font_cache:
            return self._font_cache[size]
        
        try:
            if self._font_path:
                font = ImageFont.truetype(self._font_path, size)
            else:
                # Use default font
                font = ImageFont.load_default()
        except Exception as e:
            logger.warning(f"Failed to load font: {e}, using default")
            font = ImageFont.load_default()
        
        self._font_cache[size] = font
        return font
    
    def measure_text_width(self, text: str, font_size: int) -> int:
        """Measure the width of text at a given font size.
        
        考虑描边宽度和额外的安全边距，确保文字不会被截断。
        
        Args:
            text: Text to measure
            font_size: Font size in pixels
            
        Returns:
            Width of the text in pixels (包含描边和安全边距)
        """
        font = self._get_font(font_size)
        
        try:
            # Use getbbox for accurate measurement (PIL 9.2.0+)
            bbox = font.getbbox(text)
            if bbox:
                base_width = bbox[2] - bbox[0]
            else:
                base_width = 0
        except AttributeError:
            # Fallback for older PIL versions
            try:
                base_width = font.getlength(text)
            except AttributeError:
                # Very old PIL
                base_width = font.getsize(text)[0]
        
        # 添加描边宽度（如果启用）
        # 描边会在文字左右两侧各增加stroke_width像素
        if self.enable_stroke and self.stroke_width > 0:
            base_width += self.stroke_width * 2
        
        # 添加额外的安全边距（5%或至少5像素）
        # 这是为了应对字体渲染的微小差异和边缘效果
        safety_margin = max(int(base_width * 0.05), 5)
        
        return int(base_width + safety_margin)
    
    def measure_text_height(self, text: str, font_size: int) -> int:
        """Measure the height of text at a given font size.
        
        Args:
            text: Text to measure
            font_size: Font size in pixels
            
        Returns:
            Height of the text in pixels
        """
        font = self._get_font(font_size)
        
        try:
            bbox = font.getbbox(text)
            if bbox:
                return bbox[3] - bbox[1]
        except AttributeError:
            pass
        
        try:
            return font.getsize(text)[1]
        except AttributeError:
            return font_size

    def calculate_font_size(self, text: str, region: TextRegion, 
                           max_lines: int = 1, text_type: Optional[TextType] = None) -> int:
        """Calculate the appropriate font size to fit text within a region.
        
        This method supports two modes:
        1. Standardized mode (when text_type is provided): Uses FontSizeCalculator
           to calculate font size based on text type and bbox height according to
           standard scaling ratios.
        2. Legacy mode (when text_type is None): Uses the original adaptive algorithm
           that adjusts font size based on text length and region dimensions.
        
        改进策略（Legacy mode）：
        1. 智能估算文字长度变化（中文→英文通常变长2-4倍）
        2. 从合适的初始字体大小开始，避免过大或过小
        3. 优先保证文字完整显示，而不是严格遵循原始字体大小
        4. 支持自动换行以适应长文本
        
        Args:
            text: Text to render
            region: Target text region
            max_lines: Maximum number of lines allowed (for wrapping)
            text_type: Optional text type for standardized font size calculation.
                      If provided, uses FontSizeCalculator with standard ratios.
                      If None, uses legacy adaptive algorithm.
            
        Returns:
            Calculated font size in pixels
            
        **Validates: Requirements 5.2, 5.3**
        """
        # Standardized mode: Use FontSizeCalculator if text_type is provided
        if text_type is not None:
            try:
                standardized_size = FontSizeCalculator.calculate_font_size(
                    text_type, region.height
                )
                logger.info(
                    f"Using standardized font size: {standardized_size}px "
                    f"for {text_type.value} (bbox_height={region.height})"
                )
                return standardized_size
            except (TypeError, ValueError) as e:
                logger.warning(
                    f"Failed to calculate standardized font size: {e}. "
                    f"Falling back to legacy algorithm."
                )
                # Fall through to legacy algorithm
        
        # Legacy mode: Original adaptive algorithm
        if not text:
            return region.font_size
        
        region_width = region.width
        region_height = region.height
        
        # Get font scale factor from config (default 0.9 = 90% of original size)
        # 提高默认值从0.8到0.9，减少字体缩小程度
        font_scale = self.config.get('rendering.font_scale', 0.9)
        
        # 智能估算初始字体大小
        # 如果文字很长（可能是翻译后的英文），使用更小的初始字体
        text_length = len(text)
        if text_length > 200:
            # 超超长文本（>200字符，如"经营范围"）：使用极保守的缩放
            font_scale = min(font_scale, 0.5)
        elif text_length > 100:
            # 超长文本（>100字符）：使用非常保守的缩放
            font_scale = min(font_scale, 0.5)
        elif text_length > 50:
            # 很长文本（>50字符）：使用更保守的缩放
            font_scale = min(font_scale, 0.6)
        elif text_length > 20:
            # 长文本：使用更保守的缩放
            font_scale = min(font_scale, 0.7)
        elif text_length > 10:
            # 中等长度：使用标准缩放
            font_scale = min(font_scale, 0.85)
        
        # Start with the original font size, scaled down
        font_size = int(region.font_size * font_scale)
        font_size = min(font_size, self.MAX_FONT_SIZE)
        
        # 动态最小字体大小：
        # - 对于原始字体很小的文字（< 15px），使用8px作为最小值
        # - 对于其他文字，使用12px作为最小值（避免印章文字过小）
        dynamic_min_font_size = 8 if region.font_size < 15 else self.MIN_FONT_SIZE
        font_size = max(font_size, dynamic_min_font_size)
        
        # If multi-line is allowed, calculate based on wrapped text
        if max_lines > 1:
            # Try to fit with wrapping
            wrapped_lines = self.wrap_text(text, region_width, font_size, region)
            
            # 特殊处理：如果换行结果是2行，尝试允许超出边界以保持1行
            # 但是对于字段内容，不允许超出太多（最多10%），避免超出右边界
            # 对于底部文字（y > 图片高度的80%），不允许超出边界
            if len(wrapped_lines) == 2:
                # 检查如果允许超出宽度，能否保持在1行
                text_width = self.measure_text_width(text, font_size)
                
                # 判断是否是字段内容
                is_field_content = hasattr(region, 'is_field_content') and region.is_field_content
                
                # 判断是否是底部文字（y坐标 > 图片高度的80%）
                # 注意：这里需要从config或其他地方获取图片高度
                # 暂时使用region的y坐标判断：如果y > 1200，认为是底部文字
                is_bottom_text = region.bbox[1] > 1200
                
                # 字段内容：最多允许超出10%
                # 底部文字：不允许超出（0%）
                # 其他文本：允许超出20%
                if is_bottom_text:
                    max_overflow = 1.0  # 底部文字不允许超出
                elif is_field_content:
                    max_overflow = 1.10  # 字段内容最多超出10%
                else:
                    max_overflow = 1.20  # 其他文本允许超出20%
                
                if text_width <= region_width * max_overflow:
                    # 可以保持1行，直接返回单行
                    wrapped_lines = [text]
                    overflow_pct = (text_width / region_width - 1) * 100
                    print(f"[WRAP OPTIMIZATION] Text '{text[:30]}...' can fit in 1 line with {overflow_pct:.1f}% overflow (width={text_width}, max={region_width}, is_field_content={is_field_content}, is_bottom={is_bottom_text})")
                else:
                    overflow_pct = (text_width / region_width - 1) * 100
                    print(f"[WRAP OPTIMIZATION] Text '{text[:30]}...' needs 2 lines (width={text_width}, max={region_width}, overflow={overflow_pct:.1f}%, is_field_content={is_field_content}, is_bottom={is_bottom_text})")
            
            while len(wrapped_lines) > max_lines and font_size > self.MIN_FONT_SIZE:
                font_size -= 1
                wrapped_lines = self.wrap_text(text, region_width, font_size, region)
                
                # 再次检查2行的情况
                if len(wrapped_lines) == 2:
                    text_width = self.measure_text_width(text, font_size)
                    if text_width <= region_width * 1.2:
                        wrapped_lines = [text]
            
            # Check if height fits
            line_height = int(font_size * self.DEFAULT_LINE_SPACING_RATIO)
            total_height = line_height * len(wrapped_lines)
            
            while total_height > region_height and font_size > self.MIN_FONT_SIZE:
                font_size -= 1
                wrapped_lines = self.wrap_text(text, region_width, font_size, region)
                
                # 再次检查2行的情况
                if len(wrapped_lines) == 2:
                    text_width = self.measure_text_width(text, font_size)
                    if text_width <= region_width * 1.2:
                        wrapped_lines = [text]
                
                line_height = int(font_size * self.DEFAULT_LINE_SPACING_RATIO)
                total_height = line_height * len(wrapped_lines)
            
            # 额外检查：确保每一行都不超出宽度
            # 对于长文本，使用更宽松的宽度检查（允许超出一些）
            # 但如果只有1行（通过上面的优化得到的），允许超出20%
            if len(wrapped_lines) == 1:
                safe_width = region_width * 1.2  # 单行允许超出20%
            else:
                safe_width = region_width * 0.98  # 多行使用98%的宽度，留出2%的安全边距
            
            max_line_width = max(self.measure_text_width(line, font_size) for line in wrapped_lines)
            
            while max_line_width > safe_width and font_size > self.MIN_FONT_SIZE:
                font_size -= 1
                wrapped_lines = self.wrap_text(text, region_width, font_size, region)
                
                # 再次检查2行的情况
                if len(wrapped_lines) == 2:
                    text_width = self.measure_text_width(text, font_size)
                    if text_width <= region_width * 1.2:
                        wrapped_lines = [text]
                
                if len(wrapped_lines) == 1:
                    safe_width = region_width * 1.2
                else:
                    safe_width = region_width * 0.98
                
                max_line_width = max(self.measure_text_width(line, font_size) for line in wrapped_lines)
            
            return font_size
        
        # Single line mode - reduce font size until text fits
        text_width = self.measure_text_width(text, font_size)
        
        # 智能安全边距：短文本使用更宽松的边距，长文本使用保守边距
        # 由于measure_text_width已经包含了描边和额外边距，这里不需要太保守
        word_count = len(text.split())
        if word_count <= 3:
            # 短文本（1-3个单词）：使用95%宽度，优先保持单行
            safe_width = region_width * 0.95
        else:
            # 长文本（4+个单词）：使用90%宽度，为换行预留空间
            safe_width = region_width * 0.90
        
        while text_width > safe_width and font_size > self.MIN_FONT_SIZE:
            font_size -= 1
            text_width = self.measure_text_width(text, font_size)
        
        # Also check height
        text_height = self.measure_text_height(text, font_size)
        safe_height = region_height * 0.90
        
        while text_height > safe_height and font_size > self.MIN_FONT_SIZE:
            font_size -= 1
            text_height = self.measure_text_height(text, font_size)
        
        # 如果字体太小，记录警告
        if font_size <= self.MIN_FONT_SIZE:
            logger.warning(f"Font size reached minimum ({self.MIN_FONT_SIZE}px) for text: '{text[:30]}...'")
        
        return font_size
    
    def wrap_text(self, text: str, max_width: int, font_size: int, region: Optional[TextRegion] = None) -> List[str]:
        """Wrap text to fit within a maximum width.
        
        Breaks text into multiple lines so each line fits within max_width.
        优化策略：
        1. 优先按单词换行
        2. 如果单词太长，按字符拆分
        3. 避免最后一个字符单独成行（允许超出20%宽度）
        4. 避免最后剩下1-2个字符单独成行
        5. 短文本（≤3个单词）允许超出10%宽度以保持单行
        6. 避免最后一行只有1-2个单词（尝试平衡分配）
        7. 如果是字段内容且区域宽度≤100像素，不进行换行（保持单行）
        
        Args:
            text: Text to wrap
            max_width: Maximum width in pixels
            font_size: Font size to use for measurement
            region: Optional TextRegion to check if it's field content
            
        Returns:
            List of text lines
            
        **Validates: Requirements 5.3**
        """
        if not text:
            return []
        
        if max_width <= 0:
            return [text]
        
        # 新规则：如果是字段内容且区域宽度≤动态阈值，不进行换行
        # 动态阈值 = 所有字段内容中最长文本框宽度的一半
        # 只对字段内容应用此规则，其他文本不受影响
        is_field_content = region and hasattr(region, 'is_field_content') and region.is_field_content
        if is_field_content and max_width <= self._field_content_wrap_threshold:
            logger.info(
                f"字段内容区域宽度≤动态阈值 ({max_width}px ≤ {self._field_content_wrap_threshold}px)，"
                f"不进行换行: '{text[:30]}...'"
            )
            return [text]
        
        # Check if text already fits
        text_width = self.measure_text_width(text, font_size)
        if text_width <= max_width:
            return [text]
        
        # 短文本（≤3个单词）：允许超出15%宽度以保持单行
        word_count = len(text.split())
        if word_count <= 3 and text_width <= max_width * 1.05:  # 改成1.05，只允许超出5%
            return [text]
        
        # 日期类文本（包含数字、逗号、冒号）：允许超出20%宽度以保持单行
        # 例如："Date of Establishment: April 13, 2018"
        # 注意：只对纯日期文本（如"April 13, 2016"）应用此规则，不对字段名（如"Date of Establishment:"）应用
        # 判断条件：必须包含数字，且单词数量≤4（纯日期通常不超过4个单词）
        has_numbers = any(char.isdigit() for char in text)
        has_punctuation = any(char in ',:' for char in text)
        if has_numbers and has_punctuation and word_count <= 4 and text_width <= max_width * 1.05:  # 改成1.05，只允许超出5%
            return [text]
        
        lines = []
        words = text.split()
        
        if not words:
            return [text]
        
        current_line = words[0]
        
        for word in words[1:]:
            test_line = current_line + ' ' + word
            test_width = self.measure_text_width(test_line, font_size)
            
            # 调试：输出每次测试的宽度
            if 'Date' in text or 'Establishment' in text:
                print(f"[WRAP TEST] current='{current_line}', word='{word}', test='{test_line}', width={test_width}, max={max_width}")
            
            # 严格控制宽度，不允许超出
            if test_width <= max_width:  # 改成不允许超出
                current_line = test_line
            else:
                lines.append(current_line)
                current_line = word
        
        if current_line:
            lines.append(current_line)
        
        # 优化：如果最后一行只有1-2个单词，尝试从上一行移一些单词过来
        if len(lines) >= 2:
            last_line_words = lines[-1].split()
            if len(last_line_words) <= 2:
                # 尝试从倒数第二行移一些单词到最后一行
                second_last_words = lines[-2].split()
                if len(second_last_words) > 2:
                    # 尝试移动1-2个单词
                    for move_count in range(1, min(3, len(second_last_words))):
                        # 重新分配单词
                        new_second_last = ' '.join(second_last_words[:-move_count])
                        new_last = ' '.join(second_last_words[-move_count:] + last_line_words)
                        
                        # 检查两行是否都能放下（允许超出10%）
                        second_last_width = self.measure_text_width(new_second_last, font_size)
                        last_width = self.measure_text_width(new_last, font_size)
                        
                        if second_last_width <= max_width * 1.10 and last_width <= max_width * 1.10:
                            lines[-2] = new_second_last
                            lines[-1] = new_last
                            break
        
        # If a single word is too long, break it character by character
        # 改进：只对非常长的单词（>15个字符）才按字符拆分
        # 对于正常长度的单词（如"Establishment"），允许超出宽度以保持完整
        final_lines = []
        for line in lines:
            line_width = self.measure_text_width(line, font_size)
            
            # 如果这一行只有一个单词，且长度≤15个字符，允许超出20%宽度
            if ' ' not in line and len(line) <= 15 and line_width <= max_width * 1.20:
                final_lines.append(line)
                continue
            
            # 如果超出宽度，需要拆分
            if line_width > max_width * 1.05:  # 允许超出5%
                # Break long word
                chars = list(line)
                if len(chars) <= 2:
                    # 如果整个单词只有1-2个字符，直接保留
                    final_lines.append(line)
                    continue
                
                current = chars[0]
                for i, char in enumerate(chars[1:], 1):
                    test = current + char
                    is_last_char = (i == len(chars) - 1)
                    is_second_last = (i == len(chars) - 2)
                    
                    if self.measure_text_width(test, font_size) <= max_width:
                        current = test
                    else:
                        # 如果当前字符是最后一个或倒数第二个，尝试把它加到当前行
                        # 即使稍微超出宽度也可以接受（避免1-2个字符单独一行）
                        if (is_last_char or is_second_last) and len(current) > 2:
                            # 检查加上剩余字符是否超出太多（超出25%以内可接受）
                            remaining = ''.join(chars[i:])
                            test_with_remaining = current + remaining
                            test_width = self.measure_text_width(test_with_remaining, font_size)
                            if test_width <= max_width * 1.25:
                                # 可以接受，把剩余字符都加上
                                current = test_with_remaining
                                break
                        
                        final_lines.append(current)
                        current = char
                
                if current:
                    # 如果最后剩下的是1-2个字符，尝试合并到上一行
                    if len(current) <= 2 and final_lines:
                        last_line = final_lines[-1]
                        combined = last_line + current
                        combined_width = self.measure_text_width(combined, font_size)
                        # 如果合并后不超过宽度的125%，就合并
                        if combined_width <= max_width * 1.25:
                            final_lines[-1] = combined
                        else:
                            final_lines.append(current)
                    else:
                        final_lines.append(current)
            else:
                final_lines.append(line)
        
        return final_lines if final_lines else [text]
    
    def should_wrap_text(self, text: str, region: TextRegion) -> bool:
        """Determine if text should be wrapped to multiple lines.
        
        Args:
            text: Text to check
            region: Target region
            
        Returns:
            True if text should be wrapped
        """
        if not text:
            return False
        
        # Calculate minimum font size that would fit single line
        font_size = self.calculate_font_size(text, region, max_lines=1)
        
        # If font size is too small, wrapping is needed
        return font_size < self.MIN_FONT_SIZE + 2

    def calculate_text_position(self, text: str, region: TextRegion, 
                                font_size: int,
                                alignment: TextAlignment = TextAlignment.CENTER) -> Tuple[int, int]:
        """Calculate the position to draw text within a region.
        
        Calculates the (x, y) position based on the region's center point
        and the specified alignment. Considers font baseline for proper
        vertical alignment.
        
        Args:
            text: Text to position
            region: Target text region
            font_size: Font size being used
            alignment: Text alignment (left, center, right)
            
        Returns:
            Tuple of (x, y) coordinates for text drawing
            
        **Validates: Requirements 5.4, 6.1, 6.2, 6.4**
        """
        x1, y1, x2, y2 = region.bbox
        region_width = x2 - x1
        region_height = y2 - y1
        
        # 优先级顺序（与calculate_multiline_positions保持一致）：
        # 1. 字段标签：使用unified_left_boundary（字段标签的统一左边界）
        # 2. 字段内容：使用unified_content_left_boundary（字段内容的统一左边界）
        # 3. 其他：使用bbox的x1
        original_x1 = x1
        is_field_label = hasattr(region, 'is_field_label') and region.is_field_label
        is_field_content = hasattr(region, 'is_field_content') and region.is_field_content
        
        if is_field_label and hasattr(region, 'unified_left_boundary') and alignment == TextAlignment.LEFT:
            # 字段标签：使用字段标签的统一左边界
            x1 = region.unified_left_boundary
            logger.info(f"[POSITION] Field label '{region.text}' using unified_left_boundary={x1} (bbox_x1={original_x1})")
        elif is_field_content and hasattr(region, 'unified_content_left_boundary') and alignment == TextAlignment.LEFT:
            # 字段内容：使用字段内容的统一左边界
            x1 = region.unified_content_left_boundary
            logger.info(f"[POSITION] Field content '{region.text[:20]}...' using unified_content_left_boundary={x1} (bbox_x1={original_x1})")
        else:
            logger.debug(f"[POSITION] Using bbox x1={x1}")
        
        # Get text dimensions
        text_width = self.measure_text_width(text, font_size)
        text_height = self.measure_text_height(text, font_size)
        
        # Calculate horizontal position based on alignment
        if alignment == TextAlignment.LEFT:
            x = x1
        elif alignment == TextAlignment.RIGHT:
            x = x2 - text_width
        else:  # CENTER
            center_x = (x1 + x2) // 2
            x = center_x - text_width // 2
        
        # Calculate vertical position (center vertically with baseline adjustment)
        center_y = (y1 + y2) // 2
        
        # Get font metrics for baseline alignment
        font = self._get_font(font_size)
        try:
            ascent, descent = font.getmetrics()
            # Adjust for baseline
            y = center_y - (ascent - descent) // 2
        except AttributeError:
            # Fallback: center based on text height
            y = center_y - text_height // 2
        
        return (int(x), int(y))
    
    def calculate_multiline_positions(self, lines: List[str], region: TextRegion,
                                      font_size: int,
                                      alignment: TextAlignment = TextAlignment.CENTER,
                                      line_spacing: Optional[float] = None) -> List[Tuple[int, int]]:
        """Calculate positions for multiple lines of text.
        
        Args:
            lines: List of text lines
            region: Target text region
            font_size: Font size being used
            alignment: Text alignment
            line_spacing: Line spacing ratio (default: 1.2)
            
        Returns:
            List of (x, y) positions for each line
            
        **Validates: Requirements 6.3**
        """
        if not lines:
            return []
        
        if line_spacing is None:
            line_spacing = self.DEFAULT_LINE_SPACING_RATIO
        
        x1, y1, x2, y2 = region.bbox
        region_height = y2 - y1
        
        # 优先级顺序：
        # 1. 字段标签：使用unified_left_boundary（字段标签的统一左边界）
        # 2. 字段内容：使用unified_content_left_boundary（字段内容的统一左边界）
        # 3. 其他：使用original_x1（单个字段的原始左边界）或bbox的x1
        original_x1 = x1
        is_field_label = hasattr(region, 'is_field_label') and region.is_field_label
        is_field_content = hasattr(region, 'is_field_content') and region.is_field_content
        
        if is_field_label and hasattr(region, 'unified_left_boundary') and alignment == TextAlignment.LEFT:
            # 字段标签：使用字段标签的统一左边界
            x1 = region.unified_left_boundary
            logger.info(f"[MULTILINE] Field label '{region.text}' using unified_left_boundary={x1} (bbox_x1={original_x1})")
        elif is_field_content and hasattr(region, 'unified_content_left_boundary') and alignment == TextAlignment.LEFT:
            # 字段内容：使用字段内容的统一左边界
            x1 = region.unified_content_left_boundary
            logger.info(f"[MULTILINE] Field content '{region.text[:20]}...' using unified_content_left_boundary={x1} (bbox_x1={original_x1})")
        elif hasattr(region, 'original_x1') and alignment == TextAlignment.LEFT:
            # 其他：使用原始左边界
            x1 = region.original_x1
            logger.debug(f"[MULTILINE] Using original_x1={x1} (bbox_x1={original_x1})")
        else:
            logger.debug(f"[MULTILINE] Using bbox x1={x1}")
        
        # Calculate line height
        line_height = int(font_size * line_spacing)
        total_height = line_height * len(lines)
        
        # 判断是否是字段内容（有is_field_content标记）
        is_field_content = hasattr(region, 'is_field_content') and region.is_field_content
        
        # 判断是否是段落合并后的内容（有is_paragraph_merged标记）
        is_paragraph_merged = hasattr(region, 'is_paragraph_merged') and region.is_paragraph_merged
        
        # 判断是否是高度很大的区域（高度 > 宽度的0.6倍）
        # 这种区域通常是OCR识别的长文本区域，应该使用顶部对齐
        region_width = x2 - x1
        is_tall_region = region_height > region_width * 0.6
        
        # 判断是否是多行文本（行数 >= 3）
        # 即使区域高度不大，如果文本需要多行显示，也应该使用顶部对齐
        is_multiline_text = len(lines) >= 3
        
        # Start position
        # 使用顶部对齐的条件：
        # 1. 字段内容行数很多（>= 4行）
        # 2. 高度很大的区域（高度 > 宽度的0.6倍）
        # 3. 段落合并后的内容（is_paragraph_merged=True）
        # 4. 多行文本（>= 3行）- 新增条件
        # 这样可以避免多行内容（如营业范围、二维码说明文字）顶部有大量空白
        # 对于1-2行的字段内容（非段落合并），仍然使用垂直居中，保证和字段标签在同一水平线上
        if (is_field_content and len(lines) >= 4) or is_tall_region or is_paragraph_merged or is_multiline_text:
            # 多行字段内容（>=4行）或高度很大的区域或段落合并后的内容或多行文本：从顶部开始渲染（顶部对齐）
            # 字段内容：完全与上边界对齐（不加margin），确保翻译文本框的上边界与原始文本框对齐
            # 非字段内容：横版增加3px margin，避免超出边界
            is_horizontal = self.config.get_orientation() == 'horizontal'
            top_margin = 0 if is_field_content else (3 if is_horizontal else 0)
            start_y = y1 + top_margin
            if is_tall_region:
                margin_info = f" (y1+{top_margin})" if top_margin > 0 else ""
                print(f"[MULTILINE DEBUG] Tall region '{lines[0][:20]}...': bbox=({x1},{y1},{x2},{y2}), height={region_height}, width={region_width}, start_y={start_y}{margin_info}, lines={len(lines)}, top-aligned{'with margin' if top_margin > 0 else ''}")
                logger.debug(f"[MULTILINE] Tall region (height={region_height} > width*0.6={region_width*0.6:.0f}): top-aligned at y={start_y} (with {top_margin}px margin)")
            elif is_paragraph_merged:
                margin_info = f" (y1+{top_margin})" if top_margin > 0 else ""
                print(f"[MULTILINE DEBUG] Paragraph-merged content '{lines[0][:20]}...': bbox=({x1},{y1},{x2},{y2}), start_y={start_y}{margin_info}, lines={len(lines)}, top-aligned{'with margin' if top_margin > 0 else ''}")
                logger.debug(f"[MULTILINE] Paragraph-merged content: top-aligned at y={start_y} (with {top_margin}px margin)")
            elif is_multiline_text:
                margin_info = f" (y1+{top_margin})" if top_margin > 0 else ""
                print(f"[MULTILINE DEBUG] Multiline text (>= 3 lines) '{lines[0][:20]}...': bbox=({x1},{y1},{x2},{y2}), start_y={start_y}{margin_info}, lines={len(lines)}, top-aligned{'with margin' if top_margin > 0 else ''}")
                logger.debug(f"[MULTILINE] Multiline text (>= 3 lines): top-aligned at y={start_y} (with {top_margin}px margin)")
            else:
                margin_info = f" (y1+{top_margin})" if top_margin > 0 else ""
                print(f"[MULTILINE DEBUG] Multi-line field content (>= 4 lines) '{lines[0][:20]}...': bbox=({x1},{y1},{x2},{y2}), start_y={start_y}{margin_info}, lines={len(lines)}, top-aligned{'with margin' if top_margin > 0 else ''}")
                logger.debug(f"[MULTILINE] Multi-line field content (>= 4 lines): top-aligned at y={start_y} (with {top_margin}px margin)")
        else:
            # 单行或少量多行字段内容（1-2行）或其他文本：垂直居中
            # 这样可以保证字段内容和字段标签在同一水平线上
            center_y = (y1 + y2) // 2
            start_y = center_y - total_height // 2 + font_size // 2
            if is_field_content:
                print(f"[MULTILINE DEBUG] Field content (1-2 lines) '{lines[0][:20]}...': bbox=({x1},{y1},{x2},{y2}), start_y={start_y}, lines={len(lines)}, vertically centered")
                logger.debug(f"[MULTILINE] Field content (1-2 lines): vertically centered at y={start_y}")
            else:
                logger.debug(f"[MULTILINE] Non-field content: vertically centered at y={start_y}")
        
        positions = []
        for i, line in enumerate(lines):
            # Calculate x position based on alignment
            text_width = self.measure_text_width(line, font_size)
            
            if alignment == TextAlignment.LEFT:
                x = x1
            elif alignment == TextAlignment.RIGHT:
                x = x2 - text_width
            else:  # CENTER
                center_x = (x1 + x2) // 2
                x = center_x - text_width // 2
            
            y = start_y + i * line_height
            positions.append((int(x), int(y)))
        
        return positions
    
    def detect_alignment(self, region: TextRegion) -> TextAlignment:
        """Detect the likely text alignment based on region properties.
        
        Args:
            region: Text region to analyze
            
        Returns:
            Detected text alignment
        """
        # 使用左对齐，避免文本变长时向左偏移导致重叠
        # 左对齐可以确保文本从原始区域的左边界开始，不会侵占左边区域的空间
        return TextAlignment.LEFT
    
    def _is_normalizable_field(self, text: str) -> bool:
        """检查文本是否是需要标准化的字段。
        
        Args:
            text: 文本内容（中文或英文）
            
        Returns:
            True如果是需要标准化的字段
        """
        normalize_config = self.config.get('rendering.normalize_fields', {})
        
        if not normalize_config.get('enabled', False):
            return False
        
        field_groups = normalize_config.get('field_groups', [])
        
        for group in field_groups:
            chinese_fields = group.get('chinese_fields', [])
            english_fields = group.get('english_fields', [])
            
            # 检查是否匹配中文或英文字段名
            if text in chinese_fields or text in english_fields:
                logger.info(f"字段 '{text}' 匹配标准化列表")
                return True
        
        return False
    
    def _normalize_field_bbox(self, region: TextRegion, all_regions: list) -> TextRegion:
        """标准化字段的边界框，使同组字段具有相同的宽度。
        
        Args:
            region: 当前文本区域
            all_regions: 所有文本区域列表（用于计算最大宽度）
            
        Returns:
            标准化后的TextRegion（如果不需要标准化则返回原region）
        """
        print(f"[NORMALIZE_BBOX] 调用标准化方法，字段='{region.text}'")
        
        normalize_config = self.config.get('rendering.normalize_fields', {})
        if not normalize_config.get('enabled', False):
            logger.debug(f"字段标准化未启用")
            print(f"[NORMALIZE_BBOX] 标准化未启用")
            return region
        
        print(f"[NORMALIZE_BBOX] 标准化已启用，开始检查字段")
        
        # 检查当前字段是否需要标准化
        if not self._is_normalizable_field(region.text):
            logger.debug(f"字段 '{region.text}' 不在标准化列表中")
            return region
        
        logger.info(f"检测到需要标准化的字段: '{region.text}'")
        print(f"[NORMALIZE_BBOX] OK 检测到需要标准化的字段: '{region.text}'")
        
        # 找到当前字段所属的组
        field_groups = normalize_config.get('field_groups', [])
        current_group = None
        
        for group in field_groups:
            chinese_fields = group.get('chinese_fields', [])
            english_fields = group.get('english_fields', [])
            
            if region.text in chinese_fields or region.text in english_fields:
                current_group = group
                break
        
        if current_group is None:
            logger.debug(f"字段 '{region.text}' 未找到所属组")
            return region
        
        # 计算该组所有字段的最大宽度
        chinese_fields = current_group.get('chinese_fields', [])
        english_fields = current_group.get('english_fields', [])
        all_field_names = chinese_fields + english_fields
        
        max_width = region.width  # 默认使用当前宽度
        
        for other_region in all_regions:
            if other_region.text in all_field_names:
                max_width = max(max_width, other_region.width)
                logger.debug(f"  字段 '{other_region.text}' 宽度={other_region.width}px")
                print(f"[NORMALIZE_BBOX]   字段 '{other_region.text}' 宽度={other_region.width}px")
        
        logger.info(f"该组最大宽度: {max_width}px")
        print(f"[NORMALIZE_BBOX] 该组最大宽度: {max_width}px")
        
        # 如果当前宽度已经是最大的，不需要扩展
        if region.width >= max_width:
            logger.info(f"字段 '{region.text}' 宽度已是最大，无需扩展")
            print(f"[NORMALIZE_BBOX] 字段 '{region.text}' 宽度已是最大，无需扩展")
            
            # 即使不需要扩展，也要确保所有属性被正确设置
            # 如果原region已经有所有必要的属性，直接返回
            if (hasattr(region, 'is_field_label') and region.is_field_label and
                hasattr(region, 'uses_unified_font_size') and region.uses_unified_font_size):
                logger.debug(f"字段 '{region.text}' 已有所有必要属性，直接返回")
                return region
            
            # 如果原region缺少某些属性，创建一个新的region并设置所有属性
            no_expand_region = TextRegion(
                bbox=region.bbox,
                text=region.text,
                confidence=region.confidence,
                font_size=region.font_size,
                angle=region.angle,
                is_vertical_merged=getattr(region, 'is_vertical_merged', False)
            )
            
            # 保留所有属性
            if hasattr(region, 'original_x1'):
                no_expand_region.original_x1 = region.original_x1
            if hasattr(region, 'unified_left_boundary'):
                no_expand_region.unified_left_boundary = region.unified_left_boundary
            if hasattr(region, 'unified_content_left_boundary'):
                no_expand_region.unified_content_left_boundary = region.unified_content_left_boundary
            if hasattr(region, 'is_field_content'):
                no_expand_region.is_field_content = region.is_field_content
            if hasattr(region, 'is_paragraph_merged'):
                no_expand_region.is_paragraph_merged = region.is_paragraph_merged
            if hasattr(region, 'belongs_to_field'):
                no_expand_region.belongs_to_field = region.belongs_to_field
            if hasattr(region, 'uses_unified_font_size'):
                no_expand_region.uses_unified_font_size = region.uses_unified_font_size
                logger.info(f"[FIELD_LABEL] 保留字段 '{region.text}' 的统一字体大小标记 (uses_unified_font_size=True)")
            
            # 标记为字段标签
            no_expand_region.is_field_label = True
            logger.info(f"[FIELD_LABEL] 字段 '{region.text}' 无需扩展，但标记为字段标签 (is_field_label=True)")
            
            return no_expand_region
        
        # 扩展边界框（保持左边界不变，只扩展右边界）
        x1, y1, x2, y2 = region.bbox
        new_x2 = x1 + max_width
        
        logger.info(
            f"标准化字段边界框: '{region.text}' "
            f"原宽度={region.width}px, 新宽度={max_width}px, "
            f"bbox=({x1},{y1})->({new_x2},{y2})"
        )
        
        # 创建新的TextRegion，保留original_x1、unified_left_boundary和unified_content_left_boundary属性（如果存在）
        normalized_region = TextRegion(
            bbox=(x1, y1, new_x2, y2),
            text=region.text,
            confidence=region.confidence,
            font_size=region.font_size,
            angle=region.angle,
            is_vertical_merged=getattr(region, 'is_vertical_merged', False)
        )
        
        # 保留original_x1属性（如果原region有的话）
        if hasattr(region, 'original_x1'):
            normalized_region.original_x1 = region.original_x1
        
        # 保留unified_left_boundary属性（如果原region有的话）
        if hasattr(region, 'unified_left_boundary'):
            normalized_region.unified_left_boundary = region.unified_left_boundary
        
        # 保留unified_content_left_boundary属性（如果原region有的话）
        if hasattr(region, 'unified_content_left_boundary'):
            normalized_region.unified_content_left_boundary = region.unified_content_left_boundary
        
        # 保留is_field_content属性（如果原region有的话）
        if hasattr(region, 'is_field_content'):
            normalized_region.is_field_content = region.is_field_content
        
        # 保留is_paragraph_merged属性（如果原region有的话）
        if hasattr(region, 'is_paragraph_merged'):
            normalized_region.is_paragraph_merged = region.is_paragraph_merged
        
        # 保留belongs_to_field属性（如果原region有的话）
        if hasattr(region, 'belongs_to_field'):
            normalized_region.belongs_to_field = region.belongs_to_field
        
        # 保留uses_unified_font_size属性（如果原region有的话）
        if hasattr(region, 'uses_unified_font_size'):
            normalized_region.uses_unified_font_size = region.uses_unified_font_size
            logger.info(f"[FIELD_LABEL] 保留字段 '{region.text}' 的统一字体大小标记 (uses_unified_font_size=True)")
        
        # 标记为字段标签（横版模式下，标准化的字段都是字段标签）
        normalized_region.is_field_label = True
        logger.info(f"[FIELD_LABEL] 标记字段 '{region.text}' 为字段标签 (is_field_label=True)")
        
        return normalized_region

    def _detect_seal_overlap(self, image: np.ndarray, text_bbox: Tuple[int, int, int, int]) -> bool:
        """检测文本框是否与印章重叠
        
        通过检测文本框区域内的红色像素来判断是否与印章重叠
        
        Args:
            image: 图像数组 (BGR格式)
            text_bbox: 文本框边界 (x1, y1, x2, y2)
            
        Returns:
            True 如果检测到重叠，False 否则
        """
        x1, y1, x2, y2 = text_bbox
        height, width = image.shape[:2]
        
        # 确保边界在图像范围内
        x1 = max(0, min(x1, width - 1))
        y1 = max(0, min(y1, height - 1))
        x2 = max(0, min(x2, width))
        y2 = max(0, min(y2, height))
        
        if x2 <= x1 or y2 <= y1:
            return False
        
        # 提取文本框区域
        region_img = image[y1:y2, x1:x2]
        
        if region_img.size == 0:
            return False
        
        # 转换到HSV色彩空间以检测红色
        hsv = cv2.cvtColor(region_img, cv2.COLOR_BGR2HSV)
        
        # 定义红色范围（印章通常是红色）
        # 红色在HSV中有两个范围
        lower_red1 = np.array([0, 100, 100])
        upper_red1 = np.array([10, 255, 255])
        lower_red2 = np.array([160, 100, 100])
        upper_red2 = np.array([180, 255, 255])
        
        # 创建红色掩码
        mask1 = cv2.inRange(hsv, lower_red1, upper_red1)
        mask2 = cv2.inRange(hsv, lower_red2, upper_red2)
        red_mask = cv2.bitwise_or(mask1, mask2)
        
        # 计算红色像素比例
        red_pixel_count = np.count_nonzero(red_mask)
        total_pixels = region_img.shape[0] * region_img.shape[1]
        red_ratio = red_pixel_count / total_pixels if total_pixels > 0 else 0
        
        # 如果红色像素超过5%，认为与印章重叠
        overlap_threshold = 0.05
        is_overlap = red_ratio > overlap_threshold
        
        if is_overlap:
            logger.info(
                f"检测到文本框与印章重叠: bbox={text_bbox}, "
                f"红色像素比例={red_ratio:.2%}"
            )
        
        return is_overlap

    def render_text(self, image: np.ndarray, region: TextRegion, text: str,
                   background_color: Tuple[int, int, int],
                   text_color: Optional[Tuple[int, int, int]] = None,
                   alignment: TextAlignment = TextAlignment.CENTER) -> np.ndarray:
        """Render text onto an image within a specified region.
        
        Renders the text with anti-aliasing, proper positioning, and
        optional rotation to match the original text orientation.
        
        如果检测到文本框与印章重叠，会自动调整换行策略以避免遮挡印章。
        
        Args:
            image: Input image as numpy array (BGR format)
            region: Target text region
            text: Text to render
            background_color: Background color (RGB)
            text_color: Text color (RGB), auto-calculated if None
            alignment: Text alignment
            
        Returns:
            Image with rendered text (BGR format)
            
        **Validates: Requirements 5.1, 6.3, 6.5**
        """
        if not text:
            return image.copy()
        
        # 检测文本框是否与印章重叠
        seal_overlap_detected = self._detect_seal_overlap(image, region.bbox)
        if seal_overlap_detected:
            logger.info(f"文本 '{text[:30]}...' 与印章重叠，将调整换行策略")
        
        # Convert numpy array to PIL Image
        image_rgb = image[:, :, ::-1]  # BGR to RGB
        pil_image = Image.fromarray(image_rgb)
        
        # 对于垂直合并的文本（或高度远大于宽度的文本），扩展区域以容纳横排显示
        render_region = region
        
        # 判断是否是竖排文本：
        # 1. 有is_vertical_merged标记
        # 2. 或者高度远大于宽度（height > width * 3）
        is_vertical_text = (
            (hasattr(region, 'is_vertical_merged') and region.is_vertical_merged) or
            (region.height > region.width * 3)
        )
        
        # 调试输出
        if "Important" in text or "重要提示" in text:
            print(f"\n[VERTICAL TEXT DEBUG] text='{text}'")
            print(f"[VERTICAL TEXT DEBUG] region.bbox={region.bbox}")
            print(f"[VERTICAL TEXT DEBUG] region.width={region.width}, region.height={region.height}")
            print(f"[VERTICAL TEXT DEBUG] height/width ratio={region.height/region.width:.2f}")
            print(f"[VERTICAL TEXT DEBUG] is_vertical_merged={getattr(region, 'is_vertical_merged', False)}")
            print(f"[VERTICAL TEXT DEBUG] is_vertical_text={is_vertical_text}\n")
        
        if is_vertical_text:
            # 对于竖排文本，需要扩展区域以容纳横排显示
            x1, y1, x2, y2 = region.bbox
            original_width = x2 - x1
            original_height = y2 - y1
            
            print(f"[EXPAND DEBUG] Entering vertical text expansion logic")
            print(f"[EXPAND DEBUG] original bbox=({x1},{y1})->({x2},{y2})")
            
            # 计算原始区域的中心点
            original_center_x = (x1 + x2) // 2
            original_center_y = (y1 + y2) // 2
            
            # 新的宽度：根据文本长度动态调整
            # 对于"Important Notice"这样的短文本，使用较小的宽度以触发按单词换行
            words = text.split()
            word_count = len(words)
            
            # 如果是多个单词，使用较小的宽度以触发换行
            # 宽度设置为能容纳最长单词的宽度（约为原高度的0.6倍）
            if word_count > 1:
                new_width = int(original_height * 0.6)  # 减小宽度以强制按单词换行
            else:
                new_width = int(original_height * 1.0)  # 单个单词不需要换行
            
            # 为竖排文本设置更小的字体大小
            # 将原始字体大小缩小到60%
            render_region_font_size = int(region.font_size * 0.6)
            
            # 新的高度 = 原高度（保持不变，足够容纳多行文本）
            new_height = int(original_height * 1.0)
            
            # X坐标：向左移动以避免与右边印章重叠
            # 使用原始区域的左边界，再向左移动一些（30px）
            adjusted_x1 = max(0, x1 - 30)
            adjusted_y1 = original_center_y - new_height // 2
            
            print(f"[EXPAND DEBUG] new_width={new_width}, new_height={new_height}")
            print(f"[EXPAND DEBUG] adjusted_x1={adjusted_x1}, adjusted_y1={adjusted_y1}")
            
            # 创建扩展后的区域（以中心点为基准）
            expanded_bbox = (adjusted_x1, adjusted_y1, adjusted_x1 + new_width, adjusted_y1 + new_height)
            
            render_region = TextRegion(
                bbox=expanded_bbox,
                text=region.text,
                confidence=region.confidence,
                font_size=render_region_font_size,  # 使用缩小后的字体大小
                angle=0.0,  # 横排显示，角度为0
                is_vertical_merged=True
            )
            
            # 保留 unified_content_left_boundary 属性（如果存在）
            if hasattr(region, 'unified_content_left_boundary'):
                render_region.unified_content_left_boundary = region.unified_content_left_boundary
            
            # 保留 unified_left_boundary 属性（如果存在）
            if hasattr(region, 'unified_left_boundary'):
                render_region.unified_left_boundary = region.unified_left_boundary
            
            # 保留 original_x1 属性（如果存在）
            if hasattr(region, 'original_x1'):
                render_region.original_x1 = region.original_x1
            
            # 保留 is_field_label 属性（如果存在）
            if hasattr(region, 'is_field_label'):
                render_region.is_field_label = region.is_field_label
            
            # 保留 is_field_content 属性（如果存在）
            if hasattr(region, 'is_field_content'):
                render_region.is_field_content = region.is_field_content
            
            # 保留 is_paragraph_merged 属性（如果存在）
            if hasattr(region, 'is_paragraph_merged'):
                render_region.is_paragraph_merged = region.is_paragraph_merged
            
            # 保留 belongs_to_field 属性（如果存在）
            if hasattr(region, 'belongs_to_field'):
                render_region.belongs_to_field = region.belongs_to_field
            
            # 保留 uses_unified_font_size 属性（如果存在）
            if hasattr(region, 'uses_unified_font_size'):
                render_region.uses_unified_font_size = region.uses_unified_font_size
            
            logger.info(
                f"Expanded region for vertical merged text: "
                f"original=({x1},{y1})->({x2},{y2}) [{original_width}x{original_height}], "
                f"original_center=({original_center_x},{original_center_y}), "
                f"expanded=({adjusted_x1},{adjusted_y1})->({adjusted_x1+new_width},{adjusted_y1+new_height}) [{new_width}x{new_height}], "
                f"font_size={region.font_size}->{render_region_font_size}, "
                f"text='{text}'"
            )
        
        # 检查是否是字段标签，如果是，限制其最大宽度
        is_field_label = hasattr(render_region, 'is_field_label') and render_region.is_field_label
        
        if is_field_label:
            # 字段标签：限制最大宽度为标准化字段框的宽度
            # 这样可以防止英文翻译超出标准化字段框，覆盖右侧的字段内容
            x1, y1, x2, y2 = render_region.bbox
            
            # 使用unified_left_boundary作为起始位置（如果存在）
            start_x = x1
            if hasattr(render_region, 'unified_left_boundary'):
                start_x = render_region.unified_left_boundary
            
            # 最大宽度 = 标准化字段框的右边界 - 起始位置
            # 这样可以确保字段标签不会超出标准化字段框
            max_label_width = x2 - start_x
            
            # 创建一个临时区域用于宽度限制
            limited_region = TextRegion(
                bbox=(start_x, y1, x2, y2),  # 使用原始的x2作为右边界
                text=render_region.text,
                confidence=render_region.confidence,
                font_size=render_region.font_size,
                angle=render_region.angle
            )
            
            # 保留所有属性
            if hasattr(render_region, 'unified_left_boundary'):
                limited_region.unified_left_boundary = render_region.unified_left_boundary
            if hasattr(render_region, 'unified_content_left_boundary'):
                limited_region.unified_content_left_boundary = render_region.unified_content_left_boundary
            if hasattr(render_region, 'original_x1'):
                limited_region.original_x1 = render_region.original_x1
            if hasattr(render_region, 'is_field_label'):
                limited_region.is_field_label = render_region.is_field_label
            if hasattr(render_region, 'is_field_content'):
                limited_region.is_field_content = render_region.is_field_content
            if hasattr(render_region, 'is_paragraph_merged'):
                limited_region.is_paragraph_merged = render_region.is_paragraph_merged
            if hasattr(render_region, 'belongs_to_field'):
                limited_region.belongs_to_field = render_region.belongs_to_field
            if hasattr(render_region, 'uses_unified_font_size'):
                limited_region.uses_unified_font_size = render_region.uses_unified_font_size
            
            render_region = limited_region
            
            logger.info(
                f"字段标签宽度限制: '{text}' 起始x={start_x}, 最大宽度={max_label_width}px, "
                f"bbox=({start_x},{y1},{x2},{y2})"
            )
        
        # Calculate font size
        # 对于超长文本（如"经营范围"内容），动态扩展区域高度
        original_region_height = render_region.height
        text_length = len(text)
        
        # 如果文本超长（>300字符），扩展区域高度以容纳更多行
        if text_length > 300:
            # 估算需要的行数：假设每行平均35个字符
            estimated_lines = (text_length // 35) + 1
            # 扩展高度：每行需要约15px（假设字体大小约12px）
            needed_height = estimated_lines * 15
            if needed_height > original_region_height:
                # 扩展区域高度
                x1, y1, x2, y2 = render_region.bbox
                new_y2 = y1 + needed_height
                
                # 创建新的 TextRegion，保留所有属性
                new_region = TextRegion(
                    bbox=(x1, y1, x2, new_y2),
                    text=render_region.text,
                    confidence=render_region.confidence,
                    font_size=render_region.font_size,
                    angle=render_region.angle
                )
                
                # 保留 unified_content_left_boundary 属性（如果存在）
                if hasattr(render_region, 'unified_content_left_boundary'):
                    new_region.unified_content_left_boundary = render_region.unified_content_left_boundary
                
                # 保留 unified_left_boundary 属性（如果存在）
                if hasattr(render_region, 'unified_left_boundary'):
                    new_region.unified_left_boundary = render_region.unified_left_boundary
                
                # 保留 original_x1 属性（如果存在）
                if hasattr(render_region, 'original_x1'):
                    new_region.original_x1 = render_region.original_x1
                
                # 保留 is_field_label 属性（如果存在）
                if hasattr(render_region, 'is_field_label'):
                    new_region.is_field_label = render_region.is_field_label
                
                # 保留 is_field_content 属性（如果存在）
                if hasattr(render_region, 'is_field_content'):
                    new_region.is_field_content = render_region.is_field_content
                
                # 保留 is_paragraph_merged 属性（如果存在）
                if hasattr(render_region, 'is_paragraph_merged'):
                    new_region.is_paragraph_merged = render_region.is_paragraph_merged
                
                # 保留 belongs_to_field 属性（如果存在）
                if hasattr(render_region, 'belongs_to_field'):
                    new_region.belongs_to_field = render_region.belongs_to_field
                
                # 保留 uses_unified_font_size 属性（如果存在）
                if hasattr(render_region, 'uses_unified_font_size'):
                    new_region.uses_unified_font_size = render_region.uses_unified_font_size
                
                render_region = new_region
                
                logger.info(
                    f"扩展超长文本区域: 原高度={original_region_height}px, "
                    f"新高度={needed_height}px, 文本长度={text_length}字符, "
                    f"估算行数={estimated_lines}行"
                )
        
        # 对于字段标签或字段内容，使用统一字体大小（不重新计算，避免因换行而缩小）
        if hasattr(render_region, 'uses_unified_font_size') and render_region.uses_unified_font_size:
            # 使用已经设置好的统一字体大小
            font_size = render_region.font_size
            content_type = "字段标签" if is_field_label else "字段内容"
            is_field_content = hasattr(render_region, 'is_field_content') and render_region.is_field_content
            
            # 强制输出调试信息
            print(f"\n{'='*80}")
            print(f"[字体大小追踪] {content_type}: '{text[:50]}...'")
            print(f"[字体大小追踪] region.font_size = {render_region.font_size}px")
            print(f"[字体大小追踪] uses_unified_font_size = True")
            print(f"[字体大小追踪] is_field_content = {is_field_content}")
            print(f"[字体大小追踪] 最终 font_size = {font_size}px (直接使用，不缩放)")
            print(f"{'='*80}\n")
            
            logger.info(
                f"{content_type}使用统一字体大小: '{text[:30]}...' font_size={font_size}px (不因换行而缩小)"
            )
        else:
            # 其他文本：正常计算字体大小
            print(f"\n[字体大小追踪] 非字段内容: '{text[:50]}...'")
            print(f"[字体大小追踪] 调用 calculate_font_size() 计算字体大小")
            font_size = self.calculate_font_size(text, render_region, max_lines=20)  # 增加到20行以容纳超长文本（如"经营范围"）
            print(f"[字体大小追踪] 计算结果 font_size = {font_size}px\n")
            
            # 对于竖排文本，缩小字体大小
            if is_vertical_text:
                original_font_size = font_size
                font_size = int(font_size * 0.5)  # 缩小到50%
                font_size = max(font_size, self.MIN_FONT_SIZE)  # 确保不小于最小字体
                print(f"[VERTICAL TEXT FONT] 竖排文本字体缩小: {original_font_size}px -> {font_size}px\n")
        
        # Determine if we need to wrap text
        # 对于垂直文本（竖排文本），强制按单词换行
        # 判断是否是竖排文本：
        # 1. 有is_vertical_merged标记
        # 2. 或者高度远大于宽度（height > width * 3)
        is_vertical_text = (
            (hasattr(region, 'is_vertical_merged') and region.is_vertical_merged) or
            (region.height > region.width * 3)
        )
        
        # 如果检测到印章重叠，使用更激进的换行策略
        if seal_overlap_detected:
            # 减小最大宽度，强制换行以避免遮挡印章
            # 使用原宽度的70%作为最大宽度
            adjusted_width = int(render_region.width * 0.7)
            wrapped_lines = self.wrap_text(text, adjusted_width, font_size, render_region)
            logger.info(
                f"印章重叠换行: '{text[:30]}...' -> {len(wrapped_lines)}行, "
                f"原宽度={render_region.width}px, 调整后宽度={adjusted_width}px"
            )
        elif is_vertical_text:
            # 按单词换行，使用wrap_text方法以确保每个单词都能放下
            # 这样可以避免单词被截断，同时保持按单词换行的效果
            wrapped_lines = self.wrap_text(text, render_region.width, font_size, render_region)
            
            # 如果wrap_text返回的行数少于单词数，说明某些单词被合并了
            # 这时强制每个单词一行（如果单词能放下的话）
            words = text.split()
            if len(wrapped_lines) < len(words) and len(words) > 1:
                # 检查每个单词是否能单独放下
                all_words_fit = all(
                    self.measure_text_width(word, font_size) <= render_region.width 
                    for word in words
                )
                if all_words_fit:
                    wrapped_lines = words  # 每个单词一行
                # 否则使用wrap_text的结果（某些长单词可能需要拆分）
        # 对于字段标签，强制换行（不允许超出边界）
        elif is_field_label:
            wrapped_lines = self.wrap_text(text, render_region.width, font_size, render_region)
            logger.info(
                f"字段标签换行: '{text}' -> {len(wrapped_lines)}行, "
                f"lines={[line[:20]+'...' if len(line)>20 else line for line in wrapped_lines]}"
            )
        else:
            wrapped_lines = self.wrap_text(text, render_region.width, font_size, render_region)
            
            # 特殊优化：如果换行结果是2行，检查是否可以通过允许超出边界来保持1行
            # 但是对于字段内容，不允许超出太多（最多10%），避免超出右边界
            # 对于底部文字，不允许超出边界
            if len(wrapped_lines) == 2:
                text_width = self.measure_text_width(text, font_size)
                
                # 判断是否是字段内容
                is_field_content = hasattr(render_region, 'is_field_content') and render_region.is_field_content
                
                # 判断是否是底部文字（y坐标 > 1200）
                is_bottom_text = render_region.bbox[1] > 1200
                
                # 字段内容：最多允许超出10%
                # 底部文字：不允许超出（0%）
                # 其他文本：允许超出20%
                if is_bottom_text:
                    max_overflow = 1.0  # 底部文字不允许超出
                elif is_field_content:
                    max_overflow = 1.10  # 字段内容最多超出10%
                else:
                    max_overflow = 1.20  # 其他文本允许超出20%
                
                # 允许超出指定比例
                if text_width <= render_region.width * max_overflow:
                    wrapped_lines = [text]
                    overflow_pct = (text_width / render_region.width - 1) * 100
                    print(f"[WRAP OPTIMIZATION RENDER] Text '{text[:30]}...' kept in 1 line with {overflow_pct:.1f}% overflow (width={text_width}, max={render_region.width}, is_field_content={is_field_content}, is_bottom={is_bottom_text})")
                else:
                    overflow_pct = (text_width / render_region.width - 1) * 100
                    print(f"[WRAP OPTIMIZATION RENDER] Text '{text[:30]}...' needs 2 lines (width={text_width}, max={render_region.width}, overflow={overflow_pct:.1f}%, is_field_content={is_field_content}, is_bottom={is_bottom_text})")
        
        # 强制输出调试信息（所有字段内容）
        if hasattr(render_region, 'unified_content_left_boundary'):
            print(f"[FIELD CONTENT DEBUG] text='{text[:50]}...', region_width={render_region.width}, font_size={font_size}")
            print(f"[FIELD CONTENT DEBUG] wrapped_lines={[line[:30]+'...' if len(line)>30 else line for line in wrapped_lines]}, line_count={len(wrapped_lines)}")
            print(f"[FIELD CONTENT DEBUG] region_bbox={render_region.bbox}, alignment={alignment}")
            print(f"[FIELD CONTENT DEBUG] unified_content_left_boundary={render_region.unified_content_left_boundary}")
        
        if hasattr(region, 'is_vertical_merged') and region.is_vertical_merged:
            print(f"[WRAP DEBUG] text='{text}', region_width={render_region.width}, font_size={font_size}")
            print(f"[WRAP DEBUG] wrapped_lines={wrapped_lines}, line_count={len(wrapped_lines)}")
        
        # 强制输出经营范围内容的渲染信息
        if 'Brand design' in text or 'advertising' in text or 'Business Scope' in text:
            print(f"[RENDER DEBUG] Rendering: '{text[:50]}...'")
            print(f"[RENDER DEBUG]   bbox: {render_region.bbox}")
            print(f"[RENDER DEBUG]   is_field_content: {getattr(render_region, 'is_field_content', False)}")
            print(f"[RENDER DEBUG]   is_paragraph_merged: {getattr(render_region, 'is_paragraph_merged', False)}")
            print(f"[RENDER DEBUG]   belongs_to_field: {getattr(render_region, 'belongs_to_field', None)}")
            print(f"[RENDER DEBUG]   font_size: {font_size}")
            print(f"[RENDER DEBUG]   lines: {len(wrapped_lines)}")
        
        logger.info(
            f"Rendering text: '{text[:30]}...', "
            f"region=({render_region.bbox[0]},{render_region.bbox[1]})->({render_region.bbox[2]},{render_region.bbox[3]}), "
            f"font_size={font_size}, "
            f"lines={len(wrapped_lines)}, "
            f"wrapped={wrapped_lines}"
        )
        
        # Calculate text color if not provided
        if text_color is None:
            text_color = self._calculate_contrast_color(background_color)
        
        # 统一使用横排渲染，忽略旋转角度
        result = self._render_straight_text(
            pil_image, render_region, wrapped_lines, font_size,
            text_color, alignment
        )
        
        # Convert back to numpy array (BGR)
        result_array = np.array(result)
        return result_array[:, :, ::-1]  # RGB to BGR
    
    def _render_straight_text(self, image: Image.Image, region: TextRegion,
                              lines: List[str], font_size: int,
                              text_color: Tuple[int, int, int],
                              alignment: TextAlignment) -> Image.Image:
        """Render non-rotated text onto an image.
        
        Args:
            image: PIL Image
            region: Target region
            lines: Text lines to render
            font_size: Font size
            text_color: Text color (RGB)
            alignment: Text alignment
            
        Returns:
            PIL Image with rendered text
        """
        result = image.copy()
        draw = ImageDraw.Draw(result)
        font = self._get_font(font_size)
        
        # 获取图片尺寸
        image_width, image_height = image.size
        
        # 获取边界限制配置
        boundary_enabled = self.config.get('rendering.boundary_limits.enabled', False)
        if boundary_enabled:
            right_margin = self.config.get('rendering.boundary_limits.right_margin', 0)
            left_margin = self.config.get('rendering.boundary_limits.left_margin', 0)
            top_margin = self.config.get('rendering.boundary_limits.top_margin', 0)
            bottom_margin = self.config.get('rendering.boundary_limits.bottom_margin', 0)
            
            # 计算有效渲染区域
            max_x = image_width - right_margin
            min_x = left_margin
            max_y = image_height - bottom_margin
            min_y = top_margin
            
            logger.debug(f"Boundary limits enabled: x=[{min_x}, {max_x}], y=[{min_y}, {max_y}]")
        else:
            # 没有边界限制,使用整个图片
            max_x = image_width
            min_x = 0
            max_y = image_height
            min_y = 0
        
        # Debug: Check region attributes before calling calculate_multiline_positions
        if '品牌设计' in ''.join(lines) or '布国内各类广告' in ''.join(lines) or 'Brand design' in ''.join(lines):
            print(f"[RENDER_TEXT DEBUG] Before calculate_multiline_positions:")
            print(f"[RENDER_TEXT DEBUG]   text: '{lines[0][:50]}...'")
            print(f"[RENDER_TEXT DEBUG]   is_field_content: {getattr(region, 'is_field_content', False)}")
            print(f"[RENDER_TEXT DEBUG]   is_paragraph_merged: {getattr(region, 'is_paragraph_merged', False)}")
            print(f"[RENDER_TEXT DEBUG]   belongs_to_field: {getattr(region, 'belongs_to_field', None)}")
        
        # Calculate positions for each line
        positions = self.calculate_multiline_positions(
            lines, region, font_size, alignment
        )
        
        # 判断是否是字段内容（需要顶部对齐）
        is_field_content = hasattr(region, 'unified_content_left_boundary')
        is_multiline = len(lines) > 1
        
        # Render each line
        for i, (line, (x, y)) in enumerate(zip(lines, positions)):
            # 检查文字是否会超出边界
            text_width = self.measure_text_width(line, font_size)
            text_right = x + text_width
            
            # 如果文字超出右边界,跳过渲染并记录警告
            if boundary_enabled and text_right > max_x:
                overflow = text_right - max_x
                logger.warning(
                    f"Text line '{line[:30]}...' would overflow right boundary by {overflow}px "
                    f"(x={x}, width={text_width}, max_x={max_x}). Skipping render."
                )
                print(f"[BOUNDARY WARNING] Skipped rendering '{line[:30]}...' - would overflow right boundary by {overflow}px")
                continue
            
            # 如果文字超出左边界,调整x坐标
            if boundary_enabled and x < min_x:
                logger.warning(f"Text line '{line[:30]}...' starts before left boundary (x={x}, min_x={min_x}). Adjusting.")
                x = min_x
            
            # 调试输出：显示每一行的实际渲染位置
            if is_field_content:
                print(f"[RENDER DEBUG] Line {i}: '{line[:30]}...' at x={x}, y={y}")
            
            # 注意：y坐标调整已经在calculate_multiline_positions中完成
            # 这里不需要再做调整，直接使用计算好的y坐标
            # 多行字段内容已经在calculate_multiline_positions中设置为顶部对齐
            # 单行字段内容已经在calculate_multiline_positions中设置为垂直居中
            
            if self.enable_stroke and self.stroke_width > 0:
                stroke_color = self._calculate_stroke_color(text_color)
                self._draw_text_with_stroke(
                    draw, (x, y), line, font, text_color, stroke_color
                )
            else:
                draw.text((x, y), line, font=font, fill=text_color)
        
        return result
    
    def _render_rotated_text(self, image: Image.Image, region: TextRegion,
                             lines: List[str], font_size: int,
                             text_color: Tuple[int, int, int],
                             alignment: TextAlignment) -> Image.Image:
        """Render rotated text onto an image.
        
        Creates a temporary image for the text, rotates it, and composites
        it onto the main image.
        
        Args:
            image: PIL Image
            region: Target region (with rotation angle)
            lines: Text lines to render
            font_size: Font size
            text_color: Text color (RGB)
            alignment: Text alignment
            
        Returns:
            PIL Image with rendered rotated text
            
        **Validates: Requirements 6.5**
        """
        result = image.copy()
        
        # Calculate text dimensions
        font = self._get_font(font_size)
        line_height = int(font_size * self.DEFAULT_LINE_SPACING_RATIO)
        
        # Find max line width
        max_width = max(self.measure_text_width(line, font_size) for line in lines)
        total_height = line_height * len(lines)
        
        # Create a larger canvas for rotation (to avoid clipping)
        padding = max(max_width, total_height)
        canvas_size = (max_width + padding * 2, total_height + padding * 2)
        
        # Create transparent text layer
        text_layer = Image.new('RGBA', canvas_size, (0, 0, 0, 0))
        text_draw = ImageDraw.Draw(text_layer)
        
        # Draw text centered on the canvas
        center_x = canvas_size[0] // 2
        start_y = padding
        
        for i, line in enumerate(lines):
            text_width = self.measure_text_width(line, font_size)
            
            if alignment == TextAlignment.LEFT:
                x = padding
            elif alignment == TextAlignment.RIGHT:
                x = canvas_size[0] - padding - text_width
            else:  # CENTER
                x = center_x - text_width // 2
            
            y = start_y + i * line_height
            
            if self.enable_stroke and self.stroke_width > 0:
                stroke_color = self._calculate_stroke_color(text_color)
                self._draw_text_with_stroke(
                    text_draw, (x, y), line, font, 
                    text_color + (255,), stroke_color + (255,)
                )
            else:
                text_draw.text((x, y), line, font=font, fill=text_color + (255,))
        
        # Rotate the text layer
        rotated = text_layer.rotate(
            -region.angle,  # PIL rotates counter-clockwise
            expand=True,
            resample=Image.BICUBIC if self.enable_antialiasing else Image.NEAREST
        )
        
        # Calculate paste position
        region_center = region.center
        paste_x = region_center[0] - rotated.width // 2
        paste_y = region_center[1] - rotated.height // 2
        
        # Composite onto result
        result.paste(rotated, (paste_x, paste_y), rotated)
        
        return result
    
    def _draw_text_with_stroke(self, draw: ImageDraw.ImageDraw,
                               position: Tuple[int, int], text: str,
                               font: ImageFont.FreeTypeFont,
                               fill_color: Tuple, stroke_color: Tuple) -> None:
        """Draw text with a stroke/outline effect.
        
        Args:
            draw: PIL ImageDraw object
            position: (x, y) position
            text: Text to draw
            font: Font to use
            fill_color: Fill color
            stroke_color: Stroke color
        """
        x, y = position
        
        # Draw stroke by drawing text multiple times offset
        for dx in range(-self.stroke_width, self.stroke_width + 1):
            for dy in range(-self.stroke_width, self.stroke_width + 1):
                if dx != 0 or dy != 0:
                    draw.text((x + dx, y + dy), text, font=font, fill=stroke_color)
        
        # Draw main text
        draw.text((x, y), text, font=font, fill=fill_color)
    
    def _calculate_contrast_color(self, background: Tuple[int, int, int]) -> Tuple[int, int, int]:
        """Calculate a contrasting text color for the given background.
        
        Uses luminance to determine if black or white text is more readable.
        
        Args:
            background: Background color (RGB)
            
        Returns:
            Contrasting text color (RGB)
        """
        r, g, b = background
        # Calculate relative luminance
        luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
        
        # Return black for light backgrounds, white for dark
        if luminance > 0.5:
            return (0, 0, 0)
        else:
            return (255, 255, 255)
    
    def _calculate_stroke_color(self, text_color: Tuple[int, int, int]) -> Tuple[int, int, int]:
        """Calculate a stroke color that contrasts with the text color.
        
        Args:
            text_color: Text color (RGB)
            
        Returns:
            Stroke color (RGB)
        """
        r, g, b = text_color
        luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
        
        if luminance > 0.5:
            return (0, 0, 0)
        else:
            return (255, 255, 255)
    
    def apply_antialiasing(self, image: np.ndarray) -> np.ndarray:
        """Apply anti-aliasing effect to an image.
        
        Uses a slight blur to smooth jagged edges.
        
        Args:
            image: Input image (BGR format)
            
        Returns:
            Anti-aliased image (BGR format)
            
        **Validates: Requirements 5.1**
        """
        if not self.enable_antialiasing:
            return image.copy()
        
        # Convert to PIL for processing
        image_rgb = image[:, :, ::-1]
        pil_image = Image.fromarray(image_rgb)
        
        # Apply slight smoothing
        smoothed = pil_image.filter(ImageFilter.SMOOTH)
        
        # Convert back
        result = np.array(smoothed)
        return result[:, :, ::-1]

    def render_text_with_shadow(self, image: np.ndarray, region: TextRegion, 
                                text: str, background_color: Tuple[int, int, int],
                                text_color: Optional[Tuple[int, int, int]] = None,
                                shadow_color: Optional[Tuple[int, int, int]] = None,
                                shadow_offset: Tuple[int, int] = (2, 2),
                                alignment: TextAlignment = TextAlignment.CENTER) -> np.ndarray:
        """Render text with a shadow effect.
        
        Args:
            image: Input image (BGR format)
            region: Target text region
            text: Text to render
            background_color: Background color (RGB)
            text_color: Text color (RGB), auto-calculated if None
            shadow_color: Shadow color (RGB), auto-calculated if None
            shadow_offset: Shadow offset (x, y) in pixels
            alignment: Text alignment
            
        Returns:
            Image with rendered text and shadow (BGR format)
            
        **Validates: Requirements 5.5**
        """
        if not text:
            return image.copy()
        
        # Convert to PIL
        image_rgb = image[:, :, ::-1]
        pil_image = Image.fromarray(image_rgb)
        
        # Calculate font size
        font_size = self.calculate_font_size(text, region, max_lines=3)
        wrapped_lines = self.wrap_text(text, region.width, font_size, region)
        
        # Calculate colors
        if text_color is None:
            text_color = self._calculate_contrast_color(background_color)
        
        if shadow_color is None:
            # Shadow is typically darker version of background or opposite of text
            shadow_color = self._calculate_shadow_color(text_color, background_color)
        
        # Get font and positions
        font = self._get_font(font_size)
        positions = self.calculate_multiline_positions(
            wrapped_lines, region, font_size, alignment
        )
        
        draw = ImageDraw.Draw(pil_image)
        
        # Draw shadow first (offset)
        for line, (x, y) in zip(wrapped_lines, positions):
            shadow_x = x + shadow_offset[0]
            shadow_y = y + shadow_offset[1]
            draw.text((shadow_x, shadow_y), line, font=font, fill=shadow_color)
        
        # Draw main text
        for line, (x, y) in zip(wrapped_lines, positions):
            if self.enable_stroke and self.stroke_width > 0:
                stroke_color = self._calculate_stroke_color(text_color)
                self._draw_text_with_stroke(
                    draw, (x, y), line, font, text_color, stroke_color
                )
            else:
                draw.text((x, y), line, font=font, fill=text_color)
        
        # Convert back to BGR
        result = np.array(pil_image)
        return result[:, :, ::-1]
    
    def _calculate_shadow_color(self, text_color: Tuple[int, int, int],
                                background_color: Tuple[int, int, int]) -> Tuple[int, int, int]:
        """Calculate an appropriate shadow color.
        
        Args:
            text_color: Text color (RGB)
            background_color: Background color (RGB)
            
        Returns:
            Shadow color (RGB)
        """
        # Shadow should be darker than background but visible
        r, g, b = background_color
        
        # Darken the background color
        shadow_r = max(0, r - 50)
        shadow_g = max(0, g - 50)
        shadow_b = max(0, b - 50)
        
        return (shadow_r, shadow_g, shadow_b)
    
    def set_font(self, font_name: str) -> bool:
        """Set the font to use for rendering.
        
        Args:
            font_name: Name of the font to use
            
        Returns:
            True if font was found and set, False otherwise
            
        **Validates: Requirements 5.6**
        """
        # Clear font cache
        self._font_cache.clear()
        
        # Try to find the font
        if font_name in self.FONT_PATHS:
            for path in self.FONT_PATHS[font_name]:
                if Path(path).exists():
                    self._font_path = path
                    self.font_family = font_name
                    logger.info(f"Font set to: {font_name} ({path})")
                    return True
        
        # Try to find font by searching common paths
        common_paths = [
            'C:/Windows/Fonts',
            '/usr/share/fonts/truetype',
            '/usr/share/fonts',
            '/System/Library/Fonts',
        ]
        
        font_name_lower = font_name.lower()
        for base_path in common_paths:
            base = Path(base_path)
            if base.exists():
                for font_file in base.rglob('*.ttf'):
                    if font_name_lower in font_file.stem.lower():
                        self._font_path = str(font_file)
                        self.font_family = font_name
                        logger.info(f"Font set to: {font_name} ({font_file})")
                        return True
        
        logger.warning(f"Font not found: {font_name}")
        return False
    
    def get_available_fonts(self) -> List[str]:
        """Get a list of available fonts.
        
        Returns:
            List of available font names
            
        **Validates: Requirements 5.6**
        """
        available = []
        
        for font_name, paths in self.FONT_PATHS.items():
            for path in paths:
                if Path(path).exists():
                    available.append(font_name)
                    break
        
        return available
    
    def render_text_with_effects(self, image: np.ndarray, region: TextRegion,
                                 text: str, background_color: Tuple[int, int, int],
                                 text_color: Optional[Tuple[int, int, int]] = None,
                                 enable_shadow: bool = False,
                                 shadow_offset: Tuple[int, int] = (2, 2),
                                 enable_stroke: Optional[bool] = None,
                                 stroke_width: Optional[int] = None,
                                 alignment: TextAlignment = TextAlignment.CENTER) -> np.ndarray:
        """Render text with configurable effects.
        
        A unified method for rendering text with various effects.
        
        Args:
            image: Input image (BGR format)
            region: Target text region
            text: Text to render
            background_color: Background color (RGB)
            text_color: Text color (RGB), auto-calculated if None
            enable_shadow: Whether to add shadow effect
            shadow_offset: Shadow offset (x, y) in pixels
            enable_stroke: Whether to add stroke effect (overrides config)
            stroke_width: Stroke width (overrides config)
            alignment: Text alignment
            
        Returns:
            Image with rendered text (BGR format)
            
        **Validates: Requirements 5.5, 5.6**
        """
        # Save original settings
        original_stroke = self.enable_stroke
        original_stroke_width = self.stroke_width
        
        try:
            # Apply overrides
            if enable_stroke is not None:
                self.enable_stroke = enable_stroke
            if stroke_width is not None:
                self.stroke_width = stroke_width
            
            if enable_shadow:
                return self.render_text_with_shadow(
                    image, region, text, background_color,
                    text_color=text_color,
                    shadow_offset=shadow_offset,
                    alignment=alignment
                )
            else:
                return self.render_text(
                    image, region, text, background_color,
                    text_color=text_color,
                    alignment=alignment
                )
        finally:
            # Restore original settings
            self.enable_stroke = original_stroke
            self.stroke_width = original_stroke_width
    
    def get_line_spacing(self, font_size: int, 
                         line_spacing_ratio: Optional[float] = None) -> int:
        """Calculate line spacing for a given font size.
        
        Args:
            font_size: Font size in pixels
            line_spacing_ratio: Line spacing ratio (default: 1.2)
            
        Returns:
            Line spacing in pixels
            
        **Validates: Requirements 6.3**
        """
        if line_spacing_ratio is None:
            line_spacing_ratio = self.DEFAULT_LINE_SPACING_RATIO
        
        return int(font_size * line_spacing_ratio)
