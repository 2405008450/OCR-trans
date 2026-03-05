"""Core data models for image translation system.

This module defines the fundamental data structures used throughout the
image translation pipeline, including TextRegion, TranslationResult, and QualityReport.

核心数据模型模块

本模块定义了图像翻译系统中使用的基础数据结构，包括文本区域、翻译结果、质量报告和配置相关的数据模型。
"""

from dataclasses import dataclass, field
from typing import Tuple, Optional, List, Dict, Any
from enum import Enum


@dataclass
class TextRegion:
    """Text region detected in an image.
    
    Represents a detected text area with its bounding box, content,
    and associated metadata like confidence score and font information.
    
    Attributes:
        bbox: Bounding box coordinates as (x1, y1, x2, y2)
        text: The detected text content
        confidence: OCR confidence score (0.0 to 1.0)
        font_size: Estimated font size in pixels
        angle: Rotation angle in degrees (-180 to 180)
        is_vertical_merged: Whether this is vertically merged text (e.g., "重要提示")
        is_field_label: Whether this is a field label (e.g., "名称", "类型")
        is_field_content: Whether this is field content (already paired with a label)
        field_label_name: Name of the field label if this is a field label
        belongs_to_field: Name of the field this region belongs to
        is_paragraph_merged: Whether this is a result of paragraph merging
        original_regions: List of original regions that were merged to create this region
        merge_strategy: The merge strategy used to create this region
    """
    bbox: Tuple[int, int, int, int]
    text: str
    confidence: float
    font_size: int
    angle: float = 0.0
    is_vertical_merged: bool = False  # 标记是否是垂直合并的文本（如"重要提示"）
    is_field_label: bool = False  # 标记是否是字段标签（如"名称"、"类型"等）
    is_field_content: bool = False  # 标记是否是字段内容（已配对的内容）
    field_label_name: Optional[str] = None  # 字段标签名称
    belongs_to_field: Optional[str] = None  # 属于哪个字段
    is_paragraph_merged: bool = False  # 是否是段落合并的结果
    original_regions: Optional[List['TextRegion']] = None  # 原始区域列表
    merge_strategy: Optional[str] = None  # 使用的合并策略
    
    def __post_init__(self):
        """Validate field values after initialization."""
        if not isinstance(self.bbox, tuple) or len(self.bbox) != 4:
            raise ValueError("bbox must be a tuple of 4 integers (x1, y1, x2, y2)")
        
        x1, y1, x2, y2 = self.bbox
        if not all(isinstance(v, (int, float)) for v in self.bbox):
            raise ValueError("bbox coordinates must be numeric")
        
        if not isinstance(self.confidence, (int, float)):
            raise ValueError("confidence must be a number")
        if not 0.0 <= self.confidence <= 1.0:
            raise ValueError("confidence must be between 0.0 and 1.0")
        
        if not isinstance(self.font_size, (int, float)) or self.font_size < 0:
            raise ValueError("font_size must be a non-negative number")
        
        if not isinstance(self.angle, (int, float)):
            raise ValueError("angle must be a number")
    
    @property
    def center(self) -> Tuple[int, int]:
        """Return the center point of the text region.
        
        Returns:
            Tuple of (x, y) coordinates representing the center point
        """
        x1, y1, x2, y2 = self.bbox
        return ((x1 + x2) // 2, (y1 + y2) // 2)
    
    @property
    def area(self) -> int:
        """Return the area of the text region in pixels.
        
        Returns:
            Area in square pixels
        """
        x1, y1, x2, y2 = self.bbox
        width = abs(x2 - x1)
        height = abs(y2 - y1)
        return width * height
    
    @property
    def aspect_ratio(self) -> float:
        """Return the aspect ratio of the text region.
        
        The aspect ratio is calculated as min(width, height) / max(width, height),
        resulting in a value between 0 and 1. A value close to 1 indicates
        a square-like region.
        
        Returns:
            Aspect ratio between 0 and 1
        """
        x1, y1, x2, y2 = self.bbox
        width = abs(x2 - x1)
        height = abs(y2 - y1)
        if max(width, height) == 0:
            return 0.0
        return min(width, height) / max(width, height)
    
    @property
    def width(self) -> int:
        """Return the width of the text region."""
        x1, _, x2, _ = self.bbox
        return abs(x2 - x1)
    
    @property
    def height(self) -> int:
        """Return the height of the text region."""
        _, y1, _, y2 = self.bbox
        return abs(y2 - y1)


@dataclass
class TranslationResult:
    """Result of a translation operation.
    
    Contains the source text, translated text, and metadata about
    the translation quality and success status.
    
    Attributes:
        source_text: Original text before translation
        translated_text: Translated text (empty string if failed)
        confidence: Translation confidence score (0.0 to 1.0)
        success: Whether the translation was successful
        error_message: Error description if translation failed
    """
    source_text: str
    translated_text: str
    confidence: float
    success: bool
    error_message: Optional[str] = None
    
    def __post_init__(self):
        """Validate field values after initialization."""
        if not isinstance(self.source_text, str):
            raise ValueError("source_text must be a string")
        
        if not isinstance(self.translated_text, str):
            raise ValueError("translated_text must be a string")
        
        if not isinstance(self.confidence, (int, float)):
            raise ValueError("confidence must be a number")
        if not 0.0 <= self.confidence <= 1.0:
            raise ValueError("confidence must be between 0.0 and 1.0")
        
        if not isinstance(self.success, bool):
            raise ValueError("success must be a boolean")
        
        if self.error_message is not None and not isinstance(self.error_message, str):
            raise ValueError("error_message must be a string or None")


class QualityLevel(Enum):
    """Quality level enumeration for translation output.
    
    Represents the overall quality assessment of a translated image.
    """
    EXCELLENT = "excellent"
    GOOD = "good"
    FAIR = "fair"
    POOR = "poor"


@dataclass
class QualityReport:
    """Quality report for a translation task.
    
    Contains comprehensive quality metrics and assessment for
    a completed translation operation.
    
    Attributes:
        translation_coverage: Ratio of successfully translated regions (0.0 to 1.0)
        total_regions: Total number of text regions detected
        translated_regions: Number of successfully translated regions
        failed_regions: List of TextRegion objects that failed translation
        has_artifacts: Whether visual artifacts were detected
        artifact_locations: List of bounding boxes where artifacts were found
        overall_quality: Overall quality assessment level
    """
    translation_coverage: float
    total_regions: int
    translated_regions: int
    failed_regions: List[TextRegion]
    has_artifacts: bool
    artifact_locations: List[Tuple[int, int, int, int]]
    overall_quality: QualityLevel
    
    def __post_init__(self):
        """Validate field values after initialization."""
        if not isinstance(self.translation_coverage, (int, float)):
            raise ValueError("translation_coverage must be a number")
        if not 0.0 <= self.translation_coverage <= 1.0:
            raise ValueError("translation_coverage must be between 0.0 and 1.0")
        
        if not isinstance(self.total_regions, int) or self.total_regions < 0:
            raise ValueError("total_regions must be a non-negative integer")
        
        if not isinstance(self.translated_regions, int) or self.translated_regions < 0:
            raise ValueError("translated_regions must be a non-negative integer")
        
        if self.translated_regions > self.total_regions:
            raise ValueError("translated_regions cannot exceed total_regions")
        
        if not isinstance(self.failed_regions, list):
            raise ValueError("failed_regions must be a list")
        
        if not isinstance(self.has_artifacts, bool):
            raise ValueError("has_artifacts must be a boolean")
        
        if not isinstance(self.artifact_locations, list):
            raise ValueError("artifact_locations must be a list")
        
        if not isinstance(self.overall_quality, QualityLevel):
            raise ValueError("overall_quality must be a QualityLevel enum value")
    
    def to_dict(self) -> dict:
        """Convert the quality report to a dictionary.
        
        Returns:
            Dictionary representation of the quality report
        """
        return {
            "translation_coverage": self.translation_coverage,
            "total_regions": self.total_regions,
            "translated_regions": self.translated_regions,
            "failed_regions": len(self.failed_regions),
            "has_artifacts": self.has_artifacts,
            "artifact_count": len(self.artifact_locations),
            "overall_quality": self.overall_quality.value
        }


# ============================================================================
# 配置相关的数据模型 | Configuration-related Data Models
# ============================================================================

class DocumentOrientation(Enum):
    """文档方向枚举
    
    定义文档的方向类型，用于选择不同的处理策略。
    
    Document orientation enumeration.
    
    Defines the orientation type of documents for selecting different processing strategies.
    
    Attributes:
        VERTICAL: 竖版文档 | Vertical document orientation
        HORIZONTAL: 横版文档 | Horizontal document orientation
    """
    VERTICAL = "vertical"
    HORIZONTAL = "horizontal"


@dataclass
class ConfigBlock:
    """配置区块数据类
    
    表示一个完整的配置区块，包含所有模块的配置参数。
    
    Configuration block data class.
    
    Represents a complete configuration block containing parameters for all modules.
    
    Attributes:
        ocr: OCR 相关配置 | OCR-related configuration
        rendering: 渲染相关配置 | Rendering-related configuration
        image: 图像处理相关配置 | Image processing-related configuration
        translation: 翻译相关配置 | Translation-related configuration
    """
    ocr: Dict[str, Any] = field(default_factory=dict)
    rendering: Dict[str, Any] = field(default_factory=dict)
    image: Dict[str, Any] = field(default_factory=dict)
    translation: Dict[str, Any] = field(default_factory=dict)
    
    def __post_init__(self):
        """验证字段值 | Validate field values after initialization."""
        if not isinstance(self.ocr, dict):
            raise ValueError("ocr must be a dictionary")
        if not isinstance(self.rendering, dict):
            raise ValueError("rendering must be a dictionary")
        if not isinstance(self.image, dict):
            raise ValueError("image must be a dictionary")
        if not isinstance(self.translation, dict):
            raise ValueError("translation must be a dictionary")


@dataclass
class DualConfig:
    """双配置结构数据类
    
    表示完整的双配置架构，包含竖版和横版两套独立的配置。
    
    Dual configuration structure data class.
    
    Represents the complete dual-configuration architecture with separate
    configurations for vertical and horizontal orientations.
    
    Attributes:
        orientation: 当前文档方向 | Current document orientation
        vertical_config: 竖版配置区块 | Vertical configuration block
        horizontal_config: 横版配置区块 | Horizontal configuration block
        global_config: 全局配置（两种方向共享）| Global configuration (shared by both orientations)
    """
    orientation: DocumentOrientation
    vertical_config: ConfigBlock
    horizontal_config: ConfigBlock
    global_config: Dict[str, Any] = field(default_factory=dict)
    
    def __post_init__(self):
        """验证字段值 | Validate field values after initialization."""
        if not isinstance(self.orientation, DocumentOrientation):
            raise ValueError("orientation must be a DocumentOrientation enum value")
        if not isinstance(self.vertical_config, ConfigBlock):
            raise ValueError("vertical_config must be a ConfigBlock instance")
        if not isinstance(self.horizontal_config, ConfigBlock):
            raise ValueError("horizontal_config must be a ConfigBlock instance")
        if not isinstance(self.global_config, dict):
            raise ValueError("global_config must be a dictionary")
    
    def get_active_config(self) -> ConfigBlock:
        """获取当前激活的配置区块
        
        根据当前的 orientation 返回对应的配置区块。
        
        Get the currently active configuration block.
        
        Returns the configuration block corresponding to the current orientation.
        
        Returns:
            当前激活的配置区块 | Currently active configuration block
        """
        if self.orientation == DocumentOrientation.VERTICAL:
            return self.vertical_config
        else:
            return self.horizontal_config


@dataclass
class MergeConfig:
    """段落合并配置数据类
    
    包含段落合并功能的所有配置参数，支持字段感知的合并策略。
    
    Paragraph merge configuration data class.
    
    Contains all configuration parameters for paragraph merging functionality,
    supporting field-aware merge strategies.
    
    Attributes:
        enabled: 是否启用段落合并 | Whether to enable paragraph merging
        x_tolerance: 水平对齐容差（像素）| Horizontal alignment tolerance (pixels)
        y_gap_max: 垂直间距阈值（像素）| Vertical gap threshold (pixels)
        font_size_diff_threshold: 字体大小差异阈值（比例）| Font size difference threshold (ratio)
        field_aware_merge_enabled: 是否启用字段感知合并 | Whether to enable field-aware merging
        known_field_labels: 已知字段标签列表 | List of known field labels
        long_text_fields: 长文本字段列表 | List of long text fields
        long_text_field_y_gap_max: 长文本字段的垂直间距阈值 | Vertical gap threshold for long text fields
        log_merge_decisions: 是否记录合并决策 | Whether to log merge decisions
        log_level: 日志级别 | Log level
    """
    enabled: bool = True
    x_tolerance: int = 10
    y_gap_max: int = 5
    font_size_diff_threshold: float = 0.2
    field_aware_merge_enabled: bool = True
    known_field_labels: List[str] = field(default_factory=lambda: [
        "经营范围", "住所", "法定代表人", "注册资本", "成立日期", "营业期限"
    ])
    long_text_fields: List[str] = field(default_factory=lambda: [
        "经营范围", "住所"
    ])
    long_text_field_y_gap_max: int = 10
    log_merge_decisions: bool = True
    log_level: str = "INFO"
    
    def __post_init__(self):
        """验证字段值 | Validate field values after initialization."""
        if not isinstance(self.enabled, bool):
            raise ValueError("enabled must be a boolean")
        
        if not isinstance(self.x_tolerance, int) or self.x_tolerance < 0:
            self.x_tolerance = 10  # Use default value
        
        if not isinstance(self.y_gap_max, int) or self.y_gap_max < 0:
            self.y_gap_max = 5  # Use default value
        
        if not isinstance(self.font_size_diff_threshold, (int, float)) or not 0.0 <= self.font_size_diff_threshold <= 1.0:
            self.font_size_diff_threshold = 0.2  # Use default value
        
        if not isinstance(self.field_aware_merge_enabled, bool):
            raise ValueError("field_aware_merge_enabled must be a boolean")
        
        if not isinstance(self.known_field_labels, list):
            self.known_field_labels = ["经营范围", "住所", "法定代表人", "注册资本", "成立日期", "营业期限"]
        
        if not isinstance(self.long_text_fields, list):
            self.long_text_fields = ["经营范围", "住所"]
        
        if not isinstance(self.long_text_field_y_gap_max, int) or self.long_text_field_y_gap_max < 0:
            self.long_text_field_y_gap_max = 10  # Use default value
        
        if not isinstance(self.log_merge_decisions, bool):
            raise ValueError("log_merge_decisions must be a boolean")
        
        if not isinstance(self.log_level, str):
            self.log_level = "INFO"  # Use default value
