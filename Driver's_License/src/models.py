"""数据模型定义"""

from dataclasses import dataclass
from typing import List, Tuple, Optional
import numpy as np


@dataclass
class TextBlock:
    """表示 OCR 识别到的单个文字块"""
    text: str                    # 文字内容
    bounding_box: List[Tuple[int, int]]  # 边界框坐标 [(x1,y1), (x2,y2), (x3,y3), (x4,y4)]
    confidence: float            # 识别置信度 (0-1)
    
    def get_rect(self) -> Tuple[int, int, int, int]:
        """获取矩形边界框 (x, y, width, height)"""
        xs = [p[0] for p in self.bounding_box]
        ys = [p[1] for p in self.bounding_box]
        x, y = min(xs), min(ys)
        width = max(xs) - x
        height = max(ys) - y
        return (x, y, width, height)
    
    def get_center(self) -> Tuple[int, int]:
        """获取中心点坐标"""
        x, y, width, height = self.get_rect()
        return (x + width // 2, y + height // 2)


@dataclass
class LicenseField:
    """表示驾驶证的一个字段"""
    field_name: str              # 字段名称（如 "姓名"、"性别"）
    field_value: str             # 字段值
    position: Tuple[int, int]    # 字段值的位置坐标
    translated_value: str = ""   # 翻译后的值


@dataclass
class ExtractedImage:
    """表示提取的照片或印章"""
    image_type: str              # 图像类型："photo" 或 "seal"
    image_data: np.ndarray       # 图像数据
    position: Tuple[int, int]    # 位置坐标 (x, y)
    size: Tuple[int, int]        # 尺寸 (width, height)
    temp_path: str = ""          # 临时文件路径


@dataclass
class LicenseData:
    """表示完整的驾驶证数据"""
    fields: List[LicenseField]   # 所有字段
    images: List[ExtractedImage] # 提取的图像
    image_size: Tuple[int, int]  # 原图尺寸 (width, height)
    text_blocks: List[TextBlock] # 所有文字块（用于布局参考）
    seal_texts: List[str] = None # 印章区域的文字（已翻译）
    has_duplicate: bool = False  # 是否有副页
    barcode_number: str = None   # 条形码数字（准驾车型代号规定页面）
    is_old_version: bool = False # 是否是旧版驾驶证（有 Valid From 字段）
    has_main_page: bool = True   # 是否有主页
    has_legend_page: bool = False # 是否有准驾车型代号规定页
    
    def __post_init__(self):
        if self.seal_texts is None:
            self.seal_texts = []
