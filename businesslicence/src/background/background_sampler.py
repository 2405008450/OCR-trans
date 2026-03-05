"""Background sampler for image translation system.

This module provides the BackgroundSampler class for analyzing and extracting
background colors from text regions, performing inpainting to fill text areas,
and applying feathering effects for smooth edge transitions.
"""

import logging
from typing import Tuple, List, Optional
import numpy as np
import cv2

from src.config import ConfigManager
from src.models import TextRegion


logger = logging.getLogger(__name__)


class BackgroundSampler:
    """Samples and processes background colors for text regions.
    
    The BackgroundSampler provides functionality to:
    - Sample background colors around text regions
    - Use median filtering to calculate representative colors
    - Apply inpainting to fill text regions
    - Apply feathering effects for smooth edge transitions
    
    Attributes:
        config: Configuration manager instance
        sample_radius: Radius for background color sampling
        use_inpainting: Whether to use inpainting for background restoration
    """
    
    def __init__(self, config: ConfigManager):
        """初始化背景采样器。
        
        参数：
            config: 配置管理器实例
        """
        self.config = config
        self.sample_radius = config.get('rendering.background_sample_radius', 10)  # 增加到10
        self.sample_points = config.get('rendering.background_sample_points', 20)  # 增加到20个点
        self.use_iqr_filter = config.get('rendering.background_use_iqr_filter', True)  # 启用IQR过滤
        self.use_inpainting = config.get('rendering.use_inpainting', True)
        
        # Padding配置
        self.padding_mode = config.get('rendering.background_padding', 'auto')
        self.padding_min = config.get('rendering.background_padding_min', 10)
        self.padding_max = config.get('rendering.background_padding_max', 30)
        self.padding_ratio = config.get('rendering.background_padding_ratio', 0.6)
        
        # 羽化配置
        self.feather_mode = config.get('rendering.feather_radius', 'auto')
        self.feather_min = config.get('rendering.feather_radius_min', 3)
        self.feather_max = config.get('rendering.feather_radius_max', 10)
        
        logger.info(f"BackgroundSampler 已初始化: sample_radius={self.sample_radius}, sample_points={self.sample_points}, padding_mode={self.padding_mode}")
    
    def calculate_padding(self, region: TextRegion) -> int:
        """计算智能padding值。
        
        改进策略：
        1. 确保padding足够大以完全覆盖原文
        2. 基于字体大小动态调整
        3. 对于长文本使用更大的padding
        
        参数：
            region: 文本区域
        
        返回：
            padding值（像素）
        """
        # 如果不是auto模式，返回固定值
        if self.padding_mode != 'auto':
            try:
                return int(self.padding_mode)
            except (ValueError, TypeError):
                logger.warning(f"Invalid padding_mode '{self.padding_mode}', using auto")
        
        font_size = region.font_size
        area = region.area
        width = region.width
        height = region.height
        
        # 基于字体大小的padding（确保完全覆盖原文）
        if font_size < 15:
            base_padding = 8  # 小字体：8px padding
        elif font_size < 25:
            base_padding = font_size * 0.5  # 中等字体：50%
        else:
            base_padding = font_size * 0.4  # 大字体：40%
        
        # 对于长文本（如"经营范围"），使用更大的padding
        if width > 400 or height > 60:
            base_padding = max(base_padding, 15)  # 至少15px
        
        # 基于面积的调整因子
        if area < 1000:
            area_factor = 1.0  # 小区域：不调整
        elif area < 5000:
            area_factor = 1.1  # 中等区域：略微增加
        else:
            area_factor = 1.2  # 大区域：增加
        
        padding = int(base_padding * area_factor)
        
        # 限制在配置的最小值和最大值之间
        padding = max(self.padding_min, min(self.padding_max, padding))
        
        logger.debug(f"智能Padding: font_size={font_size}, area={area}, width={width}, height={height}, padding={padding}px")
        return padding
    
    def sample_background(self, image: np.ndarray, region: TextRegion) -> Tuple[int, int, int]:
        """Sample the background color around a text region.
        
        Samples multiple pixels around the text region boundary and uses
        median filtering to calculate a representative background color.
        
        Args:
            image: Input image as numpy array (BGR format)
            region: Text region to sample around
            
        Returns:
            RGB color tuple (R, G, B) with values in range 0-255
        """
        # 使用改进的采样方法
        return self.sample_background_improved(image, region)
    
    def sample_background_improved(self, image: np.ndarray, region: TextRegion) -> Tuple[int, int, int]:
        """改进的背景采样方法。
        
        使用更多采样点、更大半径和IQR过滤来获取更准确的背景色。
        
        参数：
            image: 输入图片（BGR格式）
            region: 文本区域
        
        返回：
            RGB颜色元组 (R, G, B)，值在0-255范围内
        """
        height, width = image.shape[:2]
        
        # 使用密集采样点
        sample_points = self._get_sample_points_dense(
            region,
            self.sample_radius,
            self.sample_points // 4  # 每条边采样点数
        )
        
        if not sample_points:
            logger.warning(f"No valid sample points for region {region.bbox}, using white")
            return (255, 255, 255)
        
        # 采样颜色值
        pixel_values = []
        for px, py in sample_points:
            if self._is_valid_point(px, py, image):
                # OpenCV使用BGR格式
                bgr = image[py, px]
                pixel_values.append(bgr)
        
        if not pixel_values:
            logger.warning(f"No valid pixels sampled for region {region.bbox}, using white")
            return (255, 255, 255)
        
        # 使用IQR过滤异常值
        if self.use_iqr_filter and len(pixel_values) >= 4:
            pixel_values = self._filter_outliers_iqr(pixel_values)
        
        # 使用中值滤波获取代表性颜色
        pixel_array = np.array(pixel_values)
        median_color = np.median(pixel_array, axis=0).astype(int)
        
        # 转换BGR到RGB
        r, g, b = int(median_color[2]), int(median_color[1]), int(median_color[0])
        
        # 确保值在有效范围内
        r = max(0, min(255, r))
        g = max(0, min(255, g))
        b = max(0, min(255, b))
        
        logger.info(f"🎨 采样背景色 RGB({r}, {g}, {b})，使用{len(pixel_values)}个采样点")
        return (r, g, b)
    
    def calculate_feather_radius(self, region: TextRegion) -> int:
        """计算自适应羽化半径。
        
        根据区域面积动态计算最佳羽化半径。
        
        参数：
            region: 文本区域
        
        返回：
            羽化半径（像素）
        """
        # 如果不是auto模式，返回固定值
        if self.feather_mode != 'auto':
            try:
                return int(self.feather_mode)
            except (ValueError, TypeError):
                logger.warning(f"Invalid feather_mode '{self.feather_mode}', using auto")
        
        area = region.area
        
        # 根据区域面积计算羽化半径（增加半径以获得更平滑的过渡）
        if area < 1000:
            radius = 8  # 小区域用较大半径
        elif area < 5000:
            radius = 10  # 中等区域用大半径
        else:
            radius = 12  # 大区域用更大半径
        
        # 限制在配置的最小值和最大值之间
        radius = max(self.feather_min, min(self.feather_max, radius))
        
        logger.info(f"🎨 羽化半径计算: area={area}, radius={radius}")
        return radius
    
    def _get_sample_points(
        self, 
        x1: int, 
        y1: int, 
        x2: int, 
        y2: int, 
        img_width: int, 
        img_height: int
    ) -> List[Tuple[int, int]]:
        """Get sample points around a bounding box.
        
        Samples points along the boundary of the region, offset by the
        sample radius to avoid sampling text pixels.
        
        Args:
            x1, y1, x2, y2: Bounding box coordinates
            img_width, img_height: Image dimensions
            
        Returns:
            List of (x, y) sample point coordinates
        """
        sample_points = []
        radius = self.sample_radius
        
        # Sample along top edge (above the region)
        for x in range(x1, x2, max(1, (x2 - x1) // 10)):
            py = y1 - radius
            if 0 <= py < img_height and 0 <= x < img_width:
                sample_points.append((x, py))
        
        # Sample along bottom edge (below the region)
        for x in range(x1, x2, max(1, (x2 - x1) // 10)):
            py = y2 + radius
            if 0 <= py < img_height and 0 <= x < img_width:
                sample_points.append((x, py))
        
        # Sample along left edge (left of the region)
        for y in range(y1, y2, max(1, (y2 - y1) // 10)):
            px = x1 - radius
            if 0 <= px < img_width and 0 <= y < img_height:
                sample_points.append((px, y))
        
        # Sample along right edge (right of the region)
        for y in range(y1, y2, max(1, (y2 - y1) // 10)):
            px = x2 + radius
            if 0 <= px < img_width and 0 <= y < img_height:
                sample_points.append((px, y))
        
        # Sample corners
        corners = [
            (x1 - radius, y1 - radius),
            (x2 + radius, y1 - radius),
            (x1 - radius, y2 + radius),
            (x2 + radius, y2 + radius),
        ]
        for px, py in corners:
            if 0 <= px < img_width and 0 <= py < img_height:
                sample_points.append((px, py))
        
        return sample_points
    
    def _get_sample_points_dense(
        self,
        region: TextRegion,
        radius: int,
        num_points: int
    ) -> List[Tuple[int, int]]:
        """获取密集采样点（改进版）。
        
        在区域周围生成更多的采样点，确保采样的背景色更准确。
        
        参数：
            region: 文本区域
            radius: 采样半径（像素）
            num_points: 每条边的采样点数量
        
        返回：
            采样点坐标列表 [(x, y), ...]
        """
        x1, y1, x2, y2 = region.bbox
        sample_points = []
        
        # 计算每条边的采样间隔
        width = x2 - x1
        height = y2 - y1
        
        # 上边：在区域上方采样
        if width > 0:
            x_step = max(1, width // num_points)
            for x in range(x1, x2, x_step):
                py = y1 - radius
                if py >= 0:
                    sample_points.append((x, py))
        
        # 下边：在区域下方采样
        if width > 0:
            x_step = max(1, width // num_points)
            for x in range(x1, x2, x_step):
                py = y2 + radius
                sample_points.append((x, py))
        
        # 左边：在区域左侧采样
        if height > 0:
            y_step = max(1, height // num_points)
            for y in range(y1, y2, y_step):
                px = x1 - radius
                if px >= 0:
                    sample_points.append((px, y))
        
        # 右边：在区域右侧采样
        if height > 0:
            y_step = max(1, height // num_points)
            for y in range(y1, y2, y_step):
                px = x2 + radius
                sample_points.append((px, y))
        
        # 四个角落
        corners = [
            (x1 - radius, y1 - radius),
            (x2 + radius, y1 - radius),
            (x1 - radius, y2 + radius),
            (x2 + radius, y2 + radius),
        ]
        for px, py in corners:
            if px >= 0 and py >= 0:
                sample_points.append((px, py))
        
        return sample_points

    def _filter_outliers_iqr(self, colors: List[np.ndarray]) -> List[np.ndarray]:
        """使用四分位数范围（IQR）过滤异常值。
        
        IQR = Q3 - Q1
        异常值定义：< Q1 - 1.5*IQR 或 > Q3 + 1.5*IQR
        
        参数：
            colors: 颜色值列表，每个元素是BGR数组
        
        返回：
            过滤后的颜色值列表
        """
        if len(colors) < 4:
            # 样本太少，不进行过滤
            return colors
        
        colors_array = np.array(colors)
        
        # 对每个通道分别计算IQR并过滤
        valid_mask = np.ones(len(colors), dtype=bool)
        
        for channel_idx in range(colors_array.shape[1]):  # 遍历BGR通道
            channel = colors_array[:, channel_idx]
            
            # 计算四分位数
            q1 = np.percentile(channel, 25)
            q3 = np.percentile(channel, 75)
            iqr = q3 - q1
            
            # 计算边界
            lower_bound = q1 - 1.5 * iqr
            upper_bound = q3 + 1.5 * iqr
            
            # 更新mask（所有通道都要在范围内）
            channel_mask = (channel >= lower_bound) & (channel <= upper_bound)
            valid_mask = valid_mask & channel_mask
        
        # 应用mask
        filtered_colors = colors_array[valid_mask]
        
        if len(filtered_colors) == 0:
            # 如果全部被过滤，返回原始数据
            logger.warning("IQR过滤后没有剩余颜色，返回原始数据")
            return colors
        
        logger.debug(f"IQR过滤: {len(colors)} -> {len(filtered_colors)} 个采样点")
        return filtered_colors.tolist()
    
    def _is_valid_point(self, x: int, y: int, image: np.ndarray) -> bool:
        """检查采样点是否有效。
        
        参数：
            x, y: 采样点坐标
            image: 图片数组
        
        返回：
            True如果点在图片范围内
        """
        height, width = image.shape[:2]
        return 0 <= x < width and 0 <= y < height
    
    def _evaluate_background_complexity(self, image: np.ndarray, region: TextRegion) -> str:
        """评估背景复杂度。
        
        通过分析采样点的颜色分布来判断背景的复杂程度。
        
        参数：
            image: 输入图片（BGR格式）
            region: 文本区域
        
        返回：
            "simple" | "medium" | "complex"
        """
        # 使用更多采样点来评估复杂度
        sample_points = self._get_sample_points_dense(
            region,
            self.sample_radius,
            10  # 每条边10个点
        )
        
        if len(sample_points) < 10:
            logger.debug("采样点太少，默认为simple")
            return "simple"
        
        # 采样颜色
        colors = []
        for px, py in sample_points:
            if self._is_valid_point(px, py, image):
                colors.append(image[py, px])
        
        if len(colors) < 10:
            logger.debug("有效采样点太少，默认为simple")
            return "simple"
        
        colors_array = np.array(colors)
        
        # 计算每个通道的标准差
        std_per_channel = np.std(colors_array, axis=0)
        avg_std = np.mean(std_per_channel)
        
        # 计算整体颜色变化
        max_std = np.max(std_per_channel)
        
        # 根据标准差判断复杂度
        if avg_std < 5 and max_std < 10:
            complexity = "simple"  # 纯色或非常接近的颜色
        elif avg_std < 15 and max_std < 30:
            complexity = "medium"  # 渐变或轻微纹理
        else:
            complexity = "complex"  # 复杂图案或纹理
        
        logger.info(f"🔍 背景复杂度评估: avg_std={avg_std:.1f}, max_std={max_std:.1f}, complexity={complexity}")
        return complexity
    
    def _select_fill_strategy(self, complexity: str) -> dict:
        """根据背景复杂度选择填充策略。
        
        参数：
            complexity: 背景复杂度 ("simple" | "medium" | "complex")
        
        返回：
            策略配置字典
        """
        if complexity == "simple":
            # 纯色背景：直接用采样颜色填充，不需要inpainting
            strategy = {
                "use_inpainting": False,
                "fill_method": "color",
                "feather_radius_multiplier": 1.0,
                "padding_multiplier": 0.8
            }
        elif complexity == "medium":
            # 渐变背景：使用小范围inpainting
            strategy = {
                "use_inpainting": True,
                "inpaint_radius": 3,
                "fill_method": "hybrid",  # 先填充颜色，再inpainting边缘
                "feather_radius_multiplier": 1.2,
                "padding_multiplier": 1.0
            }
        else:  # complex
            # 复杂背景：使用纹理复制方法
            strategy = {
                "use_inpainting": False,  # 复杂背景不用inpainting，会产生伪影
                "fill_method": "texture",  # 复制周围纹理
                "feather_radius_multiplier": 1.5,
                "padding_multiplier": 0.7  # 减小padding以减少伪影
            }
        
        logger.info(f"📋 选择填充策略: complexity={complexity}, strategy={strategy}")
        return strategy

    def create_region_mask(self, image: np.ndarray, region: TextRegion) -> np.ndarray:
        """Create a binary mask for a text region.
        
        Args:
            image: Input image as numpy array
            region: Text region to create mask for
            
        Returns:
            Binary mask with the region area set to 255
        """
        height, width = image.shape[:2]
        mask = np.zeros((height, width), dtype=np.uint8)
        
        x1, y1, x2, y2 = region.bbox
        
        # Ensure coordinates are within image bounds
        x1 = max(0, min(x1, width))
        y1 = max(0, min(y1, height))
        x2 = max(0, min(x2, width))
        y2 = max(0, min(y2, height))
        
        # Fill the region with white (255)
        mask[y1:y2, x1:x2] = 255
        
        return mask
    
    def inpaint_region(self, image: np.ndarray, region: TextRegion) -> np.ndarray:
        """Use inpainting to fill a text region with background.
        
        Uses OpenCV's inpaint function to intelligently fill the text
        region based on surrounding pixels.
        
        Args:
            image: Input image as numpy array (BGR format)
            region: Text region to inpaint
            
        Returns:
            Image with the region inpainted
        """
        # Create a copy to avoid modifying the original
        result = image.copy()
        
        # Create mask for the region
        mask = self.create_region_mask(image, region)
        
        # Check if mask has any non-zero pixels
        if not np.any(mask):
            logger.warning(f"Empty mask for region {region.bbox}, returning original image")
            return result
        
        # Apply inpainting using Telea algorithm (faster) or Navier-Stokes (better quality)
        # Using INPAINT_TELEA for better performance
        inpaint_radius = max(3, min(region.width, region.height) // 10)
        
        try:
            result = cv2.inpaint(result, mask, inpaint_radius, cv2.INPAINT_TELEA)
            logger.debug(f"Inpainted region {region.bbox} with radius {inpaint_radius}")
        except cv2.error as e:
            logger.error(f"Inpainting failed for region {region.bbox}: {e}")
            # Fallback: fill with sampled background color
            bg_color = self.sample_background(image, region)
            # Convert RGB to BGR for OpenCV
            bgr_color = (bg_color[2], bg_color[1], bg_color[0])
            x1, y1, x2, y2 = region.bbox
            height, width = image.shape[:2]
            x1 = max(0, min(x1, width))
            y1 = max(0, min(y1, height))
            x2 = max(0, min(x2, width))
            y2 = max(0, min(y2, height))
            result[y1:y2, x1:x2] = bgr_color
        
        return result
    
    def inpaint_regions(self, image: np.ndarray, regions: List[TextRegion]) -> np.ndarray:
        """Inpaint multiple text regions.
        
        Args:
            image: Input image as numpy array (BGR format)
            regions: List of text regions to inpaint
            
        Returns:
            Image with all regions inpainted
        """
        result = image.copy()
        
        for region in regions:
            result = self.inpaint_region(result, region)
        
        return result

    def create_feather_mask(
        self, 
        width: int, 
        height: int, 
        region: TextRegion, 
        feather_radius: int = 5
    ) -> np.ndarray:
        """Create a feathered (soft-edged) mask for a region.
        
        Creates a mask with soft edges using Gaussian blur to enable
        smooth blending between the inpainted region and the original image.
        
        Args:
            width: Image width
            height: Image height
            region: Text region to create feathered mask for
            feather_radius: Radius for the feathering effect
            
        Returns:
            Feathered mask as float array with values 0.0-1.0
        """
        # Create binary mask
        mask = np.zeros((height, width), dtype=np.float32)
        
        x1, y1, x2, y2 = region.bbox
        
        # Ensure coordinates are within image bounds
        x1 = max(0, min(x1, width))
        y1 = max(0, min(y1, height))
        x2 = max(0, min(x2, width))
        y2 = max(0, min(y2, height))
        
        # Fill the region
        mask[y1:y2, x1:x2] = 1.0
        
        # Apply Gaussian blur for feathering
        # Kernel size must be odd
        kernel_size = feather_radius * 2 + 1
        feathered_mask = cv2.GaussianBlur(mask, (kernel_size, kernel_size), 0)
        
        return feathered_mask
    
    def apply_feathering(
        self, 
        original: np.ndarray, 
        inpainted: np.ndarray, 
        region: TextRegion,
        feather_radius: int = 5
    ) -> np.ndarray:
        """Apply feathering effect to blend inpainted region with original.
        
        Uses a feathered mask to smoothly blend the inpainted region
        with the original image, creating natural edge transitions.
        
        Args:
            original: Original image as numpy array (BGR format)
            inpainted: Inpainted image as numpy array (BGR format)
            region: Text region that was inpainted
            feather_radius: Radius for the feathering effect
            
        Returns:
            Blended image with feathered edges
        """
        height, width = original.shape[:2]
        
        # Create feathered mask
        mask = self.create_feather_mask(width, height, region, feather_radius)
        
        # Expand mask to 3 channels for blending
        mask_3ch = np.stack([mask] * 3, axis=-1)
        
        # Blend: result = inpainted * mask + original * (1 - mask)
        result = (inpainted.astype(np.float32) * mask_3ch + 
                  original.astype(np.float32) * (1 - mask_3ch))
        
        return result.astype(np.uint8)
    
    def _check_icon_overlap(
        self,
        padded_bbox: Tuple[int, int, int, int],
        icon_regions: List[TextRegion]
    ) -> Optional[TextRegion]:
        """检查扩大后的区域是否会覆盖图标。
        
        参数：
            padded_bbox: 扩大后的边界框 (x1, y1, x2, y2)
            icon_regions: 图标区域列表
        
        返回：
            如果有重叠,返回重叠的图标区域;否则返回None
        """
        if not icon_regions:
            return None
        
        px1, py1, px2, py2 = padded_bbox
        
        for icon in icon_regions:
            ix1, iy1, ix2, iy2 = icon.bbox
            
            # 检查是否有重叠
            overlap_x = max(0, min(px2, ix2) - max(px1, ix1))
            overlap_y = max(0, min(py2, iy2) - max(py1, iy1))
            
            if overlap_x > 0 and overlap_y > 0:
                overlap_area = overlap_x * overlap_y
                icon_area = (ix2 - ix1) * (iy2 - iy1)
                overlap_ratio = overlap_area / icon_area if icon_area > 0 else 0
                
                # 如果重叠超过图标面积的10%,认为会影响图标
                if overlap_ratio > 0.1:
                    logger.info(f"⚠️ 检测到padding区域会覆盖图标: 重叠比={overlap_ratio:.1%}, "
                               f"图标位置=({ix1},{iy1},{ix2},{iy2})")
                    return icon
        
        return None
    
    def _adjust_padding_for_icon(
        self,
        region_bbox: Tuple[int, int, int, int],
        padding: int,
        icon: TextRegion
    ) -> int:
        """调整padding以避免覆盖图标。
        
        参数：
            region_bbox: 文本区域边界框 (x1, y1, x2, y2)
            padding: 原始padding值
            icon: 要保护的图标区域
        
        返回：
            调整后的padding值
        """
        x1, y1, x2, y2 = region_bbox
        ix1, iy1, ix2, iy2 = icon.bbox
        
        # 计算文本区域到图标的距离
        # 上方距离
        dist_top = y1 - iy2 if y1 > iy2 else float('inf')
        # 下方距离
        dist_bottom = iy1 - y2 if iy1 > y2 else float('inf')
        # 左侧距离
        dist_left = x1 - ix2 if x1 > ix2 else float('inf')
        # 右侧距离
        dist_right = ix1 - x2 if ix1 > x2 else float('inf')
        
        # 找到最小距离
        min_dist = min(dist_top, dist_bottom, dist_left, dist_right)
        
        if min_dist == float('inf'):
            # 文本区域与图标重叠,使用最小padding
            adjusted_padding = 5
            logger.info(f"🔧 文本与图标重叠,使用最小padding={adjusted_padding}px")
        elif min_dist < padding:
            # 距离小于padding,调整padding以留出5px的安全距离
            adjusted_padding = max(5, int(min_dist - 5))
            logger.info(f"🔧 调整padding以保护图标: {padding}px -> {adjusted_padding}px (距离={min_dist:.0f}px)")
        else:
            # 距离足够,不需要调整
            adjusted_padding = padding
        
        return adjusted_padding
    
    def process_region(
        self, 
        image: np.ndarray, 
        region: TextRegion,
        background_color: Tuple[int, int, int],
        feather_radius: int = None,
        padding: int = None,
        icon_regions: List[TextRegion] = None
    ) -> np.ndarray:
        """处理文本区域，使用简化的纯色填充策略。
        
        **简化策略（用户反馈）：直接使用文字底色作为背景，不使用复杂的采样和多层填充。**
        
        流程：
        1. 计算智能padding（确保完全覆盖原文）
        2. 检查是否会覆盖图标（国徽、二维码等）并调整padding
        3. 用纯色填充扩大区域
        4. 应用羽化效果（平滑边缘）
        
        参数：
            image: 输入图片（BGR格式）
            region: 要处理的文本区域
            background_color: 背景颜色（RGB格式）
            feather_radius: 羽化效果的半径（None则自动计算）
            padding: 扩大区域的像素数（None则自动计算）
            icon_regions: 图标区域列表（用于保护图标不被抹除）
            
        返回：
            处理后的图片，区域已填充并羽化
        """
        height, width = image.shape[:2]
        x1, y1, x2, y2 = region.bbox
        
        # 获取精确擦除模式配置
        # 优先使用竖版专用配置（如果存在）
        if hasattr(self, '_portrait_precise_erase'):
            precise_erase = self._portrait_precise_erase
        else:
            precise_erase = self.config.get('rendering.precise_erase_mode', False)
        
        # 判断是否在底部区域（底部20%的区域）
        bottom_threshold = height * 0.8  # 图片底部20%
        is_bottom_region = y1 > bottom_threshold
        
        # 判断是否是横版（宽 > 高）
        is_landscape = width > height
        
        # 检查是否是"营业执照"标题
        is_business_license_title = self._is_business_license_title(region.text)
        
        # 检查是否在二维码附近
        is_near_qrcode = self._is_near_qrcode(region, icon_regions)
        
        # 检查是否是二维码说明文字(通过文本内容判断)
        is_qrcode_description = self._is_qrcode_description_text(region.text)
        
        # 判断是否应该使用精确擦除模式
        # 1. 竖版文档：对底部区域使用精确擦除模式
        # 2. 横版文档：对底部区域也使用精确擦除模式（新增）
        # 3. "营业执照"标题文字使用精确擦除
        # 4. 二维码附近的文字使用精确擦除
        # 5. 二维码说明文字使用精确擦除
        should_use_precise_erase = (
            (precise_erase and is_bottom_region) or 
            is_business_license_title or 
            is_near_qrcode or
            is_qrcode_description
        )
        
        # 特殊处理：如果是二维码说明文字，检查是否与二维码区域重叠
        # 如果重叠，裁剪bbox以避开二维码
        if is_qrcode_description and icon_regions:
            # 遍历所有图标区域，找到二维码
            for icon in icon_regions:
                if not self._is_likely_qrcode(icon):
                    continue
                
                # 找到了二维码，检查文本是否与它重叠
                qr_x1, qr_y1, qr_x2, qr_y2 = icon.bbox
                
                # 如果文本区域与二维码区域重叠
                if (x1 < qr_x2 and x2 > qr_x1 and 
                    y1 < qr_y2 and y2 > qr_y1):
                    
                    # 计算重叠区域
                    overlap_x1 = max(x1, qr_x1)
                    overlap_x2 = min(x2, qr_x2)
                    overlap_width = overlap_x2 - overlap_x1
                    text_width = x2 - x1
                    overlap_ratio = overlap_width / text_width if text_width > 0 else 0
                    
                    logger.info(
                        f"🔍 检测到二维码说明文字与二维码区域重叠: "
                        f"text='{region.text[:30]}...', "
                        f"text_bbox=({x1},{y1},{x2},{y2}), "
                        f"qr_bbox=({qr_x1},{qr_y1},{qr_x2},{qr_y2}), "
                        f"overlap_ratio={overlap_ratio:.1%}"
                    )
                    
                    # 如果重叠超过50%，完全跳过擦除
                    if overlap_ratio > 0.5:
                        logger.info(
                            f"⚠️ 跳过擦除：重叠比例过高({overlap_ratio:.1%})"
                        )
                        return image.copy()
                    
                    # 否则，裁剪bbox以避开二维码
                    # 如果文本左边界在二维码内，移到二维码右侧
                    if x1 < qr_x2 and x2 > qr_x2:
                        original_x1 = x1
                        x1 = qr_x2 + 5  # 留5px间隙
                        logger.info(
                            f"✂️ 裁剪bbox（左侧）：x1 {original_x1} -> {x1}"
                        )
                        region.bbox = (x1, y1, x2, y2)
                    
                    # 如果文本右边界在二维码内，移到二维码左侧
                    elif x2 > qr_x1 and x1 < qr_x1:
                        original_x2 = x2
                        x2 = qr_x1 - 5  # 留5px间隙
                        logger.info(
                            f"✂️ 裁剪bbox（右侧）：x2 {original_x2} -> {x2}"
                        )
                        region.bbox = (x1, y1, x2, y2)
                    
                    # 如果文本上边界在二维码内，移到二维码下方
                    elif y1 < qr_y2 and y2 > qr_y2:
                        original_y1 = y1
                        y1 = qr_y2 + 5  # 留5px间隙
                        logger.info(
                            f"✂️ 裁剪bbox（上侧）：y1 {original_y1} -> {y1}"
                        )
                        region.bbox = (x1, y1, x2, y2)
                    
                    # 如果文本下边界在二维码内，移到二维码上方
                    elif y2 > qr_y1 and y1 < qr_y1:
                        original_y2 = y2
                        y2 = qr_y1 - 5  # 留5px间隙
                        logger.info(
                            f"✂️ 裁剪bbox（下侧）：y2 {original_y2} -> {y2}"
                        )
                        region.bbox = (x1, y1, x2, y2)
                    
                    # 只处理第一个匹配的二维码
                    break
        
        if should_use_precise_erase:
            # 精确擦除模式：使用配置的 erase_padding 扩展擦除区域
            # 获取 erase_padding 配置（默认2像素）
            erase_padding = self.config.get('rendering.erase_padding', 2)
            
            # 对于底部区域，使用更大的 padding 以确保完全擦除
            # 对于其他区域，使用配置的 padding
            if is_bottom_region:
                erase_padding = 1  # 底部区域：使用5像素padding，确保完全擦除
                logger.info(f"🔧 底部区域检测：使用5px padding，确保完全擦除")
            
            # 定义边框保护区域（只保护最边缘的装饰边框）
            border_margin = 10  # 减小到10像素，减少对底部文字的影响
            
            # 扩展擦除区域（向外扩展 erase_padding 像素）
            x1_erase = max(border_margin, x1 - erase_padding)  # 保护左边框
            y1_erase = max(border_margin, y1 - erase_padding)  # 保护上边框
            x2_erase = min(width - border_margin, x2 + erase_padding)  # 保护右边框
            y2_erase = min(height - border_margin, y2 + erase_padding)  # 保护下边框
            
            # 直接在原图上填充，不创建新的region对象
            result = image.copy()
            
            # 转换RGB到BGR
            bgr_color = (background_color[2], background_color[1], background_color[0])
            
            # 使用扩展后的边界框填充
            result[y1_erase:y2_erase, x1_erase:x2_erase] = bgr_color
            
            # 生成日志标签
            if is_business_license_title:
                reason = "营业执照标题"
            elif is_qrcode_description:
                reason = "二维码说明文字"
            elif is_near_qrcode:
                reason = "二维码附近"
            elif is_bottom_region:
                orientation_label = "横版" if is_landscape else "竖版"
                reason = f"{orientation_label}底部"
            else:
                reason = "未知原因"
            
            logger.info(f"✅ 精确擦除模式（{reason}）: 原始bbox=({x1},{y1},{x2},{y2}), 扩展后=({x1_erase},{y1_erase},{x2_erase},{y2_erase}), padding={erase_padding}px, 边框保护={border_margin}px, RGB{background_color}, 无羽化, 文本='{region.text[:20]}...'")
            
            return result
        
        # 原有的智能padding逻辑（非精确模式）
        # 使用智能padding计算
        if padding is None:
            padding = self.calculate_padding(region)
        
        # 横版文档增强padding（仅当配置文件是横版配置时）
        # 检查配置方向，只对横版配置的横版图片增强padding
        config_orientation = self.config.get('document_orientation', 'vertical')
        if is_landscape and config_orientation == 'horizontal':
            # 横版配置 + 横版图片：使用更大的padding以确保完全擦除底部文字
            original_padding = padding
            padding = int(padding * 1.5)  # 增加50%的padding
            logger.info(f"🔧 横版文档增强padding: {original_padding}px -> {padding}px (增加50%)")
        
        # === 图标保护逻辑：检查padding后是否会覆盖图标 ===
        if icon_regions:
            # 先计算初步的padded区域
            temp_x1 = max(0, x1 - padding)
            temp_y1 = max(0, y1 - padding)
            temp_x2 = min(width, x2 + padding)
            temp_y2 = min(height, y2 + padding)
            
            # 检查是否会覆盖图标
            overlapping_icon = self._check_icon_overlap(
                (temp_x1, temp_y1, temp_x2, temp_y2),
                icon_regions
            )
            
            if overlapping_icon:
                # 调整padding以避免覆盖图标
                padding = self._adjust_padding_for_icon(
                    (x1, y1, x2, y2),
                    padding,
                    overlapping_icon
                )
                logger.info(f"🛡️ 图标保护：调整padding以避免覆盖图标 (国徽/二维码等)")
        
        # 创建扩大的区域，同时保护图片边框
        height, width = image.shape[:2]
        x1, y1, x2, y2 = region.bbox
        
        # 定义边框保护区域（使用百分比，适应不同尺寸的图片）
        # 营业执照的花纹边框通常占据边缘2-3%的区域
        border_margin_ratio = 0.025  # 2.5%的边缘区域
        border_margin = int(min(width, height) * border_margin_ratio)
        border_margin = max(40, min(border_margin, 100))  # 限制在40-100像素之间
        
        # 检测阈值：使用百分比而不是固定像素
        # 如果文本距离边缘小于5%的图片尺寸，认为靠近边框
        near_threshold_ratio = 0.05  # 5%
        near_threshold_x = int(width * near_threshold_ratio)
        near_threshold_y = int(height * near_threshold_ratio)
        
        # 检测文本是否靠近边框
        near_left = x1 < border_margin + near_threshold_x
        near_top = y1 < border_margin + near_threshold_y
        near_right = x2 > width - border_margin - near_threshold_x
        near_bottom = y2 > height - border_margin - near_threshold_y
        
        # 如果靠近边框，增加向内的padding以确保完全覆盖原文
        # 同时限制向外的padding以保护边框
        if near_left or near_top or near_right or near_bottom:
            # 靠近边框时，使用更大的inner padding（2.0倍）以彻底擦除原文
            # 同时减小outer padding（0.5倍）以更好地保护边框
            inner_padding = int(padding * 2.0)
            outer_padding = max(5, int(padding * 0.5))  # 至少5px,避免太小
            
            # 向内扩展（远离边框方向）使用更大的padding
            # 向外扩展（靠近边框方向）使用正常padding并限制在边框内
            x1_padded = max(border_margin, x1 - (inner_padding if not near_left else outer_padding))
            y1_padded = max(border_margin, y1 - (inner_padding if not near_top else outer_padding))
            x2_padded = min(width - border_margin, x2 + (inner_padding if not near_right else outer_padding))
            y2_padded = min(height - border_margin, y2 + (inner_padding if not near_bottom else outer_padding))
            
            logger.info(
                f"Border protection with enhanced inner padding for '{region.text[:20]}...': "
                f"near_edges=(L:{near_left},T:{near_top},R:{near_right},B:{near_bottom}), "
                f"inner_padding={inner_padding}px, outer_padding={outer_padding}px, "
                f"bbox=({x1_padded},{y1_padded},{x2_padded},{y2_padded})"
            )
        else:
            # 不靠近边框，正常扩展
            x1_padded = max(0, x1 - padding)
            y1_padded = max(0, y1 - padding)
            x2_padded = min(width, x2 + padding)
            y2_padded = min(height, y2 + padding)
        
        padded_region = TextRegion(
            bbox=(x1_padded, y1_padded, x2_padded, y2_padded),
            text=region.text,
            confidence=region.confidence,
            font_size=region.font_size,
            angle=region.angle
        )
        
        # 纯色填充
        result = self.fill_region_with_color(image, padded_region, background_color)
        logger.debug(f"纯色填充: RGB{background_color}, padding={padding}px")
        
        # 羽化平滑过渡
        if feather_radius is None:
            feather_radius = self.calculate_feather_radius(region)
        
        result = self.apply_feathering(image, result, padded_region, feather_radius)
        logger.debug(f"羽化: radius={feather_radius}px")
        
        logger.info(f"✅ 区域处理完成（简化策略）: bbox={region.bbox}, padding={padding}px, feather={feather_radius}px")
        return result
    
    def fill_region_with_color(
        self, 
        image: np.ndarray, 
        region: TextRegion, 
        color: Tuple[int, int, int]
    ) -> np.ndarray:
        """Fill a text region with a solid color.
        
        Args:
            image: Input image as numpy array (BGR format)
            region: Text region to fill
            color: RGB color tuple to fill with
            
        Returns:
            Image with the region filled
        """
        result = image.copy()
        
        x1, y1, x2, y2 = region.bbox
        height, width = image.shape[:2]
        
        # Ensure coordinates are within image bounds
        x1 = max(0, min(x1, width))
        y1 = max(0, min(y1, height))
        x2 = max(0, min(x2, width))
        y2 = max(0, min(y2, height))
        
        # Convert RGB to BGR for OpenCV
        bgr_color = (color[2], color[1], color[0])
        result[y1:y2, x1:x2] = bgr_color
        
        return result
    
    def _expand_region(self, region: TextRegion, padding: int, image_shape: Optional[Tuple[int, int]] = None) -> TextRegion:
        """扩大区域边界，同时保护图片边框。
        
        参数：
            region: 原始文本区域
            padding: 扩大的像素数
            image_shape: 图片形状 (height, width)，用于边框保护
        
        返回：
            扩大后的TextRegion
        """
        x1, y1, x2, y2 = region.bbox
        
        # 扩大边界
        x1_padded = x1 - padding
        y1_padded = y1 - padding
        x2_padded = x2 + padding
        y2_padded = y2 + padding
        
        # 如果提供了图片尺寸，进行边框保护
        if image_shape is not None:
            height, width = image_shape
            
            # 定义边框保护区域（距离边缘的像素数）
            border_margin = 40  # 保护边缘40像素的区域（花纹边框通常在这个范围内）
            
            # 限制扩展范围，避免覆盖边框
            x1_padded = max(border_margin, x1_padded)
            y1_padded = max(border_margin, y1_padded)
            x2_padded = min(width - border_margin, x2_padded)
            y2_padded = min(height - border_margin, y2_padded)
            
            logger.debug(
                f"Border protection applied: original=({x1-padding},{y1-padding},{x2+padding},{y2+padding}), "
                f"protected=({x1_padded},{y1_padded},{x2_padded},{y2_padded})"
            )
        
        return TextRegion(
            bbox=(x1_padded, y1_padded, x2_padded, y2_padded),
            text=region.text,
            confidence=region.confidence,
            font_size=region.font_size,
            angle=region.angle
        )
    
    def _fill_with_texture(
        self,
        image: np.ndarray,
        region: TextRegion,
        padding: int
    ) -> np.ndarray:
        """使用纹理复制方法填充区域。
        
        从区域周围复制纹理来填充，适用于复杂背景。
        
        参数：
            image: 输入图片
            region: 文本区域
            padding: padding值
        
        返回：
            填充后的图片
        """
        result = image.copy()
        height, width = image.shape[:2]
        
        x1, y1, x2, y2 = region.bbox
        
        # 确保坐标在图片范围内
        x1 = max(0, x1)
        y1 = max(0, y1)
        x2 = min(width, x2)
        y2 = min(height, y2)
        
        # 计算扩大后的区域
        x1_padded = max(0, x1 - padding)
        y1_padded = max(0, y1 - padding)
        x2_padded = min(width, x2 + padding)
        y2_padded = min(height, y2 + padding)
        
        # 从上方复制纹理
        if y1_padded > 0:
            sample_height = min(padding, y1_padded)
            if sample_height > 0:
                sample_region = image[y1_padded - sample_height:y1_padded, x1_padded:x2_padded]
                if sample_region.size > 0:
                    # 重复填充
                    fill_height = y2_padded - y1_padded
                    if fill_height > 0:
                        # 使用tile重复纹理
                        repeat_times = (fill_height // sample_height) + 1
                        tiled = np.tile(sample_region, (repeat_times, 1, 1))
                        result[y1_padded:y2_padded, x1_padded:x2_padded] = tiled[:fill_height, :, :]
        
        logger.debug(f"使用纹理填充: region={region.bbox}, padded=({x1_padded},{y1_padded},{x2_padded},{y2_padded})")
        return result
    
    def _fill_background_multi_layer(
        self,
        image: np.ndarray,
        region: TextRegion,
        padding: int,
        strategy: dict
    ) -> np.ndarray:
        """多层背景填充。
        
        根据策略分层填充背景：
        1. 第一层：基础填充（颜色或纹理）
        2. 第二层：边缘修复（可选的inpainting）
        3. 第三层：羽化平滑过渡
        
        参数：
            image: 输入图片
            region: 文本区域
            padding: padding值
            strategy: 填充策略
        
        返回：
            处理后的图片
        """
        result = image.copy()
        height, width = image.shape[:2]
        
        # 应用padding倍数
        padding = int(padding * strategy.get('padding_multiplier', 1.0))
        
        # 创建扩大的区域
        x1, y1, x2, y2 = region.bbox
        x1_padded = max(0, x1 - padding)
        y1_padded = max(0, y1 - padding)
        x2_padded = min(width, x2 + padding)
        y2_padded = min(height, y2 + padding)
        
        padded_region = TextRegion(
            bbox=(x1_padded, y1_padded, x2_padded, y2_padded),
            text=region.text,
            confidence=region.confidence,
            font_size=region.font_size,
            angle=region.angle
        )
        
        fill_method = strategy.get('fill_method', 'color')
        
        # 第一层：基础填充
        if fill_method == 'color':
            # 纯色填充
            bg_color = self.sample_background_improved(image, region)
            result = self.fill_region_with_color(result, padded_region, bg_color)
            logger.debug(f"第一层：纯色填充 RGB{bg_color}")
            
        elif fill_method == 'texture':
            # 纹理填充
            result = self._fill_with_texture(result, region, padding)
            logger.debug("第一层：纹理填充")
            
        elif fill_method == 'hybrid':
            # 混合填充：先颜色，再inpainting边缘
            bg_color = self.sample_background_improved(image, region)
            result = self.fill_region_with_color(result, padded_region, bg_color)
            logger.debug(f"第一层：混合填充（颜色） RGB{bg_color}")
        
        # 第二层：边缘修复（可选）
        if strategy.get('use_inpainting', False):
            inpaint_radius = strategy.get('inpaint_radius', 3)
            try:
                mask = self.create_region_mask(image, padded_region)
                if np.any(mask):
                    result = cv2.inpaint(result, mask, inpaint_radius, cv2.INPAINT_TELEA)
                    logger.debug(f"第二层：inpainting修复，radius={inpaint_radius}")
            except Exception as e:
                logger.warning(f"Inpainting失败: {e}")
        
        # 第三层：羽化平滑过渡
        base_feather_radius = self.calculate_feather_radius(region)
        feather_radius = int(base_feather_radius * strategy.get('feather_radius_multiplier', 1.0))
        feather_radius = max(self.feather_min, min(self.feather_max, feather_radius))
        
        result = self.apply_feathering(image, result, padded_region, feather_radius)
        logger.debug(f"第三层：羽化，radius={feather_radius}")
        
        logger.info(f"✅ 多层填充完成: method={fill_method}, padding={padding}px, feather={feather_radius}px")
        return result
    
    def is_region_covered(
        self, 
        original: np.ndarray, 
        processed: np.ndarray, 
        region: TextRegion,
        tolerance: float = 0.01
    ) -> bool:
        """Check if a text region has been fully covered/processed.
        
        Compares the original and processed images to verify that
        the text region has been modified.
        
        Args:
            original: Original image as numpy array
            processed: Processed image as numpy array
            region: Text region to check
            tolerance: Tolerance for considering pixels as changed
            
        Returns:
            True if the region has been covered/modified
        """
        x1, y1, x2, y2 = region.bbox
        height, width = original.shape[:2]
        
        # Ensure coordinates are within image bounds
        x1 = max(0, min(x1, width))
        y1 = max(0, min(y1, height))
        x2 = max(0, min(x2, width))
        y2 = max(0, min(y2, height))
        
        if x2 <= x1 or y2 <= y1:
            return True  # Empty region is considered covered
        
        # Extract region from both images
        original_region = original[y1:y2, x1:x2].astype(np.float32)
        processed_region = processed[y1:y2, x1:x2].astype(np.float32)
        
        # Calculate difference
        diff = np.abs(original_region - processed_region)
        max_diff = np.max(diff)
        
        # If there's any significant difference, the region has been modified
        # We use a small tolerance to account for floating point errors
        return max_diff > tolerance * 255

    def detect_seal_region(self, image: np.ndarray, region: TextRegion) -> Optional[Tuple[int, int, int, int]]:
        """智能检测区域内的红色印章位置（通用版本）。
        
        智能策略（适用于多种图片）：
        1. 自适应红色检测：根据图片整体色调动态调整红色阈值
        2. 多层次判断：
           - 检查文字区域本身的红色特征（颜色、占比、分布）
           - 检查周围是否有印章结构（形状、大小、位置关系）
        3. 智能过滤：区分红色文字、红色装饰和真正的印章
        4. 扩大检测范围：考虑翻译后文字可能覆盖印章的情况
        
        参数：
            image: 输入图片（BGR格式）
            region: 文本区域
        
        返回：
            (x1, y1, x2, y2) 印章的边界框，如果没有印章返回None
        """
        x1, y1, x2, y2 = region.bbox
        height, width = image.shape[:2]
        region_area = (x2 - x1) * (y2 - y1)
        
        # 获取配置
        search_padding = self.config.get('rendering.seal_detection.search_padding', 100)  # 增加到100
        base_red_threshold = self.config.get('rendering.seal_detection.red_threshold', [120, 120, 120])
        
        # === 步骤1：分析文字区域本身的红色特征 ===
        region_img = image[y1:y2, x1:x2]
        b, g, r = cv2.split(region_img)
        
        # 计算平均颜色和标准差
        avg_r, avg_g, avg_b = np.mean(r), np.mean(g), np.mean(b)
        std_r, std_g, std_b = np.std(r), np.std(g), np.std(b)
        
        # 判断是否是红色文字（文字本身是红色）
        # 条件：R通道明显高于G和B，且颜色比较均匀（标准差小）
        is_red_text = (avg_r > avg_g + 30 and avg_r > avg_b + 30 and 
                      avg_r > 150 and std_r < 50)
        
        # 计算红色像素（使用自适应阈值）
        red_pixels = ((r > base_red_threshold[0]) & 
                     (g < base_red_threshold[1]) & 
                     (b < base_red_threshold[2])).sum()
        red_ratio = red_pixels / region_area if region_area > 0 else 0
        
        # 判断红色像素的分布特征
        # 如果红色像素集中在边缘，可能是装饰；如果分布均匀，可能是印章内文字
        if red_pixels > 0:
            red_mask = ((r > base_red_threshold[0]) & 
                       (g < base_red_threshold[1]) & 
                       (b < base_red_threshold[2])).astype(np.uint8)
            # 计算红色像素的中心
            red_y, red_x = np.where(red_mask > 0)
            if len(red_y) > 0:
                red_center_y = np.mean(red_y)
                red_center_x = np.mean(red_x)
                region_center_y = (y2 - y1) / 2
                region_center_x = (x2 - x1) / 2
                # 红色像素中心距离区域中心的距离
                center_dist = np.sqrt((red_center_y - region_center_y)**2 + 
                                     (red_center_x - region_center_x)**2)
                max_dist = np.sqrt(region_center_y**2 + region_center_x**2)
                is_centered = center_dist < max_dist * 0.3  # 红色像素集中在中心
            else:
                is_centered = False
        else:
            is_centered = False
        
        # 方法1：如果是明显的红色文字（高占比+红色+集中分布），判定为印章内
        if is_red_text and red_ratio > 0.2 and is_centered:
            logger.info(f"🔴 方法1：检测到红色文字 - 占比{red_ratio:.1%}, "
                       f"平均RGB({avg_r:.0f},{avg_g:.0f},{avg_b:.0f}), 判定为印章内文字")
            return (x1 - 50, y1 - 50, x2 + 50, y2 + 50)
        
        # === 步骤2：在周围搜索印章结构 ===
        # 动态调整搜索范围（小文字用小范围，大文字用大范围）
        adaptive_padding = min(search_padding, max(30, int(np.sqrt(region_area) * 2)))
        sx1 = max(0, x1 - adaptive_padding)
        sy1 = max(0, y1 - adaptive_padding)
        sx2 = min(width, x2 + adaptive_padding)
        sy2 = min(height, y2 + adaptive_padding)
        
        search_area = image[sy1:sy2, sx1:sx2]
        search_b, search_g, search_r = cv2.split(search_area)
        
        # 使用自适应阈值检测红色
        red_mask = ((search_r > base_red_threshold[0]) & 
                   (search_g < base_red_threshold[1]) & 
                   (search_b < base_red_threshold[2])).astype(np.uint8) * 255
        
        red_pixel_count = red_mask.sum() // 255
        
        # 动态调整最小红色像素阈值（基于搜索区域大小）
        search_area_size = (sx2 - sx1) * (sy2 - sy1)
        min_red_pixels = max(200, int(search_area_size * 0.05))  # 至少占搜索区域的5%
        
        if red_pixel_count < min_red_pixels:
            logger.debug(f"未检测到印章: 红色像素={red_pixel_count} < {min_red_pixels}")
            return None
        
        # 找到红色区域的轮廓
        contours, _ = cv2.findContours(red_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        if not contours:
            return None
        
        # 找到最大的轮廓（可能是印章）
        largest_contour = max(contours, key=cv2.contourArea)
        contour_area = cv2.contourArea(largest_contour)
        
        # 检查轮廓是否像印章（圆形或方形，不是细长条）
        seal_x, seal_y, seal_w, seal_h = cv2.boundingRect(largest_contour)
        aspect_ratio = seal_w / seal_h if seal_h > 0 else 0
        
        # 印章通常是接近正方形或圆形的（宽高比在0.5-2之间）
        is_seal_shape = 0.5 <= aspect_ratio <= 2.0
        
        # 印章面积应该比文字区域大（至少2倍）
        is_seal_size = contour_area > region_area * 2
        
        if not (is_seal_shape and is_seal_size):
            logger.debug(f"红色区域不像印章: aspect_ratio={aspect_ratio:.2f}, "
                        f"area_ratio={contour_area/region_area:.2f}")
            return None
        
        # 转换回原图坐标
        seal_x1 = sx1 + seal_x
        seal_y1 = sy1 + seal_y
        seal_x2 = seal_x1 + seal_w
        seal_y2 = seal_y1 + seal_h
        
        # === 步骤3：判断文字与印章的位置关系 ===
        # 计算重叠
        overlap_x = max(0, min(x2, seal_x2) - max(x1, seal_x1))
        overlap_y = max(0, min(y2, seal_y2) - max(y1, seal_y1))
        overlap_area = overlap_x * overlap_y
        overlap_ratio = overlap_area / region_area if region_area > 0 else 0
        
        # 计算文字中心点
        region_center_x = (x1 + x2) // 2
        region_center_y = (y1 + y2) // 2
        
        # 检查中心点是否在印章内
        is_center_in_seal = (seal_x1 <= region_center_x <= seal_x2 and 
                            seal_y1 <= region_center_y <= seal_y2)
        
        # 新的判断标准：只要有任何重叠就跳过（降低阈值到1%）
        # 这样可以避免翻译文字覆盖到印章
        if overlap_ratio > 0.01 or is_center_in_seal:
            logger.info(f"🔴 检测到印章重叠 ({seal_x1}, {seal_y1}) -> ({seal_x2}, {seal_y2}), "
                       f"文字区域=({x1}, {y1}) -> ({x2}, {y2}), "
                       f"红色像素={red_pixel_count}, 重叠比={overlap_ratio:.1%}, "
                       f"中心在印章内={is_center_in_seal}, 跳过翻译")
            return (seal_x1, seal_y1, seal_x2, seal_y2)
        
        logger.debug(f"检测到印章但文字不在印章内: 重叠比={overlap_ratio:.1%}, "
                    f"中心在印章内={is_center_in_seal}, "
                    f"文字区域=({x1}, {y1}) -> ({x2}, {y2}), "
                    f"印章区域=({seal_x1}, {seal_y1}) -> ({seal_x2}, {seal_y2})")
        return None
    
    def calculate_below_seal_position(
        self, 
        seal_bbox: Tuple[int, int, int, int],
        text_width: int,
        text_height: int,
        image_height: int,
        gap: int = None
    ) -> Optional[Tuple[int, int]]:
        """计算印章下方的文字位置。
        
        参数：
            seal_bbox: 印章边界 (x1, y1, x2, y2)
            text_width: 文字宽度
            text_height: 文字高度
            image_height: 图片高度（用于检查是否超出边界）
            gap: 印章和文字之间的间距（None则从配置读取）
        
        返回：
            (x, y) 文字左上角坐标，如果空间不足返回None
        """
        if gap is None:
            gap = self.config.get('rendering.seal_handling.gap', 10)
        
        seal_x1, seal_y1, seal_x2, seal_y2 = seal_bbox
        
        # 印章中心的x坐标
        seal_center_x = (seal_x1 + seal_x2) // 2
        
        # 文字居中对齐印章
        text_x = seal_center_x - text_width // 2
        
        # 文字放在印章下方，留gap间距
        text_y = seal_y2 + gap
        
        # 检查是否超出图片底部
        if text_y + text_height > image_height - 10:
            logger.warning(f"⚠️ 印章下方空间不足: text_y={text_y}, text_height={text_height}, image_height={image_height}")
            return None
        
        # 检查文字是否会与印章重叠
        if text_y < seal_y2 + gap:
            logger.warning(f"⚠️ 文字会与印章重叠: text_y={text_y}, seal_y2={seal_y2}")
            return None
        
        logger.debug(f"📍 计算印章下方位置: seal_center_x={seal_center_x}, text_pos=({text_x}, {text_y}), gap={gap}px")
        return (text_x, text_y)

    def _is_business_license_title(self, text: str) -> bool:
        """检查文本是否是"营业执照"标题。
        
        参数：
            text: 要检查的文本
        
        返回：
            True如果是营业执照标题
        """
        # 去除空格和标点符号
        cleaned_text = text.strip().replace(' ', '').replace('　', '')
        
        # 检查是否包含"营业执照"
        if '营业执照' in cleaned_text:
            logger.info(f"检测到营业执照标题: '{text}'")
            return True
        
        return False
    
    def _is_near_qrcode(
        self, 
        region: TextRegion, 
        icon_regions: List[TextRegion]
    ) -> bool:
        """检查文本区域是否在二维码附近。
        
        参数：
            region: 要检查的文本区域
            icon_regions: 图标区域列表（包含二维码）
        
        返回：
            True如果在二维码附近
        """
        if not icon_regions:
            return False
        
        # 定义"附近"的距离阈值（像素）
        # 使用配置或默认值
        distance_threshold = self.config.get('rendering.qrcode_proximity_threshold', 150)
        
        x1, y1, x2, y2 = region.bbox
        region_center_x = (x1 + x2) // 2
        region_center_y = (y1 + y2) // 2
        
        # 检查每个图标区域
        for icon in icon_regions:
            # 检查是否是二维码（通过文本内容或其他特征）
            # 二维码通常被OCR识别为空文本或乱码
            if not self._is_likely_qrcode(icon):
                continue
            
            # 计算文本区域到二维码的距离
            # 使用边界框之间的最短距离,而不是中心点距离
            ix1, iy1, ix2, iy2 = icon.bbox
            
            # 计算水平距离
            if x2 < ix1:  # 文本在二维码左侧
                dx = ix1 - x2
            elif x1 > ix2:  # 文本在二维码右侧
                dx = x1 - ix2
            else:  # 水平重叠
                dx = 0
            
            # 计算垂直距离
            if y2 < iy1:  # 文本在二维码上方
                dy = iy1 - y2
            elif y1 > iy2:  # 文本在二维码下方
                dy = y1 - iy2
            else:  # 垂直重叠
                dy = 0
            
            # 计算欧几里得距离
            distance = (dx ** 2 + dy ** 2) ** 0.5
            
            if distance <= distance_threshold:
                logger.info(
                    f"检测到二维码附近的文字: 文本='{region.text[:20]}...', "
                    f"距离={distance:.1f}px, 阈值={distance_threshold}px, "
                    f"二维码位置={icon.bbox}"
                )
                return True
        
        return False
    
    def _is_likely_qrcode(self, icon: TextRegion) -> bool:
        """判断图标区域是否可能是二维码。
        
        二维码的特征：
        1. 接近正方形（高宽比接近1）
        2. 面积适中（通常在2500-250000像素之间，即50x50到500x500）
        3. OCR文本为空或很短（二维码不包含可读文本）
        4. 位置可以在图片任何地方（不限制位置）
        
        参数：
            icon: 图标区域
        
        返回：
            True如果可能是二维码
        """
        # 检查高宽比（二维码接近正方形）
        aspect_ratio = icon.aspect_ratio
        if aspect_ratio < 0.7:  # 不够接近正方形
            logger.debug(f"Not QR: aspect_ratio={aspect_ratio:.2f} < 0.7")
            return False
        
        # 检查面积（二维码通常在合理范围内）
        # 最小：50x50 = 2500
        # 最大：500x500 = 250000
        if icon.area < 2500:
            logger.debug(f"Not QR: area={icon.area} < 2500 (too small)")
            return False
        
        if icon.area > 250000:
            logger.debug(f"Not QR: area={icon.area} > 250000 (too large)")
            return False
        
        # 检查文本长度（二维码通常被识别为空或很短的文本）
        text_length = len(icon.text.strip())
        if text_length > 30:  # 如果文本太长，不太可能是二维码
            logger.debug(f"Not QR: text_length={text_length} > 30")
            return False
        
        # 综合判断：如果满足以上所有条件，很可能是二维码
        # 额外加分项：
        # - 文本为空或很短（<5个字符）
        # - 高宽比非常接近1（0.85-1.15）
        # - 面积在典型二维码范围（6400-160000，即80x80到400x400）
        
        score = 0
        
        # 基础分：满足基本条件
        score += 1
        
        # 文本为空或很短
        if text_length == 0:
            score += 2  # 空文本，很可能是二维码
        elif text_length <= 5:
            score += 1  # 很短的文本，可能是二维码
        
        # 高宽比非常接近1
        if 0.85 <= aspect_ratio <= 1.15:
            score += 1
        
        # 面积在典型范围
        if 6400 <= icon.area <= 160000:
            score += 1
        
        # 判断阈值：得分>=3认为是二维码
        is_qrcode = score >= 3
        
        if is_qrcode:
            logger.info(
                f"✅ 识别为二维码: bbox={icon.bbox}, "
                f"aspect_ratio={aspect_ratio:.2f}, area={icon.area}, "
                f"text_length={text_length}, text='{icon.text[:20]}...', "
                f"score={score}"
            )
        else:
            logger.debug(
                f"Not QR: score={score} < 3 "
                f"(aspect={aspect_ratio:.2f}, area={icon.area}, text_len={text_length})"
            )
        
        return is_qrcode

    def _is_qrcode_description_text(self, text: str) -> bool:
        """检查文本是否是二维码说明文字。
        
        通过文本内容判断是否是二维码相关的说明文字。
        
        参数：
            text: 要检查的文本
        
        返回：
            True如果是二维码说明文字
        """
        # 去除空格和标点符号
        cleaned_text = text.strip().lower()
        
        # 二维码相关的关键词
        qrcode_keywords = [
            'qr code', 'qrcode', 'qr',
            '二维码', '扫码', '扫描',
            'scan', 'scanning',
            'enterprise credit', '企业信用',
            'publicity system', '公示系统',
            'log in', '登录',
            'learn more', '了解更多',
            'registration', '注册',
            'licensing', '许可',
            'regulatory', '监管'
        ]
        
        # 检查是否包含任何关键词
        for keyword in qrcode_keywords:
            if keyword in cleaned_text:
                logger.info(f"检测到二维码说明文字(关键词匹配): '{text[:50]}...', 关键词='{keyword}'")
                return True
        
        return False
