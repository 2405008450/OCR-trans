"""Image processor for image translation system.

This module provides the ImageProcessor class for loading, saving,
and validating images. It handles various image formats and provides
robust error handling for image operations.
"""

import logging
import os
from pathlib import Path
from typing import Optional, Tuple

import cv2
import numpy as np

from src.config import ConfigManager
from src.exceptions import ImageLoadError, ImageSaveError


logger = logging.getLogger(__name__)


class ImageProcessor:
    """Handles image loading, saving, and validation operations.
    
    The ImageProcessor provides:
    - Image loading with format validation
    - Image saving with quality preservation
    - Image validation for array integrity
    - Resolution and channel verification
    
    Attributes:
        config: Configuration manager instance
    """
    
    # Supported image formats
    SUPPORTED_FORMATS = {'.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.webp'}
    
    # Valid channel counts for images
    VALID_CHANNELS = {1, 3, 4}  # Grayscale, BGR, BGRA
    
    # Recommended maximum resolution for OCR processing
    # PaddleOCR recommends short edge around 960 pixels
    DEFAULT_MAX_SHORT_EDGE = 960
    DEFAULT_MAX_LONG_EDGE = 2000
    
    def __init__(self, config: ConfigManager):
        """Initialize the image processor.
        
        Args:
            config: Configuration manager instance
        """
        self.config = config
        
        # Get auto-resize configuration
        self.auto_resize_enabled = config.get('image.auto_resize.enabled', True)
        self.max_short_edge = config.get('image.auto_resize.max_short_edge', self.DEFAULT_MAX_SHORT_EDGE)
        self.max_long_edge = config.get('image.auto_resize.max_long_edge', self.DEFAULT_MAX_LONG_EDGE)
        
        logger.debug(
            f"ImageProcessor initialized (auto_resize={self.auto_resize_enabled}, "
            f"max_short_edge={self.max_short_edge}, max_long_edge={self.max_long_edge})"
        )
    
    def load_image(self, path: str) -> np.ndarray:
        """Load an image from the specified path.
        
        Uses OpenCV to read the image file. Validates the file exists,
        has a supported format, and can be successfully decoded.
        
        Args:
            path: Path to the image file
            
        Returns:
            Image as a numpy array in BGR format
            
        Raises:
            ImageLoadError: If the image cannot be loaded due to:
                - File not found
                - Unsupported format
                - Corrupted or invalid image data
                - Read permission denied
        """
        logger.debug(f"Loading image from: {path}")
        
        # Validate path exists
        if not os.path.exists(path):
            error_msg = f"Image file not found: {path}"
            logger.error(error_msg)
            raise ImageLoadError(error_msg)
        
        # Validate file is not a directory
        if os.path.isdir(path):
            error_msg = f"Path is a directory, not an image file: {path}"
            logger.error(error_msg)
            raise ImageLoadError(error_msg)
        
        # Validate file extension
        file_ext = Path(path).suffix.lower()
        if file_ext not in self.SUPPORTED_FORMATS:
            error_msg = (
                f"Unsupported image format: {file_ext}. "
                f"Supported formats: {', '.join(sorted(self.SUPPORTED_FORMATS))}"
            )
            logger.error(error_msg)
            raise ImageLoadError(error_msg)
        
        # Attempt to load the image
        try:
            image = cv2.imread(path, cv2.IMREAD_COLOR)
        except Exception as e:
            error_msg = f"Failed to read image file: {path}. Error: {str(e)}"
            logger.error(error_msg)
            raise ImageLoadError(error_msg)
        
        # Validate image was successfully loaded
        if image is None:
            error_msg = f"Failed to decode image: {path}. The file may be corrupted or in an unsupported format."
            logger.error(error_msg)
            raise ImageLoadError(error_msg)
        
        # Validate image has valid dimensions
        if image.size == 0:
            error_msg = f"Image has zero size: {path}"
            logger.error(error_msg)
            raise ImageLoadError(error_msg)
        
        # Detect image orientation (portrait vs landscape)
        orientation = self.detect_orientation(image)
        logger.info(f"Image orientation: {orientation}")
        
        # Auto-resize if enabled and image is too large
        if self.auto_resize_enabled:
            image = self._auto_resize_if_needed(image, path)
        
        logger.info(f"Successfully loaded image: {path} (shape: {image.shape}, orientation: {orientation})")
        return image
    
    def save_image(self, image: np.ndarray, path: str, quality: Optional[int] = None) -> None:
        """Save an image to the specified path.
        
        Uses OpenCV to write the image file. Preserves original quality
        by default, or uses the specified quality setting.
        
        Args:
            image: Image as a numpy array
            path: Path where the image should be saved
            quality: Optional JPEG quality (0-100) or PNG compression (0-9).
                    If None, uses default high quality settings.
            
        Raises:
            ImageSaveError: If the image cannot be saved due to:
                - Invalid image array
                - Unsupported output format
                - Write permission denied
                - Disk full or other I/O errors
        """
        logger.debug(f"Saving image to: {path}")
        
        # Validate image array
        if not self.validate_image(image):
            error_msg = "Invalid image array: cannot save"
            logger.error(error_msg)
            raise ImageSaveError(error_msg)
        
        # Validate output format
        file_ext = Path(path).suffix.lower()
        if file_ext not in self.SUPPORTED_FORMATS:
            error_msg = (
                f"Unsupported output format: {file_ext}. "
                f"Supported formats: {', '.join(sorted(self.SUPPORTED_FORMATS))}"
            )
            logger.error(error_msg)
            raise ImageSaveError(error_msg)
        
        # Ensure parent directory exists
        parent_dir = Path(path).parent
        if parent_dir and not parent_dir.exists():
            try:
                parent_dir.mkdir(parents=True, exist_ok=True)
                logger.debug(f"Created directory: {parent_dir}")
            except Exception as e:
                error_msg = f"Failed to create directory: {parent_dir}. Error: {str(e)}"
                logger.error(error_msg)
                raise ImageSaveError(error_msg)
        
        # Set encoding parameters based on format
        encode_params = self._get_encode_params(file_ext, quality)
        
        # Attempt to save the image.
        # Use cv2.imencode + Python open() instead of cv2.imwrite() to correctly
        # handle non-ASCII paths on Windows (cv2.imwrite uses ANSI code page
        # which corrupts Unicode filenames such as Chinese characters).
        try:
            success, buf = cv2.imencode(file_ext, image, encode_params)
            if not success or buf is None:
                error_msg = f"Failed to encode image for path: {path}. cv2.imencode returned False."
                logger.error(error_msg)
                raise ImageSaveError(error_msg)
            with open(path, "wb") as f:
                f.write(buf.tobytes())
        except ImageSaveError:
            raise
        except Exception as e:
            error_msg = f"Failed to write image file: {path}. Error: {str(e)}"
            logger.error(error_msg)
            raise ImageSaveError(error_msg)
        
        logger.info(f"Successfully saved image: {path}")
    
    def validate_image(self, image: np.ndarray) -> bool:
        """Validate that an image array is valid.
        
        Checks:
        - Array is a numpy ndarray
        - Array has valid dimensions (2D or 3D)
        - Array has valid channel count (1, 3, or 4)
        - Array has non-zero size
        - Array has valid dtype (uint8)
        
        Args:
            image: Image array to validate
            
        Returns:
            True if the image is valid, False otherwise
        """
        # Check type
        if not isinstance(image, np.ndarray):
            logger.debug("Image validation failed: not a numpy array")
            return False
        
        # Check dimensions
        if image.ndim not in (2, 3):
            logger.debug(f"Image validation failed: invalid dimensions {image.ndim}")
            return False
        
        # Check size
        if image.size == 0:
            logger.debug("Image validation failed: zero size")
            return False
        
        # Check channel count for 3D arrays
        if image.ndim == 3:
            channels = image.shape[2]
            if channels not in self.VALID_CHANNELS:
                logger.debug(f"Image validation failed: invalid channel count {channels}")
                return False
        
        # Check dtype
        if image.dtype != np.uint8:
            logger.debug(f"Image validation failed: invalid dtype {image.dtype}")
            return False
        
        # Check for valid dimensions (width and height > 0)
        height, width = image.shape[:2]
        if height <= 0 or width <= 0:
            logger.debug(f"Image validation failed: invalid dimensions {width}x{height}")
            return False
        
        return True
    
    def get_image_info(self, image: np.ndarray) -> dict:
        """Get information about an image.
        
        Args:
            image: Image array
            
        Returns:
            Dictionary containing image information:
                - width: Image width in pixels
                - height: Image height in pixels
                - channels: Number of color channels
                - dtype: Data type of the array
                - size: Total number of elements
        """
        if not isinstance(image, np.ndarray):
            return {}
        
        height, width = image.shape[:2]
        channels = image.shape[2] if image.ndim == 3 else 1
        
        return {
            'width': width,
            'height': height,
            'channels': channels,
            'dtype': str(image.dtype),
            'size': image.size
        }
    
    def get_resolution(self, image: np.ndarray) -> Tuple[int, int]:
        """Get the resolution (width, height) of an image.
        
        Args:
            image: Image array
            
        Returns:
            Tuple of (width, height) in pixels
        """
        if not isinstance(image, np.ndarray) or image.ndim < 2:
            return (0, 0)
        
        height, width = image.shape[:2]
        return (width, height)
    
    def copy_image(self, image: np.ndarray) -> np.ndarray:
        """Create a deep copy of an image.
        
        Args:
            image: Image array to copy
            
        Returns:
            Deep copy of the image array
        """
        return image.copy()
    
    def _get_encode_params(self, file_ext: str, quality: Optional[int]) -> list:
        """Get encoding parameters for the specified format.
        
        Args:
            file_ext: File extension (e.g., '.jpg', '.png')
            quality: Optional quality setting
            
        Returns:
            List of encoding parameters for cv2.imwrite
        """
        params = []
        
        if file_ext in {'.jpg', '.jpeg'}:
            # JPEG quality (0-100, higher is better)
            jpeg_quality = quality if quality is not None else 95
            jpeg_quality = max(0, min(100, jpeg_quality))
            params = [cv2.IMWRITE_JPEG_QUALITY, jpeg_quality]
        
        elif file_ext == '.png':
            # PNG compression (0-9, lower is less compression)
            png_compression = quality if quality is not None else 3
            png_compression = max(0, min(9, png_compression))
            params = [cv2.IMWRITE_PNG_COMPRESSION, png_compression]
        
        elif file_ext == '.webp':
            # WebP quality (0-100, higher is better)
            webp_quality = quality if quality is not None else 95
            webp_quality = max(0, min(100, webp_quality))
            params = [cv2.IMWRITE_WEBP_QUALITY, webp_quality]
        
        return params
    
    def _auto_resize_if_needed(self, image: np.ndarray, path: str) -> np.ndarray:
        """自动缩放图片如果尺寸超过推荐值。
        
        策略：
        1. 检查图片的短边和长边
        2. 如果短边 > max_short_edge 或 长边 > max_long_edge，进行缩放
        3. 保持宽高比，使用高质量的插值算法
        
        Args:
            image: 原始图片
            path: 图片路径（用于日志）
            
        Returns:
            缩放后的图片（如果需要缩放），否则返回原图
        """
        height, width = image.shape[:2]
        short_edge = min(height, width)
        long_edge = max(height, width)
        
        # 检查是否需要缩放
        needs_resize = False
        scale_factor = 1.0
        
        if short_edge > self.max_short_edge:
            # 短边超过限制，按短边缩放
            scale_factor = self.max_short_edge / short_edge
            needs_resize = True
            logger.info(
                f"Image short edge ({short_edge}px) exceeds limit ({self.max_short_edge}px), "
                f"will resize by factor {scale_factor:.2f}"
            )
        
        if long_edge > self.max_long_edge:
            # 长边超过限制，按长边缩放
            long_edge_scale = self.max_long_edge / long_edge
            if long_edge_scale < scale_factor:
                scale_factor = long_edge_scale
                needs_resize = True
                logger.info(
                    f"Image long edge ({long_edge}px) exceeds limit ({self.max_long_edge}px), "
                    f"will resize by factor {scale_factor:.2f}"
                )
        
        if not needs_resize:
            logger.debug(f"Image size is within limits, no resize needed: {width}x{height}")
            return image
        
        # 计算新的尺寸
        new_width = int(width * scale_factor)
        new_height = int(height * scale_factor)
        
        # 使用高质量插值算法进行缩放
        # INTER_AREA 适合缩小图片，能保持较好的质量
        resized_image = cv2.resize(
            image, 
            (new_width, new_height), 
            interpolation=cv2.INTER_AREA
        )
        
        logger.info(
            f"Resized image: {width}x{height} -> {new_width}x{new_height} "
            f"(scale={scale_factor:.2f})"
        )
        
        return resized_image
    
    def detect_orientation(self, image: np.ndarray) -> str:
        """检测图片方向（横版 or 竖版）。
        
        根据宽高比判断：
        - 宽 > 高：横版（landscape）
        - 高 > 宽：竖版（portrait）
        - 宽 ≈ 高：正方形（square）
        
        Args:
            image: 图片数组
            
        Returns:
            'landscape'（横版）、'portrait'（竖版）或 'square'（正方形）
        """
        if not isinstance(image, np.ndarray) or image.ndim < 2:
            return 'unknown'
        
        height, width = image.shape[:2]
        aspect_ratio = width / height if height > 0 else 1.0
        
        # 定义正方形的阈值范围（0.9 ~ 1.1）
        square_threshold_min = 0.9
        square_threshold_max = 1.1
        
        if square_threshold_min <= aspect_ratio <= square_threshold_max:
            orientation = 'square'
        elif aspect_ratio > 1.0:
            orientation = 'landscape'  # 横版
        else:
            orientation = 'portrait'   # 竖版
        
        logger.debug(
            f"Image orientation detected: {orientation} "
            f"(width={width}, height={height}, aspect_ratio={aspect_ratio:.2f})"
        )
        
        return orientation

