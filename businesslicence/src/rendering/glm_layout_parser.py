"""GLM-OCR Layout Parser for seal text recognition.

This module provides the GLMOCRLayoutParser class for calling GLM-OCR's
layout parsing API to enhance seal text recognition.

本模块提供 GLMOCRLayoutParser 类，用于调用 GLM-OCR 的布局解析 API
来增强印章文字识别能力。
"""

import base64
import io
import logging
import time
from dataclasses import dataclass
from typing import List, Dict, Optional, Tuple

import numpy as np
from PIL import Image

from src.config import ConfigManager
from src.exceptions import OCRError

logger = logging.getLogger(__name__)


@dataclass
class SealRegionImage:
    """Seal region image with cropping information.
    
    印章区域图片及裁剪信息。
    
    Attributes:
        image: Cropped image as numpy array
        seal_bbox_in_full: Seal bbox in full image (x1, y1, x2, y2)
        seal_bbox_in_crop: Seal bbox in cropped image (x1, y1, x2, y2)
        crop_offset: Offset of crop region in full image (offset_x, offset_y)
        margin: Margin added during cropping
    """
    image: np.ndarray
    seal_bbox_in_full: Tuple[int, int, int, int]
    seal_bbox_in_crop: Tuple[int, int, int, int]
    crop_offset: Tuple[int, int]
    margin: int


class GLMOCRLayoutParser:
    """GLM-OCR layout parser for seal text recognition.
    
    This class uses GLM-OCR's layout parsing API to identify text
    within and overlapping with seals. It provides enhanced recognition
    for small and blurry text that may be difficult for standard OCR.
    
    此类使用 GLM-OCR 的布局解析 API 来识别印章内和与印章重叠的文字。
    它为标准 OCR 难以识别的小字和模糊文字提供增强识别能力。
    """
    
    def __init__(self, config: ConfigManager):
        """Initialize the layout parser.
        
        Args:
            config: Configuration manager instance
        """
        self.config = config
        self.enabled = config.get(
            'rendering.seal_text_handling.glm_layout_parsing.enabled', 
            False
        )
        self.api_key = config.get('ocr.glm_api_key', None)
        self.timeout = config.get(
            'rendering.seal_text_handling.glm_layout_parsing.timeout', 
            30
        )
        self.fallback_to_existing = config.get(
            'rendering.seal_text_handling.glm_layout_parsing.fallback_to_existing',
            True
        )
        self.max_image_size_mb = config.get(
            'rendering.seal_text_handling.glm_layout_parsing.max_image_size_mb',
            10
        )
        self.client = None
        
        # Performance monitoring: track API call count
        self.api_call_count = 0
        
        logger.info(
            f"GLMOCRLayoutParser initialized: enabled={self.enabled}, "
            f"timeout={self.timeout}s, fallback={self.fallback_to_existing}"
        )
    
    def initialize(self) -> None:
        """Initialize the API client.
        
        This method validates the API key and initializes the ZhipuAI client.
        
        Raises:
            OCRError: If API key is missing or initialization fails
        """
        if not self.enabled:
            logger.info("GLM-OCR layout parsing is disabled")
            return
        
        # Validate API key
        if not self.api_key:
            raise OCRError(
                "GLM-OCR API key not found. "
                "Please set 'ocr.glm_api_key' in your config file to use "
                "GLM-OCR layout parsing for seal text recognition."
            )
        
        if not self.api_key.strip():
            raise OCRError(
                "GLM-OCR API key is empty. "
                "Please provide a valid API key in 'ocr.glm_api_key'."
            )
        
        # Check if API key is still the placeholder
        if self.api_key == "your-api-key-here":
            raise OCRError(
                "GLM-OCR API key is still set to placeholder value. "
                "Please replace 'your-api-key-here' with your actual API key "
                "from https://open.bigmodel.cn/"
            )
        
        try:
            # Import zhipuai SDK
            try:
                from zhipuai import ZhipuAI
            except ImportError:
                raise OCRError(
                    "zhipuai SDK not installed. "
                    "Please install it with: pip install zhipuai>=2.0.0"
                )
            
            # Initialize client
            self.client = ZhipuAI(api_key=self.api_key)
            
            logger.info("✅ GLM-OCR layout parsing API client initialized successfully")
            logger.info(f"   Timeout: {self.timeout}s, Max image size: {self.max_image_size_mb}MB")
            
        except OCRError:
            raise
        except Exception as e:
            raise OCRError(f"Failed to initialize GLM-OCR layout parsing client: {e}")

    def parse_seal_region(
        self,
        image: np.ndarray,
        seal_bbox: Tuple[int, int, int, int]
    ) -> Optional[List[Dict]]:
        """Parse seal region to extract text and layout information.
        
        This is the main entry point for seal text recognition using GLM-OCR.
        It orchestrates the complete workflow:
        1. Extract seal region with margin
        2. Validate image size
        3. Call GLM-OCR API
        4. Parse response and convert coordinates
        
        这是使用 GLM-OCR 进行印章文字识别的主要入口点。
        它协调完整的工作流程：
        1. 提取带边距的印章区域
        2. 验证图片大小
        3. 调用 GLM-OCR API
        4. 解析响应并转换坐标
        
        Args:
            image: Full image as numpy array (H, W, C)
            seal_bbox: Seal bounding box in full image (x1, y1, x2, y2)
            
        Returns:
            List of text regions, each containing:
            - bbox: Bounding box in full image (x1, y1, x2, y2)
            - text: Recognized text content
            - confidence: Confidence score (0.0-1.0)
            - label: Layout label ('text')
            - source: 'glm_layout' to indicate source
            
            Returns None if:
            - GLM-OCR is not enabled
            - API client is not initialized
            - Image validation fails
            - API call fails
            
        Example:
            >>> parser = GLMOCRLayoutParser(config)
            >>> parser.initialize()
            >>> image = load_image("document.jpg")
            >>> seal_bbox = (100, 100, 300, 300)
            >>> regions = parser.parse_seal_region(image, seal_bbox)
            >>> if regions:
            ...     for region in regions:
            ...         print(f"Text: {region['text']}, Bbox: {region['bbox']}")
            ... else:
            ...     # Fall back to existing OCR
            ...     pass
        """
        if not self.enabled or self.client is None:
            logger.debug("GLM-OCR layout parsing not available")
            return None
        
        try:
            logger.info(
                f"Starting seal region parsing for bbox {seal_bbox}"
            )
            
            # Step 1: Extract seal region with margin
            seal_region_image = self._prepare_seal_image(image, seal_bbox)
            
            # Step 2: Validate image size (already done in _call_api, but log here)
            if not self._validate_image_size(seal_region_image.image):
                logger.warning(
                    f"Seal region image exceeds size limit, skipping API call"
                )
                return None
            
            # Step 3: Call GLM-OCR API
            response = self._call_api(seal_region_image)
            
            if response is None:
                logger.warning(
                    "API call failed or returned None, falling back to existing OCR"
                )
                return None
            
            # Step 4: Parse response and convert coordinates
            text_regions = self._parse_response(response, seal_region_image)
            
            if not text_regions:
                logger.warning(
                    "No text regions found in API response"
                )
                return None
            
            logger.info(
                f"✅ Successfully parsed seal region: "
                f"found {len(text_regions)} text regions"
            )
            
            return text_regions
            
        except Exception as e:
            logger.error(
                f"Unexpected error in parse_seal_region: {e}. "
                f"Falling back to existing OCR."
            )
            return None
    
    def _prepare_seal_image(
        self,
        image: np.ndarray,
        seal_bbox: Tuple[int, int, int, int],
        margin_ratio: float = 0.3
    ) -> SealRegionImage:
        """Extract seal region from full image with margin.
        
        This method crops the seal region from the full image, including
        a margin around the seal to capture surrounding text. The margin
        is calculated as a percentage of the seal size (default 30%).
        
        从完整图片中提取印章区域，包含周围的边距以捕获周围的文字。
        边距按印章尺寸的百分比计算（默认 30%）。
        
        Args:
            image: Full image as numpy array (H, W, C)
            seal_bbox: Seal bounding box in full image (x1, y1, x2, y2)
            margin_ratio: Margin as ratio of seal size (default 0.3 = 30%)
            
        Returns:
            SealRegionImage object containing:
            - Cropped image
            - Seal bbox in full image
            - Seal bbox in cropped image
            - Crop offset (for coordinate conversion)
            - Margin size
            
        Example:
            >>> image = np.zeros((1000, 1000, 3), dtype=np.uint8)
            >>> seal_bbox = (400, 400, 600, 600)  # 200x200 seal
            >>> result = parser._prepare_seal_image(image, seal_bbox)
            >>> # Margin = max(200, 200) * 0.3 = 60 pixels
            >>> # Crop region: (340, 340, 660, 660)
            >>> result.margin
            60
        """
        x1, y1, x2, y2 = seal_bbox
        img_height, img_width = image.shape[:2]
        
        # Calculate seal dimensions
        seal_width = x2 - x1
        seal_height = y2 - y1
        
        # Calculate margin (30% of the larger dimension)
        margin = int(max(seal_width, seal_height) * margin_ratio)
        
        # Calculate crop region with boundary checks
        crop_x1 = max(0, x1 - margin)
        crop_y1 = max(0, y1 - margin)
        crop_x2 = min(img_width, x2 + margin)
        crop_y2 = min(img_height, y2 + margin)
        
        # Crop the image
        cropped_image = image[crop_y1:crop_y2, crop_x1:crop_x2].copy()
        
        # Calculate seal bbox in cropped image
        seal_bbox_in_crop = (
            x1 - crop_x1,  # x1 in crop
            y1 - crop_y1,  # y1 in crop
            x2 - crop_x1,  # x2 in crop
            y2 - crop_y1   # y2 in crop
        )
        
        # Record crop offset for coordinate conversion
        crop_offset = (crop_x1, crop_y1)
        
        logger.debug(
            f"Prepared seal region: "
            f"seal_bbox={seal_bbox}, "
            f"margin={margin}, "
            f"crop_region=({crop_x1}, {crop_y1}, {crop_x2}, {crop_y2}), "
            f"cropped_size={cropped_image.shape[:2]}"
        )
        
        return SealRegionImage(
            image=cropped_image,
            seal_bbox_in_full=seal_bbox,
            seal_bbox_in_crop=seal_bbox_in_crop,
            crop_offset=crop_offset,
            margin=margin
        )

    def _encode_image_to_base64(self, image: np.ndarray) -> str:
        """Encode image to base64 string for API transmission.
        
        Converts a numpy array image to PNG format and encodes it as a
        base64 string suitable for sending to the GLM-OCR API.
        
        将 numpy 数组图片转换为 PNG 格式并编码为 base64 字符串，
        用于发送到 GLM-OCR API。
        
        Args:
            image: Image as numpy array (H, W, C) with values 0-255
            
        Returns:
            Base64 encoded string of the PNG image
            
        Raises:
            OCRError: If image encoding fails
            
        Example:
            >>> image = np.zeros((100, 100, 3), dtype=np.uint8)
            >>> base64_str = parser._encode_image_to_base64(image)
            >>> isinstance(base64_str, str)
            True
            >>> len(base64_str) > 0
            True
        """
        try:
            # Convert numpy array to PIL Image
            if image.dtype != np.uint8:
                # Ensure image is in uint8 format
                image = image.astype(np.uint8)
            
            # Handle grayscale images
            if len(image.shape) == 2:
                pil_image = Image.fromarray(image, mode='L')
            elif image.shape[2] == 3:
                pil_image = Image.fromarray(image, mode='RGB')
            elif image.shape[2] == 4:
                pil_image = Image.fromarray(image, mode='RGBA')
            else:
                raise OCRError(f"Unsupported image shape: {image.shape}")
            
            # Encode to PNG in memory
            buffer = io.BytesIO()
            pil_image.save(buffer, format='PNG')
            buffer.seek(0)
            
            # Convert to base64
            base64_bytes = base64.b64encode(buffer.read())
            base64_str = base64_bytes.decode('utf-8')
            
            logger.debug(
                f"Encoded image to base64: "
                f"shape={image.shape}, size={len(base64_str)} chars"
            )
            
            return base64_str
            
        except Exception as e:
            raise OCRError(f"Failed to encode image to base64: {e}")

    def _validate_image_size(self, image: np.ndarray) -> bool:
        """Validate that image size does not exceed 10MB.
        
        Checks if the image size (when encoded as PNG) is within the
        API's size limit of 10MB.
        
        检查图片大小（编码为 PNG 后）是否在 API 的 10MB 限制内。
        
        Args:
            image: Image as numpy array
            
        Returns:
            True if image size is valid (≤ 10MB), False otherwise
            
        Example:
            >>> small_image = np.zeros((100, 100, 3), dtype=np.uint8)
            >>> parser._validate_image_size(small_image)
            True
            >>> # Very large image would return False
        """
        try:
            # Convert to PIL Image
            if image.dtype != np.uint8:
                image = image.astype(np.uint8)
            
            if len(image.shape) == 2:
                pil_image = Image.fromarray(image, mode='L')
            elif image.shape[2] == 3:
                pil_image = Image.fromarray(image, mode='RGB')
            elif image.shape[2] == 4:
                pil_image = Image.fromarray(image, mode='RGBA')
            else:
                logger.warning(f"Unsupported image shape for size validation: {image.shape}")
                return False
            
            # Encode to PNG and check size
            buffer = io.BytesIO()
            pil_image.save(buffer, format='PNG')
            size_bytes = buffer.tell()
            size_mb = size_bytes / (1024 * 1024)
            
            max_size_mb = self.max_image_size_mb
            is_valid = size_mb <= max_size_mb
            
            if not is_valid:
                logger.warning(
                    f"Image size {size_mb:.2f}MB exceeds maximum {max_size_mb}MB"
                )
            else:
                logger.debug(
                    f"Image size {size_mb:.2f}MB is within limit {max_size_mb}MB"
                )
            
            return is_valid
            
        except Exception as e:
            logger.error(f"Failed to validate image size: {e}")
            return False

    def _validate_image_format(self, image: np.ndarray) -> bool:
        """Validate that image format is supported (PDF/JPG/PNG).
        
        Since we receive numpy arrays, we check if the image can be
        successfully encoded to PNG format (which is always supported).
        This method primarily validates the image structure.
        
        由于我们接收的是 numpy 数组，我们检查图片是否可以成功编码为
        PNG 格式（始终支持）。此方法主要验证图片结构。
        
        Args:
            image: Image as numpy array
            
        Returns:
            True if image format is valid, False otherwise
            
        Note:
            For numpy arrays, we always encode to PNG which is supported.
            This method validates the array structure is compatible.
            
        Example:
            >>> valid_image = np.zeros((100, 100, 3), dtype=np.uint8)
            >>> parser._validate_image_format(valid_image)
            True
            >>> invalid_image = np.zeros((100, 100, 5), dtype=np.uint8)
            >>> parser._validate_image_format(invalid_image)
            False
        """
        try:
            # Check if image is a valid numpy array
            if not isinstance(image, np.ndarray):
                logger.warning(f"Image is not a numpy array: {type(image)}")
                return False
            
            # Check dimensions
            if len(image.shape) not in [2, 3]:
                logger.warning(
                    f"Invalid image dimensions: {image.shape}. "
                    f"Expected 2D (grayscale) or 3D (color) array"
                )
                return False
            
            # Check channel count for color images
            if len(image.shape) == 3:
                channels = image.shape[2]
                if channels not in [3, 4]:  # RGB or RGBA
                    logger.warning(
                        f"Invalid number of channels: {channels}. "
                        f"Expected 3 (RGB) or 4 (RGBA)"
                    )
                    return False
            
            # Check if image has valid size
            if image.size == 0:
                logger.warning("Image is empty (size = 0)")
                return False
            
            # Try to convert to PIL Image to ensure compatibility
            if image.dtype != np.uint8:
                image = image.astype(np.uint8)
            
            if len(image.shape) == 2:
                Image.fromarray(image, mode='L')
            elif image.shape[2] == 3:
                Image.fromarray(image, mode='RGB')
            elif image.shape[2] == 4:
                Image.fromarray(image, mode='RGBA')
            
            logger.debug(f"Image format is valid: shape={image.shape}, dtype={image.dtype}")
            return True
            
        except Exception as e:
            logger.warning(f"Image format validation failed: {e}")
            return False

    def _call_api(self, seal_image: SealRegionImage) -> Optional[object]:
        """Call GLM-OCR layout parsing API.
        
        Calls the GLM-OCR layout parsing API with the seal region image.
        Handles timeouts, network errors, and API errors gracefully by
        returning None to trigger fallback mechanism.
        
        调用 GLM-OCR 布局解析 API。优雅地处理超时、网络错误和 API 错误，
        通过返回 None 来触发回退机制。
        
        Args:
            seal_image: SealRegionImage object containing the cropped image
            
        Returns:
            API response object if successful, None if failed
            
        Note:
            This method implements the fallback mechanism by returning None
            on any error. The caller should check for None and fall back to
            existing OCR engine.
            
        Handles:
            - Timeout errors (after self.timeout seconds)
            - Network connection errors
            - API errors (invalid key, rate limits, etc.)
            - Response parsing errors
            
        Example:
            >>> response = parser._call_api(seal_image)
            >>> if response is None:
            ...     # Fall back to existing OCR
            ...     pass
        """
        if self.client is None:
            logger.error(
                "API client not initialized. Cannot call GLM-OCR API. "
                "Falling back to existing OCR."
            )
            return None
        
        try:
            # Validate image format
            if not self._validate_image_format(seal_image.image):
                logger.warning(
                    "Invalid image format detected. Cannot call GLM-OCR API. "
                    "Falling back to existing OCR."
                )
                return None
            
            # Validate image size
            if not self._validate_image_size(seal_image.image):
                logger.warning(
                    f"Image size exceeds {self.max_image_size_mb}MB limit. "
                    f"Cannot call GLM-OCR API. Falling back to existing OCR."
                )
                return None
            
            # Encode image to base64
            try:
                base64_image = self._encode_image_to_base64(seal_image.image)
            except OCRError as e:
                logger.error(
                    f"Failed to encode image to base64: {e}. "
                    f"Falling back to existing OCR."
                )
                return None
            except Exception as e:
                logger.error(
                    f"Unexpected error during image encoding: {e}. "
                    f"Falling back to existing OCR."
                )
                return None
            
            # Record start time for performance monitoring
            start_time = time.time()
            
            # Increment API call count for performance monitoring
            self.api_call_count += 1
            
            logger.info(
                f"📞 Calling GLM-OCR layout parsing API "
                f"(call #{self.api_call_count}, timeout={self.timeout}s, "
                f"image_size={seal_image.image.shape})..."
            )
            
            # Call API with timeout handling
            # We use a manual timeout wrapper since zhipuai SDK may not support timeout directly
            import signal
            
            def timeout_handler(signum, frame):
                raise TimeoutError(f"API call exceeded timeout of {self.timeout}s")
            
            # Set up timeout (only on Unix-like systems)
            timeout_supported = hasattr(signal, 'SIGALRM')
            if timeout_supported:
                old_handler = signal.signal(signal.SIGALRM, timeout_handler)
                signal.alarm(self.timeout)
            
            try:
                # 使用 GLM-OCR 版面分析 API 识别印章文字
                logger.info(f"📞 调用 GLM-OCR 版面分析 API 识别印章文字...")
                
                # GLM-OCR API 使用 POST 请求到 /layout_parsing
                # 参考文档: https://docs.z.ai/api-reference/tools/layout_parsing
                # 注意：SDK 会自动添加 /api/paas/v4 前缀
                # file 参数需要是 data URI 格式: data:image/png;base64,<base64_string>
                data_uri = f"data:image/png;base64,{base64_image}"
                
                response = self.client.post(
                    "/layout_parsing",
                    body={
                        "model": "glm-ocr",
                        "file": data_uri  # data URI 格式的图片
                    },
                    cast_type=object  # 返回原始对象
                )
                
                # 检查响应
                if response is None:
                    logger.warning("GLM-OCR 返回空响应")
                    return None
                
                # 添加详细日志查看响应结构
                logger.info(f"📋 响应类型: {type(response)}")
                if isinstance(response, dict):
                    logger.info(f"📋 响应键: {list(response.keys())}")
                    # 打印部分响应内容（避免太长）
                    for key in ['id', 'model', 'created', 'layout_details', 'md_results']:
                        if key in response:
                            value = response[key]
                            if isinstance(value, str) and len(value) > 100:
                                logger.info(f"   {key}: {value[:100]}...")
                            elif isinstance(value, list):
                                logger.info(f"   {key}: list with {len(value)} items")
                            else:
                                logger.info(f"   {key}: {value}")
                else:
                    logger.info(f"📋 响应属性: {dir(response) if hasattr(response, '__dir__') else 'N/A'}")
                
                # 解析响应 - layout_details 是一个嵌套数组（每页一个数组）
                layout_details = None
                if isinstance(response, dict):
                    layout_details = response.get('layout_details')
                elif hasattr(response, 'layout_details'):
                    layout_details = response.layout_details
                
                if layout_details:
                    logger.info(f"✅ 找到 layout_details")
                    logger.info(f"   类型: {type(layout_details)}")
                    logger.info(f"   长度: {len(layout_details)}")
                    # layout_details 是一个二维数组，第一维是页码
                    total_elements = sum(len(page) if isinstance(page, list) else 0 
                                       for page in layout_details)
                    logger.info(f"✅ GLM-OCR 版面分析成功，识别到 {total_elements} 个元素")
                    
                    # 记录识别的文字内容（遍历所有页）
                    element_count = 0
                    for page_idx, page_elements in enumerate(layout_details):
                        if not isinstance(page_elements, list):
                            logger.warning(f"   页 {page_idx+1} 不是列表: {type(page_elements)}")
                            continue
                        logger.info(f"   页 {page_idx+1}: {len(page_elements)} 个元素")
                        for element in page_elements:
                            # 处理字典或对象
                            if isinstance(element, dict):
                                label = element.get('label', '')
                                content = element.get('content', '')
                            else:
                                label = getattr(element, 'label', '')
                                content = getattr(element, 'content', '')
                            
                            if label == 'text':
                                element_count += 1
                                logger.info(f"  元素 {element_count} (页{page_idx+1}): {content}")
                    
                    return response
                else:
                    logger.warning("layout_details 为空或不存在")
                    return None
                
            except Exception as e:
                logger.error(
                    f"API call failed: {type(e).__name__}: {e}. "
                    f"Falling back to existing OCR."
                )
                return None
            finally:
                # Cancel timeout alarm
                if timeout_supported:
                    signal.alarm(0)
                    signal.signal(signal.SIGALRM, old_handler)
            
            # Record end time
            elapsed_time = time.time() - start_time
            
            logger.info(
                f"✅ GLM-OCR API call #{self.api_call_count} completed successfully "
                f"in {elapsed_time:.2f}s"
            )
            
            # Check response status
            if hasattr(response, 'status'):
                if response.status != 'succeeded':
                    logger.error(
                        f"API returned error status: {response.status}. "
                        f"Falling back to existing OCR."
                    )
                    if hasattr(response, 'error'):
                        error_details = response.error
                        logger.error(f"Error details: {error_details}")
                    return None
            
            # Validate response has expected structure
            if not hasattr(response, 'layout_details') and not hasattr(response, 'data'):
                logger.warning(
                    "API response missing expected fields (layout_details or data). "
                    "Falling back to existing OCR."
                )
                return None
            
            return response
            
        except TimeoutError as e:
            logger.warning(
                f"⏱️ API call timed out after {self.timeout}s: {e}. "
                f"Falling back to existing OCR."
            )
            return None
            
        except ConnectionError as e:
            logger.warning(
                f"🌐 Network connection error: {type(e).__name__}: {e}. "
                f"Falling back to existing OCR."
            )
            return None
        
        except OSError as e:
            # OSError includes network-related errors
            logger.warning(
                f"🌐 Network/IO error: {type(e).__name__}: {e}. "
                f"Falling back to existing OCR."
            )
            return None
            
        except Exception as e:
            # Catch all other errors (API errors, authentication errors, etc.)
            error_msg = str(e)
            error_type = type(e).__name__
            
            # Check for common error types in the error message
            if 'timeout' in error_msg.lower():
                logger.warning(
                    f"⏱️ API call timed out: {error_type}: {e}. "
                    f"Falling back to existing OCR."
                )
            elif 'unauthorized' in error_msg.lower() or '401' in error_msg:
                logger.error(
                    f"🔑 API authentication failed (invalid API key): {error_type}: {e}. "
                    f"Please check your API key configuration. "
                    f"Falling back to existing OCR."
                )
            elif 'forbidden' in error_msg.lower() or '403' in error_msg:
                logger.error(
                    f"🔑 API access forbidden: {error_type}: {e}. "
                    f"Please check your API key permissions. "
                    f"Falling back to existing OCR."
                )
            elif 'rate limit' in error_msg.lower() or '429' in error_msg:
                logger.warning(
                    f"⚠️ API rate limit exceeded: {error_type}: {e}. "
                    f"Please wait before retrying. "
                    f"Falling back to existing OCR."
                )
            elif 'network' in error_msg.lower() or 'connection' in error_msg.lower():
                logger.warning(
                    f"🌐 Network error: {error_type}: {e}. "
                    f"Falling back to existing OCR."
                )
            elif 'not found' in error_msg.lower() or '404' in error_msg:
                logger.error(
                    f"❌ API endpoint not found: {error_type}: {e}. "
                    f"Please check your zhipuai SDK version. "
                    f"Falling back to existing OCR."
                )
            elif 'server error' in error_msg.lower() or '500' in error_msg or '502' in error_msg or '503' in error_msg:
                logger.warning(
                    f"🔧 API server error: {error_type}: {e}. "
                    f"The service may be temporarily unavailable. "
                    f"Falling back to existing OCR."
                )
            else:
                logger.error(
                    f"❌ API call failed with unexpected error: {error_type}: {e}. "
                    f"Falling back to existing OCR."
                )
            
            return None

    def _is_date_text(self, text: str) -> bool:
        """Check if text contains date information.
        
        Detects Chinese date formats like:
        - 2022年12月16日
        - 2016年04月13日
        - 2022-12-16
        - 2022/12/16
        - 2022.12.16
        
        检测文字是否包含日期信息。
        
        Args:
            text: Text content to check
            
        Returns:
            True if text appears to be a date, False otherwise
            
        Example:
            >>> parser._is_date_text("2022年12月16日")
            True
            >>> parser._is_date_text("佛山市顺德区市场监督管理局")
            False
        """
        if not text:
            return False
        
        text = text.strip()
        
        # 首先排除英文单词（避免误判"April"、"Registration"等）
        # First exclude English words (avoid false positives like "April", "Registration", etc.)
        import re
        
        # 如果包含英文字母（除了日期分隔符），则不是日期
        # If contains English letters (except date separators), it's not a date
        if re.search(r'[a-zA-Z]', text):
            return False
        
        # 检查是否包含中文日期关键字
        # Check for Chinese date keywords
        has_year = '年' in text
        has_month = '月' in text
        has_day = '日' in text
        
        # 如果包含"年月日"组合，很可能是日期
        # If contains "year-month-day" combination, likely a date
        if (has_year and has_month) or (has_year and has_day) or (has_month and has_day):
            return True
        
        # 单独的"年"、"月"、"日"也可能是日期的一部分
        # Single "year", "month", "day" characters might be part of a date
        if has_year or has_month or has_day:
            return True
        
        # 检查是否包含数字分隔符格式的日期
        # Check for numeric date formats with separators
        import re
        
        # 匹配格式：YYYY-MM-DD, YYYY/MM/DD, YYYY.MM.DD
        # Match formats: YYYY-MM-DD, YYYY/MM/DD, YYYY.MM.DD
        date_pattern = r'\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2}'
        if re.search(date_pattern, text):
            return True
        
        # 匹配格式：YYYYMMDD（8位数字）
        # Match format: YYYYMMDD (8 digits)
        if re.match(r'^\d{8}$', text):
            return True
        
        # 匹配4位年份数字（如"2018"）
        # Match 4-digit year numbers (like "2018")
        if re.match(r'^\d{4}$', text):
            return True
        
        # 匹配1-2位数字（可能是月份或日期，如"04"、"17"）
        # Match 1-2 digit numbers (might be month or day, like "04", "17")
        if re.match(r'^\d{1,2}$', text):
            return True
        
        return False
    
    def _parse_response(
        self,
        response: object,
        seal_region_image: SealRegionImage
    ) -> List[Dict]:
        """Parse API response to extract text regions.
        
        Extracts layout_details from the API response, filters for text
        elements (label='text'), and converts coordinates from the cropped
        image coordinate system to the full image coordinate system.
        
        Only includes text regions that are inside or overlapping with the seal.
        
        从 API 响应中提取 layout_details，过滤文字元素（label='text'），
        并将坐标从裁剪图片坐标系转换为完整图片坐标系。
        
        只包含在印章内或与印章重叠的文字区域。
        
        Args:
            response: API response object from GLM-OCR
            seal_region_image: SealRegionImage with crop offset information
            
        Returns:
            List of text regions, each containing:
            - bbox: Bounding box in full image (x1, y1, x2, y2)
            - text: Recognized text content
            - confidence: Confidence score (0.0-1.0)
            - label: Layout label (e.g., 'text')
            - source: 'glm_layout' to indicate source
            
        Example:
            >>> # Response with layout_details
            >>> response = MockResponse(layout_details=[
            ...     {'label': 'text', 'bbox_2d': [10, 20, 50, 40], 
            ...      'content': 'Hello', 'confidence': 0.95},
            ...     {'label': 'image', 'bbox_2d': [60, 70, 100, 110]},
            ... ])
            >>> seal_image = SealRegionImage(
            ...     image=np.zeros((100, 100, 3)),
            ...     seal_bbox_in_full=(100, 100, 200, 200),
            ...     seal_bbox_in_crop=(50, 50, 150, 150),
            ...     crop_offset=(50, 50),
            ...     margin=50
            ... )
            >>> regions = parser._parse_response(response, seal_image)
            >>> len(regions)
            1
            >>> regions[0]['text']
            'Hello'
            >>> regions[0]['bbox']  # Converted to full image coords
            (60, 70, 100, 90)
        """
        text_regions = []
        
        try:
            # Validate response is not None
            if response is None:
                logger.error("Response is None, cannot parse")
                return []
            
            # Extract layout_details from response
            layout_details = None
            
            # Try different possible attribute names
            try:
                if hasattr(response, 'layout_details'):
                    layout_details = response.layout_details
                elif hasattr(response, 'data'):
                    data = response.data
                    if hasattr(data, 'layout_details'):
                        layout_details = data.layout_details
                    elif isinstance(data, dict) and 'layout_details' in data:
                        layout_details = data['layout_details']
                elif isinstance(response, dict):
                    if 'layout_details' in response:
                        layout_details = response['layout_details']
                    elif 'data' in response:
                        data = response['data']
                        if isinstance(data, dict) and 'layout_details' in data:
                            layout_details = data['layout_details']
            except Exception as e:
                logger.error(
                    f"Error accessing response attributes: {type(e).__name__}: {e}"
                )
                return []
            
            if layout_details is None:
                logger.warning(
                    "Could not find layout_details in API response. "
                    "Response type: " + str(type(response)) + ", "
                    "Available attributes: " + 
                    str(dir(response) if hasattr(response, '__dir__') else 'N/A')
                )
                return []
            
            # layout_details 是一个二维数组（每页一个数组）
            # 我们需要展平它来处理所有元素
            all_elements = []
            try:
                if isinstance(layout_details, list):
                    for page_elements in layout_details:
                        if isinstance(page_elements, list):
                            all_elements.extend(page_elements)
                        else:
                            # 如果不是嵌套数组，直接添加
                            all_elements.append(page_elements)
                else:
                    logger.warning(
                        f"layout_details is not a list: {type(layout_details)}"
                    )
                    return []
            except Exception as e:
                logger.error(
                    f"Error flattening layout_details: {type(e).__name__}: {e}"
                )
                return []
            
            if len(all_elements) == 0:
                logger.info("API response contains no layout elements")
                return []
            
            # Get crop offset for coordinate conversion
            try:
                offset_x, offset_y = seal_region_image.crop_offset
                seal_bbox_in_crop = seal_region_image.seal_bbox_in_crop
            except Exception as e:
                logger.error(
                    f"Failed to get crop offset or seal bbox: {type(e).__name__}: {e}"
                )
                return []
            
            logger.debug(
                f"Parsing response with {len(all_elements)} layout elements, "
                f"crop_offset=({offset_x}, {offset_y}), "
                f"seal_bbox_in_crop={seal_bbox_in_crop}"
            )
            
            # Filter and extract text elements
            skipped_count = 0
            for idx, element in enumerate(all_elements):
                try:
                    # Get element attributes (handle both dict and object)
                    if isinstance(element, dict):
                        label = element.get('label', '')
                        bbox_2d = element.get('bbox_2d', None)
                        content = element.get('content', '')
                        confidence = element.get('confidence', 1.0)
                    else:
                        label = getattr(element, 'label', '')
                        bbox_2d = getattr(element, 'bbox_2d', None)
                        content = getattr(element, 'content', '')
                        confidence = getattr(element, 'confidence', 1.0)
                    
                    # Filter for text elements only
                    if label != 'text':
                        continue
                    
                    # Validate bbox_2d
                    if bbox_2d is None:
                        logger.warning(
                            f"Element {idx}: bbox_2d is None, skipping"
                        )
                        skipped_count += 1
                        continue
                    
                    if not isinstance(bbox_2d, (list, tuple)) or len(bbox_2d) < 4:
                        logger.warning(
                            f"Element {idx}: Invalid bbox_2d format: {bbox_2d}, skipping"
                        )
                        skipped_count += 1
                        continue
                    
                    # Extract bbox coordinates (in cropped image)
                    try:
                        x1_crop, y1_crop, x2_crop, y2_crop = bbox_2d[:4]
                        # Ensure coordinates are numeric
                        x1_crop = float(x1_crop)
                        y1_crop = float(y1_crop)
                        x2_crop = float(x2_crop)
                        y2_crop = float(y2_crop)
                    except (ValueError, TypeError) as e:
                        logger.warning(
                            f"Element {idx}: Failed to parse bbox coordinates: {bbox_2d}, error: {e}"
                        )
                        skipped_count += 1
                        continue
                    
                    # Validate bbox coordinates
                    if x2_crop <= x1_crop or y2_crop <= y1_crop:
                        logger.warning(
                            f"Element {idx}: Invalid bbox dimensions: "
                            f"({x1_crop}, {y1_crop}, {x2_crop}, {y2_crop}), skipping"
                        )
                        skipped_count += 1
                        continue
                    
                    # Convert to full image coordinates
                    x1_full = int(x1_crop + offset_x)
                    y1_full = int(y1_crop + offset_y)
                    x2_full = int(x2_crop + offset_x)
                    y2_full = int(y2_crop + offset_y)
                    
                    # 检查文字区域是否在印章内或与印章重叠
                    # Check if text region is inside or overlapping with seal
                    seal_x1, seal_y1, seal_x2, seal_y2 = seal_region_image.seal_bbox_in_full
                    
                    # 计算重叠区域
                    overlap_x1 = max(x1_full, seal_x1)
                    overlap_y1 = max(y1_full, seal_y1)
                    overlap_x2 = min(x2_full, seal_x2)
                    overlap_y2 = min(y2_full, seal_y2)
                    
                    # 检查是否有重叠
                    has_overlap = (overlap_x1 < overlap_x2 and overlap_y1 < overlap_y2)
                    
                    # 计算文字中心点
                    text_center_x = (x1_full + x2_full) / 2
                    text_center_y = (y1_full + y2_full) / 2
                    
                    # 检查文字中心是否在印章内
                    center_in_seal = (seal_x1 <= text_center_x <= seal_x2 and 
                                    seal_y1 <= text_center_y <= seal_y2)
                    
                    # 计算文字大小
                    text_width = x2_full - x1_full
                    text_height = y2_full - y1_full
                    text_area = text_width * text_height
                    
                    # 计算重叠比例
                    overlap_ratio = 0.0
                    if has_overlap:
                        overlap_area = (overlap_x2 - overlap_x1) * (overlap_y2 - overlap_y1)
                        overlap_ratio = overlap_area / text_area if text_area > 0 else 0
                    
                    # 检测是否为日期文字或数字
                    # Detect if this is date text or numeric text
                    is_date_text = self._is_date_text(content)
                    is_numeric = content.strip().isdigit()  # 纯数字（如"2022"）
                    
                    # 如果是日期文字或数字，检查是否在印章附近
                    # If it's date text or numeric, check if it's near the seal
                    is_near_seal = False
                    if is_date_text or is_numeric:
                        # 检查文字是否在印章下方附近（垂直距离 < 100像素，扩大范围以支持竖版营业执照）
                        # Check if text is near below the seal (vertical distance < 100 pixels, expanded for vertical licenses)
                        vertical_distance = y1_full - seal_y2  # 文字在印章下方
                        is_below_seal = (vertical_distance >= 0 and vertical_distance < 100)
                        
                        # 检查文字是否在印章上方附近（垂直距离 < 100像素）
                        # Check if text is near above the seal (vertical distance < 100 pixels)
                        vertical_distance_above = seal_y1 - y2_full  # 文字在印章上方
                        is_above_seal = (vertical_distance_above >= 0 and vertical_distance_above < 100)
                        
                        # 检查水平位置是否接近（中心点水平距离 < 印章宽度 * 2.0，扩大范围）
                        # Check if horizontal position is close (center horizontal distance < seal width * 2.0, expanded)
                        seal_center_x = (seal_x1 + seal_x2) / 2
                        horizontal_distance = abs(text_center_x - seal_center_x)
                        seal_width = seal_x2 - seal_x1
                        is_horizontally_aligned = (horizontal_distance < seal_width * 2.0)
                        
                        is_near_seal = ((is_below_seal or is_above_seal) and is_horizontally_aligned)
                        
                        logger.debug(
                            f"Element {idx}: Date/numeric text detected: '{content}', "
                            f"is_date_text={is_date_text}, is_numeric={is_numeric}, "
                            f"is_below_seal={is_below_seal}, is_above_seal={is_above_seal}, "
                            f"is_horizontally_aligned={is_horizontally_aligned}, "
                            f"is_near_seal={is_near_seal}, "
                            f"vertical_distance={vertical_distance if is_below_seal else vertical_distance_above:.1f}px, "
                            f"horizontal_distance={horizontal_distance:.1f}px"
                        )
                    
                    # 保留条件：
                    # 1. 重叠比例 > 60%（大部分在印章内）
                    # 2. 或者文字中心在印章内
                    # 3. 或者是日期文字/数字且有任何重叠或在印章附近
                    should_keep = (
                        overlap_ratio >= 0.6 or 
                        center_in_seal or 
                        ((is_date_text or is_numeric) and (has_overlap or is_near_seal))
                    )
                    
                    if not should_keep:
                        logger.debug(
                            f"Element {idx}: Text region does not meet criteria "
                            f"(overlap={overlap_ratio:.2%}, center_in_seal={center_in_seal}, "
                            f"is_date_text={is_date_text}, is_numeric={is_numeric}, "
                            f"has_overlap={has_overlap}, is_near_seal={is_near_seal}), "
                            f"skipping. Text: '{content[:30]}...'"
                        )
                        skipped_count += 1
                        continue
                    
                    logger.debug(
                        f"Element {idx}: Text region meets criteria "
                        f"(overlap={overlap_ratio:.2%}, center_in_seal={center_in_seal}, "
                        f"is_date_text={is_date_text}, is_numeric={is_numeric}, "
                        f"has_overlap={has_overlap}, is_near_seal={is_near_seal}), "
                        f"keeping. Text: '{content[:30]}...'"
                    )
                    
                    # Validate confidence
                    try:
                        confidence = float(confidence)
                        if not (0.0 <= confidence <= 1.0):
                            logger.debug(
                                f"Element {idx}: Confidence {confidence} out of range [0,1], clamping"
                            )
                            confidence = max(0.0, min(1.0, confidence))
                    except (ValueError, TypeError):
                        logger.debug(
                            f"Element {idx}: Invalid confidence value: {confidence}, using 1.0"
                        )
                        confidence = 1.0
                    
                    # Create text region
                    text_region = {
                        'bbox': (x1_full, y1_full, x2_full, y2_full),
                        'text': str(content) if content else '',
                        'confidence': confidence,
                        'label': label,
                        'source': 'glm_layout'
                    }
                    
                    text_regions.append(text_region)
                    
                    logger.debug(
                        f"Extracted text region {len(text_regions)}: "
                        f"bbox={text_region['bbox']}, "
                        f"text='{content[:20] if content else ''}...', "
                        f"confidence={confidence:.2f}"
                    )
                    
                except Exception as e:
                    logger.warning(
                        f"Failed to process element {idx}: {type(e).__name__}: {e}, skipping"
                    )
                    skipped_count += 1
                    continue
            
            if skipped_count > 0:
                logger.info(
                    f"Skipped {skipped_count} invalid elements during parsing"
                )
            
            logger.info(
                f"✅ Successfully parsed {len(text_regions)} text regions from API response "
                f"({len(all_elements)} total elements, {skipped_count} skipped)"
            )
            
            return text_regions
            
        except Exception as e:
            logger.error(
                f"Unexpected error parsing API response: {type(e).__name__}: {e}. "
                f"Returning empty list."
            )
            import traceback
            logger.debug(f"Traceback: {traceback.format_exc()}")
            return []
