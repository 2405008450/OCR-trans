"""GLM-OCR API Adapter - Wrapper for GLM-OCR cloud API.

This module provides an adapter for GLM-OCR API that implements the
BaseOCREngine interface.

本模块提供 GLM-OCR API 的适配器，实现 BaseOCREngine 接口。
"""

import logging
import base64
from typing import List, Tuple
from io import BytesIO

import numpy as np
from PIL import Image

from src.ocr.base_ocr_engine import BaseOCREngine
from src.exceptions import OCRError

logger = logging.getLogger(__name__)


class GLMOCRAPIAdapter(BaseOCREngine):
    """Adapter for GLM-OCR cloud API.
    
    This adapter uses ZhipuAI's cloud API for OCR processing.
    Requires an API key from https://open.bigmodel.cn/
    
    此适配器使用智谱 AI 的云端 API 进行 OCR 处理。
    需要从 https://open.bigmodel.cn/ 获取 API Key。
    """
    
    def __init__(self, config):
        """Initialize GLM-OCR API adapter.
        
        Args:
            config: Configuration manager instance
        """
        super().__init__(config)
        self.engine_name = "GLM-OCR-API"
        self._client = None
        self._api_key = None
    
    def initialize(self) -> None:
        """Initialize GLM-OCR API client.
        
        Raises:
            OCRError: If initialization fails
        """
        try:
            # Get API key from config
            self._api_key = self.config.get('ocr.glm_api_key', None)
            
            if not self._api_key:
                raise OCRError(
                    "GLM-OCR API key not found. "
                    "Please set 'ocr.glm_api_key' in your config file."
                )
            
            # Import zai-sdk (for OCR service API)
            try:
                from zai import ZhipuAiClient
            except ImportError:
                raise OCRError(
                    "zai-sdk not installed. "
                    "Please install it with: pip install zai-sdk"
                )
            
            # Initialize client
            self._client = ZhipuAiClient(api_key=self._api_key)
            
            logger.info("GLM-OCR API: Initialized successfully")
            logger.info("GLM-OCR API: Using OCR Service API (can recognize all text including bottom text)")
            
        except OCRError:
            raise
        except Exception as e:
            raise OCRError(f"Failed to initialize GLM-OCR API: {e}")
    
    def detect_and_recognize(self, image: np.ndarray) -> List[Tuple]:
        """Detect and recognize text using GLM-OCR Service API.
        
        Args:
            image: Input image as numpy array
            
        Returns:
            List of tuples (box_points, text, confidence)
            
        Raises:
            OCRError: If API call fails
        """
        if self._client is None:
            raise OCRError("GLM-OCR API not initialized. Call initialize() first.")
        
        try:
            # Save image to temporary file
            import tempfile
            import os
            
            # Convert numpy array to PIL Image
            if len(image.shape) == 3 and image.shape[2] == 3:
                image_rgb = image[:, :, ::-1]  # BGR to RGB
            else:
                image_rgb = image
            
            pil_image = Image.fromarray(image_rgb)
            
            # Save to temporary file
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_file:
                tmp_path = tmp_file.name
                pil_image.save(tmp_path, format='PNG')
            
            try:
                # Call GLM-OCR Service API using zai-sdk
                logger.debug("GLM-OCR API: Sending request to OCR Service API...")
                
                with open(tmp_path, 'rb') as f:
                    response = self._client.ocr.handwriting_ocr(
                        file=f,
                        tool_type="hand_write",
                        language_type="CHN_ENG",
                        probability=True
                    )
                
                # Check response status
                status = response.get('status')
                if status != 'succeeded':
                    message = response.get('message', 'Unknown error')
                    raise OCRError(f"GLM-OCR Service API failed: {message}")
                
                # Parse response
                ocr_results = self._parse_response(response)
                
                logger.debug(f"GLM-OCR API: Detected {len(ocr_results)} text regions")
                return ocr_results
                
            finally:
                # Clean up temporary file
                if os.path.exists(tmp_path):
                    os.unlink(tmp_path)
            
        except OCRError:
            raise
        except Exception as e:
            raise OCRError(f"GLM-OCR API call failed: {e}")
    
    def _image_to_base64(self, image: np.ndarray) -> str:
        """Convert numpy image to base64 string.
        
        Args:
            image: Input image as numpy array (BGR format)
            
        Returns:
            Base64 encoded image string
        """
        # Convert BGR to RGB
        if len(image.shape) == 3 and image.shape[2] == 3:
            image_rgb = image[:, :, ::-1]
        else:
            image_rgb = image
        
        # Convert to PIL Image
        pil_image = Image.fromarray(image_rgb)
        
        # Convert to base64
        buffered = BytesIO()
        pil_image.save(buffered, format="PNG")
        img_str = base64.b64encode(buffered.getvalue()).decode()
        
        return img_str
    
    def _parse_response(self, response: dict) -> List[Tuple]:
        """Parse GLM-OCR Service API response.
        
        Args:
            response: API response dictionary from OCR Service API
            
        Returns:
            List of tuples (box_points, text, confidence)
        """
        ocr_results = []
        
        try:
            # GLM-OCR Service API returns words_result
            words_result = response.get('words_result', [])
            
            if not words_result:
                logger.warning("GLM-OCR API: No words_result in response")
                return ocr_results
            
            for item in words_result:
                # Extract location (left, top, width, height)
                location = item.get('location', {})
                left = location.get('left', 0)
                top = location.get('top', 0)
                width = location.get('width', 0)
                height = location.get('height', 0)
                
                # Convert to corner points format (required by our system)
                box_points = [
                    [left, top],                    # top-left
                    [left + width, top],            # top-right
                    [left + width, top + height],   # bottom-right
                    [left, top + height]            # bottom-left
                ]
                
                # Extract text content
                text = item.get('words', '').strip()
                if not text:
                    continue
                
                # Extract confidence score
                probability = item.get('probability', {})
                confidence = probability.get('average', 1.0) if probability else 1.0
                
                ocr_results.append((box_points, text, confidence))
            
            logger.debug(f"GLM-OCR API: Parsed {len(ocr_results)} text regions from response")
        
        except Exception as e:
            logger.error(f"GLM-OCR API: Failed to parse response: {e}")
            logger.debug(f"Response keys: {response.keys() if isinstance(response, dict) else 'not a dict'}")
        
        return ocr_results
    
    def get_engine_info(self) -> dict:
        """Get GLM-OCR API engine information.
        
        Returns:
            Dictionary with engine information
        """
        return {
            "name": "GLM-OCR-API",
            "version": "cloud",
            "type": "api",
            "capabilities": ["text_detection", "text_recognition", "cloud_based"]
        }
    
    def supports_batch_processing(self) -> bool:
        """GLM-OCR API may support batch processing.
        
        Returns:
            False (process one image at a time to avoid rate limits)
        """
        return False
    
    def cleanup(self) -> None:
        """Clean up API client resources."""
        if self._client:
            # Close any open connections if needed
            self._client = None
            logger.debug("GLM-OCR API: Cleaned up resources")
