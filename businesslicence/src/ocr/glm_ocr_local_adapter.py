"""GLM-OCR Local Adapter - Wrapper for local GLM-OCR model.

This module provides an adapter for locally deployed GLM-OCR model
that implements the BaseOCREngine interface.

本模块提供本地部署的 GLM-OCR 模型适配器，实现 BaseOCREngine 接口。
"""

import logging
from typing import List, Tuple
from pathlib import Path

import numpy as np
from PIL import Image

from src.ocr.base_ocr_engine import BaseOCREngine
from src.exceptions import OCRError

logger = logging.getLogger(__name__)


class GLMOCRLocalAdapter(BaseOCREngine):
    """Adapter for locally deployed GLM-OCR model.
    
    This adapter loads and runs GLM-OCR model locally without
    requiring internet connection or API keys.
    
    此适配器在本地加载和运行 GLM-OCR 模型，无需互联网连接或 API Key。
    """
    
    def __init__(self, config):
        """Initialize GLM-OCR local adapter.
        
        Args:
            config: Configuration manager instance
        """
        super().__init__(config)
        self.engine_name = "GLM-OCR-Local"
        self._model = None
        self._processor = None
        self._device = None
    
    def initialize(self) -> None:
        """Initialize local GLM-OCR model.
        
        Raises:
            OCRError: If model loading fails
        """
        try:
            # Import required libraries
            try:
                import torch
                from transformers import AutoModel, AutoProcessor
            except ImportError as e:
                raise OCRError(
                    f"Required libraries not installed: {e}\n"
                    "Please install with: pip install torch transformers"
                )
            
            # Get model path from config
            model_path = self.config.get('ocr.glm_model_path', './models/glm-ocr')
            model_path = Path(model_path)
            
            if not model_path.exists():
                raise OCRError(
                    f"GLM-OCR model not found at: {model_path}\n"
                    "Please download the model first or update 'ocr.glm_model_path' in config."
                )
            
            # Determine device (GPU or CPU)
            if torch.cuda.is_available():
                self._device = torch.device('cuda')
                logger.info("GLM-OCR Local: Using GPU for inference")
            else:
                self._device = torch.device('cpu')
                logger.warning("GLM-OCR Local: GPU not available, using CPU (will be slower)")
            
            # Load model and processor
            logger.info(f"GLM-OCR Local: Loading model from {model_path}...")
            
            self._model = AutoModel.from_pretrained(
                str(model_path),
                trust_remote_code=True  # Required for custom models
            )
            self._model.to(self._device)
            self._model.eval()  # Set to evaluation mode
            
            self._processor = AutoProcessor.from_pretrained(
                str(model_path),
                trust_remote_code=True
            )
            
            logger.info("GLM-OCR Local: Model loaded successfully")
            logger.info(f"GLM-OCR Local: Device: {self._device}")
            
        except OCRError:
            raise
        except Exception as e:
            raise OCRError(f"Failed to initialize GLM-OCR local model: {e}")
    
    def detect_and_recognize(self, image: np.ndarray) -> List[Tuple]:
        """Detect and recognize text using local GLM-OCR model.
        
        Args:
            image: Input image as numpy array
            
        Returns:
            List of tuples (box_points, text, confidence)
            
        Raises:
            OCRError: If inference fails
        """
        if self._model is None or self._processor is None:
            raise OCRError("GLM-OCR model not initialized. Call initialize() first.")
        
        try:
            import torch
            
            # Convert numpy array to PIL Image
            if len(image.shape) == 3 and image.shape[2] == 3:
                # BGR to RGB
                image_rgb = image[:, :, ::-1]
            else:
                image_rgb = image
            
            pil_image = Image.fromarray(image_rgb)
            
            # Preprocess image
            inputs = self._processor(images=pil_image, return_tensors="pt")
            inputs = {k: v.to(self._device) for k, v in inputs.items()}
            
            # Run inference
            logger.debug("GLM-OCR Local: Running inference...")
            
            with torch.no_grad():
                outputs = self._model(**inputs)
            
            # Parse outputs
            ocr_results = self._parse_model_output(outputs, pil_image.size)
            
            logger.debug(f"GLM-OCR Local: Detected {len(ocr_results)} text regions")
            return ocr_results
            
        except Exception as e:
            raise OCRError(f"GLM-OCR local inference failed: {e}")
    
    def _parse_model_output(self, outputs, image_size: tuple) -> List[Tuple]:
        """Parse GLM-OCR model output.
        
        Args:
            outputs: Model output tensors
            image_size: Original image size (width, height)
            
        Returns:
            List of tuples (box_points, text, confidence)
        """
        ocr_results = []
        
        try:
            # Note: The exact parsing logic depends on GLM-OCR's output format
            # This is a placeholder implementation that needs to be adjusted
            # based on the actual model's output structure
            
            # Example parsing (adjust based on actual model):
            if hasattr(outputs, 'boxes') and hasattr(outputs, 'texts'):
                boxes = outputs.boxes
                texts = outputs.texts
                confidences = getattr(outputs, 'confidences', None)
                
                for i, (box, text) in enumerate(zip(boxes, texts)):
                    # Convert box coordinates to corner points
                    # Assuming box is in format [x1, y1, x2, y2]
                    if len(box) == 4:
                        x1, y1, x2, y2 = box
                        box_points = [
                            [float(x1), float(y1)],
                            [float(x2), float(y1)],
                            [float(x2), float(y2)],
                            [float(x1), float(y2)]
                        ]
                    else:
                        box_points = [[float(x) for x in point] for point in box]
                    
                    # Get confidence
                    if confidences is not None:
                        confidence = float(confidences[i])
                    else:
                        confidence = 1.0
                    
                    # Get text
                    text_str = str(text).strip()
                    if not text_str:
                        continue
                    
                    ocr_results.append((box_points, text_str, confidence))
            
            else:
                logger.warning("GLM-OCR Local: Unexpected model output format")
                logger.debug(f"Output keys: {dir(outputs)}")
        
        except Exception as e:
            logger.error(f"GLM-OCR Local: Failed to parse model output: {e}")
            logger.debug(f"Output: {outputs}")
        
        return ocr_results
    
    def get_engine_info(self) -> dict:
        """Get GLM-OCR local engine information.
        
        Returns:
            Dictionary with engine information
        """
        device_info = str(self._device) if self._device else "unknown"
        
        return {
            "name": "GLM-OCR-Local",
            "version": self._get_model_version(),
            "type": "local",
            "device": device_info,
            "capabilities": ["text_detection", "text_recognition", "offline"]
        }
    
    def _get_model_version(self) -> str:
        """Get model version.
        
        Returns:
            Version string or "unknown"
        """
        try:
            if self._model and hasattr(self._model, 'config'):
                return getattr(self._model.config, 'version', 'unknown')
        except:
            pass
        return "unknown"
    
    def supports_batch_processing(self) -> bool:
        """GLM-OCR local model may support batch processing.
        
        Returns:
            True (can process multiple images in a batch)
        """
        return True
    
    def cleanup(self) -> None:
        """Clean up model resources."""
        if self._model:
            # Move model to CPU and clear cache
            try:
                import torch
                self._model.cpu()
                if torch.cuda.is_available():
                    torch.cuda.empty_cache()
            except:
                pass
            
            self._model = None
            self._processor = None
            logger.debug("GLM-OCR Local: Cleaned up model resources")
