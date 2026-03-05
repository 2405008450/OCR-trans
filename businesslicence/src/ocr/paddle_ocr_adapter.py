"""PaddleOCR Adapter - Wrapper for PaddleOCR engine.

This module provides an adapter for PaddleOCR that implements the
BaseOCREngine interface.

本模块提供 PaddleOCR 的适配器，实现 BaseOCREngine 接口。
"""

import logging
from typing import List, Tuple

import numpy as np

from src.ocr.base_ocr_engine import BaseOCREngine
from src.exceptions import OCRError

logger = logging.getLogger(__name__)


class PaddleOCRAdapter(BaseOCREngine):
    """Adapter for PaddleOCR engine.
    
    This adapter wraps PaddleOCR and provides a unified interface
    compatible with the BaseOCREngine abstract class.
    
    此适配器封装 PaddleOCR 并提供与 BaseOCREngine 抽象类兼容的统一接口。
    """
    
    def __init__(self, config):
        """Initialize PaddleOCR adapter.
        
        Args:
            config: Configuration manager instance
        """
        super().__init__(config)
        self.engine_name = "PaddleOCR"
        self._ocr = None
    
    def initialize(self) -> None:
        """Initialize PaddleOCR engine.
        
        Raises:
            OCRError: If PaddleOCR initialization fails
        """
        try:
            from paddleocr import PaddleOCR
            
            # 检查是否使用服务器版模型
            use_server = self.config.get('ocr.use_server_model', False)
            
            # 从配置文件读取是否启用角度分类器
            use_angle_cls = self.config.get('ocr.use_angle_cls', True)
            logger.info(f"PaddleOCR: Text angle classifier: {'enabled' if use_angle_cls else 'disabled'}")
            
            # 检查GPU是否真正可用
            gpu_available = False
            try:
                import paddle
                gpu_available = paddle.device.cuda.device_count() > 0
            except:
                gpu_available = False
            
            if not gpu_available:
                logger.warning("PaddleOCR: GPU not available, falling back to CPU mode")
            
            # 基础参数
            ocr_params = {
                'use_angle_cls': use_angle_cls,
                'lang': 'ch',
                'use_gpu': gpu_available,
                'show_log': False
            }
            
            # 检查是否有竖版模式的检测参数覆盖
            portrait_det_params = self.config.get('ocr.portrait_det_params', None)
            
            if use_server:
                logger.info("PaddleOCR: Initializing with server model (more accurate, slower)")
                ocr_params.update({
                    'det_model_dir': None,
                    'rec_model_dir': None,
                    'det_limit_side_len': 960,
                    'rec_batch_num': 6
                })
                
                if portrait_det_params:
                    logger.info("PaddleOCR: Using portrait mode detection parameters")
                    if portrait_det_params.get('det_db_thresh') is not None:
                        ocr_params['det_db_thresh'] = portrait_det_params['det_db_thresh']
                    if portrait_det_params.get('det_db_box_thresh') is not None:
                        ocr_params['det_db_box_thresh'] = portrait_det_params['det_db_box_thresh']
                    if portrait_det_params.get('det_db_unclip_ratio') is not None:
                        ocr_params['det_db_unclip_ratio'] = portrait_det_params['det_db_unclip_ratio']
                else:
                    ocr_params['det_db_thresh'] = 0.5
                    ocr_params['det_db_box_thresh'] = 0.6
            else:
                logger.info("PaddleOCR: Initializing with mobile model (faster, standard accuracy)")
                
                if portrait_det_params:
                    logger.info("PaddleOCR: Using portrait mode detection parameters")
                    if portrait_det_params.get('det_db_thresh') is not None:
                        ocr_params['det_db_thresh'] = portrait_det_params['det_db_thresh']
                    if portrait_det_params.get('det_db_box_thresh') is not None:
                        ocr_params['det_db_box_thresh'] = portrait_det_params['det_db_box_thresh']
                    if portrait_det_params.get('det_db_unclip_ratio') is not None:
                        ocr_params['det_db_unclip_ratio'] = portrait_det_params['det_db_unclip_ratio']
                else:
                    ocr_params['det_db_thresh'] = 0.3
                    ocr_params['det_db_box_thresh'] = 0.5
            
            self._ocr = PaddleOCR(**ocr_params)
            logger.info("PaddleOCR: Initialized successfully")
            
        except ImportError as e:
            raise OCRError(f"PaddleOCR is not installed: {e}")
        except Exception as e:
            raise OCRError(f"Failed to initialize PaddleOCR: {e}")
    
    def detect_and_recognize(self, image: np.ndarray) -> List[Tuple]:
        """Detect and recognize text using PaddleOCR.
        
        Args:
            image: Input image as numpy array
            
        Returns:
            List of tuples (box_points, text, confidence)
            
        Raises:
            OCRError: If OCR processing fails
        """
        if self._ocr is None:
            raise OCRError("PaddleOCR not initialized. Call initialize() first.")
        
        try:
            # Try new API first (PaddleOCR 3.x), fall back to old API
            try:
                result = self._ocr.predict(image)
            except (TypeError, AttributeError):
                # Fall back to old API (PaddleOCR 2.x)
                try:
                    result = self._ocr.ocr(image, cls=True)
                except TypeError:
                    result = self._ocr.ocr(image)
            
            if result is None or len(result) == 0:
                logger.debug("PaddleOCR: No text detected")
                return []
            
            # Convert PaddleOCR result to standard format
            ocr_results = []
            
            for line in result:
                if line is None:
                    continue
                for item in line:
                    try:
                        # Extract box points, text, and confidence
                        if len(item) == 2:
                            box_points = item[0]
                            text_info = item[1]
                            
                            if isinstance(text_info, (tuple, list)) and len(text_info) == 2:
                                text, confidence = text_info
                            else:
                                text = str(text_info)
                                confidence = 1.0
                        elif len(item) == 3:
                            box_points, text, confidence = item
                        else:
                            logger.warning(f"PaddleOCR: Unexpected result format: {item}")
                            continue
                        
                        # Ensure text is string
                        text = str(text).strip()
                        if not text:
                            continue
                        
                        # Ensure confidence is float
                        try:
                            confidence = float(confidence)
                        except (ValueError, TypeError):
                            confidence = 1.0
                        
                        # Ensure box_points is list
                        if isinstance(box_points, np.ndarray):
                            box_points = box_points.tolist()
                        
                        ocr_results.append((box_points, text, confidence))
                        
                    except Exception as e:
                        logger.warning(f"PaddleOCR: Failed to parse result item: {e}")
                        continue
            
            logger.debug(f"PaddleOCR: Detected {len(ocr_results)} text regions")
            return ocr_results
            
        except Exception as e:
            raise OCRError(f"PaddleOCR processing failed: {e}")
    
    def get_engine_info(self) -> dict:
        """Get PaddleOCR engine information.
        
        Returns:
            Dictionary with engine information
        """
        return {
            "name": "PaddleOCR",
            "version": self._get_paddle_version(),
            "type": "local",
            "capabilities": ["text_detection", "text_recognition", "angle_classification"]
        }
    
    def _get_paddle_version(self) -> str:
        """Get PaddleOCR version.
        
        Returns:
            Version string or "unknown"
        """
        try:
            import paddleocr
            return getattr(paddleocr, '__version__', 'unknown')
        except:
            return "unknown"
    
    def supports_batch_processing(self) -> bool:
        """PaddleOCR supports batch processing.
        
        Returns:
            True
        """
        return True
