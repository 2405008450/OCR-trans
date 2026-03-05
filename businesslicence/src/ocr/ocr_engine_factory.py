"""OCR Engine Factory - Factory for creating OCR engine instances.

This module provides a factory class for instantiating different OCR engines
based on configuration settings.

本模块提供工厂类，根据配置设置实例化不同的 OCR 引擎。
"""

import logging
from typing import Optional

from src.config import ConfigManager
from src.ocr.base_ocr_engine import BaseOCREngine
from src.exceptions import OCRError

logger = logging.getLogger(__name__)


class OCREngineFactory:
    """Factory for creating OCR engine instances.
    
    This factory creates the appropriate OCR engine adapter based on
    the configuration settings.
    
    此工厂根据配置设置创建适当的 OCR 引擎适配器。
    """
    
    # Supported engine types
    SUPPORTED_ENGINES = {
        'paddle': 'PaddleOCR',
        'paddleocr': 'PaddleOCR',
        'glm-api': 'GLM-OCR-API',
        'glm_api': 'GLM-OCR-API',
        'glm-local': 'GLM-OCR-Local',
        'glm_local': 'GLM-OCR-Local',
    }
    
    @staticmethod
    def create_engine(config: ConfigManager) -> BaseOCREngine:
        """Create an OCR engine based on configuration.
        
        Args:
            config: Configuration manager instance
            
        Returns:
            BaseOCREngine instance (PaddleOCR, GLM-OCR-API, or GLM-OCR-Local)
            
        Raises:
            OCRError: If engine type is not supported or initialization fails
        """
        # Get engine type from config (default to 'paddle' for backward compatibility)
        engine_type = config.get('ocr.engine', 'paddle').lower()
        
        logger.info(f"Creating OCR engine: {engine_type}")
        
        # Normalize engine type
        if engine_type not in OCREngineFactory.SUPPORTED_ENGINES:
            raise OCRError(
                f"Unsupported OCR engine: '{engine_type}'. "
                f"Supported engines: {list(OCREngineFactory.SUPPORTED_ENGINES.keys())}"
            )
        
        engine_name = OCREngineFactory.SUPPORTED_ENGINES[engine_type]
        
        # Create the appropriate adapter
        try:
            if engine_name == 'PaddleOCR':
                from src.ocr.paddle_ocr_adapter import PaddleOCRAdapter
                engine = PaddleOCRAdapter(config)
                
            elif engine_name == 'GLM-OCR-API':
                from src.ocr.glm_ocr_api_adapter import GLMOCRAPIAdapter
                engine = GLMOCRAPIAdapter(config)
                
            elif engine_name == 'GLM-OCR-Local':
                from src.ocr.glm_ocr_local_adapter import GLMOCRLocalAdapter
                engine = GLMOCRLocalAdapter(config)
                
            else:
                raise OCRError(f"Engine implementation not found: {engine_name}")
            
            # Initialize the engine
            engine.initialize()
            
            # Log engine info
            info = engine.get_engine_info()
            logger.info(
                f"OCR engine initialized: {info['name']} "
                f"(type={info['type']}, version={info.get('version', 'unknown')})"
            )
            
            return engine
            
        except ImportError as e:
            raise OCRError(
                f"Failed to import {engine_name} adapter: {e}\n"
                f"Please ensure required dependencies are installed."
            )
        except Exception as e:
            raise OCRError(f"Failed to create {engine_name} engine: {e}")
    
    @staticmethod
    def get_supported_engines() -> list:
        """Get list of supported engine types.
        
        Returns:
            List of supported engine type strings
        """
        return list(OCREngineFactory.SUPPORTED_ENGINES.keys())
    
    @staticmethod
    def get_engine_info(engine_type: str) -> Optional[str]:
        """Get engine name for a given engine type.
        
        Args:
            engine_type: Engine type string (e.g., 'paddle', 'glm-api')
            
        Returns:
            Engine name or None if not supported
        """
        return OCREngineFactory.SUPPORTED_ENGINES.get(engine_type.lower())
