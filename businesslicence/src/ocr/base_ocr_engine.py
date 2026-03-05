"""Base OCR Engine - Abstract interface for OCR engines.

This module defines the abstract base class for OCR engines, allowing
multiple OCR implementations (PaddleOCR, GLM-OCR, etc.) to be used
interchangeably.

本模块定义了 OCR 引擎的抽象基类，允许多种 OCR 实现（PaddleOCR、GLM-OCR 等）
可互换使用。
"""

from abc import ABC, abstractmethod
from typing import List, Tuple
import numpy as np

from src.models import TextRegion


class BaseOCREngine(ABC):
    """Abstract base class for OCR engines.
    
    All OCR engine implementations must inherit from this class and
    implement the required methods.
    
    所有 OCR 引擎实现都必须继承此类并实现所需方法。
    """
    
    def __init__(self, config):
        """Initialize the OCR engine.
        
        Args:
            config: Configuration manager instance
        """
        self.config = config
        self.engine_name = "base"
    
    @abstractmethod
    def initialize(self) -> None:
        """Initialize the OCR engine.
        
        This method should handle:
        - Loading models
        - Setting up API clients
        - Configuring engine parameters
        
        Raises:
            Exception: If initialization fails
        """
        pass
    
    @abstractmethod
    def detect_and_recognize(self, image: np.ndarray) -> List[Tuple]:
        """Detect and recognize text in an image.
        
        This is the core OCR method that must be implemented by all engines.
        
        Args:
            image: Input image as numpy array (BGR or RGB format)
            
        Returns:
            List of tuples, each containing:
            - box_points: List of 4 corner points [[x1,y1], [x2,y2], [x3,y3], [x4,y4]]
            - text: Recognized text string
            - confidence: Confidence score (0.0 to 1.0)
            
        Example:
            [
                ([[100, 50], [200, 50], [200, 80], [100, 80]], "Hello", 0.95),
                ([[100, 90], [250, 90], [250, 120], [100, 120]], "World", 0.92),
            ]
            
        Raises:
            Exception: If OCR processing fails
        """
        pass
    
    def get_engine_info(self) -> dict:
        """Get information about the OCR engine.
        
        Returns:
            Dictionary containing engine information:
            - name: Engine name
            - version: Engine version (if available)
            - type: Engine type (local/api)
            - capabilities: List of supported features
        """
        return {
            "name": self.engine_name,
            "version": "unknown",
            "type": "unknown",
            "capabilities": []
        }
    
    def supports_batch_processing(self) -> bool:
        """Check if the engine supports batch processing.
        
        Returns:
            True if batch processing is supported, False otherwise
        """
        return False
    
    def cleanup(self) -> None:
        """Clean up resources used by the engine.
        
        This method is called when the engine is no longer needed.
        Subclasses can override this to release resources (close connections, etc.)
        """
        pass
