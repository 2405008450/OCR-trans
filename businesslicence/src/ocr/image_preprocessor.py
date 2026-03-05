"""Image preprocessing module for OCR accuracy improvement.

This module provides various image preprocessing techniques to enhance
OCR recognition accuracy, including denoising, binarization, contrast
enhancement, and sharpening.
"""

import logging
from typing import Optional, Tuple
import numpy as np
import cv2

from src.config import ConfigManager

logger = logging.getLogger(__name__)


class ImagePreprocessor:
    """Image preprocessor for OCR enhancement.
    
    Provides multiple preprocessing strategies:
    - Denoising: Remove noise from images
    - Binarization: Convert to black and white for better contrast
    - Contrast enhancement: Improve text visibility
    - Sharpening: Enhance text edges
    - Deskewing: Correct image rotation
    
    Attributes:
        config: Configuration manager instance
        enabled: Whether preprocessing is enabled
        strategy: Preprocessing strategy to use
    """
    
    def __init__(self, config: ConfigManager):
        """Initialize the image preprocessor.
        
        Args:
            config: Configuration manager instance
        """
        self.config = config
        self.enabled = config.get('ocr.preprocessing.enabled', False)
        self.strategy = config.get('ocr.preprocessing.strategy', 'auto')
        
        # Preprocessing parameters
        self.denoise_strength = config.get('ocr.preprocessing.denoise_strength', 10)
        self.contrast_clip_limit = config.get('ocr.preprocessing.contrast_clip_limit', 2.0)
        self.sharpen_strength = config.get('ocr.preprocessing.sharpen_strength', 1.0)
        
        logger.info(
            f"ImagePreprocessor initialized: "
            f"enabled={self.enabled}, strategy={self.strategy}"
        )
    
    def preprocess(self, image: np.ndarray) -> np.ndarray:
        """Apply preprocessing to an image.
        
        Args:
            image: Input image (BGR format)
            
        Returns:
            Preprocessed image
        """
        if not self.enabled:
            return image
        
        if self.strategy == 'auto':
            return self._auto_preprocess(image)
        elif self.strategy == 'denoise':
            return self._denoise_only(image)
        elif self.strategy == 'binarize':
            return self._binarize_only(image)
        elif self.strategy == 'enhance':
            return self._enhance_only(image)
        elif self.strategy == 'full':
            return self._full_preprocess(image)
        else:
            logger.warning(f"Unknown preprocessing strategy: {self.strategy}")
            return image
    
    def _auto_preprocess(self, image: np.ndarray) -> np.ndarray:
        """Automatically select best preprocessing strategy.
        
        Analyzes image characteristics and applies appropriate preprocessing.
        
        Args:
            image: Input image
            
        Returns:
            Preprocessed image
        """
        # Analyze image quality
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY) if len(image.shape) == 3 else image
        
        # Calculate image metrics
        mean_brightness = np.mean(gray)
        std_brightness = np.std(gray)
        
        logger.debug(
            f"Image analysis: brightness={mean_brightness:.1f}, "
            f"std={std_brightness:.1f}"
        )
        
        # Low contrast image: apply contrast enhancement
        if std_brightness < 40:
            logger.info("Low contrast detected, applying contrast enhancement")
            return self._enhance_contrast(image)
        
        # Noisy image: apply denoising
        elif std_brightness > 80:
            logger.info("High noise detected, applying denoising")
            return self._denoise(image)
        
        # Normal image: light preprocessing
        else:
            logger.info("Normal image quality, applying light preprocessing")
            return self._light_preprocess(image)
    
    def _denoise_only(self, image: np.ndarray) -> np.ndarray:
        """Apply denoising only.
        
        Args:
            image: Input image
            
        Returns:
            Denoised image
        """
        return self._denoise(image)
    
    def _binarize_only(self, image: np.ndarray) -> np.ndarray:
        """Apply binarization only.
        
        Args:
            image: Input image
            
        Returns:
            Binarized image
        """
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY) if len(image.shape) == 3 else image
        binary = self._adaptive_threshold(gray)
        return cv2.cvtColor(binary, cv2.COLOR_GRAY2BGR)
    
    def _enhance_only(self, image: np.ndarray) -> np.ndarray:
        """Apply contrast enhancement only.
        
        Args:
            image: Input image
            
        Returns:
            Enhanced image
        """
        return self._enhance_contrast(image)
    
    def _full_preprocess(self, image: np.ndarray) -> np.ndarray:
        """Apply full preprocessing pipeline.
        
        Steps:
        1. Denoise
        2. Enhance contrast
        3. Sharpen
        4. Binarize (optional)
        
        Args:
            image: Input image
            
        Returns:
            Fully preprocessed image
        """
        # Step 1: Denoise
        denoised = self._denoise(image)
        
        # Step 2: Enhance contrast
        enhanced = self._enhance_contrast(denoised)
        
        # Step 3: Sharpen
        sharpened = self._sharpen(enhanced)
        
        logger.info("Full preprocessing applied: denoise + enhance + sharpen")
        return sharpened
    
    def _light_preprocess(self, image: np.ndarray) -> np.ndarray:
        """Apply light preprocessing (contrast + sharpen).
        
        Args:
            image: Input image
            
        Returns:
            Lightly preprocessed image
        """
        enhanced = self._enhance_contrast(image)
        sharpened = self._sharpen(enhanced)
        return sharpened
    
    def _denoise(self, image: np.ndarray) -> np.ndarray:
        """Remove noise from image.
        
        Uses Non-local Means Denoising algorithm.
        
        Args:
            image: Input image
            
        Returns:
            Denoised image
        """
        if len(image.shape) == 3:
            # Color image
            denoised = cv2.fastNlMeansDenoisingColored(
                image,
                None,
                h=self.denoise_strength,
                hColor=self.denoise_strength,
                templateWindowSize=7,
                searchWindowSize=21
            )
        else:
            # Grayscale image
            denoised = cv2.fastNlMeansDenoising(
                image,
                None,
                h=self.denoise_strength,
                templateWindowSize=7,
                searchWindowSize=21
            )
        
        return denoised
    
    def _enhance_contrast(self, image: np.ndarray) -> np.ndarray:
        """Enhance image contrast using CLAHE.
        
        CLAHE (Contrast Limited Adaptive Histogram Equalization) improves
        local contrast while preventing over-amplification of noise.
        
        Args:
            image: Input image
            
        Returns:
            Contrast-enhanced image
        """
        # Convert to LAB color space for better results
        if len(image.shape) == 3:
            lab = cv2.cvtColor(image, cv2.COLOR_BGR2LAB)
            l, a, b = cv2.split(lab)
            
            # Apply CLAHE to L channel
            clahe = cv2.createCLAHE(
                clipLimit=self.contrast_clip_limit,
                tileGridSize=(8, 8)
            )
            l_enhanced = clahe.apply(l)
            
            # Merge channels
            lab_enhanced = cv2.merge([l_enhanced, a, b])
            enhanced = cv2.cvtColor(lab_enhanced, cv2.COLOR_LAB2BGR)
        else:
            # Grayscale image
            clahe = cv2.createCLAHE(
                clipLimit=self.contrast_clip_limit,
                tileGridSize=(8, 8)
            )
            enhanced = clahe.apply(image)
        
        return enhanced
    
    def _sharpen(self, image: np.ndarray) -> np.ndarray:
        """Sharpen image to enhance text edges.
        
        Uses unsharp masking technique.
        
        Args:
            image: Input image
            
        Returns:
            Sharpened image
        """
        # Create Gaussian blur
        blurred = cv2.GaussianBlur(image, (0, 0), 3)
        
        # Unsharp mask: original + (original - blurred) * strength
        sharpened = cv2.addWeighted(
            image, 1.0 + self.sharpen_strength,
            blurred, -self.sharpen_strength,
            0
        )
        
        return sharpened
    
    def _adaptive_threshold(self, gray: np.ndarray) -> np.ndarray:
        """Apply adaptive thresholding for binarization.
        
        Args:
            gray: Grayscale image
            
        Returns:
            Binary image
        """
        binary = cv2.adaptiveThreshold(
            gray,
            255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY,
            11,
            2
        )
        
        return binary
    
    def deskew(self, image: np.ndarray) -> Tuple[np.ndarray, float]:
        """Detect and correct image skew.
        
        Args:
            image: Input image
            
        Returns:
            Tuple of (deskewed image, rotation angle in degrees)
        """
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY) if len(image.shape) == 3 else image
        
        # Detect edges
        edges = cv2.Canny(gray, 50, 150, apertureSize=3)
        
        # Detect lines using Hough transform
        lines = cv2.HoughLines(edges, 1, np.pi / 180, 200)
        
        if lines is None or len(lines) == 0:
            logger.debug("No lines detected for deskewing")
            return image, 0.0
        
        # Calculate average angle
        angles = []
        for line in lines:
            rho, theta = line[0]
            angle = np.degrees(theta) - 90
            
            # Filter out near-vertical lines
            if abs(angle) < 45:
                angles.append(angle)
        
        if not angles:
            logger.debug("No valid angles for deskewing")
            return image, 0.0
        
        # Use median angle to avoid outliers
        median_angle = np.median(angles)
        
        # Only deskew if angle is significant (> 0.5 degrees)
        if abs(median_angle) < 0.5:
            logger.debug(f"Skew angle too small: {median_angle:.2f}°")
            return image, 0.0
        
        # Rotate image
        height, width = image.shape[:2]
        center = (width // 2, height // 2)
        rotation_matrix = cv2.getRotationMatrix2D(center, median_angle, 1.0)
        
        deskewed = cv2.warpAffine(
            image,
            rotation_matrix,
            (width, height),
            flags=cv2.INTER_CUBIC,
            borderMode=cv2.BORDER_REPLICATE
        )
        
        logger.info(f"Image deskewed by {median_angle:.2f}°")
        return deskewed, median_angle
