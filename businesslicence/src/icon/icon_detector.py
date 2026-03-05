"""Icon Detector for identifying and filtering non-text elements.

This module provides the IconDetector class that identifies icons, seals,
QR codes, and other non-text graphical elements in images to prevent
them from being incorrectly processed as text.
"""

import logging
from typing import List, Tuple, Optional

import cv2
import numpy as np

from src.config import ConfigManager
from src.models import TextRegion

logger = logging.getLogger(__name__)


class IconDetector:
    """Detector for identifying icons and non-text graphical elements.
    
    This class analyzes text regions to determine if they are actually
    icons, seals, QR codes, or other graphical elements that should
    not be translated. It uses multiple features including:
    - Aspect ratio (icons tend to be square-like)
    - Content complexity (icons have high edge density)
    - Text density (icons have low text-to-area ratio)
    
    Attributes:
        config: Configuration manager instance
        aspect_ratio_threshold: Threshold for square-like detection (0-1)
        complexity_threshold: Threshold for content complexity (0-1)
        text_density_threshold: Threshold for text density (0-1)
        whitelist: List of icon types to always preserve
    """
    
    def __init__(self, config: ConfigManager):
        """Initialize the icon detector.
        
        Args:
            config: Configuration manager instance
        """
        self.config = config
        self.aspect_ratio_threshold = config.get(
            'icon_detection.aspect_ratio_threshold', 0.8
        )
        self.complexity_threshold = config.get(
            'icon_detection.complexity_threshold', 0.7
        )
        self.text_density_threshold = config.get(
            'icon_detection.text_density_threshold', 0.3
        )
        self.whitelist = config.get(
            'icon_detection.whitelist', ['qrcode', 'seal', 'logo']
        )
        
        # Check if QR code hard protection is enabled
        self.hard_protection_enabled = config.get('qr_pre_detection.hard_protection', False)
        
        logger.info(
            f"IconDetector initialized with thresholds: "
            f"aspect_ratio={self.aspect_ratio_threshold}, "
            f"complexity={self.complexity_threshold}, "
            f"text_density={self.text_density_threshold}, "
            f"hard_protection={self.hard_protection_enabled}"
        )
    
    def calculate_aspect_ratio(self, region: TextRegion) -> float:
        """Calculate the aspect ratio of a text region.
        
        The aspect ratio is calculated as min(width, height) / max(width, height),
        resulting in a value between 0 and 1. Values close to 1 indicate
        square-like regions which are more likely to be icons.
        
        Args:
            region: Text region to analyze
            
        Returns:
            Aspect ratio between 0 and 1
        """
        return region.aspect_ratio
    
    def calculate_complexity(
        self, 
        region: TextRegion, 
        image: np.ndarray
    ) -> float:
        """Calculate the content complexity of a region using edge detection.
        
        Complexity is measured by the density of edges within the region.
        Icons and graphical elements typically have higher edge density
        than plain text.
        
        Args:
            region: Text region to analyze
            image: Original image as numpy array
            
        Returns:
            Complexity score between 0 and 1
        """
        if image is None or image.size == 0:
            return 0.0
        
        x1, y1, x2, y2 = region.bbox
        
        # Ensure coordinates are within image bounds
        height, width = image.shape[:2]
        x1 = max(0, min(x1, width - 1))
        y1 = max(0, min(y1, height - 1))
        x2 = max(x1 + 1, min(x2, width))
        y2 = max(y1 + 1, min(y2, height))
        
        # Extract region of interest
        roi = image[y1:y2, x1:x2]
        
        if roi.size == 0:
            return 0.0
        
        # Convert to grayscale if needed
        if len(roi.shape) == 3:
            gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
        else:
            gray = roi
        
        # Apply Canny edge detection
        edges = cv2.Canny(gray, 50, 150)
        
        # Calculate edge density (ratio of edge pixels to total pixels)
        total_pixels = edges.size
        edge_pixels = np.count_nonzero(edges)
        
        if total_pixels == 0:
            return 0.0
        
        complexity = edge_pixels / total_pixels
        
        # Normalize to 0-1 range (edge density rarely exceeds 0.5)
        complexity = min(1.0, complexity * 2)
        
        logger.debug(
            f"Region complexity: {complexity:.3f} "
            f"(edges={edge_pixels}, total={total_pixels})"
        )
        
        return complexity
    
    def calculate_text_density(self, region: TextRegion) -> float:
        """Calculate the text density of a region.
        
        Text density is the ratio of text character count to the region area.
        Real text regions typically have higher text density than icons.
        
        Args:
            region: Text region to analyze
            
        Returns:
            Text density score between 0 and 1
        """
        if region.area == 0:
            return 0.0
        
        # Calculate expected characters per pixel for typical text
        # Assuming average character width of ~font_size/2 pixels
        text_length = len(region.text.strip())
        
        if text_length == 0:
            return 0.0
        
        # Estimate expected area for the text
        # Average character width is roughly 0.6 * font_size
        # Character height is roughly font_size
        char_width = region.font_size * 0.6
        char_height = region.font_size
        expected_text_area = text_length * char_width * char_height
        
        # Text density is the ratio of expected text area to actual region area
        density = expected_text_area / region.area
        
        # Clamp to 0-1 range
        density = max(0.0, min(1.0, density))
        
        logger.debug(
            f"Region text density: {density:.3f} "
            f"(text_len={text_length}, area={region.area})"
        )
        
        return density

    def detect_qrcode(
        self, 
        region: TextRegion, 
        image: np.ndarray
    ) -> bool:
        """Detect if a region contains a QR code.
        
        Uses multiple methods to identify QR codes:
        1. OpenCV's QRCodeDetector (most reliable)
        2. Visual pattern analysis (shape, complexity, texture)
        3. Heuristic checks (aspect ratio, size, text content)
        
        Args:
            region: Text region to analyze
            image: Original image as numpy array
            
        Returns:
            True if the region contains a QR code, False otherwise
        """
        if image is None or image.size == 0:
            return False
        
        x1, y1, x2, y2 = region.bbox
        
        # Ensure coordinates are within image bounds
        height, width = image.shape[:2]
        x1 = max(0, min(x1, width - 1))
        y1 = max(0, min(y1, height - 1))
        x2 = max(x1 + 1, min(x2, width))
        y2 = max(y1 + 1, min(y2, height))
        
        # Extract region of interest
        roi = image[y1:y2, x1:x2]
        
        if roi.size == 0:
            return False
        
        # Method 1: Try OpenCV QR code detector
        try:
            detector = cv2.QRCodeDetector()
            data, bbox, _ = detector.detectAndDecode(roi)
            
            if data:
                logger.info(f"✅ QR code detected (OpenCV): {region.bbox}, data='{data[:30]}...'")
                return True
        except Exception as e:
            logger.debug(f"OpenCV QR detection failed: {e}")
        
        # Method 2: Visual pattern analysis for QR-like structures
        if self._has_qrcode_pattern(roi, region):
            logger.info(f"✅ QR code detected (pattern analysis): {region.bbox}")
            return True
        
        # Method 3: Heuristic checks
        if self._is_qrcode_by_heuristics(region, roi):
            logger.info(f"✅ QR code detected (heuristics): {region.bbox}")
            return True
        
        return False
    
    def _has_qrcode_pattern(self, roi: np.ndarray, region: TextRegion) -> bool:
        """Check if ROI has QR code-like visual patterns.
        
        QR codes have distinctive visual characteristics:
        - Three position markers (squares in corners)
        - High contrast black/white pattern
        - Regular grid structure
        - High edge density
        
        Args:
            roi: Region of interest image
            region: Text region metadata
            
        Returns:
            True if the region has QR code-like patterns
        """
        if roi.size == 0:
            return False
        
        # QR codes must be reasonably square
        aspect_ratio = region.aspect_ratio
        if aspect_ratio < 0.7:  # Not square enough
            return False
        
        # QR codes must be large enough (at least 50x50 pixels)
        if region.width < 50 or region.height < 50:
            return False
        
        try:
            # Convert to grayscale
            if len(roi.shape) == 3:
                gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
            else:
                gray = roi
            
            # Apply binary threshold to enhance pattern
            _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            
            # Check for high contrast (QR codes are black and white)
            black_pixels = np.sum(binary == 0)
            white_pixels = np.sum(binary == 255)
            total_pixels = binary.size
            
            # QR codes should have significant amounts of both black and white
            black_ratio = black_pixels / total_pixels
            white_ratio = white_pixels / total_pixels
            
            if black_ratio < 0.2 or white_ratio < 0.2:
                return False  # Not enough contrast
            
            # Check for high edge density (QR codes have many edges)
            edges = cv2.Canny(gray, 50, 150)
            edge_density = np.count_nonzero(edges) / edges.size
            
            if edge_density < 0.15:  # QR codes typically have >15% edge density
                return False
            
            # Look for corner markers (three squares in corners)
            # This is a simplified check - real QR detection would be more sophisticated
            contours, _ = cv2.findContours(binary, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
            
            # Count square-like contours
            square_contours = 0
            for contour in contours:
                area = cv2.contourArea(contour)
                if area < 100:  # Too small
                    continue
                
                # Check if contour is square-like
                x, y, w, h = cv2.boundingRect(contour)
                contour_aspect = w / h if h > 0 else 0
                if 0.8 <= contour_aspect <= 1.2:  # Square-ish
                    square_contours += 1
            
            # QR codes typically have multiple square patterns
            if square_contours >= 3:
                logger.debug(
                    f"QR pattern detected: aspect={aspect_ratio:.2f}, "
                    f"black={black_ratio:.2f}, white={white_ratio:.2f}, "
                    f"edges={edge_density:.2f}, squares={square_contours}"
                )
                return True
            
        except Exception as e:
            logger.debug(f"Pattern analysis failed: {e}")
        
        return False
    
    def _is_qrcode_by_heuristics(self, region: TextRegion, roi: np.ndarray) -> bool:
        """Use heuristic rules to identify QR codes.
        
        Heuristics include:
        - Nearly square shape
        - Appropriate size range
        - Low or garbled text content
        - High complexity
        
        Args:
            region: Text region metadata
            roi: Region of interest image
            
        Returns:
            True if heuristics suggest this is a QR code
        """
        # Check aspect ratio (QR codes are square)
        aspect_ratio = region.aspect_ratio
        if aspect_ratio < 0.75:
            return False
        
        # Check size (QR codes are typically 50-500 pixels)
        if region.width < 50 or region.height < 50:
            return False
        if region.width > 500 or region.height > 500:
            return False
        
        # Check text content (QR codes have no readable text or garbled text)
        text = region.text.strip()
        text_length = len(text)
        
        # Empty or very short text is suspicious
        if text_length == 0:
            # Empty text + square shape + right size = likely QR code
            if aspect_ratio >= 0.85 and 80 <= region.width <= 300:
                logger.debug(f"Heuristic match: empty text, square, right size")
                return True
        
        # Very short garbled text (< 5 chars) with square shape
        if text_length <= 5 and aspect_ratio >= 0.85:
            logger.debug(f"Heuristic match: short text '{text}', square shape")
            return True
        
        # Check complexity (QR codes have high complexity)
        complexity = self.calculate_complexity(region, roi)
        if complexity > 0.6 and aspect_ratio >= 0.85:
            logger.debug(f"Heuristic match: high complexity {complexity:.2f}, square")
            return True
        
        return False
    
    def is_in_whitelist(self, region: TextRegion, image: np.ndarray) -> bool:
        """Check if a region matches any whitelist pattern.
        
        Whitelist patterns include:
        - 'qrcode': QR codes detected by OpenCV
        - 'seal': Circular seal-like patterns
        - 'logo': Logo-like graphical elements
        
        Args:
            region: Text region to analyze
            image: Original image as numpy array
            
        Returns:
            True if the region matches a whitelist pattern
        """
        if not self.whitelist:
            return False
        
        # Check for QR code
        if 'qrcode' in self.whitelist:
            if self.detect_qrcode(region, image):
                logger.info(f"Region matched whitelist: qrcode")
                return True
        
        # Check for seal (circular pattern with high aspect ratio)
        if 'seal' in self.whitelist:
            if self._detect_seal(region, image):
                logger.info(f"Region matched whitelist: seal")
                return True
        
        # Check for logo (high complexity, square-ish)
        if 'logo' in self.whitelist:
            if self._detect_logo(region, image):
                logger.info(f"Region matched whitelist: logo")
                return True
        
        return False
    
    def _detect_seal(self, region: TextRegion, image: np.ndarray) -> bool:
        """Detect if a region contains a seal (circular stamp pattern).
        
        Seals are typically:
        - Nearly square (high aspect ratio)
        - Contain circular patterns
        - Have red or other distinctive colors
        
        Args:
            region: Text region to analyze
            image: Original image as numpy array
            
        Returns:
            True if the region appears to be a seal
        """
        if image is None or image.size == 0:
            return False
        
        x1, y1, x2, y2 = region.bbox
        
        # Ensure coordinates are within image bounds
        height, width = image.shape[:2]
        x1 = max(0, min(x1, width - 1))
        y1 = max(0, min(y1, height - 1))
        x2 = max(x1 + 1, min(x2, width))
        y2 = max(y1 + 1, min(y2, height))
        
        roi = image[y1:y2, x1:x2]
        
        if roi.size == 0:
            return False
        
        # Check aspect ratio (seals are typically square)
        aspect_ratio = region.aspect_ratio
        if aspect_ratio < 0.7:
            return False
        
        try:
            # Convert to grayscale
            if len(roi.shape) == 3:
                gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
            else:
                gray = roi
            
            # Apply Gaussian blur
            blurred = cv2.GaussianBlur(gray, (5, 5), 0)
            
            # Detect circles using Hough transform
            circles = cv2.HoughCircles(
                blurred,
                cv2.HOUGH_GRADIENT,
                dp=1,
                minDist=min(roi.shape[:2]) // 4,
                param1=50,
                param2=30,
                minRadius=min(roi.shape[:2]) // 6,
                maxRadius=min(roi.shape[:2]) // 2
            )
            
            if circles is not None and len(circles[0]) > 0:
                return True
            
            return False
            
        except Exception as e:
            logger.debug(f"Seal detection failed: {e}")
            return False
    
    def _detect_logo(self, region: TextRegion, image: np.ndarray) -> bool:
        """Detect if a region contains a logo.
        
        Logos are typically:
        - Square-ish (high aspect ratio)
        - High complexity (many edges)
        - Low text density
        
        Args:
            region: Text region to analyze
            image: Original image as numpy array
            
        Returns:
            True if the region appears to be a logo
        """
        # Logos have high aspect ratio (square-like)
        aspect_ratio = self.calculate_aspect_ratio(region)
        if aspect_ratio < 0.85:
            return False
        
        # Logos have high complexity
        complexity = self.calculate_complexity(region, image)
        if complexity < 0.6:
            return False
        
        # Logos have very low text density
        text_density = self.calculate_text_density(region)
        if text_density > 0.15:
            return False
        
        return True
    
    def is_icon(self, region: TextRegion, image: np.ndarray) -> bool:
        """Determine if a region is an icon rather than text.
        
        A region is classified as an icon if:
        1. It matches a whitelist pattern (QR code, seal, logo), OR
        2. It has icon-like features:
           - High aspect ratio (square-like) AND
           - High content complexity AND
           - Low text density
        3. It contains a QR code pattern (even if mixed with text)
        
        In hard protection mode, regions containing QR codes with text are NOT
        classified as icons, so they can be preserved (but marked as hard_protected).
        
        Args:
            region: Text region to analyze
            image: Original image as numpy array
            
        Returns:
            True if the region is classified as an icon
        """
        # CRITICAL: Check hard protection FIRST before whitelist
        # In hard protection mode, if region contains QR code with text, don't filter it
        if self.hard_protection_enabled:
            if self._contains_qrcode_with_text(region, image):
                logger.info(
                    f"🛡️ Hard protection: Region contains QR code with text, "
                    f"NOT filtering as icon (will be marked as hard_protected later). "
                    f"bbox={region.bbox}, text='{region.text[:30]}...'"
                )
                return False  # Don't filter in hard protection mode
        
        # Check whitelist - whitelist takes priority (but after hard protection check)
        if self.is_in_whitelist(region, image):
            logger.info(
                f"Region classified as icon (whitelist): "
                f"bbox={region.bbox}, text='{region.text[:20]}...'"
            )
            return True
        
        # Special check: if region contains QR code keywords and has QR-like visual patterns
        # This handles cases where OCR merges QR code with nearby text
        # (This check is redundant now since hard protection is checked first, but keep for non-hard-protection mode)
        if self._contains_qrcode_with_text(region, image):
            logger.info(
                f"Region classified as icon (contains QR code): "
                f"bbox={region.bbox}, text='{region.text[:20]}...'"
            )
            return True
        
        # Calculate features
        aspect_ratio = self.calculate_aspect_ratio(region)
        complexity = self.calculate_complexity(region, image)
        text_density = self.calculate_text_density(region)
        
        logger.debug(
            f"Icon features: aspect_ratio={aspect_ratio:.3f}, "
            f"complexity={complexity:.3f}, text_density={text_density:.3f}"
        )
        
        # Check if region meets icon criteria
        is_square_like = aspect_ratio >= self.aspect_ratio_threshold
        is_complex = complexity >= self.complexity_threshold
        is_low_text = text_density <= self.text_density_threshold
        
        # Region is an icon if it's square-like AND complex AND has low text density
        if is_square_like and is_complex and is_low_text:
            logger.info(
                f"Region classified as icon (features): "
                f"bbox={region.bbox}, "
                f"aspect_ratio={aspect_ratio:.3f}, "
                f"complexity={complexity:.3f}, "
                f"text_density={text_density:.3f}"
            )
            return True
        
        return False
    
    def _contains_qrcode_with_text(self, region: TextRegion, image: np.ndarray) -> bool:
        """Check if region contains a QR code mixed with text.
        
        This handles cases where OCR merges QR code with nearby description text.
        Uses a simplified approach: if text contains QR keywords and the region
        has high edge density, it likely contains a QR code.
        
        Args:
            region: Text region to analyze
            image: Original image as numpy array
            
        Returns:
            True if the region contains a QR code (even if mixed with text)
        """
        # Check if text contains QR code keywords
        text_lower = region.text.lower()
        qr_keywords = ['二维码', 'qr code', 'qrcode', 'qr', '扫描', 'scan']
        
        has_qr_keyword = any(keyword in text_lower for keyword in qr_keywords)
        
        if not has_qr_keyword:
            return False
        
        # If has QR keyword, check if region has QR-like visual characteristics
        x1, y1, x2, y2 = region.bbox
        height, width = image.shape[:2]
        
        # Ensure coordinates are within bounds
        x1 = max(0, min(x1, width - 1))
        y1 = max(0, min(y1, height - 1))
        x2 = max(x1 + 1, min(x2, width))
        y2 = max(y1 + 1, min(y2, height))
        
        roi = image[y1:y2, x1:x2]
        
        if roi.size == 0:
            return False
        
        try:
            # Convert to grayscale
            if len(roi.shape) == 3:
                gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
            else:
                gray = roi
            
            # Method 1: Check overall edge density
            # QR codes have very high edge density even when mixed with text
            edges = cv2.Canny(gray, 50, 150)
            edge_density = np.count_nonzero(edges) / edges.size
            
            # Method 2: Check for high-contrast black/white pattern
            _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            black_pixels = np.sum(binary == 0)
            white_pixels = np.sum(binary == 255)
            total_pixels = binary.size
            
            black_ratio = black_pixels / total_pixels
            white_ratio = white_pixels / total_pixels
            
            # QR codes have high edge density (>12%) and balanced black/white
            has_high_edges = edge_density > 0.12
            has_contrast = black_ratio > 0.15 and white_ratio > 0.15
            
            logger.debug(
                f"QR keyword region analysis: "
                f"edge_density={edge_density:.3f}, "
                f"black_ratio={black_ratio:.3f}, "
                f"white_ratio={white_ratio:.3f}"
            )
            
            # If region has QR keywords AND (high edges OR good contrast), likely contains QR
            if has_qr_keyword and (has_high_edges or has_contrast):
                logger.info(
                    f"✅ Found QR code mixed with text: "
                    f"bbox={region.bbox}, "
                    f"text='{region.text[:30]}...', "
                    f"edge_density={edge_density:.3f}, "
                    f"has_high_edges={has_high_edges}, "
                    f"has_contrast={has_contrast}"
                )
                return True
            
        except Exception as e:
            logger.debug(f"Error checking for QR code in text region: {e}")
        
        return False

    def filter_icons(
        self, 
        regions: List[TextRegion], 
        image: np.ndarray
    ) -> List[TextRegion]:
        """Filter out icon regions from a list of text regions.
        
        Iterates through all text regions and removes those classified
        as icons, returning only the regions that contain actual text
        for translation.
        
        Special handling: If a region contains both QR code and text,
        attempt to split it into separate regions. If split fails, use
        fallback strategy based on edge density analysis.
        
        Args:
            regions: List of TextRegion objects to filter
            image: Original image as numpy array
            
        Returns:
            List of TextRegion objects that are not icons (text only)
        """
        if not regions:
            return []
        
        logger.info(f"[ICON FILTER] Starting with {len(regions)} regions")
        
        text_regions = []
        icon_regions = []
        
        for i, region in enumerate(regions):
            logger.debug(f"[ICON FILTER] Region {i+1}/{len(regions)}: bbox={region.bbox}, text='{region.text[:30]}...'")
            
            if self.is_icon(region, image):
                logger.info(f"[ICON FILTER] Region {i+1} classified as ICON: bbox={region.bbox}")
                
                # Check if this region contains QR code mixed with text
                if self._contains_qrcode_with_text(region, image):
                    logger.info(f"[ICON FILTER] Region {i+1} contains QR code with text, attempting split...")
                    
                    # Try to split the region into QR code part and text part
                    split_regions = self._split_qrcode_and_text(region, image)
                    if split_regions:
                        logger.info(
                            f"[ICON FILTER] ✅ Split QR code region into {len(split_regions)} parts: "
                            f"original_bbox={region.bbox}"
                        )
                        # Add the split regions to appropriate lists
                        for split_region in split_regions:
                            if split_region.get('is_qrcode', False):
                                icon_regions.append(split_region['region'])
                                logger.info(f"[ICON FILTER]   - QR code part: {split_region['region'].bbox}")
                            else:
                                text_regions.append(split_region['region'])
                                logger.info(f"[ICON FILTER]   - Text part: {split_region['region'].bbox}, text='{split_region['region'].text[:30]}...'")
                        continue
                    else:
                        # Split failed - use fallback strategy
                        logger.warning(f"[ICON FILTER] ❌ Split failed for region {i+1}, using fallback strategy")
                        
                        # Analyze edge density to determine if region is primarily QR code or text
                        x1, y1, x2, y2 = region.bbox
                        height, width = image.shape[:2]
                        x1 = max(0, min(x1, width - 1))
                        y1 = max(0, min(y1, height - 1))
                        x2 = max(x1 + 1, min(x2, width))
                        y2 = max(y1 + 1, min(y2, height))
                        
                        roi = image[y1:y2, x1:x2]
                        
                        if roi.size > 0:
                            # Convert to grayscale
                            if len(roi.shape) == 3:
                                gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
                            else:
                                gray = roi
                            
                            # Calculate edge density
                            edges = cv2.Canny(gray, 50, 150)
                            edge_density = np.count_nonzero(edges) / edges.size
                            
                            # Get configuration thresholds
                            qr_edge_threshold = self.config.get('icon_detection.qr_edge_density_threshold', 0.15)
                            text_edge_threshold = self.config.get('icon_detection.text_edge_density_threshold', 0.10)
                            protect_qr_by_default = self.config.get('icon_detection.protect_qr_by_default', True)
                            
                            logger.info(
                                f"[ICON FILTER] Fallback analysis: edge_density={edge_density:.3f}, "
                                f"qr_threshold={qr_edge_threshold}, text_threshold={text_edge_threshold}"
                            )
                            
                            # Decide based on edge density
                            if edge_density > qr_edge_threshold:
                                # Primarily QR code - protect it (filter as icon)
                                logger.info(
                                    f"[ICON FILTER] 🔒 Fallback: Primarily QR code (edge_density={edge_density:.3f} > {qr_edge_threshold}), "
                                    f"protecting QR code, sacrificing text translation"
                                )
                                icon_regions.append(region)
                                continue
                            elif edge_density < text_edge_threshold:
                                # Primarily text - translate it (keep as text)
                                logger.info(
                                    f"[ICON FILTER] 📝 Fallback: Primarily text (edge_density={edge_density:.3f} < {text_edge_threshold}), "
                                    f"translating text, may affect QR code"
                                )
                                text_regions.append(region)
                                continue
                            else:
                                # Ambiguous - use default strategy
                                if protect_qr_by_default:
                                    logger.info(
                                        f"[ICON FILTER] ⚠️ Fallback: Ambiguous region (edge_density={edge_density:.3f}), "
                                        f"using protect_qr_by_default=True, protecting QR code"
                                    )
                                    icon_regions.append(region)
                                else:
                                    logger.info(
                                        f"[ICON FILTER] ⚠️ Fallback: Ambiguous region (edge_density={edge_density:.3f}), "
                                        f"using protect_qr_by_default=False, translating text"
                                    )
                                    text_regions.append(region)
                                continue
                
                # Regular icon (not splittable)
                icon_regions.append(region)
                logger.debug(
                    f"[ICON FILTER] Filtered icon region: bbox={region.bbox}, "
                    f"text='{region.text[:30] if len(region.text) > 30 else region.text}'"
                )
            else:
                text_regions.append(region)
                logger.debug(f"[ICON FILTER] Region {i+1} classified as TEXT")
        
        logger.info(
            f"[ICON FILTER] Complete: {len(regions)} regions -> "
            f"{len(icon_regions)} icons + {len(text_regions)} text = {len(icon_regions) + len(text_regions)} total"
        )
        
        return text_regions
    
    def _split_qrcode_and_text(self, region: TextRegion, image: np.ndarray) -> Optional[List[dict]]:
        """Split a region containing both QR code and text into separate regions.
        
        Args:
            region: Text region containing both QR code and text
            image: Original image as numpy array
            
        Returns:
            List of dicts with 'region' and 'is_qrcode' keys, or None if split failed
        """
        x1, y1, x2, y2 = region.bbox
        height, width = image.shape[:2]
        
        # Ensure coordinates are within bounds
        x1 = max(0, min(x1, width - 1))
        y1 = max(0, min(y1, height - 1))
        x2 = max(x1 + 1, min(x2, width))
        y2 = max(y1 + 1, min(y2, height))
        
        roi = image[y1:y2, x1:x2]
        
        if roi.size == 0:
            return None
        
        try:
            # Convert to grayscale
            if len(roi.shape) == 3:
                gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
            else:
                gray = roi
            
            # Apply binary threshold
            _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            
            # Find the QR code area by looking for high edge density regions
            edges = cv2.Canny(gray, 50, 150)
            
            # Divide the region into left and right halves
            roi_width = x2 - x1
            roi_height = y2 - y1
            
            # Try vertical split (QR code on left, text on right)
            # Check edge density in left half vs right half
            left_half = edges[:, :roi_width//2]
            right_half = edges[:, roi_width//2:]
            
            left_density = np.count_nonzero(left_half) / left_half.size if left_half.size > 0 else 0
            right_density = np.count_nonzero(right_half) / right_half.size if right_half.size > 0 else 0
            
            logger.debug(
                f"Split analysis: left_density={left_density:.3f}, right_density={right_density:.3f}"
            )
            
            # If left half has much higher edge density, it's likely the QR code
            if left_density > 0.12 and left_density > right_density * 1.5:
                # Split vertically: left = QR code, right = text
                split_x = x1 + roi_width // 2
                
                # QR code region (left half)
                qr_region = TextRegion(
                    bbox=(x1, y1, split_x, y2),
                    text="",  # QR code has no text
                    confidence=region.confidence,
                    font_size=region.font_size,
                    angle=region.angle
                )
                
                # Text region (right half)
                text_region = TextRegion(
                    bbox=(split_x, y1, x2, y2),
                    text=region.text,  # Keep the original text
                    confidence=region.confidence,
                    font_size=region.font_size,
                    angle=region.angle
                )
                
                return [
                    {'region': qr_region, 'is_qrcode': True},
                    {'region': text_region, 'is_qrcode': False}
                ]
            
            # Try horizontal split (QR code on top, text on bottom)
            top_half = edges[:roi_height//2, :]
            bottom_half = edges[roi_height//2:, :]
            
            top_density = np.count_nonzero(top_half) / top_half.size if top_half.size > 0 else 0
            bottom_density = np.count_nonzero(bottom_half) / bottom_half.size if bottom_half.size > 0 else 0
            
            if top_density > 0.12 and top_density > bottom_density * 1.5:
                # Split horizontally: top = QR code, bottom = text
                split_y = y1 + roi_height // 2
                
                qr_region = TextRegion(
                    bbox=(x1, y1, x2, split_y),
                    text="",
                    confidence=region.confidence,
                    font_size=region.font_size,
                    angle=region.angle
                )
                
                text_region = TextRegion(
                    bbox=(x1, split_y, x2, y2),
                    text=region.text,
                    confidence=region.confidence,
                    font_size=region.font_size,
                    angle=region.angle
                )
                
                return [
                    {'region': qr_region, 'is_qrcode': True},
                    {'region': text_region, 'is_qrcode': False}
                ]
            
            # Cannot split reliably
            logger.debug("Cannot split QR code and text reliably")
            return None
            
        except Exception as e:
            logger.debug(f"Error splitting QR code and text: {e}")
            return None
    
    def get_icon_regions(
        self, 
        regions: List[TextRegion], 
        image: np.ndarray
    ) -> List[TextRegion]:
        """Get all regions classified as icons.
        
        This is the inverse of filter_icons - returns only the regions
        that are classified as icons.
        
        Args:
            regions: List of TextRegion objects to analyze
            image: Original image as numpy array
            
        Returns:
            List of TextRegion objects that are classified as icons
        """
        if not regions:
            return []
        
        icon_regions = []
        
        for region in regions:
            if self.is_icon(region, image):
                icon_regions.append(region)
        
        return icon_regions
