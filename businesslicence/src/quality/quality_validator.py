"""Quality validator for image translation system.

This module provides the QualityValidator class for validating
the quality of translated images, including translation coverage
calculation, artifact detection, and quality report generation.

_Requirements: 9.1, 9.3, 9.4, 9.5_
"""

import logging
from typing import List, Tuple, Optional

import cv2
import numpy as np

from src.config import ConfigManager
from src.models.data_models import TextRegion, TranslationResult, QualityReport, QualityLevel


logger = logging.getLogger(__name__)


class QualityValidator:
    """Validates the quality of translated images.
    
    The QualityValidator provides functionality for:
    - Calculating translation coverage (ratio of successfully translated regions)
    - Detecting visual artifacts in the output image
    - Generating comprehensive quality reports
    
    Attributes:
        config: Configuration manager instance
        min_coverage: Minimum acceptable translation coverage
        check_artifacts: Whether to check for visual artifacts
    """
    
    def __init__(self, config: ConfigManager):
        """Initialize the quality validator.
        
        Args:
            config: Configuration manager instance
        """
        self.config = config
        self.min_coverage = config.get('quality.min_translation_coverage', 0.9)
        self.check_artifacts_enabled = config.get('quality.check_artifacts', True)
        
        # Artifact detection parameters
        self._edge_threshold_low = 50
        self._edge_threshold_high = 150
        self._color_discontinuity_threshold = 30
        self._min_artifact_area = 100

    def calculate_coverage(
        self,
        regions: List[TextRegion],
        results: List[TranslationResult]
    ) -> float:
        """Calculate the translation coverage ratio.
        
        Coverage is defined as the ratio of successfully translated regions
        to the total number of regions.
        
        Args:
            regions: List of text regions detected in the image
            results: List of translation results corresponding to regions
            
        Returns:
            Coverage ratio between 0.0 and 1.0
            
        _Requirements: 9.1, 9.5_
        """
        if not regions:
            logger.debug("No regions provided, returning coverage of 1.0")
            return 1.0
        
        total_regions = len(regions)
        
        # Count successful translations
        successful_count = sum(1 for result in results if result.success)
        
        coverage = successful_count / total_regions
        
        logger.debug(
            f"Translation coverage: {successful_count}/{total_regions} = {coverage:.2%}"
        )
        
        return coverage
    
    def get_failed_regions(
        self,
        regions: List[TextRegion],
        results: List[TranslationResult]
    ) -> List[TextRegion]:
        """Get the list of regions that failed translation.
        
        Args:
            regions: List of text regions
            results: List of translation results
            
        Returns:
            List of TextRegion objects that failed translation
        """
        failed = []
        
        # Match regions with results
        for i, result in enumerate(results):
            if not result.success and i < len(regions):
                failed.append(regions[i])
                logger.debug(
                    f"Region {i} failed translation: {result.error_message}"
                )
        
        return failed

    def check_artifacts(
        self,
        image: np.ndarray
    ) -> Tuple[bool, List[Tuple[int, int, int, int]]]:
        """Detect visual artifacts in the image.
        
        Uses edge detection and color discontinuity analysis to identify
        areas that may contain visual artifacts such as:
        - Sharp color transitions (color banding)
        - Unnatural edges from poor inpainting
        - Visible seams from text replacement
        
        Args:
            image: Output image as numpy array (BGR format)
            
        Returns:
            Tuple of (has_artifacts, artifact_locations) where:
            - has_artifacts: True if artifacts were detected
            - artifact_locations: List of bounding boxes (x1, y1, x2, y2)
              where artifacts were found
              
        _Requirements: 9.3_
        """
        if image is None or not isinstance(image, np.ndarray):
            logger.warning("Invalid image provided for artifact detection")
            return False, []
        
        if image.size == 0:
            logger.warning("Empty image provided for artifact detection")
            return False, []
        
        artifact_locations = []
        
        # Convert to grayscale for edge detection
        if len(image.shape) == 3:
            gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        else:
            gray = image.copy()
        
        # Detect edges using Canny edge detector
        edges = cv2.Canny(
            gray,
            self._edge_threshold_low,
            self._edge_threshold_high
        )
        
        # Find contours of edge regions
        contours, _ = cv2.findContours(
            edges,
            cv2.RETR_EXTERNAL,
            cv2.CHAIN_APPROX_SIMPLE
        )
        
        # Analyze each contour for potential artifacts
        for contour in contours:
            area = cv2.contourArea(contour)
            
            # Skip small contours
            if area < self._min_artifact_area:
                continue
            
            # Get bounding box
            x, y, w, h = cv2.boundingRect(contour)
            
            # Check for color discontinuity in this region
            if self._has_color_discontinuity(image, x, y, w, h):
                artifact_locations.append((x, y, x + w, y + h))
                logger.debug(
                    f"Artifact detected at ({x}, {y}, {x + w}, {y + h})"
                )
        
        has_artifacts = len(artifact_locations) > 0
        
        if has_artifacts:
            logger.info(
                f"Detected {len(artifact_locations)} potential artifacts"
            )
        else:
            logger.debug("No artifacts detected")
        
        return has_artifacts, artifact_locations
    
    def _has_color_discontinuity(
        self,
        image: np.ndarray,
        x: int,
        y: int,
        w: int,
        h: int
    ) -> bool:
        """Check if a region has significant color discontinuity.
        
        Analyzes the gradient magnitude within a region to detect
        unnatural color transitions that may indicate artifacts.
        
        Args:
            image: Input image
            x, y: Top-left corner of the region
            w, h: Width and height of the region
            
        Returns:
            True if significant color discontinuity is detected
        """
        # Ensure bounds are within image
        height, width = image.shape[:2]
        x = max(0, x)
        y = max(0, y)
        x2 = min(width, x + w)
        y2 = min(height, y + h)
        
        if x2 <= x or y2 <= y:
            return False
        
        # Extract region
        region = image[y:y2, x:x2]
        
        if region.size == 0:
            return False
        
        # Convert to grayscale if needed
        if len(region.shape) == 3:
            gray_region = cv2.cvtColor(region, cv2.COLOR_BGR2GRAY)
        else:
            gray_region = region
        
        # Calculate gradient magnitude
        grad_x = cv2.Sobel(gray_region, cv2.CV_64F, 1, 0, ksize=3)
        grad_y = cv2.Sobel(gray_region, cv2.CV_64F, 0, 1, ksize=3)
        gradient_magnitude = np.sqrt(grad_x**2 + grad_y**2)
        
        # Check if maximum gradient exceeds threshold
        max_gradient = np.max(gradient_magnitude)
        
        return max_gradient > self._color_discontinuity_threshold

    def determine_quality_level(
        self,
        coverage: float,
        has_artifacts: bool,
        artifact_count: int
    ) -> QualityLevel:
        """Determine the overall quality level based on metrics.
        
        Quality levels are determined as follows:
        - EXCELLENT: coverage >= 95% and no artifacts
        - GOOD: coverage >= 90% and artifacts <= 2
        - FAIR: coverage >= 70% or artifacts <= 5
        - POOR: coverage < 70% or artifacts > 5
        
        Args:
            coverage: Translation coverage ratio (0.0 to 1.0)
            has_artifacts: Whether artifacts were detected
            artifact_count: Number of artifact locations
            
        Returns:
            QualityLevel enum value
        """
        if coverage >= 0.95 and not has_artifacts:
            return QualityLevel.EXCELLENT
        elif coverage >= 0.90 and artifact_count <= 2:
            return QualityLevel.GOOD
        elif coverage >= 0.70 or artifact_count <= 5:
            return QualityLevel.FAIR
        else:
            return QualityLevel.POOR
    
    def validate(
        self,
        original_image: np.ndarray,
        output_image: np.ndarray,
        regions: List[TextRegion],
        results: List[TranslationResult]
    ) -> QualityReport:
        """Validate the quality of a translated image.
        
        Performs comprehensive quality validation including:
        - Translation coverage calculation
        - Artifact detection (if enabled)
        - Overall quality assessment
        
        Args:
            original_image: Original input image
            output_image: Translated output image
            regions: List of text regions detected
            results: List of translation results
            
        Returns:
            QualityReport containing all quality metrics
            
        _Requirements: 9.1, 9.3, 9.4, 9.5_
        """
        logger.info("Starting quality validation")
        
        # Calculate translation coverage
        coverage = self.calculate_coverage(regions, results)
        
        # Get failed regions
        failed_regions = self.get_failed_regions(regions, results)
        
        # Check for artifacts if enabled
        if self.check_artifacts_enabled and output_image is not None:
            has_artifacts, artifact_locations = self.check_artifacts(output_image)
        else:
            has_artifacts = False
            artifact_locations = []
        
        # Determine overall quality level
        quality_level = self.determine_quality_level(
            coverage,
            has_artifacts,
            len(artifact_locations)
        )
        
        # Calculate counts
        total_regions = len(regions)
        translated_regions = sum(1 for r in results if r.success)
        
        # Create quality report
        report = QualityReport(
            translation_coverage=coverage,
            total_regions=total_regions,
            translated_regions=translated_regions,
            failed_regions=failed_regions,
            has_artifacts=has_artifacts,
            artifact_locations=artifact_locations,
            overall_quality=quality_level
        )
        
        logger.info(
            f"Quality validation complete: "
            f"coverage={coverage:.2%}, "
            f"artifacts={len(artifact_locations)}, "
            f"quality={quality_level.value}"
        )
        
        return report
    
    def generate_report_summary(self, report: QualityReport) -> str:
        """Generate a human-readable summary of the quality report.
        
        Args:
            report: QualityReport to summarize
            
        Returns:
            Formatted string summary
        """
        summary_lines = [
            "=" * 50,
            "Quality Report Summary",
            "=" * 50,
            f"Translation Coverage: {report.translation_coverage:.1%}",
            f"Total Regions: {report.total_regions}",
            f"Translated Regions: {report.translated_regions}",
            f"Failed Regions: {len(report.failed_regions)}",
            f"Artifacts Detected: {'Yes' if report.has_artifacts else 'No'}",
            f"Artifact Count: {len(report.artifact_locations)}",
            f"Overall Quality: {report.overall_quality.value.upper()}",
            "=" * 50,
        ]
        
        return "\n".join(summary_lines)
