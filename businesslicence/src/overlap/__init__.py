"""Overlap detection and position adjustment for text boxes.

This module provides components for detecting and preventing text box overlaps
in the image translation system.
"""

from src.overlap.bounding_box import BoundingBox
from src.overlap.overlap_detector import OverlapDetector
from src.overlap.position_adjuster import PositionAdjuster

__all__ = ['BoundingBox', 'OverlapDetector', 'PositionAdjuster']
