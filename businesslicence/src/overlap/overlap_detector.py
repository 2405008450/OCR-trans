"""Overlap detection for bounding boxes.

This module provides the OverlapDetector class for detecting overlaps
between rectangular bounding boxes using the Separating Axis Theorem.
"""

from src.overlap.bounding_box import BoundingBox


class OverlapDetector:
    """Detects overlaps between bounding boxes.
    
    Uses the Separating Axis Theorem (SAT) to efficiently determine if
    two rectangular bounding boxes overlap. Two boxes overlap if they
    share any interior area. Edge or corner touching is not considered
    an overlap.
    
    The SAT states that two convex shapes do not overlap if there exists
    a separating axis along which their projections do not overlap. For
    axis-aligned rectangles, we only need to check the x and y axes.
    
    Example:
        >>> detector = OverlapDetector()
        >>> box1 = BoundingBox(x=0, y=0, width=10, height=10)
        >>> box2 = BoundingBox(x=5, y=5, width=10, height=10)
        >>> detector.detect_overlap(box1, box2)
        True
        >>> box3 = BoundingBox(x=10, y=0, width=10, height=10)
        >>> detector.detect_overlap(box1, box3)
        False
    """
    
    @staticmethod
    def detect_overlap(box1: BoundingBox, box2: BoundingBox) -> bool:
        """Detect if two bounding boxes overlap.
        
        Two boxes overlap if they share any interior area. Boxes that only
        touch at edges or corners are not considered overlapping.
        
        Uses the Separating Axis Theorem: if the boxes are separated on
        either the x-axis or y-axis, they do not overlap.
        
        Args:
            box1: First bounding box
            box2: Second bounding box
        
        Returns:
            True if the boxes overlap (share interior area), False otherwise
        
        Example:
            >>> box1 = BoundingBox(x=0, y=0, width=10, height=10)
            >>> box2 = BoundingBox(x=5, y=5, width=10, height=10)
            >>> OverlapDetector.detect_overlap(box1, box2)
            True
            
            >>> # Edge touching - not overlapping
            >>> box3 = BoundingBox(x=10, y=0, width=10, height=10)
            >>> OverlapDetector.detect_overlap(box1, box3)
            False
            
            >>> # One box inside another
            >>> box4 = BoundingBox(x=2, y=2, width=5, height=5)
            >>> OverlapDetector.detect_overlap(box1, box4)
            True
        """
        # Check for separation on x-axis
        # Boxes are separated if one ends before the other starts
        if box1.x2 <= box2.x or box2.x2 <= box1.x:
            return False
        
        # Check for separation on y-axis
        if box1.y2 <= box2.y or box2.y2 <= box1.y:
            return False
        
        # Not separated on either axis, so they overlap
        return True
