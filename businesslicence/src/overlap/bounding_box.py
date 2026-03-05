"""Bounding box data structure for overlap detection.

This module provides the BoundingBox dataclass for representing rectangular
regions in the image translation system.
"""

from dataclasses import dataclass


@dataclass
class BoundingBox:
    """Represents a rectangular bounding box.
    
    A bounding box is defined by its top-left corner (x, y) and its
    dimensions (width, height). All values should be non-negative.
    
    Attributes:
        x: X-coordinate of the top-left corner (pixels)
        y: Y-coordinate of the top-left corner (pixels)
        width: Width of the bounding box (pixels)
        height: Height of the bounding box (pixels)
    
    Properties:
        x2: X-coordinate of the bottom-right corner
        y2: Y-coordinate of the bottom-right corner
    
    Example:
        >>> box = BoundingBox(x=10, y=20, width=100, height=50)
        >>> box.x2
        110
        >>> box.y2
        70
        >>> moved = box.move_right(5)
        >>> moved.x
        15
    """
    
    x: float
    y: float
    width: float
    height: float
    
    def __post_init__(self):
        """Validate bounding box parameters.
        
        Raises:
            ValueError: If width or height is negative
        """
        if self.width < 0:
            raise ValueError(f"Width must be non-negative, got {self.width}")
        if self.height < 0:
            raise ValueError(f"Height must be non-negative, got {self.height}")
    
    @property
    def x2(self) -> float:
        """Get the x-coordinate of the bottom-right corner.
        
        Returns:
            X-coordinate of bottom-right corner (x + width)
        """
        return self.x + self.width
    
    @property
    def y2(self) -> float:
        """Get the y-coordinate of the bottom-right corner.
        
        Returns:
            Y-coordinate of bottom-right corner (y + height)
        """
        return self.y + self.height
    
    def move_right(self, offset: float) -> 'BoundingBox':
        """Move the bounding box horizontally to the right.
        
        Creates a new BoundingBox with the same dimensions but shifted
        horizontally by the specified offset. The y-coordinate and
        dimensions remain unchanged.
        
        Args:
            offset: Number of pixels to move right (can be negative to move left)
        
        Returns:
            New BoundingBox instance with adjusted x-coordinate
        
        Example:
            >>> box = BoundingBox(x=10, y=20, width=100, height=50)
            >>> moved = box.move_right(5)
            >>> moved.x
            15
            >>> moved.y
            20
        """
        return BoundingBox(
            x=self.x + offset,
            y=self.y,
            width=self.width,
            height=self.height
        )
