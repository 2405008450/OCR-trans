"""Position adjustment for text boxes to avoid overlaps.

This module provides the PositionAdjuster class for adjusting text box
positions to prevent overlaps with normalized field boxes.
"""

import logging
from typing import List, Optional

from src.overlap.bounding_box import BoundingBox
from src.overlap.overlap_detector import OverlapDetector


logger = logging.getLogger(__name__)


class PositionAdjuster:
    """Adjusts text box positions to avoid overlaps with obstacles.
    
    The PositionAdjuster moves text boxes horizontally to the right until
    they no longer overlap with any obstacle boxes (normalized field boxes).
    It uses an iterative approach with configurable step size and maximum
    iterations to prevent infinite loops.
    
    Attributes:
        step_size: Number of pixels to move right in each iteration
        max_iterations: Maximum number of adjustment iterations
        detector: OverlapDetector instance for checking overlaps
    
    Example:
        >>> adjuster = PositionAdjuster(step_size=5.0, max_iterations=1000)
        >>> text_box = BoundingBox(x=10, y=20, width=50, height=20)
        >>> obstacle = BoundingBox(x=15, y=20, width=60, height=20)
        >>> adjusted = adjuster.adjust_position(text_box, [obstacle])
        >>> adjusted.x > text_box.x  # Moved to the right
        True
    """
    
    def __init__(self, step_size: float = 5.0, max_iterations: int = 1000):
        """Initialize the position adjuster.
        
        Args:
            step_size: Number of pixels to move right in each iteration.
                      Smaller values give more precise positioning but may
                      require more iterations. Default is 5.0 pixels.
            max_iterations: Maximum number of iterations to prevent infinite
                          loops. Default is 1000 iterations.
        
        Raises:
            ValueError: If step_size is not positive or max_iterations < 1
        """
        if step_size <= 0:
            raise ValueError(f"step_size must be positive, got {step_size}")
        if max_iterations < 1:
            raise ValueError(f"max_iterations must be at least 1, got {max_iterations}")
        
        self.step_size = step_size
        self.max_iterations = max_iterations
        self.detector = OverlapDetector()
        
        logger.debug(
            f"PositionAdjuster initialized: step_size={step_size}, "
            f"max_iterations={max_iterations}"
        )
    
    def adjust_position(
        self,
        text_box: BoundingBox,
        obstacles: List[BoundingBox],
        image_width: Optional[float] = None
    ) -> BoundingBox:
        """Adjust text box position to avoid overlaps with obstacles.
        
        Moves the text box horizontally to the right until it no longer
        overlaps with any obstacle. The vertical position (y-coordinate)
        and dimensions remain unchanged.
        
        Strategy:
        1. Check if text_box overlaps with any obstacle
        2. If overlap exists, move text_box right by step_size
        3. Repeat until no overlap or max_iterations reached
        4. Stop if text_box exceeds image_width (if provided)
        
        Args:
            text_box: The text box to adjust
            obstacles: List of obstacle boxes (normalized field boxes) to avoid
            image_width: Optional image width for boundary checking. If provided,
                        adjustment stops when text_box would exceed this width.
        
        Returns:
            Adjusted BoundingBox. If no obstacles or no overlap, returns a copy
            of the original text_box. If adjustment fails (max iterations or
            boundary exceeded), returns the best position found.
        
        Example:
            >>> adjuster = PositionAdjuster(step_size=5.0)
            >>> text_box = BoundingBox(x=10, y=20, width=50, height=20)
            >>> obstacle = BoundingBox(x=15, y=20, width=60, height=20)
            >>> adjusted = adjuster.adjust_position(text_box, [obstacle])
            >>> adjusted.x >= 75  # Moved past the obstacle
            True
        """
        # Handle empty obstacles list
        if not obstacles:
            logger.debug("No obstacles, returning original position")
            return BoundingBox(
                x=text_box.x,
                y=text_box.y,
                width=text_box.width,
                height=text_box.height
            )
        
        logger.debug(
            f"Adjusting position for text_box at ({text_box.x}, {text_box.y}), "
            f"size=({text_box.width}x{text_box.height}), "
            f"obstacles={len(obstacles)}"
        )
        
        adjusted_box = BoundingBox(
            x=text_box.x,
            y=text_box.y,
            width=text_box.width,
            height=text_box.height
        )
        last_valid_box = adjusted_box
        iterations = 0
        
        while iterations < self.max_iterations:
            # Check if current position overlaps with any obstacle
            has_overlap = any(
                self.detector.detect_overlap(adjusted_box, obstacle)
                for obstacle in obstacles
            )
            
            if not has_overlap:
                # No overlap, position is good
                logger.debug(
                    f"Position adjusted successfully after {iterations} iterations: "
                    f"({text_box.x}, {text_box.y}) -> ({adjusted_box.x}, {adjusted_box.y})"
                )
                break
            
            # Save current position as last valid before moving
            last_valid_box = adjusted_box
            
            # Move right
            adjusted_box = adjusted_box.move_right(self.step_size)
            
            # Check image boundary
            if image_width is not None and adjusted_box.x2 > image_width:
                logger.warning(
                    f"Adjusted text_box exceeds image boundary: "
                    f"x2={adjusted_box.x2}, image_width={image_width}. "
                    f"Stopping adjustment at iteration {iterations}"
                )
                # Return last valid position that didn't exceed boundary
                return last_valid_box
            
            iterations += 1
        
        if iterations >= self.max_iterations:
            logger.warning(
                f"Reached max iterations ({self.max_iterations}) "
                f"while adjusting text_box at ({text_box.x}, {text_box.y}). "
                f"Final position: ({adjusted_box.x}, {adjusted_box.y})"
            )
        
        return adjusted_box
