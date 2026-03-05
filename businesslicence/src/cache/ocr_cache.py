"""OCR result caching for image translation system.

This module provides caching functionality for OCR results using
image hashing as cache keys. This avoids redundant OCR processing
for identical or similar images.

_Requirements: 14.2_
"""

import hashlib
import logging
from collections import OrderedDict
from functools import wraps
from typing import Callable, List, Optional, Tuple, Any

import numpy as np

from src.models import TextRegion


logger = logging.getLogger(__name__)


class ImageHasher:
    """Computes hash values for images to use as cache keys.
    
    Uses a combination of image content hashing and perceptual hashing
    to generate unique identifiers for images. The hash is computed
    from a downsampled version of the image for efficiency.
    
    Attributes:
        hash_size: Size of the downsampled image for hashing (default: 16)
    """
    
    def __init__(self, hash_size: int = 16):
        """Initialize the image hasher.
        
        Args:
            hash_size: Size to downsample images to before hashing.
                      Larger values provide more unique hashes but are slower.
        """
        self.hash_size = hash_size
    
    def compute_hash(self, image: np.ndarray) -> str:
        """Compute a hash value for an image.
        
        The hash is computed by:
        1. Converting to grayscale if needed
        2. Downsampling to hash_size x hash_size
        3. Computing MD5 hash of the pixel values
        
        Args:
            image: Input image as numpy array
            
        Returns:
            Hexadecimal hash string
        """
        if image is None or not isinstance(image, np.ndarray):
            return ""
        
        if image.size == 0:
            return ""
        
        try:
            # Convert to grayscale if color image
            if len(image.shape) == 3:
                # Simple grayscale conversion using luminosity method
                gray = np.dot(image[..., :3], [0.299, 0.587, 0.114]).astype(np.uint8)
            else:
                gray = image
            
            # Downsample using simple averaging
            h, w = gray.shape[:2]
            if h > self.hash_size or w > self.hash_size:
                # Calculate block sizes
                block_h = max(1, h // self.hash_size)
                block_w = max(1, w // self.hash_size)
                
                # Truncate to exact multiple of block size
                new_h = block_h * self.hash_size
                new_w = block_w * self.hash_size
                truncated = gray[:new_h, :new_w]
                
                # Reshape and compute mean for each block
                reshaped = truncated.reshape(
                    self.hash_size, block_h, 
                    self.hash_size, block_w
                )
                downsampled = reshaped.mean(axis=(1, 3)).astype(np.uint8)
            else:
                downsampled = gray
            
            # Compute MD5 hash of the downsampled image
            hash_bytes = hashlib.md5(downsampled.tobytes()).hexdigest()
            
            return hash_bytes
            
        except Exception as e:
            logger.warning(f"Failed to compute image hash: {e}")
            return ""
    
    def compute_hash_with_shape(self, image: np.ndarray) -> str:
        """Compute a hash that includes image shape information.
        
        This provides a more unique hash by including the original
        image dimensions in the hash computation.
        
        Args:
            image: Input image as numpy array
            
        Returns:
            Hexadecimal hash string including shape info
        """
        content_hash = self.compute_hash(image)
        if not content_hash:
            return ""
        
        # Include shape in hash
        shape_str = f"{image.shape}"
        combined = f"{content_hash}_{shape_str}"
        
        return hashlib.md5(combined.encode()).hexdigest()


class OCRCache:
    """LRU cache for OCR results.
    
    Caches OCR detection results using image hashes as keys.
    Implements a Least Recently Used (LRU) eviction policy to
    limit memory usage.
    
    Attributes:
        max_size: Maximum number of entries in the cache
        hasher: ImageHasher instance for computing cache keys
        cache: OrderedDict storing cached results
        hits: Number of cache hits
        misses: Number of cache misses
    
    _Requirements: 14.2_
    """
    
    def __init__(self, max_size: int = 100, hash_size: int = 16):
        """Initialize the OCR cache.
        
        Args:
            max_size: Maximum number of cached results (default: 100)
            hash_size: Size for image hashing (default: 16)
        """
        self.max_size = max_size
        self.hasher = ImageHasher(hash_size)
        self.cache: OrderedDict[str, List[TextRegion]] = OrderedDict()
        self.hits = 0
        self.misses = 0
        
        logger.debug(f"OCRCache initialized with max_size={max_size}")
    
    def get(self, image: np.ndarray) -> Optional[List[TextRegion]]:
        """Get cached OCR results for an image.
        
        Args:
            image: Input image
            
        Returns:
            Cached list of TextRegion objects, or None if not cached
        """
        key = self.hasher.compute_hash_with_shape(image)
        if not key:
            self.misses += 1
            return None
        
        if key in self.cache:
            # Move to end (most recently used)
            self.cache.move_to_end(key)
            self.hits += 1
            logger.debug(f"Cache hit for image hash: {key[:8]}...")
            # Return a copy to prevent modification of cached data
            return [self._copy_region(r) for r in self.cache[key]]
        
        self.misses += 1
        return None
    
    def put(self, image: np.ndarray, regions: List[TextRegion]) -> None:
        """Store OCR results in the cache.
        
        Args:
            image: Input image
            regions: List of detected TextRegion objects
        """
        key = self.hasher.compute_hash_with_shape(image)
        if not key:
            return
        
        # Store a copy to prevent external modification
        cached_regions = [self._copy_region(r) for r in regions]
        
        # If key exists, update and move to end
        if key in self.cache:
            self.cache.move_to_end(key)
            self.cache[key] = cached_regions
            return
        
        # Add new entry
        self.cache[key] = cached_regions
        
        # Evict oldest entries if over capacity
        while len(self.cache) > self.max_size:
            oldest_key = next(iter(self.cache))
            del self.cache[oldest_key]
            logger.debug(f"Evicted cache entry: {oldest_key[:8]}...")
        
        logger.debug(f"Cached OCR results for image hash: {key[:8]}...")
    
    def _copy_region(self, region: TextRegion) -> TextRegion:
        """Create a copy of a TextRegion.
        
        Args:
            region: TextRegion to copy
            
        Returns:
            New TextRegion with same values
        """
        return TextRegion(
            bbox=region.bbox,
            text=region.text,
            confidence=region.confidence,
            font_size=region.font_size,
            angle=region.angle
        )
    
    def clear(self) -> None:
        """Clear all cached entries."""
        self.cache.clear()
        logger.debug("OCR cache cleared")
    
    def get_stats(self) -> dict:
        """Get cache statistics.
        
        Returns:
            Dictionary containing:
                - size: Current number of cached entries
                - max_size: Maximum cache size
                - hits: Number of cache hits
                - misses: Number of cache misses
                - hit_rate: Cache hit rate (0-1)
        """
        total = self.hits + self.misses
        hit_rate = self.hits / total if total > 0 else 0.0
        
        return {
            'size': len(self.cache),
            'max_size': self.max_size,
            'hits': self.hits,
            'misses': self.misses,
            'hit_rate': hit_rate
        }
    
    def reset_stats(self) -> None:
        """Reset cache statistics."""
        self.hits = 0
        self.misses = 0


def cached_ocr(cache: OCRCache):
    """Decorator to add caching to OCR detection methods.
    
    This decorator wraps an OCR detection function to check the cache
    before performing detection and store results after detection.
    
    Args:
        cache: OCRCache instance to use for caching
        
    Returns:
        Decorator function
        
    Example:
        cache = OCRCache(max_size=50)
        
        @cached_ocr(cache)
        def detect_text(self, image):
            # OCR detection logic
            return regions
    
    _Requirements: 14.2_
    """
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(self, image: np.ndarray, *args, **kwargs) -> List[TextRegion]:
            # Check cache first
            cached_result = cache.get(image)
            if cached_result is not None:
                return cached_result
            
            # Perform OCR detection
            result = func(self, image, *args, **kwargs)
            
            # Cache the result
            cache.put(image, result)
            
            return result
        
        return wrapper
    
    return decorator
