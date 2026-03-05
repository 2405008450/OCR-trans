"""Multi-parameter OCR recognition with voting mechanism.

This module implements a voting-based OCR strategy that runs OCR with
multiple parameter sets and combines results to improve accuracy.
"""

import logging
from typing import List, Dict, Tuple, Optional
from collections import Counter
import numpy as np

from src.config import ConfigManager
from src.models import TextRegion

logger = logging.getLogger(__name__)


class MultiParamRecognizer:
    """Multi-parameter OCR recognizer with voting mechanism.
    
    Runs OCR with different parameter combinations and uses voting
    to select the most reliable results.
    
    Attributes:
        config: Configuration manager instance
        enabled: Whether multi-param recognition is enabled
        param_sets: List of parameter sets to try
        voting_strategy: Strategy for combining results
    """
    
    def __init__(self, config: ConfigManager):
        """Initialize the multi-parameter recognizer.
        
        Args:
            config: Configuration manager instance
        """
        self.config = config
        self.enabled = config.get('ocr.multi_param.enabled', False)
        self.voting_strategy = config.get('ocr.multi_param.voting_strategy', 'confidence')
        self.min_confidence_diff = config.get('ocr.multi_param.min_confidence_diff', 0.1)
        
        # Define parameter sets to try
        self.param_sets = self._load_param_sets()
        
        logger.info(
            f"MultiParamRecognizer initialized: "
            f"enabled={self.enabled}, "
            f"param_sets={len(self.param_sets)}, "
            f"voting_strategy={self.voting_strategy}"
        )
    
    def _load_param_sets(self) -> List[Dict]:
        """Load parameter sets from configuration.
        
        Returns:
            List of parameter dictionaries
        """
        # Check if custom param sets are defined in config
        custom_sets = self.config.get('ocr.multi_param.param_sets', [])
        
        if custom_sets:
            logger.info(f"Loaded {len(custom_sets)} custom parameter sets")
            return custom_sets
        
        # Default parameter sets
        default_sets = [
            {
                'name': 'standard',
                'det_db_thresh': 0.3,
                'det_db_box_thresh': 0.5,
                'det_db_unclip_ratio': 1.5,
            },
            {
                'name': 'sensitive',
                'det_db_thresh': 0.2,
                'det_db_box_thresh': 0.4,
                'det_db_unclip_ratio': 2.0,
            },
            {
                'name': 'strict',
                'det_db_thresh': 0.4,
                'det_db_box_thresh': 0.6,
                'det_db_unclip_ratio': 1.2,
            },
        ]
        
        logger.info(f"Using {len(default_sets)} default parameter sets")
        return default_sets
    
    def recognize_with_voting(
        self,
        image: np.ndarray,
        ocr_engine
    ) -> List[TextRegion]:
        """Run OCR with multiple parameters and vote on results.
        
        Args:
            image: Input image
            ocr_engine: OCR engine instance
            
        Returns:
            List of TextRegion objects with voted results
        """
        if not self.enabled or len(self.param_sets) <= 1:
            # Multi-param disabled or only one param set
            return []
        
        logger.info(f"Running multi-parameter OCR with {len(self.param_sets)} param sets")
        
        # Store original parameters
        original_params = self._save_original_params(ocr_engine)
        
        # Run OCR with each parameter set
        all_results = []
        
        for i, param_set in enumerate(self.param_sets):
            param_name = param_set.get('name', f'set_{i}')
            logger.info(f"Running OCR with param set '{param_name}'")
            
            # Apply parameters
            self._apply_params(ocr_engine, param_set)
            
            # Reinitialize OCR with new parameters
            ocr_engine._initialize_ocr()
            
            # Run OCR
            try:
                regions = ocr_engine.detect_text(image)
                all_results.append({
                    'name': param_name,
                    'regions': regions,
                    'params': param_set
                })
                logger.info(f"  Detected {len(regions)} regions with '{param_name}'")
            except Exception as e:
                logger.warning(f"  Failed with param set '{param_name}': {e}")
                continue
        
        # Restore original parameters
        self._restore_original_params(ocr_engine, original_params)
        ocr_engine._initialize_ocr()
        
        if not all_results:
            logger.warning("All parameter sets failed")
            return []
        
        # Vote on results
        voted_regions = self._vote_on_results(all_results)
        
        logger.info(f"Voting complete: {len(voted_regions)} final regions")
        return voted_regions
    
    def _save_original_params(self, ocr_engine) -> Dict:
        """Save original OCR parameters.
        
        Args:
            ocr_engine: OCR engine instance
            
        Returns:
            Dictionary of original parameters
        """
        return {
            'det_db_thresh': getattr(ocr_engine, '_portrait_det_params', {}).get('det_db_thresh'),
            'det_db_box_thresh': getattr(ocr_engine, '_portrait_det_params', {}).get('det_db_box_thresh'),
            'det_db_unclip_ratio': getattr(ocr_engine, '_portrait_det_params', {}).get('det_db_unclip_ratio'),
        }
    
    def _apply_params(self, ocr_engine, param_set: Dict) -> None:
        """Apply parameter set to OCR engine.
        
        Args:
            ocr_engine: OCR engine instance
            param_set: Parameter dictionary
        """
        if not hasattr(ocr_engine, '_portrait_det_params'):
            ocr_engine._portrait_det_params = {}
        
        for key, value in param_set.items():
            if key != 'name':
                ocr_engine._portrait_det_params[key] = value
    
    def _restore_original_params(self, ocr_engine, original_params: Dict) -> None:
        """Restore original OCR parameters.
        
        Args:
            ocr_engine: OCR engine instance
            original_params: Dictionary of original parameters
        """
        if not hasattr(ocr_engine, '_portrait_det_params'):
            ocr_engine._portrait_det_params = {}
        
        for key, value in original_params.items():
            if value is not None:
                ocr_engine._portrait_det_params[key] = value
    
    def _vote_on_results(self, all_results: List[Dict]) -> List[TextRegion]:
        """Vote on OCR results from multiple parameter sets.
        
        Args:
            all_results: List of result dictionaries
            
        Returns:
            List of voted TextRegion objects
        """
        if self.voting_strategy == 'confidence':
            return self._vote_by_confidence(all_results)
        elif self.voting_strategy == 'majority':
            return self._vote_by_majority(all_results)
        elif self.voting_strategy == 'union':
            return self._vote_by_union(all_results)
        else:
            logger.warning(f"Unknown voting strategy: {self.voting_strategy}")
            return self._vote_by_confidence(all_results)
    
    def _vote_by_confidence(self, all_results: List[Dict]) -> List[TextRegion]:
        """Vote by selecting regions with highest confidence.
        
        For each spatial location, select the region with highest confidence.
        
        Args:
            all_results: List of result dictionaries
            
        Returns:
            List of voted TextRegion objects
        """
        # Collect all regions
        all_regions = []
        for result in all_results:
            all_regions.extend(result['regions'])
        
        if not all_regions:
            return []
        
        # Group regions by spatial location
        region_groups = self._group_by_location(all_regions)
        
        # Select best region from each group
        voted_regions = []
        for group in region_groups:
            # Sort by confidence
            group.sort(key=lambda r: r.confidence, reverse=True)
            best_region = group[0]
            
            # Check if confidence is significantly higher
            if len(group) > 1:
                second_best = group[1]
                conf_diff = best_region.confidence - second_best.confidence
                
                if conf_diff < self.min_confidence_diff:
                    # Confidence too close, use text voting
                    best_region = self._vote_by_text(group)
            
            voted_regions.append(best_region)
        
        logger.info(f"Confidence voting: {len(region_groups)} groups -> {len(voted_regions)} regions")
        return voted_regions
    
    def _vote_by_majority(self, all_results: List[Dict]) -> List[TextRegion]:
        """Vote by majority text consensus.
        
        For each spatial location, select the text that appears most frequently.
        
        Args:
            all_results: List of result dictionaries
            
        Returns:
            List of voted TextRegion objects
        """
        # Collect all regions
        all_regions = []
        for result in all_results:
            all_regions.extend(result['regions'])
        
        if not all_regions:
            return []
        
        # Group regions by spatial location
        region_groups = self._group_by_location(all_regions)
        
        # Select most common text from each group
        voted_regions = []
        for group in region_groups:
            voted_region = self._vote_by_text(group)
            voted_regions.append(voted_region)
        
        logger.info(f"Majority voting: {len(region_groups)} groups -> {len(voted_regions)} regions")
        return voted_regions
    
    def _vote_by_union(self, all_results: List[Dict]) -> List[TextRegion]:
        """Vote by union (keep all unique regions).
        
        Keeps all regions that don't significantly overlap.
        
        Args:
            all_results: List of result dictionaries
            
        Returns:
            List of all unique TextRegion objects
        """
        # Collect all regions
        all_regions = []
        for result in all_results:
            all_regions.extend(result['regions'])
        
        if not all_regions:
            return []
        
        # Remove duplicates by location
        unique_regions = []
        for region in all_regions:
            is_duplicate = False
            
            for existing in unique_regions:
                if self._regions_overlap(region, existing, threshold=0.5):
                    is_duplicate = True
                    break
            
            if not is_duplicate:
                unique_regions.append(region)
        
        logger.info(f"Union voting: {len(all_regions)} regions -> {len(unique_regions)} unique")
        return unique_regions
    
    def _group_by_location(self, regions: List[TextRegion]) -> List[List[TextRegion]]:
        """Group regions by spatial location.
        
        Args:
            regions: List of TextRegion objects
            
        Returns:
            List of region groups
        """
        if not regions:
            return []
        
        groups = []
        used = set()
        
        for i, region1 in enumerate(regions):
            if i in used:
                continue
            
            group = [region1]
            used.add(i)
            
            for j, region2 in enumerate(regions):
                if j <= i or j in used:
                    continue
                
                # Check if regions overlap significantly
                if self._regions_overlap(region1, region2, threshold=0.5):
                    group.append(region2)
                    used.add(j)
            
            groups.append(group)
        
        return groups
    
    def _regions_overlap(
        self,
        region1: TextRegion,
        region2: TextRegion,
        threshold: float = 0.5
    ) -> bool:
        """Check if two regions overlap significantly.
        
        Args:
            region1: First region
            region2: Second region
            threshold: Overlap ratio threshold (0-1)
            
        Returns:
            True if regions overlap above threshold
        """
        x1_1, y1_1, x2_1, y2_1 = region1.bbox
        x1_2, y1_2, x2_2, y2_2 = region2.bbox
        
        # Calculate intersection
        x1_i = max(x1_1, x1_2)
        y1_i = max(y1_1, y1_2)
        x2_i = min(x2_1, x2_2)
        y2_i = min(y2_1, y2_2)
        
        if x1_i >= x2_i or y1_i >= y2_i:
            return False
        
        intersection_area = (x2_i - x1_i) * (y2_i - y1_i)
        
        area1 = (x2_1 - x1_1) * (y2_1 - y1_1)
        area2 = (x2_2 - x1_2) * (y2_2 - y1_2)
        
        min_area = min(area1, area2)
        
        if min_area == 0:
            return False
        
        overlap_ratio = intersection_area / min_area
        
        return overlap_ratio >= threshold
    
    def _vote_by_text(self, group: List[TextRegion]) -> TextRegion:
        """Vote on text within a group of regions.
        
        Selects the most common text, or highest confidence if tied.
        
        Args:
            group: List of TextRegion objects at same location
            
        Returns:
            Selected TextRegion
        """
        if len(group) == 1:
            return group[0]
        
        # Count text occurrences
        text_counts = Counter(r.text for r in group)
        most_common_text, count = text_counts.most_common(1)[0]
        
        # If there's a clear winner, use it
        if count > len(group) / 2:
            # Find region with this text and highest confidence
            candidates = [r for r in group if r.text == most_common_text]
            candidates.sort(key=lambda r: r.confidence, reverse=True)
            return candidates[0]
        
        # No clear winner, use highest confidence
        group.sort(key=lambda r: r.confidence, reverse=True)
        return group[0]
