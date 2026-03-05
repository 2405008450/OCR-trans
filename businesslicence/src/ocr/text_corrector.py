"""Text correction module for post-processing OCR results.

This module provides text correction capabilities including:
- Common OCR error correction
- Dictionary-based validation
- Pattern-based validation (dates, numbers, etc.)
- Context-aware correction
"""

import logging
import re
from typing import List, Dict, Optional, Tuple
from datetime import datetime

from src.config import ConfigManager
from src.models import TextRegion

logger = logging.getLogger(__name__)


class TextCorrector:
    """Text corrector for post-processing OCR results.
    
    Applies various correction strategies to improve OCR accuracy:
    - Common character substitutions (e.g., 0 vs O)
    - Dictionary-based validation
    - Pattern validation (dates, IDs, etc.)
    - Field-specific corrections
    
    Attributes:
        config: Configuration manager instance
        enabled: Whether correction is enabled
        correction_rules: Dictionary of correction rules
    """
    
    def __init__(self, config: ConfigManager):
        """Initialize the text corrector.
        
        Args:
            config: Configuration manager instance
        """
        self.config = config
        self.enabled = config.get('ocr.correction.enabled', False)
        
        # Load correction rules
        self.char_substitutions = self._load_char_substitutions()
        self.field_patterns = self._load_field_patterns()
        self.common_words = self._load_common_words()
        
        logger.info(
            f"TextCorrector initialized: "
            f"enabled={self.enabled}, "
            f"substitutions={len(self.char_substitutions)}, "
            f"patterns={len(self.field_patterns)}"
        )
    
    def _load_char_substitutions(self) -> Dict[str, str]:
        """Load common character substitution rules.
        
        Returns:
            Dictionary mapping incorrect chars to correct chars
        """
        # Check for custom rules in config
        custom_rules = self.config.get('ocr.correction.char_substitutions', {})
        
        # Default substitution rules
        default_rules = {
            # Numbers vs Letters
            'O': '0',  # Letter O -> Number 0 (in numeric contexts)
            'l': '1',  # Lowercase L -> Number 1 (in numeric contexts)
            'I': '1',  # Uppercase I -> Number 1 (in numeric contexts)
            'S': '5',  # Letter S -> Number 5 (in numeric contexts)
            'Z': '2',  # Letter Z -> Number 2 (in numeric contexts)
            'B': '8',  # Letter B -> Number 8 (in numeric contexts)
            
            # Common Chinese character confusions
            '囗': '口',  # Similar looking characters
            '巳': '已',
            '己': '已',
            '戸': '户',
            '厂': '广',
        }
        
        # Merge custom rules
        default_rules.update(custom_rules)
        
        return default_rules
    
    def _load_field_patterns(self) -> Dict[str, str]:
        """Load field validation patterns.
        
        Returns:
            Dictionary mapping field names to regex patterns
        """
        # Check for custom patterns in config
        custom_patterns = self.config.get('ocr.correction.field_patterns', {})
        
        # Default patterns
        default_patterns = {
            '统一社会信用代码': r'^\d{18}$',  # 18 digits
            '注册号': r'^\d{13,15}$',  # 13-15 digits
            '成立日期': r'^\d{4}年\d{1,2}月\d{1,2}日$',  # YYYY年MM月DD日
            '营业期限': r'^(\d{4}年\d{1,2}月\d{1,2}日|长期)$',  # Date or "长期"
            '注册资本': r'^.+元.*$',  # Contains "元"
        }
        
        # Merge custom patterns
        default_patterns.update(custom_patterns)
        
        return default_patterns
    
    def _load_common_words(self) -> set:
        """Load common words dictionary.
        
        Returns:
            Set of common words
        """
        # Check for custom dictionary in config
        custom_words = self.config.get('ocr.correction.common_words', [])
        
        # Default common words (business license fields)
        default_words = {
            '营业执照', '统一社会信用代码', '名称', '类型', '法定代表人',
            '注册资本', '成立日期', '营业期限', '经营范围', '住所',
            '登记机关', '有限责任公司', '股份有限公司', '个人独资企业',
            '合伙企业', '个体工商户', '长期', '年', '月', '日',
            '元', '万元', '人民币', '省', '市', '区', '县', '街道',
            '路', '号', '楼', '室', '层', '座', '栋', '单元',
        }
        
        # Add custom words
        default_words.update(custom_words)
        
        return default_words
    
    def correct_regions(self, regions: List[TextRegion]) -> List[TextRegion]:
        """Apply corrections to all text regions.
        
        Args:
            regions: List of TextRegion objects
            
        Returns:
            List of corrected TextRegion objects
        """
        if not self.enabled:
            return regions
        
        logger.info(f"Correcting {len(regions)} text regions")
        
        corrected_regions = []
        correction_count = 0
        
        for region in regions:
            original_text = region.text
            corrected_text = self.correct_text(original_text, region)
            
            if corrected_text != original_text:
                # Create new region with corrected text
                corrected_region = TextRegion(
                    bbox=region.bbox,
                    text=corrected_text,
                    confidence=region.confidence,
                    font_size=region.font_size,
                    angle=region.angle
                )
                
                # Copy additional attributes
                for attr in ['is_field_label', 'is_field_content', 'is_paragraph_merged', 'belongs_to_field']:
                    if hasattr(region, attr):
                        setattr(corrected_region, attr, getattr(region, attr))
                
                corrected_regions.append(corrected_region)
                correction_count += 1
                
                logger.debug(f"Corrected: '{original_text}' -> '{corrected_text}'")
            else:
                corrected_regions.append(region)
        
        if correction_count > 0:
            logger.info(f"Applied {correction_count} corrections")
        
        return corrected_regions
    
    def correct_text(self, text: str, region: Optional[TextRegion] = None) -> str:
        """Correct a single text string.
        
        Args:
            text: Input text
            region: Optional TextRegion for context
            
        Returns:
            Corrected text
        """
        if not text:
            return text
        
        corrected = text
        
        # 1. Apply field-specific corrections
        if region and hasattr(region, 'is_field_label') and region.is_field_label:
            corrected = self._correct_field_label(corrected)
        elif region and hasattr(region, 'belongs_to_field') and region.belongs_to_field:
            corrected = self._correct_field_content(corrected, region.belongs_to_field)
        
        # 2. Apply pattern-based corrections
        corrected = self._correct_by_pattern(corrected)
        
        # 3. Apply character substitutions (context-aware)
        corrected = self._correct_characters(corrected)
        
        # 4. Apply dictionary validation
        corrected = self._validate_with_dictionary(corrected)
        
        return corrected
    
    def _correct_field_label(self, text: str) -> str:
        """Correct field label text.
        
        Args:
            text: Field label text
            
        Returns:
            Corrected text
        """
        # Common field label corrections
        label_corrections = {
            '名 称': '名称',
            '类 型': '类型',
            '住 所': '住所',
            '法定代 表人': '法定代表人',
            '法定 代表人': '法定代表人',
            '注册 资本': '注册资本',
            '成立 日期': '成立日期',
            '营业 期限': '营业期限',
            '经营 范围': '经营范围',
            '经营范 围': '经营范围',
            '登记 机关': '登记机关',
        }
        
        # Remove extra spaces
        text_no_space = text.replace(' ', '')
        
        # Check if it matches a known label
        if text_no_space in label_corrections.values():
            return text_no_space
        
        # Check if it's a corrupted version
        for corrupted, correct in label_corrections.items():
            if text == corrupted or text_no_space == corrupted.replace(' ', ''):
                return correct
        
        return text
    
    def _correct_field_content(self, text: str, field_name: str) -> str:
        """Correct field content based on field type.
        
        Args:
            text: Field content text
            field_name: Name of the field
            
        Returns:
            Corrected text
        """
        # Unified credit code: should be 18 digits/letters
        if '统一社会信用代码' in field_name or '信用代码' in field_name:
            return self._correct_credit_code(text)
        
        # Date fields
        elif '日期' in field_name:
            return self._correct_date(text)
        
        # Capital fields
        elif '资本' in field_name:
            return self._correct_capital(text)
        
        # Period fields
        elif '期限' in field_name:
            return self._correct_period(text)
        
        return text
    
    def _correct_credit_code(self, text: str) -> str:
        """Correct unified social credit code.
        
        Args:
            text: Credit code text
            
        Returns:
            Corrected credit code
        """
        # Remove spaces and special characters
        cleaned = re.sub(r'[^0-9A-Z]', '', text.upper())
        
        # Apply character substitutions for alphanumeric codes
        corrected = ''
        for char in cleaned:
            # In credit codes, these are typically numbers
            if char == 'O':
                corrected += '0'
            elif char == 'I':
                corrected += '1'
            elif char == 'S':
                corrected += '5'
            elif char == 'Z':
                corrected += '2'
            else:
                corrected += char
        
        # Validate length (should be 18)
        if len(corrected) == 18:
            return corrected
        
        # If length is wrong, return original
        return text
    
    def _correct_date(self, text: str) -> str:
        """Correct date format.
        
        Args:
            text: Date text
            
        Returns:
            Corrected date
        """
        # Try to parse and reformat date
        # Expected format: YYYY年MM月DD日
        
        # Extract numbers
        numbers = re.findall(r'\d+', text)
        
        if len(numbers) >= 3:
            year, month, day = numbers[0], numbers[1], numbers[2]
            
            # Validate year (should be 1900-2100)
            try:
                year_int = int(year)
                if 1900 <= year_int <= 2100:
                    # Validate month (1-12)
                    month_int = int(month)
                    if 1 <= month_int <= 12:
                        # Validate day (1-31)
                        day_int = int(day)
                        if 1 <= day_int <= 31:
                            return f"{year}年{month:0>2}月{day:0>2}日"
            except ValueError:
                pass
        
        return text
    
    def _correct_capital(self, text: str) -> str:
        """Correct registered capital format.
        
        Args:
            text: Capital text
            
        Returns:
            Corrected capital
        """
        # Ensure it contains "元"
        if '元' not in text:
            # Try to add it
            if re.search(r'\d', text):
                text = text + '元'
        
        # Correct "万元" spacing
        text = re.sub(r'万\s*元', '万元', text)
        
        return text
    
    def _correct_period(self, text: str) -> str:
        """Correct business period format.
        
        Args:
            text: Period text
            
        Returns:
            Corrected period
        """
        # Common corrections
        if '长 期' in text:
            return text.replace('长 期', '长期')
        
        # If it contains date, correct the date part
        if '年' in text and '月' in text:
            return self._correct_date(text)
        
        return text
    
    def _correct_by_pattern(self, text: str) -> str:
        """Apply pattern-based corrections.
        
        Args:
            text: Input text
            
        Returns:
            Corrected text
        """
        # Correct common spacing issues
        text = re.sub(r'(\d)\s+(\d)', r'\1\2', text)  # Remove spaces between digits
        
        # Correct punctuation
        text = text.replace('：', ':')  # Normalize colons
        text = text.replace('（', '(').replace('）', ')')  # Normalize parentheses
        
        return text
    
    def _correct_characters(self, text: str) -> str:
        """Apply character-level corrections.
        
        Context-aware character substitution.
        
        Args:
            text: Input text
            
        Returns:
            Corrected text
        """
        # Check if text is primarily numeric
        digit_count = sum(c.isdigit() for c in text)
        total_count = len(text)
        
        if total_count == 0:
            return text
        
        is_numeric_context = digit_count / total_count > 0.5
        
        if is_numeric_context:
            # Apply number-focused substitutions
            corrected = ''
            for char in text:
                if char in self.char_substitutions:
                    corrected += self.char_substitutions[char]
                else:
                    corrected += char
            return corrected
        
        return text
    
    def _validate_with_dictionary(self, text: str) -> str:
        """Validate text against dictionary.
        
        Args:
            text: Input text
            
        Returns:
            Validated/corrected text
        """
        # Check if text is in common words
        if text in self.common_words:
            return text
        
        # Check for close matches (simple edit distance)
        for word in self.common_words:
            if self._is_close_match(text, word):
                logger.debug(f"Dictionary correction: '{text}' -> '{word}'")
                return word
        
        return text
    
    def _is_close_match(self, text1: str, text2: str, max_distance: int = 1) -> bool:
        """Check if two strings are close matches.
        
        Uses simple edit distance (Levenshtein distance).
        
        Args:
            text1: First string
            text2: Second string
            max_distance: Maximum edit distance
            
        Returns:
            True if strings are close matches
        """
        if abs(len(text1) - len(text2)) > max_distance:
            return False
        
        # Simple check: count different characters
        if len(text1) != len(text2):
            return False
        
        diff_count = sum(c1 != c2 for c1, c2 in zip(text1, text2))
        
        return diff_count <= max_distance
