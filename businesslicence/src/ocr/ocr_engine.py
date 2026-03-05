"""OCR Engine for text detection and recognition.

This module provides the OCREngine class that wraps PaddleOCR for detecting
and recognizing text in images, with support for region merging and filtering.

_Requirements: 14.2_
"""

import logging
import math
from typing import List, Optional, Tuple

import numpy as np

from src.config import ConfigManager
from src.exceptions import OCRError
from src.models import TextRegion
from src.cache import OCRCache

logger = logging.getLogger(__name__)


class OCREngine:
    """OCR engine for detecting and recognizing text in images.
    
    This class wraps PaddleOCR and provides additional functionality for:
    - Text detection with bounding boxes
    - Region merging for adjacent text areas
    - Filtering by confidence and area thresholds
    - Caching of OCR results to avoid redundant processing
    
    Attributes:
        config: Configuration manager instance
        confidence_threshold: Minimum confidence for text detection
        min_text_area: Minimum area in pixels for text regions
        merge_threshold: Distance threshold for merging adjacent regions
        cache: OCRCache instance for caching results (if enabled)
        cache_enabled: Whether caching is enabled
    
    _Requirements: 14.2_
    """
    
    def __init__(self, config: ConfigManager, cache: Optional[OCRCache] = None):
        """Initialize the OCR engine.
        
        Args:
            config: Configuration manager instance
            cache: Optional OCRCache instance. If None and caching is enabled
                   in config, a new cache will be created.
        """
        self.config = config
        self.confidence_threshold = config.get('ocr.confidence_threshold', 0.6)
        self.min_text_area = config.get('ocr.min_text_area', 100)
        self.merge_threshold = config.get('ocr.merge_threshold', 10)
        
        # 文本黑名单模式
        import re
        blacklist_patterns = config.get('ocr.text_blacklist_patterns', [])
        self.text_blacklist = [re.compile(pattern) for pattern in blacklist_patterns]
        if self.text_blacklist:
            logger.info(f"Loaded {len(self.text_blacklist)} text blacklist patterns")
        
        # 区分水平和垂直合并阈值（方案一改进版）
        # 水平阈值：用于合并同一行的文字
        # 垂直阈值：用于合并不同行的文字（设置较小以避免多行合并）
        self.merge_threshold_horizontal = config.get('ocr.merge_threshold_horizontal', self.merge_threshold)
        self.merge_threshold_vertical = config.get('ocr.merge_threshold_vertical', 3)  # 默认3像素，避免多行合并
        
        # 字体大小差异阈值：只有字体大小相近的文字才会合并（默认20%）
        self.font_size_diff_threshold = config.get('ocr.font_size_diff_threshold', 0.2)
        
        # 字段碎片合并规则（营业执照等表单常见字段）
        # 格式: (碎片1, 碎片2, 合并后的完整字段, 最大水平距离)
        self.field_fragment_rules = [
            # 基本信息字段
            ("名", "称", "名称", 150),
            ("类", "型", "类型", 150),
            ("住", "所", "住所", 150),
            ("法定代", "表人", "法定代表人", 150),
            ("法定", "代表人", "法定代表人", 150),
            ("法", "定代表人", "法定代表人", 150),
            ("代", "表人", "法定代表人", 150),  # 新增：处理"代"+"表人"的情况
            ("表", "人", "表人", 150),  # 先合并成"表人"
            ("法定代", "表人", "法定代表人", 150),  # 然后"法定代"+"表人"
            ("注册", "资本", "注册资本", 150),
            ("注", "册资本", "注册资本", 150),
            ("资", "本", "资本", 150),  # 先合并成"资本"
            ("注册", "资本", "注册资本", 150),  # 然后"注册"+"资本"
            ("成立", "日期", "成立日期", 150),
            ("成", "立日期", "成立日期", 150),
            ("立", "日期", "立日期", 150),  # 先合并成"立日期"
            ("成", "立日期", "成立日期", 150),  # 然后"成"+"立日期"
            ("日", "期", "日期", 150),  # 先合并成"日期"
            ("立", "日期", "立日期", 150),  # 然后"立"+"日期"
            # 营业期限的多种合并路径
            ("营业", "期限", "营业期限", 150),
            ("营", "业期限", "营业期限", 150),
            ("业", "期限", "业期限", 150),  # 先合并成"业期限"
            ("营", "业期限", "营业期限", 150),  # 然后"营"+"业期限"
            ("期", "限", "期限", 150),  # 先合并成"期限"
            ("业", "期限", "业期限", 150),  # 然后"业"+"期限"
            # 新增：处理"经"+"营"+"期限"的情况（用户报告的问题）
            ("经", "营", "经营", 150),  # 第一步：合并"经"+"营"
            ("经营", "期限", "经营期限", 150),  # 第二步：合并"经营"+"期限"（会被标准化为"营业期限"）
            ("营", "期限", "营期限", 150),  # 处理"营"+"期限"（会被标准化为"营业期限"）
            ("经营范", "围", "经营范围", 150),
            ("经营", "范围", "经营范围", 150),
            ("经", "营范围", "经营范围", 150),
            ("营", "范围", "营范围", 150),  # 先合并成"营范围"
            ("经", "营范围", "经营范围", 150),  # 然后"经"+"营范围"
            ("范", "围", "范围", 150),  # 先合并成"范围"
            ("营", "范围", "营范围", 150),  # 然后"营"+"范围"
            ("登记", "机关", "登记机关", 150),
            
            # 其他常见字段
            ("统一社会", "信用代码", "统一社会信用代码", 150),
            ("统一", "社会信用代码", "统一社会信用代码", 150),
            ("组织机构", "代码", "组织机构代码", 150),
            ("税务登记", "证号", "税务登记证号", 150),
            ("开户", "银行", "开户银行", 150),
            ("银行", "账号", "银行账号", 150),
        ]
        
        # 字段名标准化映射（部分字段名 → 完整字段名）
        # 用于处理OCR无法识别完整字段名的情况
        self.field_name_normalization = {
            "表人": "法定代表人",
            "代表人": "法定代表人",
            "资本": "注册资本",
            "立日期": "成立日期",
            "业期限": "营业期限",
            "营范围": "经营范围",
            # 新增：处理"经营期限"和"营期限"的标准化（用户报告的问题）
            "经营期限": "营业期限",  # "经"+"营"+"期限"合并后的结果
            "营期限": "营业期限",  # "营"+"期限"合并后的结果
        }
        
        # 从配置文件加载自定义规则（如果有）
        custom_rules = config.get('ocr.field_fragment_rules', [])
        if custom_rules:
            for rule in custom_rules:
                if len(rule) >= 3:
                    # 格式: [碎片1, 碎片2, 完整字段, 最大距离(可选)]
                    max_distance = rule[3] if len(rule) > 3 else 50
                    self.field_fragment_rules.append((rule[0], rule[1], rule[2], max_distance))
            logger.info(f"Loaded {len(custom_rules)} custom field fragment rules")
        
        logger.info(f"Initialized with {len(self.field_fragment_rules)} field fragment rules")
        
        # Initialize caching
        self.cache_enabled = config.get('performance.cache_ocr_results', True)
        cache_size = config.get('performance.ocr_cache_size', 100)
        
        if cache is not None:
            self.cache = cache
        elif self.cache_enabled:
            self.cache = OCRCache(max_size=cache_size)
            logger.info(f"OCR cache enabled with max_size={cache_size}")
        else:
            self.cache = None
            logger.info("OCR cache disabled")
        
        # Initialize enhancement modules
        from src.ocr.image_preprocessor import ImagePreprocessor
        from src.ocr.multi_param_recognizer import MultiParamRecognizer
        from src.ocr.text_corrector import TextCorrector
        
        self.preprocessor = ImagePreprocessor(config)
        self.multi_param_recognizer = MultiParamRecognizer(config)
        self.text_corrector = TextCorrector(config)
        
        # Initialize OCR engine using factory pattern
        self._ocr_adapter = None
        self._initialize_ocr()
    
    def _initialize_ocr(self) -> None:
        """Initialize the OCR engine using factory pattern.
        
        This method creates an OCR engine adapter based on the configuration.
        Supports multiple engines: PaddleOCR, GLM-OCR API, GLM-OCR Local.
        
        Raises:
            OCRError: If OCR engine initialization fails
        """
        try:
            from src.ocr.ocr_engine_factory import OCREngineFactory
            
            # Create OCR engine using factory
            self._ocr_adapter = OCREngineFactory.create_engine(self.config)
            
            # Get engine info for logging
            engine_info = self._ocr_adapter.get_engine_info()
            logger.info(
                f"OCR engine ready: {engine_info['name']} "
                f"(type={engine_info['type']}, capabilities={engine_info.get('capabilities', [])})"
            )
            
        except Exception as e:
            raise OCRError(f"Failed to initialize OCR engine: {e}")

    def detect_text(self, image: np.ndarray) -> List[TextRegion]:
        """Detect text regions in an image.
        
        If caching is enabled, checks the cache first before performing
        OCR detection. Results are cached for future use.
        
        Enhanced with:
        - Image preprocessing for better quality
        - Multi-parameter recognition with voting
        - Text correction for common OCR errors
        
        Args:
            image: Input image as numpy array (BGR or RGB format)
            
        Returns:
            List of TextRegion objects representing detected text areas
            
        Raises:
            OCRError: If OCR processing fails
            
        _Requirements: 14.2_
        """
        if image is None or not isinstance(image, np.ndarray):
            raise OCRError("Invalid image input: must be a numpy array")
        
        if image.size == 0:
            raise OCRError("Invalid image input: image is empty")
        
        # 保存图片尺寸，用于二维码保护
        self._image_height, self._image_width = image.shape[:2]
        
        # Check cache first if enabled
        if self.cache_enabled and self.cache is not None:
            cached_result = self.cache.get(image)
            if cached_result is not None:
                logger.debug("Returning cached OCR results")
                return cached_result
        
        # Step 1: Preprocess image (if enabled)
        preprocessed_image = self.preprocessor.preprocess(image)
        
        # Step 1.5: QR code pre-detection (if enabled) - Requirements 14
        qr_regions = []
        qr_pre_detection_enabled = self.config.get('qr_pre_detection.enabled', True)
        hard_protection = self.config.get('qr_pre_detection.hard_protection', False)
        
        if qr_pre_detection_enabled:
            logger.info("🔍 QR code pre-detection enabled, detecting QR codes before OCR...")
            if hard_protection:
                logger.info("🛡️ Hard protection mode: QR codes will not be masked, translated, or processed")
                qr_regions, _ = self._detect_and_mask_qrcodes(preprocessed_image, mask_qr=False)
            else:
                qr_regions, preprocessed_image = self._detect_and_mask_qrcodes(preprocessed_image, mask_qr=True)
            
            if qr_regions:
                if hard_protection:
                    logger.info(f"✅ Pre-detected {len(qr_regions)} QR code(s) with hard protection (no masking)")
                else:
                    logger.info(f"✅ Pre-detected {len(qr_regions)} QR code(s), masked from OCR")
        
        # Step 2: Try multi-parameter recognition (if enabled)
        multi_param_regions = self.multi_param_recognizer.recognize_with_voting(
            preprocessed_image, self
        )
        
        # Step 3: Standard OCR detection using adapter
        try:
            # Use the OCR adapter's unified interface
            # Returns: List of tuples (box_points, text, confidence)
            ocr_results = self._ocr_adapter.detect_and_recognize(preprocessed_image)
            
            if ocr_results is None or len(ocr_results) == 0:
                logger.info("No text detected in image")
                regions = []
            else:
                regions = []
                
                # Convert OCR results to TextRegion objects
                # The adapter returns a standardized format: [(box_points, text, confidence), ...]
                for box_points, text, confidence in ocr_results:
                    region = self._convert_ocr_result_from_adapter(
                        box_points, text, confidence, preprocessed_image
                    )
                    if region is not None:
                        regions.append(region)
                
                logger.info(f"Detected {len(regions)} text regions")
            
            # Step 4: Merge multi-param results with standard results (if available)
            if multi_param_regions:
                logger.info(f"Merging {len(multi_param_regions)} multi-param regions with {len(regions)} standard regions")
                regions = self._merge_multi_param_results(regions, multi_param_regions)
            
            # Step 4.5: Add pre-detected QR code regions (Requirements 14)
            if qr_regions:
                logger.info(f"Adding {len(qr_regions)} pre-detected QR code region(s) to results")
                regions.extend(qr_regions)
                
                # Step 4.6: In hard protection mode, mark text regions near QR codes as hard protected
                if hard_protection:
                    regions = self._mark_qr_nearby_text_as_protected(regions, qr_regions)
            
            # Step 5: Apply text correction (if enabled)
            regions = self.text_corrector.correct_regions(regions)
            
            # 优先检测并合并"重要提示"（在缓存之前）
            # 这样可以确保"重要提示"的合并优先级最高，不会被其他逻辑干扰
            regions = self._merge_important_notice(regions)
            
            # 拆分字段标签和内容（如果它们被OCR错误地识别成一个区域）
            regions = self._split_field_labels(regions)
            
            # 移除重复或重叠的区域
            regions = self._remove_duplicate_regions(regions)
            
            # Cache the result if enabled
            if self.cache_enabled and self.cache is not None:
                self.cache.put(image, regions)
            
            return regions
            
        except Exception as e:
            raise OCRError(f"OCR processing failed: {e}")
            regions = self._split_field_labels(regions)
            
            # 移除重复或重叠的区域
            regions = self._remove_duplicate_regions(regions)
            
            # Cache the result if enabled
            if self.cache_enabled and self.cache is not None:
                self.cache.put(image, regions)
            
            return regions
            
        except Exception as e:
            raise OCRError(f"OCR processing failed: {e}")
    
    def _convert_paddlex_result(
        self,
        ocr_result,
        image: np.ndarray
    ) -> List[TextRegion]:
        """Convert PaddleX OCRResult object to list of TextRegions.
        
        Args:
            ocr_result: PaddleX OCRResult object (dict-like, contains batch data)
            image: Original image for validation
            
        Returns:
            List of TextRegion objects
        """
        regions = []
        
        try:
            # PaddleX OCRResult 包含批量数据：
            # - dt_polys: 检测框坐标列表
            # - rec_texts: 识别文本列表
            # - rec_scores: 识别
            # 置信度列表
            
            dt_polys = ocr_result.get('dt_polys', [])
            rec_texts = ocr_result.get('rec_texts', [])
            rec_scores = ocr_result.get('rec_scores', [])
            
            logger.debug(f"OCRResult contains {len(dt_polys)} detections, {len(rec_texts)} texts, {len(rec_scores)} scores")
            
            # 确保所有列表长度一致
            min_len = min(len(dt_polys), len(rec_texts), len(rec_scores))
            
            for i in range(min_len):
                try:
                    box_points = dt_polys[i]
                    text = rec_texts[i]
                    confidence = rec_scores[i]
                    
                    # 确保 text 是字符串
                    text = str(text).strip()
                    if not text:
                        continue
                    
                    # 确保 confidence 是数字
                    try:
                        confidence = float(confidence)
                    except (ValueError, TypeError):
                        confidence = 1.0
                    
                    # 确保 box_points 是正确的格式
                    if isinstance(box_points, np.ndarray):
                        box_points = box_points.tolist()
                    
                    # Extract bounding box from polygon points
                    x_coords = [p[0] for p in box_points]
                    y_coords = [p[1] for p in box_points]
                    
                    x1 = int(min(x_coords))
                    y1 = int(min(y_coords))
                    x2 = int(max(x_coords))
                    y2 = int(max(y_coords))
                    
                    # Validate coordinates are within image bounds
                    height, width = image.shape[:2]
                    x1 = max(0, min(x1, width))
                    y1 = max(0, min(y1, height))
                    x2 = max(0, min(x2, width))
                    y2 = max(0, min(y2, height))
                    
                    # Calculate rotation angle from box points
                    angle = self._calculate_angle(box_points)
                    
                    # Estimate font size from box height
                    box_height = y2 - y1
                    font_size = max(8, int(box_height * 0.8))
                    
                    region = TextRegion(
                        bbox=(x1, y1, x2, y2),
                        text=text,
                        confidence=float(confidence),
                        font_size=font_size,
                        angle=angle
                    )
                    regions.append(region)
                    
                except Exception as e:
                    logger.warning(f"Failed to convert OCR item {i}: {e}")
                    continue
            
            logger.debug(f"Converted {len(regions)} regions from PaddleX OCRResult")
            return regions
            
        except Exception as e:
            logger.warning(f"Failed to convert PaddleX OCRResult: {e}")
            import traceback
            logger.debug(f"Traceback: {traceback.format_exc()}")
            return []
    
    def _convert_ocr_result_from_adapter(
        self,
        box_points: List[List[float]],
        text: str,
        confidence: float,
        image: np.ndarray
    ) -> Optional[TextRegion]:
        """Convert OCR adapter result to TextRegion.
        
        This method handles the standardized output from OCR adapters.
        
        Args:
            box_points: List of 4 corner points [[x1,y1], [x2,y2], [x3,y3], [x4,y4]]
            text: Recognized text string
            confidence: Confidence score (0.0 to 1.0)
            image: Original image for validation
            
        Returns:
            TextRegion object or None if conversion fails
        """
        try:
            # Validate inputs
            if not box_points or not text:
                return None
            
            # Ensure text is string
            text = str(text).strip()
            if not text:
                return None
            
            # Ensure confidence is float
            try:
                confidence = float(confidence)
            except (ValueError, TypeError):
                confidence = 1.0
            
            # Ensure box_points is list
            if isinstance(box_points, np.ndarray):
                box_points = box_points.tolist()
            
            # Extract bounding box from polygon points
            x_coords = [p[0] for p in box_points]
            y_coords = [p[1] for p in box_points]
            
            x1 = int(min(x_coords))
            y1 = int(min(y_coords))
            x2 = int(max(x_coords))
            y2 = int(max(y_coords))
            
            # Validate coordinates are within image bounds
            height, width = image.shape[:2]
            x1 = max(0, min(x1, width))
            y1 = max(0, min(y1, height))
            x2 = max(0, min(x2, width))
            y2 = max(0, min(y2, height))
            
            # Calculate rotation angle from box points
            angle = self._calculate_angle(box_points)
            
            # Estimate font size from box height
            box_height = y2 - y1
            font_size = max(8, int(box_height * 0.8))
            
            return TextRegion(
                bbox=(x1, y1, x2, y2),
                text=text,
                confidence=float(confidence),
                font_size=font_size,
                angle=angle
            )
            
        except (ValueError, TypeError, IndexError) as e:
            logger.warning(f"Failed to convert OCR adapter result: {e}")
            logger.debug(f"box_points={box_points}, text={text}, confidence={confidence}")
            return None
    
    def _convert_ocr_result(
        self, 
        item: Tuple, 
        image: np.ndarray
    ) -> Optional[TextRegion]:
        """Convert PaddleOCR result item to TextRegion.
        
        Args:
            item: OCR result item containing box coordinates and text info
            image: Original image for validation
            
        Returns:
            TextRegion object or None if conversion fails
        """
        try:
            # PaddleOCR 有多种返回格式：
            # 格式1 (v4): [box_points, (text, confidence)]
            # 格式2 (v5): [box_points, text, confidence]
            # 格式3: [box_points, text]
            # 格式4 (v5_server): 字典格式 {'dt_polys': [...], 'rec_text': '...', 'rec_score': 0.xx, ...}
            
            box_points = None
            text = None
            confidence = 1.0
            
            # 检查是否是字典格式（PP-OCRv5_server）
            if isinstance(item, dict):
                # 格式4: 字典格式
                box_points = item.get('dt_polys')
                text = item.get('rec_text')
                confidence = item.get('rec_score', 1.0)
                
                # 如果没有 dt_polys，尝试其他可能的键名
                if box_points is None:
                    box_points = item.get('bbox') or item.get('box') or item.get('polygon')
                
                if text is None:
                    text = item.get('text') or item.get('transcription')
                    
            elif len(item) == 2:
                # 格式1 或格式3
                box_points = item[0]
                text_info = item[1]
                
                if isinstance(text_info, (tuple, list)) and len(text_info) == 2:
                    # 格式1: [box_points, (text, confidence)]
                    text, confidence = text_info
                else:
                    # 格式3: [box_points, text]
                    text = str(text_info)
                    confidence = 1.0
                    
            elif len(item) == 3:
                # 格式2: [box_points, text, confidence]
                box_points, text, confidence = item
                
            elif len(item) > 3:
                # PP-OCRv5_server 可能返回多元素列表
                # 尝试解析：通常第一个是坐标，然后是文本和置信度
                logger.debug(f"Attempting to parse {len(item)}-element OCR result")
                
                # 查找坐标数据（通常是列表的列表或numpy数组）
                for i, element in enumerate(item):
                    if isinstance(element, (list, np.ndarray)):
                        try:
                            # 检查是否是坐标点（4个点，每个点2个坐标）
                            if len(element) == 4 and all(len(p) == 2 for p in element):
                                box_points = element
                                break
                        except (TypeError, AttributeError):
                            continue
                
                # 查找文本（通常是字符串）
                for element in item:
                    if isinstance(element, str) and element.strip():
                        text = element
                        break
                
                # 查找置信度（通常是0-1之间的浮点数）
                for element in item:
                    if isinstance(element, (int, float)):
                        try:
                            conf = float(element)
                            if 0 <= conf <= 1:
                                confidence = conf
                                break
                        except (ValueError, TypeError):
                            continue
                
                if box_points is None or text is None:
                    logger.warning(f"Could not parse {len(item)}-element OCR result")
                    logger.debug(f"Item types: {[type(x).__name__ for x in item]}")
                    return None
            else:
                logger.warning(f"Unexpected OCR result format with {len(item)} elements")
                return None
            
            # 验证必需字段
            if box_points is None or text is None:
                logger.warning("Missing box_points or text in OCR result")
                return None
            
            # 确保 text 是字符串
            text = str(text).strip()
            if not text:
                return None
            
            # 确保 confidence 是数字
            try:
                confidence = float(confidence)
            except (ValueError, TypeError):
                confidence = 1.0
            
            # 确保 box_points 是正确的格式
            if isinstance(box_points, np.ndarray):
                box_points = box_points.tolist()
            
            # Extract bounding box from polygon points
            x_coords = [p[0] for p in box_points]
            y_coords = [p[1] for p in box_points]
            
            x1 = int(min(x_coords))
            y1 = int(min(y_coords))
            x2 = int(max(x_coords))
            y2 = int(max(y_coords))
            
            # Validate coordinates are within image bounds
            height, width = image.shape[:2]
            x1 = max(0, min(x1, width))
            y1 = max(0, min(y1, height))
            x2 = max(0, min(x2, width))
            y2 = max(0, min(y2, height))
            
            # Calculate rotation angle from box points
            angle = self._calculate_angle(box_points)
            
            # Estimate font size from box height
            box_height = y2 - y1
            font_size = max(8, int(box_height * 0.8))
            
            return TextRegion(
                bbox=(x1, y1, x2, y2),
                text=text,
                confidence=float(confidence),
                font_size=font_size,
                angle=angle
            )
            
        except (ValueError, TypeError, IndexError) as e:
            logger.warning(f"Failed to convert OCR result: {e}")
            logger.debug(f"OCR item structure: {item}")
            return None
    
    def _calculate_angle(self, box_points: List[List[float]]) -> float:
        """Calculate rotation angle from box polygon points.
        
        Args:
            box_points: List of 4 corner points [[x1,y1], [x2,y2], ...]
            
        Returns:
            Rotation angle in degrees
        """
        try:
            # Calculate angle from top edge of the box
            p1, p2 = box_points[0], box_points[1]
            dx = p2[0] - p1[0]
            dy = p2[1] - p1[1]
            angle = math.degrees(math.atan2(dy, dx))
            return angle
        except (IndexError, TypeError):
            return 0.0
    
    def _detect_and_mask_qrcodes(
        self, 
        image: np.ndarray,
        mask_qr: bool = True
    ) -> Tuple[List[TextRegion], np.ndarray]:
        """Detect QR codes before OCR and optionally mask them to prevent misrecognition.
        
        This method implements Requirements 14 & 15:
        - Detects QR codes using OpenCV's QRCodeDetector
        - Optionally creates a mask to cover QR code areas (if mask_qr=True)
        - Returns QR code regions and optionally masked image
        
        Args:
            image: Preprocessed image as numpy array
            mask_qr: Whether to mask QR codes (default: True). If False, QR codes are
                    detected but not masked (hard protection mode)
            
        Returns:
            Tuple of (qr_regions, masked_image):
            - qr_regions: List of TextRegion objects for detected QR codes
            - masked_image: Image with QR codes masked out (if mask_qr=True), 
                           or original image (if mask_qr=False)
            
        _Requirements: 14, 15_
        """
        import cv2
        
        qr_regions = []
        masked_image = image.copy()
        
        try:
            # Get configuration
            mask_padding = self.config.get('qr_pre_detection.mask_padding', 10)
            mask_color = self.config.get('qr_pre_detection.mask_color', 255)
            max_image_size = self.config.get('qr_pre_detection.max_image_size', 2000)
            downsample_ratio = self.config.get('qr_pre_detection.downsample_ratio', 0.5)
            
            height, width = image.shape[:2]
            
            # Downsample if image is too large (Requirements 15.2)
            scale_factor = 1.0
            detection_image = image
            
            if width > max_image_size or height > max_image_size:
                scale_factor = downsample_ratio
                new_width = int(width * scale_factor)
                new_height = int(height * scale_factor)
                detection_image = cv2.resize(image, (new_width, new_height), interpolation=cv2.INTER_AREA)
                logger.debug(
                    f"Downsampling image for QR detection: "
                    f"{width}x{height} -> {new_width}x{new_height} (scale={scale_factor})"
                )
            
            # Detect QR codes using OpenCV (Requirements 14.2, 15.1)
            detector = cv2.QRCodeDetector()
            
            # Try method 1: Detect multiple QR codes on full image
            retval, decoded_info, points, straight_qrcode = detector.detectAndDecodeMulti(detection_image)
            
            detected_qr_codes = []
            
            if retval and points is not None and len(points) > 0:
                logger.info(f"Detected {len(points)} QR code(s) using detectAndDecodeMulti on full image")
                for i, qr_points in enumerate(points):
                    if qr_points is not None and len(qr_points) > 0:
                        qr_data = decoded_info[i] if decoded_info and i < len(decoded_info) else ""
                        detected_qr_codes.append((qr_points, qr_data))
            else:
                # Try method 2: Detect single QR code on full image
                logger.debug("detectAndDecodeMulti failed, trying detectAndDecode for single QR code")
                data, qr_points, _ = detector.detectAndDecode(detection_image)
                
                if qr_points is not None and len(qr_points) > 0:
                    logger.info(f"Detected 1 QR code using detectAndDecode on full image")
                    detected_qr_codes.append((qr_points, data))
                else:
                    # Try method 3: Scan common QR code regions (top-right, top-left, etc.)
                    logger.debug("Full image detection failed, trying region-based detection")
                    
                    # Define regions to scan (as percentage of image size)
                    scan_regions = [
                        ("top-right", 0.65, 0.0, 1.0, 0.35),  # Right 35%, top 35%
                        ("top-left", 0.0, 0.0, 0.35, 0.35),   # Left 35%, top 35%
                        ("bottom-right", 0.65, 0.65, 1.0, 1.0),  # Right 35%, bottom 35%
                        ("bottom-left", 0.0, 0.65, 0.35, 1.0),   # Left 35%, bottom 35%
                    ]
                    
                    det_height, det_width = detection_image.shape[:2]
                    
                    for region_name, x1_ratio, y1_ratio, x2_ratio, y2_ratio in scan_regions:
                        roi_x1 = int(det_width * x1_ratio)
                        roi_y1 = int(det_height * y1_ratio)
                        roi_x2 = int(det_width * x2_ratio)
                        roi_y2 = int(det_height * y2_ratio)
                        
                        roi = detection_image[roi_y1:roi_y2, roi_x1:roi_x2]
                        
                        if roi.size == 0:
                            continue
                        
                        # Try to detect QR code in this region
                        roi_data, roi_points, _ = detector.detectAndDecode(roi)
                        
                        if roi_points is not None and len(roi_points) > 0:
                            logger.info(f"Detected QR code in {region_name} region")
                            
                            # Adjust points to full image coordinates
                            adjusted_points = roi_points.copy()
                            adjusted_points[:, :, 0] += roi_x1
                            adjusted_points[:, :, 1] += roi_y1
                            
                            detected_qr_codes.append((adjusted_points[0], roi_data))
                            break  # Found one, stop scanning
            
            if detected_qr_codes:
                logger.info(f"Total detected {len(detected_qr_codes)} QR code(s) using OpenCV QRCodeDetector")
                
                for i, (qr_points, qr_data) in enumerate(detected_qr_codes):
                    if qr_points is None or len(qr_points) == 0:
                        continue
                    
                    # Scale coordinates back to original image size (Requirements 15.3)
                    if scale_factor != 1.0:
                        qr_points = qr_points / scale_factor
                    
                    # Extract bounding box from QR code points
                    # Use tight bounding box without padding for the region
                    x_coords = [p[0] for p in qr_points]
                    y_coords = [p[1] for p in qr_points]
                    
                    qr_x1 = int(max(0, min(x_coords)))
                    qr_y1 = int(max(0, min(y_coords)))
                    qr_x2 = int(min(width, max(x_coords)))
                    qr_y2 = int(min(height, max(y_coords)))
                    
                    # For masking, use a smaller padding to avoid covering nearby text
                    mask_x1 = int(max(0, qr_x1 - mask_padding))
                    mask_y1 = int(max(0, qr_y1 - mask_padding))
                    mask_x2 = int(min(width, qr_x2 + mask_padding))
                    mask_y2 = int(min(height, qr_y2 + mask_padding))
                    
                    logger.info(
                        f"QR code #{i+1}: bbox=({qr_x1}, {qr_y1}, {qr_x2}, {qr_y2}), "
                        f"mask=({mask_x1}, {mask_y1}, {mask_x2}, {mask_y2}), "
                        f"data='{qr_data[:30]}...'" if qr_data else f"QR code #{i+1}: bbox=({qr_x1}, {qr_y1}, {qr_x2}, {qr_y2})"
                    )
                    
                    # Create TextRegion for QR code (Requirements 14.6)
                    # Use tight bbox without padding
                    qr_region = TextRegion(
                        bbox=(qr_x1, qr_y1, qr_x2, qr_y2),
                        text="",  # Empty text - should not be rendered
                        confidence=1.0,
                        font_size=max(8, int((qr_y2 - qr_y1) * 0.8)),
                        angle=0.0
                    )
                    # Mark as QR code for icon detector
                    qr_region.is_qrcode = True
                    qr_region.should_render = False  # Don't render this region
                    qr_region.hard_protected = True  # Mark as hard protected
                    qr_regions.append(qr_region)
                    
                    # Mask the QR code area only if mask_qr is True (Requirements 14.3, 14.4)
                    if mask_qr:
                        # Fill with white or background color to prevent OCR from detecting it
                        # Use the padded mask coordinates
                        cv2.rectangle(
                            masked_image,
                            (mask_x1, mask_y1),
                            (mask_x2, mask_y2),
                            (mask_color, mask_color, mask_color),
                            -1  # Filled rectangle
                        )
                        
                        logger.debug(
                            f"Masked QR code area: ({mask_x1}, {mask_y1}, {mask_x2}, {mask_y2}), "
                            f"padding={mask_padding}px, color={mask_color}"
                        )
                    else:
                        logger.debug(
                            f"Hard protection: QR code area NOT masked: ({qr_x1}, {qr_y1}, {qr_x2}, {qr_y2})"
                        )
            else:
                logger.debug("No QR codes detected in pre-detection phase")
        
        except Exception as e:
            # If pre-detection fails, continue with normal OCR (Requirements 14.11)
            logger.warning(f"QR code pre-detection failed: {e}, continuing with normal OCR")
            import traceback
            logger.debug(f"Traceback: {traceback.format_exc()}")
            return [], image
        
        return qr_regions, masked_image
    
    def _mark_qr_nearby_text_as_protected(
        self,
        regions: List[TextRegion],
        qr_regions: List[TextRegion]
    ) -> List[TextRegion]:
        """Mark text regions near QR codes as hard protected, and filter out overlapping regions.
        
        In hard protection mode:
        1. Filter out text regions that significantly overlap with QR codes (likely OCR errors)
        2. Mark text regions near QR codes as hard protected
        
        Args:
            regions: All text regions (including QR codes)
            qr_regions: QR code regions
            
        Returns:
            Filtered and updated regions with hard_protected flag set for nearby text
        """
        if not qr_regions:
            return regions
        
        # Define thresholds
        overlap_threshold = 0.5  # If region overlaps >50% with QR code, it's likely OCR error
        proximity_threshold = 50  # Text within 50px of QR code is considered "nearby"
        
        filtered_regions = []
        protected_count = 0
        filtered_count = 0
        
        for region in regions:
            # Skip if already a QR code
            if hasattr(region, 'is_qrcode') and region.is_qrcode:
                filtered_regions.append(region)
                continue
            
            # Check if this text region overlaps with any QR code
            should_filter = False
            for qr_region in qr_regions:
                overlap_ratio = self._calculate_overlap_ratio(region, qr_region)
                if overlap_ratio > overlap_threshold:
                    # This region significantly overlaps with QR code - likely OCR error
                    filtered_count += 1
                    should_filter = True
                    logger.info(
                        f"🗑️ Filtering OCR error region (overlaps {overlap_ratio*100:.1f}% with QR code): "
                        f"'{region.text[:30]}...' at {region.bbox}"
                    )
                    break
            
            if should_filter:
                continue  # Don't add this region to filtered_regions
            
            # Check if this text region is near any QR code
            for qr_region in qr_regions:
                if self._is_region_near_qr(region, qr_region, proximity_threshold):
                    region.hard_protected = True
                    protected_count += 1
                    logger.info(
                        f"🛡️ Marking text region as hard protected (near QR code): "
                        f"'{region.text[:30]}...' at {region.bbox}"
                    )
                    break
            
            filtered_regions.append(region)
        
        if filtered_count > 0:
            logger.info(f"🗑️ Filtered {filtered_count} OCR error region(s) overlapping with QR codes")
        if protected_count > 0:
            logger.info(f"🛡️ Marked {protected_count} text region(s) near QR codes as hard protected")
        
        return filtered_regions
    
    def _is_region_near_qr(
        self,
        text_region: TextRegion,
        qr_region: TextRegion,
        threshold: int
    ) -> bool:
        """Check if a text region is near a QR code region.
        
        Args:
            text_region: Text region to check
            qr_region: QR code region
            threshold: Distance threshold in pixels
            
        Returns:
            True if text region is within threshold distance of QR code
        """
        tx1, ty1, tx2, ty2 = text_region.bbox
        qx1, qy1, qx2, qy2 = qr_region.bbox
        
        # Calculate center points
        text_center_x = (tx1 + tx2) / 2
        text_center_y = (ty1 + ty2) / 2
        qr_center_x = (qx1 + qx2) / 2
        qr_center_y = (qy1 + qy2) / 2
        
        # Calculate distance between centers
        distance = ((text_center_x - qr_center_x) ** 2 + (text_center_y - qr_center_y) ** 2) ** 0.5
        
        return distance <= threshold
    
    def _calculate_overlap_ratio(
        self,
        region1: TextRegion,
        region2: TextRegion
    ) -> float:
        """Calculate the overlap ratio between two regions.
        
        Args:
            region1: First region
            region2: Second region
            
        Returns:
            Overlap ratio (0.0 to 1.0) relative to region1's area
        """
        x1_1, y1_1, x2_1, y2_1 = region1.bbox
        x1_2, y1_2, x2_2, y2_2 = region2.bbox
        
        # Calculate intersection
        x1_i = max(x1_1, x1_2)
        y1_i = max(y1_1, y1_2)
        x2_i = min(x2_1, x2_2)
        y2_i = min(y2_1, y2_2)
        
        if x1_i >= x2_i or y1_i >= y2_i:
            # No overlap
            return 0.0
        
        # Calculate areas
        intersection_area = (x2_i - x1_i) * (y2_i - y1_i)
        region1_area = (x2_1 - x1_1) * (y2_1 - y1_1)
        
        if region1_area == 0:
            return 0.0
        
        return intersection_area / region1_area
        distance = ((text_center_x - qr_center_x) ** 2 + (text_center_y - qr_center_y) ** 2) ** 0.5
        
        return distance <= threshold
    
    def _merge_multi_param_results(
        self,
        standard_regions: List[TextRegion],
        multi_param_regions: List[TextRegion]
    ) -> List[TextRegion]:
        """Merge results from standard OCR and multi-parameter OCR.
        
        Strategy:
        - For regions that appear in both, use the one with higher confidence
        - Add unique regions from multi-param results
        
        Args:
            standard_regions: Regions from standard OCR
            multi_param_regions: Regions from multi-parameter OCR
            
        Returns:
            Merged list of TextRegion objects
        """
        if not multi_param_regions:
            return standard_regions
        
        if not standard_regions:
            return multi_param_regions
        
        merged = []
        used_multi_param = set()
        
        # For each standard region, check if there's a better multi-param match
        for std_region in standard_regions:
            best_match = None
            best_match_idx = -1
            max_overlap = 0
            
            for i, mp_region in enumerate(multi_param_regions):
                if i in used_multi_param:
                    continue
                
                # Calculate overlap
                overlap = self._calculate_region_overlap(std_region, mp_region)
                
                if overlap > 0.5 and overlap > max_overlap:
                    max_overlap = overlap
                    best_match = mp_region
                    best_match_idx = i
            
            if best_match and best_match.confidence > std_region.confidence:
                # Use multi-param result (higher confidence)
                merged.append(best_match)
                used_multi_param.add(best_match_idx)
                logger.debug(
                    f"Using multi-param result: '{best_match.text}' "
                    f"(conf={best_match.confidence:.2f} vs {std_region.confidence:.2f})"
                )
            else:
                # Use standard result
                merged.append(std_region)
        
        # Add unique multi-param regions
        for i, mp_region in enumerate(multi_param_regions):
            if i not in used_multi_param:
                merged.append(mp_region)
                logger.debug(f"Adding unique multi-param region: '{mp_region.text}'")
        
        logger.info(
            f"Merged results: {len(standard_regions)} standard + "
            f"{len(multi_param_regions)} multi-param -> {len(merged)} final"
        )
        
        return merged
    
    def _calculate_region_overlap(
        self,
        region1: TextRegion,
        region2: TextRegion
    ) -> float:
        """Calculate overlap ratio between two regions.
        
        Args:
            region1: First region
            region2: Second region
            
        Returns:
            Overlap ratio (0-1)
        """
        x1_1, y1_1, x2_1, y2_1 = region1.bbox
        x1_2, y1_2, x2_2, y2_2 = region2.bbox
        
        # Calculate intersection
        x1_i = max(x1_1, x1_2)
        y1_i = max(y1_1, y1_2)
        x2_i = min(x2_1, x2_2)
        y2_i = min(y2_1, y2_2)
        
        if x1_i >= x2_i or y1_i >= y2_i:
            return 0.0
        
        intersection_area = (x2_i - x1_i) * (y2_i - y1_i)
        
        area1 = (x2_1 - x1_1) * (y2_1 - y1_1)
        area2 = (x2_2 - x1_2) * (y2_2 - y1_2)
        
        min_area = min(area1, area2)
        
        if min_area == 0:
            return 0.0
        
        return intersection_area / min_area

    def _merge_important_notice(self, regions: List[TextRegion]) -> List[TextRegion]:
        """检测并合并竖向排列的"重要提示"字符。
        
        这个方法专门处理横版文档中竖向排列的"重要提示"文字，
        确保它们被合并成一个整体进行翻译。
        
        检测条件：
        1. 找到"重"、"要"、"提"、"示"中的至少3个字
        2. 它们竖向排列（在同一列）
        3. 垂直距离合理
        4. 字体大小相近
        
        Args:
            regions: 输入的文本区域列表
            
        Returns:
            合并后的文本区域列表
        """
        if not regions or len(regions) < 3:
            return regions
        
        logger.info(f"开始检测竖向排列的'重要提示'，输入 {len(regions)} 个区域")
        
        # 显示输入中包含目标字符的区域
        target_chars_set = set(['重', '要', '提', '示'])
        input_target_regions = []
        for r in regions:
            text = r.text.strip()
            if any(c in text for c in target_chars_set):
                input_target_regions.append((text, r.bbox))
        
        if input_target_regions:
            logger.info(f"输入中包含目标字符的区域: {len(input_target_regions)}个")
            for text, bbox in input_target_regions[:10]:  # 只显示前10个
                logger.info(f"  '{text[:30]}...' - bbox={bbox}")
        
        # 目标字符序列
        target_chars = ['重', '要', '提', '示']
        
        # 按从上到下、从左到右排序
        sorted_regions = sorted(regions, key=lambda r: (r.bbox[1], r.bbox[0]))
        
        merged_regions = []
        used_indices = set()
        found_important_notice = False
        
        # 遍历所有区域，寻找"重"字开始的序列
        for i, region in enumerate(sorted_regions):
            if i in used_indices:
                continue
            
            text = region.text.strip()
            
            # 检查是否是"重"字
            if text != '重':
                continue
            
            logger.debug(f"找到'重'字 (index={i}, bbox={region.bbox})")
            
            # 尝试找到后续的"要"、"提"、"示"
            candidate_group = [region]
            candidate_indices = [i]
            candidate_chars = ['重']
            
            # 从当前位置开始查找后续字符
            for target_idx, target_char in enumerate(target_chars[1:], start=1):
                found_next = False
                
                # 在后续区域中查找目标字符
                for j in range(i + 1, len(sorted_regions)):
                    if j in used_indices or j in candidate_indices:
                        continue
                    
                    next_region = sorted_regions[j]
                    next_text = next_region.text.strip()
                    
                    # 检查是否是目标字符
                    if next_text != target_char:
                        continue
                    
                    # 检查是否可以合并到当前组
                    last_in_group = candidate_group[-1]
                    can_merge = self._can_merge_vertical_important_notice(last_in_group, next_region)
                    
                    logger.debug(
                        f"  检查'{target_char}'字 (index={j}): "
                        f"can_merge={can_merge}, bbox={next_region.bbox}"
                    )
                    
                    if can_merge:
                        candidate_group.append(next_region)
                        candidate_indices.append(j)
                        candidate_chars.append(target_char)
                        found_next = True
                        logger.debug(f"  ✓ 找到'{target_char}'字，添加到候选组")
                        break
                
                # 如果没有找到下一个字符，继续查找其他字符（不放弃）
                if not found_next:
                    logger.debug(f"  - 未找到'{target_char}'字，继续查找其他字符")
            
            # 特殊处理：如果找到了3个字，尝试强制查找第4个字（放宽条件）
            if len(candidate_group) == 3 and '示' not in candidate_chars:
                logger.info(f"找到3个字 {candidate_chars}，尝试强制查找第4个字'示'（放宽条件）")
                logger.info(f"当前候选组最后一个字: '{candidate_group[-1].text}', bbox={candidate_group[-1].bbox}")
                
                # 使用更宽松的条件查找"示"字或以"示"开头的文本
                checked_count = 0
                for j in range(i + 1, len(sorted_regions)):
                    if j in used_indices or j in candidate_indices:
                        continue
                    
                    next_region = sorted_regions[j]
                    next_text = next_region.text.strip()
                    
                    # 检查是否是"示"字，或者以"示"开头的文本
                    is_shi_char = (next_text == '示')
                    starts_with_shi = (len(next_text) > 1 and next_text[0] == '示')
                    
                    if not (is_shi_char or starts_with_shi):
                        continue
                    
                    checked_count += 1
                    logger.info(f"  检查区域{j}: text='{next_text[:20]}...', is_shi={is_shi_char}, starts_with_shi={starts_with_shi}")
                    
                    # 使用更宽松的合并条件
                    last_in_group = candidate_group[-1]
                    
                    # 特殊处理：如果是以"示"开头的长文本，使用"示"字的估算位置来判断
                    if starts_with_shi and len(next_text) > 1:
                        # 创建一个临时的"示"字区域用于判断
                        x1, y1, x2, y2 = next_region.bbox
                        char_width = (x2 - x1) / len(next_text)
                        shi_x2 = int(x1 + char_width)
                        
                        temp_shi_region = TextRegion(
                            bbox=(x1, y1, shi_x2, y2),
                            text='示',
                            confidence=next_region.confidence,
                            font_size=next_region.font_size,
                            angle=next_region.angle
                        )
                        can_merge_relaxed = self._can_merge_vertical_important_notice_relaxed(
                            last_in_group, temp_shi_region
                        )
                    else:
                        can_merge_relaxed = self._can_merge_vertical_important_notice_relaxed(
                            last_in_group, next_region
                        )
                    
                    logger.info(
                        f"  强制检查'示'字 (index={j}): "
                        f"text='{next_text[:30]}...', is_shi={is_shi_char}, starts_with_shi={starts_with_shi}, "
                        f"can_merge_relaxed={can_merge_relaxed}, bbox={next_region.bbox}"
                    )
                    
                    if can_merge_relaxed:
                        if is_shi_char:
                            # 如果是单独的"示"字，直接添加
                            candidate_group.append(next_region)
                            candidate_indices.append(j)
                            candidate_chars.append('示')
                            logger.info(f"  ✓✓✓ 强制找到单独的'示'字，完成'重要提示'四字合并")
                        else:
                            # 如果是以"示"开头的文本（如"示年度报告"），需要拆分
                            # 创建一个只包含"示"字的新区域
                            x1, y1, x2, y2 = next_region.bbox
                            # 估算"示"字的宽度（假设是整个区域宽度的1/文本长度）
                            char_width = (x2 - x1) / len(next_text)
                            shi_x2 = int(x1 + char_width)
                            
                            shi_region = TextRegion(
                                bbox=(x1, y1, shi_x2, y2),
                                text='示',
                                confidence=next_region.confidence,
                                font_size=next_region.font_size,
                                angle=next_region.angle
                            )
                            
                            candidate_group.append(shi_region)
                            candidate_indices.append(j)  # 标记原区域已使用
                            candidate_chars.append('示')
                            
                            # 创建剩余文本的新区域（保留"年度报告..."部分）
                            remaining_text = next_text[1:]  # 去掉第一个"示"字
                            if remaining_text:
                                remaining_x1 = shi_x2
                                
                                # 调整剩余文本的Y坐标，使其与前面的文本对齐
                                # 查找前面最近的文本区域（应该是"3.各类商事主休..."）
                                # 使用前面文本的Y坐标和高度
                                prev_region = candidate_group[0] if len(candidate_group) > 0 else None
                                if prev_region and len(candidate_group) >= 3:
                                    # 如果有前面的"重要提示"组，查找它们右边的文本
                                    # 简化处理：使用原始区域的Y坐标，但调整字体大小为13（与前面文本一致）
                                    remaining_font_size = 13  # 与前面的文本保持一致
                                else:
                                    # 重新计算剩余文本的字体大小（基于剩余区域的高度）
                                    remaining_height = y2 - y1
                                    remaining_font_size = max(8, int(remaining_height * 0.8))
                                
                                remaining_region = TextRegion(
                                    bbox=(remaining_x1, y1, x2, y2),
                                    text=remaining_text,
                                    confidence=next_region.confidence,
                                    font_size=remaining_font_size,  # 使用调整后的字体大小
                                    angle=next_region.angle
                                )
                                # 将剩余文本区域添加到sorted_regions中，以便后续处理
                                sorted_regions.append(remaining_region)
                                logger.info(f"  ✓ 保留剩余文本: '{remaining_text[:30]}...' (字体大小: {remaining_font_size})")
                            
                            logger.info(f"  ✓✓✓ 从'{next_text[:30]}...'中提取'示'字，完成'重要提示'四字合并")
                        break
                
                if checked_count == 0:
                    logger.info(f"  未找到任何包含'示'字的区域")
                elif '示' not in candidate_chars:
                    logger.info(f"  检查了{checked_count}个包含'示'的区域，但都无法合并")
            
            # 检查是否找到至少3个字符（重要提示的3个或更多）
            if len(candidate_group) >= 3:
                # 合并这些字
                merged_region = self._merge_vertical_group(candidate_group)
                
                # 根据找到的字符生成文本
                merged_text = ''.join(candidate_chars)
                merged_region.text = merged_text
                
                merged_regions.append(merged_region)
                
                # 标记这些区域已使用
                for idx in candidate_indices:
                    used_indices.add(idx)
                
                found_important_notice = True
                logger.info(
                    f"✓✓✓ 成功检测并合并竖向'重要提示'相关字符: "
                    f"text='{merged_text}', bbox={merged_region.bbox}, "
                    f"合并了 {len(candidate_group)} 个单字: {candidate_chars}"
                )
        
        # 添加未使用的区域
        for i, region in enumerate(sorted_regions):
            if i not in used_indices:
                merged_regions.append(region)
        
        if found_important_notice:
            logger.info(f"'重要提示'检测完成：找到并合并了竖向排列的相关字符")
        else:
            logger.info("'重要提示'检测完成：未找到竖向排列的相关字符")
        
        return merged_regions
    
    def _split_field_labels(self, regions: List[TextRegion]) -> List[TextRegion]:
        """拆分字段标签和内容（如果它们被OCR错误地识别成一个区域）。
        
        支持三种格式：
        1. 直接连接：如"法定代表人柯婷" → "法定代表人" + "柯婷"
        2. 冒号分隔：如"经营范围：商事主体..." → "经营范围" + "商事主体..."
        3. 空格分隔：如"名称 深圳市弘远贸易有限公司" → "名称" + "深圳市弘远贸易有限公司"
        4. 模糊匹配：如"经营 营期限长期" → "营业期限" + "长期" (处理OCR识别错误)
        
        Args:
            regions: 输入的文本区域列表
            
        Returns:
            拆分后的文本区域列表
        """
        # 字段标签列表
        # 注意：这里只包含左侧的字段标签，不包括"统一社会信用代码"等独立标题
        field_labels = [
            "法定代表人",
            "注册资本",
            "成立日期",
            "营业期限",
            "经营范围",
            "登记机关",
            "名称",
            "类型",
            "住所",
        ]
        
        # 模糊匹配规则: (OCR错误文本模式, 正确的字段标签, 前缀模式)
        # 用于处理OCR识别错误,如"营期限"→"营业期限"
        fuzzy_patterns = [
            ("营期限", "营业期限", ["经营 ", "经营"]),  # "经营 营期限长期" → "营业期限" + "长期"
        ]
        
        split_regions = []
        split_count = 0
        
        for region in regions:
            text = region.text.strip()
            
            # 首先检查模糊匹配模式（处理OCR识别错误）
            found_fuzzy_match = False
            for error_pattern, correct_label, prefixes in fuzzy_patterns:
                if error_pattern in text:
                    # 检查是否有前缀需要去除
                    cleaned_text = text
                    removed_prefix = ""
                    
                    for prefix in prefixes:
                        if text.startswith(prefix):
                            cleaned_text = text[len(prefix):]
                            removed_prefix = prefix
                            break
                    
                    # 检查清理后的文本是否以错误模式开头
                    if cleaned_text.startswith(error_pattern):
                        content_text = cleaned_text[len(error_pattern):].strip()
                        
                        if content_text:  # 确保内容不为空
                            x1, y1, x2, y2 = region.bbox
                            width = x2 - x1
                            
                            # 估算标签宽度（包括被去除的前缀）
                            # 前缀 + 错误模式的字符数
                            prefix_and_pattern_len = len(removed_prefix) + len(error_pattern)
                            label_ratio = prefix_and_pattern_len / len(text)
                            label_width = int(width * label_ratio)
                            
                            # 创建标签区域（使用正确的标签文本）
                            label_region = TextRegion(
                                bbox=(x1, y1, x1 + label_width, y2),
                                text=correct_label,  # 使用正确的标签
                                confidence=region.confidence,
                                font_size=region.font_size,
                                angle=region.angle
                            )
                            # 标记为字段标签
                            label_region.is_field_label = True
                            
                            # 创建内容区域
                            content_region = TextRegion(
                                bbox=(x1 + label_width, y1, x2, y2),
                                text=content_text,
                                confidence=region.confidence,
                                font_size=region.font_size,
                                angle=region.angle
                            )
                            # 标记为字段内容
                            content_region.is_field_content = True
                            
                            split_regions.append(label_region)
                            split_regions.append(content_region)
                            split_count += 1
                            found_fuzzy_match = True
                            
                            logger.info(
                                f"拆分字段标签（模糊匹配）: '{text}' → '{correct_label}' + '{content_text}' "
                                f"(OCR错误: '{error_pattern}' → '{correct_label}', 去除前缀: '{removed_prefix}')"
                            )
                            break
            
            if found_fuzzy_match:
                continue
            
            # 优先检查冒号分隔的情况（如"经营范围：xxx"）
            found_label_with_colon = None
            separator = None
            
            for label in field_labels:
                # 检查中文冒号
                if text.startswith(label + "：") and len(text) > len(label) + 1:
                    found_label_with_colon = label
                    separator = "："
                    break
                # 检查英文冒号
                elif text.startswith(label + ":") and len(text) > len(label) + 1:
                    found_label_with_colon = label
                    separator = ":"
                    break
            
            if found_label_with_colon:
                # 拆分字段标签和内容（冒号分隔）
                label_text = found_label_with_colon
                content_text = text[len(label_text) + len(separator):].strip()
                
                if content_text:  # 确保内容不为空
                    x1, y1, x2, y2 = region.bbox
                    width = x2 - x1
                    
                    # 估算标签和内容的宽度比例（包括冒号）
                    label_with_colon = label_text + separator
                    label_ratio = len(label_with_colon) / len(text)
                    label_width = int(width * label_ratio)
                    
                    # 创建标签区域（不包括冒号）
                    label_region = TextRegion(
                        bbox=(x1, y1, x1 + label_width, y2),
                        text=label_text,
                        confidence=region.confidence,
                        font_size=region.font_size,
                        angle=region.angle
                    )
                    # 标记为字段标签
                    label_region.is_field_label = True
                    
                    # 创建内容区域
                    content_region = TextRegion(
                        bbox=(x1 + label_width, y1, x2, y2),
                        text=content_text,
                        confidence=region.confidence,
                        font_size=region.font_size,
                        angle=region.angle
                    )
                    # 标记为字段内容
                    content_region.is_field_content = True
                    
                    split_regions.append(label_region)
                    split_regions.append(content_region)
                    split_count += 1
                    
                    logger.info(
                        f"拆分字段标签（冒号分隔）: '{text[:50]}...' → '{label_text}' + '{content_text[:30]}...'"
                    )
                else:
                    # 内容为空，保留原区域
                    split_regions.append(region)
                continue
            
            # 检查空格分隔的情况（如"名称 深圳市弘远贸易有限公司"）
            found_label_with_space = None
            for label in field_labels:
                if text.startswith(label + " ") and len(text) > len(label) + 1:
                    found_label_with_space = label
                    break
            
            if found_label_with_space:
                # 拆分字段标签和内容（空格分隔）
                label_text = found_label_with_space
                content_text = text[len(label_text) + 1:].strip()  # +1 for space
                
                if content_text:  # 确保内容不为空
                    x1, y1, x2, y2 = region.bbox
                    width = x2 - x1
                    
                    # 估算标签和内容的宽度比例（包括空格）
                    label_with_space = label_text + " "
                    label_ratio = len(label_with_space) / len(text)
                    label_width = int(width * label_ratio)
                    
                    # 创建标签区域（不包括空格）
                    label_region = TextRegion(
                        bbox=(x1, y1, x1 + label_width, y2),
                        text=label_text,
                        confidence=region.confidence,
                        font_size=region.font_size,
                        angle=region.angle
                    )
                    # 标记为字段标签
                    label_region.is_field_label = True
                    
                    # 创建内容区域
                    content_region = TextRegion(
                        bbox=(x1 + label_width, y1, x2, y2),
                        text=content_text,
                        confidence=region.confidence,
                        font_size=region.font_size,
                        angle=region.angle
                    )
                    # 标记为字段内容
                    content_region.is_field_content = True
                    
                    split_regions.append(label_region)
                    split_regions.append(content_region)
                    split_count += 1
                    
                    logger.info(
                        f"拆分字段标签（空格分隔）: '{text}' → '{label_text}' + '{content_text}'"
                    )
                else:
                    # 内容为空，保留原区域
                    split_regions.append(region)
                continue
            
            # 检查直接连接的情况（如"法定代表人柯婷"）
            found_label = None
            for label in field_labels:
                if text.startswith(label) and len(text) > len(label):
                    # 确保不是空格或冒号分隔（已经在上面处理过了）
                    next_char = text[len(label)]
                    if next_char not in [' ', '：', ':']:
                        found_label = label
                        break
            
            if found_label:
                # 拆分字段标签和内容
                label_text = found_label
                content_text = text[len(found_label):].strip()
                
                if content_text:  # 确保内容不为空
                    x1, y1, x2, y2 = region.bbox
                    width = x2 - x1
                    
                    # 估算标签和内容的宽度比例
                    label_ratio = len(label_text) / len(text)
                    label_width = int(width * label_ratio)
                    
                    # 创建标签区域
                    label_region = TextRegion(
                        bbox=(x1, y1, x1 + label_width, y2),
                        text=label_text,
                        confidence=region.confidence,
                        font_size=region.font_size,
                        angle=region.angle
                    )
                    # 标记为字段标签
                    label_region.is_field_label = True
                    
                    # 创建内容区域
                    content_region = TextRegion(
                        bbox=(x1 + label_width, y1, x2, y2),
                        text=content_text,
                        confidence=region.confidence,
                        font_size=region.font_size,
                        angle=region.angle
                    )
                    # 标记为字段内容
                    content_region.is_field_content = True
                    
                    split_regions.append(label_region)
                    split_regions.append(content_region)
                    split_count += 1
                    
                    logger.info(
                        f"拆分字段标签（直接连接）: '{text}' → '{label_text}' + '{content_text}'"
                    )
                else:
                    # 内容为空，保留原区域
                    split_regions.append(region)
            else:
                # 不包含字段标签，保留原区域
                split_regions.append(region)
        
        if split_count > 0:
            logger.info(f"字段标签拆分完成：拆分了 {split_count} 个区域")
        
        return split_regions
    
    def _remove_duplicate_regions(self, regions: List[TextRegion]) -> List[TextRegion]:
        """移除重复或重叠的文本区域。
        
        当两个区域高度重叠且其中一个是另一个的子集时，保留较大的区域。
        例如：如果有"经营范围"和"经"两个区域，且"经"完全包含在"经营范围"中，
        则移除"经"这个重复区域。
        
        Args:
            regions: 输入的文本区域列表
            
        Returns:
            去重后的文本区域列表
        """
        if not regions or len(regions) < 2:
            return regions
        
        # 按面积从大到小排序
        sorted_regions = sorted(regions, key=lambda r: r.width * r.height, reverse=True)
        
        filtered_regions = []
        removed_count = 0
        
        for i, region1 in enumerate(sorted_regions):
            is_duplicate = False
            x1_1, y1_1, x2_1, y2_1 = region1.bbox
            text1 = region1.text.strip()
            
            # 检查是否与已保留的区域重叠或文本包含
            for region2 in filtered_regions:
                x1_2, y1_2, x2_2, y2_2 = region2.bbox
                text2 = region2.text.strip()
                
                # 检查1：文本包含关系（如果text1是text2的子串，且位置接近）
                if text1 in text2 and text1 != text2:
                    # 检查Y坐标是否接近（在同一行）
                    y_center1 = (y1_1 + y2_1) / 2
                    y_center2 = (y1_2 + y2_2) / 2
                    y_diff = abs(y_center1 - y_center2)
                    
                    # 如果Y坐标差异小于较大区域高度的50%，认为在同一行
                    max_height = max(y2_1 - y1_1, y2_2 - y1_2)
                    if y_diff < max_height * 0.5:
                        is_duplicate = True
                        removed_count += 1
                        logger.debug(
                            f"移除重复区域（文本包含）: '{text1}' (包含在 '{text2}' 中, "
                            f"Y差异={y_diff:.1f}px)"
                        )
                        break
                
                # 检查2：空间重叠
                # 计算重叠区域
                overlap_x1 = max(x1_1, x1_2)
                overlap_y1 = max(y1_1, y1_2)
                overlap_x2 = min(x2_1, x2_2)
                overlap_y2 = min(y2_1, y2_2)
                
                # 检查是否有重叠
                if overlap_x1 < overlap_x2 and overlap_y1 < overlap_y2:
                    overlap_area = (overlap_x2 - overlap_x1) * (overlap_y2 - overlap_y1)
                    region1_area = region1.width * region1.height
                    
                    # 如果region1的80%以上被region2覆盖，认为是重复
                    overlap_ratio = overlap_area / region1_area if region1_area > 0 else 0
                    
                    if overlap_ratio > 0.8:
                        # 如果text1是text2的子串，或者两者相同，认为是重复
                        if text1 in text2 or text1 == text2:
                            is_duplicate = True
                            removed_count += 1
                            logger.debug(
                                f"移除重复区域（空间重叠）: '{text1}' (被 '{text2}' 覆盖, "
                                f"重叠率={overlap_ratio*100:.1f}%)"
                            )
                            break
            
            if not is_duplicate:
                filtered_regions.append(region1)
        
        if removed_count > 0:
            logger.info(f"去重完成：移除了 {removed_count} 个重复区域")
        
        # 恢复原始顺序（按Y坐标，然后X坐标）
        filtered_regions.sort(key=lambda r: (r.bbox[1], r.bbox[0]))
        
        return filtered_regions
    
    def _can_merge_vertical_important_notice(self, region1: TextRegion, region2: TextRegion) -> bool:
        """判断两个字符是否可以作为"重要提示"的一部分进行垂直合并。
        
        Args:
            region1: 第一个文本区域
            region2: 第二个文本区域
            
        Returns:
            True如果可以合并
        """
        x1_1, y1_1, x2_1, y2_1 = region1.bbox
        x1_2, y1_2, x2_2, y2_2 = region2.bbox
        
        # 计算中心点
        center_x1 = (x1_1 + x2_1) / 2
        center_x2 = (x1_2 + x2_2) / 2
        
        # 计算宽度
        width1 = x2_1 - x1_1
        width2 = x2_2 - x1_2
        avg_width = (width1 + width2) / 2
        
        # 1. 检查是否在同一列（水平中心点接近）
        # 对于"重要提示"，使用更宽松的阈值（60%）
        horizontal_center_diff = abs(center_x1 - center_x2)
        if horizontal_center_diff >= avg_width * 0.6:
            logger.debug(f"    水平偏移过大: {horizontal_center_diff:.1f}px >= {avg_width * 0.6:.1f}px")
            return False
        
        # 2. 检查垂直距离
        if y2_1 < y1_2:
            vertical_distance = y1_2 - y2_1
        elif y2_2 < y1_1:
            vertical_distance = y1_1 - y2_2
        else:
            vertical_distance = 0  # 垂直重叠
        
        # 对于"重要提示"，使用更宽松的垂直距离阈值（150像素）
        max_vertical_distance = 150
        if vertical_distance > max_vertical_distance:
            logger.debug(f"    垂直距离过大: {vertical_distance:.1f}px > {max_vertical_distance}px")
            return False
        
        # 3. 检查字体大小是否相近（使用更宽松的阈值40%）
        font_size1 = region1.font_size
        font_size2 = region2.font_size
        size_diff_ratio = abs(font_size1 - font_size2) / max(font_size1, font_size2)
        
        if size_diff_ratio > 0.4:
            logger.debug(f"    字体差异过大: {size_diff_ratio:.1%} > 40%")
            return False
        
        logger.debug(f"    ✓ 可以合并: 水平偏移={horizontal_center_diff:.1f}px, 垂直距离={vertical_distance:.1f}px, 字体差异={size_diff_ratio:.1%}")
        return True
    
    def _can_merge_vertical_important_notice_relaxed(self, region1: TextRegion, region2: TextRegion) -> bool:
        """判断两个字符是否可以作为"重要提示"的一部分进行垂直合并（放宽条件）。
        
        这个方法用于强制查找第4个字时，使用更宽松的条件。
        
        Args:
            region1: 第一个文本区域
            region2: 第二个文本区域
            
        Returns:
            True如果可以合并
        """
        x1_1, y1_1, x2_1, y2_1 = region1.bbox
        x1_2, y1_2, x2_2, y2_2 = region2.bbox
        
        # 计算中心点
        center_x1 = (x1_1 + x2_1) / 2
        center_x2 = (x1_2 + x2_2) / 2
        
        # 计算宽度
        width1 = x2_1 - x1_1
        width2 = x2_2 - x1_2
        avg_width = (width1 + width2) / 2
        
        # 1. 检查是否在同一列（水平中心点接近）
        # 使用绝对像素值（50px）或相对值（80%），取较大者
        horizontal_center_diff = abs(center_x1 - center_x2)
        max_horizontal_diff = max(50, avg_width * 0.8)
        
        if horizontal_center_diff >= max_horizontal_diff:
            logger.info(f"    [放宽] 水平偏移过大: {horizontal_center_diff:.1f}px >= {max_horizontal_diff:.1f}px")
            return False
        
        # 2. 检查垂直距离
        if y2_1 < y1_2:
            vertical_distance = y1_2 - y2_1
        elif y2_2 < y1_1:
            vertical_distance = y1_1 - y2_2
        else:
            vertical_distance = 0  # 垂直重叠
        
        # 放宽到300像素
        max_vertical_distance = 300
        if vertical_distance > max_vertical_distance:
            logger.info(f"    [放宽] 垂直距离过大: {vertical_distance:.1f}px > {max_vertical_distance}px")
            return False
        
        # 3. 检查字体大小是否相近（放宽到60%）
        font_size1 = region1.font_size
        font_size2 = region2.font_size
        size_diff_ratio = abs(font_size1 - font_size2) / max(font_size1, font_size2)
        
        if size_diff_ratio > 0.6:
            logger.info(f"    [放宽] 字体差异过大: {size_diff_ratio:.1%} > 60%")
            return False
        
        logger.info(f"    ✓ [放宽] 可以合并: 水平偏移={horizontal_center_diff:.1f}px, 垂直距离={vertical_distance:.1f}px, 字体差异={size_diff_ratio:.1%}")
        return True

    def merge_regions(self, regions: List[TextRegion]) -> List[TextRegion]:
        """Merge adjacent text regions based on distance threshold.
        
        Supports two merge strategies:
        1. Smart merge: Multi-feature intelligent grouping (for forms)
        2. Standard merge: Distance-based merging (default)
        
        Note: "重要提示" detection and merging is already done in detect_text()
        with highest priority, so it won't be affected by other merge logic.
        
        Args:
            regions: List of TextRegion objects to merge
            
        Returns:
            List of merged TextRegion objects
        """
        if not regions:
            return []
        
        if len(regions) == 1:
            return regions.copy()
        
        # 优先合并字段碎片（如"名"+"称"→"名称"）
        # 这样可以确保字段标签被正确识别
        regions = self._merge_field_fragments(regions)
        
        # Check if smart merge is enabled
        smart_merge_enabled = self.config.get('ocr.smart_merge.enabled', False)
        
        if smart_merge_enabled:
            logger.info("Using smart merge strategy (multi-feature grouping)")
            merged = self._smart_merge_regions(regions)
        else:
            logger.info("Using standard merge strategy (distance-based)")
            merged = self._standard_merge_regions(regions)
        
        # 调整"重要提示"与其他区域的重叠
        merged = self._adjust_important_notice_overlap(merged)
        
        return merged
    
    def _adjust_important_notice_overlap(self, regions: List[TextRegion]) -> List[TextRegion]:
        """调整"重要提示"与其他区域的重叠。
        
        如果"重要提示"的边界框与右边的文本框重叠，
        将右边文本框向右移动直到不重叠。
        
        Args:
            regions: 合并后的文本区域列表
            
        Returns:
            调整后的文本区域列表
        """
        # 查找"重要提示"区域
        important_notice = None
        important_notice_idx = -1
        
        for i, region in enumerate(regions):
            if region.text.strip() == '重要提示':
                important_notice = region
                important_notice_idx = i
                break
        
        if not important_notice:
            return regions
        
        x1_notice, y1_notice, x2_notice, y2_notice = important_notice.bbox
        logger.info(f"检查'重要提示'区域的重叠: bbox={important_notice.bbox}")
        
        # 调整后的区域列表
        adjusted_regions = []
        adjusted_count = 0
        
        for i, region in enumerate(regions):
            if i == important_notice_idx:
                # "重要提示"区域不需要调整
                adjusted_regions.append(region)
                continue
            
            x1, y1, x2, y2 = region.bbox
            
            # 检查Y轴是否有重叠
            y_overlap = not (y2 < y1_notice or y1 > y2_notice)
            
            if not y_overlap:
                # Y轴没有重叠，不需要调整
                adjusted_regions.append(region)
                continue
            
            # 检查X轴是否有重叠
            x_overlap = not (x2 < x1_notice or x1 > x2_notice)
            
            if not x_overlap:
                # X轴没有重叠，不需要调整
                adjusted_regions.append(region)
                continue
            
            # 计算重叠宽度
            overlap_x1 = max(x1, x1_notice)
            overlap_x2 = min(x2, x2_notice)
            overlap_width = overlap_x2 - overlap_x1
            
            # 如果重叠，将右边文本框向右移动
            if overlap_width > 0:
                # 计算需要移动的距离（移动到"重要提示"右边，加上一点间距）
                shift_distance = x2_notice - x1 + 5  # 加5px间距
                
                new_x1 = x1 + shift_distance
                new_x2 = x2 + shift_distance
                
                # 创建调整后的区域
                adjusted_region = TextRegion(
                    bbox=(new_x1, y1, new_x2, y2),
                    text=region.text,
                    confidence=region.confidence,
                    font_size=region.font_size,
                    angle=region.angle
                )
                
                adjusted_regions.append(adjusted_region)
                adjusted_count += 1
                
                logger.info(
                    f"调整区域 {i} 的X坐标: "
                    f"'{region.text[:30]}...' "
                    f"重叠{overlap_width}px, "
                    f"X: {x1}-{x2} → {new_x1}-{new_x2} (右移{shift_distance}px)"
                )
            else:
                adjusted_regions.append(region)
        
        if adjusted_count > 0:
            logger.info(f"'重要提示'重叠调整完成：调整了 {adjusted_count} 个区域")
        else:
            logger.info("'重要提示'重叠调整完成：没有需要调整的区域")
        
        return adjusted_regions
    
    def _standard_merge_regions(self, regions: List[TextRegion]) -> List[TextRegion]:
        """Standard distance-based merging (original logic).
        
        改进策略：
        1. 第一轮：优先合并垂直排列的单字（如"重要提示"）
        2. 第二轮：合并水平排列的文字
        3. 重复直到没有更多合并
        4. 二维码保护：避免合并跨越二维码边界的区域
        
        Args:
            regions: List of TextRegion objects to merge
            
        Returns:
            List of merged TextRegion objects
        """
        # 估算二维码位置（如果存在）
        # 二维码通常在右上角
        qr_bbox = None
        if hasattr(self, '_image_width') and hasattr(self, '_image_height'):
            width = self._image_width
            height = self._image_height
            qr_size = int(min(width, height) * 0.15)
            qr_x1 = int(width * 0.68)
            qr_y1 = int(height * 0.12)
            qr_x2 = qr_x1 + qr_size
            qr_y2 = qr_y1 + qr_size
            qr_bbox = (qr_x1, qr_y1, qr_x2, qr_y2)
            logger.info(f"二维码保护：估算二维码位置 {qr_bbox}")
        
        # 第一轮：优先合并垂直排列的单字
        current_regions = self._merge_vertical_single_chars(list(regions))
        
        # 第二轮：标准合并（水平和垂直）
        changed = True
        
        while changed:
            changed = False
            merged = []
            used = set()
            
            for i, region1 in enumerate(current_regions):
                if i in used:
                    continue
                
                current = region1
                
                for j, region2 in enumerate(current_regions):
                    if j <= i or j in used:
                        continue
                    
                    distance = self._calculate_region_distance(current, region2)
                    
                    # 判断是水平还是垂直方向的合并
                    is_horizontal = self._is_horizontal_merge(current, region2)
                    threshold = self.merge_threshold_horizontal if is_horizontal else self.merge_threshold_vertical
                    
                    # 详细日志：显示合并判断过程
                    logger.debug(
                        f"Region {i} vs {j}: "
                        f"text1='{current.text[:20]}', text2='{region2.text[:20]}', "
                        f"font1={current.font_size}, font2={region2.font_size}, "
                        f"distance={distance:.2f}, threshold={threshold}, "
                        f"is_horizontal={is_horizontal}"
                    )
                    
                    if distance <= threshold:
                        # 二维码保护：检查合并后是否会跨越二维码边界
                        if qr_bbox and self._would_cross_qrcode(current, region2, qr_bbox):
                            logger.info(
                                f"⚠️ 二维码保护：跳过合并 '{current.text[:20]}' + '{region2.text[:20]}' "
                                f"(会跨越二维码边界)"
                            )
                            continue
                        
                        current = self._merge_two_regions(current, region2)
                        used.add(j)
                        changed = True
                        logger.debug(f"Merged regions ({'horizontal' if is_horizontal else 'vertical'}): distance={distance:.2f}, threshold={threshold}")
                
                merged.append(current)
            
            current_regions = merged
        
        logger.info(f"Merged {len(regions)} regions into {len(current_regions)} regions")
        return current_regions
    
    def _merge_vertical_single_chars(self, regions: List[TextRegion]) -> List[TextRegion]:
        """优先合并垂直排列的单字（如"重要提示"）。
        
        这个方法专门处理垂直排列的单字，确保它们优先合并成一个整体，
        而不是被合并到旁边的大段文字中。
        
        合并条件：
        1. 两个区域都是单字（1-2个字符）
        2. 垂直排列（在同一列）
        3. 垂直距离在合理范围内
        4. 字体大小相近
        
        Args:
            regions: 输入的文本区域列表
            
        Returns:
            合并后的文本区域列表
        """
        # 检查配置是否启用垂直单字合并
        # 优先检查竖版模式的覆盖配置
        if hasattr(self, '_portrait_vertical_merge_enabled'):
            vertical_merge_enabled = self._portrait_vertical_merge_enabled
            logger.info(f"使用竖版模式的垂直单字合并配置: {vertical_merge_enabled}")
        else:
            vertical_merge_enabled = self.config.get(
                'ocr.smart_merge.vertical_single_char_merge.enabled', 
                True  # 默认启用
            )
        
        if not vertical_merge_enabled:
            logger.info("垂直单字合并已禁用，跳过")
            return regions
        
        if not regions or len(regions) < 2:
            return regions
        
        logger.info(f"开始垂直单字合并，输入 {len(regions)} 个区域")
        
        # 按从上到下、从左到右排序
        sorted_regions = sorted(regions, key=lambda r: (r.bbox[1], r.bbox[0]))
        
        merged_regions = []
        used_indices = set()
        merge_count = 0
        
        for i, region1 in enumerate(sorted_regions):
            if i in used_indices:
                continue
            
            text1 = region1.text.strip()
            
            # 只处理单字或双字
            if len(text1) > 2:
                merged_regions.append(region1)
                continue
            
            logger.debug(f"检查单字 '{text1}' (index={i}, bbox={region1.bbox})")
            
            # 尝试找到垂直排列的其他单字
            vertical_group = [region1]
            used_indices.add(i)
            
            for j in range(i + 1, len(sorted_regions)):
                if j in used_indices:
                    continue
                
                region2 = sorted_regions[j]
                text2 = region2.text.strip()
                
                # 只处理单字或双字
                if len(text2) > 2:
                    continue
                
                # 检查是否可以合并到当前垂直组
                last_in_group = vertical_group[-1]
                
                can_merge = self._can_merge_vertical_chars(last_in_group, region2)
                
                logger.debug(
                    f"  检查是否可以合并 '{last_in_group.text}' + '{text2}': "
                    f"can_merge={can_merge}, "
                    f"bbox1={last_in_group.bbox}, bbox2={region2.bbox}"
                )
                
                if can_merge:
                    vertical_group.append(region2)
                    used_indices.add(j)
                    logger.debug(f"  OK 添加到垂直组: '{text2}' (组大小: {len(vertical_group)})")
                else:
                    # 如果不能合并，检查是否因为距离太远
                    # 如果是，停止查找（因为已经排序，后面的更远）
                    x1_1, y1_1, x2_1, y2_1 = last_in_group.bbox
                    x1_2, y1_2, x2_2, y2_2 = region2.bbox
                    
                    # 如果垂直距离太大，停止查找
                    vertical_distance = y1_2 - y2_1
                    if vertical_distance > self.merge_threshold_horizontal:
                        logger.debug(f"  X 垂直距离太大 ({vertical_distance}px > {self.merge_threshold_horizontal}px)，停止查找")
                        break
            
            # 如果找到了垂直组（至少2个字符），合并它们
            if len(vertical_group) >= 2:
                merged_region = self._merge_vertical_group(vertical_group)
                merged_regions.append(merged_region)
                merge_count += 1
                logger.info(
                    f"✓ 合并垂直单字: "
                    f"{' + '.join([r.text for r in vertical_group])} → '{merged_region.text}'"
                )
            else:
                merged_regions.append(region1)
        
        if merge_count > 0:
            logger.info(f"垂直单字合并完成：合并了 {merge_count} 个垂直组")
        else:
            logger.info("垂直单字合并完成：没有找到可合并的垂直组")
        
        return merged_regions
    
    def _can_merge_vertical_chars(self, region1: TextRegion, region2: TextRegion) -> bool:
        """判断两个单字是否可以垂直合并。
        
        Args:
            region1: 第一个文本区域
            region2: 第二个文本区域
            
        Returns:
            True如果可以合并
        """
        # 0. 检查是否是标签字段（不应该合并）
        # 标签字段列表（从配置中获取）
        label_keywords = self.config.get('ocr.smart_merge.label_keywords', [])
        text1 = region1.text.strip()
        text2 = region2.text.strip()
        
        # 如果两个文本都是标签关键词，不合并
        if text1 in label_keywords and text2 in label_keywords:
            logger.debug(f"  X 两个文本都是标签关键词，不合并: '{text1}' + '{text2}'")
            return False
        
        x1_1, y1_1, x2_1, y2_1 = region1.bbox
        x1_2, y1_2, x2_2, y2_2 = region2.bbox
        
        # 计算中心点
        center_x1 = (x1_1 + x2_1) / 2
        center_x2 = (x1_2 + x2_2) / 2
        
        # 计算宽度
        width1 = x2_1 - x1_1
        width2 = x2_2 - x1_2
        avg_width = (width1 + width2) / 2
        
        # 1. 检查是否在同一列（水平中心点接近）
        horizontal_center_diff = abs(center_x1 - center_x2)
        if horizontal_center_diff >= avg_width * 0.5:
            return False
        
        # 2. 检查垂直距离
        if y2_1 < y1_2:
            vertical_distance = y1_2 - y2_1
        elif y2_2 < y1_1:
            vertical_distance = y1_1 - y2_2
        else:
            vertical_distance = 0  # 垂直重叠
        
        # 使用较宽松的垂直距离阈值（因为垂直文本字间距较大）
        if vertical_distance > self.merge_threshold_horizontal:
            return False
        
        # 3. 检查字体大小是否相近
        font_size1 = region1.font_size
        font_size2 = region2.font_size
        size_diff_ratio = abs(font_size1 - font_size2) / max(font_size1, font_size2)
        
        if size_diff_ratio > self.font_size_diff_threshold:
            return False
        
        return True
    
    def _merge_vertical_group(self, regions: List[TextRegion]) -> TextRegion:
        """合并垂直排列的单字组。
        
        Args:
            regions: 垂直排列的文本区域列表（从上到下排序）
            
        Returns:
            合并后的文本区域
        """
        if len(regions) == 1:
            return regions[0]
        
        # 计算合并后的边界框
        x1 = min(r.bbox[0] for r in regions)
        y1 = min(r.bbox[1] for r in regions)
        x2 = max(r.bbox[2] for r in regions)
        y2 = max(r.bbox[3] for r in regions)
        
        merged_bbox = (x1, y1, x2, y2)
        
        # 合并文本（从上到下）
        merged_text = ''.join(r.text.strip() for r in regions)
        
        # 使用平均置信度
        merged_confidence = sum(r.confidence for r in regions) / len(regions)
        
        # 使用平均字体大小
        merged_font_size = int(sum(r.font_size for r in regions) / len(regions))
        
        # 使用平均角度
        merged_angle = sum(r.angle for r in regions) / len(regions)
        
        return TextRegion(
            bbox=merged_bbox,
            text=merged_text,
            confidence=merged_confidence,
            font_size=merged_font_size,
            angle=merged_angle,
            is_vertical_merged=True  # 标记为垂直合并的文本
        )
    
    def _smart_merge_regions(self, regions: List[TextRegion]) -> List[TextRegion]:
        """Smart merge using multi-feature intelligent grouping.
        
        This method groups regions on the same line and uses multiple features
        (font size, text length, semantics, spacing) to detect label-content boundaries.
        
        改进策略：
        1. 第一步：优先合并垂直排列的单字（如"重要提示"）
        2. 第二步：按行分组
        3. 第三步：智能合并每一行
        4. 第四步：合并段落
        
        Args:
            regions: List of TextRegion objects to merge
            
        Returns:
            List of merged TextRegion objects
        """
        # Step 0: 优先合并垂直排列的单字
        regions = self._merge_vertical_single_chars(regions)
        logger.debug(f"After vertical single char merge: {len(regions)} regions")
        
        # Step 1: Group regions by line (Y coordinate)
        lines = self._group_regions_by_line(regions)
        logger.debug(f"Grouped {len(regions)} regions into {len(lines)} lines")
        
        # Step 2: Process each line with smart grouping
        merged_regions = []
        for line_idx, line_regions in enumerate(lines):
            logger.debug(f"Processing line {line_idx + 1} with {len(line_regions)} regions")
            line_merged = self._smart_merge_line(line_regions)
            merged_regions.extend(line_merged)
            logger.debug(f"Line {line_idx + 1} merged into {len(line_merged)} groups")
        
        # Step 3: Merge paragraph regions (vertically stacked, aligned regions)
        merged_regions = self._merge_paragraph_regions(merged_regions)
        
        logger.info(f"Smart merge: {len(regions)} regions → {len(merged_regions)} regions")
        return merged_regions
    
    def _group_regions_by_line(self, regions: List[TextRegion]) -> List[List[TextRegion]]:
        """Group regions into lines based on Y coordinate.
        
        改进策略：
        - 对于左侧区域（X < 图片宽度的30%），使用更严格的行分组阈值
        - 这样可以避免将垂直排列的字段标签错误地分到同一行
        
        Args:
            regions: List of TextRegion objects
            
        Returns:
            List of lines, where each line is a list of regions
        """
        if not regions:
            return []
        
        # 获取图片宽度（从第一个区域估算）
        # 假设图片宽度至少是最右边区域的X坐标的1.2倍
        max_x = max(r.bbox[2] for r in regions)
        estimated_image_width = max_x * 1.2
        left_region_threshold = estimated_image_width * 0.3  # 左侧区域阈值（30%）
        
        # Sort by Y coordinate
        sorted_regions = sorted(regions, key=lambda r: (r.bbox[1] + r.bbox[3]) / 2)
        
        lines = []
        current_line = [sorted_regions[0]]
        
        for region in sorted_regions[1:]:
            prev_region = current_line[-1]
            
            # 特殊处理：如果当前区域或前一个区域是垂直合并的文本，强制分行
            is_prev_vertical_merged = hasattr(prev_region, 'is_vertical_merged') and prev_region.is_vertical_merged
            is_curr_vertical_merged = hasattr(region, 'is_vertical_merged') and region.is_vertical_merged
            
            if is_prev_vertical_merged or is_curr_vertical_merged:
                # 垂直合并的文本单独成行
                lines.append(current_line)
                current_line = [region]
                continue
            
            # Calculate vertical distance between centers
            prev_center_y = (prev_region.bbox[1] + prev_region.bbox[3]) / 2
            curr_center_y = (region.bbox[1] + region.bbox[3]) / 2
            vertical_diff = abs(curr_center_y - prev_center_y)
            
            # Calculate average height
            avg_height = ((prev_region.bbox[3] - prev_region.bbox[1]) + 
                         (region.bbox[3] - region.bbox[1])) / 2
            
            # 检查是否在左侧区域（可能是垂直排列的字段标签）
            prev_x = prev_region.bbox[0]
            curr_x = region.bbox[0]
            is_prev_in_left = prev_x < left_region_threshold
            is_curr_in_left = curr_x < left_region_threshold
            
            # 对于左侧区域，使用更严格的阈值（30%而不是50%）
            # 这样可以避免将垂直排列的字段标签错误地分到同一行
            if is_prev_in_left and is_curr_in_left:
                threshold_ratio = 0.3  # 左侧区域使用30%阈值
                logger.debug(
                    f"Left region detected: prev='{prev_region.text[:10]}', curr='{region.text[:10]}', "
                    f"using strict threshold (30%)"
                )
            else:
                threshold_ratio = 0.5  # 其他区域使用50%阈值
            
            # If vertical difference < threshold of average height, same line
            if vertical_diff < avg_height * threshold_ratio:
                current_line.append(region)
            else:
                lines.append(current_line)
                current_line = [region]
        
        lines.append(current_line)
        return lines
    
    def _smart_merge_line(self, line_regions: List[TextRegion]) -> List[TextRegion]:
        """Smart merge regions on the same line using multi-feature detection.
        
        Args:
            line_regions: List of regions on the same line
            
        Returns:
            List of merged regions
        """
        if len(line_regions) <= 1:
            return line_regions
        
        # Sort by X coordinate
        sorted_regions = sorted(line_regions, key=lambda r: r.bbox[0])
        
        # Detect group boundaries
        groups = []
        current_group = [sorted_regions[0]]
        
        for i in range(1, len(sorted_regions)):
            prev = sorted_regions[i - 1]
            curr = sorted_regions[i]
            
            # Debug log for specific regions
            if ('统一社会信用代码' in prev.text or '91440300' in prev.text) and \
               ('统一社会信用代码' in curr.text or '91440300' in curr.text):
                logger.info(f"  Checking boundary: prev='{prev.text[:20]}', curr='{curr.text[:20]}'")
            
            # Check if there's a boundary between prev and curr
            is_boundary, reason = self._is_group_boundary(prev, curr)
            
            if is_boundary:
                logger.debug(
                    f"Group boundary detected: '{prev.text}' | '{curr.text}' - {reason}"
                )
                groups.append(current_group)
                current_group = [curr]
            else:
                current_group.append(curr)
        
        groups.append(current_group)
        
        # Merge each group
        merged = []
        for group in groups:
            if len(group) == 1:
                merged.append(group[0])
            else:
                merged_region = self._merge_group(group)
                merged.append(merged_region)
        
        return merged
    
    def _merge_paragraph_regions(self, regions: List[TextRegion]) -> List[TextRegion]:
        """Merge vertically stacked regions that belong to the same paragraph.
        
        This method:
        1. Groups regions into columns (based on X coordinate)
        2. Within each column, merges regions based on vertical spacing
        
        Args:
            regions: List of TextRegion objects
            
        Returns:
            List of merged TextRegion objects
        """
        if not regions:
            return []
        
        # 检查是否启用段落合并
        # 优先检查竖版模式的覆盖配置
        if hasattr(self, '_portrait_paragraph_merge_enabled'):
            paragraph_merge_enabled = self._portrait_paragraph_merge_enabled
            logger.info(f"使用竖版模式的段落合并配置: {paragraph_merge_enabled}")
        else:
            paragraph_merge_enabled = self.config.get('ocr.paragraph_merge.enabled', True)
        
        if not paragraph_merge_enabled:
            logger.info("段落合并已禁用")
            return regions
        
        # Get configuration
        # 优先使用竖版模式的参数
        if hasattr(self, '_portrait_paragraph_y_gap_max'):
            max_vertical_gap = self._portrait_paragraph_y_gap_max
            logger.info(f"使用竖版模式的y_gap_max: {max_vertical_gap}")
        else:
            max_vertical_gap = self.config.get('ocr.paragraph_merge.max_vertical_gap', 5)
        
        max_font_size_diff = self.config.get('ocr.paragraph_merge.max_font_size_diff', 0.4)
        
        # Step 1: Group regions into columns based on X coordinate
        columns = self._group_regions_by_column(regions)
        logger.debug(f"Grouped {len(regions)} regions into {len(columns)} columns")
        
        # Step 2: Merge paragraphs within each column
        all_merged_regions = []
        
        for col_idx, column_regions in enumerate(columns):
            logger.debug(f"Processing column {col_idx + 1} with {len(column_regions)} regions")
            
            # Sort by Y coordinate within column
            sorted_regions = sorted(column_regions, key=lambda r: r.bbox[1])
            
            # Group into paragraphs based on vertical spacing
            paragraphs = []
            current_paragraph = [sorted_regions[0]]
            
            for i in range(1, len(sorted_regions)):
                prev = sorted_regions[i - 1]
                curr = sorted_regions[i]
                
                # Calculate vertical gap
                vertical_gap = curr.bbox[1] - prev.bbox[3]
                
                # Calculate font size difference
                font_diff = abs(curr.font_size - prev.font_size) / max(curr.font_size, prev.font_size)
                
                # Merge rules:
                should_merge = False
                reason = ""
                
                # 计算宽高比（用于垂直文本检测）
                prev_width = prev.bbox[2] - prev.bbox[0]
                curr_width = curr.bbox[2] - curr.bbox[0]
                prev_height = prev.bbox[3] - prev.bbox[1]
                curr_height = curr.bbox[3] - curr.bbox[1]
                
                # 如果宽度远小于高度，可能是垂直文本
                is_prev_vertical = prev_width < prev_height * 0.5
                is_curr_vertical = curr_width < curr_height * 0.5
                
                logger.debug(
                    f"  Checking merge: prev='{prev.text[:20]}' (w={prev_width}, h={prev_height}, vertical={is_prev_vertical}, "
                    f"is_vertical_merged={getattr(prev, 'is_vertical_merged', False)}), "
                    f"curr='{curr.text[:20]}' (w={curr_width}, h={curr_height}, vertical={is_curr_vertical}, "
                    f"is_vertical_merged={getattr(curr, 'is_vertical_merged', False)})"
                )
                
                # 特殊规则1：如果当前或前一个区域是垂直合并的文本（如"重要提示"），不要合并
                if hasattr(prev, 'is_vertical_merged') and prev.is_vertical_merged:
                    reason = "前一个区域是垂直合并的文本"
                elif hasattr(curr, 'is_vertical_merged') and curr.is_vertical_merged:
                    reason = "当前区域是垂直合并的文本"
                # 特殊规则1.5：如果当前或前一个区域是字段标签，不要合并
                elif hasattr(prev, 'is_field_label') and prev.is_field_label:
                    reason = "前一个区域是字段标签"
                elif hasattr(curr, 'is_field_label') and curr.is_field_label:
                    reason = "当前区域是字段标签"
                # 特殊规则2：如果当前或前一个区域是垂直排列的文本（通过宽高比判断）
                elif is_prev_vertical:
                    reason = "前一个区域是垂直排列的文本（宽高比检测）"
                elif is_curr_vertical:
                    reason = "当前区域是垂直排列的文本（宽高比检测）"
                elif vertical_gap > max_vertical_gap:
                    reason = f"垂直间距过大 ({vertical_gap}px > {max_vertical_gap}px)"
                elif font_diff > max_font_size_diff:
                    reason = f"字体差异过大 ({font_diff:.1%} > {max_font_size_diff:.1%})"
                elif self._is_label_keyword(curr.text.strip()):
                    # 如果前一个区域是长文本（>10个字），即使当前区域是标签关键词，也尝试合并
                    prev_text_len = len(prev.text.strip())
                    if prev_text_len > 10:
                        should_merge = True
                        reason = f"符合段落特征 (gap={vertical_gap}px, 前一个区域是长文本 len={prev_text_len})"
                        logger.debug(f"  标签关键词 '{curr.text.strip()}' 但前一个区域是长文本，尝试合并")
                    else:
                        reason = "是标签关键词"
                elif self._is_section_title(curr.text.strip()) and not self._is_continuation(curr.text.strip()):
                    reason = "是段落标题"
                elif vertical_gap <= max_vertical_gap:
                    # 只有垂直间距小于等于阈值时才合并
                    should_merge = True
                    reason = f"符合段落特征 (gap={vertical_gap}px <= {max_vertical_gap}px)"
                else:
                    # 垂直间距过大，不合并
                    reason = f"垂直间距过大 ({vertical_gap}px > {max_vertical_gap}px)"
                
                if should_merge:
                    current_paragraph.append(curr)
                    logger.debug(f"  Added to paragraph: '{curr.text[:30]}...' - {reason}")
                else:
                    if len(current_paragraph) > 0:
                        paragraphs.append(current_paragraph)
                    current_paragraph = [curr]
                    logger.debug(f"  New paragraph: '{curr.text[:30]}...' - {reason}")
            
            # Add last paragraph
            if len(current_paragraph) > 0:
                paragraphs.append(current_paragraph)
            
            # Merge each paragraph
            for paragraph in paragraphs:
                if len(paragraph) == 1:
                    all_merged_regions.append(paragraph[0])
                elif len(paragraph) >= 2:
                    merged_region = self._merge_paragraph(paragraph)
                    all_merged_regions.append(merged_region)
                    logger.info(
                        f"Merged paragraph with {len(paragraph)} regions in column {col_idx + 1}: "
                        f"'{merged_region.text[:50]}...'"
                    )
        
        logger.info(f"Paragraph merge: {len(regions)} regions → {len(all_merged_regions)} regions")
        
        # Step 3: 跨列合并 - 合并被竖排文本分隔的段落
        # 例如："1.经营范围:..." + "重要提示"(竖排) + "明,个人独资企业..."
        # 应该合并成："1.经营范围:...明,个人独资企业..."
        print(f"[DEBUG] 开始跨列合并，输入 {len(all_merged_regions)} 个区域")
        all_merged_regions = self._merge_across_vertical_text(all_merged_regions, max_vertical_gap)
        print(f"[DEBUG] 跨列合并完成，输出 {len(all_merged_regions)} 个区域")
        
        # Step 4: URL合并 - 合并URL和前缀文本
        # 例如："国家企业信用信息公示系统网址:" + "http://www.gsxt.gov.cn"
        print(f"[DEBUG] 开始URL合并，输入 {len(all_merged_regions)} 个区域")
        all_merged_regions = self._merge_url_with_prefix(all_merged_regions)
        print(f"[DEBUG] URL合并完成，输出 {len(all_merged_regions)} 个区域")
        
        return all_merged_regions
    
    def _group_regions_by_column(self, regions: List[TextRegion]) -> List[List[TextRegion]]:
        """Group regions into columns based on X coordinate clustering.
        
        Uses a simple clustering algorithm:
        - Regions with similar X start positions (within 50px) are in the same column
        - Vertical merged regions (is_vertical_merged=True) are always in separate columns
        
        Args:
            regions: List of TextRegion objects
            
        Returns:
            List of columns, where each column is a list of regions
        """
        if not regions:
            return []
        
        # Separate vertical merged regions from normal regions
        vertical_merged_regions = [r for r in regions if getattr(r, 'is_vertical_merged', False)]
        normal_regions = [r for r in regions if not getattr(r, 'is_vertical_merged', False)]
        
        # Sort normal regions by X coordinate
        sorted_regions = sorted(normal_regions, key=lambda r: r.bbox[0])
        
        # Cluster normal regions into columns
        columns = []
        if sorted_regions:
            current_column = [sorted_regions[0]]
            column_x_start = sorted_regions[0].bbox[0]
            
            column_threshold = 50  # Regions within 50px are in the same column
            
            for region in sorted_regions[1:]:
                x_start = region.bbox[0]
                
                # Check if this region belongs to the current column
                if abs(x_start - column_x_start) <= column_threshold:
                    current_column.append(region)
                else:
                    # Start a new column
                    columns.append(current_column)
                    current_column = [region]
                    column_x_start = x_start
            
            # Add last column
            if current_column:
                columns.append(current_column)
        
        # Add each vertical merged region as a separate column
        for vm_region in vertical_merged_regions:
            columns.append([vm_region])
        
        logger.debug(f"Grouped regions: {len(normal_regions)} normal + {len(vertical_merged_regions)} vertical merged → {len(columns)} columns")
        
        return columns
    
    def _merge_across_vertical_text(self, regions: List[TextRegion], max_vertical_gap: int) -> List[TextRegion]:
        """跨列合并：合并被竖排文本分隔的段落
        
        例如：
        区域A："1.经营范围:商事主体的经营范围在章程中载明..."
        区域B："重要提示"（竖排文本，宽高比 < 0.5）
        区域C："明,个人独资企业和个体工商户的经营范围..."
        
        如果A和C在垂直方向上接近（y坐标差距小），且B是竖排文本，
        则合并A和C，跳过B。
        
        Args:
            regions: 已经过段落合并的区域列表
            max_vertical_gap: 最大垂直间距
            
        Returns:
            合并后的区域列表
        """
        if len(regions) < 3:
            return regions
        
        # 识别竖排文本区域
        vertical_text_regions = []
        normal_regions = []
        
        for region in regions:
            width = region.bbox[2] - region.bbox[0]
            height = region.bbox[3] - region.bbox[1]
            
            # 竖排文本特征：宽度远小于高度（宽高比 < 0.3）
            if width < height * 0.3:
                vertical_text_regions.append(region)
                logger.debug(f"识别竖排文本: '{region.text[:20]}...' (w={width}, h={height}, ratio={width/height:.2f})")
            else:
                normal_regions.append(region)
        
        if not vertical_text_regions:
            logger.debug("没有竖排文本，跳过跨列合并")
            return regions
        
        logger.info(f"跨列合并: 发现 {len(vertical_text_regions)} 个竖排文本区域")
        
        # 按y坐标排序
        sorted_regions = sorted(regions, key=lambda r: r.bbox[1])
        
        merged_regions = []
        skip_indices = set()
        
        for i in range(len(sorted_regions)):
            if i in skip_indices:
                continue
            
            current = sorted_regions[i]
            
            # 检查是否是竖排文本
            curr_width = current.bbox[2] - current.bbox[0]
            curr_height = current.bbox[3] - current.bbox[1]
            is_curr_vertical = curr_width < curr_height * 0.3
            
            if is_curr_vertical:
                # 竖排文本不参与合并，直接添加
                merged_regions.append(current)
                continue
            
            # 查找下一个非竖排文本区域
            merge_candidates = [current]
            j = i + 1
            
            while j < len(sorted_regions):
                next_region = sorted_regions[j]
                next_width = next_region.bbox[2] - next_region.bbox[0]
                next_height = next_region.bbox[3] - next_region.bbox[1]
                is_next_vertical = next_width < next_height * 0.3
                
                if is_next_vertical:
                    # 跳过竖排文本
                    j += 1
                    continue
                
                # 检查是否可以合并
                # 条件1：y坐标接近（考虑竖排文本的高度）
                y_gap = next_region.bbox[1] - current.bbox[3]
                
                # 条件2：x坐标对齐（左边界差距 < 20px）
                x_diff = abs(next_region.bbox[0] - current.bbox[0])
                
                # 条件3：字体大小相近
                font_diff = abs(next_region.font_size - current.font_size) / max(next_region.font_size, current.font_size)
                
                logger.debug(
                    f"检查跨列合并: '{current.text[:20]}...' + '{next_region.text[:20]}...' "
                    f"(y_gap={y_gap}px, x_diff={x_diff}px, font_diff={font_diff:.1%})"
                )
                
                # 放宽y_gap限制，因为中间可能有竖排文本
                if y_gap < 200 and x_diff < 20 and font_diff < 0.3:
                    merge_candidates.append(next_region)
                    skip_indices.add(j)
                    logger.info(
                        f"跨列合并: '{current.text[:30]}...' + '{next_region.text[:30]}...' "
                        f"(y_gap={y_gap}px, x_diff={x_diff}px)"
                    )
                    j += 1
                else:
                    break
            
            # 合并候选区域
            if len(merge_candidates) == 1:
                merged_regions.append(current)
            else:
                merged_region = self._merge_paragraph(merge_candidates)
                merged_regions.append(merged_region)
                logger.info(f"跨列合并完成: {len(merge_candidates)} 个区域 -> '{merged_region.text[:50]}...'")
        
        logger.info(f"跨列合并: {len(regions)} 个区域 -> {len(merged_regions)} 个区域")
        return merged_regions
    
    def _merge_url_with_prefix(self, regions: List[TextRegion]) -> List[TextRegion]:
        """合并URL和前缀文本
        
        例如：
        区域A："国家企业信用信息公示系统网址:"
        区域B："http://www.gsxt.gov.cn"
        
        如果A以冒号结尾，B是URL，且它们在同一行（y坐标接近），
        则合并A和B。
        
        Args:
            regions: 区域列表
            
        Returns:
            合并后的区域列表
        """
        import re
        
        if len(regions) < 2:
            logger.debug(f"URL合并: 区域数量不足 ({len(regions)} < 2)，跳过")
            return regions
        
        logger.info(f"URL合并: 开始处理 {len(regions)} 个区域")
        
        # URL模式 - 支持OCR识别错误（如 ∥ 被识别成 :∥ 或 ://）
        url_pattern = re.compile(r'^(https?[:∥/]+|http[:∥/]+|www\.|ftp[:∥/]+)', re.IGNORECASE)
        
        merged_regions = []
        skip_indices = set()
        
        for i in range(len(regions)):
            if i in skip_indices:
                continue
            
            current = regions[i]
            current_text = current.text.strip()
            
            logger.debug(f"URL合并: 检查区域 #{i}: '{current_text[:30]}...'")
            
            # 检查当前区域是否以冒号结尾（可能是URL前缀）
            if not (current_text.endswith(':') or current_text.endswith('：')):
                logger.debug(f"  -> 不以冒号结尾，跳过")
                merged_regions.append(current)
                continue
            
            logger.debug(f"  -> 以冒号结尾，查找下一个区域")
            
            # 查找下一个区域
            if i + 1 >= len(regions):
                logger.debug(f"  -> 没有下一个区域")
                merged_regions.append(current)
                continue
            
            next_region = regions[i + 1]
            next_text = next_region.text.strip()
            
            logger.debug(f"  -> 下一个区域: '{next_text[:30]}...'")
            
            # 检查下一个区域是否是URL
            if not url_pattern.match(next_text):
                logger.debug(f"  -> 不是URL格式，跳过")
                merged_regions.append(current)
                continue
            
            logger.debug(f"  -> 是URL格式，检查位置")
            
            # 检查是否在同一行（y坐标接近）
            y_diff = abs(current.bbox[1] - next_region.bbox[1])
            
            # 检查水平距离（应该很近）
            x_gap = next_region.bbox[0] - current.bbox[2]
            
            logger.info(
                f"URL合并检查: '{current_text}' + '{next_text}' "
                f"(y_diff={y_diff}px, x_gap={x_gap}px, "
                f"current_bbox={current.bbox}, next_bbox={next_region.bbox})"
            )
            
            # 如果在同一行（y差距 < 10px）且水平距离合理（< 200px）
            if y_diff < 10 and -10 < x_gap < 200:
                # 合并
                merged_region = self._merge_paragraph([current, next_region])
                merged_regions.append(merged_region)
                skip_indices.add(i + 1)
                logger.info(
                    f"✓ URL合并成功: '{current_text}' + '{next_text}' -> '{merged_region.text}'"
                )
            else:
                logger.debug(f"  -> 位置不符合条件 (y_diff={y_diff}, x_gap={x_gap})，不合并")
                merged_regions.append(current)
        
        logger.info(f"URL合并完成: {len(regions)} 个区域 -> {len(merged_regions)} 个区域")
        return merged_regions
    
    def _is_section_title(self, text: str) -> bool:
        """Check if text is a section title (ends with : or ：).
        
        Args:
            text: Text to check
            
        Returns:
            True if text is a section title
        """
        return text.endswith('：') or text.endswith(':')
    
    def _is_continuation(self, text: str) -> bool:
        """Check if text is a continuation of previous paragraph.
        
        Continuations are texts that:
        - Are long (>10 chars) and end with ：
        - Start with common business verbs (售、发、批、造、用、产、制、销)
        
        Args:
            text: Text to check
            
        Returns:
            True if text is a continuation
        """
        # Long text ending with ： is likely a continuation
        if len(text) > 10 and (text.endswith('：') or text.endswith(':')):
            return True
        
        # Short text starting with a verb
        if len(text) <= 5 and text[0] in ['售', '发', '批', '造', '用', '产', '制', '销']:
            return True
        
        return False
    
    def _is_valid_paragraph(self, regions: List[TextRegion]) -> bool:
        """Check if a group of regions forms a valid paragraph.
        
        Args:
            regions: List of regions to check
            
        Returns:
            True if regions form a valid paragraph
        """
        if len(regions) < 2:
            return False
        
        # Check 1: Consistent horizontal alignment
        x_starts = [r.bbox[0] for r in regions]
        x_std = np.std(x_starts)
        
        # If X coordinates vary too much, not a paragraph
        if x_std > 30:
            return False
        
        # Check 2: Consistent vertical spacing
        vertical_gaps = []
        for i in range(len(regions) - 1):
            gap = regions[i + 1].bbox[1] - regions[i].bbox[3]
            vertical_gaps.append(gap)
        
        # If vertical gaps vary too much, not a paragraph
        if vertical_gaps:
            gap_std = np.std(vertical_gaps)
            if gap_std > 10:
                return False
        
        # Check 3: Similar font sizes
        font_sizes = [r.font_size for r in regions]
        font_std = np.std(font_sizes)
        avg_font = np.mean(font_sizes)
        
        # If font sizes vary too much (>20% of average), not a paragraph
        if font_std > avg_font * 0.2:
            return False
        
        return True
    
    def _merge_paragraph(self, regions: List[TextRegion]) -> TextRegion:
        """Merge a list of regions into a single paragraph region.
        
        Args:
            regions: List of regions to merge (should be sorted top to bottom)
            
        Returns:
            Merged TextRegion
        """
        if len(regions) == 1:
            return regions[0]
        
        # Calculate merged bounding box
        x1 = min(r.bbox[0] for r in regions)
        y1 = min(r.bbox[1] for r in regions)
        x2 = max(r.bbox[2] for r in regions)
        y2 = max(r.bbox[3] for r in regions)
        
        merged_bbox = (x1, y1, x2, y2)
        
        # Combine text (top to bottom order)
        # For Chinese text, we should not add spaces between regions
        # Only add space if both regions contain non-Chinese characters
        merged_parts = []
        for i, r in enumerate(regions):
            text = r.text.strip()
            if not text:
                continue
            
            # Add the text
            if i == 0:
                merged_parts.append(text)
            else:
                # Check if we need a space between this and previous text
                prev_text = merged_parts[-1] if merged_parts else ""
                
                # Add space only if both texts are primarily non-Chinese (Latin, numbers, etc.)
                # For Chinese text, no space is needed
                prev_is_chinese = any('\u4e00' <= c <= '\u9fff' for c in prev_text[-3:])
                curr_is_chinese = any('\u4e00' <= c <= '\u9fff' for c in text[:3])
                
                if prev_is_chinese or curr_is_chinese:
                    # No space for Chinese text
                    merged_parts.append(text)
                else:
                    # Add space for non-Chinese text
                    merged_parts.append(' ' + text)
        
        merged_text = ''.join(merged_parts)
        
        # Calculate weighted average confidence
        total_area = sum(r.area for r in regions)
        if total_area > 0:
            merged_confidence = sum(r.confidence * r.area for r in regions) / total_area
        else:
            merged_confidence = sum(r.confidence for r in regions) / len(regions)
        
        # Use average font size
        merged_font_size = int(sum(r.font_size for r in regions) / len(regions))
        
        # Use weighted average angle
        if total_area > 0:
            merged_angle = sum(r.angle * r.area for r in regions) / total_area
        else:
            merged_angle = sum(r.angle for r in regions) / len(regions)
        
        return TextRegion(
            bbox=merged_bbox,
            text=merged_text,
            confidence=merged_confidence,
            font_size=merged_font_size,
            angle=merged_angle
        )
    
    def _is_group_boundary(
        self, 
        prev: TextRegion, 
        curr: TextRegion
    ) -> tuple[bool, str]:
        """Determine if there's a group boundary between two regions.
        
        Uses multiple features to detect label-content boundaries:
        1. Check vertical spacing (if too large, always create boundary)
        2. Check if prev or curr is a field label (is_field_label=True)
        3. Check if prev or curr is a label keyword
        4. Font size difference (only for label detection)
        5. Text length jump (only for label → content)
        6. Semantic matching (label keywords)
        7. Spacing ratio (auxiliary)
        
        Args:
            prev: Previous region
            curr: Current region
            
        Returns:
            Tuple of (is_boundary, reason)
        """
        # Priority check 0: Vertical spacing
        # If vertical spacing is too large, they should not be merged
        # (even if they're in the same "line" due to similar X coordinates)
        prev_y_bottom = prev.bbox[3]
        prev_y_top = prev.bbox[1]
        curr_y_top = curr.bbox[1]
        curr_y_bottom = curr.bbox[3]
        
        # Calculate vertical gap (can be negative if curr is above prev)
        # Use the minimum distance between the two regions
        if curr_y_top >= prev_y_bottom:
            # curr is below prev
            vertical_gap = curr_y_top - prev_y_bottom
        elif prev_y_top >= curr_y_bottom:
            # prev is below curr
            vertical_gap = prev_y_top - curr_y_bottom
        else:
            # They overlap vertically
            vertical_gap = 0
        
        # If vertical gap > 10px, create boundary
        if vertical_gap > 10:
            logger.debug(f"  Vertical gap check: gap={vertical_gap}px")
            return True, f"垂直间距过大 ({vertical_gap}px > 10px)"
        
        # Priority check: If prev or curr is marked as a field label, always create boundary
        prev_is_field_label = hasattr(prev, 'is_field_label') and prev.is_field_label
        curr_is_field_label = hasattr(curr, 'is_field_label') and curr.is_field_label
        
        # If both are field labels, they should NOT be merged (each is a separate field)
        if prev_is_field_label and curr_is_field_label:
            return True, "两个字段标签不合并"
        
        # If prev is a field label and curr is not, create boundary (label → content)
        if prev_is_field_label and not curr_is_field_label:
            return True, "字段标签→内容"
        
        # If curr is a field label and prev is not, create boundary (content → label)
        if curr_is_field_label and not prev_is_field_label:
            return True, "内容→字段标签"
        
        # Get configuration thresholds
        font_diff_threshold = self.config.get('ocr.smart_merge.font_size_diff_threshold', 0.2)
        length_jump_min = self.config.get('ocr.smart_merge.length_jump_threshold', [2, 5])[0]
        length_jump_max = self.config.get('ocr.smart_merge.length_jump_threshold', [2, 5])[1]
        gap_ratio_threshold = self.config.get('ocr.smart_merge.gap_ratio_threshold', 2.0)
        
        prev_len = len(prev.text.strip())
        curr_len = len(curr.text.strip())
        
        # Check if prev or curr is a label keyword
        prev_is_label = self._is_label_keyword(prev.text)
        curr_is_label = self._is_label_keyword(curr.text)
        
        # Strategy 1: If NEITHER prev NOR curr is a label, use relaxed merging
        # Only check for very large gaps or very large font differences
        if not prev_is_label and not curr_is_label:
            # Check for very large spacing (indicates different fields)
            gap = curr.bbox[0] - prev.bbox[2]
            avg_width = ((prev.bbox[2] - prev.bbox[0]) + (curr.bbox[2] - curr.bbox[0])) / 2
            gap_ratio = gap / avg_width if avg_width > 0 else 0
            
            # Get absolute gap threshold from config
            absolute_gap_threshold = self.config.get('ocr.smart_merge.absolute_gap_threshold', 100)
            
            # Check both gap ratio and absolute gap
            # If gap is very large (>=threshold) or gap ratio is large (>2.5), it's a boundary
            if gap >= absolute_gap_threshold or gap_ratio > 2.5:
                return True, f"间距过大 (gap={gap:.0f}px, gap_ratio={gap_ratio:.1f})"
            
            # Check for very large font size difference (>40%)
            font_diff = abs(curr.font_size - prev.font_size) / max(curr.font_size, prev.font_size)
            if font_diff > 0.4:
                return True, f"字体大小差异过大 ({font_diff:.1%} > 40.0%)"
            
            # Otherwise, merge (relaxed strategy for non-labels)
            return False, "非标签区域，宽松合并"
        
        # Strategy 2: If prev OR curr IS a label, use strict boundary detection
        
        # Special case: Both are very short (≤3 chars) and close together
        # This handles cases like "成" + "立日期" which should be merged
        if prev_len <= 3 and curr_len <= 3:
            gap = curr.bbox[0] - prev.bbox[2]
            avg_width = ((prev.bbox[2] - prev.bbox[0]) + (curr.bbox[2] - curr.bbox[0])) / 2
            gap_ratio = gap / avg_width if avg_width > 0 else 0
            
            # If gap is small, merge them (they're likely parts of the same label)
            if gap_ratio < gap_ratio_threshold:
                return False, f"短文本近距离合并 (gap_ratio={gap_ratio:.1f})"
        
        # Feature 1: Font size difference (strict for labels)
        font_diff = abs(curr.font_size - prev.font_size) / max(curr.font_size, prev.font_size)
        if font_diff > font_diff_threshold:
            return True, f"字体大小差异 ({font_diff:.1%} > {font_diff_threshold:.1%})"
        
        # Feature 2: Text length jump (short label → long content OR long content → short label)
        # BUT: Skip if prev is a single-char that's part of a multi-char label
        # (e.g., "围" is part of "范围", should not trigger boundary)
        if prev_len <= length_jump_min and curr_len >= length_jump_max:
            # Check if prev is a single char that's part of a larger label
            if prev_len == 1 and prev_is_label:
                label_keywords = self.config.get('ocr.smart_merge.label_keywords', [])
                is_part_of_larger_label = any(
                    prev.text.strip() in kw and len(kw) > 1
                    for kw in label_keywords
                )
                if is_part_of_larger_label:
                    # This is likely a field fragment (e.g., "围" from "范围")
                    # Don't create boundary - let it merge with content
                    return False, f"单字标签碎片，不分割 ('{prev.text.strip()}'是多字标签的一部分)"
            
            return True, f"文字长度突变 ({prev_len}字 → {curr_len}字)"
        if prev_len >= length_jump_max and curr_len <= length_jump_min:
            return True, f"文字长度突变 ({prev_len}字 → {curr_len}字)"
        
        # Feature 3: Semantic matching
        # Case 1: label → content
        if prev_is_label and not curr_is_label:
            return True, "语义边界 (标签→内容)"
        # Case 2: content → label
        if not prev_is_label and curr_is_label:
            return True, "语义边界 (内容→标签)"
        
        # Feature 4: Spacing (both absolute and relative)
        gap = curr.bbox[0] - prev.bbox[2]
        avg_width = ((prev.bbox[2] - prev.bbox[0]) + (curr.bbox[2] - curr.bbox[0])) / 2
        gap_ratio = gap / avg_width if avg_width > 0 else 0
        
        # Get absolute gap threshold from config (default: 100px)
        absolute_gap_threshold = self.config.get('ocr.smart_merge.absolute_gap_threshold', 100)
        
        # Check absolute gap first (for cases like "市场主体..." | "国家市场监督...")
        # If gap is very large (>=100px by default), it's likely a boundary
        if gap >= absolute_gap_threshold:
            return True, f"绝对间距过大 (gap={gap:.0f}px >= {absolute_gap_threshold}px)"
        
        # Then check gap ratio
        if gap_ratio > gap_ratio_threshold:
            # Check if font and length are similar (might be same group despite large gap)
            if font_diff < 0.1 and abs(prev_len - curr_len) <= 1:
                # Similar font and length, probably same group (e.g., "名" and "称")
                return False, f"间距大但特征相似 (gap_ratio={gap_ratio:.1f})"
            else:
                return True, f"间距比例过大 (gap_ratio={gap_ratio:.1f} > {gap_ratio_threshold})"
        
        return False, "同一组"
    
    def _is_label_keyword(self, text: str) -> bool:
        """Check if text is a label keyword or starts with a label keyword.
        
        Args:
            text: Text to check
            
        Returns:
            True if text is a label keyword or starts with a label keyword
        """
        # Get label keywords from config
        label_keywords = self.config.get('ocr.smart_merge.label_keywords', [
            "名", "称", "名称",
            "注", "册", "资", "本", "注册资本",
            "法", "定", "代", "表", "人", "法定代表人",
            "住", "所", "住所",
            "成", "立", "日", "期", "成立日期",
            "营", "业", "期", "限", "营业期限",
            "经", "营", "范", "围", "经营范围",
            "类", "型", "类型",
        ])
        
        text_stripped = text.strip()
        
        # Check if text is exactly a label keyword
        if text_stripped in label_keywords:
            return True
        
        # Check if text starts with a label keyword (for cases like "名 称佛山...")
        # Only check multi-character keywords to avoid false positives
        for keyword in label_keywords:
            if len(keyword) >= 2 and text_stripped.startswith(keyword):
                # Make sure there's a clear boundary after the keyword
                # (e.g., space, punctuation, or different character type)
                if len(text_stripped) > len(keyword):
                    next_char = text_stripped[len(keyword)]
                    # If next char is space, punctuation, or uppercase letter, it's likely a boundary
                    if next_char in ' ：:，,。.！!？?；;、' or next_char.isupper():
                        return True
                    # If next char is a Chinese character that's not part of the keyword, it's likely content
                    if '\u4e00' <= next_char <= '\u9fff':
                        return True
        
        return False
    
    def _merge_group(self, group: List[TextRegion]) -> TextRegion:
        """Merge a group of regions into one.
        
        Args:
            group: List of regions to merge
            
        Returns:
            Merged TextRegion
        """
        if len(group) == 1:
            return group[0]
        
        # Merge sequentially
        merged = group[0]
        for region in group[1:]:
            merged = self._merge_two_regions(merged, region)
        
        return merged
    
    def _calculate_region_distance(
        self, 
        region1: TextRegion, 
        region2: TextRegion
    ) -> float:
        """Calculate the minimum distance between two text regions.
        
        Args:
            region1: First text region
            region2: Second text region
            
        Returns:
            Minimum distance between the bounding boxes in pixels
        """
        x1_1, y1_1, x2_1, y2_1 = region1.bbox
        x1_2, y1_2, x2_2, y2_2 = region2.bbox
        
        # Calculate horizontal distance
        if x2_1 < x1_2:
            dx = x1_2 - x2_1
        elif x2_2 < x1_1:
            dx = x1_1 - x2_2
        else:
            dx = 0  # Overlapping horizontally
        
        # Calculate vertical distance
        if y2_1 < y1_2:
            dy = y1_2 - y2_1
        elif y2_2 < y1_1:
            dy = y1_1 - y2_2
        else:
            dy = 0  # Overlapping vertically
        
        return math.sqrt(dx * dx + dy * dy)
    
    def _is_horizontal_merge(
        self,
        region1: TextRegion,
        region2: TextRegion
    ) -> bool:
        """判断两个区域是否应该进行水平合并（或垂直合并）.
        
        改进逻辑：
        1. 首先判断是水平排列还是垂直排列
        2. 对于水平排列：检查是否在同一行、水平距离、字体大小
        3. 对于垂直排列：检查是否在同一列、垂直距离、字体大小
        4. 所有条件都满足才返回 True（允许合并）
        
        这样可以：
        - 合并相隔较远但字体相同的文字 ✅
        - 不合并单字（大字）和紧接着的小字 ✅
        - 不合并图片左右两边相隔太远的文字 ✅
        - 合并垂直排列的文字（如"重要提示"）✅
        
        Args:
            region1: 第一个文本区域
            region2: 第二个文本区域
            
        Returns:
            True 表示可以合并（水平或垂直），False 表示不合并
        """
        x1_1, y1_1, x2_1, y2_1 = region1.bbox
        x1_2, y1_2, x2_2, y2_2 = region2.bbox
        
        # 计算中心点
        center_x1 = (x1_1 + x2_1) / 2
        center_y1 = (y1_1 + y2_1) / 2
        center_x2 = (x1_2 + x2_2) / 2
        center_y2 = (y1_2 + y2_2) / 2
        
        # 计算宽度和高度
        width1 = x2_1 - x1_1
        height1 = y2_1 - y1_1
        width2 = x2_2 - x1_2
        height2 = y2_2 - y1_2
        
        avg_width = (width1 + width2) / 2
        avg_height = (height1 + height2) / 2
        
        # 计算中心点距离
        horizontal_center_diff = abs(center_x1 - center_x2)
        vertical_center_diff = abs(center_y1 - center_y2)
        
        # 判断是水平排列还是垂直排列
        # 如果水平距离 > 垂直距离，认为是水平排列
        # 如果垂直距离 > 水平距离，认为是垂直排列
        is_horizontal_layout = horizontal_center_diff > vertical_center_diff
        
        # 检查字体大小是否相近（对于水平和垂直都适用）
        font_size1 = region1.font_size
        font_size2 = region2.font_size
        size_diff_ratio = abs(font_size1 - font_size2) / max(font_size1, font_size2)
        
        if size_diff_ratio > self.font_size_diff_threshold:
            logger.debug(
                f"Font size mismatch: "
                f"font1={font_size1}, font2={font_size2}, "
                f"diff={size_diff_ratio:.2%}, threshold={self.font_size_diff_threshold:.2%}"
            )
            return False
        
        if is_horizontal_layout:
            # 水平排列：检查是否在同一行
            if vertical_center_diff >= avg_height * 0.5:
                logger.debug(f"Not same line: vertical_diff={vertical_center_diff:.2f}, avg_height={avg_height:.2f}")
                return False
            
            # 检查水平距离
            if x2_1 < x1_2:
                horizontal_distance = x1_2 - x2_1
            elif x2_2 < x1_1:
                horizontal_distance = x1_1 - x2_2
            else:
                horizontal_distance = 0  # 水平重叠
            
            if horizontal_distance > self.merge_threshold_horizontal:
                logger.debug(
                    f"Horizontal distance too large: "
                    f"distance={horizontal_distance:.2f}, threshold={self.merge_threshold_horizontal}"
                )
                return False
            
            logger.debug(
                f"Can merge (horizontal): same line, distance={horizontal_distance:.2f}px, "
                f"similar font size (font1={font_size1}, font2={font_size2})"
            )
            return True
        else:
            # 垂直排列：检查是否在同一列
            if horizontal_center_diff >= avg_width * 0.5:
                logger.debug(f"Not same column: horizontal_diff={horizontal_center_diff:.2f}, avg_width={avg_width:.2f}")
                return False
            
            # 检查垂直距离
            if y2_1 < y1_2:
                vertical_distance = y1_2 - y2_1
            elif y2_2 < y1_1:
                vertical_distance = y1_1 - y2_2
            else:
                vertical_distance = 0  # 垂直重叠
            
            # 对于垂直文本，使用更宽松的垂直距离阈值
            # 因为垂直文本的字间距通常比水平文本大
            vertical_threshold = self.merge_threshold_horizontal  # 使用相同的阈值
            
            if vertical_distance > vertical_threshold:
                logger.debug(
                    f"Vertical distance too large: "
                    f"distance={vertical_distance:.2f}, threshold={vertical_threshold}"
                )
                return False
            
            logger.debug(
                f"Can merge (vertical): same column, distance={vertical_distance:.2f}px, "
                f"similar font size (font1={font_size1}, font2={font_size2})"
            )
            return True
    
    def _merge_two_regions(
        self, 
        region1: TextRegion, 
        region2: TextRegion
    ) -> TextRegion:
        """Merge two text regions into one.
        
        Args:
            region1: First text region
            region2: Second text region
            
        Returns:
            New TextRegion containing both regions
        """
        x1_1, y1_1, x2_1, y2_1 = region1.bbox
        x1_2, y1_2, x2_2, y2_2 = region2.bbox
        
        # Calculate merged bounding box
        merged_bbox = (
            min(x1_1, x1_2),
            min(y1_1, y1_2),
            max(x2_1, x2_2),
            max(y2_1, y2_2)
        )
        
        # Combine text (order by position: left-to-right, top-to-bottom)
        # If Y coordinates are very close (< 5 pixels), consider them on the same line
        y_diff = abs(y1_1 - y1_2)
        if y_diff < 5:
            # Same line, order by X coordinate only
            if x1_1 <= x1_2:
                merged_text = region1.text + " " + region2.text
            else:
                merged_text = region2.text + " " + region1.text
        elif y1_1 < y1_2:
            # region1 is above region2
            merged_text = region1.text + " " + region2.text
        else:
            # region2 is above region1
            merged_text = region2.text + " " + region1.text
        
        # Use weighted average for confidence based on area
        total_area = region1.area + region2.area
        if total_area > 0:
            merged_confidence = (
                region1.confidence * region1.area + 
                region2.confidence * region2.area
            ) / total_area
        else:
            merged_confidence = (region1.confidence + region2.confidence) / 2
        
        # Use the larger font size
        merged_font_size = max(region1.font_size, region2.font_size)
        
        # Use weighted average for angle
        if total_area > 0:
            merged_angle = (
                region1.angle * region1.area + 
                region2.angle * region2.area
            ) / total_area
        else:
            merged_angle = (region1.angle + region2.angle) / 2
        
        return TextRegion(
            bbox=merged_bbox,
            text=merged_text,
            confidence=merged_confidence,
            font_size=merged_font_size,
            angle=merged_angle
        )
    
    def _would_cross_qrcode(
        self,
        region1: TextRegion,
        region2: TextRegion,
        qr_bbox: Tuple[int, int, int, int]
    ) -> bool:
        """检查合并两个区域后是否会跨越二维码边界。
        
        Args:
            region1: 第一个区域
            region2: 第二个区域
            qr_bbox: 二维码边界框 (x1, y1, x2, y2)
            
        Returns:
            True如果合并后会跨越二维码边界
        """
        # 计算合并后的bbox
        x1_1, y1_1, x2_1, y2_1 = region1.bbox
        x1_2, y1_2, x2_2, y2_2 = region2.bbox
        
        merged_x1 = min(x1_1, x1_2)
        merged_y1 = min(y1_1, y1_2)
        merged_x2 = max(x2_1, x2_2)
        merged_y2 = max(y2_1, y2_2)
        
        qr_x1, qr_y1, qr_x2, qr_y2 = qr_bbox
        
        # 检查合并后的bbox是否与二维码重叠
        # 如果两个区域都不与二维码重叠，但合并后会重叠，说明跨越了二维码
        
        # 检查region1是否与二维码重叠
        r1_overlaps = not (x2_1 < qr_x1 or x1_1 > qr_x2 or y2_1 < qr_y1 or y1_1 > qr_y2)
        
        # 检查region2是否与二维码重叠
        r2_overlaps = not (x2_2 < qr_x1 or x1_2 > qr_x2 or y2_2 < qr_y1 or y1_2 > qr_y2)
        
        # 检查合并后是否与二维码重叠
        merged_overlaps = not (merged_x2 < qr_x1 or merged_x1 > qr_x2 or merged_y2 < qr_y1 or merged_y1 > qr_y2)
        
        # 如果合并后会重叠，但两个区域都不重叠，说明跨越了二维码
        # 或者，如果合并后的bbox包含了二维码的大部分区域，也认为是跨越
        if merged_overlaps:
            # 计算重叠区域
            overlap_x1 = max(merged_x1, qr_x1)
            overlap_y1 = max(merged_y1, qr_y1)
            overlap_x2 = min(merged_x2, qr_x2)
            overlap_y2 = min(merged_y2, qr_y2)
            
            if overlap_x1 < overlap_x2 and overlap_y1 < overlap_y2:
                overlap_area = (overlap_x2 - overlap_x1) * (overlap_y2 - overlap_y1)
                qr_area = (qr_x2 - qr_x1) * (qr_y2 - qr_y1)
                overlap_ratio = overlap_area / qr_area if qr_area > 0 else 0
                
                # 如果重叠超过二维码面积的10%，认为是跨越
                if overlap_ratio > 0.1:
                    logger.debug(
                        f"合并会跨越二维码: "
                        f"region1={region1.bbox}, region2={region2.bbox}, "
                        f"merged=({merged_x1},{merged_y1},{merged_x2},{merged_y2}), "
                        f"qr={qr_bbox}, overlap_ratio={overlap_ratio:.1%}"
                    )
                    return True
        
        return False

    def filter_low_confidence(
        self, 
        regions: List[TextRegion]
    ) -> List[TextRegion]:
        """Filter out text regions with low confidence scores.
        
        Args:
            regions: List of TextRegion objects to filter
            
        Returns:
            List of TextRegion objects with confidence >= threshold
        """
        filtered = []
        removed_count = 0
        
        for region in regions:
            # 基本置信度过滤
            if region.confidence < self.confidence_threshold:
                removed_count += 1
                logger.debug(
                    f"Filtered low confidence region: "
                    f"text='{region.text[:20]}...', "
                    f"confidence={region.confidence:.2f}"
                )
                continue
            
            # 智能过滤：检测可能的误识别
            text = region.text.strip()
            
            # 规则1: 过滤纯英文字母的短文本（可能是水印）
            # 但保留常见的英文单词和缩写
            if len(text) <= 10 and text.replace(' ', '').isalpha():
                # 检查是否全是大写字母（水印特征）
                if text.replace(' ', '').isupper():
                    # 检查是否是常见的合法文本
                    common_words = ['HTTP', 'HTTPS', 'WWW', 'COM', 'CN', 'GOV', 'ORG', 'NET']
                    if not any(word in text.upper() for word in common_words):
                        removed_count += 1
                        logger.debug(
                            f"Filtered suspected watermark (uppercase): "
                            f"text='{text}', confidence={region.confidence:.2f}"
                        )
                        continue
            
            # 规则2: 过滤包含大量生僻字或乱码的文本（置信度较低时）
            if region.confidence < 0.8:
                # 检查是否包含过多生僻字
                rare_chars = sum(1 for c in text if ord(c) > 0x9FFF)
                if rare_chars > len(text) * 0.3:  # 超过30%是生僻字
                    removed_count += 1
                    logger.debug(
                        f"Filtered text with rare characters: "
                        f"text='{text[:20]}...', confidence={region.confidence:.2f}"
                    )
                    continue
            
            # 规则3: 过滤明显的乱码文本（包含大量不常见汉字）
            # 注意：不能简单地用Unicode值判断生僻字，因为很多常用字（如"若"、"言"、"育"、"询"、"限"）
            # 的Unicode值都在0x8000以上。这个规则已禁用，避免误过滤正常文本。
            # 如果需要过滤乱码，应该使用更准确的方法，如检查字符是否在常用汉字表中。
            
            filtered.append(region)
        
        if removed_count > 0:
            logger.info(
                f"Filtered {removed_count} low confidence/suspicious regions "
                f"(threshold={self.confidence_threshold})"
            )
        
        return filtered
    
    def _filter_watermarks(self, regions: List[TextRegion]) -> List[TextRegion]:
        """Filter out watermark text regions using blacklist patterns.
        
        This method is called BEFORE merging to prevent watermarks from
        being merged with legitimate text.
        
        Args:
            regions: List of TextRegion objects to filter
            
        Returns:
            List of TextRegion objects with watermarks removed
        """
        if not self.text_blacklist:
            return regions
        
        filtered = []
        removed_count = 0
        
        for region in regions:
            text = region.text.strip()
            is_watermark = False
            
            # Check against blacklist patterns
            for pattern in self.text_blacklist:
                if pattern.search(text):
                    removed_count += 1
                    logger.debug(
                        f"Filtered watermark (pre-merge): "
                        f"text='{text}', pattern='{pattern.pattern}', "
                        f"confidence={region.confidence:.2f}"
                    )
                    is_watermark = True
                    break
            
            if not is_watermark:
                filtered.append(region)
        
        if removed_count > 0:
            logger.info(
                f"Filtered {removed_count} watermark regions before merging"
            )
        
        return filtered
    
    def filter_small_area(self, regions: List[TextRegion]) -> List[TextRegion]:
        """Filter out text regions with small area.
        
        Args:
            regions: List of TextRegion objects to filter
            
        Returns:
            List of TextRegion objects with area >= min_text_area
        """
        filtered = []
        removed_count = 0
        
        for region in regions:
            if region.area >= self.min_text_area:
                filtered.append(region)
            else:
                removed_count += 1
                logger.debug(
                    f"Filtered small area region: "
                    f"text='{region.text[:20] if len(region.text) > 20 else region.text}', "
                    f"area={region.area}"
                )
        
        if removed_count > 0:
            logger.info(
                f"Filtered {removed_count} small area regions "
                f"(min_area={self.min_text_area})"
            )
        
        return filtered
    
    def filter_regions(self, regions: List[TextRegion], image: Optional[np.ndarray] = None) -> List[TextRegion]:
        """Apply all filters to text regions.
        
        Applies both confidence and area filters in sequence.
        
        Args:
            regions: List of TextRegion objects to filter
            image: Optional image array for edge noise detection
            
        Returns:
            List of filtered TextRegion objects
        """
        # Apply confidence filter first
        filtered = self.filter_low_confidence(regions)
        
        # Then apply area filter
        filtered = self.filter_small_area(filtered)
        
        # Finally apply noise filter (remove single chars and small text)
        filtered = self.filter_noise(filtered, image)
        
        return filtered
    
    def filter_noise(self, regions: List[TextRegion], image: Optional[np.ndarray] = None) -> List[TextRegion]:
        """Filter out noise regions (single characters, small text, edge noise).
        
        过滤噪点的策略：
        1. 过滤单个字符（如"T"、"。"、"："等）
        2. 过滤面积特别小的文本（< 200像素）
        3. 保留重要的单字标签（如"名称"、"类型"等字段名的一部分）
        4. 过滤边缘区域的单字符（即使是重要字符）
        
        Args:
            regions: List of TextRegion objects to filter
            image: Optional image array for calculating edge threshold
            
        Returns:
            List of filtered TextRegion objects without noise
        """
        # 从配置读取噪点过滤设置
        # 优先使用竖版专用配置（如果存在）
        if hasattr(self, '_portrait_noise_filter_enabled'):
            noise_filter_enabled = self._portrait_noise_filter_enabled
        else:
            noise_filter_enabled = self.config.get('ocr.noise_filter.enabled', False)
        
        if not noise_filter_enabled:
            return regions
        
        # 读取其他配置参数
        if hasattr(self, '_portrait_noise_filter_min_area'):
            min_noise_area = self._portrait_noise_filter_min_area
        else:
            min_noise_area = self.config.get('ocr.noise_filter.min_area', 200)
        
        if hasattr(self, '_portrait_noise_filter_single_char'):
            filter_single_char = self._portrait_noise_filter_single_char
        else:
            filter_single_char = self.config.get('ocr.noise_filter.filter_single_char', True)
        
        # 重要的单字列表（不应该被过滤）
        important_single_chars = set(self.config.get('ocr.noise_filter.important_single_chars', [
            '名', '称', '型', '类', '址', '住', '所', '法', '定', '代', '表', '人',
            '注', '册', '资', '本', '成', '立', '日', '期', '营', '业', '限',
            '经', '范', '围', '登', '记', '机', '关', '统', '一', '社', '会',
            '信', '用', '码', '组', '织', '构', '税', '务', '证', '号',
            '开', '户', '银', '行', '账'
        ]))
        
        # 计算边缘阈值（用于过滤边缘噪点）
        right_edge_threshold = None
        if image is not None:
            # 获取图片宽度
            image_width = image.shape[1]
            # 右边缘阈值：图片宽度的85%
            # 任何在右边缘85%之后的单字符都会被过滤（即使是重要字符）
            right_edge_threshold = image_width * 0.85
            logger.debug(f"Edge noise filter: right_edge_threshold = {right_edge_threshold:.0f} (image_width={image_width})")
        
        filtered = []
        removed_count = 0
        edge_noise_count = 0
        
        for region in regions:
            text = region.text.strip()
            
            # 跳过空文本
            if not text:
                removed_count += 1
                logger.debug(f"Filtered empty text region: bbox={region.bbox}")
                continue
            
            # 边缘噪点检测：如果是单字符且在右边缘，过滤掉（即使是重要字符）
            if right_edge_threshold is not None and len(text) <= 2:
                x1, y1, x2, y2 = region.bbox
                if x1 > right_edge_threshold:
                    removed_count += 1
                    edge_noise_count += 1
                    logger.info(f"Filtered edge noise: '{text}' at x={x1:.0f} (threshold={right_edge_threshold:.0f}), area={region.area}")
                    continue
            
            # 过滤单个字符（除非是重要字符）
            if filter_single_char and len(text) == 1:
                if text not in important_single_chars:
                    removed_count += 1
                    logger.debug(f"Filtered single char noise: '{text}', area={region.area}, bbox={region.bbox}")
                    continue
            
            # 过滤面积特别小的文本（可能是噪点）
            if region.area < min_noise_area:
                # 如果是重要的单字，保留
                if len(text) == 1 and text in important_single_chars:
                    filtered.append(region)
                else:
                    removed_count += 1
                    logger.debug(f"Filtered small noise: '{text}', area={region.area}, bbox={region.bbox}")
                    continue
            
            # 通过所有过滤条件
            filtered.append(region)
        
        if removed_count > 0:
            logger.info(f"Filtered {removed_count} noise regions (single chars: {removed_count - edge_noise_count}, edge noise: {edge_noise_count})")
        
        return filtered
    
    def _split_merged_regions(self, regions: List[TextRegion], image: np.ndarray) -> List[TextRegion]:
        """Split regions that were incorrectly merged by PaddleOCR.
        
        This method detects regions that contain both labels and content
        (e.g., "称佛山怡柔坊纺织品制造有限公司") and attempts to split them.
        
        Detection criteria:
        - Text starts with a SHORT label keyword (1-2 chars ONLY)
        - Followed by substantial content (5+ chars)
        - Content is NOT another label keyword
        
        Args:
            regions: List of TextRegion objects
            image: Original image for re-OCR if needed
            
        Returns:
            List of TextRegion objects with split regions
        """
        if not self.config.get('ocr.smart_merge.enabled', False):
            logger.info("Smart merge disabled, skipping region splitting")
            return regions
        
        logger.info(f"开始拆分错误合并的区域，输入 {len(regions)} 个区域")
        
        # Print first few regions for debugging
        for i, r in enumerate(regions[:10]):
            logger.debug(f"  Input region #{i+1}: '{r.text}'")
        
        # First, sort regions by position (top to bottom, left to right)
        sorted_regions = sorted(regions, key=lambda r: (r.bbox[1], r.bbox[0]))
        
        split_regions = []
        split_count = 0
        
        for idx, region in enumerate(sorted_regions):
            text = region.text.strip()
            
            logger.debug(f"检查区域 #{idx+1}: '{text}'")
            
            # Skip short text (no need to split)
            if len(text) < 3:
                logger.debug(f"  跳过：文本太短 (len={len(text)})")
                split_regions.append(region)
                continue
            
            # Get label keywords
            label_keywords = self.config.get('ocr.smart_merge.label_keywords', [])
            
            # Try to find the longest matching label keyword at the beginning
            # Split for ALL label keywords (not just 1-2 char ones)
            # IMPORTANT: Also check for keywords with spaces (e.g., "名 称")
            found_label = None
            label_end_pos = 0
            
            # Sort by length (longest first) to match longer labels first
            sorted_keywords = sorted(label_keywords, key=len, reverse=True)
            
            logger.debug(f"  检查 {len(sorted_keywords)} 个标签关键词...")
            
            for keyword in sorted_keywords:
                # Check both with and without spaces
                # e.g., "名称" and "名 称"
                keyword_with_spaces = ' '.join(keyword)  # "名称" -> "名 称"
                
                logger.debug(f"    尝试匹配: '{keyword}' 或 '{keyword_with_spaces}'")
                
                # Try exact match first
                if text.startswith(keyword):
                    match_keyword = keyword
                    match_end_pos = len(keyword)
                    logger.debug(f"      匹配成功（无空格）: '{keyword}'")
                # Try with spaces
                elif text.startswith(keyword_with_spaces):
                    match_keyword = keyword_with_spaces
                    match_end_pos = len(keyword_with_spaces)
                    logger.debug(f"      匹配成功（带空格）: '{keyword_with_spaces}'")
                else:
                    continue
                
                # IMPORTANT: Skip single-char keywords that are likely part of a field name
                # (e.g., "围" in "围以审批..." is part of "范围", not a standalone label)
                # Only split if:
                # 1. Keyword is 2+ chars (e.g., "名称", "类型"), OR
                # 2. Keyword is 1 char AND is a complete field name (e.g., "名", "称" when standalone)
                if len(keyword) == 1:
                    # Check if this single char is likely part of a larger field name
                    # by checking if it appears in any multi-char field names
                    is_part_of_larger_field = any(
                        keyword in kw and len(kw) > 1
                        for kw in label_keywords
                    )
                    
                    if is_part_of_larger_field:
                        # **关键修复**: 即使单字是更大字段的一部分,如果后面跟着非标签内容,仍然应该拆分
                        # 例如: "称佛山..." 中的"称"是"名称"的一部分,但"佛山"不是标签,应该拆分
                        remaining = text[match_end_pos:].strip()
                        
                        # 检查remaining是否以标签关键词开头
                        is_remaining_label = any(
                            remaining.startswith(kw) for kw in label_keywords 
                            if len(kw) <= 3  # 只检查短关键词
                        )
                        
                        if is_remaining_label:
                            # remaining是标签,跳过拆分(例如: "名称"不应该拆分成"名"+"称")
                            logger.debug(f"    跳过单字关键词'{keyword}': 是更大字段的一部分,且后面是标签")
                            continue
                        else:
                            # remaining不是标签,应该拆分(例如: "称佛山..."应该拆分成"称"+"佛山...")
                            logger.debug(f"    保留单字关键词'{keyword}': 虽然是更大字段的一部分,但后面是非标签内容")
                            # 继续处理,不跳过
                    else:
                        # 单字不是更大字段的一部分,继续处理
                        logger.debug(f"    保留单字关键词'{keyword}': 不是更大字段的一部分")
                
                # Check what comes after the keyword
                remaining = text[match_end_pos:].strip()
                
                # Skip if remaining is empty or too short
                if len(remaining) < 3:
                    continue
                
                # Check if remaining text is also a label keyword
                # If so, don't split (e.g., "注册资本" should not split "注" and "册资本")
                is_remaining_label = any(
                    remaining.startswith(kw) for kw in label_keywords 
                    if len(kw) <= 2  # Only check short keywords
                )
                
                if is_remaining_label:
                    continue
                
                # Check if remaining starts with a non-label character
                # (e.g., company name, address, etc.)
                first_char = remaining[0]
                
                # If first char is a label keyword, don't split
                if first_char in label_keywords:
                    continue
                
                # Found a valid split point
                found_label = match_keyword
                label_end_pos = match_end_pos
                break
            
            # If no valid split point found, keep region as is
            if not found_label:
                split_regions.append(region)
                continue
            
            # Split the region
            label_text = text[:label_end_pos]
            content_text = text[label_end_pos:].strip()
            
            logger.info(
                f"Splitting region: '{text}' → '{label_text}' | '{content_text}'"
            )
            
            # Calculate split position with improved accuracy
            x1, y1, x2, y2 = region.bbox
            width = x2 - x1
            
            # Check if there's a previous region on the same line that might overlap
            # Look at ALL regions in the ORIGINAL sorted list (both before and after current)
            prev_right_boundary = None
            for i in range(len(sorted_regions)):
                if i == idx:
                    continue  # Skip current region
                    
                prev_region = sorted_regions[i]
                prev_y_center = (prev_region.bbox[1] + prev_region.bbox[3]) / 2
                curr_y_center = (y1 + y2) / 2
                
                # Check if on the same line (within 10px vertically)
                if abs(prev_y_center - curr_y_center) < 10:
                    # Check if prev is to the left of current (prev's right <= current's right)
                    if prev_region.bbox[2] <= x2 and prev_region.bbox[0] < x1:
                        if prev_right_boundary is None or prev_region.bbox[2] > prev_right_boundary:
                            prev_right_boundary = prev_region.bbox[2]
                            logger.debug(
                                f"Found previous region on same line: '{prev_region.text}' "
                                f"(right boundary: {prev_right_boundary}px)"
                            )
            
            # Adjust left boundary if it overlaps with previous region
            adjusted_x1 = x1
            if prev_right_boundary is not None:
                logger.debug(
                    f"Checking overlap: current left={x1}px, prev right={prev_right_boundary}px"
                )
                if x1 < prev_right_boundary:
                    # Add a small gap (5px) to avoid touching
                    adjusted_x1 = prev_right_boundary + 5
                    logger.info(
                        f"Adjusted left boundary from {x1}px to {adjusted_x1}px "
                        f"to avoid overlap with previous region (right boundary: {prev_right_boundary}px)"
                    )
                elif x1 - prev_right_boundary < 10:
                    # If gap is too small, add a bit more space
                    adjusted_x1 = prev_right_boundary + 10
                    logger.info(
                        f"Adjusted left boundary from {x1}px to {adjusted_x1}px "
                        f"to ensure minimum gap with previous region (right boundary: {prev_right_boundary}px)"
                    )
            
            # Recalculate width with adjusted boundary
            adjusted_width = x2 - adjusted_x1
            
            # Method 1: Character ratio (baseline)
            char_ratio = len(label_text) / len(text)
            split_x_char = int(adjusted_x1 + adjusted_width * char_ratio)
            
            # Method 2: Estimate based on average character width
            # Assume Chinese characters are roughly square (width ≈ height)
            avg_char_width = adjusted_width / len(text)
            label_width_estimate = len(label_text) * avg_char_width
            split_x_width = int(adjusted_x1 + label_width_estimate)
            
            # Method 3: Conservative approach - use the larger split position
            # This ensures content region doesn't overlap with label region
            # Add a small buffer (5% of width or 5px, whichever is larger)
            buffer = max(int(adjusted_width * 0.05), 5)
            split_x = max(split_x_char, split_x_width) + buffer
            
            # Ensure split_x doesn't exceed the right boundary
            split_x = min(split_x, x2 - 10)  # Leave at least 10px for content
            
            logger.debug(
                f"Split position calculation: "
                f"char_ratio={split_x_char}, width_estimate={split_x_width}, "
                f"final={split_x} (buffer={buffer}px), adjusted_x1={adjusted_x1}"
            )
            
            # Create two new regions with adjusted boundaries
            label_region = TextRegion(
                bbox=(adjusted_x1, y1, split_x, y2),
                text=label_text,
                confidence=region.confidence,
                font_size=region.font_size,
                angle=region.angle
            )
            
            content_region = TextRegion(
                bbox=(split_x, y1, x2, y2),
                text=content_text,
                confidence=region.confidence,
                font_size=region.font_size,
                angle=region.angle
            )
            
            split_regions.append(label_region)
            split_regions.append(content_region)
            split_count += 1
        
        if split_count > 0:
            logger.info(f"Split {split_count} incorrectly merged regions")
        
        return split_regions
    
    def _merge_field_fragments(self, regions: List[TextRegion]) -> List[TextRegion]:
        """合并字段碎片（如"名"+"称"→"名称"）。
        
        根据预定义的字段碎片规则,将被OCR错误分离的字段名合并回完整字段。
        这个方法专门处理营业执照等表单文档中常见的字段碎片问题。
        
        **迭代合并策略**：
        - 第一轮：合并"日"+"期"→"日期"
        - 第二轮：合并"立"+"日期"→"立日期"
        - 第三轮：合并"成"+"立日期"→"成立日期"
        - 重复直到没有更多合并
        
        合并条件:
        1. 两个区域的文本匹配某个字段碎片规则
        2. 两个区域在同一行（垂直位置接近）
        3. 两个区域水平距离在规则定义的最大距离内
        4. 字体大小相近
        
        Args:
            regions: 输入的文本区域列表
            
        Returns:
            合并后的文本区域列表
        """
        if not regions or len(regions) < 2:
            logger.info(f"字段碎片合并：输入区域数量不足（{len(regions) if regions else 0}），跳过")
            return regions
        
        logger.info(f"=" * 80)
        logger.info(f"开始字段碎片合并，输入 {len(regions)} 个区域")
        logger.info(f"=" * 80)
        
        # 打印所有输入区域
        for i, r in enumerate(regions):
            logger.info(f"  输入区域 #{i+1}: '{r.text}' (bbox={r.bbox}, font={r.font_size}px)")
        
        # 迭代合并，直到没有更多合并
        current_regions = regions
        total_merge_count = 0
        iteration = 0
        max_iterations = 10  # 防止无限循环
        
        while iteration < max_iterations:
            iteration += 1
            logger.debug(f"字段碎片合并 - 第 {iteration} 轮，输入 {len(current_regions)} 个区域")
            
            # 按从左到右、从上到下排序
            sorted_regions = sorted(current_regions, key=lambda r: (r.bbox[1], r.bbox[0]))
            
            merged_regions = []
            used_indices = set()
            merge_count = 0
            
            for i, region1 in enumerate(sorted_regions):
                if i in used_indices:
                    continue
                
                text1 = region1.text.strip()
                
                # 只检查短文本（可能是字段碎片）
                if len(text1) > 5:
                    merged_regions.append(region1)
                    continue
                
                # 尝试找到匹配的碎片对
                merged = False
                
                for j in range(i + 1, len(sorted_regions)):
                    if j in used_indices:
                        continue
                    
                    region2 = sorted_regions[j]
                    text2 = region2.text.strip()
                    
                    # 检查是否匹配某个字段碎片规则（支持部分匹配）
                    match_result = self._check_field_fragment_match(region1, region2)
                    
                    if match_result:
                        merged_text, max_distance, is_partial_match, remaining_content = match_result
                        
                        logger.debug(
                            f"找到匹配规则: '{text1}' + '{text2}' → '{merged_text}' "
                            f"(max_distance={max_distance}, partial={is_partial_match})"
                        )
                        
                        # 检查位置条件
                        if self._should_merge_fragments(region1, region2, max_distance):
                            if is_partial_match:
                                # 部分匹配：需要分离标签和内容
                                # 1. 创建合并后的标签region
                                merged_label_region = self._merge_two_field_fragments(
                                    region1, region2, merged_text
                                )
                                merged_regions.append(merged_label_region)
                                
                                # 2. 创建剩余内容的region
                                # 使用region2的bbox，但调整x坐标（内容从标签后面开始）
                                x1_2, y1_2, x2_2, y2_2 = region2.bbox
                                # 估算每个字符的宽度
                                char_width = (x2_2 - x1_2) / len(text2) if len(text2) > 0 else 0
                                # 内容的起始x坐标 = 原始x + 碎片长度 * 字符宽度
                                fragment2_len = len(text2) - len(remaining_content)
                                content_x1 = x1_2 + int(char_width * fragment2_len)
                                
                                content_region = TextRegion(
                                    text=remaining_content,
                                    bbox=(content_x1, y1_2, x2_2, y2_2),
                                    confidence=region2.confidence,
                                    font_size=region2.font_size
                                )
                                merged_regions.append(content_region)
                                
                                used_indices.add(i)
                                used_indices.add(j)
                                merged = True
                                merge_count += 1
                                
                                logger.info(
                                    f"合并字段碎片(部分匹配,第{iteration}轮): '{region1.text}' + '{region2.text}' → "
                                    f"标签='{merged_text}' + 内容='{remaining_content[:20]}...'"
                                )
                            else:
                                # 完全匹配：直接合并两个区域
                                merged_region = self._merge_two_field_fragments(
                                    region1, region2, merged_text
                                )
                                merged_regions.append(merged_region)
                                used_indices.add(i)
                                used_indices.add(j)
                                merged = True
                                merge_count += 1
                                
                                logger.info(
                                    f"合并字段碎片(完全匹配,第{iteration}轮): '{region1.text}' + '{region2.text}' → '{merged_text}'"
                                )
                            break
                
                if not merged:
                    merged_regions.append(region1)
            
            total_merge_count += merge_count
            
            # 如果这一轮没有合并任何碎片，停止迭代
            if merge_count == 0:
                logger.debug(f"字段碎片合并 - 第 {iteration} 轮没有合并，停止迭代")
                break
            
            # 更新当前区域列表，准备下一轮
            current_regions = merged_regions
        
        if total_merge_count > 0:
            logger.info(f"字段碎片合并完成：共 {iteration} 轮，合并了 {total_merge_count} 对碎片")
        else:
            logger.info("字段碎片合并完成：没有找到可合并的碎片")
        
        # 应用字段名标准化（处理部分字段名）
        current_regions = self._normalize_field_names(current_regions)
        
        return current_regions
    
    def _normalize_field_names(self, regions: List[TextRegion]) -> List[TextRegion]:
        """标准化字段名（将部分字段名替换为完整字段名）。
        
        这个方法处理OCR无法识别完整字段名的情况，例如：
        - "表人" → "法定代表人"
        - "资本" → "注册资本"
        - "立日期" → "成立日期"
        
        Args:
            regions: 输入的文本区域列表
            
        Returns:
            标准化后的文本区域列表
        """
        if not regions:
            return regions
        
        normalized_count = 0
        
        for region in regions:
            text = region.text.strip()
            
            # 检查是否需要标准化
            if text in self.field_name_normalization:
                normalized_text = self.field_name_normalization[text]
                logger.info(
                    f"字段名标准化: '{text}' → '{normalized_text}'"
                )
                region.text = normalized_text
                normalized_count += 1
        
        if normalized_count > 0:
            logger.info(f"字段名标准化完成：标准化了 {normalized_count} 个字段")
        else:
            logger.info("字段名标准化完成：没有需要标准化的字段")
        
        return regions
    
    def _check_field_fragment_match(
        self, 
        region1: TextRegion, 
        region2: TextRegion
    ) -> Optional[Tuple[str, int, bool, str]]:
        """检查两个区域是否匹配某个字段碎片规则。
        
        支持两种匹配模式：
        1. 完全匹配：text2完全等于fragment2（如"住"+"所"）
        2. 部分匹配：text2以fragment2开头但后面还有内容（如"住"+"所广东省..."）
        
        特殊处理：
        - 对于"经营范"+"围..."这种情况，需要特别小心
        - 如果剩余内容看起来像字段内容（不是字段标签），才进行部分匹配
        
        Args:
            region1: 第一个文本区域
            region2: 第二个文本区域
            
        Returns:
            如果匹配,返回(合并后的完整字段, 最大距离, 是否部分匹配, 剩余内容); 否则返回None
            - 合并后的完整字段: 如"住所"
            - 最大距离: 规则定义的最大水平距离
            - 是否部分匹配: True表示text2后面还有内容需要保留
            - 剩余内容: 如果是部分匹配，返回text2中fragment2之后的内容
        """
        text1 = region1.text.strip()
        text2 = region2.text.strip()
        
        # 检查所有规则
        for fragment1, fragment2, merged_text, max_distance in self.field_fragment_rules:
            # 正向匹配: region1是fragment1, region2是fragment2或以fragment2开头
            if text1 == fragment1:
                # 完全匹配
                if text2 == fragment2:
                    return (merged_text, max_distance, False, "")
                
                # 部分匹配：text2以fragment2开头
                if text2.startswith(fragment2):
                    remaining_content = text2[len(fragment2):]
                    
                    # 特殊检查：如果剩余内容为空或只有空格，视为完全匹配
                    if not remaining_content.strip():
                        return (merged_text, max_distance, False, "")
                    
                    # 特殊检查：防止误合并
                    # 如果fragment2只有1个字符，且剩余内容很长（>10字符），
                    # 可能是误匹配，需要更严格的检查
                    if len(fragment2) == 1 and len(remaining_content) > 10:
                        # 检查剩余内容是否看起来像字段内容（包含常见的内容特征）
                        # 如果剩余内容以常见的字段标签开头，可能是误匹配
                        content_indicators = ["一般项目", "许可项目", "：", "；", "、"]
                        is_likely_content = any(remaining_content.startswith(ind) for ind in content_indicators)
                        
                        if not is_likely_content:
                            # 不太像字段内容，可能是误匹配，跳过
                            logger.debug(
                                f"跳过可疑的部分匹配: '{text1}' + '{text2}' "
                                f"(fragment2='{fragment2}' 只有1字符，剩余内容不像字段内容)"
                            )
                            continue
                    
                    logger.debug(
                        f"部分匹配字段碎片: '{text1}' + '{text2}' → "
                        f"标签='{merged_text}', 剩余内容='{remaining_content[:20]}...'"
                    )
                    return (merged_text, max_distance, True, remaining_content)
            
            # 反向匹配: region1是fragment2, region2是fragment1
            # (虽然通常不会发生,但为了健壮性还是检查一下)
            if text1 == fragment2 and text2 == fragment1:
                return (merged_text, max_distance, False, "")
        
        return None
    
    def _should_merge_fragments(
        self, 
        region1: TextRegion, 
        region2: TextRegion,
        max_horizontal_distance: int
    ) -> bool:
        """判断两个字段碎片是否应该合并。
        
        使用自适应策略:
        1. 在同一行（垂直位置接近）
        2. 水平距离检查（自适应）:
           - 对于单字碎片: 使用字体大小的倍数作为阈值（更宽松）
           - 对于多字碎片: 使用规则定义的固定阈值
        3. 字体大小相近（放宽到30%）
        
        Args:
            region1: 第一个文本区域
            region2: 第二个文本区域
            max_horizontal_distance: 最大允许的水平距离（规则定义的基准值）
            
        Returns:
            True如果应该合并
        """
        x1_1, y1_1, x2_1, y2_1 = region1.bbox
        x1_2, y1_2, x2_2, y2_2 = region2.bbox
        
        text1 = region1.text.strip()
        text2 = region2.text.strip()
        
        # 1. 检查是否在同一行（垂直中心点距离）
        center_y1 = (y1_1 + y2_1) / 2
        center_y2 = (y1_2 + y2_2) / 2
        vertical_distance = abs(center_y1 - center_y2)
        
        # 平均高度
        avg_height = ((y2_1 - y1_1) + (y2_2 - y1_2)) / 2
        
        # 如果垂直距离超过平均高度的50%,认为不在同一行
        if vertical_distance > avg_height * 0.5:
            logger.debug(
                f"Fragment merge check: '{text1}' + '{text2}' - "
                f"不在同一行 (vertical_distance={vertical_distance:.1f}, avg_height={avg_height:.1f})"
            )
            return False
        
        # 2. 自适应水平距离检查
        # 计算两个区域之间的水平间距
        if x1_2 > x2_1:
            # region2在region1右边
            horizontal_distance = x1_2 - x2_1
        elif x1_1 > x2_2:
            # region1在region2右边
            horizontal_distance = x1_1 - x2_2
        else:
            # 重叠
            horizontal_distance = 0
        
        # 自适应距离阈值策略
        # 对于单字碎片（如"住"+"所"），使用字体大小的倍数作为阈值
        # 对于多字碎片，使用规则定义的固定阈值
        is_single_char_pair = (len(text1) <= 2 and len(text2) <= 2)
        
        if is_single_char_pair:
            # 单字碎片: 使用平均字体大小的10倍作为最大距离
            # 这样可以适应不同分辨率和布局的图片
            avg_font_size = (region1.font_size + region2.font_size) / 2
            adaptive_max_distance = avg_font_size * 10
            
            # 但不能无限大，设置一个上限（300px）
            adaptive_max_distance = min(adaptive_max_distance, 300)
            
            logger.debug(
                f"Fragment merge check: '{text1}' + '{text2}' - "
                f"单字碎片，使用自适应阈值 (adaptive_max={adaptive_max_distance:.1f}px, "
                f"avg_font_size={avg_font_size:.1f}px)"
            )
        else:
            # 多字碎片: 使用规则定义的固定阈值
            adaptive_max_distance = max_horizontal_distance
            logger.debug(
                f"Fragment merge check: '{text1}' + '{text2}' - "
                f"多字碎片，使用固定阈值 (max={adaptive_max_distance}px)"
            )
        
        if horizontal_distance > adaptive_max_distance:
            logger.debug(
                f"Fragment merge check: '{text1}' + '{text2}' - "
                f"水平距离过大 (horizontal_distance={horizontal_distance:.1f} > max={adaptive_max_distance:.1f})"
            )
            return False
        
        # 3. 检查字体大小是否相近（放宽到40%）
        if region1.font_size > 0 and region2.font_size > 0:
            font_diff = abs(region1.font_size - region2.font_size) / max(region1.font_size, region2.font_size)
            # 对于字段碎片合并，放宽字体差异阈值到40%
            fragment_font_threshold = 0.4
            
            # 特殊情况：对于某些特定的字段碎片（如"代"+"表人"），进一步放宽到50%
            # 因为OCR可能将单字识别为较小的字体
            special_fragments = [("代", "表人"), ("表人", "代"), ("法", "定"), ("定", "法")]
            is_special = any((text1 == f1 and text2 == f2) or (text1 == f2 and text2 == f1) 
                           for f1, f2 in special_fragments)
            
            if is_special:
                fragment_font_threshold = 0.5  # 放宽到50%
                logger.debug(
                    f"Fragment merge check: '{text1}' + '{text2}' - "
                    f"特殊字段碎片，放宽字体阈值到50%"
                )
            
            if font_diff > fragment_font_threshold:
                logger.debug(
                    f"Fragment merge check: '{text1}' + '{text2}' - "
                    f"字体大小差异过大 (font_diff={font_diff:.1%} > threshold={fragment_font_threshold:.1%})"
                )
                return False
        
        logger.debug(
            f"Fragment merge check: '{text1}' + '{text2}' - "
            f"✓ 可以合并 (h_dist={horizontal_distance:.1f}, v_dist={vertical_distance:.1f}, "
            f"adaptive_max={adaptive_max_distance:.1f})"
        )
        return True
    
    def _merge_two_field_fragments(
        self, 
        region1: TextRegion, 
        region2: TextRegion,
        merged_text: str
    ) -> TextRegion:
        """合并两个字段碎片为一个完整字段。
        
        Args:
            region1: 第一个文本区域
            region2: 第二个文本区域
            merged_text: 合并后的完整文本
            
        Returns:
            合并后的文本区域
        """
        x1_1, y1_1, x2_1, y2_1 = region1.bbox
        x1_2, y1_2, x2_2, y2_2 = region2.bbox
        
        # 合并边界框（取最小和最大值）
        merged_bbox = (
            min(x1_1, x1_2),
            min(y1_1, y1_2),
            max(x2_1, x2_2),
            max(y2_1, y2_2)
        )
        
        # 使用较高的置信度
        merged_confidence = max(region1.confidence, region2.confidence)
        
        # 使用较大的字体大小
        merged_font_size = max(region1.font_size, region2.font_size)
        
        # 使用第一个区域的角度（通常字段碎片角度相同）
        merged_angle = region1.angle
        
        return TextRegion(
            bbox=merged_bbox,
            text=merged_text,
            confidence=merged_confidence,
            font_size=merged_font_size,
            angle=merged_angle,
            is_field_label=True  # 标记为字段标签，避免后续合并
        )
    
    def _detect_and_split_vertical_labels(
        self, 
        regions: List[TextRegion],
        image: np.ndarray
    ) -> List[TextRegion]:
        """检测并拆分垂直排列的字段标签。
        
        这个方法处理那些在原图中垂直排列的字段名（如营业执照左侧的字段），
        它们被OCR识别为一个文本块（可能是横向或纵向）。
        
        检测策略：
        1. 宽高比检测：height/width > threshold（适用于明显的垂直文本）
        2. 内容检测：文本包含多个字段名的碎片（适用于OCR错误识别的情况）
        3. 位置检测：文本在图片左侧（X < 图片宽度的30%）
        
        处理流程：
        1. 检测符合条件的文本区域
        2. 将文本按字符拆分成单字列表
        3. 使用字段名词典进行贪婪匹配，识别完整字段名
        4. 为每个字段创建独立的TextRegion
        
        Args:
            regions: 输入的文本区域列表
            image: 原始图像（用于验证）
            
        Returns:
            处理后的文本区域列表
        """
        # 检查是否启用垂直标签检测
        vertical_label_enabled = self.config.get('ocr.vertical_label_detection.enabled', True)
        
        if not vertical_label_enabled:
            logger.info("垂直标签检测已禁用")
            return regions
        
        if not regions:
            return regions
        
        # 获取配置参数
        aspect_ratio_threshold = self.config.get('ocr.vertical_label_detection.aspect_ratio_threshold', 1.3)
        field_dictionary = self.config.get('ocr.vertical_label_detection.field_dictionary', [
            # 按长度排序（长的在前，用于贪婪匹配）
            "统一社会信用代码",
            "法定代表人",
            "注册资本",
            "成立日期",
            "营业期限",
            "经营范围",
            "登记机关",
            "名称",
            "类型",
            "住所",
            "名",
            "称",
            "类",
            "型",
            "住",
            "所",
        ])
        
        # 获取图片尺寸
        image_height, image_width = image.shape[:2]
        left_region_threshold = image_width * 0.3  # 左侧区域阈值（30%）
        
        logger.info(f"开始垂直标签检测，输入 {len(regions)} 个区域")
        
        # 调试：打印所有区域的文本和位置
        logger.info("=" * 60)
        logger.info("垂直标签检测 - 输入区域列表：")
        for i, r in enumerate(regions):
            x1, y1, x2, y2 = r.bbox
            width = x2 - x1
            height = y2 - y1
            ratio = height / width if width > 0 else 0
            logger.info(
                f"  区域 #{i+1}: '{r.text}' "
                f"(w={width}, h={height}, ratio={ratio:.2f}, x={x1})"
            )
        logger.info("=" * 60)
        
        # 调试：打印所有区域的文本和位置（只显示长度>=3的）
        for i, r in enumerate(regions):
            if len(r.text) >= 3:
                x1, y1, x2, y2 = r.bbox
                width = x2 - x1
                height = y2 - y1
                ratio = height / width if width > 0 else 0
                logger.debug(
                    f"  区域 #{i+1}: '{r.text}' "
                    f"(w={width}, h={height}, ratio={ratio:.2f}, x={x1})"
                )
        
        processed_regions = []
        split_count = 0
        
        for region in regions:
            x1, y1, x2, y2 = region.bbox
            width = x2 - x1
            height = y2 - y1
            text = region.text.strip()
            
            # 避免除零错误
            if width == 0:
                processed_regions.append(region)
                continue
            
            aspect_ratio = height / width
            
            # 策略1：宽高比检测（适用于明显的垂直文本）
            # 注意：有些垂直标签被OCR识别成横向文本块（宽度>高度），所以也要检查这种情况
            is_vertical_by_ratio = aspect_ratio > aspect_ratio_threshold
            
            # 策略2：内容检测（检查是否包含多个字段名的碎片）
            is_vertical_by_content = self._contains_multiple_field_fragments(text, field_dictionary)
            
            # 策略3：位置检测（是否在图片左侧）
            is_in_left_region = x1 < left_region_threshold
            
            # 调试日志
            if len(text) >= 3:
                logger.debug(
                    f"检查区域 '{text}': "
                    f"ratio={aspect_ratio:.2f}, "
                    f"x={x1}, "
                    f"is_vertical_by_ratio={is_vertical_by_ratio}, "
                    f"is_vertical_by_content={is_vertical_by_content}, "
                    f"is_in_left_region={is_in_left_region}"
                )
            
            # 综合判断：满足任一条件即可
            should_split = False
            reason = ""
            
            if is_vertical_by_ratio:
                should_split = True
                reason = f"宽高比检测 (ratio={aspect_ratio:.2f} > {aspect_ratio_threshold})"
            elif is_vertical_by_content and is_in_left_region:
                should_split = True
                reason = f"内容+位置检测 (left={x1}px < {left_region_threshold:.0f}px, 包含多个字段碎片)"
            
            if should_split:
                logger.info(
                    f"检测到垂直排列标签: '{text}' "
                    f"(width={width}, height={height}, ratio={aspect_ratio:.2f}, x={x1}, 原因={reason})"
                )
                
                # 拆分并重组字段
                split_regions = self._split_vertical_label(region, field_dictionary, image)
                
                if split_regions and len(split_regions) > 1:
                    # 成功拆分
                    processed_regions.extend(split_regions)
                    split_count += 1
                    logger.info(
                        f"成功拆分垂直标签: '{text}' → "
                        f"{[r.text for r in split_regions]}"
                    )
                else:
                    # 拆分失败或只有一个字段，保持原样
                    processed_regions.append(region)
            else:
                # 不是垂直排列的标签，保持原样
                processed_regions.append(region)
        
        if split_count > 0:
            logger.info(f"垂直标签检测完成：拆分了 {split_count} 个垂直标签")
        else:
            logger.info("垂直标签检测完成：没有找到需要拆分的垂直标签")
        
        return processed_regions
    
    def _contains_multiple_field_fragments(
        self,
        text: str,
        field_dictionary: List[str]
    ) -> bool:
        """检查文本是否包含多个字段名的碎片。
        
        例如："类型表人" 包含 "类型" 和 "表人"（"法定代表人"的一部分）
        
        Args:
            text: 要检查的文本
            field_dictionary: 字段名词典
            
        Returns:
            True 如果包含多个字段碎片
        """
        if len(text) < 3:
            return False
        
        # 统计匹配到的字段数量
        matched_fields = []
        
        for field_name in field_dictionary:
            # 检查字段名是否在文本中（完全匹配或部分匹配）
            if field_name in text:
                matched_fields.append(field_name)
                logger.debug(f"  文本 '{text}' 完全匹配字段: '{field_name}'")
            else:
                # 检查部分匹配（至少2个字符）
                if len(field_name) >= 2:
                    for i in range(len(text) - 1):
                        substring = text[i:i+2]
                        if substring in field_name:
                            matched_fields.append(field_name)
                            logger.debug(f"  文本 '{text}' 部分匹配字段: '{field_name}' (子串='{substring}')")
                            break
        
        # 如果匹配到2个或更多不同的字段，认为是垂直标签
        unique_matched = list(set(matched_fields))
        
        if len(unique_matched) >= 2:
            logger.debug(
                f"文本 '{text}' 包含多个字段碎片: {unique_matched}"
            )
            return True
        else:
            if len(text) >= 3:
                logger.debug(f"  文本 '{text}' 只匹配到 {len(unique_matched)} 个字段: {unique_matched}")
        
        return False
    
    def _split_vertical_label(
        self,
        region: TextRegion,
        field_dictionary: List[str],
        image: np.ndarray
    ) -> List[TextRegion]:
        """拆分单个垂直排列的字段标签。
        
        使用贪婪匹配算法：
        1. 从第一个字符开始
        2. 尝试匹配最长的字段名
        3. 如果匹配成功，创建TextRegion并继续处理剩余字符
        4. 如果匹配失败，尝试匹配较短的字段名
        5. 如果所有匹配都失败，将当前字符作为单字字段处理
        
        Args:
            region: 要拆分的文本区域
            field_dictionary: 字段名词典（已按长度排序）
            image: 原始图像
            
        Returns:
            拆分后的文本区域列表
        """
        text = region.text.strip()
        x1, y1, x2, y2 = region.bbox
        width = x2 - x1
        height = y2 - y1
        
        # 将文本拆分成字符列表
        chars = list(text)
        
        if not chars:
            return [region]
        
        # 计算平均字符高度
        char_height = height / len(chars)
        
        # 贪婪匹配字段名
        split_regions = []
        i = 0
        
        while i < len(chars):
            matched = False
            
            # 尝试匹配最长的字段名（从当前位置开始）
            for field_name in field_dictionary:
                field_len = len(field_name)
                
                # 检查是否有足够的字符
                if i + field_len > len(chars):
                    continue
                
                # 提取候选字符串
                candidate = ''.join(chars[i:i+field_len])
                
                # 检查是否匹配
                if candidate == field_name:
                    # 匹配成功，创建TextRegion
                    field_y1 = int(y1 + i * char_height)
                    field_y2 = int(y1 + (i + field_len) * char_height)
                    
                    # 确保边界在图像范围内
                    image_height, image_width = image.shape[:2]
                    field_y1 = max(0, min(field_y1, image_height))
                    field_y2 = max(0, min(field_y2, image_height))
                    
                    field_region = TextRegion(
                        bbox=(x1, field_y1, x2, field_y2),
                        text=field_name,
                        confidence=region.confidence,
                        font_size=region.font_size,
                        angle=region.angle
                    )
                    
                    split_regions.append(field_region)
                    
                    logger.debug(
                        f"匹配字段: '{field_name}' at position {i} "
                        f"(bbox: ({x1}, {field_y1}, {x2}, {field_y2}))"
                    )
                    
                    # 移动到下一个位置
                    i += field_len
                    matched = True
                    break
            
            if not matched:
                # 没有匹配到任何字段名，将当前字符作为单字字段
                single_char = chars[i]
                field_y1 = int(y1 + i * char_height)
                field_y2 = int(y1 + (i + 1) * char_height)
                
                # 确保边界在图像范围内
                image_height, image_width = image.shape[:2]
                field_y1 = max(0, min(field_y1, image_height))
                field_y2 = max(0, min(field_y2, image_height))
                
                field_region = TextRegion(
                    bbox=(x1, field_y1, x2, field_y2),
                    text=single_char,
                    confidence=region.confidence,
                    font_size=region.font_size,
                    angle=region.angle
                )
                
                split_regions.append(field_region)
                
                logger.debug(
                    f"单字字段: '{single_char}' at position {i} "
                    f"(bbox: ({x1}, {field_y1}, {x2}, {field_y2}))"
                )
                
                i += 1
        
        return split_regions
    
    def process_image(self, image: np.ndarray) -> List[TextRegion]:
        """Process an image through the complete OCR pipeline.
        
        This method performs:
        1. Text detection
        2. Region splitting (split incorrectly merged regions)
        3. Field fragment merging (merge field fragments like "名"+"称"→"名称")
        4. Pre-merge filtering (remove watermarks before merging)
        5. Region merging
        6. Post-merge filtering (confidence and area)
        
        Args:
            image: Input image as numpy array
            
        Returns:
            List of processed TextRegion objects
            
        Raises:
            OCRError: If OCR processing fails
        """
        # Detect text regions
        regions = self.detect_text(image)
        
        if not regions:
            return []
        
        # Split incorrectly merged regions (e.g., "称佛山..." → "称" | "佛山...")
        regions = self._split_merged_regions(regions, image)
        
        # Merge field fragments (e.g., "名" + "称" → "名称")
        regions = self._merge_field_fragments(regions)
        
        # Pre-merge filtering: Remove watermarks BEFORE merging
        # This prevents watermarks from being merged with legitimate text
        regions = self._filter_watermarks(regions)
        
        # Merge adjacent regions
        regions = self.merge_regions(regions)
        
        # Post-merge filtering: Apply confidence and area filters
        # Pass image for edge noise detection
        regions = self.filter_regions(regions, image)
        
        logger.info(f"OCR pipeline complete: {len(regions)} regions after processing")
        return regions
    
    def get_cache_stats(self) -> Optional[dict]:
        """Get OCR cache statistics.
        
        Returns:
            Dictionary containing cache statistics, or None if caching is disabled
            
        _Requirements: 14.2_
        """
        if self.cache is not None:
            return self.cache.get_stats()
        return None
    
    def clear_cache(self) -> None:
        """Clear the OCR cache.
        
        _Requirements: 14.2_
        """
        if self.cache is not None:
            self.cache.clear()
            logger.info("OCR cache cleared")
    
    def reset_cache_stats(self) -> None:
        """Reset OCR cache statistics.
        
        _Requirements: 14.2_
        """
        if self.cache is not None:
            self.cache.reset_stats()
            logger.debug("OCR cache statistics reset")
