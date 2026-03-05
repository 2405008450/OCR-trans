# -*- coding: utf-8 -*-
"""Translation pipeline for image translation system.

This module provides the TranslationPipeline class that orchestrates
the complete image translation workflow, coordinating all components
from OCR detection through quality validation.

_Requirements: 10.1, 10.2, 10.4, 10.5, 13.2, 14.1, 14.3, 14.4_
"""

import gc
import logging
import os
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Generator, Iterator, List, Optional, Tuple

import numpy as np
import cv2

from src.config import ConfigManager
from src.ocr import OCREngine
from src.icon import IconDetector
from src.translation import TranslationService
from src.background import BackgroundSampler
from src.rendering import TextRenderer
from src.rendering.seal_text_handler import SealTextHandler
from src.image import ImageProcessor
from src.quality import QualityValidator
from src.models import TextRegion, TranslationResult, QualityReport, QualityLevel
from src.overlap import PositionAdjuster
from src.paragraph_merge import FieldAwareParagraphMerger
from src.exceptions import (
    PipelineError,
    OCRError,
    TranslationError,
    ImageLoadError,
    ImageSaveError
)


logger = logging.getLogger(__name__)


class TranslationPipeline:
    """Orchestrates the complete image translation workflow.
    
    The TranslationPipeline coordinates all components of the image
    translation system:
    - OCR Engine for text detection
    - Icon Detector for filtering non-text elements
    - Translation Service for text translation
    - Background Sampler for background color extraction
    - Text Renderer for rendering translated text
    - Image Processor for image I/O
    - Quality Validator for output validation
    
    The pipeline supports:
    - Single image translation
    - Batch image translation with parallel processing
    - Memory-efficient generator-based batch processing
    - Partial failure tolerance
    - Original image protection
    
    Attributes:
        config: Configuration manager instance
        ocr: OCR engine for text detection
        icon_detector: Icon detector for filtering
        translator: Translation service
        background_sampler: Background color sampler
        text_renderer: Text renderer
        image_processor: Image processor
        validator: Quality validator
        max_workers: Maximum number of parallel workers
        parallel_enabled: Whether parallel processing is enabled
        memory_efficient: Whether to use memory-efficient processing
    
    _Requirements: 10.1, 14.1, 14.3_
    """
    
    # Default maximum workers for parallel processing
    DEFAULT_MAX_WORKERS = 4
    
    # Minimum workers to prevent resource starvation
    MIN_WORKERS = 1
    
    # Maximum workers to prevent excessive memory usage
    MAX_WORKERS_LIMIT = 8
    
    def __init__(self, config: ConfigManager):
        """Initialize the translation pipeline.
        
        Creates instances of all required components and configures
        the pipeline based on the provided configuration.
        
        Args:
            config: Configuration manager instance
            
        _Requirements: 10.1, 14.1, 14.3_
        """
        self.config = config
        
        # Set up logging
        self._setup_logging()
        
        logger.info("Initializing TranslationPipeline")
        
        # Initialize all components
        self.ocr = OCREngine(config)
        self.icon_detector = IconDetector(config)
        self.translator = TranslationService(config)
        self.background_sampler = BackgroundSampler(config)
        self.text_renderer = TextRenderer(config)
        self.seal_text_handler = SealTextHandler(config)
        self.image_processor = ImageProcessor(config)
        self.validator = QualityValidator(config)
        
        # Initialize field-aware paragraph merger
        self.paragraph_merger = FieldAwareParagraphMerger(config)
        
        # Initialize position adjuster for overlap prevention
        overlap_step_size = config.get('rendering.overlap_prevention.step_size', 5.0)
        overlap_max_iterations = config.get('rendering.overlap_prevention.max_iterations', 1000)
        self.position_adjuster = PositionAdjuster(
            step_size=overlap_step_size,
            max_iterations=overlap_max_iterations
        )
        
        # Pipeline configuration with memory optimization
        configured_workers = config.get('performance.max_workers', self.DEFAULT_MAX_WORKERS)
        self.max_workers = max(
            self.MIN_WORKERS, 
            min(configured_workers, self.MAX_WORKERS_LIMIT)
        )
        self.parallel_enabled = config.get('performance.parallel_translation', True)
        self.memory_efficient = config.get('performance.memory_efficient', True)
        
        logger.info(
            f"TranslationPipeline initialized successfully "
            f"(max_workers={self.max_workers}, memory_efficient={self.memory_efficient})"
        )
    
    def _setup_logging(self) -> None:
        """Set up logging for the pipeline.
        
        Configures logging based on configuration settings.
        """
        log_level = self.config.get('logging.level', 'INFO')
        
        # Configure logger if not already configured
        if not logger.handlers:
            handler = logging.StreamHandler()
            handler.setFormatter(
                logging.Formatter(
                    '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
                )
            )
            logger.addHandler(handler)
        
        logger.setLevel(getattr(logging, log_level.upper(), logging.INFO))
    
    def _release_memory(self, *arrays: Optional[np.ndarray]) -> None:
        """Release memory by deleting large arrays and triggering garbage collection.
        
        This method helps manage memory usage when processing large images
        by explicitly releasing numpy arrays and triggering garbage collection.
        
        Args:
            *arrays: Variable number of numpy arrays to release
            
        _Requirements: 14.1_
        """
        for arr in arrays:
            if arr is not None:
                del arr
        
        # Trigger garbage collection to free memory
        gc.collect()
        logger.debug("Memory released and garbage collection triggered")


    def translate_image(
        self, 
        input_path: str, 
        output_path: str,
        source_lang: str = "zh",
        target_lang: str = "en"
    ) -> QualityReport:
        """Translate text in an image from source to target language.
        
        Performs the complete translation workflow:
        1. Load the input image
        2. Detect text regions using OCR
        3. Filter out icon regions
        4. Translate each text region
        5. Sample background colors
        6. Render translated text
        7. Validate output quality
        8. Save the output image
        
        Args:
            input_path: Path to the input image
            output_path: Path where the translated image should be saved
            source_lang: Source language code (default: "zh")
            target_lang: Target language code (default: "en")
            
        Returns:
            QualityReport containing translation quality metrics
            
        Raises:
            PipelineError: If a critical error occurs during processing
            
        _Requirements: 1.1, 2.2, 3.1, 4.1, 5.2, 9.1_
        """
        logger.info(f"Starting translation: {input_path} -> {output_path}")
        
        # Track processing state for error recovery
        original_image = None
        output_image = None
        regions = []
        text_regions = []
        translation_results = []
        
        try:
            # Step 1: Load the input image
            logger.debug("Step 1: Loading input image")
            original_image = self.image_processor.load_image(input_path)
            
            # Detect image orientation
            orientation = self.image_processor.detect_orientation(original_image)
            logger.info(f"Image orientation: {orientation}")
            
            # Save orientation for later use
            self.current_orientation = orientation
            
            # 如果配置是 auto 模式，根据图片方向自动切换配置
            if self.config.is_auto_mode():
                # 映射 orientation 到配置方向
                # landscape -> horizontal, portrait -> vertical
                config_orientation = 'horizontal' if orientation == 'landscape' else 'vertical'
                self.config.switch_orientation(config_orientation)
                logger.info(f"🔄 Auto mode: switched to {config_orientation} configuration")
            
            # Apply orientation-specific settings
            if orientation == 'portrait':
                self._apply_portrait_settings()
            else:
                # Landscape mode
                self._apply_landscape_settings()
            
            # Create a working copy to protect the original
            output_image = original_image.copy()
            
            # Step 2: Detect text regions using OCR
            logger.debug("Step 2: Detecting text regions")
            regions = self.ocr.process_image(original_image)
            
            if not regions:
                logger.info("No text regions detected in image")
                # Create empty quality report
                report = QualityReport(
                    translation_coverage=1.0,
                    total_regions=0,
                    translated_regions=0,
                    failed_regions=[],
                    has_artifacts=False,
                    artifact_locations=[],
                    overall_quality=QualityLevel.EXCELLENT
                )
                # Save the original image as output
                self.image_processor.save_image(output_image, output_path)
                return report
            
            logger.info(f"Detected {len(regions)} text regions")
            
            # Step 2.5: Merge continuous regions based on line spacing
            logger.debug("Step 2.5: Merging continuous regions based on line spacing")
            regions = self._merge_continuous_regions(regions)
            
            # Step 2.6: Separate watermark from mixed text regions
            logger.debug("Step 2.6: Separating watermark from mixed text regions")
            regions = self._separate_watermark_regions(regions, original_image)
            
            # Step 3: Filter out icon regions
            logger.debug("Step 3: Filtering icon regions")
            # 保存图标区域（包括二维码）以便后续使用
            icon_regions = self.icon_detector.get_icon_regions(regions, original_image)
            text_regions = self.icon_detector.filter_icons(regions, original_image)
            
            if not text_regions:
                logger.info("All regions were filtered as icons")
                report = QualityReport(
                    translation_coverage=1.0,
                    total_regions=len(regions),
                    translated_regions=0,
                    failed_regions=[],
                    has_artifacts=False,
                    artifact_locations=[],
                    overall_quality=QualityLevel.EXCELLENT
                )
                self.image_processor.save_image(output_image, output_path)
                return report
            
            logger.info(f"Processing {len(text_regions)} text regions after icon filtering")
            
            # Step 3.5: Filter out regions that overlap with seals
            logger.debug("Step 3.5: Filtering seal regions")
            text_regions = self._filter_seal_regions(original_image, text_regions)
            
            if not text_regions:
                logger.info("All regions were filtered as seal-overlapping")
                report = QualityReport(
                    translation_coverage=1.0,
                    total_regions=len(regions),
                    translated_regions=0,
                    failed_regions=[],
                    has_artifacts=False,
                    artifact_locations=[],
                    overall_quality=QualityLevel.EXCELLENT
                )
                self.image_processor.save_image(output_image, output_path)
                return report
            
            logger.info(f"Processing {len(text_regions)} text regions after seal filtering")
            
            # Step 3.6: 标准化特定字段的边界框（在翻译之前）
            logger.debug("Step 3.6: Normalizing field bounding boxes")
            text_regions = self._normalize_field_regions(text_regions)
            
            # Step 4: Translate each text region
            logger.debug("Step 4: Translating text regions")
            
            # Step 4.1: 合并被同一个印章覆盖的文本区域
            logger.debug("Step 4.1: Merging seal-covered text regions")
            text_regions = self._merge_seal_covered_regions(text_regions)
            
            # 过滤掉忽略关键词（纯水印文本），但保留字段标签用于单独翻译
            # 注意：混合文本（水印+有效内容）已经在Step 2.6中被分离处理
            ignore_keywords = self.config.get('ocr.ignore_keywords', [])
            texts_to_translate = []
            filtered_regions = []
            field_label_regions = []
            field_label_texts = []
            
            # 字段标签的标准翻译映射（中文 -> 英文）
            # 注意：这里只包含左侧的字段标签，不包括"统一社会信用代码"等独立标题
            field_label_translations = {
                "法定代表人": "Legal Representative",
                "注册资本": "Registered Capital",
                "成立日期": "Date of Establishment",
                "营业期限": "Business Term",
                "经营范围": "Business Scope",
                "登记机关": "Registration Authority",
                "名称": "Name",
                "类型": "Type",
                "住所": "Domicile",
                "商事主体类型": "Type of Business Entity",
            }
            
            for region in text_regions:
                # 检查是否是字段标签（在_split_field_labels中标记的）
                is_field_label = getattr(region, 'is_field_label', False)
                
                if is_field_label:
                    # 字段标签单独处理，使用标准翻译
                    field_label_regions.append(region)
                    field_label_texts.append(region.text)
                    logger.info(f"识别字段标签: '{region.text}'")
                    continue
                
                # 水印过滤：检查是否包含水印关键词
                should_ignore = False
                if ignore_keywords:
                    import re
                    
                    # 只过滤真正包含水印关键词的文本
                    for keyword in ignore_keywords:
                        if keyword.lower() in region.text.lower():
                            should_ignore = True
                            logger.info(f"忽略水印文本: '{region.text}'")
                            break
                
                if not should_ignore:
                    texts_to_translate.append(region.text)
                    filtered_regions.append(region)
            
            # 记录过滤统计
            if field_label_regions:
                logger.info(f"字段标签识别: 找到 {len(field_label_regions)} 个字段标签")
            
            if len(filtered_regions) < len(text_regions):
                total_filtered = len(text_regions) - len(filtered_regions) - len(field_label_regions)
                logger.info(f"总过滤统计: {len(text_regions)} 个区域 -> {len(filtered_regions)} 个内容区域 + {len(field_label_regions)} 个标签区域 (跳过 {total_filtered} 个)")
            
            # 翻译字段内容（不包括字段标签）
            translation_results = self.translator.translate_batch(
                texts_to_translate, 
                source_lang, 
                target_lang
            )
            
            # 为字段标签创建翻译结果（使用标准翻译）
            from src.models import TranslationResult
            field_label_results = []
            for label_text in field_label_texts:
                # 使用标准翻译映射
                translated_label = field_label_translations.get(label_text, label_text)
                field_label_results.append(
                    TranslationResult(
                        source_text=label_text,
                        translated_text=translated_label,
                        confidence=1.0,
                        success=True
                    )
                )
                logger.info(f"字段标签翻译: '{label_text}' -> '{translated_label}'")
            
            # 合并字段标签和内容的区域和翻译结果
            # 注意：需要保持原始顺序，以便后续处理
            all_regions = []
            all_results = []
            label_idx = 0
            content_idx = 0
            
            for region in text_regions:
                is_field_label = getattr(region, 'is_field_label', False)
                
                if is_field_label:
                    # 这是字段标签
                    if label_idx < len(field_label_regions):
                        all_regions.append(field_label_regions[label_idx])
                        all_results.append(field_label_results[label_idx])
                        label_idx += 1
                else:
                    # 这是内容区域（或被过滤的区域）
                    if content_idx < len(filtered_regions):
                        all_regions.append(filtered_regions[content_idx])
                        all_results.append(translation_results[content_idx])
                        content_idx += 1
            
            # 更新变量，使用合并后的列表
            filtered_regions = all_regions
            translation_results = all_results
            
            # Debug: 检查经营范围的内容是否有标志
            for i, region in enumerate(filtered_regions):
                if '品牌设计' in region.text or '布国内各类广告' in region.text:
                    print(f"[FILTERED_REGIONS DEBUG] Region {i}: '{region.text[:50]}...'")
                    print(f"[FILTERED_REGIONS DEBUG]   is_field_content: {getattr(region, 'is_field_content', False)}")
                    print(f"[FILTERED_REGIONS DEBUG]   is_paragraph_merged: {getattr(region, 'is_paragraph_merged', False)}")
                    print(f"[FILTERED_REGIONS DEBUG]   belongs_to_field: {getattr(region, 'belongs_to_field', None)}")
            
            # Step 5 & 6: Process each region (background sampling and text rendering)
            logger.debug("Step 5 & 6: Processing regions (background + rendering)")
            output_image, translation_results = self._process_regions(
                output_image,
                filtered_regions,  # 使用过滤后的区域
                translation_results,
                icon_regions  # 传递图标区域（包括二维码）
            )
            
            # Step 7: Validate output quality
            logger.debug("Step 7: Validating output quality")
            report = self.validator.validate(
                original_image,
                output_image,
                text_regions,
                translation_results
            )
            
            # Step 8: Save the output image
            logger.debug("Step 8: Saving output image")
            self.image_processor.save_image(output_image, output_path)
            
            logger.info(
                f"Translation complete: coverage={report.translation_coverage:.1%}, "
                f"quality={report.overall_quality.value}"
            )
            
            return report
            
        except ImageLoadError as e:
            logger.error(f"Failed to load image: {e}")
            raise PipelineError(f"Failed to load image: {e}")
        except ImageSaveError as e:
            logger.error(f"Failed to save image: {e}")
            raise PipelineError(f"Failed to save image: {e}")
        except OCRError as e:
            logger.error(f"OCR processing failed: {e}")
            raise PipelineError(f"OCR processing failed: {e}")
        except Exception as e:
            logger.error(f"Unexpected error during translation: {e}")
            raise PipelineError(f"Translation pipeline failed: {e}")
    
    def _apply_portrait_settings(self) -> None:
        """应用竖版图片的特殊配置。
        
        竖版图片（高 > 宽）通常是营业执照等文档的扫描件，
        需要使用不同的 OCR 和渲染参数来获得更好的效果。
        
        注意：在双配置架构下，配置已经通过 switch_orientation() 切换到竖版，
        所以直接从当前配置读取即可，不需要再找 portrait_mode 路径。
        """
        logger.info("=" * 60)
        logger.info("PORTRAIT MODE ACTIVATED - Applying vertical configuration")
        logger.info("=" * 60)
        
        # 直接从当前配置读取（配置已经切换到 vertical_config）
        # 应用竖版 OCR 配置
        logger.info("  [OCR Configuration]")
        logger.info(f"    - confidence_threshold = {self.config.get('ocr.confidence_threshold')}")
        logger.info(f"    - min_text_area = {self.config.get('ocr.min_text_area')}")
        logger.info(f"    - merge_threshold_horizontal = {self.config.get('ocr.merge_threshold_horizontal')}")
        logger.info(f"    - merge_threshold_vertical = {self.config.get('ocr.merge_threshold_vertical')}")
        logger.info(f"    - smart_merge.enabled = {self.config.get('ocr.smart_merge.enabled', False)}")
        logger.info(f"    - vertical_single_char_merge.enabled = {self.config.get('ocr.smart_merge.vertical_single_char_merge.enabled', False)}")
        logger.info(f"    - paragraph_merge.enabled = {self.config.get('ocr.paragraph_merge.enabled', False)}")
        logger.info(f"    - noise_filter.enabled = {self.config.get('ocr.noise_filter.enabled', False)}")
        
        # 应用竖版渲染配置
        logger.info("  [Rendering Configuration]")
        logger.info(f"    - font_scale = {self.config.get('rendering.font_scale')}")
        logger.info(f"    - overlap_prevention.enabled = {self.config.get('rendering.overlap_prevention.enabled', False)}")
        logger.info(f"    - normalize_fields.enabled = {self.config.get('rendering.normalize_fields.enabled', False)}")
        logger.info(f"    - seal_handling.enabled = {self.config.get('rendering.seal_handling.enabled', False)}")
        logger.info(f"    - precise_erase_mode = {self.config.get('rendering.precise_erase_mode', False)}")
        
        # 设置精确擦除模式标志
        precise_erase = self.config.get('rendering.precise_erase_mode', False)
        if precise_erase:
            self.background_sampler._portrait_precise_erase = True
            logger.info(f"    - Enabled precise erase mode for portrait")
        
        logger.info("=" * 60)
        logger.info("Portrait configuration applied successfully")
        logger.info("=" * 60)
        logger.info("Portrait-specific settings applied successfully")
        logger.info("=" * 60)
    
    def _apply_landscape_settings(self) -> None:
        """应用横版图片的配置（基于双配置架构）。
        
        横版图片（宽 > 高）通常是表格、证书等文档，
        需要使用简单的处理逻辑和不同的参数来获得更好的效果。
        
        双配置架构：配置已经根据 document_orientation 自动加载，
        这里只需要记录日志，不需要再次切换配置。
        """
        logger.info("=" * 60)
        logger.info("LANDSCAPE MODE ACTIVATED - Using horizontal configuration")
        logger.info(f"Configuration orientation: {self.config.get_orientation()}")
        logger.info("=" * 60)
        
        # 双配置架构：配置已经在 ConfigManager 中根据 document_orientation 加载
        # 这里只需要记录关键配置参数，不需要再次设置
        
        # 记录横版 OCR 配置
        logger.info("  [OCR Configuration]")
        logger.info(f"    - confidence_threshold = {self.config.get('ocr.confidence_threshold')}")
        logger.info(f"    - min_text_area = {self.config.get('ocr.min_text_area')}")
        logger.info(f"    - merge_threshold_horizontal = {self.config.get('ocr.merge_threshold_horizontal')}")
        logger.info(f"    - merge_threshold_vertical = {self.config.get('ocr.merge_threshold_vertical')}")
        logger.info(f"    - smart_merge.enabled = {self.config.get('ocr.smart_merge.enabled', False)}")
        logger.info(f"    - vertical_single_char_merge.enabled = {self.config.get('ocr.smart_merge.vertical_single_char_merge.enabled', False)}")
        logger.info(f"    - paragraph_merge.enabled = {self.config.get('ocr.paragraph_merge.enabled', False)}")
        
        # 记录横版渲染配置
        logger.info("  [Rendering Configuration]")
        logger.info(f"    - font_scale = {self.config.get('rendering.font_scale')}")
        logger.info(f"    - overlap_prevention.enabled = {self.config.get('rendering.overlap_prevention.enabled', False)}")
        logger.info(f"    - normalize_fields.enabled = {self.config.get('rendering.normalize_fields.enabled', False)}")
        logger.info(f"    - seal_handling.enabled = {self.config.get('rendering.seal_handling.enabled', False)}")
        logger.info(f"    - align_field_contents.enabled = {self.config.get('rendering.align_field_contents.enabled', False)}")
        logger.info(f"    - align_field_contents.y_threshold = {self.config.get('rendering.align_field_contents.y_threshold', 15)}px")
        
        logger.info("=" * 60)
        logger.info("Landscape configuration loaded from horizontal_config block")
        logger.info("=" * 60)
    
    def _process_regions(
        self,
        image: np.ndarray,
        regions: List[TextRegion],
        results: List[TranslationResult],
        icon_regions: List[TextRegion] = None
    ) -> Tuple[np.ndarray, List[TranslationResult]]:
        """Process text regions by sampling background and rendering text.
        
        For each successfully translated region:
        1. Sample the background color
        2. Fill the region with background (using inpainting if enabled)
        3. Render the translated text
        
        Args:
            image: Working image to modify
            regions: List of text regions to process (已经标准化过的)
            results: List of translation results
            icon_regions: List of icon regions (包括二维码等)
            
        Returns:
            Tuple of (processed_image, updated_results)
        """
        output_image = image.copy()
        
        # ========== 第零阶段：设置动态换行阈值 ==========
        # 在渲染之前，根据所有字段内容的宽度计算动态阈值
        logger.debug("Step 0: Setting dynamic wrap threshold for field content")
        self.text_renderer.set_field_content_wrap_threshold_from_regions(regions)
        
        # ========== 第一阶段：位置调整决策（在抹除任何内容之前） ==========
        
        # 获取所有标准化字段框作为障碍物
        normalized_field_boxes = self._get_normalized_field_boxes(regions)
        # 获取标准化字段的原始文本（中文），用于判断当前区域是否是标准化字段
        normalized_field_texts = self._get_normalized_field_texts()
        image_height, image_width = output_image.shape[:2]
        
        # 预先计算所有区域的调整后位置
        adjusted_regions = []
        
        logger.info(f"[DEBUG] 开始计算调整后位置，共 {len(regions)} 个区域")
        
        for i, (region, result) in enumerate(zip(regions, results)):
            # 对于所有区域（包括失败的），都要添加到adjusted_regions列表中
            # 这样才能保证adjusted_regions和regions长度一致
            if not result.success or not result.translated_text.strip():
                # 失败或空文本，保持原region
                adjusted_regions.append(region)
                # 注意：这里不能continue，要继续处理下一个区域
            else:
                # 检查当前区域是否是标准化字段本身
                is_normalized_field = region.text in normalized_field_texts
                
                # 检查当前区域是否是字段内容（已经在_normalize_field_regions中调整过y坐标）
                is_field_content = getattr(region, 'is_field_content', False)
                
                # 调试日志
                text_preview = region.text[:30] + '...' if len(region.text) > 30 else region.text
                logger.info(
                    f"[DEBUG] Region {i}: '{text_preview}', "
                    f"is_normalized_field={is_normalized_field}, is_field_content={is_field_content}"
                )
                
                if is_normalized_field or is_field_content:
                    # 标准化字段本身或字段内容，不调整位置
                    adjusted_regions.append(region)
                    if is_normalized_field:
                        logger.debug(f"Region '{region.text}' is a normalized field, keeping original position")
                    if is_field_content:
                        logger.info(f"[DEBUG] Region '{text_preview}' is field content, keeping adjusted position")
                        # 额外的debug输出
                        if '品牌设计' in region.text or '布国内各类广告' in region.text:
                            print(f"[ADJUSTED_REGIONS DEBUG] Adding field content: '{region.text[:50]}...'")
                            print(f"[ADJUSTED_REGIONS DEBUG]   is_field_content: {getattr(region, 'is_field_content', False)}")
                            print(f"[ADJUSTED_REGIONS DEBUG]   is_paragraph_merged: {getattr(region, 'is_paragraph_merged', False)}")
                            print(f"[ADJUSTED_REGIONS DEBUG]   belongs_to_field: {getattr(region, 'belongs_to_field', None)}")
                else:
                    # 非标准化字段，检查并调整位置以避免与标准化字段框重叠
                    adjusted_region = self._adjust_region_position(
                        region, normalized_field_boxes, image_width
                    )
                    adjusted_regions.append(adjusted_region)
        
        # ========== 第二阶段：抹除和渲染（使用调整后的位置） ==========
        
        # 步骤1：先抹除所有需要翻译的原始文字
        logger.debug("Step 2.1: Erasing all original text regions")
        for i, (region, result) in enumerate(zip(regions, results)):
            if not result.success or not result.translated_text.strip():
                continue
            
            # Skip erasing if should_render is False (e.g., QR codes)
            if hasattr(region, 'should_render') and not region.should_render:
                logger.info(f"Skipping erasing for region {i} (should_render=False)")
                continue
            
            # Skip erasing if hard_protected is True (hard protection mode for QR codes)
            if hasattr(region, 'hard_protected') and region.hard_protected:
                logger.info(f"🛡️ Skipping erasing for region {i} (hard_protected=True, QR code area)")
                continue
            
            # 🔴 跳过印章文字区域的抹除（不抹除印章内或被印章覆盖的文字）
            if hasattr(region, 'is_seal_text') and region.is_seal_text:
                logger.info(f"🔴 Skipping erasing for seal text region {i}: '{region.text}' (will render translation nearby)")
                continue
            
            try:
                # 采样背景色
                bg_color = self.background_sampler.sample_background(
                    output_image, region
                )
                
                # 抹除原始位置的文字（传递图标区域）
                output_image = self.background_sampler.process_region(
                    output_image, region, bg_color, icon_regions=icon_regions
                )
                logger.debug(f"Erased region {i}: '{region.text[:20]}...'")
                
            except Exception as e:
                logger.warning(f"Failed to erase region {i}: {e}")
        
        # 步骤1.5：抹除字段标签（如果有）
        if hasattr(self, '_field_label_regions') and self._field_label_regions:
            logger.info(f"Step 2.1.5: Erasing {len(self._field_label_regions)} field label regions")
            
            # 判断是否是竖版模式
            is_portrait = hasattr(self, 'current_orientation') and self.current_orientation == 'portrait'
            
            if is_portrait:
                # 竖版模式：使用统一最左边界抹除（确保完全擦除左侧字符）
                # 计算所有字段标签的统一最左边界
                min_left_x = float('inf')
                for label_region in self._field_label_regions:
                    # 使用 unified_left_boundary（如果存在），否则使用 bbox 的 x1
                    left_x = getattr(label_region, 'unified_left_boundary', label_region.bbox[0])
                    min_left_x = min(min_left_x, left_x)
                
                # 获取字段内容的统一左边界（如果存在）
                # 这样可以确保抹除范围覆盖从字段标签到字段内容之间的所有残留字符
                max_right_x = 0
                if hasattr(self, '_field_content_regions') and self._field_content_regions:
                    for content_region in self._field_content_regions:
                        if hasattr(content_region, 'unified_content_left_boundary'):
                            max_right_x = max(max_right_x, content_region.unified_content_left_boundary)
                            break  # 所有字段内容的统一左边界都相同，只需要取一次
                
                # 如果没有字段内容，使用字段标签的最右边界
                if max_right_x == 0:
                    for label_region in self._field_label_regions:
                        max_right_x = max(max_right_x, label_region.bbox[2])
                
                logger.info(f"竖版模式：使用统一最左边界抹除字段标签 x={min_left_x} 到 x={max_right_x}")
                
                for i, label_region in enumerate(self._field_label_regions):
                    try:
                        # 创建扩展到统一最左边界和字段内容左边界的区域
                        x1, y1, x2, y2 = label_region.bbox
                        expanded_bbox = (int(min_left_x), y1, int(max_right_x), y2)
                        
                        # 创建扩展后的 TextRegion
                        expanded_label_region = TextRegion(
                            bbox=expanded_bbox,
                            text=label_region.text,
                            confidence=label_region.confidence,
                            font_size=label_region.font_size,
                            angle=label_region.angle
                        )
                        
                        # 采样背景色
                        bg_color = self.background_sampler.sample_background(
                            output_image, expanded_label_region
                        )
                        
                        # 抹除字段标签（使用扩展后的区域）
                        output_image = self.background_sampler.process_region(
                            output_image, expanded_label_region, bg_color, icon_regions=icon_regions
                        )
                        logger.info(
                            f"Erased field label {i}: '{label_region.text}' "
                            f"from x={int(min_left_x)} to x={int(max_right_x)} (width={int(max_right_x - min_left_x)}px)"
                        )
                        
                    except Exception as e:
                        logger.warning(f"Failed to erase field label {i}: {e}")
            else:
                # 横版模式：使用原始 bbox 抹除（不扩展）
                logger.info(f"横版模式：使用原始边界抹除字段标签")
                
                for i, label_region in enumerate(self._field_label_regions):
                    try:
                        # 采样背景色
                        bg_color = self.background_sampler.sample_background(
                            output_image, label_region
                        )
                        
                        # 抹除字段标签（使用原始区域）
                        output_image = self.background_sampler.process_region(
                            output_image, label_region, bg_color, icon_regions=icon_regions
                        )
                        logger.info(f"Erased field label {i}: '{label_region.text}'")
                        
                    except Exception as e:
                        logger.warning(f"Failed to erase field label {i}: {e}")
        
        # 步骤1.6：抹除竖排组合字段（如果有）
        if hasattr(self, '_filtered_vertical_combined_regions') and self._filtered_vertical_combined_regions:
            logger.debug(f"Step 2.1.6: Erasing {len(self._filtered_vertical_combined_regions)} vertical combined regions")
            for i, combined_region in enumerate(self._filtered_vertical_combined_regions):
                try:
                    # 采样背景色
                    bg_color = self.background_sampler.sample_background(
                        output_image, combined_region
                    )
                    
                    # 抹除组合字段
                    output_image = self.background_sampler.process_region(
                        output_image, combined_region, bg_color, icon_regions=icon_regions
                    )
                    logger.debug(f"Erased vertical combined region {i}: '{combined_region.text}'")
                    
                except Exception as e:
                    logger.warning(f"Failed to erase vertical combined region {i}: {e}")
        
        # 步骤2：渲染所有翻译文字（使用调整后的位置）
        logger.debug("Step 2.2: Rendering all translated text")
        
        # 步骤2.1：处理印章文字区域（如果启用）
        seal_text_enabled = self.config.get('rendering.seal_text_handling.enabled', False)
        seal_text_positions = {}
        
        if seal_text_enabled and hasattr(self, '_seal_info'):
            logger.info("📍 处理印章文字区域的翻译位置...")
            
            # 获取印章文字区域
            seal_text_regions = [r for r in regions if getattr(r, 'is_seal_text', False)]
            
            logger.info(f"📍 检查 {len(regions)} 个区域中的印章文字标记...")
            for i, r in enumerate(regions):
                has_seal_text = getattr(r, 'is_seal_text', False)
                if has_seal_text:
                    logger.info(f"  区域 {i}: '{r.text}' - is_seal_text=True")
            
            if seal_text_regions:
                # 获取非印章文字区域(用于碰撞检测)
                # 使用adjusted_regions，因为它包含了调整后的位置
                non_seal_adjusted_regions = [adj_r for r, adj_r in zip(regions, adjusted_regions) 
                                            if not getattr(r, 'is_seal_text', False)]
                
                logger.info(
                    f"📍 印章文字区域: {len(seal_text_regions)}, "
                    f"非印章区域: {len(non_seal_adjusted_regions)}"
                )
                
                # 为印章文字寻找翻译位置
                # 传递adjusted_regions中的非印章区域，这样碰撞检测会使用调整后的位置
                # 从配置中读取是否启用碰撞检测
                use_collision_detection = self.config.get(
                    'rendering.seal_text_handling.use_collision_detection', 
                    True  # 默认启用
                )
                
                # 为每个印章文字区域准备翻译文本（用于精确尺寸计算）
                seal_text_translations = {}
                logger.info(f"📍 开始准备印章文字翻译，seal_text_regions 数量 = {len(seal_text_regions)}")
                for region in seal_text_regions:
                    region_idx = regions.index(region)
                    logger.info(f"📍 处理区域 {region_idx}: '{region.text}'")
                    if region_idx < len(results) and results[region_idx].success:
                        translated_text = results[region_idx].translated_text
                        
                        # 如果是印章内文字，添加[Seal:]前缀
                        seal_text_type = getattr(region, 'seal_text_type', 'seal_overlap')
                        if seal_text_type == 'seal_inner':
                            translated_text = f"[Seal: {translated_text}]"
                        
                        seal_text_translations[id(region)] = translated_text
                        logger.info(
                            f"📍 印章文字翻译: '{region.text}' -> '{translated_text}'"
                        )
                    else:
                        logger.warning(f"📍 区域 {region_idx} 翻译失败或不成功")
                
                logger.info(f"📍 印章文字翻译准备完成，共 {len(seal_text_translations)} 个")
                
                # 查找"Registration Authority"文本的位置
                registration_authority_bbox = None
                for region, result in zip(regions, results):
                    if result.success and result.translated_text:
                        # 检查翻译后的文本是否包含"Registration Authority"
                        if "Registration Authority" in result.translated_text:
                            registration_authority_bbox = region.bbox
                            logger.info(f"📍 找到Registration Authority位置: {registration_authority_bbox}")
                            break
                
                if not registration_authority_bbox:
                    logger.warning("📍 未找到Registration Authority文本，将使用印章位置作为参考")
                
                logger.info(f"📍 准备调用 find_translation_positions...")
                logger.info(f"📍 seal_text_handler = {self.seal_text_handler}")
                logger.info(f"📍 seal_text_regions 数量 = {len(seal_text_regions)}")
                logger.info(f"📍 use_collision_detection = {use_collision_detection}")
                logger.info(f"📍 registration_authority_bbox = {registration_authority_bbox}")
                
                try:
                    positioned_regions = self.seal_text_handler.find_translation_positions(
                        seal_text_regions,
                        non_seal_adjusted_regions,  # 使用调整后的非印章区域进行碰撞检测
                        self._seal_info['image_shape'],
                        use_collision_detection=use_collision_detection,
                        translations=seal_text_translations,  # 传递实际翻译文本
                        image=None,  # 第一步不传入图片，使用原始OCR区域
                        registration_authority_bbox=registration_authority_bbox  # 传递Registration Authority位置
                    )
                    logger.info(f"📍 find_translation_positions 调用成功，返回 {len(positioned_regions)} 个位置")
                except Exception as e:
                    logger.error(f"📍 find_translation_positions 调用失败: {e}")
                    import traceback
                    traceback.print_exc()
                    positioned_regions = []
                
                # 保存位置映射
                for region, new_bbox in positioned_regions:
                    seal_text_positions[id(region)] = new_bbox
                    logger.info(
                        f"📍 印章文字 '{region.text}' 新位置: {new_bbox}"
                    )
            else:
                logger.warning("📍 未找到标记为 is_seal_text=True 的区域!")
        
        # 步骤2.2：分两步渲染翻译文字
        # 第一步：渲染所有普通文字（非印章文字）
        # 第二步：使用渲染后的图片检测空白区域，然后渲染印章文字
        
        logger.info("=" * 80)
        logger.info("🎨 第一步：渲染所有普通文字（非印章文字）")
        logger.info("=" * 80)
        
        seal_text_to_render = []  # 保存印章文字，稍后渲染
        
        for i, (region, adjusted_region, result) in enumerate(zip(regions, adjusted_regions, results)):
            if not result.success:
                logger.debug(f"Skipping failed translation for region {i}")
                continue
            
            if not result.translated_text.strip():
                logger.debug(f"Skipping empty translation for region {i}")
                continue
            
            # Skip rendering if should_render is False (e.g., QR codes)
            if hasattr(region, 'should_render') and not region.should_render:
                logger.info(f"Skipping rendering for region {i} (should_render=False)")
                continue
            
            # Skip rendering if hard_protected is True (hard protection mode for QR codes)
            if hasattr(region, 'hard_protected') and region.hard_protected:
                logger.info(f"🛡️ Skipping rendering for region {i} (hard_protected=True, QR code area)")
                continue
            
            # 检查是否是印章文字区域（延迟渲染）
            is_seal_text = getattr(region, 'is_seal_text', False)
            if is_seal_text:
                # 这是印章文字，暂时不渲染，保存到列表中
                seal_text_to_render.append((i, region, adjusted_region, result))
                logger.info(f"🔴 跳过印章文字渲染（延迟到第二步）: '{region.text}'")
                continue
            
            # 旧的检查方式（保留作为备用）
            region_id = id(region)
            if region_id in seal_text_positions:
                # 这是印章文字，暂时不渲染，保存到列表中
                seal_text_to_render.append((i, region, adjusted_region, result))
                logger.info(f"📍 印章文字 '{region.text}' 将在第二步渲染")
                continue
            
            try:
                # 渲染普通文字
                render_region = adjusted_region
                translated_text = result.translated_text
                
                # 采样背景色
                bg_color = self.background_sampler.sample_background(
                    output_image, render_region
                )
                
                # 渲染翻译文本
                from src.rendering.text_renderer import TextAlignment
                output_image = self.text_renderer.render_text(
                    output_image,
                    render_region,
                    translated_text,
                    bg_color,
                    alignment=TextAlignment.LEFT
                )
                
                logger.debug(
                    f"Rendered region {i}: '{region.text[:20]}...' -> "
                    f"'{result.translated_text[:20]}...'"
                )
                
            except Exception as e:
                logger.warning(f"Failed to render region {i}: {e}")
                results[i] = TranslationResult(
                    source_text=result.source_text,
                    translated_text=result.translated_text,
                    confidence=result.confidence,
                    success=False,
                    error_message=f"Rendering failed: {e}"
                )
        
        # 第二步：使用渲染后的图片重新检测空白区域，然后渲染印章文字
        if seal_text_to_render:
            logger.info("=" * 80)
            logger.info(f"🎨 第二步：使用渲染后的图片检测空白区域，渲染 {len(seal_text_to_render)} 个印章文字")
            logger.info("=" * 80)
            
            # 重新调用 find_translation_positions，传入渲染后的图片
            if seal_text_regions:
                logger.info("📍 使用渲染后的图片重新查找印章文字位置...")
                
                # 重新查找"Registration Authority"文本的位置（可能已经被渲染）
                # 这次需要从渲染后的图片中查找
                registration_authority_bbox = None
                for region, result in zip(regions, results):
                    if result.success and result.translated_text:
                        # 检查翻译后的文本是否包含"Registration Authority"
                        if "Registration Authority" in result.translated_text:
                            # 使用调整后的区域位置
                            region_idx = regions.index(region)
                            if region_idx < len(adjusted_regions):
                                registration_authority_bbox = adjusted_regions[region_idx].bbox
                            else:
                                registration_authority_bbox = region.bbox
                            logger.info(f"📍 找到Registration Authority位置（渲染后）: {registration_authority_bbox}")
                            break
                
                if not registration_authority_bbox:
                    logger.warning("📍 未找到Registration Authority文本（渲染后），将使用印章位置作为参考")
                
                try:
                    # 传入渲染后的图片
                    positioned_regions_updated = self.seal_text_handler.find_translation_positions(
                        seal_text_regions,
                        non_seal_adjusted_regions,
                        self._seal_info['image_shape'],
                        use_collision_detection=use_collision_detection,
                        translations=seal_text_translations,
                        image=output_image,  # 传入渲染后的图片！
                        registration_authority_bbox=registration_authority_bbox  # 传递Registration Authority位置
                    )
                    
                    # 更新位置映射
                    seal_text_positions.clear()
                    for region, new_bbox in positioned_regions_updated:
                        seal_text_positions[id(region)] = new_bbox
                        logger.info(
                            f"📍 印章文字 '{region.text}' 更新后的位置: {new_bbox}"
                        )
                    
                    logger.info(f"📍 位置更新成功，共 {len(positioned_regions_updated)} 个位置")
                except Exception as e:
                    logger.error(f"📍 重新查找位置失败: {e}")
                    import traceback
                    traceback.print_exc()
            
            # 渲染印章文字
            for i, region, adjusted_region, result in seal_text_to_render:
                try:
                    region_id = id(region)
                    if region_id in seal_text_positions:
                        # 使用更新后的位置
                        new_bbox = seal_text_positions[region_id]
                        
                        # 获取印章文字翻译的字体大小配置
                        seal_font_size = self.config.get('rendering.seal_text_handling.font_size', 12)
                        
                        # 为印章内文字添加[Seal:]前缀
                        translated_text = result.translated_text
                        seal_text_type = getattr(region, 'seal_text_type', 'seal_overlap')
                        
                        if seal_text_type == 'seal_inner':
                            translated_text = f"[Seal: {translated_text}]"
                            logger.info(
                                f"📍 渲染印章内文字 '{region.text}' -> '{translated_text}' 到位置: {new_bbox}"
                            )
                        else:
                            logger.info(
                                f"📍 渲染印章覆盖文字 '{region.text}' -> '{translated_text}' 到位置: {new_bbox}"
                            )
                        
                        # 检测并处理文本与印章的重叠（自动换行）
                        seal_bbox = getattr(region, 'overlapping_seal', None)
                        logger.info(f"🔍 检查印章文本 '{region.text}' 的overlapping_seal属性: {seal_bbox}")
                        
                        if seal_bbox:
                            logger.info(f"🔍 检测印章文本是否需要换行: text='{translated_text[:50]}...'")
                            wrapped_text, adjusted_font_size, adjusted_bbox = self.seal_text_handler.check_text_seal_overlap_and_wrap(
                                translated_text=translated_text,
                                font_size=seal_font_size,
                                position_bbox=new_bbox,
                                seal_bbox=seal_bbox,
                                max_width=None  # 使用默认宽度
                            )
                            
                            # 如果换行成功，更新文本、字体大小和位置
                            if wrapped_text != translated_text or adjusted_bbox != new_bbox:
                                logger.info(
                                    f"✅ 文本已换行: "
                                    f"lines={len(wrapped_text.split(chr(10)))}, "
                                    f"font_size={seal_font_size}->{adjusted_font_size}, "
                                    f"bbox={new_bbox}->{adjusted_bbox}"
                                )
                                translated_text = wrapped_text
                                seal_font_size = adjusted_font_size
                                new_bbox = adjusted_bbox
                        else:
                            logger.warning(f"⚠️ 印章文本 '{region.text}' 没有overlapping_seal属性")
                        
                        # 创建新的 TextRegion 用于渲染
                        render_region = TextRegion(
                            bbox=new_bbox,
                            text=region.text,
                            confidence=region.confidence,
                            font_size=seal_font_size,
                            angle=region.angle
                        )
                        
                        # 采样背景色
                        bg_color = self.background_sampler.sample_background(
                            output_image, render_region
                        )
                        
                        # 渲染翻译文本
                        from src.rendering.text_renderer import TextAlignment
                        output_image = self.text_renderer.render_text(
                            output_image,
                            render_region,
                            translated_text,
                            bg_color,
                            alignment=TextAlignment.LEFT
                        )
                        
                        logger.info(f"✅ 印章文字渲染完成: '{region.text}'")
                    else:
                        logger.warning(f"⚠️ 印章文字 '{region.text}' 没有找到位置")
                        
                except Exception as e:
                    logger.warning(f"Failed to render seal text region {i}: {e}")
                    results[i] = TranslationResult(
                        source_text=result.source_text,
                        translated_text=result.translated_text,
                        confidence=result.confidence,
                        success=False,
                        error_message=f"Rendering failed: {e}"
                    )
        
        logger.info("=" * 80)
        logger.info("🎨 渲染完成")
        logger.info("=" * 80)
        
        return output_image, results
    

    def _filter_seal_regions(
        self,
        image: np.ndarray,
        regions: List[TextRegion]
    ) -> List[TextRegion]:
        """Filter text regions that overlap with seals.
        
        Strategy:
        1. Detect all seals in the image
        2. Check each text region for overlap with any seal
        3. If seal_text_handling is enabled:
           - Identify seal text regions
           - Keep them for later processing (translation near seal)
        4. Otherwise:
           - Filter out overlapping regions (original behavior)
        
        Args:
            image: Input image
            regions: List of text regions
            
        Returns:
            Filtered list of text regions (excluding those overlapping with seals,
            unless seal_text_handling is enabled)
        """
        seal_enabled = self.config.get('rendering.seal_handling.enabled', True)
        if not seal_enabled:
            return regions
        
        # 步骤1：检测图片中所有的印章
        seals = self._detect_all_seals(image)
        
        if not seals:
            logger.debug("未检测到印章，不进行过滤")
            return regions
        
        logger.info(f"🔴 检测到 {len(seals)} 个印章")
        for i, seal in enumerate(seals):
            logger.info(f"  印章 #{i+1}: ({seal[0]}, {seal[1]}) -> ({seal[2]}, {seal[3]})")
        
        # 检查是否启用了印章文字处理
        seal_text_enabled = self.config.get('rendering.seal_text_handling.enabled', False)
        
        if seal_text_enabled:
            logger.info("📍 印章文字翻译功能已启用，识别印章文字区域...")
            
            # 使用 SealTextHandler 识别印章文字区域
            seal_text_regions, non_seal_regions = self.seal_text_handler.identify_seal_text_regions(
                image, regions, seals
            )
            
            if seal_text_regions:
                logger.info(f"📍 识别到 {len(seal_text_regions)} 个印章文字区域")
                
                # 标记这些区域，以便后续特殊处理
                for region in seal_text_regions:
                    region.is_seal_text = True
                    region.needs_repositioning = True
                
                # 保存印章信息和图片尺寸，供后续使用
                self._seal_info = {
                    'seals': seals,
                    'seal_text_regions': seal_text_regions,
                    'image_shape': image.shape[:2],
                    'image': image  # 保存图片数据用于框检测
                }
                
                # 合并印章文字区域和非印章区域
                # 注意:seal_text_regions可能包含新创建的TextRegion(从GLM-OCR)
                # 这些区域不在原始regions列表中,需要添加进去
                all_regions = seal_text_regions + non_seal_regions
                
                logger.info(
                    f"📍 合并后总区域数: {len(all_regions)} "
                    f"(印章文字: {len(seal_text_regions)}, 非印章: {len(non_seal_regions)})"
                )
                
                # 返回合并后的所有区域
                return all_regions
            else:
                logger.info("📍 未识别到印章文字区域")
                return non_seal_regions
        else:
            # 原始行为：过滤与印章重叠的区域
            filtered_regions = []
            skipped_count = 0
            overlap_threshold = 0.01  # 重叠阈值：1%
            
            for region in regions:
                x1, y1, x2, y2 = region.bbox
                region_bbox = (x1, y1, x2, y2)
                
                # 检查该区域是否与任何印章重叠
                is_overlapping = False
                for seal_bbox in seals:
                    if self._check_overlap(region_bbox, seal_bbox, overlap_threshold):
                        is_overlapping = True
                        logger.info(f"🔴 区域 '{region.text}' ({x1}, {y1}) -> ({x2}, {y2}) 与印章 ({seal_bbox[0]}, {seal_bbox[1]}) -> ({seal_bbox[2]}, {seal_bbox[3]}) 重叠，跳过翻译")
                        break
                
                if is_overlapping:
                    skipped_count += 1
                else:
                    filtered_regions.append(region)
            
            logger.info(f"印章过滤: {len(regions)} 个区域 -> {len(filtered_regions)} 个区域 (跳过 {skipped_count} 个)")
            return filtered_regions
    
    def _merge_seal_covered_regions(self, regions: List[TextRegion]) -> List[TextRegion]:
        """合并被同一个印章覆盖的文本区域。
        
        当多个文本区域被同一个印章覆盖时（比如日期的不同部分），
        将它们合并成一个区域，以便一起翻译。
        
        合并策略：
        1. 识别所有被印章覆盖的区域（is_seal_text = True）
        2. 按照overlapping_seal分组（同一个印章覆盖的区域）
        3. 对每组区域：
           - 合并文本（按照从左到右、从上到下的顺序）
           - 计算合并后的边界框（包含所有子区域）
           - 创建新的TextRegion
        4. 保留未被印章覆盖的区域
        
        Args:
            regions: 文本区域列表
            
        Returns:
            合并后的文本区域列表
        """
        # 检查是否启用了印章文字处理
        seal_text_enabled = self.config.get('rendering.seal_text_handling.enabled', False)
        if not seal_text_enabled:
            return regions
        
        # 分离印章文本区域和非印章区域
        seal_text_regions = []
        non_seal_regions = []
        
        for region in regions:
            if getattr(region, 'is_seal_text', False):
                seal_text_regions.append(region)
            else:
                non_seal_regions.append(region)
        
        if not seal_text_regions:
            logger.debug("没有印章文本区域需要合并")
            return regions
        
        logger.info(f"🔍 开始合并印章覆盖的文本区域，共 {len(seal_text_regions)} 个印章文本区域")
        
        # 按照overlapping_seal分组
        seal_groups = {}
        for region in seal_text_regions:
            seal_bbox = getattr(region, 'overlapping_seal', None)
            if seal_bbox:
                # 使用seal_bbox作为key（转换为tuple以便作为dict key）
                seal_key = tuple(seal_bbox)
                if seal_key not in seal_groups:
                    seal_groups[seal_key] = []
                seal_groups[seal_key].append(region)
        
        logger.info(f"📍 识别到 {len(seal_groups)} 个印章组")
        
        # 合并每组区域
        merged_seal_regions = []
        for seal_bbox, group_regions in seal_groups.items():
            if len(group_regions) == 1:
                # 只有一个区域，不需要合并
                merged_seal_regions.append(group_regions[0])
                logger.debug(f"印章 {seal_bbox} 只有1个文本区域，不需要合并")
            else:
                # 多个区域，需要智能合并
                logger.info(f"📍 印章 {seal_bbox} 有 {len(group_regions)} 个文本区域，开始智能合并...")
                
                # 分离印章内文字和日期文字
                import re
                seal_inner_regions = []
                date_regions = []
                
                for region in group_regions:
                    text = region.text.strip()
                    
                    # 检查是否包含机构名称关键词（这些不是日期）
                    has_org_keywords = bool(re.search(
                        r'(管理局|监督|市场|政府|公司|有限|股份|集团|部门|委员会|办公室)',
                        text
                    ))
                    
                    # 检查是否包含日期模式
                    has_date = bool(re.search(
                        r'\d{4}年\d{1,2}月\d{1,2}日|\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2}|\d{4}年|\d{1,2}月|\d{1,2}日|\d{4}|\d{1,2}',
                        text
                    ))
                    
                    # 如果包含机构关键词，即使有数字也不是日期
                    if has_org_keywords:
                        seal_inner_regions.append(region)
                    elif has_date:
                        date_regions.append(region)
                    else:
                        seal_inner_regions.append(region)
                
                logger.info(
                    f"  分类结果: {len(seal_inner_regions)} 个印章内文字, "
                    f"{len(date_regions)} 个日期文字"
                )
                
                # 合并日期区域（如果有多个）
                if len(date_regions) > 1:
                    # 按照从左到右、从上到下的顺序排序
                    sorted_date_regions = sorted(date_regions, key=lambda r: (r.bbox[1], r.bbox[0]))
                    
                    # 合并文本
                    merged_date_text = ''.join([r.text.strip() for r in sorted_date_regions])
                    
                    # 计算合并后的边界框
                    min_x1 = min(r.bbox[0] for r in sorted_date_regions)
                    min_y1 = min(r.bbox[1] for r in sorted_date_regions)
                    max_x2 = max(r.bbox[2] for r in sorted_date_regions)
                    max_y2 = max(r.bbox[3] for r in sorted_date_regions)
                    merged_bbox = (min_x1, min_y1, max_x2, max_y2)
                    
                    # 计算平均置信度和字体大小
                    avg_confidence = sum(r.confidence for r in sorted_date_regions) / len(sorted_date_regions)
                    avg_font_size = sum(r.font_size for r in sorted_date_regions) / len(sorted_date_regions)
                    
                    # 创建合并后的日期区域
                    merged_date_region = TextRegion(
                        bbox=merged_bbox,
                        text=merged_date_text,
                        confidence=avg_confidence,
                        font_size=int(avg_font_size),
                        angle=sorted_date_regions[0].angle
                    )
                    
                    # 保留印章相关属性
                    merged_date_region.is_seal_text = True
                    merged_date_region.needs_repositioning = True
                    merged_date_region.overlapping_seal = seal_bbox
                    merged_date_region.seal_text_type = 'seal_overlap'  # 日期是被覆盖的
                    
                    merged_seal_regions.append(merged_date_region)
                    
                    logger.info(
                        f"✅ 合并日期文本: {[r.text for r in sorted_date_regions]} → '{merged_date_text}'"
                    )
                elif len(date_regions) == 1:
                    # 只有一个日期区域，直接添加
                    date_regions[0].seal_text_type = 'seal_overlap'
                    merged_seal_regions.append(date_regions[0])
                
                # 添加印章内文字区域（不合并）
                for region in seal_inner_regions:
                    region.seal_text_type = 'seal_inner'
                    merged_seal_regions.append(region)
                    logger.info(
                        f"✅ 保留印章内文字: '{region.text}'"
                    )
        
        # 合并所有区域
        all_regions = merged_seal_regions + non_seal_regions
        
        logger.info(
            f"🎯 印章文本合并完成: {len(seal_text_regions)} 个印章文本区域 → "
            f"{len(merged_seal_regions)} 个合并后的区域 "
            f"(总区域数: {len(regions)} → {len(all_regions)})"
        )
        
        # ========== 交互式验证：让用户确认印章文字识别结果 ==========
        interactive_enabled = self.config.get(
            'rendering.seal_text_handling.interactive_verification.enabled',
            False
        )
        
        if interactive_enabled and merged_seal_regions:
            logger.info("=" * 60)
            logger.info("🔍 交互式验证：开始验证印章文字识别结果")
            logger.info("=" * 60)
            
            verified_seal_regions = []
            
            # 检查是否有GUI回调函数
            gui_callback = getattr(self, '_gui_verification_callback', None)
            
            for i, region in enumerate(merged_seal_regions, 1):
                text_type = getattr(region, 'seal_text_type', 'seal_overlap')
                type_name = "印章内文字" if text_type == 'seal_inner' else "被覆盖的日期"
                
                # 如果有GUI回调，使用GUI模式
                if gui_callback:
                    region_info = {
                        'type_name': type_name,
                        'text_type': text_type,
                        'bbox': region.bbox,
                        'text': region.text,
                        'confidence': region.confidence,
                        'index': i,
                        'total': len(merged_seal_regions)
                    }
                    
                    print(f"\n[GUI模式] 正在显示验证对话框...")
                    print(f"  类型: {type_name}")
                    print(f"  内容: \"{region.text}\"")
                    
                    try:
                        action, corrected_text = gui_callback(region_info)
                        
                        print(f"[GUI模式] 用户选择: {action}")
                        
                        if action == 'confirm':
                            print(f"确认识别正确: \"{region.text}\"")
                            verified_seal_regions.append(region)
                        elif action == 'correct' and corrected_text:
                            print(f"修正内容: \"{region.text}\" -> \"{corrected_text}\"")
                            region.text = corrected_text
                            region.user_corrected = True
                            verified_seal_regions.append(region)
                            logger.info(f"✏️  用户修正内容: 原文='{region.text}' -> 修正='{corrected_text}'")
                        elif action == 'skip':
                            print(f"跳过区域: \"{region.text}\"")
                            logger.info(f"⏭️  用户跳过区域: '{region.text}'")
                        else:
                            # 默认确认
                            print(f"[GUI模式] 默认确认")
                            verified_seal_regions.append(region)
                    except Exception as e:
                        logger.error(f"GUI回调错误: {e}")
                        import traceback
                        traceback.print_exc()
                        # 出错时自动确认
                        verified_seal_regions.append(region)
                    
                    continue
                
                # 终端模式
                # 显示识别结果
                print("\n" + "=" * 60)
                print(f"检测到{type_name} (#{i}/{len(merged_seal_regions)})：")
                print("-" * 60)
                print(f"类型: {text_type}")
                print(f"位置: {region.bbox}")
                print(f"识别内容: \"{region.text}\"")
                print(f"置信度: {region.confidence:.2f}")
                print("-" * 60)
                
                # 获取用户输入
                while True:
                    try:
                        response = input("识别是否正确？(y=是/n=否/s=跳过): ").strip().lower()
                        
                        if response in ['y', 'yes', '是', '']:
                            print("✅ 确认识别正确，继续处理")
                            verified_seal_regions.append(region)
                            break
                        
                        elif response in ['n', 'no', '否']:
                            corrected_text = input("请输入正确的文字内容: ").strip()
                            if corrected_text:
                                print(f"✅ 使用修正后的内容: \"{corrected_text}\"")
                                region.text = corrected_text
                                region.user_corrected = True
                                verified_seal_regions.append(region)
                                logger.info(f"✏️  用户修正内容: 原文='{region.text}' -> 修正='{corrected_text}'")
                                break
                            else:
                                print("❌ 输入为空，请重新选择")
                                continue
                        
                        elif response in ['s', 'skip', '跳过']:
                            print("⏭️  跳过此区域")
                            logger.info(f"⏭️  用户跳过区域: '{region.text}'")
                            break
                        
                        else:
                            print("❌ 无效输入，请输入 y/n/s")
                            continue
                    
                    except KeyboardInterrupt:
                        print("\n⚠️  用户中断操作")
                        logger.warning("用户中断交互式验证")
                        # 保留已验证的区域
                        break
                    except EOFError:
                        print("\n⚠️  输入流结束，自动确认当前区域")
                        logger.warning("输入流结束，自动确认")
                        verified_seal_regions.append(region)
                        break
                        break
            
            # 更新合并后的印章区域列表
            merged_seal_regions = verified_seal_regions
            
            # 重新合并所有区域
            all_regions = merged_seal_regions + non_seal_regions
            
            logger.info("=" * 60)
            logger.info(f"✅ 交互式验证完成：保留 {len(merged_seal_regions)} 个印章文字区域")
            logger.info("=" * 60)
        
        return all_regions
    
    def _get_normalized_field_boxes(self, regions: List[TextRegion]) -> List:
        """获取所有标准化字段框作为障碍物。
        
        从regions列表中提取所有被标准化的字段框（通过检查字段名是否在配置的
        标准化字段列表中）。
        
        Args:
            regions: 所有文本区域列表
            
        Returns:
            BoundingBox对象列表，表示标准化字段框
        """
        from src.overlap import BoundingBox
        
        normalize_config = self.config.get('rendering.normalize_fields', {})
        if not normalize_config.get('enabled', False):
            return []
        
        field_groups = normalize_config.get('field_groups', [])
        all_normalized_field_names = []
        
        for group in field_groups:
            chinese_fields = group.get('chinese_fields', [])
            english_fields = group.get('english_fields', [])
            all_normalized_field_names.extend(chinese_fields)
            all_normalized_field_names.extend(english_fields)
        
        normalized_boxes = []
        for region in regions:
            if region.text in all_normalized_field_names:
                x1, y1, x2, y2 = region.bbox
                box = BoundingBox(
                    x=x1,
                    y=y1,
                    width=x2 - x1,
                    height=y2 - y1
                )
                normalized_boxes.append(box)
                logger.debug(f"Added normalized field box: '{region.text}' at ({x1}, {y1})")
        
        logger.info(f"Found {len(normalized_boxes)} normalized field boxes as obstacles")
        return normalized_boxes
    
    def _get_normalized_field_texts(self) -> List[str]:
        """获取所有标准化字段的文本名称（中文）。
        
        Returns:
            标准化字段文本列表（中文字段名）
        """
        normalize_config = self.config.get('rendering.normalize_fields', {})
        if not normalize_config.get('enabled', False):
            return []
        
        field_groups = normalize_config.get('field_groups', [])
        all_normalized_field_names = []
        
        for group in field_groups:
            chinese_fields = group.get('chinese_fields', [])
            # 只返回中文字段名，因为region.text是翻译前的中文
            all_normalized_field_names.extend(chinese_fields)
        
        return all_normalized_field_names
    
    def _adjust_region_position(
        self,
        region: TextRegion,
        obstacles: List,
        image_width: int
    ) -> TextRegion:
        """调整文本区域位置以避免与障碍物重叠。
        
        策略：只调整那些**左边界**与标准化字段重叠的文本框。
        如果文本框在标准化字段的右边（正常的字段-值布局），不调整。
        
        Args:
            region: 原始文本区域
            obstacles: 障碍物列表（BoundingBox对象）
            image_width: 图像宽度
            
        Returns:
            调整后的TextRegion（如果不需要调整则返回原region）
        """
        from src.overlap import BoundingBox, OverlapDetector
        
        # 如果没有障碍物或重叠防止未启用，直接返回原region
        overlap_enabled = self.config.get('rendering.overlap_prevention.enabled', True)
        if not overlap_enabled or not obstacles:
            return region
        
        # 创建当前region的BoundingBox
        x1, y1, x2, y2 = region.bbox
        text_box = BoundingBox(
            x=x1,
            y=y1,
            width=x2 - x1,
            height=y2 - y1
        )
        
        # 检查是否有需要调整的重叠
        # 策略：只有当文本框的左边界在标准化字段内部时才调整
        # 如果文本框完全在标准化字段右边，说明是正常的字段-值布局，不调整
        detector = OverlapDetector()
        needs_adjustment = False
        
        for obstacle in obstacles:
            if detector.detect_overlap(text_box, obstacle):
                # 有重叠，但需要判断是否是正常的左右布局
                # 如果文本框的左边界在障碍物的右边界之后，说明是正常布局
                if text_box.x >= obstacle.x2:
                    # 文本框在障碍物右边，是正常布局，不调整
                    logger.debug(
                        f"Region '{region.text}' is to the right of normalized field, "
                        f"keeping original position (normal field-value layout)"
                    )
                    continue
                else:
                    # 文本框的左边界在障碍物内部或左边，需要调整
                    needs_adjustment = True
                    logger.info(
                        f"Region '{region.text}' overlaps with normalized field "
                        f"(text_x={text_box.x}, obstacle_x2={obstacle.x2}), needs adjustment"
                    )
                    break
        
        # 如果不需要调整，直接返回原region
        if not needs_adjustment:
            return region
        
        # 需要调整，进行位置调整
        adjusted_box = self.position_adjuster.adjust_position(
            text_box, obstacles, image_width
        )
        
        # 如果位置改变了，创建新的TextRegion
        if adjusted_box.x != text_box.x:
            logger.info(
                f"Adjusted region position: '{region.text}' "
                f"from x={text_box.x} to x={adjusted_box.x} "
                f"(moved {adjusted_box.x - text_box.x:.1f} pixels)"
            )
            
            # 检查调整后是否超出图像边界太多
            if adjusted_box.x2 > image_width:
                logger.warning(
                    f"Adjusted region '{region.text}' exceeds image width: "
                    f"x2={adjusted_box.x2}, image_width={image_width}. "
                    f"Keeping original position to avoid losing content."
                )
                return region
            
            # 创建调整后的TextRegion，保留original_x1、unified_left_boundary和unified_content_left_boundary属性（如果存在）
            adjusted_region = TextRegion(
                bbox=(adjusted_box.x, adjusted_box.y, adjusted_box.x2, adjusted_box.y2),
                text=region.text,
                confidence=region.confidence,
                font_size=region.font_size,
                angle=region.angle,
                is_vertical_merged=getattr(region, 'is_vertical_merged', False)
            )
            
            # 保留original_x1属性（如果原region有的话）
            if hasattr(region, 'original_x1'):
                adjusted_region.original_x1 = region.original_x1
                logger.debug(f"Preserved original_x1={region.original_x1} for adjusted region '{region.text}'")
            
            # 保留unified_left_boundary属性（如果原region有的话）
            if hasattr(region, 'unified_left_boundary'):
                adjusted_region.unified_left_boundary = region.unified_left_boundary
                logger.debug(f"Preserved unified_left_boundary={region.unified_left_boundary} for adjusted region '{region.text}'")
            
            # 保留unified_content_left_boundary属性（如果原region有的话）
            if hasattr(region, 'unified_content_left_boundary'):
                adjusted_region.unified_content_left_boundary = region.unified_content_left_boundary
                logger.debug(f"Preserved unified_content_left_boundary={region.unified_content_left_boundary} for adjusted region '{region.text}'")
            
            # 保留is_field_label属性（如果原region有的话）
            if hasattr(region, 'is_field_label'):
                adjusted_region.is_field_label = region.is_field_label
                logger.debug(f"Preserved is_field_label={region.is_field_label} for adjusted region '{region.text}'")
            
            # 保留is_field_content属性（如果原region有的话）
            if hasattr(region, 'is_field_content'):
                adjusted_region.is_field_content = region.is_field_content
                logger.debug(f"Preserved is_field_content={region.is_field_content} for adjusted region '{region.text}'")
            
            # 保留is_paragraph_merged属性（如果原region有的话）
            if hasattr(region, 'is_paragraph_merged'):
                adjusted_region.is_paragraph_merged = region.is_paragraph_merged
                logger.debug(f"Preserved is_paragraph_merged={region.is_paragraph_merged} for adjusted region '{region.text}'")
                # 额外的debug输出
                if region.is_paragraph_merged:
                    print(f"[ADJUST DEBUG] Preserved is_paragraph_merged=True for '{region.text[:50]}...'")
            
            # 保留belongs_to_field属性（如果原region有的话）
            if hasattr(region, 'belongs_to_field'):
                adjusted_region.belongs_to_field = region.belongs_to_field
                logger.debug(f"Preserved belongs_to_field={region.belongs_to_field} for adjusted region '{region.text}'")
                # 额外的debug输出
                if region.belongs_to_field:
                    print(f"[ADJUST DEBUG] Preserved belongs_to_field='{region.belongs_to_field}' for '{region.text[:50]}...'")
            
            # 保留is_seal_text属性（如果原region有的话）
            if hasattr(region, 'is_seal_text'):
                adjusted_region.is_seal_text = region.is_seal_text
                logger.debug(f"Preserved is_seal_text={region.is_seal_text} for adjusted region '{region.text}'")
            
            # 保留overlapping_seal属性（如果原region有的话）
            if hasattr(region, 'overlapping_seal'):
                adjusted_region.overlapping_seal = region.overlapping_seal
                logger.debug(f"Preserved overlapping_seal={region.overlapping_seal} for adjusted region '{region.text}'")
            
            return adjusted_region
        
        return region
    
    def _erase_left_residual_chars(
        self,
        image: np.ndarray,
        region: TextRegion,
        translated_text: str,
        bg_color: Tuple[int, int, int]
    ) -> np.ndarray:
        """擦除字段标签左侧可能存在的残留中文字。
        
        改进策略：
        1. 从安全边界（避开边框）一直擦除到字段标签的左边界
        2. 只擦除字段标签高度范围内的区域，不影响上下内容
        3. 使用智能边界检测，避免擦除边框
        
        Args:
            image: 输入图片（BGR格式）
            region: 文本区域（原始位置，可能只包含部分字段名）
            translated_text: 翻译后的文本（英文）
            bg_color: 背景颜色（RGB格式）
            
        Returns:
            处理后的图片
        """
        # 需要擦除左侧残留的字段标签（英文）
        field_labels_to_erase = [
            "Name",
            "Type",
            "Legal Representative",
            "Registered Capital",
            "Date of Establishment",
            "Term of Operation",
            "Business Scope",
            "Domicile"  # 住所的翻译结果
        ]
        
        # 检查当前翻译文本是否是需要处理的字段标签
        if translated_text not in field_labels_to_erase:
            logger.debug(f"Skipped erasing for '{translated_text}': not a field label")
            return image
        
        logger.info(f"[ERASE] Processing field label: '{translated_text}'")
        
        # 计算左侧擦除区域
        x1, y1, x2, y2 = region.bbox
        height, width = image.shape[:2]
        
        # 关键改进：从安全边界（避开边框）一直擦除到字段标签的左边界
        # 营业执照的左边框通常在x=50-80px左右，我们从x=100开始擦除以确保安全
        safe_left_boundary = 100
        
        # 计算擦除区域的坐标
        # 从安全边界一直擦除到region的左边界
        erase_x1 = safe_left_boundary
        erase_y1 = y1
        erase_x2 = x1  # 擦除到region的左边界
        erase_y2 = y2
        
        # 确保坐标在图片范围内
        erase_x1 = max(0, min(erase_x1, width))
        erase_y1 = max(0, min(erase_y1, height))
        erase_x2 = max(0, min(erase_x2, width))
        erase_y2 = max(0, min(erase_y2, height))
        
        logger.info(
            f"[ERASE] Erase region calculated: "
            f"x1={erase_x1}, y1={erase_y1}, x2={erase_x2}, y2={erase_y2}, "
            f"width={erase_x2 - erase_x1}px, height={erase_y2 - erase_y1}px"
        )
        
        # 如果擦除区域有效且宽度合理，进行擦除
        if erase_x2 > erase_x1 and erase_y2 > erase_y1:
            # 创建擦除区域的TextRegion
            erase_region = TextRegion(
                bbox=(erase_x1, erase_y1, erase_x2, erase_y2),
                text="",
                confidence=1.0,
                font_size=region.font_size,
                angle=0
            )
            
            # 使用background_sampler擦除该区域
            result = self.background_sampler.process_region(
                image, erase_region, bg_color
            )
            
            logger.info(
                f"[ERASE] Successfully erased left residual chars for '{translated_text}': "
                f"region=({erase_x1}, {erase_y1}, {erase_x2}, {erase_y2}), "
                f"width={erase_x2 - erase_x1}px"
            )
            
            return result
        else:
            logger.warning(
                f"[ERASE] Invalid erase region for '{translated_text}': "
                f"x1={erase_x1}, x2={erase_x2}, skipped"
            )
            return image
    
    def _merge_continuous_regions(self, regions: List[TextRegion]) -> List[TextRegion]:
        """根据行间距合并连续的文本区域。
        
        使用字段感知的段落合并器来智能合并文本区域。
        该方法保持向后兼容性，同时提供更智能的合并策略。
        
        策略：
        1. 使用FieldAwareParagraphMerger进行字段感知合并
        2. 识别字段标签和对应的内容块
        3. 对长文本字段（如"经营范围"）应用特殊的合并规则
        4. 保持现有功能（字段标签分离、图标过滤等）不受影响
        
        Args:
            regions: 原始文本区域列表
            
        Returns:
            合并后的文本区域列表
        """
        if not regions:
            return []
        
        # 使用FieldAwareParagraphMerger进行合并
        return self.paragraph_merger.merge_regions(regions)
    
    def _merge_single_char_field_fragments(self, regions: List[TextRegion]) -> List[TextRegion]:
        """合并单字字段碎片（如"名"+"称"→"名称"）。
        
        这个方法专门处理OCR把字段标签识别成单个字符的情况。
        它会检测竖排或横排的单字碎片，并将它们合并成完整的字段标签。
        
        支持的字段碎片模式：
        - 横排：("名", "称") → "名称"
        - 竖排：("注", "册", "资", "本") → "注册资本"（从上到下）
        
        Args:
            regions: 输入的文本区域列表
            
        Returns:
            合并后的文本区域列表
        """
        if not regions or len(regions) < 2:
            return regions
        
        logger.info(f"=" * 80)
        logger.info(f"开始合并单字字段碎片，输入 {len(regions)} 个区域")
        logger.info(f"=" * 80)
        
        # 定义单字碎片模式（按顺序）
        # 格式：(碎片元组, 完整字段名, 最大距离, 方向)
        # 方向：'horizontal' 表示横排，'vertical' 表示竖排
        single_char_patterns = [
            # 横排模式（从左到右）
            (("名", "称"), "名称", 100, 'horizontal'),
            (("类", "型"), "类型", 100, 'horizontal'),
            (("住", "所"), "住所", 100, 'horizontal'),
            
            # 竖排模式（从上到下）
            (("注", "册", "资", "本"), "注册资本", 50, 'vertical'),
            (("注", "册", "资"), "注册资本", 50, 'vertical'),  # 缺少"本"的情况
            (("成", "立", "日", "期"), "成立日期", 50, 'vertical'),
            (("成", "立", "日"), "成立日期", 50, 'vertical'),  # 缺少"期"的情况
            (("成", "立"), "成立日期", 50, 'vertical'),  # 只有"成立"的情况
            (("经", "营", "期", "限"), "营业期限", 50, 'vertical'),
            (("经", "营", "期"), "营业期限", 50, 'vertical'),  # 缺少"限"的情况
            (("经", "营"), "营业期限", 50, 'vertical'),  # 只有"经营"的情况
            (("法", "定", "代", "表", "人"), "法定代表人", 50, 'vertical'),
            (("法", "定", "代", "表"), "法定代表人", 50, 'vertical'),
            (("法", "定", "代"), "法定代表人", 50, 'vertical'),
        ]
        
        merged_regions = []
        used_indices = set()
        merge_count = 0
        
        # 按从上到下、从左到右排序
        sorted_regions = sorted(regions, key=lambda r: (r.bbox[1], r.bbox[0]))
        
        for i, region in enumerate(sorted_regions):
            if i in used_indices:
                continue
            
            text = region.text.strip()
            
            # 只处理单字或双字区域
            if len(text) > 2:
                merged_regions.append(region)
                continue
            
            # 尝试匹配每个模式
            matched = False
            
            for pattern, full_field, max_distance, direction in single_char_patterns:
                # 检查当前区域是否是模式的第一个字符
                if text != pattern[0]:
                    continue
                
                # 尝试找到后续的字符
                candidate_regions = [region]
                candidate_indices = [i]
                
                for target_char in pattern[1:]:
                    found = False
                    
                    # 在后续区域中查找目标字符
                    for j in range(i + 1, len(sorted_regions)):
                        if j in used_indices or j in candidate_indices:
                            continue
                        
                        next_region = sorted_regions[j]
                        next_text = next_region.text.strip()
                        
                        # 检查是否是目标字符
                        if next_text != target_char:
                            continue
                        
                        # 检查位置关系
                        last_region = candidate_regions[-1]
                        
                        if direction == 'horizontal':
                            # 横排：检查是否在同一行且水平距离合理
                            y_diff = abs(next_region.bbox[1] - last_region.bbox[1])
                            x_diff = next_region.bbox[0] - last_region.bbox[2]
                            
                            if y_diff < 10 and 0 <= x_diff <= max_distance:
                                candidate_regions.append(next_region)
                                candidate_indices.append(j)
                                found = True
                                break
                        
                        elif direction == 'vertical':
                            # 竖排：检查是否在同一列且垂直距离合理
                            x_diff = abs(next_region.bbox[0] - last_region.bbox[0])
                            y_diff = next_region.bbox[1] - last_region.bbox[3]
                            
                            if x_diff < 20 and 0 <= y_diff <= max_distance:
                                candidate_regions.append(next_region)
                                candidate_indices.append(j)
                                found = True
                                break
                    
                    # 如果没有找到目标字符，停止匹配这个模式
                    if not found:
                        break
                
                # 检查是否找到了完整的模式
                if len(candidate_regions) == len(pattern):
                    # 合并这些区域
                    merged_region = self._merge_char_regions(candidate_regions, full_field)
                    merged_regions.append(merged_region)
                    
                    # 标记这些区域已被使用
                    for idx in candidate_indices:
                        used_indices.add(idx)
                    
                    matched = True
                    merge_count += 1
                    
                    chars_str = " + ".join([f"'{r.text}'" for r in candidate_regions])
                    logger.info(
                        f"✅ 合并单字碎片: {chars_str} → '{full_field}' "
                        f"(方向={direction}, 区域数={len(candidate_regions)})"
                    )
                    break
            
            # 如果没有匹配任何模式，保留原区域
            if not matched:
                merged_regions.append(region)
        
        logger.info(f"=" * 80)
        logger.info(f"单字碎片合并完成: {len(regions)} 个区域 → {len(merged_regions)} 个区域 (合并了 {merge_count} 组)")
        logger.info(f"=" * 80)
        
        return merged_regions
    
    def _merge_char_regions(
        self, 
        regions: List[TextRegion], 
        merged_text: str
    ) -> TextRegion:
        """合并多个字符区域成一个完整的字段标签。
        
        Args:
            regions: 要合并的字符区域列表
            merged_text: 合并后的文本
            
        Returns:
            合并后的TextRegion
        """
        # 计算合并后的边界框
        min_x1 = min(r.bbox[0] for r in regions)
        min_y1 = min(r.bbox[1] for r in regions)
        max_x2 = max(r.bbox[2] for r in regions)
        max_y2 = max(r.bbox[3] for r in regions)
        
        # 使用第一个区域的其他属性
        first_region = regions[0]
        
        # 创建合并后的区域
        merged_region = TextRegion(
            bbox=(min_x1, min_y1, max_x2, max_y2),
            text=merged_text,
            confidence=first_region.confidence,
            font_size=first_region.font_size,
            angle=first_region.angle
        )
        
        # 标记为字段标签
        merged_region.is_field_label = True
        
        return merged_region
    
    def _split_partial_label_content(self, regions: List[TextRegion]) -> List[TextRegion]:
        """分割"部分标签+内容"的区域（如"称佛山市..."→"称"+"佛山市..."）。
        
        这个方法处理OCR把字段标签的最后一个字和内容识别成一个区域的情况。
        例如：
        - "称佛山市炜烨餐饮管理有限公司" → "称" + "佛山市炜烨餐饮管理有限公司"
        - "所佛山市南海区桂城街道..." → "所" + "佛山市南海区桂城街道..."
        
        Args:
            regions: 输入的文本区域列表
            
        Returns:
            分割后的文本区域列表
        """
        if not regions:
            return regions
        
        logger.info(f"=" * 80)
        logger.info(f"开始分割部分标签+内容区域，输入 {len(regions)} 个区域")
        logger.info(f"=" * 80)
        
        # 定义可能的部分标签模式
        # 格式：(部分标签字符, 完整标签, 前面应该有的字符)
        partial_patterns = [
            ("称", "名称", "名"),
            ("所", "住所", "住"),
            ("本", "注册资本", "资"),
            ("期", "成立日期", "日"),
            ("限", "营业期限", "期"),
            ("围", "经营范围", "范"),
        ]
        
        split_regions = []
        split_count = 0
        
        for region in regions:
            text = region.text.strip()
            
            # 检查是否匹配部分标签模式
            matched = False
            
            for partial_char, full_label, prev_char in partial_patterns:
                # 检查是否以部分标签字符开头，且后面有内容
                if text.startswith(partial_char) and len(text) > 1:
                    # 检查后面的内容是否看起来像字段内容（不是标签的一部分）
                    remaining = text[1:].strip()
                    
                    # 如果剩余部分是标签的一部分，跳过
                    # 例如："称号"不应该被分割
                    if remaining and remaining[0] not in "名类住法注成营经登商":
                        # 分割这个区域
                        x1, y1, x2, y2 = region.bbox
                        width = x2 - x1
                        
                        # 估算部分标签字符的宽度（假设等宽）
                        char_width = width / len(text)
                        label_width = int(char_width)
                        
                        # 创建部分标签区域
                        label_region = TextRegion(
                            bbox=(x1, y1, x1 + label_width, y2),
                            text=partial_char,
                            confidence=region.confidence,
                            font_size=region.font_size,
                            angle=region.angle
                        )
                        
                        # 创建内容区域
                        content_region = TextRegion(
                            bbox=(x1 + label_width, y1, x2, y2),
                            text=remaining,
                            confidence=region.confidence,
                            font_size=region.font_size,
                            angle=region.angle
                        )
                        
                        split_regions.append(label_region)
                        split_regions.append(content_region)
                        
                        split_count += 1
                        matched = True
                        
                        logger.info(
                            f"✅ 分割部分标签+内容: '{text[:30]}...' → "
                            f"'{partial_char}' + '{remaining[:20]}...'"
                        )
                        break
            
            # 如果没有匹配任何模式，保留原区域
            if not matched:
                split_regions.append(region)
        
        logger.info(f"=" * 80)
        logger.info(f"部分标签+内容分割完成: {len(regions)} 个区域 → {len(split_regions)} 个区域 (分割了 {split_count} 个)")
        logger.info(f"=" * 80)
        
        return split_regions
    
    def _normalize_field_regions(self, regions: List[TextRegion]) -> List[TextRegion]:
        """标准化特定字段的边界框（在翻译之前）。
        
        对于配置中指定的字段（如"成立日期"、"注册资本"等），
        将它们的边界框扩展到该组字段中最大的宽度。
        
        横版和竖版使用不同的逻辑：
        - 横版：简单标准化（commit 8959018 的逻辑），不做字段内容识别和y坐标调整
        - 竖版：扩展不完整字段标签的左边界，识别字段内容
        
        Args:
            regions: 文本区域列表（中文）
            
        Returns:
            标准化后的文本区域列表
        """
        # 步骤0：合并单字碎片（如"名"+"称"→"名称"）
        # 这一步在横版和竖版模式下都需要执行
        regions = self._merge_single_char_field_fragments(regions)
        
        # 判断是否是竖版模式
        is_portrait = hasattr(self, 'current_orientation') and self.current_orientation == 'portrait'
        
        if not is_portrait:
            # 横版模式：使用 commit 8959018 的简单逻辑（只标准化字段边界框，不做其他处理）
            logger.info(f"横版模式：开始标准化字段边界框（翻译前），共 {len(regions)} 个区域")
            
            normalized_regions = []
            field_label_regions = []  # 收集所有字段标签
            
            for i, region in enumerate(regions):
                # 保存原始bbox（标准化前）
                original_bbox = region.bbox
                
                # 使用中文字段名进行标准化
                normalized_region = self.text_renderer._normalize_field_bbox(region, regions)
                
                # 如果是字段标签，保存原始bbox
                if hasattr(normalized_region, 'is_field_label') and normalized_region.is_field_label:
                    normalized_region.original_bbox = original_bbox
                    field_label_regions.append(normalized_region)
                
                normalized_regions.append(normalized_region)
                
                # 如果宽度变化了，打印信息
                if normalized_region.width != region.width:
                    logger.info(
                        f"标准化字段: '{region.text}' - "
                        f"{region.width}px -> {normalized_region.width}px"
                    )
            
            # 计算并设置统一字体大小（基于最大高度）
            unified_font_enabled = self.config.get('rendering.unified_field_label_font_size.enabled', True)
            
            if unified_font_enabled and field_label_regions:
                # 找到所有字段标签中最大的高度
                max_field_height = 0
                for region in field_label_regions:
                    field_height = region.bbox[3] - region.bbox[1]
                    max_field_height = max(max_field_height, field_height)
                    logger.debug(f"字段 '{region.text}' 高度: {field_height}px")
                
                # 计算统一字体大小：max_height × 0.5
                unified_font_size = int(max_field_height * 0.6)
                
                # 限制在合理范围内 [8, 48]
                unified_font_size = max(8, min(48, unified_font_size))
                
                logger.info(
                    f"横版模式：计算字段标签统一字体大小: max_height={max_field_height}px, "
                    f"unified_font_size={unified_font_size}px (max_height × 0.5)"
                )
                
                # 为所有字段标签设置统一字体大小
                for region in field_label_regions:
                    original_font_size = region.font_size
                    region.font_size = unified_font_size
                    # 添加标记，表示使用了统一字体大小
                    region.uses_unified_font_size = True
                    logger.info(
                        f"字段标签 '{region.text}' 字体大小: {original_font_size}px -> {unified_font_size}px (统一)"
                    )
            elif not unified_font_enabled:
                logger.info("横版模式：统一字体大小功能已禁用")
            
            # 【新增】识别左右两列的字段标签，分别对齐
            if field_label_regions:
                # 获取图片宽度
                image_width = regions[0].bbox[2] if regions else 1000  # 默认值
                for region in regions:
                    image_width = max(image_width, region.bbox[2])
                
                # 使用图片中心线分隔左右两列
                center_x = image_width / 2
                
                # 分离左右两列的字段标签
                left_column_labels = []
                right_column_labels = []
                
                for region in field_label_regions:
                    x1 = region.bbox[0]
                    # 根据左边界判断属于哪一列
                    if x1 < center_x:
                        left_column_labels.append(region)
                    else:
                        right_column_labels.append(region)
                
                logger.info(
                    f"横版模式：识别到 {len(left_column_labels)} 个左列字段标签, "
                    f"{len(right_column_labels)} 个右列字段标签 (分界线 x={center_x:.0f})"
                )
                
                # 分别对齐左列和右列的字段标签
                for column_labels, column_name in [(left_column_labels, "左列"), (right_column_labels, "右列")]:
                    if not column_labels:
                        continue
                    
                    # 找到该列中最左的边界
                    min_left_x = float('inf')
                    for region in column_labels:
                        left_x = region.bbox[0]
                        min_left_x = min(min_left_x, left_x)
                        logger.debug(f"{column_name}字段标签 '{region.text}' 左边界: {left_x}")
                    
                    logger.info(f"横版模式：{column_name}字段标签的最左边界: {min_left_x}")
                    
                    # 调整该列所有字段标签的左边界到最左边界
                    for region in column_labels:
                        x1, y1, x2, y2 = region.bbox
                        if x1 != min_left_x:
                            # 调整左边界，保持宽度不变
                            new_x2 = min_left_x + (x2 - x1)
                            adjusted_bbox = (int(min_left_x), y1, int(new_x2), y2)
                            region.bbox = adjusted_bbox
                            logger.info(
                                f"{column_name}字段标签 '{region.text}' 调整左边界: x1 {x1} -> {int(min_left_x)} "
                                f"(左移 {x1 - min_left_x:.0f}px)"
                            )
                        
                        # 设置unified_left_boundary属性，用于渲染时对齐
                        region.unified_left_boundary = min_left_x
            
            # 识别字段内容（字段标签右侧的文本）
            # 使用新的配对逻辑：每个字段标签只配对水平方向最近的一个字段内容
            field_content_regions = self._identify_field_contents_horizontal(
                normalized_regions, field_label_regions
            )
            
            # 【新增】补救步骤：查找那些没有被配对但实际上是字段内容的区域
            # 特别是经营范围等长文本字段的内容
            logger.info(f"开始补救配对，共 {len(field_label_regions)} 个字段标签")
            for label_region in field_label_regions:
                # 检查这个字段标签是否已经有配对的内容
                has_paired_content = any(
                    getattr(content, 'paired_field_label', None) == label_region
                    for content in field_content_regions
                )
                
                if has_paired_content:
                    logger.debug(f"字段标签 '{label_region.text}' 已有配对内容，跳过")
                    continue  # 已经有配对的内容，跳过
                
                logger.info(f"字段标签 '{label_region.text}' 没有配对内容，开始查找...")
                
                # 没有配对的内容，尝试查找
                label_x1, label_y1, label_x2, label_y2 = label_region.bbox
                label_center_y = (label_y1 + label_y2) / 2
                
                # 查找右侧或下方的区域
                for region in normalized_regions:
                    # 跳过字段标签
                    if getattr(region, 'is_field_label', False):
                        continue
                    
                    # 跳过已经配对的字段内容
                    if region in field_content_regions:
                        continue
                    
                    region_x1, region_y1, region_x2, region_y2 = region.bbox
                    region_center_y = (region_y1 + region_y2) / 2
                    
                    # 检查是否在右侧（x坐标接近）或下方（y坐标接近）
                    x_distance = region_x1 - label_x2
                    y_distance = region_y1 - label_y2
                    
                    # 条件1：在右侧且y坐标接近（同一行）
                    # 使用顶部y坐标差异而不是中心y坐标，因为高度很大的区域中心点会偏下很多
                    y_top_diff = abs(region_y1 - label_y1)
                    is_right_side = (-10 <= x_distance <= 50) and (y_top_diff < 30)
                    
                    # 条件2：在下方且x坐标接近（下一行）
                    is_below = (0 <= y_distance <= 20) and (abs(region_x1 - label_x1) < 50)
                    
                    # 调试输出
                    if label_region.text == "经营范围" and "一般项目" in region.text:
                        logger.info(
                            f"  检查候选区域: '{region.text[:30]}...' "
                            f"x_distance={x_distance:.0f}px, y_distance={y_distance:.0f}px, "
                            f"y_top_diff={y_top_diff:.0f}px, "
                            f"label_y1={label_y1}, region_y1={region_y1}, "
                            f"is_right_side={is_right_side}, is_below={is_below}"
                        )
                    
                    if is_right_side or is_below:
                        # 找到了！标记为字段内容
                        region.is_field_content = True
                        region.paired_field_label = label_region
                        field_content_regions.append(region)
                        
                        logger.info(
                            f"【补救配对】字段标签 '{label_region.text}' 找到字段内容: '{region.text[:30]}...' "
                            f"(x_distance={x_distance:.0f}px, y_top_diff={y_top_diff:.0f}px, "
                            f"{'右侧配对' if is_right_side else '下方配对'})"
                        )
                        break  # 只配对一个
            
            # 【新增】检查段落合并后的字段内容区域，调整与字段标签重叠的位置
            # 这些区域没有经过_identify_field_contents的重叠调整，需要在这里处理
            for region in normalized_regions:
                # 检查是否是字段内容（段落合并标记的）
                if not getattr(region, 'is_field_content', False):
                    continue
                
                # 检查是否属于某个字段
                belongs_to_field = getattr(region, 'belongs_to_field', None)
                if not belongs_to_field:
                    continue
                
                # 查找对应的字段标签
                field_label = None
                for label_region in field_label_regions:
                    if label_region.text == belongs_to_field:
                        field_label = label_region
                        break
                
                if not field_label:
                    continue
                
                # 检查是否与字段标签重叠
                label_x1, label_y1, label_x2, label_y2 = field_label.bbox
                content_x1, content_y1, content_x2, content_y2 = region.bbox
                
                if content_x1 < label_x2:
                    # 重叠了，需要往右移动
                    new_x1 = label_x2
                    original_width = content_x2 - content_x1
                    new_x2 = new_x1 + original_width  # 保持宽度不变
                    
                    # 检查调整后是否会与右侧的其他区域重叠
                    # 查找右侧同一行的其他区域
                    min_right_x = float('inf')
                    for other_region in normalized_regions:
                        if other_region == region:
                            continue
                        other_x1, other_y1, other_x2, other_y2 = other_region.bbox
                        # 检查是否在同一行（y坐标有重叠）
                        if not (other_y2 < content_y1 or other_y1 > content_y2):
                            # 在同一行，检查是否在右侧
                            if other_x1 > new_x1:
                                min_right_x = min(min_right_x, other_x1)
                    
                    # 如果右侧有其他区域，确保不会重叠（留出至少20px的间距）
                    if min_right_x != float('inf'):
                        max_allowed_x2 = min_right_x - 20  # 留出20px间距
                        if new_x2 > max_allowed_x2:
                            # 需要缩小宽度以避免重叠
                            new_x2 = max_allowed_x2
                            logger.info(
                                f"调整段落合并后的字段内容宽度: '{region.text[:30]}...' "
                                f"原宽度={original_width}px, 新宽度={new_x2 - new_x1}px "
                                f"(避免与右侧区域重叠，右侧x1={min_right_x})"
                            )
                    
                    # 创建新的bbox
                    aligned_bbox = (new_x1, content_y1, new_x2, content_y2)
                    region.bbox = aligned_bbox
                    
                    logger.info(
                        f"调整段落合并后的字段内容位置: '{region.text[:30]}...' "
                        f"属于字段='{belongs_to_field}', "
                        f"重叠调整: x1 {content_x1} -> {new_x1} (右移 {new_x1 - content_x1}px), "
                        f"bbox=({new_x1}, {content_y1}, {new_x2}, {content_y2})"
                    )
                    
                    # 添加到字段内容列表（如果还没有）
                    if region not in field_content_regions:
                        field_content_regions.append(region)
            
            # 为字段内容设置统一字体大小（基于中位数高度，排除异常值）
            if unified_font_enabled and field_content_regions:
                # 收集所有字段内容的高度
                content_heights = []
                for region in field_content_regions:
                    content_height = region.bbox[3] - region.bbox[1]
                    content_heights.append(content_height)
                    logger.debug(f"字段内容 '{region.text}' 高度: {content_height}px")
                
                # 排序后取中位数，避免被极端值（如经营范围的183px）影响
                content_heights.sort()
                median_height = content_heights[len(content_heights) // 2]
                
                # 计算统一字体大小：median_height × 0.6（提高系数以增大字体）
                unified_content_font_size = int(median_height * 0.6)
                
                # 限制在合理范围内 [10, 48]
                unified_content_font_size = max(10, min(48, unified_content_font_size))
                
                print(f"\n{'='*80}")
                print(f"[字段内容统一字体大小计算]")
                print(f"  字段内容数量: {len(field_content_regions)}")
                print(f"  所有高度: {content_heights}")
                print(f"  中位数高度: {median_height}px")
                print(f"  缩放系数: 0.85")
                print(f"  计算公式: {median_height} × 0.6 = {median_height * 0.6}")
                print(f"  统一字体大小: {unified_content_font_size}px (限制在 12-48px)")
                print(f"{'='*80}\n")
                
                logger.info(
                    f"横版模式：计算字段内容统一字体大小: median_height={median_height}px, "
                    f"heights={content_heights}, "
                    f"unified_content_font_size={unified_content_font_size}px (median_height × 0.6)"
                )
                
                # 为所有字段内容设置统一字体大小
                for region in field_content_regions:
                    original_font_size = region.font_size
                    region.font_size = unified_content_font_size
                    # 添加标记，表示使用了统一字体大小
                    region.uses_unified_font_size = True
                    
                    print(f"[设置统一字体] 字段内容 '{region.text[:30]}...'")
                    print(f"  原始 font_size: {original_font_size}px")
                    print(f"  新的 font_size: {unified_content_font_size}px")
                    print(f"  uses_unified_font_size: True\n")
                    
                    logger.info(
                        f"字段内容 '{region.text}' 字体大小: {original_font_size}px -> {unified_content_font_size}px (统一)"
                    )
            
            logger.info(f"横版模式：字段边界框标准化完成（翻译前）")
            return normalized_regions
        else:
            # 竖版模式：使用复杂的扩展和识别逻辑
            return self._normalize_field_regions_vertical(regions)
    
    def _normalize_field_regions_vertical(self, regions: List[TextRegion]) -> List[TextRegion]:
        """竖版模式：标准化字段边界框（扩展和识别逻辑）。
        
        竖版模式下，需要：
        1. 扩展不完整字段标签的左边界
        2. 识别字段内容（字段标签右侧的文本）
        3. 计算统一的左边界
        
        Args:
            regions: 文本区域列表（中文）
            
        Returns:
            标准化后的文本区域列表
        """
        logger.info(f"竖版模式：开始标准化字段边界框（翻译前），共 {len(regions)} 个区域")
        
        # 初始化保存被过滤区域的列表
        self._filtered_vertical_combined_regions = []
        
        # 步骤1：扩展不完整字段标签的左边界，并收集所有字段标签
        expanded_regions = []
        field_label_regions = []  # 存储所有字段标签的region
        
        # 字段标签列表（中文）
        field_labels = ["名称", "类型", "住所", "法定代表人", "注册资本", "成立日期", "营业期限", "经营范围"]
        
        # 横版营业执照特殊处理：左侧竖排字段可能被识别成组合字符
        # 例如："名称"+"类型" -> "名 类" 或 "名类"
        field_label_patterns = {
            "名称": ["名称", "名 称"],
            "类型": ["类型", "类 型"],
            "住所": ["住所", "住 所", "住"],  # "住"可能单独识别
            "法定代表人": ["法定代表人", "法定 代表人", "代表人"],
            "注册资本": ["注册资本", "注册 资本", "资本"],
            "成立日期": ["成立日期", "成立 日期", "立日期"],
            "营业期限": ["营业期限", "营业 期限", "业期限"],
            "经营范围": ["经营范围", "经营 范围", "营范围", "经营范 围"]
        }
        
        # 横版特殊模式：竖排字段组合（如"名 类"包含"名称"和"类型"）
        vertical_combined_patterns = {
            "名 类": ["名称", "类型"],
            "名类": ["名称", "类型"],
        }
        
        # 竖版特殊模式：多个字段标签被识别成一个区域（如"名称类型住所"）
        # 这种情况下，我们需要将其拆分成多个独立的字段标签
        multi_field_combined_patterns = {
            "名称类型住所": ["名称", "类型", "住所"],
            "名称类型": ["名称", "类型"],
            "类型住所": ["类型", "住所"],
        }
        
        for i, region in enumerate(regions):
            # 扩展不完整字段标签的左边界
            expanded_region = self._expand_incomplete_field_left_boundary(region)
            
            # 检查是否是字段标签（精确匹配）
            if region.text in field_labels:
                expanded_regions.append(expanded_region)
                field_label_regions.append(expanded_region)
                logger.info(f"识别到字段标签: '{region.text}'")
                continue
            
            # 检查是否是字段标签的变体（模糊匹配）
            matched_label = None
            for label, patterns in field_label_patterns.items():
                if region.text in patterns:
                    matched_label = label
                    break
            
            if matched_label:
                # 创建一个新的region,文本改为标准字段名
                normalized_region = TextRegion(
                    bbox=expanded_region.bbox,
                    text=matched_label,  # 使用标准字段名
                    confidence=expanded_region.confidence,
                    font_size=expanded_region.font_size,
                    angle=expanded_region.angle,
                    is_vertical_merged=getattr(expanded_region, 'is_vertical_merged', False)
                )
                # 保留原始属性
                if hasattr(expanded_region, 'original_x1'):
                    normalized_region.original_x1 = expanded_region.original_x1
                
                expanded_regions.append(normalized_region)
                field_label_regions.append(normalized_region)
                logger.info(f"识别到字段标签变体: '{region.text}' -> '{matched_label}'")
                continue
            
            # 检查是否是竖排组合字段（如"名 类"）
            if region.text in vertical_combined_patterns:
                combined_labels = vertical_combined_patterns[region.text]
                logger.info(f"识别到竖排组合字段: '{region.text}' 包含 {combined_labels}")
                
                # 为每个字段创建一个独立的region,并添加到 expanded_regions
                # 假设它们是竖排的,平均分配bbox的高度
                x1, y1, x2, y2 = expanded_region.bbox
                height_per_label = (y2 - y1) / len(combined_labels)
                
                for idx, label in enumerate(combined_labels):
                    label_y1 = y1 + idx * height_per_label
                    label_y2 = label_y1 + height_per_label
                    
                    split_region = TextRegion(
                        bbox=(x1, int(label_y1), x2, int(label_y2)),
                        text=label,
                        confidence=expanded_region.confidence,
                        font_size=expanded_region.font_size,
                        angle=expanded_region.angle,
                        is_vertical_merged=False
                    )
                    # 添加到 expanded_regions,这样它会被翻译
                    expanded_regions.append(split_region)
                    # 同时添加到 field_label_regions,用于计算统一左边界
                    field_label_regions.append(split_region)
                    logger.info(f"  拆分字段标签: '{label}' at ({x1}, {int(label_y1)}) -> ({x2}, {int(label_y2)})")
                
                # 保存原始的组合字段,用于后续抹除
                self._filtered_vertical_combined_regions.append(expanded_region)
                
                # 不要将原始的组合字段添加到 expanded_regions
                continue  # 跳过添加到 expanded_regions
            
            # 检查是否是多字段组合（如"名称类型住所"）
            if region.text in multi_field_combined_patterns:
                combined_labels = multi_field_combined_patterns[region.text]
                logger.info(f"识别到多字段组合: '{region.text}' 包含 {combined_labels}")
                
                # 为每个字段创建一个独立的region
                # 假设它们是竖排的,平均分配bbox的高度
                x1, y1, x2, y2 = expanded_region.bbox
                height_per_label = (y2 - y1) / len(combined_labels)
                
                for idx, label in enumerate(combined_labels):
                    label_y1 = y1 + idx * height_per_label
                    label_y2 = label_y1 + height_per_label
                    
                    split_region = TextRegion(
                        bbox=(x1, int(label_y1), x2, int(label_y2)),
                        text=label,
                        confidence=expanded_region.confidence,
                        font_size=expanded_region.font_size,
                        angle=expanded_region.angle,
                        is_vertical_merged=False
                    )
                    # 添加到 expanded_regions,这样它会被翻译
                    expanded_regions.append(split_region)
                    # 同时添加到 field_label_regions,用于计算统一左边界
                    field_label_regions.append(split_region)
                    logger.info(f"  拆分字段标签: '{label}' at ({x1}, {int(label_y1)}) -> ({x2}, {int(label_y2)})")
                
                # 保存原始的组合字段,用于后续抹除
                self._filtered_vertical_combined_regions.append(expanded_region)
                
                # 不要将原始的组合字段添加到 expanded_regions
                continue  # 跳过添加到 expanded_regions
            
            # 默认情况：不是字段标签的普通区域，直接添加到 expanded_regions
            expanded_regions.append(expanded_region)
        
        # 步骤2：计算所有字段标签的最左边界（使用original_x1或bbox的x1）
        if field_label_regions:
            min_left_boundary = float('inf')
            for region in field_label_regions:
                # 优先使用original_x1（如果存在），否则使用bbox的x1
                left_x = getattr(region, 'original_x1', region.bbox[0])
                min_left_boundary = min(min_left_boundary, left_x)
                logger.info(f"字段 '{region.text}' 左边界: {left_x}")
            
            logger.info(f"所有字段标签的最左边界: {min_left_boundary}")
            
            # 保存字段标签区域，用于后续抹除
            self._field_label_regions = field_label_regions
            logger.info(f"保存 {len(field_label_regions)} 个字段标签区域，用于后续抹除")
            
            # 步骤3：为所有字段标签设置统一的左边界
            for region in field_label_regions:
                region.unified_left_boundary = min_left_boundary
                # 标记为字段标签（不是字段内容）
                region.is_field_label = True
                logger.info(
                    f"字段标签 '{region.text}' 设置统一左边界: {min_left_boundary} "
                    f"(原始={getattr(region, 'original_x1', region.bbox[0])}), is_field_label=True"
                )
            
            # 步骤3.1：计算并设置统一字体大小（基于最大高度）
            # 检查配置是否启用统一字体大小功能
            unified_font_enabled = self.config.get('rendering.unified_field_label_font_size.enabled', True)
            
            if unified_font_enabled:
                # 找到所有字段标签中最大的高度
                max_field_height = 0
                for region in field_label_regions:
                    field_height = region.bbox[3] - region.bbox[1]
                    max_field_height = max(max_field_height, field_height)
                    logger.debug(f"字段 '{region.text}' 高度: {field_height}px")
                
                # 计算统一字体大小：max_height × 0.8
                unified_font_size = int(max_field_height * 0.5)
                
                # 限制在合理范围内 [8, 48]
                unified_font_size = max(8, min(48, unified_font_size))
                
                logger.info(
                    f"计算统一字体大小: max_height={max_field_height}px, "
                    f"unified_font_size={unified_font_size}px (max_height × 0.8)"
                )
                
                # 为所有字段标签设置统一字体大小
                for region in field_label_regions:
                    original_font_size = region.font_size
                    region.font_size = unified_font_size
                    # 添加标记，表示使用了统一字体大小
                    region.uses_unified_font_size = True
                    logger.info(
                        f"字段标签 '{region.text}' 字体大小: {original_font_size}px -> {unified_font_size}px (统一)"
                    )
            else:
                logger.info("统一字体大小功能已禁用（配置: rendering.unified_field_label_font_size.enabled=false）")
            
            # 步骤3.5：调整字段标签的bbox，使其从统一左边界开始
            # 这样翻译后的字段标签会从统一左边界开始渲染
            for region in field_label_regions:
                x1, y1, x2, y2 = region.bbox
                # 调整bbox的左边界到统一左边界
                adjusted_bbox = (int(min_left_boundary), y1, x2, y2)
                region.bbox = adjusted_bbox
                logger.info(
                    f"字段标签 '{region.text}' 调整bbox: ({x1},{y1},{x2},{y2}) -> ({int(min_left_boundary)},{y1},{x2},{y2})"
                )
        
        # 步骤4：识别字段内容（字段标签右侧的文本）
        field_content_regions = self._identify_field_contents(
            expanded_regions, field_label_regions
        )
        
        # 保存字段内容区域列表，供后续重叠检测使用
        self._field_content_regions = field_content_regions
        
        # 步骤5：计算所有字段内容的统一左边界
        if field_content_regions:
            min_content_left = float('inf')
            for region in field_content_regions:
                content_left = region.bbox[0]
                min_content_left = min(min_content_left, content_left)
                logger.info(f"字段内容 '{region.text}' 左边界: {content_left}")
            
            logger.info(f"所有字段内容的最左边界: {min_content_left}")
            
            # 为所有字段内容设置统一的左边界
            for region in field_content_regions:
                region.unified_content_left_boundary = min_content_left
                # 确保is_field_content标记被设置
                region.is_field_content = True
                logger.info(
                    f"字段内容 '{region.text}' 设置统一左边界: {min_content_left} "
                    f"(原始={region.bbox[0]}), is_field_content=True"
                )
            
            # 步骤5.5：计算字段内容的统一字体大小（基于中位数高度）
            # 收集所有字段内容的高度
            content_heights = []
            for region in field_content_regions:
                content_height = region.bbox[3] - region.bbox[1]
                content_heights.append(content_height)
                logger.debug(f"字段内容 '{region.text}' 高度: {content_height}px")
            
            # 排序后取中位数，避免被极端值（如经营范围的长文本）影响
            content_heights.sort()
            median_height = content_heights[len(content_heights) // 2]
            
            # 计算统一字体大小：median_height × 0.6（与横版保持一致）
            unified_content_font_size = int(median_height * 0.6)
            
            # 限制在合理范围内 [10, 48]
            unified_content_font_size = max(10, min(48, unified_content_font_size))
            
            logger.info(
                f"竖版模式：计算字段内容统一字体大小: median_height={median_height}px, "
                f"heights={content_heights}, "
                f"unified_content_font_size={unified_content_font_size}px (median_height × 0.6)"
            )
            
            # 为所有字段内容设置统一字体大小
            for region in field_content_regions:
                original_font_size = region.font_size
                region.font_size = unified_content_font_size
                # 添加标记，表示使用了统一字体大小（渲染时不会因换行而缩小）
                region.uses_unified_font_size = True
                
                logger.info(
                    f"字段内容 '{region.text}' 字体大小: {original_font_size}px -> {unified_content_font_size}px (统一)"
                )
        
        # 步骤6：使用中文字段名进行标准化（右边界扩展）
        normalized_regions = []
        for expanded_region in expanded_regions:
            # 记录标准化前的属性
            has_unified_before = hasattr(expanded_region, 'unified_content_left_boundary')
            if has_unified_before:
                logger.info(f"[NORMALIZE] 标准化前: '{expanded_region.text}' 有 unified_content_left_boundary={expanded_region.unified_content_left_boundary}")
            
            normalized_region = self.text_renderer._normalize_field_bbox(expanded_region, expanded_regions)
            
            # 记录标准化后的属性
            has_unified_after = hasattr(normalized_region, 'unified_content_left_boundary')
            if has_unified_before and not has_unified_after:
                logger.error(f"[NORMALIZE] 标准化后丢失: '{normalized_region.text}' 丢失了 unified_content_left_boundary!")
            elif has_unified_after:
                logger.info(f"[NORMALIZE] 标准化后保留: '{normalized_region.text}' 有 unified_content_left_boundary={normalized_region.unified_content_left_boundary}")
            
            normalized_regions.append(normalized_region)
            
            # 如果bbox变化了，打印信息
            if (normalized_region.bbox != expanded_region.bbox):
                logger.info(
                    f"标准化字段: '{expanded_region.text}' - "
                    f"bbox=({expanded_region.bbox[0]},{expanded_region.bbox[1]},{expanded_region.bbox[2]},{expanded_region.bbox[3]}) -> "
                    f"({normalized_region.bbox[0]},{normalized_region.bbox[1]},{normalized_region.bbox[2]},{normalized_region.bbox[3]})"
                )
        
        logger.info(f"竖版模式：字段边界框标准化完成（翻译前）")
        logger.info(f"标准化后共 {len(normalized_regions)} 个区域")
        if self._filtered_vertical_combined_regions:
            logger.info(f"保存 {len(self._filtered_vertical_combined_regions)} 个被过滤的竖排组合字段，用于后续抹除")
        
        return normalized_regions
    
    def _expand_incomplete_field_left_boundary(self, region: TextRegion) -> TextRegion:
        """扩展不完整字段标签的左边界，包含缺失的字符。
        
        策略：
        1. 识别不完整的字段标签（如"法定代表人"、"注册资本"、"成立日期"等）
        2. 根据完整字段名估算缺失字符的数量
        3. 根据字体大小估算缺失字符的宽度
        4. 向左扩展bbox的左边界
        5. 保存原始的左边界，用于渲染时的左对齐
        
        注意：这里处理的是字段碎片合并后的字段，虽然文本已经是完整的（如"法定代表人"），
        但bbox可能只包含部分字符（如只包含"代表人"的位置，缺少"法定"的位置）。
        
        Args:
            region: 原始文本区域
            
        Returns:
            扩展后的TextRegion（如果不需要扩展则返回原region）
        """
        # 完整字段标签映射：完整文本 -> 可能缺失的字符数
        # 这些字段在OCR识别时，左侧的字符可能被遗漏
        incomplete_field_mapping = {
            "法定代表人": 2,      # 可能缺少"法定"2个字（只识别到"代表人"）
            "注册资本": 2,        # 可能缺少"注册"2个字（只识别到"资本"）
            "成立日期": 2,        # 可能缺少"成立"2个字（只识别到"日期"）
            "营业期限": 2,        # 可能缺少"营业"2个字（只识别到"期限"）
            "经营范围": 2,        # 可能缺少"经营"2个字（只识别到"范围"）
        }
        
        # 检查当前文本是否是可能不完整的字段标签
        if region.text not in incomplete_field_mapping:
            return region
        
        missing_char_count = incomplete_field_mapping[region.text]
        
        # 估算缺失字符的宽度
        # 假设每个中文字符的宽度约等于字体大小
        estimated_char_width = region.font_size
        missing_width = missing_char_count * estimated_char_width
        
        # 向左扩展bbox
        x1, y1, x2, y2 = region.bbox
        new_x1 = max(0, x1 - missing_width)  # 确保不超出图片边界
        
        logger.info(
            f"[EXPAND_LEFT] 扩展字段左边界: '{region.text}' "
            f"(可能缺失{missing_char_count}字, 估算宽度={missing_width}px), "
            f"x1: {x1} -> {new_x1}, 保存原始x1={x1}用于左对齐"
        )
        
        # 创建扩展后的TextRegion
        # 注意：text保持不变（仍然是合并后的完整文本），只扩展bbox
        # 同时保存原始的左边界，用于渲染时的左对齐
        expanded_region = TextRegion(
            bbox=(new_x1, y1, x2, y2),
            text=region.text,  # 保持合并后的完整文本
            confidence=region.confidence,
            font_size=region.font_size,
            angle=region.angle,
            is_vertical_merged=getattr(region, 'is_vertical_merged', False)
        )
        
        # 保存原始的左边界，用于渲染时的左对齐
        # 这样可以确保文本从原始位置开始渲染，而不是从扩展后的位置
        expanded_region.original_x1 = x1
        
        # 保留is_paragraph_merged属性（如果原region有的话）
        if hasattr(region, 'is_paragraph_merged'):
            expanded_region.is_paragraph_merged = region.is_paragraph_merged
        
        # 保留is_field_content属性（如果原region有的话）
        if hasattr(region, 'is_field_content'):
            expanded_region.is_field_content = region.is_field_content
        
        # 保留belongs_to_field属性（如果原region有的话）
        if hasattr(region, 'belongs_to_field'):
            expanded_region.belongs_to_field = region.belongs_to_field
        
        return expanded_region
    
    def _identify_field_contents(
        self,
        all_regions: List[TextRegion],
        field_label_regions: List[TextRegion]
    ) -> List[TextRegion]:
        """识别字段标签右侧的文本（字段内容）。
        
        策略：
        1. 第一步：找到与字段标签同一行的第一个字段内容
        2. 第二步：从第一个字段内容开始，向下查找垂直连续的区域
        3. 一个字段标签可以对应多个垂直连续的字段内容（如多行的经营范围）
        4. 字段内容的右边界与字段标签的标准化右边界对齐
        
        Args:
            all_regions: 所有文本区域列表
            field_label_regions: 字段标签区域列表
            
        Returns:
            字段内容区域列表
        """
        field_content_regions = []
        
        # 获取所有字段标签的文本
        field_label_texts = {region.text for region in field_label_regions}
        
        # 用于记录每个字段标签配对的字段内容（允许多个）
        label_to_contents = {id(label): [] for label in field_label_regions}
        
        # 用于记录已经被配对的区域，避免重复配对
        paired_regions = set()
        
        # 第一步：对每个字段标签，找到同一行的第一个字段内容
        for label_region in field_label_regions:
            # 获取字段标签的初始bbox（标准化前的bbox）
            if hasattr(label_region, 'original_bbox'):
                original_x1, original_y1, original_x2, original_y2 = label_region.original_bbox
                original_center_y = (original_y1 + original_y2) / 2
            else:
                # 如果没有original_bbox，使用当前bbox
                label_x1, label_y1, label_x2, label_y2 = label_region.bbox
                original_x1 = label_x1  # 添加这一行
                original_x2 = label_x2
                original_center_y = (label_y1 + label_y2) / 2
            
            # 找到右侧距离最近的或与字段标签重叠的字段内容
            min_distance = float('inf')
            first_content = None
            
            for region in all_regions:
                # 跳过字段标签本身
                if region.text in field_label_texts:
                    continue
                
                # 跳过已经被配对的区域
                if id(region) in paired_regions:
                    continue
                
                # 跳过已经被段落合并处理过的区域（is_field_content=True 或 is_paragraph_merged=True）
                if getattr(region, 'is_field_content', False) or getattr(region, 'is_paragraph_merged', False):
                    logger.debug(
                        f"跳过已合并的区域: '{region.text[:30]}...' "
                        f"(is_field_content={getattr(region, 'is_field_content', False)}, "
                        f"is_paragraph_merged={getattr(region, 'is_paragraph_merged', False)})"
                    )
                    continue
                
                region_x1, region_y1, region_x2, region_y2 = region.bbox
                region_center_y = (region_y1 + region_y2) / 2
                
                # 判断是否在同一水平线：使用初始bbox的y坐标
                y_diff = abs(region_center_y - original_center_y)
                
                # 只考虑y坐标接近的区域（y_diff < 10px）
                if y_diff >= 10:
                    continue
                
                # 计算水平距离（使用初始bbox的x2来判断距离）
                x_distance = region_x1 - original_x2
                
                # 新增：检查是否与字段标签重叠
                # 重叠判断：字段内容的左边界 < 字段标签的右边界 AND 字段内容的右边界 > 字段标签的左边界
                is_overlapping = (region_x1 < original_x2 and region_x2 > original_x1)
                
                # 接受两种情况：
                # 1. 在右侧或紧贴（x_distance >= 0）
                # 2. 与字段标签重叠（is_overlapping = True）
                if x_distance < 0 and not is_overlapping:
                    continue
                
                # 计算距离度量：
                # - 如果重叠，距离为0（优先选择重叠的）
                # - 如果不重叠，距离为x_distance
                distance_metric = 0 if is_overlapping else x_distance
                
                # 选择距离最近的（重叠的优先，因为距离为0）
                if distance_metric < min_distance:
                    min_distance = distance_metric
                    first_content = region
                    
                    if is_overlapping:
                        logger.debug(
                            f"字段标签 '{label_region.text}' 找到重叠的字段内容: '{region.text[:30]}...' "
                            f"(label_x: {original_x1}-{original_x2}, content_x: {region_x1}-{region_x2})"
                        )
            
            # 如果找到了第一个字段内容
            if first_content is not None:
                # 添加第一个字段内容
                label_to_contents[id(label_region)].append({
                    'region': first_content,
                    'distance': min_distance,
                    'original_x2': original_x2
                })
                paired_regions.add(id(first_content))
                
                logger.debug(
                    f"字段标签 '{label_region.text}' 找到第一个字段内容: '{first_content.text[:30]}...'"
                )
                
                # 第二步：从第一个字段内容开始，向下查找垂直连续的区域
                current_content = first_content
                max_vertical_gap = 10  # 最大垂直间隙（像素）
                max_vertical_overlap = 10  # 允许的最大垂直重叠（像素）
                
                # 动态调整x坐标差异阈值：
                # - 对于经营范围等长文本字段，使用更宽松的阈值（100px）
                # - 对于其他字段，使用标准阈值（50px）
                is_long_text_field = label_region.text in ["经营范围", "Business Scope", "住所", "Address"]
                max_x_diff = 100 if is_long_text_field else 50  # 最大x坐标差异（像素）
                
                if is_long_text_field:
                    logger.info(
                        f"字段 '{label_region.text}' 是长文本字段，使用宽松的x坐标阈值: {max_x_diff}px"
                    )
                
                while True:
                    # 查找当前内容下方的连续区域
                    current_x1, current_y1, current_x2, current_y2 = current_content.bbox
                    
                    next_content = None
                    min_y_gap = float('inf')
                    
                    for region in all_regions:
                        # 跳过字段标签本身
                        if region.text in field_label_texts:
                            continue
                        
                        # 跳过已经被配对的区域
                        if id(region) in paired_regions:
                            continue
                        
                        # 跳过已经被段落合并处理过的区域
                        if getattr(region, 'is_field_content', False) or getattr(region, 'is_paragraph_merged', False):
                            continue
                        
                        region_x1, region_y1, region_x2, region_y2 = region.bbox
                        
                        # 检查是否在当前内容下方（允许轻微重叠）
                        y_gap = region_y1 - current_y2
                        
                        # 允许轻微重叠（y_gap为负数但绝对值小于max_vertical_overlap）
                        # 或者间隙不超过max_vertical_gap
                        if y_gap < -max_vertical_overlap or y_gap > max_vertical_gap:
                            continue
                        
                        # 检查x坐标是否接近（左边界接近）
                        x_diff = abs(region_x1 - current_x1)
                        if x_diff > max_x_diff:
                            continue
                        
                        # 找到最近的下方区域（使用绝对值比较）
                        abs_y_gap = abs(y_gap)
                        if abs_y_gap < abs(min_y_gap):
                            min_y_gap = y_gap  # 保留原始值（可能是负数）
                            next_content = region
                    
                    # 如果找到了下一个连续区域
                    if next_content is not None:
                        label_to_contents[id(label_region)].append({
                            'region': next_content,
                            'distance': min_distance,  # 使用第一个内容的距离
                            'original_x2': original_x2
                        })
                        paired_regions.add(id(next_content))
                        
                        logger.info(
                            f"字段标签 '{label_region.text}' 找到连续字段内容: '{next_content.text[:30]}...' "
                            f"(y_gap={min_y_gap:.1f}px)"
                        )
                        
                        # 继续向下查找
                        current_content = next_content
                    else:
                        # 没有找到更多连续区域，结束
                        break
        
        # 第三步：处理每个字段标签的内容
        for label_region in field_label_regions:
            contents = label_to_contents[id(label_region)]
            
            if not contents:
                # 没有找到字段内容
                if hasattr(label_region, 'original_bbox'):
                    original_x2 = label_region.original_bbox[2]
                    original_center_y = (label_region.original_bbox[1] + label_region.original_bbox[3]) / 2
                else:
                    original_x2 = label_region.bbox[2]
                    original_center_y = (label_region.bbox[1] + label_region.bbox[3]) / 2
                
                logger.warning(
                    f"未找到字段内容: 字段标签='{label_region.text}' "
                    f"(original_x2={original_x2}, original_center_y={original_center_y:.1f})"
                )
                continue
            
            # 获取字段标签的标准化bbox
            label_x1, label_y1, label_x2, label_y2 = label_region.bbox
            
            # 处理每个字段内容
            for i, content_info in enumerate(contents):
                region = content_info['region']
                min_distance = content_info['distance']
                original_x2 = content_info['original_x2']
                
                # 标记为字段内容
                region.is_field_content = True
                
                # 记录对应的字段标签（用于后续对齐）
                region.paired_field_label = label_region
                
                # 获取字段内容的原始bbox
                content_x1, content_y1, content_x2, content_y2 = region.bbox
                
                # 检查是否与字段标签重叠
                is_overlapping = content_x1 < label_x2 and content_x2 > label_x1
                
                # 检查是否重叠：如果字段内容的左边界 < 字段标签的右边界，说明重叠
                if content_x1 < label_x2:
                    # 重叠了，需要往右移动
                    # 新的左边界 = 字段标签的右边界
                    new_x1 = label_x2
                    new_x2 = new_x1 + (content_x2 - content_x1)  # 保持宽度不变
                    
                    # 创建新的bbox（往右移动到不重叠）
                    aligned_bbox = (new_x1, content_y1, new_x2, content_y2)
                    region.bbox = aligned_bbox
                    
                    logger.info(
                        f"配对字段内容 #{i+1}/{len(contents)}: '{region.text[:20]}...' "
                        f"<- 字段标签='{label_region.text}', "
                        f"x_distance={min_distance:.0f}px (from original_x2={original_x2}), "
                        f"{'重叠配对' if is_overlapping else '右侧配对'}，"
                        f"重叠调整: x1 {content_x1} -> {new_x1} (右移 {new_x1 - content_x1}px)"
                    )
                else:
                    # 没有重叠，保持原位置
                    logger.info(
                        f"配对字段内容 #{i+1}/{len(contents)}: '{region.text[:20]}...' "
                        f"<- 字段标签='{label_region.text}', "
                        f"x_distance={min_distance:.0f}px (from original_x2={original_x2}), "
                        f"{'重叠配对' if is_overlapping else '右侧配对'}，"
                        f"无重叠，保持原位置"
                    )
                
                # 添加到字段内容列表
                field_content_regions.append(region)
        
        # 【补救配对】查找那些没有被配对但实际上是字段内容的区域
        # 特别是经营范围等长文本字段的内容（可能不在同一水平线上）
        logger.info(f"竖版模式：开始补救配对，共 {len(field_label_regions)} 个字段标签")
        for label_region in field_label_regions:
            # 检查这个字段标签是否已经有配对的内容
            has_paired_content = any(
                getattr(content, 'paired_field_label', None) == label_region
                for content in field_content_regions
            )
            
            if has_paired_content:
                logger.debug(f"字段标签 '{label_region.text}' 已有配对内容，跳过")
                continue  # 已经有配对的内容，跳过
            
            logger.info(f"竖版模式：字段标签 '{label_region.text}' 没有配对内容，开始查找...")
            
            # 没有配对的内容，尝试查找
            # 获取字段标签的初始bbox（标准化前的bbox）
            if hasattr(label_region, 'original_bbox'):
                original_x1, original_y1, original_x2, original_y2 = label_region.original_bbox
            else:
                label_x1, label_y1, label_x2, label_y2 = label_region.bbox
                original_x1, original_y1, original_x2, original_y2 = label_x1, label_y1, label_x2, label_y2
            
            # 查找右侧或下方的区域
            for region in all_regions:
                # 跳过字段标签
                if region.text in field_label_texts:
                    continue
                
                # 跳过已经配对的字段内容
                if region in field_content_regions:
                    continue
                
                # 跳过已经被段落合并处理过的区域
                if getattr(region, 'is_field_content', False) or getattr(region, 'is_paragraph_merged', False):
                    continue
                
                region_x1, region_y1, region_x2, region_y2 = region.bbox
                
                # 检查是否在右侧（x坐标接近）或下方（y坐标接近）
                x_distance = region_x1 - original_x2
                y_distance = region_y1 - original_y2
                
                # 条件1：在右侧且y坐标接近（同一行或稍微偏下）
                # 使用顶部y坐标差异而不是中心y坐标，因为高度很大的区域中心点会偏下很多
                y_top_diff = abs(region_y1 - original_y1)
                is_right_side = (-10 <= x_distance <= 50) and (y_top_diff < 50)  # 竖版使用更宽松的y阈值
                
                # 条件2：在下方且x坐标接近（下一行）
                is_below = (0 <= y_distance <= 30) and (abs(region_x1 - original_x1) < 50)
                
                if is_right_side or is_below:
                    # 找到了！标记为字段内容
                    region.is_field_content = True
                    region.paired_field_label = label_region
                    field_content_regions.append(region)
                    
                    logger.info(
                        f"竖版模式：【补救配对】字段标签 '{label_region.text}' 找到字段内容: '{region.text[:30]}...' "
                        f"(x_distance={x_distance:.0f}px, y_top_diff={y_top_diff:.0f}px, "
                        f"{'右侧配对' if is_right_side else '下方配对'})"
                    )
                    break  # 只配对一个
        
        return field_content_regions
    
    def _identify_field_contents_horizontal(
        self,
        all_regions: List[TextRegion],
        field_label_regions: List[TextRegion]
    ) -> List[TextRegion]:
        """横版模式：识别字段标签右侧的文本（字段内容）。
        
        策略：
        1. 对于每个字段标签，找到水平方向最近的一个文本框作为字段内容
        2. 一个字段标签只对应一个字段内容
        3. 字段内容的右边界与字段标签的标准化右边界对齐
        
        Args:
            all_regions: 所有文本区域列表
            field_label_regions: 字段标签区域列表
            
        Returns:
            字段内容区域列表
        """
        # 横版模式使用相同的配对逻辑
        return self._identify_field_contents(all_regions, field_label_regions)
    
    def _detect_all_seals(self, image: np.ndarray) -> List[Tuple[int, int, int, int]]:
        """检测图片中的所有印章，使用多颜色检测和智能筛选。
        
        改进策略（适用多种场景）：
        1. 检测黑白图片：如果是黑白图片，使用边缘检测+轮廓检测
        2. 多颜色检测：红色、蓝色、黑色印章
        3. 自适应HSV阈值：适应不同亮度和饱和度
        4. 形态学操作：连接断裂的笔画
        5. 智能位置判断：根据图片布局和印章特征区分国徽和印章
        6. 多维度筛选：面积、形状、位置、颜色分布
        
        参数：
            image: 输入图片（BGR格式）
            
        返回：
            印章边界框列表 [(x1, y1, x2, y2), ...]
        """
        height, width = image.shape[:2]
        hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)
        
        # 检查是否是黑白图片（平均饱和度 < 20）
        saturation = hsv[:, :, 1]
        avg_saturation = np.mean(saturation)
        is_grayscale = avg_saturation < 20
        
        if is_grayscale:
            logger.info(f"🔍 检测到黑白图片（平均饱和度={avg_saturation:.2f}），使用边缘检测方法")
            return self._detect_seals_in_grayscale(image)
        
        # 读取配置
        detect_red = self.config.get('rendering.seal_detection.detect_red', True)
        detect_blue = self.config.get('rendering.seal_detection.detect_blue', True)
        detect_black = self.config.get('rendering.seal_detection.detect_black', False)
        red_sat_min = self.config.get('rendering.seal_detection.red_saturation_min', 40)
        red_val_min = self.config.get('rendering.seal_detection.red_value_min', 40)
        blue_sat_min = self.config.get('rendering.seal_detection.blue_saturation_min', 40)
        blue_val_min = self.config.get('rendering.seal_detection.blue_value_min', 40)
        black_sat_max = self.config.get('rendering.seal_detection.black_saturation_max', 50)
        black_val_max = self.config.get('rendering.seal_detection.black_value_max', 80)
        
        # 1. 多颜色检测
        masks = []
        
        if detect_red:
            # 红色检测（两个范围）
            lower_red1 = np.array([0, red_sat_min, red_val_min])
            upper_red1 = np.array([10, 255, 255])
            mask_red1 = cv2.inRange(hsv, lower_red1, upper_red1)
            
            lower_red2 = np.array([170, red_sat_min, red_val_min])
            upper_red2 = np.array([180, 255, 255])
            mask_red2 = cv2.inRange(hsv, lower_red2, upper_red2)
            
            mask_red = cv2.bitwise_or(mask_red1, mask_red2)
            masks.append(mask_red)
        
        if detect_blue:
            # 蓝色检测
            lower_blue = np.array([100, blue_sat_min, blue_val_min])
            upper_blue = np.array([130, 255, 255])
            mask_blue = cv2.inRange(hsv, lower_blue, upper_blue)
            masks.append(mask_blue)
        
        if detect_black:
            # 黑色检测（低饱和度、中低亮度）
            # 避免检测纯黑背景和普通黑色文字
            black_val_min = self.config.get('rendering.seal_detection.black_value_min', 20)
            lower_black = np.array([0, 0, black_val_min])
            upper_black = np.array([180, black_sat_max, black_val_max])
            mask_black = cv2.inRange(hsv, lower_black, upper_black)
            masks.append(mask_black)
        
        if not masks:
            logger.debug("未启用任何颜色检测")
            return []
        
        # 合并所有颜色mask
        combined_mask = masks[0]
        for mask in masks[1:]:
            combined_mask = cv2.bitwise_or(combined_mask, mask)
        
        # 2. 形态学操作：闭运算连接断裂的笔画
        # 使用中等大小的核，平衡印章完整性和避免与文字合并
        kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (4, 4))
        combined_mask = cv2.morphologyEx(combined_mask, cv2.MORPH_CLOSE, kernel)
        # 轻微膨胀，确保印章轮廓完整
        combined_mask = cv2.dilate(combined_mask, kernel, iterations=1)
        
        # 3. 找到轮廓
        contours, _ = cv2.findContours(combined_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        if not contours:
            logger.debug("未检测到印章候选区域")
            return []
        
        logger.info(f"🔍 检测到 {len(contours)} 个候选轮廓")
        
        # 读取形状筛选配置
        min_seal_area = self.config.get('rendering.seal_detection.min_seal_area', 1000)
        max_seal_area = self.config.get('rendering.seal_detection.max_seal_area', 100000)
        min_circularity = self.config.get('rendering.seal_detection.min_circularity', 0.25)
        min_aspect_ratio = self.config.get('rendering.seal_detection.min_aspect_ratio', 0.3)
        max_aspect_ratio = self.config.get('rendering.seal_detection.max_aspect_ratio', 3.0)
        expand_ratio = self.config.get('rendering.seal_detection.expand_radius_ratio', 1.5)
        
        # 读取国徽识别配置
        emblem_enabled = self.config.get('rendering.seal_detection.emblem_detection.enabled', True)
        emblem_top_ratio = self.config.get('rendering.seal_detection.emblem_detection.top_region_ratio', 0.33)
        emblem_h_center_ratio = self.config.get('rendering.seal_detection.emblem_detection.horizontal_center_ratio', 0.33)
        emblem_min_area_ratio = self.config.get('rendering.seal_detection.emblem_detection.min_area_ratio', 0.01)
        emblem_min_circularity = self.config.get('rendering.seal_detection.emblem_detection.min_circularity', 0.7)
        
        seals = []
        
        # 统计过滤原因
        filtered_stats = {
            'too_small': 0,
            'too_large': 0,
            'bad_aspect_ratio': 0,
            'low_circularity': 0,
            'emblem': 0
        }
        
        for contour in contours:
            contour_area = cv2.contourArea(contour)
            
            # 过滤太小或太大的轮廓
            if contour_area < min_seal_area:
                filtered_stats['too_small'] += 1
                continue
            if contour_area > max_seal_area:
                filtered_stats['too_large'] += 1
                logger.debug(f"跳过过大的轮廓: 面积={contour_area:.0f} > {max_seal_area}")
                continue
            
            # 获取边界框
            seal_x, seal_y, seal_w, seal_h = cv2.boundingRect(contour)
            seal_center_x = seal_x + seal_w // 2
            seal_center_y = seal_y + seal_h // 2
            
            # 检查形状：印章通常是接近正方形或圆形的
            aspect_ratio = seal_w / seal_h if seal_h > 0 else 0
            if not (min_aspect_ratio <= aspect_ratio <= max_aspect_ratio):
                filtered_stats['bad_aspect_ratio'] += 1
                continue
            
            # 计算圆度
            perimeter = cv2.arcLength(contour, True)
            if perimeter > 0:
                circularity = 4 * np.pi * contour_area / (perimeter * perimeter)
                if circularity < min_circularity:
                    filtered_stats['low_circularity'] += 1
                    continue
            else:
                filtered_stats['low_circularity'] += 1
                continue
            
            # 4. 智能位置判断：区分国徽和印章
            if emblem_enabled:
                # 判断是否在上方区域
                in_top_region = seal_center_y < height * emblem_top_ratio
                
                # 判断是否在水平中央位置
                h_center_left = width * emblem_h_center_ratio
                h_center_right = width * (1 - emblem_h_center_ratio)
                in_horizontal_center = h_center_left < seal_center_x < h_center_right
                
                # 判断面积是否较大
                is_large = contour_area > (height * width * emblem_min_area_ratio)
                
                # 判断圆度是否很高
                is_very_circular = circularity > emblem_min_circularity
                
                # 综合判断：如果在上方中央、面积大、圆度高，很可能是国徽
                if in_top_region and in_horizontal_center and is_large and is_very_circular:
                    filtered_stats['emblem'] += 1
                    logger.info(f"🔴 跳过疑似国徽: 位置=({seal_x}, {seal_y}), "
                               f"中心=({seal_center_x}, {seal_center_y}), "
                               f"面积={contour_area:.0f}, 圆度={circularity:.2f}")
                    continue
            
            # 5. 使用最小外接圆来获取印章的完整区域
            (center_x, center_y), radius = cv2.minEnclosingCircle(contour)
            
            # 扩大半径以确保包含完整的印章
            expanded_radius = int(radius * expand_ratio)
            
            x1_expanded = max(0, int(center_x - expanded_radius))
            y1_expanded = max(0, int(center_y - expanded_radius))
            x2_expanded = min(width, int(center_x + expanded_radius))
            y2_expanded = min(height, int(center_y + expanded_radius))
            
            seals.append((x1_expanded, y1_expanded, x2_expanded, y2_expanded))
            logger.info(f"🔴 检测到印章: 圆心=({int(center_x)}, {int(center_y)}), 半径={int(radius)}, "
                       f"面积={contour_area:.0f}, 圆度={circularity:.2f}, "
                       f"位置={'上方' if seal_center_y < height * emblem_top_ratio else '下方'}, "
                       f"扩大后边界=({x1_expanded}, {y1_expanded}) -> ({x2_expanded}, {y2_expanded})")
        
        # 输出过滤统计
        logger.info(
            f"🔍 印章检测统计: 候选轮廓={len(contours)}, "
            f"过滤={sum(filtered_stats.values())} "
            f"(太小={filtered_stats['too_small']}, "
            f"太大={filtered_stats['too_large']}, "
            f"宽高比={filtered_stats['bad_aspect_ratio']}, "
            f"圆度={filtered_stats['low_circularity']}, "
            f"国徽={filtered_stats['emblem']}), "
            f"检测到印章={len(seals)}"
        )
        
        return seals
    
    def _detect_seals_in_grayscale(self, image: np.ndarray) -> List[Tuple[int, int, int, int]]:
        """针对黑白图片的印章检测方法。
        
        使用边缘检测和轮廓检测来识别印章，不依赖颜色信息。
        对于黑白图片，优先检测右下角区域（印章通常在此位置）。
        
        参数：
            image: 输入图片（BGR格式）
            
        返回：
            印章边界框列表 [(x1, y1, x2, y2), ...]
        """
        height, width = image.shape[:2]
        
        # 转换为灰度图
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        
        # 策略：优先检测右下角区域（印章通常在此位置）
        # 检测右下角 1/2 区域
        roi_x = width // 2
        roi_y = height // 2
        roi_gray = gray[roi_y:, roi_x:]
        
        logger.info(f"🔍 检测右下角区域: ({roi_x}, {roi_y}) -> ({width}, {height})")
        
        # 二值化（使用自适应阈值，效果更好）
        roi_binary = cv2.adaptiveThreshold(
            roi_gray,
            255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY_INV,
            11,
            2
        )
        
        # 形态学操作：闭运算连接断裂的笔画
        kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (5, 5))
        roi_binary = cv2.morphologyEx(roi_binary, cv2.MORPH_CLOSE, kernel)
        
        # 找轮廓
        contours, _ = cv2.findContours(roi_binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        if not contours:
            logger.debug("未检测到轮廓")
            return []
        
        logger.info(f"🔍 在右下角区域检测到 {len(contours)} 个轮廓")
        
        # 读取形状筛选配置
        min_seal_area = self.config.get('rendering.seal_detection.min_seal_area', 2000)
        max_seal_area = self.config.get('rendering.seal_detection.max_seal_area', 100000)
        # 对于黑白图片，使用更宽松的圆度要求
        min_circularity = 0.3
        min_aspect_ratio = 0.5
        max_aspect_ratio = 2.0
        expand_ratio = self.config.get('rendering.seal_detection.expand_radius_ratio', 1.2)
        
        seals = []
        
        for contour in contours:
            contour_area = cv2.contourArea(contour)
            
            # 过滤太小或太大的轮廓
            if contour_area < min_seal_area or contour_area > max_seal_area:
                continue
            
            # 获取边界框（相对于ROI）
            seal_x, seal_y, seal_w, seal_h = cv2.boundingRect(contour)
            
            # 转换为全图坐标
            seal_x += roi_x
            seal_y += roi_y
            
            seal_center_x = seal_x + seal_w // 2
            seal_center_y = seal_y + seal_h // 2
            
            # 检查形状：印章通常是接近正方形或圆形的
            aspect_ratio = seal_w / seal_h if seal_h > 0 else 0
            if not (min_aspect_ratio <= aspect_ratio <= max_aspect_ratio):
                continue
            
            # 计算圆度
            perimeter = cv2.arcLength(contour, True)
            if perimeter == 0:
                continue
            
            circularity = 4 * np.pi * contour_area / (perimeter * perimeter)
            if circularity < min_circularity:
                continue
            
            # 稍微扩大边界框
            padding = int(max(seal_w, seal_h) * (expand_ratio - 1.0) / 2)
            x1_expanded = max(0, seal_x - padding)
            y1_expanded = max(0, seal_y - padding)
            x2_expanded = min(width, seal_x + seal_w + padding)
            y2_expanded = min(height, seal_y + seal_h + padding)
            
            seals.append((x1_expanded, y1_expanded, x2_expanded, y2_expanded))
            logger.info(
                f"🔴 检测到印章（黑白图片）: 中心=({seal_center_x}, {seal_center_y}), "
                f"大小={seal_w}x{seal_h}, 面积={contour_area:.0f}, 圆度={circularity:.2f}, "
                f"边界=({x1_expanded}, {y1_expanded}) -> ({x2_expanded}, {y2_expanded})"
            )
        
        return seals
    
    def _check_overlap(
        self,
        region_bbox: Tuple[int, int, int, int],
        seal_bbox: Tuple[int, int, int, int],
        threshold: float = 0.01
    ) -> bool:
        """检查文字区域是否与印章重叠（智能判断，适用多种场景）。
        
        改进策略：
        1. 优先判断文字中心点是否在印章内 → 过滤
        2. 判断文字是否完全在印章下方（不重叠）→ 不过滤
        3. 判断文字是否在印章边缘（轻微接触）→ 根据重叠比例决定
        4. 自适应重叠阈值：根据文字大小和位置动态调整
        5. 过滤掉文字宽度远超印章宽度的区域（如二维码说明文字）
        
        参数：
            region_bbox: 文字区域边界 (x1, y1, x2, y2)
            seal_bbox: 印章边界 (x1, y1, x2, y2)
            threshold: 基础重叠阈值（重叠面积占文字区域的比例）
            
        返回：
            True如果应该过滤该区域
        """
        r_x1, r_y1, r_x2, r_y2 = region_bbox
        s_x1, s_y1, s_x2, s_y2 = seal_bbox
        
        region_width = r_x2 - r_x1
        region_height = r_y2 - r_y1
        region_area = region_width * region_height
        seal_width = s_x2 - s_x1
        seal_height = s_y2 - s_y1
        
        if region_area == 0:
            return False
        
        # 1. 检查文字中心点是否在印章内
        region_center_x = (r_x1 + r_x2) // 2
        region_center_y = (r_y1 + r_y2) // 2
        is_center_in_seal = (s_x1 <= region_center_x <= s_x2 and 
                            s_y1 <= region_center_y <= s_y2)
        
        if is_center_in_seal:
            logger.debug(f"文字中心点在印章内: ({region_center_x}, {region_center_y})")
            return True
        
        # 2. 检查文字是否完全在印章下方（不重叠）
        if r_y1 >= s_y2:
            logger.debug(f"文字完全在印章下方，不过滤")
            return False
        
        # 3. 计算重叠区域
        overlap_x1 = max(r_x1, s_x1)
        overlap_y1 = max(r_y1, s_y1)
        overlap_x2 = min(r_x2, s_x2)
        overlap_y2 = min(r_y2, s_y2)
        
        # 检查是否有重叠
        if overlap_x1 >= overlap_x2 or overlap_y1 >= overlap_y2:
            return False
        
        # 计算重叠宽度和高度
        overlap_width = overlap_x2 - overlap_x1
        overlap_height = overlap_y2 - overlap_y1
        
        # 4. 严格的重叠判断：x轴和y轴都必须有显著重叠
        # 计算x轴和y轴的重叠比例
        x_overlap_ratio = overlap_width / region_width if region_width > 0 else 0
        y_overlap_ratio = overlap_height / region_height if region_height > 0 else 0
        
        # 如果x轴或y轴的重叠比例太小，则不认为是真正的重叠
        # x轴阈值设为50%：文字必须有一半以上在印章范围内
        # y轴阈值设为20%：允许文字在印章边缘轻微接触
        min_x_overlap_ratio = 0.5
        min_y_overlap_ratio = 0.2
        
        if x_overlap_ratio < min_x_overlap_ratio or y_overlap_ratio < min_y_overlap_ratio:
            logger.debug(
                f"x轴重叠比例({x_overlap_ratio:.1%})或y轴重叠比例({y_overlap_ratio:.1%})"
                f"不满足阈值要求(x>={min_x_overlap_ratio:.0%}, y>={min_y_overlap_ratio:.0%})，"
                f"不认为是印章相关文字"
            )
            return False
        
        # 计算重叠面积和比例
        overlap_area = overlap_width * overlap_height
        overlap_ratio = overlap_area / region_area
        
        # 5. 过滤掉文字宽度远超印章宽度的区域（如二维码说明文字）
        # 如果文字宽度 > 印章宽度的2倍，且重叠比例 < 40%，则不认为是印章相关文字
        if region_width > seal_width * 2.0 and overlap_ratio < 0.4:
            logger.debug(
                f"文字宽度({region_width}px)远超印章宽度({seal_width}px)的2倍，"
                f"且重叠比例({overlap_ratio:.1%})较低，不认为是印章相关文字"
            )
            return False
        
        # 6. 自适应重叠阈值判断（从配置读取）
        small_text_height = self.config.get('rendering.seal_detection.overlap_threshold.small_text_height', 30)
        small_text_ratio = self.config.get('rendering.seal_detection.overlap_threshold.small_text_ratio', 0.15)
        normal_text_ratio = self.config.get('rendering.seal_detection.overlap_threshold.normal_text_ratio', 0.25)
        below_seal_ratio = self.config.get('rendering.seal_detection.overlap_threshold.below_seal_ratio', 0.40)
        
        # 判断文字是否主要在印章下方
        region_bottom_in_seal = r_y2 > s_y2  # 文字底部在印章下方
        is_mostly_below = region_bottom_in_seal and overlap_height < region_height * 0.3
        
        if is_mostly_below:
            # 文字主要在印章下方，只是上边缘接触，使用更宽松的阈值
            adaptive_threshold = below_seal_ratio
            logger.debug(f"文字主要在印章下方，使用宽松阈值: {adaptive_threshold:.0%}")
        elif region_height < small_text_height:
            # 小文字使用更严格的阈值
            adaptive_threshold = small_text_ratio
            logger.debug(f"小文字，使用严格阈值: {adaptive_threshold:.0%}")
        else:
            # 标准阈值
            adaptive_threshold = normal_text_ratio
        
        # 判断是否应该过滤
        should_filter = overlap_ratio > adaptive_threshold
        
        if should_filter:
            logger.debug(f"重叠面积超过阈值: {overlap_ratio:.1%} > {adaptive_threshold:.1%}")
        else:
            logger.debug(f"重叠面积未超过阈值: {overlap_ratio:.1%} <= {adaptive_threshold:.1%}")
        
        return should_filter
        
        return output_image, results


    def translate_image_safe(
        self,
        input_path: str,
        output_path: str,
        source_lang: str = "zh",
        target_lang: str = "en"
    ) -> Tuple[Optional[QualityReport], Optional[str]]:
        """Safely translate an image with comprehensive error handling.
        
        This method wraps translate_image with additional error handling
        to ensure the original image is never modified and errors are
        gracefully handled.
        
        Args:
            input_path: Path to the input image
            output_path: Path where the translated image should be saved
            source_lang: Source language code
            target_lang: Target language code
            
        Returns:
            Tuple of (QualityReport, error_message) where:
            - QualityReport is the report if successful, None if failed
            - error_message is None if successful, error description if failed
            
        _Requirements: 10.2, 10.4, 10.5_
        """
        try:
            report = self.translate_image(
                input_path, output_path, source_lang, target_lang
            )
            return report, None
        except PipelineError as e:
            error_msg = str(e)
            logger.error(f"Pipeline error: {error_msg}")
            return None, error_msg
        except Exception as e:
            error_msg = f"Unexpected error: {str(e)}"
            logger.error(error_msg)
            return None, error_msg
    
    def _process_region_safe(
        self,
        image: np.ndarray,
        region: TextRegion,
        result: TranslationResult,
        region_index: int
    ) -> Tuple[np.ndarray, TranslationResult]:
        """Safely process a single region with error handling.
        
        If processing fails, returns the original image unchanged and
        marks the translation result as failed.
        
        Args:
            image: Working image
            region: Text region to process
            result: Translation result for this region
            region_index: Index of the region for logging
            
        Returns:
            Tuple of (processed_image, updated_result)
            
        _Requirements: 10.5_
        """
        if not result.success:
            return image, result
        
        if not result.translated_text.strip():
            return image, result
        
        # Skip rendering if should_render is False (e.g., QR codes)
        if hasattr(region, 'should_render') and not region.should_render:
            logger.info(f"Skipping rendering for region {region_index} (should_render=False)")
            return image, result
        
        try:
            # Sample background color
            bg_color = self.background_sampler.sample_background(image, region)
            
            # Fill region with background (simplified strategy)
            processed = self.background_sampler.process_region(image, region, bg_color)
            
            # Render translated text (使用左对齐)
            from src.rendering.text_renderer import TextAlignment
            processed = self.text_renderer.render_text(
                processed,
                region,
                result.translated_text,
                bg_color,
                alignment=TextAlignment.LEFT
            )
            
            return processed, result
            
        except Exception as e:
            logger.warning(
                f"Failed to process region {region_index}: {e}. "
                f"Continuing with other regions."
            )
            # Return original image and mark result as failed
            failed_result = TranslationResult(
                source_text=result.source_text,
                translated_text=result.translated_text,
                confidence=result.confidence,
                success=False,
                error_message=f"Rendering failed: {e}"
            )
            return image, failed_result
    
    def _create_partial_failure_report(
        self,
        total_regions: int,
        successful_results: List[TranslationResult],
        failed_regions: List[TextRegion],
        output_image: Optional[np.ndarray]
    ) -> QualityReport:
        """Create a quality report for partial failure scenarios.
        
        Args:
            total_regions: Total number of text regions
            successful_results: List of successful translation results
            failed_regions: List of regions that failed
            output_image: Output image (may be None)
            
        Returns:
            QualityReport reflecting partial success
            
        _Requirements: 10.5_
        """
        successful_count = len(successful_results)
        
        if total_regions > 0:
            coverage = successful_count / total_regions
        else:
            coverage = 1.0
        
        # Check for artifacts if we have an output image
        has_artifacts = False
        artifact_locations = []
        if output_image is not None:
            has_artifacts, artifact_locations = self.validator.check_artifacts(
                output_image
            )
        
        # Determine quality level
        quality_level = self.validator.determine_quality_level(
            coverage, has_artifacts, len(artifact_locations)
        )
        
        return QualityReport(
            translation_coverage=coverage,
            total_regions=total_regions,
            translated_regions=successful_count,
            failed_regions=failed_regions,
            has_artifacts=has_artifacts,
            artifact_locations=artifact_locations,
            overall_quality=quality_level
        )


    def translate_batch(
        self,
        input_paths: List[str],
        output_dir: str,
        source_lang: str = "zh",
        target_lang: str = "en",
        parallel: Optional[bool] = None
    ) -> List[Tuple[str, Optional[QualityReport], Optional[str]]]:
        """Translate multiple images in batch.
        
        Processes multiple images, optionally in parallel, and collects
        quality reports for each. Failed images do not stop processing
        of other images.
        
        Args:
            input_paths: List of input image paths
            output_dir: Directory where translated images should be saved
            source_lang: Source language code
            target_lang: Target language code
            parallel: Whether to process in parallel (None uses config setting)
            
        Returns:
            List of tuples (input_path, QualityReport, error_message) where:
            - input_path: The original input path
            - QualityReport: Report if successful, None if failed
            - error_message: None if successful, error description if failed
            
        _Requirements: 13.2, 14.4_
        """
        if not input_paths:
            logger.warning("No input paths provided for batch translation")
            return []
        
        # Ensure output directory exists
        os.makedirs(output_dir, exist_ok=True)
        
        # Determine whether to use parallel processing
        use_parallel = parallel if parallel is not None else self.parallel_enabled
        
        logger.info(
            f"Starting batch translation of {len(input_paths)} images "
            f"(parallel={use_parallel})"
        )
        
        results = []
        
        if use_parallel and len(input_paths) > 1:
            results = self._translate_batch_parallel(
                input_paths, output_dir, source_lang, target_lang
            )
        else:
            results = self._translate_batch_sequential(
                input_paths, output_dir, source_lang, target_lang
            )
        
        # Log summary
        successful = sum(1 for _, report, _ in results if report is not None)
        failed = len(results) - successful
        
        logger.info(
            f"Batch translation complete: {successful} succeeded, {failed} failed"
        )
        
        return results
    
    def _translate_batch_sequential(
        self,
        input_paths: List[str],
        output_dir: str,
        source_lang: str,
        target_lang: str
    ) -> List[Tuple[str, Optional[QualityReport], Optional[str]]]:
        """Process images sequentially.
        
        Args:
            input_paths: List of input image paths
            output_dir: Output directory
            source_lang: Source language code
            target_lang: Target language code
            
        Returns:
            List of (input_path, QualityReport, error_message) tuples
        """
        results = []
        
        for i, input_path in enumerate(input_paths):
            logger.info(f"Processing image {i + 1}/{len(input_paths)}: {input_path}")
            
            output_path = self._generate_output_path(input_path, output_dir)
            report, error = self.translate_image_safe(
                input_path, output_path, source_lang, target_lang
            )
            
            results.append((input_path, report, error))
            
            if report:
                logger.debug(
                    f"Image {i + 1} completed: "
                    f"coverage={report.translation_coverage:.1%}"
                )
            else:
                logger.warning(f"Image {i + 1} failed: {error}")
        
        return results
    
    def _translate_batch_parallel(
        self,
        input_paths: List[str],
        output_dir: str,
        source_lang: str,
        target_lang: str
    ) -> List[Tuple[str, Optional[QualityReport], Optional[str]]]:
        """Process images in parallel using thread pool.
        
        Args:
            input_paths: List of input image paths
            output_dir: Output directory
            source_lang: Source language code
            target_lang: Target language code
            
        Returns:
            List of (input_path, QualityReport, error_message) tuples
            in the same order as input_paths
        """
        # Initialize results with None placeholders
        results: List[Tuple[str, Optional[QualityReport], Optional[str]]] = [
            (path, None, None) for path in input_paths
        ]
        
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            # Submit all tasks
            future_to_index = {}
            for i, input_path in enumerate(input_paths):
                output_path = self._generate_output_path(input_path, output_dir)
                future = executor.submit(
                    self.translate_image_safe,
                    input_path,
                    output_path,
                    source_lang,
                    target_lang
                )
                future_to_index[future] = i
            
            # Collect results as they complete
            for future in as_completed(future_to_index):
                index = future_to_index[future]
                input_path = input_paths[index]
                
                try:
                    report, error = future.result()
                    results[index] = (input_path, report, error)
                    
                    if report:
                        logger.debug(
                            f"Image completed: {input_path}, "
                            f"coverage={report.translation_coverage:.1%}"
                        )
                    else:
                        logger.warning(f"Image failed: {input_path}, error={error}")
                        
                except Exception as e:
                    error_msg = f"Unexpected error: {str(e)}"
                    results[index] = (input_path, None, error_msg)
                    logger.error(f"Image failed: {input_path}, error={error_msg}")
        
        return results
    
    def _generate_output_path(self, input_path: str, output_dir: str) -> str:
        """Generate output path for a translated image.
        
        Creates an output filename by appending '_translated' to the
        original filename.
        
        Args:
            input_path: Original input image path
            output_dir: Output directory
            
        Returns:
            Full output path
        """
        input_name = Path(input_path).stem
        input_ext = Path(input_path).suffix
        output_name = f"{input_name}_translated{input_ext}"
        return os.path.join(output_dir, output_name)
    
    def get_batch_summary(
        self,
        results: List[Tuple[str, Optional[QualityReport], Optional[str]]]
    ) -> dict:
        """Generate a summary of batch translation results.
        
        Args:
            results: List of batch translation results
            
        Returns:
            Dictionary containing summary statistics
        """
        total = len(results)
        successful = sum(1 for _, report, _ in results if report is not None)
        failed = total - successful
        
        # Calculate average coverage for successful translations
        coverages = [
            report.translation_coverage 
            for _, report, _ in results 
            if report is not None
        ]
        avg_coverage = sum(coverages) / len(coverages) if coverages else 0.0
        
        # Count quality levels
        quality_counts = {level.value: 0 for level in QualityLevel}
        for _, report, _ in results:
            if report is not None:
                quality_counts[report.overall_quality.value] += 1
        
        return {
            'total_images': total,
            'successful': successful,
            'failed': failed,
            'success_rate': successful / total if total > 0 else 0.0,
            'average_coverage': avg_coverage,
            'quality_distribution': quality_counts
        }

    def translate_batch_generator(
        self,
        input_paths: Iterator[str],
        output_dir: str,
        source_lang: str = "zh",
        target_lang: str = "en"
    ) -> Generator[Tuple[str, Optional[QualityReport], Optional[str]], None, None]:
        """Translate images using a memory-efficient generator.
        
        This method processes images one at a time using a generator,
        which is more memory-efficient for large batches as it doesn't
        load all images into memory at once.
        
        Args:
            input_paths: Iterator of input image paths
            output_dir: Directory where translated images should be saved
            source_lang: Source language code
            target_lang: Target language code
            
        Yields:
            Tuples of (input_path, QualityReport, error_message) where:
            - input_path: The original input path
            - QualityReport: Report if successful, None if failed
            - error_message: None if successful, error description if failed
            
        _Requirements: 14.1, 14.3_
        """
        # Ensure output directory exists
        os.makedirs(output_dir, exist_ok=True)
        
        logger.info("Starting memory-efficient batch translation using generator")
        
        processed_count = 0
        
        for input_path in input_paths:
            processed_count += 1
            logger.debug(f"Processing image {processed_count}: {input_path}")
            
            output_path = self._generate_output_path(input_path, output_dir)
            report, error = self.translate_image_safe(
                input_path, output_path, source_lang, target_lang
            )
            
            # Release memory after each image if memory-efficient mode is enabled
            if self.memory_efficient:
                self._release_memory()
            
            yield (input_path, report, error)
            
            if report:
                logger.debug(
                    f"Image {processed_count} completed: "
                    f"coverage={report.translation_coverage:.1%}"
                )
            else:
                logger.warning(f"Image {processed_count} failed: {error}")
        
        logger.info(f"Generator batch translation complete: {processed_count} images processed")
    
    def translate_batch_memory_efficient(
        self,
        input_paths: List[str],
        output_dir: str,
        source_lang: str = "zh",
        target_lang: str = "en",
        batch_size: int = 10
    ) -> List[Tuple[str, Optional[QualityReport], Optional[str]]]:
        """Translate images in memory-efficient batches.
        
        Processes images in smaller batches to limit memory usage,
        releasing memory between batches. This is useful for processing
        large numbers of images without running out of memory.
        
        Args:
            input_paths: List of input image paths
            output_dir: Directory where translated images should be saved
            source_lang: Source language code
            target_lang: Target language code
            batch_size: Number of images to process before releasing memory
            
        Returns:
            List of tuples (input_path, QualityReport, error_message)
            
        _Requirements: 14.1, 14.3_
        """
        if not input_paths:
            logger.warning("No input paths provided for batch translation")
            return []
        
        # Ensure output directory exists
        os.makedirs(output_dir, exist_ok=True)
        
        logger.info(
            f"Starting memory-efficient batch translation of {len(input_paths)} images "
            f"(batch_size={batch_size})"
        )
        
        all_results = []
        total_batches = (len(input_paths) + batch_size - 1) // batch_size
        
        for batch_num in range(total_batches):
            start_idx = batch_num * batch_size
            end_idx = min(start_idx + batch_size, len(input_paths))
            batch_paths = input_paths[start_idx:end_idx]
            
            logger.info(
                f"Processing batch {batch_num + 1}/{total_batches} "
                f"({len(batch_paths)} images)"
            )
            
            # Process this batch
            batch_results = self._translate_batch_sequential(
                batch_paths, output_dir, source_lang, target_lang
            )
            all_results.extend(batch_results)
            
            # Release memory between batches
            self._release_memory()
            logger.debug(f"Memory released after batch {batch_num + 1}")
        
        # Log summary
        successful = sum(1 for _, report, _ in all_results if report is not None)
        failed = len(all_results) - successful
        
        logger.info(
            f"Memory-efficient batch translation complete: "
            f"{successful} succeeded, {failed} failed"
        )
        
        return all_results
    
    def get_ocr_cache_stats(self) -> Optional[dict]:
        """Get OCR cache statistics from the OCR engine.
        
        Returns:
            Dictionary containing cache statistics, or None if caching is disabled
            
        _Requirements: 14.2_
        """
        return self.ocr.get_cache_stats()
    
    def clear_ocr_cache(self) -> None:
        """Clear the OCR cache.
        
        _Requirements: 14.2_
        """
        self.ocr.clear_cache()
    
    def get_memory_stats(self) -> dict:
        """Get memory-related statistics and configuration.
        
        Returns:
            Dictionary containing memory configuration and stats
            
        _Requirements: 14.1, 14.3_
        """
        return {
            'max_workers': self.max_workers,
            'parallel_enabled': self.parallel_enabled,
            'memory_efficient': self.memory_efficient,
            'ocr_cache_stats': self.get_ocr_cache_stats()
        }

    def _separate_watermark_regions(
        self, 
        regions: List[TextRegion],
        image: np.ndarray
    ) -> List[TextRegion]:
        """分离包含水印的混合文本区域。
        
        对于包含水印关键词的文本区域，尝试分离出纯水印部分和有效文本部分，
        创建新的TextRegion只包含有效文本。
        
        Args:
            regions: 原始文本区域列表
            image: 原始图片（用于二次OCR）
            
        Returns:
            处理后的文本区域列表（水印被分离或过滤）
        """
        ignore_keywords = self.config.get('ocr.ignore_keywords', [])
        if not ignore_keywords:
            return regions
        
        import re
        import cv2
        
        separated_regions = []
        
        for region in regions:
            # 检查是否包含水印关键词
            contains_watermark = False
            for keyword in ignore_keywords:
                if keyword.lower() in region.text.lower():
                    contains_watermark = True
                    break
            
            if not contains_watermark:
                # 不包含水印，直接保留
                separated_regions.append(region)
                continue
            
            # 包含水印，尝试分离
            logger.info(f"检测到包含水印的区域: '{region.text}'")
            
            # 移除水印关键词后的剩余文本
            text_clean = re.sub(r'[\s\(\)（）]', '', region.text)
            remaining_text = text_clean
            for keyword in ignore_keywords:
                remaining_text = re.sub(re.escape(keyword), '', remaining_text, flags=re.IGNORECASE)
            
            # 检查剩余文本是否有意义
            chinese_chars = re.findall(r'[\u4e00-\u9fff]', remaining_text)
            has_meaningful_content = len(chinese_chars) >= 2
            
            if not has_meaningful_content:
                # 纯水印，过滤掉
                logger.info(f"  → 纯水印，过滤掉 (剩余文本: '{remaining_text}')")
                continue
            
            # 有有效内容，尝试定位有效文本的位置
            logger.info(f"  → 包含有效内容: '{remaining_text}'")
            
            # 从原文本中提取有效部分（保留空格和标点，包括括号）
            effective_text = region.text
            for keyword in ignore_keywords:
                effective_text = re.sub(re.escape(keyword), '', effective_text, flags=re.IGNORECASE)
            effective_text = re.sub(r'\s+', ' ', effective_text).strip()
            # 不移除括号，保留原始格式
            
            # 尝试估算有效文本的位置
            # 需要判断水印在左侧还是右侧
            x1, y1, x2, y2 = region.bbox
            region_width = x2 - x1
            region_height = y2 - y1
            
            # 检查水印关键词在原文本中的位置
            watermark_keyword = None
            for keyword in ignore_keywords:
                if keyword.lower() in region.text.lower():
                    watermark_keyword = keyword
                    break
            
            # 判断水印在左侧还是右侧
            watermark_at_start = False
            watermark_at_end = False
            
            if watermark_keyword:
                # 移除空格后检查位置
                text_no_space = region.text.replace(' ', '')
                keyword_no_space = watermark_keyword.replace(' ', '')
                
                if text_no_space.lower().startswith(keyword_no_space.lower()):
                    watermark_at_start = True
                elif text_no_space.lower().endswith(keyword_no_space.lower()):
                    watermark_at_end = True
            
            # 使用TextRenderer来测量实际文本宽度（更准确）
            try:
                # 测量原始文本和有效文本的实际宽度
                original_text_width = self.text_renderer.measure_text_width(
                    region.text, region.font_size
                )
                effective_text_width = self.text_renderer.measure_text_width(
                    effective_text, region.font_size
                )
                
                # 计算水印部分的宽度
                watermark_width = original_text_width - effective_text_width
                
                if watermark_width > 0 and effective_text_width > 0:
                    # 根据实际文本宽度比例计算新的bbox
                    width_ratio = watermark_width / original_text_width
                    
                    if watermark_at_end:
                        # 水印在右侧，有效文本在左侧
                        # 缩小右边界
                        new_x2 = int(x2 - region_width * width_ratio)
                        new_bbox = (x1, y1, new_x2, y2)
                        position_desc = "水印在右侧"
                    else:
                        # 水印在左侧（默认），有效文本在右侧
                        # 移动左边界
                        new_x1 = int(x1 + region_width * width_ratio)
                        new_bbox = (new_x1, y1, x2, y2)
                        position_desc = "水印在左侧"
                    
                    # 创建新的边界框，只包含有效文本部分
                    new_region = TextRegion(
                        bbox=new_bbox,
                        text=effective_text,
                        confidence=region.confidence,
                        font_size=region.font_size,
                        angle=region.angle
                    )
                    
                    separated_regions.append(new_region)
                    logger.info(
                        f"  → 成功分离: '{effective_text}' at {new_bbox} "
                        f"({position_desc}, 原文={original_text_width}px, 有效={effective_text_width}px, "
                        f"水印={watermark_width}px, 比例={width_ratio:.2%})"
                    )
                else:
                    # 无法估算，使用原区域
                    new_region = TextRegion(
                        bbox=region.bbox,
                        text=effective_text,
                        confidence=region.confidence,
                        font_size=region.font_size,
                        angle=region.angle
                    )
                    separated_regions.append(new_region)
                    logger.info(f"  → 无法估算位置（宽度异常），使用原区域: '{effective_text}'")
            except Exception as e:
                # 如果测量失败，回退到使用原区域
                logger.warning(f"  → 测量文本宽度失败: {e}，使用原区域")
                new_region = TextRegion(
                    bbox=region.bbox,
                    text=effective_text,
                    confidence=region.confidence,
                    font_size=region.font_size,
                    angle=region.angle
                )
                separated_regions.append(new_region)
                logger.info(f"  → 使用原区域: '{effective_text}'")
        
        logger.info(f"水印分离完成: {len(regions)} 个区域 -> {len(separated_regions)} 个区域")
        return separated_regions
