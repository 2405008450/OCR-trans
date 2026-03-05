"""Seal text handler for extracting and rendering seal-related text.

This module provides functionality to:
1. Extract text from seals and seal-overlapping regions
2. Find suitable positions near seals to render translations
3. Ensure translations don't overlap with existing content
"""

import logging
from typing import List, Tuple, Optional, Dict, Any
import numpy as np
import cv2

from src.models import TextRegion
from src.config import ConfigManager
from src.exceptions import OCRError


logger = logging.getLogger(__name__)


class SealTextHandler:
    """Handler for seal-related text extraction and rendering.
    
    This class identifies text within or overlapping with seals,
    and finds appropriate positions to render their translations
    without overlapping existing content.
    """
    
    def __init__(self, config: ConfigManager):
        """Initialize the seal text handler.
        
        Args:
            config: Configuration manager instance
        """
        self.config = config
        self.enabled = config.get('rendering.seal_text_handling.enabled', False)
        self.min_distance = config.get('rendering.seal_text_handling.min_distance', 20)
        self.search_radius = config.get('rendering.seal_text_handling.search_radius', 150)
        self.overlap_threshold = config.get('rendering.seal_text_handling.overlap_threshold', 0.1)
        
        # Frame rendering configuration
        self.frame_rendering_enabled = config.get('rendering.seal_text_handling.frame_rendering.enabled', True)
        self.frame_margin = config.get('rendering.seal_text_handling.frame_rendering.frame_margin', 20)
        self.min_frame_area = config.get('rendering.seal_text_handling.frame_rendering.min_frame_area', 10000)
        self.frame_aspect_ratio_range = config.get('rendering.seal_text_handling.frame_rendering.frame_aspect_ratio_range', [0.8, 1.5])
        
        # Initialize GLM-OCR layout parser if enabled
        self.layout_parser = None
        glm_layout_enabled = config.get(
            'rendering.seal_text_handling.glm_layout_parsing.enabled', 
            False
        )
        
        if glm_layout_enabled:
            try:
                from src.rendering.glm_layout_parser import GLMOCRLayoutParser
                self.layout_parser = GLMOCRLayoutParser(config)
                self.layout_parser.initialize()
                logger.info("✅ GLM-OCR layout parser initialized successfully")
            except OCRError as e:
                # Log warning if API key is missing or initialization fails
                logger.warning(
                    f"⚠️  GLM-OCR layout parser initialization failed: {e}. "
                    f"Falling back to existing OCR for seal text recognition."
                )
                self.layout_parser = None
            except ImportError as e:
                logger.warning(
                    f"⚠️  Failed to import GLMOCRLayoutParser: {e}. "
                    f"Falling back to existing OCR for seal text recognition."
                )
                self.layout_parser = None
            except Exception as e:
                logger.warning(
                    f"⚠️  Unexpected error initializing GLM-OCR layout parser: {e}. "
                    f"Falling back to existing OCR for seal text recognition."
                )
                self.layout_parser = None
        
        logger.info(
            f"SealTextHandler initialized: enabled={self.enabled}, "
            f"min_distance={self.min_distance}, search_radius={self.search_radius}, "
            f"glm_layout_parser={'enabled' if self.layout_parser else 'disabled'}"
        )
    
    def identify_seal_text_regions(
        self,
        image: np.ndarray,
        regions: List[TextRegion],
        seals: List[Tuple[int, int, int, int]]
    ) -> Tuple[List[TextRegion], List[TextRegion]]:
        """Identify text regions that overlap with seals.
        
        If GLM-OCR layout parsing is enabled, uses the API to identify
        seal text. Otherwise, falls back to overlap detection logic.
        
        This method also merges date text fragments that are split by seals.
        
        Args:
            image: Full image as numpy array (for GLM-OCR API calls)
            regions: All text regions from existing OCR
            seals: List of seal bounding boxes [(x1, y1, x2, y2), ...]
            
        Returns:
            Tuple of (seal_text_regions, non_seal_regions)
        """
        if not self.enabled or not seals:
            return [], regions
        
        seal_text_regions = []
        non_seal_regions = []
        
        # Try GLM-OCR layout parsing if available
        if self.layout_parser and self.layout_parser.enabled:
            logger.info(
                f"🔍 Using GLM-OCR layout parsing for {len(seals)} seal(s)"
            )
            
            # Track which seals were successfully processed by GLM-OCR
            successfully_parsed_seals = set()
            
            # Process each seal with GLM-OCR
            for seal_bbox in seals:
                try:
                    glm_regions = self.layout_parser.parse_seal_region(
                        image, seal_bbox
                    )
                    
                    if glm_regions:
                        # Convert GLM-OCR results to TextRegion objects
                        valid_regions_count = 0
                        for glm_region in glm_regions:
                            text = glm_region['text'].strip()
                            
                            # 过滤掉 HTML 标签和无效文字
                            # Filter out HTML tags and invalid text
                            if not text or text.startswith('<') or text.endswith('>'):
                                logger.debug(
                                    f"Skipping invalid GLM-OCR text: '{text}'"
                                )
                                continue
                            
                            # 过滤掉只包含标点符号的文字
                            # Filter out text with only punctuation
                            if all(c in '.,;:!?<>()[]{}"\'-_=+*/\\|' for c in text):
                                logger.debug(
                                    f"Skipping punctuation-only text: '{text}'"
                                )
                                continue
                            
                            # 过滤掉已经是英文的文本（不需要翻译）
                            # Filter out text that is already in English (no translation needed)
                            if self._is_english_text(text):
                                logger.info(
                                    f"⏭️  Skipping English text (no translation needed): '{text}'"
                                )
                                continue
                            
                            # 过滤掉完整的英文日期（如"April 17, 2018"）
                            # Filter out complete English dates (like "April 17, 2018")
                            if self._is_english_date(text):
                                logger.info(
                                    f"⏭️  Skipping English date (no translation needed): '{text}'"
                                )
                                continue
                            
                            # 检查是否需要分离印章机关和日期
                            # 如果文本包含多行，且同时包含非日期内容和日期内容，则分离
                            separated_regions = self._separate_seal_org_and_date(
                                text, glm_region['bbox'], glm_region['confidence'], seal_bbox
                            )
                            
                            if separated_regions:
                                # 成功分离，添加所有分离后的区域
                                for sep_region in separated_regions:
                                    # 确定文本类型（用于交互式验证）
                                    # 判断是印章内文字还是日期
                                    import re
                                    is_date = bool(re.search(r'\d{4}.*\d{1,2}.*\d{1,2}', sep_region.text))
                                    text_type = 'seal_overlap' if is_date else 'seal_inner'
                                    
                                    # 交互式验证
                                    should_process, corrected_text = self.verify_seal_text_with_user(
                                        sep_region,
                                        text_type,
                                        image
                                    )
                                    
                                    if not should_process:
                                        logger.info(f"⏭️  用户跳过区域: '{sep_region.text}'")
                                        continue
                                    
                                    if corrected_text:
                                        sep_region.text = corrected_text
                                        sep_region.user_corrected = True
                                        logger.info(f"✏️  用户修正内容: '{corrected_text}'")
                                    
                                    seal_text_regions.append(sep_region)
                                    valid_regions_count += 1
                                    logger.info(
                                        f"✅ GLM-OCR identified seal text (separated): '{sep_region.text[:50]}...' "
                                        f"at {sep_region.bbox}"
                                    )
                            else:
                                # 无需分离，直接添加
                                text_region = TextRegion(
                                    bbox=glm_region['bbox'],
                                    text=text,
                                    confidence=glm_region['confidence'],
                                    font_size=self._estimate_font_size(glm_region['bbox']),
                                    angle=0.0
                                )
                                # Mark which seal this region is associated with
                                text_region.overlapping_seal = seal_bbox
                                
                                # 确定文本类型（用于交互式验证）
                                import re
                                is_date = bool(re.search(r'\d{4}.*\d{1,2}.*\d{1,2}', text))
                                text_type = 'seal_overlap' if is_date else 'seal_inner'
                                
                                # 交互式验证
                                should_process, corrected_text = self.verify_seal_text_with_user(
                                    text_region,
                                    text_type,
                                    image
                                )
                                
                                if not should_process:
                                    logger.info(f"⏭️  用户跳过区域: '{text_region.text}'")
                                    continue
                                
                                if corrected_text:
                                    text_region.text = corrected_text
                                    text_region.user_corrected = True
                                    logger.info(f"✏️  用户修正内容: '{corrected_text}'")
                                
                                seal_text_regions.append(text_region)
                                valid_regions_count += 1
                                
                                logger.info(
                                    f"✅ GLM-OCR identified seal text: '{text_region.text[:50]}...' "
                                    f"at {text_region.bbox}"
                                )
                        
                        # Only mark as successfully processed if we got valid regions
                        if valid_regions_count > 0:
                            successfully_parsed_seals.add(seal_bbox)
                            logger.info(
                                f"✅ GLM-OCR successfully processed seal {seal_bbox} "
                                f"with {valid_regions_count} valid text region(s)"
                            )
                        else:
                            logger.warning(
                                f"⚠️  GLM-OCR returned no valid text for seal {seal_bbox}, "
                                f"falling back to overlap detection"
                            )
                    else:
                        logger.warning(
                            f"⚠️  GLM-OCR returned no results for seal {seal_bbox}, "
                            f"falling back to overlap detection"
                        )
                        
                except Exception as e:
                    logger.warning(
                        f"⚠️  GLM-OCR failed for seal {seal_bbox}: {e}. "
                        f"Falling back to overlap detection"
                    )
            
            # For seals that weren't successfully processed by GLM-OCR,
            # fall back to overlap detection
            if len(successfully_parsed_seals) < len(seals):
                logger.info(
                    f"Falling back to overlap detection for "
                    f"{len(seals) - len(successfully_parsed_seals)} seal(s)"
                )
                
                for region in regions:
                    region_bbox = region.bbox
                    
                    # Check if region overlaps with any seal that wasn't
                    # successfully processed by GLM-OCR
                    overlaps_unprocessed_seal = False
                    for seal_bbox in seals:
                        if seal_bbox not in successfully_parsed_seals:
                            if self._check_overlap(region_bbox, seal_bbox):
                                overlaps_unprocessed_seal = True
                                region.overlapping_seal = seal_bbox
                                break
                    
                    if overlaps_unprocessed_seal:
                        seal_text_regions.append(region)
                        logger.info(
                            f"📍 Overlap detection identified seal text: "
                            f"'{region.text}' at {region.bbox}"
                        )
            else:
                # All seals were successfully processed by GLM-OCR
                # No need for overlap detection fallback
                logger.info(
                    f"✅ All {len(seals)} seal(s) successfully processed by GLM-OCR, "
                    f"skipping overlap detection"
                )
            
            # Filter out existing OCR regions that overlap with seals
            # For successfully processed seals, add overlapping regions to seal_text_regions
            # (they might be seal-related text like dates below the seal that GLM-OCR missed)
            for region in regions:
                region_bbox = region.bbox
                
                # Check if this region overlaps with any successfully
                # processed seal
                overlaps_processed_seal = False
                overlapping_seal_bbox = None
                for seal_bbox in successfully_parsed_seals:
                    if self._check_overlap(region_bbox, seal_bbox):
                        overlaps_processed_seal = True
                        overlapping_seal_bbox = seal_bbox
                        logger.info(
                            f"📍 Region '{region.text}' overlaps with GLM-OCR processed seal {seal_bbox}, "
                            f"adding to seal_text_regions"
                        )
                        break
                
                if overlaps_processed_seal:
                    # Add to seal_text_regions (might be seal-related text like dates)
                    region.overlapping_seal = overlapping_seal_bbox
                    seal_text_regions.append(region)
                else:
                    # For unprocessed seals, check if region overlaps
                    # (these were already added to seal_text_regions above)
                    overlaps_unprocessed_seal = False
                    for seal_bbox in seals:
                        if seal_bbox not in successfully_parsed_seals:
                            if self._check_overlap(region_bbox, seal_bbox):
                                overlaps_unprocessed_seal = True
                                break
                    
                    # Only add if it doesn't overlap with any seal
                    if not overlaps_unprocessed_seal:
                        non_seal_regions.append(region)
        
        else:
            # Fall back to existing overlap detection logic
            logger.info(
                f"Using overlap detection for {len(seals)} seal(s) "
                f"(GLM-OCR layout parsing not available)"
            )
            
            for region in regions:
                region_bbox = region.bbox
                
                # Check if region overlaps with any seal
                overlaps_seal = False
                for seal_bbox in seals:
                    if self._check_overlap(region_bbox, seal_bbox):
                        overlaps_seal = True
                        # Mark which seal this region overlaps with
                        region.overlapping_seal = seal_bbox
                        break
                
                if overlaps_seal:
                    seal_text_regions.append(region)
                    logger.info(
                        f"📍 Identified seal text: '{region.text}' at {region.bbox}"
                    )
                else:
                    non_seal_regions.append(region)
        
        # Merge date text fragments that are split by seals
        seal_text_regions = self._merge_date_fragments(seal_text_regions, seals)
        
        # Mark seal text regions with their type (seal_inner or seal_overlap)
        seal_text_regions = self._classify_seal_text_regions(seal_text_regions, seals)
        
        # 交互式验证：在分类完成后，对每个印章文字区域进行验证
        verified_seal_text_regions = []
        for region in seal_text_regions:
            text_type = getattr(region, 'seal_text_type', 'seal_overlap')
            
            # 交互式验证
            should_process, corrected_text = self.verify_seal_text_with_user(
                region,
                text_type,
                image
            )
            
            if not should_process:
                logger.info(f"⏭️  用户跳过区域: '{region.text}'")
                continue
            
            if corrected_text:
                region.text = corrected_text
                region.user_corrected = True
                logger.info(f"✏️  用户修正内容: '{corrected_text}'")
            
            verified_seal_text_regions.append(region)
        
        seal_text_regions = verified_seal_text_regions
        
        logger.info(
            f"Identified {len(seal_text_regions)} seal text regions, "
            f"{len(non_seal_regions)} non-seal regions"
        )
        
        return seal_text_regions, non_seal_regions
    
    def find_translation_positions(
        self,
        seal_text_regions: List[TextRegion],
        all_regions: List[TextRegion],
        image_shape: Tuple[int, int],
        use_collision_detection: bool = True,
        translations: Optional[Dict[int, str]] = None,
        image: Optional[np.ndarray] = None,
        registration_authority_bbox: Optional[Tuple[int, int, int, int]] = None
    ) -> List[Tuple[TextRegion, Tuple[int, int, int, int]]]:
        """Find suitable positions to render seal text translations.
        
        改进策略（启用碰撞检测）：
        1. 计算翻译文本的实际尺寸（使用 TextRenderer）
        2. 建立已占用区域集合（印章、原始文本、已渲染翻译）
        3. 通过矩形碰撞检测找到不与任何已占区域重叠的安全位置
        4. 优先级：印章正上方（居中）> 印章下方 > 印章左右两侧
        
        Args:
            seal_text_regions: Text regions overlapping with seals
            all_regions: All text regions (for collision detection)
            image_shape: Image dimensions (height, width)
            use_collision_detection: 是否使用新的碰撞检测方法（默认True）
            translations: 实际翻译文本字典 {region_id: translated_text}（可选）
            registration_authority_bbox: "Registration Authority"文本的边界框（可选）
            
        Returns:
            List of (region, new_bbox) tuples where new_bbox is (x1, y1, x2, y2)
        """
        print(f"\n[FIND_TRANSLATION_POSITIONS] 方法被调用！")
        print(f"[FIND_TRANSLATION_POSITIONS] seal_text_regions={len(seal_text_regions)}")
        print(f"[FIND_TRANSLATION_POSITIONS] use_collision_detection={use_collision_detection}")
        print(f"[FIND_TRANSLATION_POSITIONS] registration_authority_bbox={registration_authority_bbox}")
        
        logger.info(
            f"🔍 find_translation_positions 被调用: "
            f"seal_text_regions={len(seal_text_regions)}, "
            f"all_regions={len(all_regions)}, "
            f"use_collision_detection={use_collision_detection}, "
            f"translations={'已提供' if translations else '未提供'}, "
            f"image={'已提供' if image is not None else '未提供'}, "
            f"registration_authority_bbox={registration_authority_bbox}"
        )
        
        if not seal_text_regions:
            return []
        
        height, width = image_shape
        positioned_regions = []
        
        # 建立已占用区域集合（用于碰撞检测）
        occupied_regions = []
        
        # 如果提供了已渲染的图片，从图片中检测已渲染的文字区域
        if image is not None:
            logger.info("🔍 从渲染后的图片中检测已渲染的文字区域...")
            rendered_text_regions = self._detect_rendered_text_regions(image)
            logger.info(f"🔍 检测到 {len(rendered_text_regions)} 个已渲染的文字区域")
            
            # 打印前10个区域用于调试
            for i, rendered_bbox in enumerate(rendered_text_regions[:10]):
                x1, y1, x2, y2 = rendered_bbox
                logger.info(f"  📍 已渲染区域 #{i+1}: ({x1}, {y1}) -> ({x2}, {y2}), 尺寸={x2-x1}x{y2-y1}px")
            
            # 将已渲染的文字区域加入已占用区域
            for rendered_bbox in rendered_text_regions:
                occupied_regions.append(rendered_bbox)
            
            logger.info(f"🔍 已占用区域总数: {len(occupied_regions)}")
        else:
            # 如果没有提供渲染后的图片，使用原始OCR区域
            logger.info("🔍 使用原始OCR区域建立已占用区域...")
            
            # 添加所有原始文本区域（但排除印章文字区域，因为它们会被翻译替换）
            seal_text_bboxes = set(tuple(r.bbox) for r in seal_text_regions)
            
            # 获取所有印章的位置，用于判断文本片段是否在印章附近
            seal_positions = []
            for region in seal_text_regions:
                seal_bbox = getattr(region, 'overlapping_seal', None)
                if seal_bbox and seal_bbox not in seal_positions:
                    seal_positions.append(seal_bbox)
            
            logger.info(f"🔍 印章文字区域的bbox:")
            for r in seal_text_regions:
                logger.info(f"  - {r.text[:30]}... bbox={r.bbox}")
            
            logger.info(f"🔍 检测到 {len(seal_positions)} 个印章位置")
            
            # 调试：打印所有传入的区域
            logger.info(f"🔍 传入的all_regions共{len(all_regions)}个:")
            for i, r in enumerate(all_regions):
                logger.info(f"  [{i}] '{r.text[:30]}...' bbox={r.bbox}")
            
            excluded_count = 0
            fragment_count = 0
            
            import re
            
            for region in all_regions:
                region_bbox_tuple = tuple(region.bbox)
                
                # 1. 排除印章文字区域本身
                if region_bbox_tuple in seal_text_bboxes:
                    excluded_count += 1
                    logger.info(f"  ✅ 排除印章文字区域: {region.text[:30]}... bbox={region.bbox}")
                    continue
                
                # 2. 激进过滤：排除印章附近的小文本片段（可能是日期碎片）
                # 这些片段通常是OCR误检测的，不应该阻止日期翻译放置在边框区域
                text = region.text.strip()
                text_no_space = re.sub(r'\s+', '', text)
                r_x1, r_y1, r_x2, r_y2 = region.bbox
                region_height = r_y2 - r_y1
                region_width = r_x2 - r_x1
                
                # 判断是否是小文本片段
                is_small_text = (
                    region_height < 50 or  # 高度小于50px
                    region_width < 100 or  # 宽度小于100px
                    len(text_no_space) < 15  # 文本长度小于15个字符
                )
                
                # 判断是否是日期相关的片段
                is_date_fragment = bool(re.match(r'^[\d\s年月日]+$', text_no_space))
                
                # 判断是否在印章附近（印章下方300px范围内）
                is_near_seal = False
                for seal_bbox in seal_positions:
                    s_x1, s_y1, s_x2, s_y2 = seal_bbox
                    # 检查是否在印章下方300px范围内
                    if r_y1 >= s_y1 and r_y1 <= s_y2 + 300:
                        # 并且水平位置有重叠或接近
                        if not (r_x2 < s_x1 - 100 or r_x1 > s_x2 + 100):
                            is_near_seal = True
                            break
                
                # 修改过滤逻辑：只过滤真正的小碎片，不过滤完整的日期
                # 完整日期的特征：
                # 1. 包含"年"和"月"和"日"（如"2018年04月17日"）
                # 2. 或者宽度较大（>= 200px，说明是完整文本）
                is_complete_date = (
                    ('年' in text and '月' in text and '日' in text) or
                    region_width >= 200
                )
                
                # 如果是印章附近的小日期片段，但不是完整日期，则排除
                # 完整日期应该被识别为印章文字并翻译
                if is_small_text and is_date_fragment and is_near_seal and not is_complete_date:
                    fragment_count += 1
                    logger.info(
                        f"  🗑️  排除日期碎片: '{text}' bbox={region.bbox} "
                        f"(height={region_height}, width={region_width}, len={len(text_no_space)}, "
                        f"is_small={is_small_text}, is_date={is_date_fragment}, near_seal={is_near_seal})"
                    )
                    continue
                
                # 其他区域加入已占用区域
                occupied_regions.append(region.bbox)
            
            logger.info(
                f"🔍 建立已占用区域: 总区域={len(all_regions)}, "
                f"印章文字区域={len(seal_text_regions)}, "
                f"排除了{excluded_count}个印章区域, "
                f"排除了{fragment_count}个日期碎片, "
                f"已占用区域={len(occupied_regions)}"
            )
        
        # 获取印章文字翻译的字体大小配置
        seal_font_size = self.config.get('rendering.seal_text_handling.font_size', 12)
        
        # 按照文本类型排序：seal_overlap（日期）优先处理，seal_inner（印章机关）后处理
        # 这样可以确保日期先占据印章上方位置，印章机关文字可以根据日期位置智能排列
        def get_sort_key(region):
            text_type = getattr(region, 'seal_text_type', 'seal_overlap')
            # seal_overlap = 0 (优先), seal_inner = 1 (后处理)
            return 0 if text_type == 'seal_overlap' else 1
        
        sorted_seal_text_regions = sorted(seal_text_regions, key=get_sort_key)
        
        logger.info(
            f"🔍 处理顺序: "
            f"{sum(1 for r in sorted_seal_text_regions if getattr(r, 'seal_text_type', '') == 'seal_overlap')} 个日期区域优先, "
            f"{sum(1 for r in sorted_seal_text_regions if getattr(r, 'seal_text_type', '') == 'seal_inner')} 个印章机关区域后处理"
        )
        
        for i, region in enumerate(sorted_seal_text_regions):
            logger.info(f"🔍 Processing seal text region #{i}: '{region.text[:30]}...'")
            
            # Get the seal this region overlaps with
            seal_bbox = getattr(region, 'overlapping_seal', None)
            if not seal_bbox:
                logger.warning(f"Region '{region.text}' has no overlapping_seal attribute")
                continue
            
            logger.info(f"🔍 Region has overlapping_seal: {seal_bbox}")
            
            # 获取文本类型
            text_type = getattr(region, 'seal_text_type', 'seal_overlap')
            
            # 获取实际翻译文本（如果提供）或使用保守估算
            region_id = id(region)
            if translations and region_id in translations:
                # 使用实际翻译文本
                estimated_translated_text = translations[region_id]
                logger.info(
                    f"🔍 使用实际翻译文本: '{estimated_translated_text[:50]}...'"
                )
            else:
                # 保守估算：中文->英文通常变长3-5倍，再加上可能的前缀
                # 对于印章内文字，还要考虑 "[Seal: ]" 前缀（8个字符）
                if text_type == 'seal_inner':
                    # 印章内文字：原文 * 4 + "[Seal: ]"
                    estimated_translated_text = f"[Seal: {region.text * 4}]"
                else:
                    # 日期文字：原文 * 3
                    estimated_translated_text = region.text * 3
                
                logger.info(
                    f"🔍 使用估算翻译文本（长度={len(estimated_translated_text)}）"
                )
            
            # 如果是印章机关文字(seal_inner)，检查是否已经有日期文字被放置
            # 如果有，传递日期文字的位置和区域信息，以便智能排列
            date_text_bbox = None
            date_text_region = None
            if text_type == 'seal_inner' and positioned_regions:
                # 查找已放置的日期文字
                for placed_region, placed_bbox in positioned_regions:
                    if getattr(placed_region, 'seal_text_type', '') == 'seal_overlap':
                        date_text_bbox = placed_bbox
                        date_text_region = placed_region
                        logger.info(
                            f"🔍 发现已放置的日期文字: bbox={date_text_bbox}, "
                            f"将尝试与日期文字智能排列（垂直或水平）"
                        )
                        break
            
            
            # 使用新的碰撞检测方法
            if use_collision_detection:
                new_bbox = self.find_safe_position_with_collision_detection(
                    translated_text=estimated_translated_text,
                    font_size=seal_font_size,
                    seal_bbox=seal_bbox,
                    occupied_regions=occupied_regions,
                    image_shape=image_shape,
                    text_type=text_type,
                    date_text_bbox=date_text_bbox,  # 传递日期文字位置
                    date_text_region=date_text_region,  # 传递日期文字区域对象
                    rendered_image=image,  # 传递已渲染的图片（用于精确空白检测）
                    registration_authority_bbox=registration_authority_bbox  # 传递"Registration Authority"位置
                )
            else:
                # 使用旧方法（向后兼容）
                logger.info(f"🔍 使用传统方法查找位置")
                new_bbox = self._find_empty_space_near_seal(
                    seal_bbox, region, all_regions, (height, width)
                )
            
            if new_bbox:
                positioned_regions.append((region, new_bbox))
                
                # 将新位置添加到已占用区域，供后续区域使用
                occupied_regions.append(new_bbox)
                
                logger.info(
                    f"✅ Found position for '{region.text}': "
                    f"original={region.bbox}, new={new_bbox}"
                )
            else:
                logger.warning(
                    f"❌ Could not find suitable position for '{region.text}'"
                )
        
        logger.info(f"🔍 find_translation_positions returning {len(positioned_regions)} positioned regions")
        return positioned_regions
    
    def _find_empty_space_near_seal(
        self,
        seal_bbox: Tuple[int, int, int, int],
        region: TextRegion,
        all_regions: List[TextRegion],
        image_shape: Tuple[int, int]
    ) -> Optional[Tuple[int, int, int, int]]:
        """Find empty space near a seal to place translation.
        
        Search order:
        1. PRIORITY: Check directly above and below seal first (most natural positions)
        2. If no space above/below, try left and right
        3. If still no space, expand search radius
        
        Args:
            seal_bbox: Seal bounding box (x1, y1, x2, y2)
            region: Text region to place
            all_regions: All regions for collision detection
            image_shape: Image dimensions (height, width)
            
        Returns:
            New bounding box (x1, y1, x2, y2) or None if no space found
        """
        logger.info(f"🔍 _find_empty_space_near_seal called: seal={seal_bbox}, region_text='{region.text[:30]}...', all_regions_count={len(all_regions)}")
        
        s_x1, s_y1, s_x2, s_y2 = seal_bbox
        seal_center_x = (s_x1 + s_x2) // 2
        seal_center_y = (s_y1 + s_y2) // 2
        seal_width = s_x2 - s_x1
        seal_height = s_y2 - s_y1
        
        # Estimate size needed for translation (assume English is 2x longer)
        region_width = region.width
        region_height = region.height
        estimated_width = int(region_width * 2)
        
        # Get text type
        text_type = getattr(region, 'seal_text_type', 'seal_overlap')
        
        # For seal_inner text, use configured font size to estimate height
        # Since the text will be rendered with a specific font size (e.g., 12px)
        # and may wrap to multiple lines, we need to estimate the actual rendered height
        if text_type == 'seal_inner':
            seal_font_size = self.config.get('rendering.seal_text_handling.font_size', 12)
            # Estimate number of lines based on text length and estimated width
            # Assume average character width is about 0.6 * font_size for English
            chars_per_line = max(1, int(estimated_width / (seal_font_size * 0.6)))
            # Add "[Seal: " prefix (7 chars) to text length
            total_chars = len(region.text) * 2 + 7  # Assume English is 2x longer + prefix
            num_lines = max(1, (total_chars + chars_per_line - 1) // chars_per_line)
            # Height = num_lines * font_size * 1.5 (line spacing)
            estimated_height = int(num_lines * seal_font_size * 1.5)
            logger.info(
                f"📏 Estimated height for seal_inner text: {estimated_height}px "
                f"(font_size={seal_font_size}, chars={total_chars}, "
                f"chars_per_line={chars_per_line}, lines={num_lines})"
            )
        else:
            # For seal_overlap text (dates), use original height
            estimated_height = region_height
        
        logger.debug(
            f"Searching for position near seal {seal_bbox}: "
            f"region_size=({region_width}, {region_height}), "
            f"estimated_size=({estimated_width}, {estimated_height}), "
            f"text_type={text_type}"
        )
        
        height, width = image_shape
        
        # PRIORITY POSITIONS: Check directly above and below seal first
        # 对于印章内文字（seal_inner），优先放在印章正上方（留一定间隙）
        # 对于日期文字（seal_overlap），优先放在印章下方
        if text_type == 'seal_inner':
            # 印章内文字：根据用户需求，优先放在"登记机关"(Registration Authority)上方一些的位置
            # 与印章重叠的日期放在"登记机关"下方的区域
            
            # 计算印章上方的安全位置
            # 使用更小的间隙，让文字更靠近印章上方（"登记机关"上方）
            safe_gap = max(10, self.min_distance // 3)  # 减小间隙，让文字更靠近印章
            above_y = s_y1 - safe_gap - estimated_height
            
            # 如果计算出的位置会超出上边界，调整到边界内（留5px边距）
            if above_y < 5:
                above_y = 5
                logger.info(f"⚠️  Adjusted above-seal position to avoid top boundary: y={above_y}")
            
            priority_positions = [
                # Directly above seal, centered with gap - HIGHEST PRIORITY
                # 印章正上方，居中对齐，这是"登记机关"上方的位置
                (
                    seal_center_x - estimated_width // 2,
                    above_y,
                    "above-seal-centered-registration-above"
                ),
                # Left of seal (vertically centered) - SECOND PRIORITY
                (
                    s_x1 - self.min_distance - estimated_width,
                    seal_center_y - estimated_height // 2,
                    "seal-left-center"
                ),
                # Right of seal (vertically centered) - THIRD PRIORITY
                (
                    s_x2 + self.min_distance,
                    seal_center_y - estimated_height // 2,
                    "seal-right-center"
                ),
                # Directly below seal (centered) - FOURTH PRIORITY
                (
                    seal_center_x - estimated_width // 2,
                    s_y2 + self.min_distance,
                    "below-seal-centered"
                ),
                # Left-bottom of seal - FIFTH PRIORITY
                (
                    s_x1 - self.min_distance - estimated_width,
                    s_y2 - estimated_height,
                    "seal-left-bottom"
                ),
                # Right-bottom of seal - SIXTH PRIORITY
                (
                    s_x2 + self.min_distance,
                    s_y2 - estimated_height,
                    "seal-right-bottom"
                ),
            ]
        else:
            # 日期文字：根据用户需求，优先放在"登记机关"(Registration Authority)下方的框出区域
            # 这个区域通常在印章下方，是营业执照右下角的空白框区域
            priority_positions = [
                # HIGHEST PRIORITY: 印章下方居中位置（"登记机关"下方的框出区域）
                # 这是用户指定的位置
                (
                    seal_center_x - estimated_width // 2,
                    s_y2 + self.min_distance,
                    "below-seal-centered-registration-area"
                ),
                # SECOND PRIORITY: 印章下方稍微偏右
                (
                    seal_center_x - estimated_width // 2 + 20,
                    s_y2 + self.min_distance,
                    "below-seal-right-offset"
                ),
                # THIRD PRIORITY: 印章下方稍微偏左
                (
                    seal_center_x - estimated_width // 2 - 20,
                    s_y2 + self.min_distance,
                    "below-seal-left-offset"
                ),
                # RIGHT-BOTTOM BORDER BOX - 备选位置
                # 右下角边框区域（营业执照"Registration Authority"下方的空白框）
                (
                    width - estimated_width - self.min_distance * 3,
                    height - estimated_height - self.min_distance * 3,
                    "right-bottom-border-box"
                ),
                # Below seal (left-aligned)
                (
                    s_x1,
                    s_y2 + self.min_distance,
                    "below-left-aligned"
                ),
                # Below seal (right-aligned)
                (
                    s_x2 - estimated_width,
                    s_y2 + self.min_distance,
                    "below-right-aligned"
                ),
                # Below-left corner
                (
                    s_x1 - self.min_distance - estimated_width,
                    s_y2 + self.min_distance,
                    "below-left-corner"
                ),
                # Below-right corner
                (
                    s_x2 + self.min_distance,
                    s_y2 + self.min_distance,
                    "below-right-corner"
                ),
                # Directly above seal (centered) - LOWER PRIORITY
                (
                    seal_center_x - estimated_width // 2,
                    s_y1 - self.min_distance - estimated_height,
                    "above-center-priority"
                ),
                # Left border region (vertically centered with seal)
                (
                    self.min_distance,
                    seal_center_y - estimated_height // 2,
                    "left-border-region"
                ),
                # Right border region (vertically centered with seal)
                (
                    width - estimated_width - self.min_distance,
                    seal_center_y - estimated_height // 2,
                    "right-border-region"
                ),
                # Bottom border region (horizontally centered with seal)
                (
                    seal_center_x - estimated_width // 2,
                    height - estimated_height - self.min_distance,
                    "bottom-border-region"
                ),
            ]
        
        # Try priority positions first (above and below)
        logger.info(f"🔍 Checking PRIORITY positions (above/below seal) for '{region.text[:30]}...', text_type={text_type}")
        for x1, y1, position_desc in priority_positions:
            x2 = x1 + estimated_width
            y2 = y1 + estimated_height
            
            logger.info(f"  🔍 Trying position '{position_desc}': ({x1}, {y1}, {x2}, {y2})")
            
            # Check if within image bounds
            if x1 < 0 or y1 < 0 or x2 > width or y2 > height:
                logger.info(f"  ❌ Position {position_desc} out of bounds")
                continue
            
            candidate_bbox = (x1, y1, x2, y2)
            
            # Check for collisions
            # For seal_inner text, skip collision check with other text regions
            is_seal_inner = (text_type == 'seal_inner')
            
            # 对于印章下方的位置，允许轻微重叠以更好地利用空间
            # For below-seal positions, allow minor overlap to better utilize space
            allow_minor_overlap = ('below' in position_desc.lower() or 'bottom' in position_desc.lower())
            
            if self._check_position_available(candidate_bbox, seal_bbox, all_regions, is_seal_inner, allow_minor_overlap):
                logger.info(
                    f"  ✅ Found empty space at PRIORITY position '{position_desc}': {candidate_bbox}"
                )
                return candidate_bbox
            else:
                logger.info(f"  ❌ Position {position_desc} has collision")
        
        # If priority positions don't work, try secondary positions (left/right)
        logger.info(f"🔍 Priority positions occupied, checking SECONDARY positions (left/right)")
        secondary_positions = [
            # Right of seal (vertically centered)
            (
                s_x2 + self.min_distance,
                seal_center_y - estimated_height // 2,
                "right-center"
            ),
            # Left of seal (vertically centered)
            (
                s_x1 - self.min_distance - estimated_width,
                seal_center_y - estimated_height // 2,
                "left-center"
            ),
            # Below seal (left-aligned)
            (
                s_x1,
                s_y2 + self.min_distance,
                "below-left"
            ),
            # Below seal (right-aligned)
            (
                s_x2 - estimated_width,
                s_y2 + self.min_distance,
                "below-right"
            ),
            # Above seal (left-aligned)
            (
                s_x1,
                s_y1 - self.min_distance - estimated_height,
                "above-left"
            ),
            # Above seal (right-aligned)
            (
                s_x2 - estimated_width,
                s_y1 - self.min_distance - estimated_height,
                "above-right"
            ),
        ]
        
        # Try secondary positions
        for x1, y1, position_desc in secondary_positions:
            x2 = x1 + estimated_width
            y2 = y1 + estimated_height
            
            # Check if within image bounds
            if x1 < 0 or y1 < 0 or x2 > width or y2 > height:
                logger.debug(f"  ❌ Position {position_desc} out of bounds")
                continue
            
            candidate_bbox = (x1, y1, x2, y2)
            
            # Check for collisions
            # For seal_inner text, skip collision check with other text regions
            is_seal_inner = (text_type == 'seal_inner')
            
            # 对于印章下方的位置，允许轻微重叠
            # For below-seal positions, allow minor overlap
            allow_minor_overlap = ('below' in position_desc.lower())
            
            if self._check_position_available(candidate_bbox, seal_bbox, all_regions, is_seal_inner, allow_minor_overlap):
                logger.info(
                    f"  ✅ Found empty space at SECONDARY position '{position_desc}': {candidate_bbox}"
                )
                return candidate_bbox
            else:
                logger.debug(f"  ❌ Position {position_desc} has collision")
        
        # If no position found in primary/secondary search, try expanding search radius
        logger.info(f"🔍 Primary/secondary positions occupied, trying EXPANDED search")
        return self._expanded_search(
            seal_bbox, estimated_width, estimated_height, 
            all_regions, image_shape, text_type
        )
    
    def _check_position_available(
        self,
        candidate_bbox: Tuple[int, int, int, int],
        seal_bbox: Tuple[int, int, int, int],
        all_regions: List[TextRegion],
        is_seal_inner: bool = False,
        allow_minor_overlap: bool = False
    ) -> bool:
        """Check if a position is available (no collisions).
        
        Uses strict collision detection - ANY overlap is considered a collision.
        
        For seal_inner text (text inside seal), we skip collision check with ORIGINAL text regions,
        but still check collision with the seal itself. This allows placing the translation above 
        the seal even if there's other original text nearby.
        
        Note: This method is used by the old positioning method. The new collision detection method
        handles occupied regions (including already-placed translations) separately.
        
        Args:
            candidate_bbox: Candidate position to check
            seal_bbox: Seal bounding box (to avoid)
            all_regions: All text regions (to avoid)
            is_seal_inner: Whether this is seal inner text (if True, skip text region collision check)
            allow_minor_overlap: If True, allow minor overlap (< 10% area) with text regions
            
        Returns:
            True if position is available, False otherwise
        """
        # Check collision with seal (avoid overlapping with seal) - STRICT
        if self._has_any_overlap(candidate_bbox, seal_bbox):
            logger.info(f"    [X] Collision with seal: candidate={candidate_bbox}, seal={seal_bbox}")
            return False
        
        # For seal inner text, skip collision check with other text regions
        # This allows placing the translation above the seal even if there's text nearby
        if is_seal_inner:
            logger.info(f"    [OK] Position available for seal_inner: {candidate_bbox} (skipped text collision check)")
            return True
        
        # Check collision with text regions
        for i, other_region in enumerate(all_regions):
            if allow_minor_overlap:
                # 允许轻微重叠（< 10%面积）- 用于印章下方位置
                # Allow minor overlap (< 10% area) - for below-seal positions
                overlap_area = self._calculate_overlap_area(candidate_bbox, other_region.bbox)
                candidate_area = (candidate_bbox[2] - candidate_bbox[0]) * (candidate_bbox[3] - candidate_bbox[1])
                overlap_ratio = overlap_area / candidate_area if candidate_area > 0 else 0
                
                if overlap_ratio > 0.1:  # 重叠超过10%才算碰撞
                    logger.info(
                        f"    [X] Significant collision with region #{i}: candidate={candidate_bbox}, "
                        f"region={other_region.bbox}, overlap_ratio={overlap_ratio:.2%}, text='{other_region.text[:30]}...'"
                    )
                    return False
                elif overlap_ratio > 0:
                    logger.info(
                        f"    [~] Minor overlap with region #{i}: overlap_ratio={overlap_ratio:.2%}, allowing"
                    )
            else:
                # 严格检测 - STRICT
                if self._has_any_overlap(candidate_bbox, other_region.bbox):
                    logger.info(
                        f"    [X] Collision with region #{i}: candidate={candidate_bbox}, "
                        f"region={other_region.bbox}, text='{other_region.text[:30]}...'"
                    )
                    return False
        
        logger.info(f"    [OK] Position available: {candidate_bbox} (checked against {len(all_regions)} regions)")
        return True
    
    def _calculate_overlap_area(
        self,
        bbox1: Tuple[int, int, int, int],
        bbox2: Tuple[int, int, int, int]
    ) -> float:
        """Calculate the overlap area between two bounding boxes.
        
        Args:
            bbox1: First bounding box (x1, y1, x2, y2)
            bbox2: Second bounding box (x1, y1, x2, y2)
            
        Returns:
            Overlap area in pixels
        """
        x1_1, y1_1, x2_1, y2_1 = bbox1
        x1_2, y1_2, x2_2, y2_2 = bbox2
        
        # Calculate intersection
        x1_i = max(x1_1, x1_2)
        y1_i = max(y1_1, y1_2)
        x2_i = min(x2_1, x2_2)
        y2_i = min(y2_1, y2_2)
        
        # Check if there's an intersection
        if x1_i < x2_i and y1_i < y2_i:
            return (x2_i - x1_i) * (y2_i - y1_i)
        else:
            return 0.0
    
    def _has_any_overlap(
        self,
        bbox1: Tuple[int, int, int, int],
        bbox2: Tuple[int, int, int, int]
    ) -> bool:
        """Check if two bounding boxes have ANY overlap (strict check).
        
        Unlike _check_overlap which uses a threshold, this method returns
        True if there is ANY overlap at all, no matter how small.
        
        Args:
            bbox1: First bounding box (x1, y1, x2, y2)
            bbox2: Second bounding box (x1, y1, x2, y2)
            
        Returns:
            True if boxes have any overlap
        """
        x1_1, y1_1, x2_1, y2_1 = bbox1
        x1_2, y1_2, x2_2, y2_2 = bbox2
        
        # Check if boxes overlap in both x and y dimensions
        x_overlap = not (x2_1 <= x1_2 or x2_2 <= x1_1)
        y_overlap = not (y2_1 <= y1_2 or y2_2 <= y1_1)
        
        return x_overlap and y_overlap
    
    def _expanded_search(
        self,
        seal_bbox: Tuple[int, int, int, int],
        width_needed: int,
        height_needed: int,
        all_regions: List[TextRegion],
        image_shape: Tuple[int, int],
        text_type: str = 'seal_overlap'
    ) -> Optional[Tuple[int, int, int, int]]:
        """Perform expanded search in a wider radius around seal.
        
        Args:
            seal_bbox: Seal bounding box
            width_needed: Width needed for translation
            height_needed: Height needed for translation
            all_regions: All regions for collision detection
            image_shape: Image dimensions
            text_type: Type of text ('seal_inner' or 'seal_overlap')
            
        Returns:
            New bounding box or None
        """
        s_x1, s_y1, s_x2, s_y2 = seal_bbox
        seal_center_x = (s_x1 + s_x2) // 2
        seal_center_y = (s_y1 + s_y2) // 2
        
        height, width = image_shape
        
        # Search in concentric circles around seal
        for radius in range(self.min_distance, self.search_radius, 20):
            # Try 8 directions around seal
            angles = [0, 45, 90, 135, 180, 225, 270, 315]
            
            for angle in angles:
                # Calculate position
                rad = np.radians(angle)
                x1 = int(seal_center_x + radius * np.cos(rad) - width_needed // 2)
                y1 = int(seal_center_y + radius * np.sin(rad) - height_needed // 2)
                x2 = x1 + width_needed
                y2 = y1 + height_needed
                
                # Check bounds
                if x1 < 0 or y1 < 0 or x2 > width or y2 > height:
                    continue
                
                candidate_bbox = (x1, y1, x2, y2)
                
                # Check if position is available
                # For seal_inner text, skip collision check with other text regions
                is_seal_inner = (text_type == 'seal_inner')
                if self._check_position_available(candidate_bbox, seal_bbox, all_regions, is_seal_inner):
                    logger.info(
                        f"  ✅ Found space in EXPANDED search: "
                        f"radius={radius}, angle={angle}°, bbox={candidate_bbox}"
                    )
                    return candidate_bbox
        
        logger.warning(f"  ❌ Could not find suitable position within search_radius={self.search_radius}")
        return None
    
    def _find_blank_areas_in_quadrant(
        self,
        regions: List[Any],
        seals: List[Tuple[int, int, int, int]],
        quadrant_bounds: Tuple[int, int, int, int],
        margin: int = 10,
        seal_shrink_ratio: float = 0.15,
        grid_size: int = 10,
        min_blank_area: int = 100
    ) -> List[Tuple[int, int, int, int]]:
        """在指定象限中查找空白区域（使用可视化脚本的逻辑）
        
        Args:
            regions: 文字区域列表
            seals: 印章边界框列表
            quadrant_bounds: 象限边界 (x_start, y_start, x_end, y_end)
            margin: 空白区域检测的边距
            seal_shrink_ratio: 印章边界框收缩比例
            grid_size: 网格大小（像素）
            min_blank_area: 最小空白区域面积（像素²）
        
        Returns:
            空白区域列表 [(x1, y1, x2, y2), ...]
        """
        x_start, y_start, x_end, y_end = quadrant_bounds
        
        # 收缩印章边界框
        def shrink_seal_bbox(seal, shrink_ratio):
            x1, y1, x2, y2 = seal
            width = x2 - x1
            height = y2 - y1
            shrink_x = int(width * shrink_ratio)
            shrink_y = int(height * shrink_ratio)
            return (x1 + shrink_x, y1 + shrink_y, x2 - shrink_x, y2 - shrink_y)
        
        # 过滤出在象限内或附近的区域
        regions_in_or_near = []
        for region in regions:
            if hasattr(region, 'bbox'):
                x1, y1, x2, y2 = region.bbox
            else:
                x1, y1, x2, y2 = region
            
            # 检查是否与象限有交集
            if not (x2 < x_start or x1 > x_end or y2 < y_start or y1 > y_end):
                regions_in_or_near.append((x1, y1, x2, y2))
        
        # 收缩印章并过滤出在象限内或附近的印章
        shrunk_seals = []
        for seal in seals:
            x1, y1, x2, y2 = seal
            # 检查是否与象限有交集
            if not (x2 < x_start or x1 > x_end or y2 < y_start or y1 > y_end):
                shrunk_seal = shrink_seal_bbox(seal, seal_shrink_ratio)
                shrunk_seals.append(shrunk_seal)
        
        # 创建占用矩阵
        quadrant_width = x_end - x_start
        quadrant_height = y_end - y_start
        occupied = np.zeros((quadrant_height, quadrant_width), dtype=np.uint8)
        
        # 标记文字区域
        for x1, y1, x2, y2 in regions_in_or_near:
            rel_x1 = max(0, x1 - x_start - margin)
            rel_y1 = max(0, y1 - y_start - margin)
            rel_x2 = min(quadrant_width, x2 - x_start + margin)
            rel_y2 = min(quadrant_height, y2 - y_start + margin)
            
            if rel_x2 > rel_x1 and rel_y2 > rel_y1:
                occupied[rel_y1:rel_y2, rel_x1:rel_x2] = 1
        
        # 标记印章区域（使用收缩后的边界框）
        for x1, y1, x2, y2 in shrunk_seals:
            rel_x1 = max(0, x1 - x_start - margin)
            rel_y1 = max(0, y1 - y_start - margin)
            rel_x2 = min(quadrant_width, x2 - x_start + margin)
            rel_y2 = min(quadrant_height, y2 - y_start + margin)
            
            if rel_x2 > rel_x1 and rel_y2 > rel_y1:
                occupied[rel_y1:rel_y2, rel_x1:rel_x2] = 1
        
        # 查找空白区域
        blank_areas = []
        for y in range(0, quadrant_height, grid_size):
            for x in range(0, quadrant_width, grid_size):
                grid_x2 = min(x + grid_size, quadrant_width)
                grid_y2 = min(y + grid_size, quadrant_height)
                
                # 检查这个网格是否完全空白
                grid_region = occupied[y:grid_y2, x:grid_x2]
                if np.sum(grid_region) == 0:
                    area = (grid_x2 - x) * (grid_y2 - y)
                    if area >= min_blank_area:
                        # 转换回图片坐标
                        abs_x1 = x + x_start
                        abs_y1 = y + y_start
                        abs_x2 = grid_x2 + x_start
                        abs_y2 = grid_y2 + y_start
                        blank_areas.append((abs_x1, abs_y1, abs_x2, abs_y2))
        
        # 合并相邻的空白区域
        return self._merge_adjacent_blank_areas(blank_areas, grid_size)
    
    def _merge_adjacent_blank_areas(
        self,
        areas: List[Tuple[int, int, int, int]],
        tolerance: int = 5
    ) -> List[Tuple[int, int, int, int]]:
        """合并相邻的空白区域
        
        Args:
            areas: 空白区域列表
            tolerance: 合并容差（像素）
        
        Returns:
            合并后的空白区域列表
        """
        if not areas:
            return []
        
        sorted_areas = sorted(areas, key=lambda a: (a[1], a[0]))
        merged = []
        
        for area in sorted_areas:
            x1, y1, x2, y2 = area
            merged_with_existing = False
            
            for i, (mx1, my1, mx2, my2) in enumerate(merged):
                # 检查是否可以水平合并
                if abs(y1 - my1) <= tolerance and abs(y2 - my2) <= tolerance:
                    if abs(x1 - mx2) <= tolerance:
                        merged[i] = (mx1, my1, x2, my2)
                        merged_with_existing = True
                        break
                    elif abs(x2 - mx1) <= tolerance:
                        merged[i] = (x1, my1, mx2, my2)
                        merged_with_existing = True
                        break
                
                # 检查是否可以垂直合并
                if abs(x1 - mx1) <= tolerance and abs(x2 - mx2) <= tolerance:
                    if abs(y1 - my2) <= tolerance:
                        merged[i] = (mx1, my1, mx2, y2)
                        merged_with_existing = True
                        break
                    elif abs(y2 - my1) <= tolerance:
                        merged[i] = (mx1, y1, mx2, my2)
                        merged_with_existing = True
                        break
            
            if not merged_with_existing:
                merged.append(area)
        
        return merged
    
    def _check_collision(
        self,
        bbox: Tuple[int, int, int, int],
        occupied_regions: List[Tuple[int, int, int, int]]
    ) -> bool:
        """检查边界框是否与已占用区域冲突
        
        Args:
            bbox: 要检查的边界框 (x1, y1, x2, y2)
            occupied_regions: 已占用区域列表
        
        Returns:
            True 如果有冲突，False 如果没有冲突
        """
        x1, y1, x2, y2 = bbox
        
        for ox1, oy1, ox2, oy2 in occupied_regions:
            # 检查是否有重叠
            if not (x2 <= ox1 or x1 >= ox2 or y2 <= oy1 or y1 >= oy2):
                return True  # 有冲突
        
        return False  # 没有冲突
    
    def _detect_rendered_text_regions(
        self,
        image: np.ndarray,
        background_color: Tuple[int, int, int] = (255, 255, 255)
    ) -> List[Tuple[int, int, int, int]]:
        """从渲染后的图片中检测文字区域
        
        通过检测与背景色不同的区域来识别已渲染的文字
        
        Args:
            image: 渲染后的图片
            background_color: 背景颜色（默认白色）
        
        Returns:
            文字区域列表 [(x1, y1, x2, y2), ...]
        """
        # 转换为灰度图
        if len(image.shape) == 3:
            gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        else:
            gray = image.copy()
        
        # 使用Otsu自动阈值法，这是最可靠的方法
        # 它会自动找到最佳阈值来分离前景和背景
        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
        
        # 形态学操作：先腐蚀去除噪点，再膨胀连接文字
        kernel_small = np.ones((2, 2), np.uint8)
        kernel_large = np.ones((3, 3), np.uint8)
        
        # 腐蚀去除小噪点
        eroded = cv2.erode(binary, kernel_small, iterations=1)
        # 膨胀连接相近的文字
        dilated = cv2.dilate(eroded, kernel_large, iterations=3)
        
        # 查找轮廓
        contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        # 提取边界框
        text_regions = []
        image_height, image_width = image.shape[:2]
        
        for contour in contours:
            x, y, w, h = cv2.boundingRect(contour)
            
            # 过滤条件：
            # 1. 太小的区域（可能是噪点）- 至少20x20像素
            # 2. 太大的区域（接近整个图片的区域，可能是误检测）
            area = w * h
            image_area = image_width * image_height
            
            # 如果区域面积超过图片面积的50%，很可能是误检测
            if area > image_area * 0.5:
                logger.debug(
                    f"⚠️ 过滤掉过大区域: ({x}, {y}) -> ({x+w}, {y+h}), "
                    f"面积={area}px² ({area/image_area*100:.1f}% of image)"
                )
                continue
            
            # 过滤太小的区域（至少20x20）
            if w >= 20 and h >= 20:
                text_regions.append((x, y, x + w, y + h))
        
        logger.info(f"📍 从渲染图片检测到 {len(text_regions)} 个文字区域")
        
        # 打印前几个区域用于调试
        for i, (x1, y1, x2, y2) in enumerate(text_regions[:10]):
            logger.debug(
                f"  区域 #{i+1}: ({x1}, {y1}) -> ({x2}, {y2}), "
                f"尺寸={x2-x1}x{y2-y1}px"
            )
        
        return text_regions
    
    def check_text_seal_overlap_and_wrap(
        self,
        translated_text: str,
        font_size: int,
        position_bbox: Tuple[int, int, int, int],
        seal_bbox: Tuple[int, int, int, int],
        max_width: Optional[int] = None
    ) -> Tuple[str, int, Tuple[int, int, int, int]]:
        """检测文本是否与印章重叠，如果重叠则自动换行。
        
        Args:
            translated_text: 翻译后的文本
            font_size: 字体大小
            position_bbox: 文本的位置边界框 (x1, y1, x2, y2)
            seal_bbox: 印章的边界框 (x1, y1, x2, y2)
            max_width: 最大宽度限制（可选，如果不提供则使用position_bbox的宽度）
            
        Returns:
            (wrapped_text, adjusted_font_size, adjusted_bbox): 
            - wrapped_text: 换行后的文本（可能包含换行符）
            - adjusted_font_size: 调整后的字体大小
            - adjusted_bbox: 调整后的边界框
        """
        from src.rendering.text_renderer import TextRenderer
        
        # 创建临时的TextRenderer来测量文本
        temp_renderer = TextRenderer(self.config)
        
        # 计算文本的实际渲染宽度
        text_width = temp_renderer.measure_text_width(translated_text, font_size)
        text_height = temp_renderer.measure_text_height(translated_text, font_size)
        
        x1, y1, x2, y2 = position_bbox
        
        # 检查是否与印章重叠
        if not self._has_any_overlap(position_bbox, seal_bbox):
            # 没有重叠，直接返回原文本
            logger.info(f"✅ 文本不与印章重叠: text='{translated_text[:50]}...'")
            return translated_text, font_size, position_bbox
        
        logger.info(
            f"⚠️ 检测到文本与印章重叠: text='{translated_text[:50]}...', "
            f"text_bbox={position_bbox}, seal_bbox={seal_bbox}"
        )
        
        # 确定最大宽度
        if max_width is None:
            max_width = x2 - x1
        
        # 尝试换行以避免重叠
        # 策略：逐步减小每行的最大宽度，直到不再与印章重叠
        for width_ratio in [0.9, 0.8, 0.7, 0.6, 0.5, 0.4]:
            adjusted_max_width = int(max_width * width_ratio)
            
            # 使用TextRenderer的wrap_text方法进行换行
            wrapped_lines = temp_renderer.wrap_text(
                translated_text, 
                adjusted_max_width, 
                font_size
            )
            
            if not wrapped_lines:
                continue
            
            # 计算换行后的文本尺寸
            line_height = int(font_size * temp_renderer.DEFAULT_LINE_SPACING_RATIO)
            total_height = line_height * len(wrapped_lines)
            
            # 计算每行的最大宽度
            max_line_width = max(
                temp_renderer.measure_text_width(line, font_size) 
                for line in wrapped_lines
            )
            
            # 计算新的边界框
            # 如果文本在印章上方，向上移动以避免重叠
            # 如果文本在印章下方，保持原位置
            seal_y1 = seal_bbox[1]
            if y1 < seal_y1:
                # 文本在印章上方，向上移动
                new_y1 = seal_y1 - 10 - total_height  # 印章上方10px
            else:
                # 文本在印章下方或内部，保持原Y坐标
                new_y1 = y1
            
            new_bbox = (x1, new_y1, x1 + max_line_width, new_y1 + total_height)
            
            # 检查新边界框是否还与印章重叠
            if not self._has_any_overlap(new_bbox, seal_bbox):
                wrapped_text = '\n'.join(wrapped_lines)
                logger.info(
                    f"✅ 换行成功避免重叠: lines={len(wrapped_lines)}, "
                    f"width_ratio={width_ratio}, new_bbox={new_bbox}"
                )
                return wrapped_text, font_size, new_bbox
            
            logger.debug(
                f"  尝试 width_ratio={width_ratio}: lines={len(wrapped_lines)}, "
                f"仍有重叠，继续尝试..."
            )
        
        # 如果换行仍然无法避免重叠，尝试减小字体大小
        logger.info("⚠️ 换行无法完全避免重叠，尝试减小字体大小...")
        
        for reduced_font_size in range(font_size - 1, max(8, font_size - 6), -1):
            # 重新计算文本尺寸
            text_width = temp_renderer.measure_text_width(translated_text, reduced_font_size)
            text_height = temp_renderer.measure_text_height(translated_text, reduced_font_size)
            
            # 如果文本在印章上方，向上移动
            seal_y1 = seal_bbox[1]
            if y1 < seal_y1:
                new_y1 = seal_y1 - 10 - text_height
            else:
                new_y1 = y1
            
            new_bbox = (x1, new_y1, x1 + text_width, new_y1 + text_height)
            
            if not self._has_any_overlap(new_bbox, seal_bbox):
                logger.info(
                    f"✅ 减小字体大小成功避免重叠: "
                    f"font_size={font_size}->{reduced_font_size}, new_bbox={new_bbox}"
                )
                return translated_text, reduced_font_size, new_bbox
        
        # 如果所有方法都失败，返回原文本但记录警告
        logger.warning(
            f"❌ 无法通过换行或减小字体避免重叠: text='{translated_text[:50]}...'"
        )
        return translated_text, font_size, position_bbox
    
    def find_safe_position_with_collision_detection(
        self,
        translated_text: str,
        font_size: int,
        seal_bbox: Tuple[int, int, int, int],
        occupied_regions: List[Tuple[int, int, int, int]],
        image_shape: Tuple[int, int],
        text_type: str = 'seal_overlap',
        date_text_bbox: Optional[Tuple[int, int, int, int]] = None,
        date_text_region: Optional[Any] = None,
        rendered_image: Optional[np.ndarray] = None,  # 新增：已渲染翻译的图片
        registration_authority_bbox: Optional[Tuple[int, int, int, int]] = None  # 新增："Registration Authority"位置
    ) -> Optional[Tuple[int, int, int, int]]:
        """通过碰撞检测找到安全的绘制位置。
        
        改进策略：
        1. 计算翻译文本的实际尺寸（宽度和高度）
        2. 建立已占用区域集合（包括印章、已渲染的文本等）
        3. 在限定搜索区域内，通过矩形碰撞检测找到不与任何已占区域重叠的安全位置
        4. 对于印章机关文字，如果日期文字已放置，优先尝试放在日期文字的正上方
        5. 优先级：
           - seal_inner: Registration Authority上方 > 印章上方 > 印章左侧
           - seal_overlap: Registration Authority下方 > 印章下方 > 印章上方
        
        Args:
            translated_text: 翻译后的文本
            font_size: 字体大小
            seal_bbox: 印章边界框 (x1, y1, x2, y2)
            occupied_regions: 已占用区域列表 [(x1, y1, x2, y2), ...]
            image_shape: 图片尺寸 (height, width)
            text_type: 文本类型 ('seal_inner' 或 'seal_overlap')
            registration_authority_bbox: "Registration Authority"文本的边界框（可选）
            
        Returns:
            安全位置的边界框 (x1, y1, x2, y2)，如果找不到则返回 None
        """
        print(f"\n[COLLISION_DETECTION] Method called!")
        print(f"[COLLISION_DETECTION] text_type={text_type}")
        print(f"[COLLISION_DETECTION] translated_text='{translated_text[:50]}...'")
        print(f"[COLLISION_DETECTION] seal_bbox={seal_bbox}")
        print(f"[COLLISION_DETECTION] registration_authority_bbox={registration_authority_bbox}")
        print(f"[COLLISION_DETECTION] occupied_regions count={len(occupied_regions)}")
        
        from src.rendering.text_renderer import TextRenderer
        
        # 创建临时的 TextRenderer 来测量文本尺寸
        temp_renderer = TextRenderer(self.config)
        
        # 计算翻译文本的实际尺寸
        text_width = temp_renderer.measure_text_width(translated_text, font_size)
        text_height = temp_renderer.measure_text_height(translated_text, font_size)
        
        # 对于印章内文字(seal_inner),限制最大宽度以支持换行
        # 实际渲染时文本会换行,所以这里使用换行后的宽度
        # 使用更合理的最大宽度，确保文本可以在Registration Authority上方合适地换行
        max_seal_text_width = 300  # 减少到300px，使文本更容易换行
        if text_type == 'seal_inner' and text_width > max_seal_text_width:
            # 限制最大宽度，文本会自动换行
            width_needed = max_seal_text_width + 10
            # 估算换行后的行数（使用更紧凑的行间距）
            import math
            estimated_lines = math.ceil(text_width / max_seal_text_width)
            # 使用1.3倍行高（更紧凑）
            height_needed = int(font_size * estimated_lines * 1.3) + 10
            print(f"[COLLISION_DETECTION] Seal inner text too long, limiting width: text_width={text_width} -> width_needed={width_needed}, estimated_lines={estimated_lines}, height_needed={height_needed}")
        else:
            # 添加安全边距（10px）
            width_needed = text_width + 10
            height_needed = text_height + 10
        
        print(f"[COLLISION_DETECTION] text_width={text_width}, text_height={text_height}")
        print(f"[COLLISION_DETECTION] width_needed={width_needed}, height_needed={height_needed}")
        
        logger.info(
            f"📏 计算文本尺寸: text='{translated_text[:30]}...', "
            f"font_size={font_size}, width={text_width}px, height={text_height}px, "
            f"width_needed={width_needed}px, height_needed={height_needed}px"
        )
        
        s_x1, s_y1, s_x2, s_y2 = seal_bbox
        seal_center_x = (s_x1 + s_x2) // 2
        seal_center_y = (s_y1 + s_y2) // 2
        seal_width = s_x2 - s_x1
        seal_height = s_y2 - s_y1
        
        height, width = image_shape
        
        # ========== 印章内文字的候选位置策略 ==========
        # 简化策略：直接将印章内文字放在"登记机关"(Registration Authority)上方
        if text_type == 'seal_inner':
            logger.info("🔍 为印章内文字查找翻译位置（直接放在Registration Authority上方）")
            
            # === 最高优先级：Registration Authority上方 ===
            if registration_authority_bbox:
                ra_x1, ra_y1, ra_x2, ra_y2 = registration_authority_bbox
                ra_center_x = (ra_x1 + ra_x2) // 2
                
                logger.info(f"✅ 找到Registration Authority位置: {registration_authority_bbox}")
                
                # 计算在Registration Authority上方的位置
                safe_gap = 5  # Registration Authority上方的间距
                above_ra_y = ra_y1 - safe_gap - height_needed
                
                logger.info(f"📏 计算位置: ra_y1={ra_y1}, safe_gap={safe_gap}, height_needed={height_needed}, above_ra_y={above_ra_y}")
                
                if above_ra_y >= 20:  # 确保不会超出图片顶部
                    logger.info(f"🎯 直接放在Registration Authority上方 (y={above_ra_y})")
                    
                    # 与Registration Authority左边界对齐
                    x_pos = ra_x1
                    candidate_bbox = (x_pos, above_ra_y, x_pos + width_needed, above_ra_y + height_needed)
                    
                    # 检查x坐标是否在图片范围内
                    if x_pos >= 20 and x_pos + width_needed <= width - 20:
                        logger.info(f"✅ 印章内文字与Registration Authority左边界对齐: {candidate_bbox}")
                        return candidate_bbox
                    else:
                        logger.warning(f"⚠️ 左对齐位置超出图片范围，尝试其他位置")
                        
                        # 如果左对齐超出范围，尝试居中对齐
                        x_pos = ra_center_x - width_needed // 2
                        candidate_bbox = (x_pos, above_ra_y, x_pos + width_needed, above_ra_y + height_needed)
                        
                        if x_pos >= 20 and x_pos + width_needed <= width - 20:
                            logger.info(f"✅ 印章内文字居中对齐（备选）: {candidate_bbox}")
                            return candidate_bbox
                    
                    logger.warning("⚠️ Registration Authority上方没有有效的x位置")
                else:
                    logger.warning(f"⚠️ Registration Authority上方空间不足 (y={above_ra_y} < 20)")
            else:
                logger.warning("⚠️ 未提供Registration Authority位置")
        
        # ========== 日期文字的候选位置策略 ==========
        # 简化策略：直接将日期文字放在"登记机关"(Registration Authority)下方
        elif text_type == 'seal_overlap':
            logger.info("🔍 为日期文字查找翻译位置（直接放在Registration Authority下方）")
            
            # === 最高优先级：Registration Authority下方 ===
            if registration_authority_bbox:
                ra_x1, ra_y1, ra_x2, ra_y2 = registration_authority_bbox
                ra_center_x = (ra_x1 + ra_x2) // 2
                
                logger.info(f"✅ 找到Registration Authority位置: {registration_authority_bbox}")
                
                # 计算在Registration Authority下方的位置
                safe_gap = 5  # Registration Authority下方的间距
                below_ra_y = ra_y2 + safe_gap
                
                logger.info(f"📏 计算位置: ra_y2={ra_y2}, safe_gap={safe_gap}, below_ra_y={below_ra_y}")
                
                if below_ra_y + height_needed <= height - 20:  # 确保不会超出图片底部
                    logger.info(f"🎯 直接放在Registration Authority下方 (y={below_ra_y})")
                    
                    # 与Registration Authority左边界对齐
                    x_pos = ra_x1
                    candidate_bbox = (x_pos, below_ra_y, x_pos + width_needed, below_ra_y + height_needed)
                    
                    # 检查x坐标是否在图片范围内
                    if x_pos >= 20 and x_pos + width_needed <= width - 20:
                        logger.info(f"✅ 日期文字与Registration Authority左边界对齐: {candidate_bbox}")
                        return candidate_bbox
                    else:
                        logger.warning(f"⚠️ 左对齐位置超出图片范围，尝试其他位置")
                        
                        # 如果左对齐超出范围，尝试居中对齐
                        x_pos = ra_center_x - width_needed // 2
                        candidate_bbox = (x_pos, below_ra_y, x_pos + width_needed, below_ra_y + height_needed)
                        
                        if x_pos >= 20 and x_pos + width_needed <= width - 20:
                            logger.info(f"✅ 日期文字居中对齐（备选）: {candidate_bbox}")
                            return candidate_bbox
                    
                    logger.warning("⚠️ Registration Authority下方没有有效的x位置")
                else:
                    logger.warning(f"⚠️ Registration Authority下方空间不足")
            else:
                logger.warning("⚠️ 未提供Registration Authority位置")
            right_half_start = width // 2
            bottom_half_start = height // 2
            quadrant_bounds = (right_half_start, bottom_half_start, width, height)
            
            logger.info(
                f"📍 右下角象限: x=[{right_half_start}, {width}], y=[{bottom_half_start}, {height}]"
            )
            
            # 使用可视化脚本的逻辑查找空白区域
            blank_areas = self._find_blank_areas_in_quadrant(
                regions=actual_text_regions,  # 使用实际渲染的文字区域
                seals=all_seals,
                quadrant_bounds=quadrant_bounds,
                margin=10,
                seal_shrink_ratio=0.15,
                grid_size=10,
                min_blank_area=100
            )
            
            logger.info(f"📍 在右下角象限找到 {len(blank_areas)} 个空白区域")
            
            # 从空白区域中选择合适的位置
            # 按面积排序，优先选择较大的空白区域
            blank_areas_sorted = sorted(
                blank_areas,
                key=lambda area: (area[2] - area[0]) * (area[3] - area[1]),
                reverse=True
            )
            
            # 打印前几个空白区域用于调试
            for i, (x1, y1, x2, y2) in enumerate(blank_areas_sorted[:5]):
                area = (x2 - x1) * (y2 - y1)
                logger.info(f"  空白区域 #{i+1}: ({x1}, {y1}) -> ({x2}, {y2}), 面积={area}px²")
            
            # 尝试在每个空白区域中放置文本
            for blank_x1, blank_y1, blank_x2, blank_y2 in blank_areas_sorted:
                blank_width = blank_x2 - blank_x1
                blank_height = blank_y2 - blank_y1
                
                # 检查空白区域是否足够大
                if blank_width >= width_needed and blank_height >= height_needed:
                    # 在空白区域内居中放置文本
                    text_x = blank_x1 + (blank_width - width_needed) // 2
                    text_y = blank_y1 + (blank_height - height_needed) // 2
                    
                    candidate_bbox = (text_x, text_y, text_x + width_needed, text_y + height_needed)
                    
                    # 检查是否与已占用区域冲突
                    if not self._check_collision(candidate_bbox, actual_text_regions):
                        logger.info(
                            f"✅ 在空白区域找到合适位置: "
                            f"blank_area=({blank_x1}, {blank_y1}, {blank_x2}, {blank_y2}), "
                            f"text_bbox={candidate_bbox}, "
                            f"blank_size={blank_width}x{blank_height}px"
                        )
                        return candidate_bbox
                    else:
                        logger.debug(
                            f"⚠️ 空白区域位置与已占用区域冲突: "
                            f"blank_area=({blank_x1}, {blank_y1}, {blank_x2}, {blank_y2})"
                        )
                else:
                    logger.debug(
                        f"⚠️ 空白区域太小: "
                        f"blank_size={blank_width}x{blank_height}px, "
                        f"needed={width_needed}x{height_needed}px"
                    )
            
            logger.info("⚠️ 空白区域检测未找到合适位置，回退到传统候选位置算法")
        
        # ========== 原有的候选位置生成逻辑（作为后备方案）==========
        
        # 定义候选位置（按优先级排序）
        if text_type == 'seal_inner':
            # 印章机关文字：如果前面的Registration Authority上方和空白区域都失败了
            # 则尝试其他候选位置
            
            candidate_positions = []
            
            # === 第1优先级：印章上方（PRIORITY1）===
            
            seal_top_y = s_y1
            safe_gap = max(10, self.min_distance // 3)
            above_seal_y = seal_top_y - safe_gap - height_needed
            
            if above_seal_y >= 20:  # 确保不会超出图片顶部
                logger.info(f"🔍 第1优先级：印章上方区域 (y < {seal_top_y})")
                
                candidate_positions.extend([
                    # 与印章居中对齐
                    (
                        seal_center_x - width_needed // 2,
                        above_seal_y,
                        "above-seal-centered-PRIORITY1"
                    ),
                    # 与印章左对齐
                    (
                        s_x1,
                        above_seal_y,
                        "above-seal-left-aligned-PRIORITY1"
                    ),
                    # 与印章右对齐
                    (
                        s_x2 - width_needed,
                        above_seal_y,
                        "above-seal-right-aligned-PRIORITY1"
                    ),
                ])
                
                # 在印章上方，从右到左尝试多个x位置
                for x_ratio in [0.9, 0.8, 0.7, 0.6, 0.5, 0.4, 0.3, 0.2]:
                    x_pos = int(width * x_ratio) - width_needed // 2
                    if x_pos >= 20 and x_pos + width_needed <= width - 20:
                        candidate_positions.append((
                            x_pos,
                            above_seal_y,
                            f"above-seal-x-ratio-{int(x_ratio*100)}-PRIORITY1"
                        ))
            
            # === 第2优先级：印章左侧空白区域（PRIORITY2）===
            
            logger.info(f"🔍 第2优先级：印章左侧空白区域 (x < {s_x1})")
            
            # 在印章左侧，尝试多个位置
            # 从印章左边界向左，尝试不同的x坐标
            for x_offset in [40, 60, 80, 100, 120, 150, 200, 250, 300, 350]:  # 距离印章左边界的偏移
                x_pos = s_x1 - x_offset - width_needed
                if x_pos >= 30:  # 确保不会超出左边界
                    # 尝试不同的y坐标（与印章对齐）
                    # 优先级：顶部 > 中间偏上 > 中间（避免向下延伸重叠）
                    candidate_positions.extend([
                        # 与印章顶部对齐（最高优先级，向下延伸空间最大）
                        (
                            x_pos,
                            s_y1,
                            f"seal-left-blank-x{x_offset}-top-PRIORITY2"
                        ),
                        # 与印章中间偏上对齐
                        (
                            x_pos,
                            s_y1 + (seal_height // 3),  # 印章上1/3位置
                            f"seal-left-blank-x{x_offset}-upper-third-PRIORITY2"
                        ),
                        # 与印章垂直居中对齐
                        (
                            x_pos,
                            seal_center_y - height_needed // 2,
                            f"seal-left-blank-x{x_offset}-center-PRIORITY2"
                        ),
                    ])
            
            # 也尝试靠近左边界的位置
            for x_pos in [50, 100, 150, 200, 250, 300]:
                if x_pos + width_needed < s_x1 - 20:  # 确保不会太靠近印章
                    candidate_positions.extend([
                        # 与印章顶部对齐
                        (
                            x_pos,
                            s_y1,
                            f"left-edge-x{x_pos}-top-PRIORITY2"
                        ),
                        # 与印章垂直居中对齐
                        (
                            x_pos,
                            seal_center_y - height_needed // 2,
                            f"left-edge-x{x_pos}-center-PRIORITY2"
                        ),
                    ])
            
            # === 第3优先级：右下角象限（PRIORITY3）===
            # 右下角象限定义：x > width/2 且 y > height/2
            
            # 计算右下角象限的候选位置
            right_half_start = width // 2
            bottom_half_start = height // 2
            
            logger.info(
                f"🔍 第3优先级：右下角象限: "
                f"x范围=[{right_half_start}, {width}], "
                f"y范围=[{bottom_half_start}, {height}]"
            )
            
            # 在右下角象限内，尝试多个位置
            # 从右下角开始，逐步向左上移动
            for y_offset in [30, 60, 100, 150, 200]:  # 从下边界向上的偏移
                for x_offset in [30, 60, 100, 150, 200, 250]:  # 从右边界向左的偏移
                    x_pos = width - width_needed - x_offset
                    y_pos = height - height_needed - y_offset
                    
                    # 确保在右下角象限内
                    if x_pos >= right_half_start and y_pos >= bottom_half_start:
                        candidate_positions.append((
                            x_pos,
                            y_pos,
                            f"right-bottom-quadrant-offset-x{x_offset}-y{y_offset}-PRIORITY2"
                        ))
            
            # === 第4优先级：印章下方（PRIORITY3）===
            # 不限制x坐标，只要y坐标在印章下方即可
            
            seal_bottom_y = s_y2
            below_seal_y = seal_bottom_y + self.min_distance
            
            if below_seal_y + height_needed <= height - 20:  # 确保不会超出图片底部
                logger.info(f"🔍 第4优先级：印章下方区域 (y > {seal_bottom_y})")
                
                # 也尝试与印章对齐的位置
                candidate_positions.extend([
                    # 与印章居中对齐
                    (
                        seal_center_x - width_needed // 2,
                        below_seal_y,
                        "below-seal-centered-PRIORITY3"
                    ),
                    # 与印章左对齐
                    (
                        s_x1,
                        below_seal_y,
                        "below-seal-left-aligned-PRIORITY3"
                    ),
                    # 与印章右对齐
                    (
                        s_x2 - width_needed,
                        below_seal_y,
                        "below-seal-right-aligned-PRIORITY3"
                    ),
                ])
                
                # 在印章下方，从右到左尝试多个x位置
                for x_ratio in [0.9, 0.8, 0.7, 0.6, 0.5, 0.4, 0.3, 0.2]:
                    x_pos = int(width * x_ratio) - width_needed // 2
                    if x_pos >= 20 and x_pos + width_needed <= width - 20:
                        candidate_positions.append((
                            x_pos,
                            below_seal_y,
                            f"below-seal-x-ratio-{int(x_ratio*100)}-PRIORITY3"
                        ))
            
            # === 第5优先级：右半图其他空白地方（PRIORITY4）===
            # 在图片右半部分（x > width/2）尝试多个位置
            right_half_start = width // 2
            
            logger.info(f"🔍 第5优先级：右半图其他空白地方 (x > {right_half_start})")
            
            # 在右半图从上到下、从右到左扫描多个候选位置
            for y_ratio in [0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8]:
                for x_ratio in [0.9, 0.8, 0.7, 0.6]:
                    x_pos = int(width * x_ratio) - width_needed // 2
                    y_pos = int(height * y_ratio) - height_needed // 2
                    
                    # 确保在右半图内
                    if x_pos >= right_half_start and x_pos + width_needed <= width - 20:
                        if y_pos >= 20 and y_pos + height_needed <= height - 20:
                            candidate_positions.append((
                                x_pos,
                                y_pos,
                                f"right-half-blank-x{int(x_ratio*100)}-y{int(y_ratio*100)}-PRIORITY4"
                            ))
            
            # 也尝试一些固定的右半图位置
            candidate_positions.extend([
                # 右上角区域
                (
                    width - width_needed - 50,
                    50,
                    "right-half-top-PRIORITY4"
                ),
                # 右中区域
                (
                    width - width_needed - 50,
                    (height - height_needed) // 2,
                    "right-half-middle-PRIORITY4"
                ),
                # 右半图中心
                (
                    right_half_start + (width - right_half_start - width_needed) // 2,
                    (height - height_needed) // 2,
                    "right-half-center-PRIORITY4"
                ),
            ])
            
            # === 第1优先级：印章左侧的空白区域（红色框区域）===
            # 这是用户指出的左边空白位置，应该最优先考虑
            # 计算印章左侧的空白区域：从图片左边界到印章左边界之间
            # 特别注意：这个区域也包括"日期左侧的空白区域"
            
            # 计算日期可能的Y坐标位置
            # 日期通常在印章上方，距离印章上边界约30-40px
            estimated_date_y = s_y1 - 30
            
            # 如果日期已经被放置，优先尝试放在日期左侧
            if date_text_bbox:
                d_x1, d_y1, d_x2, d_y2 = date_text_bbox
                date_center_y = (d_y1 + d_y2) // 2
                
                # 印章机关文字放在日期左侧，垂直居中对齐
                # 计算合适的X坐标：日期左边界 - 间距 - 文本宽度
                seal_text_x = d_x1 - 20 - width_needed
                seal_text_y = date_center_y - height_needed // 2
                
                print(f"[COLLISION_DETECTION] Date already placed at {date_text_bbox}")
                print(f"[COLLISION_DETECTION] Trying to place seal text to the left of date")
                print(f"[COLLISION_DETECTION] seal_text position: ({seal_text_x}, {seal_text_y})")
                
                candidate_positions.extend([
                    # === 第1优先级：日期左侧（最高优先级）===
                    (
                        seal_text_x,
                        seal_text_y,
                        "left-of-date-center-aligned-PRIORITY5"
                    ),
                    # 如果左侧空间不够，尝试更靠左的位置
                    (
                        seal_text_x - 50,
                        seal_text_y,
                        "left-of-date-far-left-PRIORITY5"
                    ),
                    (
                        seal_text_x - 100,
                        seal_text_y,
                        "left-of-date-very-far-left-PRIORITY5"
                    ),
                    # 尝试不同的Y坐标（与日期顶部对齐）
                    (
                        seal_text_x,
                        d_y1,
                        "left-of-date-top-aligned-PRIORITY5"
                    ),
                    # 尝试不同的Y坐标（与日期底部对齐）
                    (
                        seal_text_x,
                        d_y2 - height_needed,
                        "left-of-date-bottom-aligned-PRIORITY5"
                    ),
                ])
            
            # === 备选：其他位置（PRIORITY5）===
            candidate_positions.extend([
                (
                    100,
                    300,
                    "fallback-top-left-PRIORITY5"
                ),
                (
                    300,
                    300,
                    "fallback-top-center-PRIORITY5"
                ),
                (
                    500,
                    300,
                    "fallback-top-right-PRIORITY5"
                ),
            ])
        else:
            # 日期文字：根据用户需求，优先放在"Registration Authority"下方
            # 如果没有提供registration_authority_bbox，则回退到印章下方
            candidate_positions = []
            
            # === 第1优先级：Registration Authority下方（PRIORITY0）===
            if registration_authority_bbox:
                ra_x1, ra_y1, ra_x2, ra_y2 = registration_authority_bbox
                ra_center_x = (ra_x1 + ra_x2) // 2
                
                # 计算在Registration Authority下方的位置
                safe_gap = 20  # Registration Authority下方的间距
                below_ra_y = ra_y2 + safe_gap
                
                if below_ra_y + height_needed <= height - 20:  # 确保不会超出图片底部
                    logger.info(f"🎯 最高优先级：Registration Authority下方 (y > {ra_y2})")
                    
                    candidate_positions.extend([
                        # 与Registration Authority居中对齐（最优）
                        (
                            ra_center_x - width_needed // 2,
                            below_ra_y,
                            "below-registration-authority-centered-PRIORITY0"
                        ),
                        # 与Registration Authority左对齐
                        (
                            ra_x1,
                            below_ra_y,
                            "below-registration-authority-left-aligned-PRIORITY0"
                        ),
                        # 与Registration Authority右对齐
                        (
                            ra_x2 - width_needed,
                            below_ra_y,
                            "below-registration-authority-right-aligned-PRIORITY0"
                        ),
                    ])
                    
                    # 在Registration Authority下方，尝试多个x位置
                    for x_offset in [-50, -30, 0, 30, 50]:
                        x_pos = ra_center_x - width_needed // 2 + x_offset
                        if x_pos >= 20 and x_pos + width_needed <= width - 20:
                            candidate_positions.append((
                                x_pos,
                                below_ra_y,
                                f"below-registration-authority-offset-{x_offset}-PRIORITY0"
                            ))
            
            # === 第2优先级：印章下方居中（PRIORITY1）===
            # 如果没有Registration Authority信息，或者下方位置不可用，则使用印章下方
            
            seal_bottom_y = s_y2
            below_seal_y = seal_bottom_y + self.min_distance
            
            if below_seal_y + height_needed <= height - 20:  # 确保不会超出图片底部
                logger.info(f"🔍 第2优先级：印章下方居中 (y > {seal_bottom_y})")
                
                # 最高优先级：与印章居中对齐
                candidate_positions.extend([
                    # 与印章居中对齐
                    (
                        seal_center_x - width_needed // 2,
                        below_seal_y,
                        "below-seal-centered-PRIORITY1"
                    ),
                    # 稍微偏右
                    (
                        seal_center_x - width_needed // 2 + 20,
                        below_seal_y,
                        "below-seal-right-offset-PRIORITY1"
                    ),
                    # 稍微偏左
                    (
                        seal_center_x - width_needed // 2 - 20,
                        below_seal_y,
                        "below-seal-left-offset-PRIORITY1"
                    ),
                    # 与印章左对齐
                    (
                        s_x1,
                        below_seal_y,
                        "below-seal-left-aligned-PRIORITY1"
                    ),
                    # 与印章右对齐
                    (
                        s_x2 - width_needed,
                        below_seal_y,
                        "below-seal-right-aligned-PRIORITY1"
                    ),
                ])
                
                # 在印章下方，从右到左尝试多个x位置
                for x_ratio in [0.9, 0.8, 0.7, 0.6, 0.5, 0.4, 0.3]:
                    x_pos = int(width * x_ratio) - width_needed // 2
                    if x_pos >= 20 and x_pos + width_needed <= width - 20:
                        candidate_positions.append((
                            x_pos,
                            below_seal_y,
                            f"below-seal-x-ratio-{int(x_ratio*100)}-PRIORITY2"
                        ))
            
            # === 第2优先级：印章上方（新理解：y < 印章y1的任何位置）===
            # 不限制x坐标，只要y坐标在印章上方即可
            
            seal_top_y = s_y1
            above_seal_y = seal_top_y - self.min_distance - height_needed
            
            if above_seal_y >= 20:  # 确保不会超出图片顶部
                logger.info(f"🔍 日期尝试印章上方区域 (y < {seal_top_y})")
                
                # 在印章上方，从右到左尝试多个x位置
                for x_ratio in [0.9, 0.8, 0.7, 0.6, 0.5, 0.4, 0.3]:
                    x_pos = int(width * x_ratio) - width_needed // 2
                    if x_pos >= 20 and x_pos + width_needed <= width - 20:
                        candidate_positions.append((
                            x_pos,
                            above_seal_y,
                            f"above-seal-x-ratio-{int(x_ratio*100)}-PRIORITY4"
                        ))
                
                # 也尝试与印章对齐的位置
                candidate_positions.extend([
                    # 与印章居中对齐
                    (
                        seal_center_x - width_needed // 2,
                        above_seal_y,
                        "above-seal-centered-PRIORITY4"
                    ),
                    # 与印章左对齐
                    (
                        s_x1,
                        above_seal_y,
                        "above-seal-left-aligned-PRIORITY4"
                    ),
                    # 与印章右对齐
                    (
                        s_x2 - width_needed,
                        above_seal_y,
                        "above-seal-right-aligned-PRIORITY4"
                    ),
                ])
            
            # === 第3优先级：印章左右两侧 ===
            candidate_positions.extend([
                # 印章右侧 - 垂直居中
                (
                    s_x2 + self.min_distance,
                    seal_center_y - height_needed // 2,
                    "seal-right-centered-PRIORITY5"
                ),
                # 印章左侧 - 垂直居中
                (
                    s_x1 - self.min_distance - width_needed,
                    seal_center_y - height_needed // 2,
                    "seal-left-centered-PRIORITY5"
                ),
                # 印章右侧 - 靠上
                (
                    s_x2 + self.min_distance,
                    s_y1,
                    "seal-right-top-PRIORITY5"
                ),
                # 印章左侧 - 靠上
                (
                    s_x1 - self.min_distance - width_needed,
                    s_y1,
                    "seal-left-top-PRIORITY5"
                ),
                # 印章右侧 - 靠下
                (
                    s_x2 + self.min_distance,
                    s_y2 - height_needed,
                    "seal-right-bottom-PRIORITY5"
                ),
                # 印章左侧 - 靠下
                (
                    s_x1 - self.min_distance - width_needed,
                    s_y2 - height_needed,
                    "seal-left-bottom-PRIORITY5"
                ),
            ])
        
        # 遍历所有候选位置，找到第一个不与已占用区域碰撞的位置
        print(f"\n[COLLISION_DETECTION] Trying {len(candidate_positions)} candidate positions")
        for idx, (x1, y1, position_desc) in enumerate(candidate_positions):
            try:
                print(f"  [COLLISION_DETECTION] === Candidate #{idx+1}/{len(candidate_positions)} ===")
                x2 = x1 + width_needed
                y2 = y1 + height_needed
                
                print(f"  [COLLISION_DETECTION] Trying position '{position_desc}': ({x1}, {y1}, {x2}, {y2})")
                logger.info(f"  🔍 尝试位置 '{position_desc}': ({x1}, {y1}, {x2}, {y2})")
            except Exception as e:
                print(f"  [COLLISION_DETECTION] X Exception: {e}")
                import traceback
                traceback.print_exc()
                continue
            
            # 检查是否在图片边界内
            if x1 < 0 or y1 < 0 or x2 > width or y2 > height:
                print(f"  [COLLISION_DETECTION] X Out of bounds (image: {width}x{height})")
                logger.info(f"  ❌ 位置 {position_desc} 超出图片边界 (image: {width}x{height})")
                continue
            
            candidate_bbox = (x1, y1, x2, y2)
            
            # 检查是否与印章重叠
            # 对于边框位置(右下角边框或印章下方边框)，允许与印章有一定重叠
            is_border_box_position = (
                position_desc.startswith("right-bottom-box") or 
                position_desc.startswith("seal-bottom-border")
            )
            
            if not is_border_box_position:
                # 非边框位置，严格检查与印章的重叠
                if self._has_any_overlap(candidate_bbox, seal_bbox):
                    print(f"  [COLLISION_DETECTION] X Overlaps with seal")
                    logger.info(f"  ❌ 位置 {position_desc} 与印章重叠")
                    continue
            else:
                # 边框位置，允许与印章有轻微重叠
                # 只有当候选位置的中心点在印章内部时才认为重叠
                cand_center_x = (x1 + x2) // 2
                cand_center_y = (y1 + y2) // 2
                s_x1, s_y1, s_x2, s_y2 = seal_bbox
                
                if s_x1 <= cand_center_x <= s_x2 and s_y1 <= cand_center_y <= s_y2:
                    logger.info(f"  ❌ 位置 {position_desc} 中心点在印章内部")
                    continue
                else:
                    logger.info(f"  ✅ 位置 {position_desc} 允许与印章边缘重叠（中心点不在印章内）")
            
            # 检查是否与已占用区域重叠
            has_collision = False
            
            # 对于右下角边框区域的候选位置，使用宽松的碰撞检测
            # 因为这个区域可能包含一些OCR碎片（如"年月日"），但实际上是空白的
            
            for occupied_bbox in occupied_regions:
                # 如果是右下角边框位置，只检查与大文本区域的碰撞
                # 忽略小文本片段（高度<30px或宽度<80px）
                if is_border_box_position:
                    o_x1, o_y1, o_x2, o_y2 = occupied_bbox
                    occupied_height = o_y2 - o_y1
                    occupied_width = o_x2 - o_x1
                    
                    # 如果是小文本片段，跳过碰撞检测
                    if occupied_height < 30 or occupied_width < 80:
                        logger.info(
                            f"    ⏭️  跳过小文本片段的碰撞检测: occupied={occupied_bbox} "
                            f"(h={occupied_height}, w={occupied_width})"
                        )
                        continue  # 跳过这个occupied_bbox，继续检查下一个
                
                if self._has_any_overlap(candidate_bbox, occupied_bbox):
                    has_collision = True
                    print(
                        f"  [COLLISION_DETECTION] X Overlaps with occupied: "
                        f"candidate={candidate_bbox}, occupied={occupied_bbox}"
                    )
                    logger.info(
                        f"  ❌ 位置 {position_desc} 与已占用区域重叠: "
                        f"candidate={candidate_bbox}, occupied={occupied_bbox}"
                    )
                    break  # 发现碰撞，停止检查
            
            print(f"  [COLLISION_DETECTION] Collision result: has_collision={has_collision}, checked {len(occupied_regions)} regions")
            logger.info(f"  🔍 位置 {position_desc} 碰撞检测结果: has_collision={has_collision}, checked {len(occupied_regions)} regions")
            
            if not has_collision:
                print(
                    f"  [COLLISION_DETECTION] OK Found safe position '{position_desc}': {candidate_bbox}"
                )
                logger.info(
                    f"  ✅ 找到安全位置 {position_desc}: {candidate_bbox}"
                )
                return candidate_bbox
        
        # 如果所有候选位置都不可用，尝试扩展搜索
        logger.info("🔍 候选位置都不可用，尝试扩展搜索...")
        
        for radius in range(self.min_distance, self.search_radius, 20):
            # 尝试8个方向
            angles = [0, 45, 90, 135, 180, 225, 270, 315]
            
            for angle in angles:
                rad = np.radians(angle)
                x1 = int(seal_center_x + radius * np.cos(rad) - width_needed // 2)
                y1 = int(seal_center_y + radius * np.sin(rad) - height_needed // 2)
                x2 = x1 + width_needed
                y2 = y1 + height_needed
                
                # 检查边界
                if x1 < 0 or y1 < 0 or x2 > width or y2 > height:
                    continue
                
                candidate_bbox = (x1, y1, x2, y2)
                
                # 检查碰撞
                if self._has_any_overlap(candidate_bbox, seal_bbox):
                    continue
                
                has_collision = False
                for occupied_bbox in occupied_regions:
                    if self._has_any_overlap(candidate_bbox, occupied_bbox):
                        has_collision = True
                        break
                
                if not has_collision:
                    logger.info(
                        f"  ✅ 扩展搜索找到位置: radius={radius}, angle={angle}°, "
                        f"bbox={candidate_bbox}"
                    )
                    return candidate_bbox
        
        logger.warning(f"  ❌ 无法找到安全位置（搜索半径={self.search_radius}）")
        return None
    
    def _check_overlap(
        self,
        bbox1: Tuple[int, int, int, int],
        bbox2: Tuple[int, int, int, int]
    ) -> bool:
        """Check if two bounding boxes overlap with strict criteria.
        
        使用严格的重叠判断：x轴和y轴都必须有显著重叠才算真正重叠。
        这可以过滤掉只是边缘轻微接触的情况（如二维码说明文字）。
        
        Args:
            bbox1: First bounding box (x1, y1, x2, y2)
            bbox2: Second bounding box (x1, y1, x2, y2)
            
        Returns:
            True if boxes overlap significantly on both x and y axes
        """
        x1_1, y1_1, x2_1, y2_1 = bbox1
        x1_2, y1_2, x2_2, y2_2 = bbox2
        
        # Calculate dimensions
        width1 = x2_1 - x1_1
        height1 = y2_1 - y1_1
        width2 = x2_2 - x1_2
        height2 = y2_2 - y1_2
        
        # Calculate overlap
        overlap_x1 = max(x1_1, x1_2)
        overlap_y1 = max(y1_1, y1_2)
        overlap_x2 = min(x2_1, x2_2)
        overlap_y2 = min(y2_1, y2_2)
        
        # Check if there's ANY overlap
        if overlap_x1 >= overlap_x2 or overlap_y1 >= overlap_y2:
            return False
        
        # Calculate overlap dimensions
        overlap_width = overlap_x2 - overlap_x1
        overlap_height = overlap_y2 - overlap_y1
        
        # Calculate overlap ratios for both boxes
        x_overlap_ratio1 = overlap_width / width1 if width1 > 0 else 0
        y_overlap_ratio1 = overlap_height / height1 if height1 > 0 else 0
        x_overlap_ratio2 = overlap_width / width2 if width2 > 0 else 0
        y_overlap_ratio2 = overlap_height / height2 if height2 > 0 else 0
        
        # Use the maximum overlap ratio (more lenient)
        x_overlap_ratio = max(x_overlap_ratio1, x_overlap_ratio2)
        y_overlap_ratio = max(y_overlap_ratio1, y_overlap_ratio2)
        
        # Require significant overlap on both axes
        # x轴阈值50%：文字必须有一半以上在印章范围内
        # y轴阈值20%：允许文字在印章边缘轻微接触
        min_x_overlap_ratio = 0.5
        min_y_overlap_ratio = 0.2
        
        if x_overlap_ratio < min_x_overlap_ratio or y_overlap_ratio < min_y_overlap_ratio:
            return False
        
        return True
    
    def _estimate_font_size(self, bbox: Tuple[int, int, int, int]) -> int:
        """Estimate font size from bounding box height.
        
        Args:
            bbox: Bounding box (x1, y1, x2, y2)
            
        Returns:
            Estimated font size in pixels
        """
        _, y1, _, y2 = bbox
        height = abs(y2 - y1)
        # Font size is approximately the height of the bounding box
        return max(1, height)
    
    def _separate_seal_org_and_date(
        self,
        text: str,
        bbox: Tuple[int, int, int, int],
        confidence: float,
        seal_bbox: Tuple[int, int, int, int]
    ) -> Optional[List[TextRegion]]:
        """分离印章机关名称和日期部分。
        
        如果文本同时包含印章机关名称和日期信息，将其分离成两个独立的区域。
        
        Args:
            text: 文本内容
            bbox: 边界框
            confidence: 置信度
            seal_bbox: 印章边界框
            
        Returns:
            分离后的区域列表，如果无需分离则返回None
        """
        import re
        
        # 按行分析
        lines = text.split('\n')
        
        seal_org_lines = []
        date_lines = []
        
        for line in lines:
            line_stripped = line.strip()
            if not line_stripped:
                continue
            
            # 检查是否只包含数字、空格和"年月日"
            if re.match(r'^[\d\s年月日]+$', line_stripped):
                date_lines.append(line_stripped)
            else:
                seal_org_lines.append(line_stripped)
        
        # 如果同时有印章机关和日期，则分离
        if seal_org_lines and date_lines:
            logger.info(
                f"🔍 检测到混合内容，分离印章机关和日期："
            )
            logger.info(f"  印章机关: {seal_org_lines}")
            logger.info(f"  日期部分: {date_lines}")
            
            separated_regions = []
            
            # 创建印章机关区域
            seal_org_text = ''.join(seal_org_lines)
            seal_org_region = TextRegion(
                bbox=bbox,  # 使用原始bbox，后续会重新分类
                text=seal_org_text,
                confidence=confidence,
                font_size=self._estimate_font_size(bbox),
                angle=0.0
            )
            seal_org_region.overlapping_seal = seal_bbox
            seal_org_region.is_seal_org = True  # 标记为印章机关
            separated_regions.append(seal_org_region)
            
            # 创建日期区域
            date_text = ''.join(date_lines)
            # 智能处理空格和格式
            # 输入: "12 16" + "年 月 日"
            # 期望输出: "年12月16日"（年字在前面，这样和"2022"合并后就是"2022年12月16日"）
            
            # 先移除所有空格
            date_text_no_space = re.sub(r'\s+', '', date_text)
            
            # 尝试匹配: "1216年月日" -> "年12月16日"
            match = re.match(r'(\d{1,2})(\d{1,2})年月日', date_text_no_space)
            if match:
                month = match.group(1)
                day = match.group(2)
                date_text = f"年{month}月{day}日"
                logger.info(f"  格式化日期: '{date_text_no_space}' -> '{date_text}'")
            else:
                # 如果没有匹配，保持原样但移除空格
                date_text = date_text_no_space
            
            date_region = TextRegion(
                bbox=bbox,  # 使用原始bbox，后续会重新分类
                text=date_text,
                confidence=confidence,
                font_size=self._estimate_font_size(bbox),
                angle=0.0
            )
            date_region.overlapping_seal = seal_bbox
            date_region.is_date_part = True  # 标记为日期部分
            separated_regions.append(date_region)
            
            logger.info(
                f"✅ 分离完成: 印章机关='{seal_org_text}', 日期='{date_text}'"
            )
            
            return separated_regions
        
        # 无需分离
        return None

    
    def _merge_date_fragments(
        self,
        seal_text_regions: List[TextRegion],
        seals: List[Tuple[int, int, int, int]]
    ) -> List[TextRegion]:
        """Merge date text fragments that are split by seals.
        
        When a seal overlaps with a date (e.g., "2022年12月16日"), the OCR
        may detect it as multiple fragments. This method merges them back
        into a complete date.
        
        改进策略：
        1. 检测是否有完整的日期文本（如"2020年11月09日"）
        2. 如果有完整日期，过滤掉可能是重复的日期片段（如单独的"2020"和"11月09"）
        3. 只合并真正需要合并的片段（如"年12月16日"需要和"2022"合并）
        
        Args:
            seal_text_regions: Text regions overlapping with seals
            seals: List of seal bounding boxes
            
        Returns:
            List of text regions with merged date fragments
        """
        import re
        
        if not seal_text_regions:
            return seal_text_regions
        
        # Group regions by their overlapping seal
        seal_groups = {}
        for region in seal_text_regions:
            seal_bbox = getattr(region, 'overlapping_seal', None)
            if seal_bbox:
                seal_key = tuple(seal_bbox)
                if seal_key not in seal_groups:
                    seal_groups[seal_key] = []
                seal_groups[seal_key].append(region)
        
        merged_regions = []
        
        # Process each seal group
        for seal_bbox, group_regions in seal_groups.items():
            # Find date fragments in this group
            date_fragments = []
            other_regions = []
            complete_dates = []  # 完整的日期（如"2020年11月09日"）
            
            for region in group_regions:
                text = region.text.strip()
                
                # 检查是否是完整的日期（包含年月日）
                # 完整日期格式：2020年11月09日、2020-11-09、2020/11/09等
                is_complete_date = bool(re.search(
                    r'\d{4}年\d{1,2}月\d{1,2}日|\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2}',
                    text
                ))
                
                if is_complete_date:
                    # 这是完整的日期，单独保留
                    complete_dates.append(region)
                    logger.info(f"✅ 识别为完整日期: '{text[:50]}...'")
                    continue
                
                # 检查是否是纯日期文本片段（不包含其他内容如印章机关名称）
                # 纯日期片段特征：
                # 1. 只包含数字和"年月日"关键字，没有其他汉字
                # 2. 或者是纯4位数字（年份）
                # 3. 或者是"月日"格式（如"11月09"）
                
                # 移除空白字符后检查
                text_no_space = re.sub(r'\s+', '', text)
                
                # 检查是否只包含数字和年月日
                is_pure_date = bool(re.match(r'^[\d年月日\s]+$', text_no_space))
                # 或者是纯4位数字
                is_year_only = bool(re.match(r'^\d{4}$', text_no_space))
                
                if is_pure_date or is_year_only:
                    # 纯日期片段
                    date_fragments.append(region)
                    logger.debug(f"识别为日期片段: '{text[:30]}...'")
                else:
                    # 包含其他内容（如印章机关名称），不是纯日期
                    other_regions.append(region)
                    logger.debug(f"不是日期片段: '{text[:30]}...'")
            
            # 如果已经有完整日期，不再合并日期片段（避免重复）
            if complete_dates:
                logger.info(
                    f"✅ 已有 {len(complete_dates)} 个完整日期，跳过 {len(date_fragments)} 个日期片段的合并"
                )
                # 保留完整日期和其他区域，丢弃日期片段
                merged_regions.extend(complete_dates)
                merged_regions.extend(other_regions)
                continue
            # If we have multiple date fragments, try to merge them
            if len(date_fragments) > 1:
                # 智能排序：优先按y坐标（上下），然后按x坐标（左右）
                # 但是对于日期，如果有纯数字（年份）和包含"年月日"的片段，
                # 应该把纯数字放在前面
                
                # 先按位置排序
                date_fragments.sort(key=lambda r: (r.bbox[1], r.bbox[0]))
                
                # 检查是否有年份片段（纯4位数字）和日期片段（包含"年月日"）
                year_fragments = []
                date_keyword_fragments = []
                other_fragments = []
                
                for frag in date_fragments:
                    text = frag.text.strip()
                    if re.match(r'^\d{4}$', text):
                        # 纯4位数字，可能是年份
                        year_fragments.append(frag)
                    elif re.search(r'年|月|日', text):
                        # 包含日期关键字
                        date_keyword_fragments.append(frag)
                    else:
                        other_fragments.append(frag)
                
                # 重新排序：年份 -> 日期关键字 -> 其他
                sorted_fragments = year_fragments + date_keyword_fragments + other_fragments
                
                # 合并日期片段
                combined_text = ''.join(r.text.strip() for r in sorted_fragments)
                
                logger.info(f"🔍 尝试合并 {len(date_fragments)} 个日期片段:")
                for i, frag in enumerate(sorted_fragments):
                    logger.info(f"  片段 {i+1}: '{frag.text[:50]}...' at {frag.bbox}")
                logger.info(f"  合并后文本: '{combined_text[:100]}...'")
                
                # 检查合并后的文本是否包含完整日期信息
                # 1. 检查是否有年份（4位数字）
                has_year = bool(re.search(r'\d{4}', combined_text))
                # 2. 检查是否有"年月日"关键字或者"月日"关键字
                has_date_keywords = bool(re.search(r'年.*月.*日|月.*日', combined_text))
                # 3. 检查是否有月份和日期数字
                has_month_day = bool(re.search(r'\d{1,2}.*\d{1,2}', combined_text))
                
                # 如果满足以下任一条件，则认为是完整日期：
                # - 有年份 + 有"月日"关键字（如"2022年" + "12月16日" = "2022年12月16日"）
                # - 标准日期格式：2022年12月16日
                is_complete_date = (has_year and has_date_keywords) or bool(re.search(
                    r'\d{4}年\d{1,2}月\d{1,2}日|\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2}',
                    combined_text
                ))
                
                logger.info(f"  日期检查: has_year={has_year}, has_date_keywords={has_date_keywords}, has_month_day={has_month_day}, is_complete_date={is_complete_date}")
                
                if is_complete_date:
                    # Merge into a single region
                    # Calculate bounding box that encompasses all fragments
                    x1 = min(r.bbox[0] for r in sorted_fragments)
                    y1 = min(r.bbox[1] for r in sorted_fragments)
                    x2 = max(r.bbox[2] for r in sorted_fragments)
                    y2 = max(r.bbox[3] for r in sorted_fragments)
                    
                    # Create merged region
                    merged_region = TextRegion(
                        bbox=(x1, y1, x2, y2),
                        text=combined_text,
                        confidence=sum(r.confidence for r in sorted_fragments) / len(sorted_fragments),
                        font_size=sorted_fragments[0].font_size,
                        angle=0.0
                    )
                    merged_region.overlapping_seal = seal_bbox
                    merged_region.is_merged_date = True
                    
                    merged_regions.append(merged_region)
                    
                    logger.info(
                        f"✅ 合并了 {len(date_fragments)} 个日期片段: '{combined_text[:100]}...' "
                        f"at {merged_region.bbox}"
                    )
                else:
                    # Not a complete date, keep fragments separate
                    logger.info(f"⚠️  不是完整日期，保持片段分离")
                    merged_regions.extend(date_fragments)
            else:
                # No merging needed
                merged_regions.extend(date_fragments)
            
            # Add other regions
            merged_regions.extend(other_regions)
        
        # Add regions that don't have an overlapping seal
        for region in seal_text_regions:
            if not hasattr(region, 'overlapping_seal') or region.overlapping_seal is None:
                merged_regions.append(region)
        
        return merged_regions
    
    def _classify_seal_text_regions(
        self,
        seal_text_regions: List[TextRegion],
        seals: List[Tuple[int, int, int, int]]
    ) -> List[TextRegion]:
        """Classify seal text regions as seal_inner or seal_overlap.
        
        - seal_inner: Text completely inside the seal (e.g., seal text)
        - seal_overlap: Text overlapping with seal (e.g., date covered by seal)
        
        Strategy:
        1. Check if text contains date patterns -> seal_overlap
        2. Check overlap ratio: >80% -> seal_inner, else -> seal_overlap
        3. Check text center position: inside seal -> seal_inner
        4. Filter out text that is just below the seal (likely year number, not seal text)
        
        Args:
            seal_text_regions: Text regions overlapping with seals
            seals: List of seal bounding boxes
            
        Returns:
            List of text regions with classification
        """
        import re
        
        filtered_regions = []
        
        for region in seal_text_regions:
            seal_bbox = getattr(region, 'overlapping_seal', None)
            if not seal_bbox:
                filtered_regions.append(region)
                continue
            
            s_x1, s_y1, s_x2, s_y2 = seal_bbox
            r_x1, r_y1, r_x2, r_y2 = region.bbox
            
            # 注释掉过滤印章下方年份的逻辑
            # 因为这些年份可能是日期的一部分，应该被识别为印章重叠文字
            # Check if text is just below the seal (likely a year number, not seal text)
            # If text is completely below the seal and is a short number, skip it
            # text = region.text.strip()
            # if r_y1 > s_y2:  # Text is below seal
            #     # Check if it's a short number (like "2022")
            #     if re.match(r'^\d{4}$', text):
            #         logger.info(
            #             f"Skipping year number '{text}' below seal - not seal text"
            #         )
            #         continue
            
            text = region.text.strip()
            
            # Check if text contains date patterns
            has_date = bool(re.search(
                r'\d{4}年\d{1,2}月\d{1,2}日|\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2}|\d{4}年|\d{1,2}月|\d{1,2}日',
                text
            ))
            
            if has_date:
                # If text contains date, it's likely seal_overlap
                region.seal_text_type = 'seal_overlap'
                logger.info(
                    f"Classified '{region.text[:30]}...' as seal_overlap "
                    f"(contains date pattern)"
                )
                filtered_regions.append(region)
                continue
            
            # Calculate overlap ratio
            overlap_x1 = max(r_x1, s_x1)
            overlap_y1 = max(r_y1, s_y1)
            overlap_x2 = min(r_x2, s_x2)
            overlap_y2 = min(r_y2, s_y2)
            
            if overlap_x1 < overlap_x2 and overlap_y1 < overlap_y2:
                overlap_area = (overlap_x2 - overlap_x1) * (overlap_y2 - overlap_y1)
                region_area = (r_x2 - r_x1) * (r_y2 - r_y1)
                overlap_ratio = overlap_area / region_area if region_area > 0 else 0
                
                # Check if text center is inside seal
                text_center_x = (r_x1 + r_x2) / 2
                text_center_y = (r_y1 + r_y2) / 2
                center_in_seal = (s_x1 <= text_center_x <= s_x2 and 
                                s_y1 <= text_center_y <= s_y2)
                
                # If more than 60% inside seal OR center is inside seal, it's seal inner text
                if overlap_ratio > 0.6 or center_in_seal:
                    region.seal_text_type = 'seal_inner'
                    logger.info(
                        f"Classified '{region.text[:30]}...' as seal_inner "
                        f"(overlap={overlap_ratio:.1%}, center_in_seal={center_in_seal})"
                    )
                else:
                    region.seal_text_type = 'seal_overlap'
                    logger.info(
                        f"Classified '{region.text[:30]}...' as seal_overlap "
                        f"(overlap={overlap_ratio:.1%}, center_in_seal={center_in_seal})"
                    )
            else:
                # No overlap, shouldn't happen but handle it
                region.seal_text_type = 'seal_overlap'
            
            filtered_regions.append(region)
        
        return filtered_regions
        

    def detect_seal_frames(self, image: np.ndarray) -> List[Tuple[int, int, int, int]]:
        """检测红色印章框区域（矩形框，非圆形印章）
        
        改进策略：
        1. 降低面积阈值，避免过滤掉较小的框
        2. 放宽宽高比限制，支持更多矩形形状
        3. 使用边缘检测辅助识别矩形框
        
        Args:
            image: 输入图片（BGR格式）
            
        Returns:
            框的边界坐标列表 [(x1, y1, x2, y2), ...]
        """
        if not self.frame_rendering_enabled:
            return []
        
        height, width = image.shape[:2]
        hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)
        
        # 检测红色区域（使用更宽松的HSV范围）
        # 降低饱和度和亮度阈值，以检测更多红色区域
        red_sat_min = self.config.get('rendering.seal_detection.red_saturation_min', 30)  # 从40降到30
        red_val_min = self.config.get('rendering.seal_detection.red_value_min', 30)  # 从40降到30
        
        lower_red1 = np.array([0, red_sat_min, red_val_min])
        upper_red1 = np.array([10, 255, 255])
        mask_red1 = cv2.inRange(hsv, lower_red1, upper_red1)
        
        lower_red2 = np.array([170, red_sat_min, red_val_min])
        upper_red2 = np.array([180, 255, 255])
        mask_red2 = cv2.inRange(hsv, lower_red2, upper_red2)
        
        mask_red = cv2.bitwise_or(mask_red1, mask_red2)
        
        # 形态学操作：闭运算连接断裂的边框
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (5, 5))
        mask_red = cv2.morphologyEx(mask_red, cv2.MORPH_CLOSE, kernel)
        
        # 找到轮廓
        contours, _ = cv2.findContours(mask_red, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        logger.info(f"🔍 检测到 {len(contours)} 个红色轮廓")
        
        if not contours:
            logger.debug("未检测到红色区域")
            return []
        
        frames = []
        # 使用配置的宽高比范围，如果没有则使用更宽松的默认值
        min_aspect, max_aspect = self.frame_aspect_ratio_range
        # 降低最小面积阈值，从10000降到5000
        min_area = max(5000, self.min_frame_area // 2)
        
        for idx, contour in enumerate(contours):
            contour_area = cv2.contourArea(contour)
            
            logger.debug(f"  轮廓#{idx}: area={contour_area:.0f}")
            
            # 过滤太小的轮廓（使用更宽松的阈值）
            if contour_area < min_area:
                logger.debug(f"    ✗ 面积太小 (< {min_area})")
                continue
            
            # 获取边界框
            x, y, w, h = cv2.boundingRect(contour)
            
            # 检查宽高比：矩形框的宽高比应该接近1（正方形）或略宽
            aspect_ratio = w / h if h > 0 else 0
            logger.debug(f"    宽高比: {aspect_ratio:.2f} (范围: {min_aspect}-{max_aspect})")
            
            if not (min_aspect <= aspect_ratio <= max_aspect):
                logger.debug(f"    ✗ 宽高比不符合")
                continue
            
            # 计算圆度：矩形框的圆度应该较低（与圆形印章区分）
            perimeter = cv2.arcLength(contour, True)
            if perimeter > 0:
                circularity = 4 * np.pi * contour_area / (perimeter * perimeter)
                logger.debug(f"    圆度: {circularity:.2f}")
                
                # 矩形框的圆度通常 < 0.8，圆形印章的圆度 > 0.8
                if circularity > 0.8:
                    logger.debug(
                        f"    ✗ 跳过圆形区域 (circularity={circularity:.2f})"
                    )
                    continue
            else:
                circularity = 0
            
            # 位置筛选：优先右下角区域（营业执照的印章框通常在右下角）
            # 但不强制要求，因为框可能在其他位置
            center_x = x + w // 2
            center_y = y + h // 2
            is_right_bottom = (center_x > width * 0.5 and center_y > height * 0.5)
            
            frame_bbox = (x, y, x + w, y + h)
            frames.append(frame_bbox)
            
            logger.info(
                f"✅ 检测到红色框#{idx}: bbox={frame_bbox}, "
                f"area={contour_area:.0f}, aspect_ratio={aspect_ratio:.2f}, "
                f"circularity={circularity:.2f}, right_bottom={is_right_bottom}"
            )
        
        logger.info(f"共检测到 {len(frames)} 个红色框")
        return frames

    def _is_english_text(self, text: str) -> bool:
        """检查文本是否主要是英文。
        
        Args:
            text: 文本内容
            
        Returns:
            True if text is primarily English, False otherwise
        """
        if not text:
            return False
        
        text = text.strip()
        
        # 统计英文字母和中文字符的数量
        import re
        english_chars = len(re.findall(r'[a-zA-Z]', text))
        chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', text))
        
        # 如果英文字母数量 > 中文字符数量，认为是英文文本
        # 并且英文字母数量至少要有3个（避免误判单个字母）
        if english_chars >= 3 and english_chars > chinese_chars:
            return True
        
        return False
    
    def _is_english_date(self, text: str) -> bool:
        """检查文本是否是英文日期格式。
        
        Args:
            text: 文本内容
            
        Returns:
            True if text is an English date, False otherwise
        """
        if not text:
            return False
        
        text = text.strip()
        
        import re
        
        # 英文月份名称
        english_months = [
            'January', 'February', 'March', 'April', 'May', 'June',
            'July', 'August', 'September', 'October', 'November', 'December',
            'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
            'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'
        ]
        
        # 检查是否包含英文月份
        for month in english_months:
            if month in text:
                return True
        
        # 检查常见的英文日期格式
        # 例如: "April 17, 2018", "17 April 2018", "2018-04-17"
        date_patterns = [
            r'[A-Z][a-z]+\s+\d{1,2},?\s+\d{4}',  # April 17, 2018
            r'\d{1,2}\s+[A-Z][a-z]+\s+\d{4}',    # 17 April 2018
            r'\d{4}-\d{2}-\d{2}',                 # 2018-04-17
            r'\d{2}/\d{2}/\d{4}',                 # 04/17/2018
            r'\d{4}/\d{2}/\d{2}',                 # 2018/04/17
        ]
        
        for pattern in date_patterns:
            if re.search(pattern, text):
                return True
        
        return False
    
    def verify_seal_text_with_user(
        self,
        region: TextRegion,
        text_type: str,
        image: Optional[np.ndarray] = None
    ) -> Tuple[bool, Optional[str]]:
        """与用户交互验证印章文字识别结果。
        
        Args:
            region: 识别到的文本区域
            text_type: 文本类型（'seal_inner' 表示印章内文字，'seal_overlap' 表示被覆盖的日期）
            image: 原始图片（用于显示预览，可选）
            
        Returns:
            (是否继续处理, 用户修正的文本或None)
            - (True, None): 确认识别正确，使用原始识别结果
            - (True, corrected_text): 用户提供了修正内容，使用修正后的文本
            - (False, None): 用户选择跳过此区域
        """
        # 检查是否启用交互式验证
        if not self.config.get('rendering.seal_text_handling.interactive_verification.enabled', False):
            return True, None
        
        # 检查置信度是否足够高，可以自动确认
        auto_threshold = self.config.get(
            'rendering.seal_text_handling.interactive_verification.auto_confirm_threshold',
            0.95
        )
        if region.confidence >= auto_threshold:
            logger.info(f"✅ 置信度 {region.confidence:.2f} >= {auto_threshold}，自动确认")
            return True, None
        
        # 显示识别结果
        type_name = "印章内文字" if text_type == 'seal_inner' else "被覆盖的日期"
        print("\n" + "=" * 60)
        print(f"检测到{type_name}：")
        print("-" * 60)
        print(f"类型: {text_type}")
        print(f"位置: {region.bbox}")
        print(f"识别内容: \"{region.text}\"")
        print(f"置信度: {region.confidence:.2f}")
        print("-" * 60)
        
        # 可选：显示图片预览
        if image is not None and self.config.get(
            'rendering.seal_text_handling.interactive_verification.show_image_preview',
            False
        ):
            self._show_region_preview(image, region.bbox)
        
        # 获取用户输入
        while True:
            response = input("识别是否正确？(y=是/n=否/s=跳过): ").strip().lower()
            
            if response in ['y', 'yes', '是', '']:  # 空输入默认为确认
                print("✅ 确认识别正确，继续处理")
                return True, None
            
            elif response in ['n', 'no', '否']:
                corrected_text = input("请输入正确的文字内容: ").strip()
                if corrected_text:
                    print(f"✅ 使用修正后的内容: \"{corrected_text}\"")
                    return True, corrected_text
                else:
                    print("❌ 输入为空，请重新选择")
                    continue
            
            elif response in ['s', 'skip', '跳过']:
                print("⏭️  跳过此区域")
                return False, None
            
            else:
                print("❌ 无效输入，请输入 y/n/s")
                continue
    
    def _show_region_preview(
        self,
        image: np.ndarray,
        bbox: Tuple[int, int, int, int],
        padding: int = 20
    ) -> None:
        """显示文本区域的图片预览。
        
        Args:
            image: 原始图片
            bbox: 文本区域边界框 (x1, y1, x2, y2)
            padding: 预览区域的边距（像素）
        """
        try:
            x1, y1, x2, y2 = bbox
            
            # 添加padding
            h, w = image.shape[:2]
            x1 = max(0, x1 - padding)
            y1 = max(0, y1 - padding)
            x2 = min(w, x2 + padding)
            y2 = min(h, y2 + padding)
            
            # 提取区域
            region_img = image[y1:y2, x1:x2].copy()
            
            # 在区域上绘制边框
            cv2.rectangle(
                region_img,
                (padding, padding),
                (region_img.shape[1] - padding, region_img.shape[0] - padding),
                (0, 255, 0),
                2
            )
            
            # 显示图片
            cv2.imshow("识别区域预览 (按任意键关闭)", region_img)
            cv2.waitKey(2000)  # 显示2秒或等待按键
            cv2.destroyAllWindows()
            
        except Exception as e:
            logger.warning(f"无法显示图片预览: {e}")
