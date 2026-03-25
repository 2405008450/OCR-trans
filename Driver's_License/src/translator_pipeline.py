"""主控制器模块 - 驾驶证翻译流程"""

import logging
import os
import cv2
from typing import List, Dict
from pathlib import Path
from datetime import datetime

from .models import LicenseData
from .ocr_service import OCRService
from .field_parser import FieldParser
from .image_extractor import ImageExtractor
from .translation_service import TranslationService
from .document_generator import DocumentGenerator
from .ocr_corrector import OCRCorrector
from .exceptions import TranslationPipelineError


class TranslatorPipeline:
    """驾驶证翻译主流程控制器"""
    
    def __init__(
        self,
        glm_api_key: str,
        deepseek_api_key: str
    ):
        """
        初始化翻译流程
        
        Args:
            glm_api_key: 智谱 GLM API 密钥
            deepseek_api_key: DeepSeek API 密钥
        """
        self.ocr_service = OCRService(glm_api_key)
        self.ocr_corrector = OCRCorrector()
        self.field_parser = FieldParser()
        self.image_extractor = ImageExtractor()
        self.translation_service = TranslationService(deepseek_api_key)
        self.document_generator = DocumentGenerator()
        self.logger = logging.getLogger(__name__)
    
    def translate_image(self, input_path: str, output_dir: str = None) -> str:
        """
        翻译驾驶证图片
        
        Args:
            input_path: 输入图片路径
            output_dir: 输出目录路径（可选，默认为当前目录）
            
        Returns:
            输出文件的完整路径
            
        Raises:
            TranslationPipelineError: 翻译流程失败
        """
        try:
            self.logger.info(f"开始处理图片: {input_path}")
            
            # 1. OCR 识别
            self.logger.info("步骤 1/7: OCR 识别文字")
            text_blocks, rotation_angle = self.ocr_service.recognize(input_path)
            self.logger.info(f"识别到 {len(text_blocks)} 个文字块")
            
            # 2. OCR 结果修正
            self.logger.info("步骤 2/7: 修正 OCR 识别错误")
            text_blocks = self.ocr_corrector.correct_text_blocks(text_blocks)
            
            # 3. 检测是否有副页
            has_duplicate = self._detect_duplicate_page(text_blocks)
            if has_duplicate:
                self.logger.info("检测到副页")
            
            # 4. 如果图片被旋转了，需要旋转图片后再提取照片
            image_for_extraction = input_path
            if rotation_angle != 0:
                self.logger.info(f"旋转图片 {rotation_angle}° 以匹配 OCR 结果")
                image_for_extraction = self._rotate_image_file(input_path, rotation_angle)
            
            # 5. 图像提取（先提取，用于后续印章文字识别）
            self.logger.info("步骤 3/7: 提取照片和印章")
            images = self.image_extractor.extract_images(image_for_extraction)
            self.logger.info(f"提取到 {len(images)} 个图像")
            
            # 清理临时旋转文件
            if rotation_angle != 0 and image_for_extraction != input_path:
                import os
                try:
                    os.unlink(image_for_extraction)
                except:
                    pass
            
            # 计算相片右边界，用于区分左右页
            photo_boundary = None
            if images:
                photo = images[0]  # 第一个图像通常是相片
                photo_x, photo_y = photo.position
                photo_w, photo_h = photo.size
                photo_boundary = photo_x + photo_w  # 相片右边界
                self.logger.info(f"相片位置: ({photo_x}, {photo_y}), 大小: ({photo_w}, {photo_h}), 右边界: {photo_boundary}")
            
            # 5. 识别印章区域的文字（增强版）
            self.logger.info("步骤 4/7: 识别印章区域文字")
            
            # 5.5. 初步检测驾驶证版本（用于条形码识别）
            is_old_version_preliminary = self._detect_license_version_from_text_blocks(text_blocks)
            
            # 5.6. 识别条形码数字
            barcode_number, is_barcode_page = self._extract_barcode_number(text_blocks, is_old_version_preliminary)
            
            if barcode_number:
                self.logger.info(f"识别到条形码数字: {barcode_number}")
            
            # 如果是条形码页面（准驾车型代号规定页面），只在左侧识别印章
            if is_barcode_page:
                # 过滤出左侧的文字块（使用相片边界作为分界线）
                if photo_boundary:
                    left_blocks = [block for block in text_blocks if block.get_center()[0] < photo_boundary]
                    seal_texts = self._extract_seal_texts_enhanced(left_blocks, has_duplicate)
                    self.logger.info(f"准驾车型代号规定页面，只在左侧识别印章")
                else:
                    seal_texts = []
                    self.logger.info("准驾车型代号规定页面，跳过印章识别")
            else:
                seal_texts = self._extract_seal_texts_enhanced(text_blocks, has_duplicate)
                self.logger.info(f"识别到印章文字: {seal_texts}")
            
            # 6. 字段解析
            self.logger.info("步骤 5/7: 解析驾驶证字段")
            fields = self.field_parser.parse_fields(text_blocks, has_duplicate, photo_boundary)
            
            # 修正字段值
            fields = self.ocr_corrector.correct_fields(fields)
            self.logger.info(f"解析到 {len(fields)} 个字段")
            
            # 检测驾驶证版本（旧版有 Valid From 字段，新版只有 Valid Period）
            is_old_version = self._detect_license_version(fields)
            if is_old_version:
                self.logger.info("检测到旧版驾驶证（有 Valid From 字段）")
            else:
                self.logger.info("检测到新版驾驶证（只有 Valid Period 字段）")
            
            # 7. 翻译字段和印章文字
            self.logger.info("步骤 6/7: 翻译字段和印章文字")
            translated_fields = self.translation_service.translate_fields(fields)
            
            # 翻译印章文字
            translated_seal_texts = []
            if seal_texts:
                for text in seal_texts:
                    translated = self.translation_service.translate_text(text)
                    translated_seal_texts.append(translated)
                    self.logger.info(f"印章文字翻译: {text} -> {translated}")
            
            # 8. 生成文档
            self.logger.info("步骤 7/7: 生成 DOCX 文档")
            
            # 获取原图尺寸
            img = cv2.imread(input_path)
            if img is None:
                raise TranslationPipelineError(f"无法读取图片: {input_path}")
            img_height, img_width = img.shape[:2]
            
            # 检测页面类型
            page_types = self._detect_page_types(text_blocks)
            has_main_page = page_types['has_main']
            has_legend_page = page_types['has_legend']
            self.logger.info(f"页面检测: 主页={has_main_page}, 副页={has_duplicate}, 准驾车型代号规定页={has_legend_page}")
            
            # 构建驾驶证数据
            license_data = LicenseData(
                fields=translated_fields,
                images=images,
                image_size=(img_width, img_height),
                text_blocks=text_blocks,
                seal_texts=translated_seal_texts,
                has_duplicate=has_duplicate,
                barcode_number=barcode_number if barcode_number else None,
                is_old_version=is_old_version,
                has_main_page=has_main_page,
                has_legend_page=has_legend_page
            )
            
            # 生成输出文件名
            output_path = self._generate_output_path(input_path, output_dir)
            
            # 生成文档
            self.document_generator.generate_document(license_data, output_path)
            
            self.logger.info(f"翻译完成，输出保存到: {output_path}")
            
            return output_path
            
        except Exception as e:
            self.logger.error(f"翻译流程失败: {str(e)}")
            raise TranslationPipelineError(f"翻译流程失败: {str(e)}")
    
    def translate_batch(
        self,
        input_paths: List[str],
        output_dir: str = None
    ) -> Dict[str, str]:
        """
        批量翻译驾驶证图片
        
        Args:
            input_paths: 输入图片路径列表
            output_dir: 输出目录路径（可选）
            
        Returns:
            {输入路径: 输出路径} 的字典，失败的条目值为错误信息
        """
        results = {}
        
        for input_path in input_paths:
            try:
                output_path = self.translate_image(input_path, output_dir)
                results[input_path] = output_path
            except Exception as e:
                self.logger.error(f"处理 {input_path} 失败: {str(e)}")
                results[input_path] = f"ERROR: {str(e)}"
        
        return results
    
    def translate_merge(
        self,
        input_paths: List[str],
        output_dir: str = None
    ) -> str:
        """
        合并翻译多张驾驶证图片（属于同一个驾驶证）
        
        将多张图片（主页、副页、延伸页等）的内容合并，生成一个完整的文档
        
        Args:
            input_paths: 输入图片路径列表（按顺序）
            output_dir: 输出目录路径（可选）
            
        Returns:
            输出文件的完整路径
            
        Raises:
            TranslationPipelineError: 翻译流程失败
        """
        try:
            self.logger.info(f"开始合并处理 {len(input_paths)} 张图片")
            
            # 存储所有图片的数据
            all_fields = []
            all_images = []
            all_seal_texts = []
            all_text_blocks = []
            has_duplicate = False
            
            # 逐个处理每张图片
            for i, input_path in enumerate(input_paths, 1):
                self.logger.info(f"处理第 {i}/{len(input_paths)} 张图片: {input_path}")
                
                # 1. OCR 识别
                self.logger.info(f"  步骤 1: OCR 识别文字")
                text_blocks, rotation_angle = self.ocr_service.recognize(input_path)
                self.logger.info(f"  识别到 {len(text_blocks)} 个文字块")
                
                # 2. OCR 结果修正
                self.logger.info(f"  步骤 2: 修正 OCR 识别错误")
                text_blocks = self.ocr_corrector.correct_text_blocks(text_blocks)
                all_text_blocks.extend(text_blocks)
                
                # 3. 检测是否有副页
                if self._detect_duplicate_page(text_blocks):
                    has_duplicate = True
                    self.logger.info(f"  检测到副页")
                
                # 4. 如果图片被旋转了，需要旋转图片后再提取照片
                image_for_extraction = input_path
                if rotation_angle != 0:
                    self.logger.info(f"  旋转图片 {rotation_angle}° 以匹配 OCR 结果")
                    image_for_extraction = self._rotate_image_file(input_path, rotation_angle)
                
                # 5. 图像提取
                self.logger.info(f"  步骤 3: 提取照片和印章")
                images = self.image_extractor.extract_images(image_for_extraction)
                self.logger.info(f"  提取到 {len(images)} 个图像")
                all_images.extend(images)
                
                # 清理临时旋转文件
                if rotation_angle != 0 and image_for_extraction != input_path:
                    import os
                    try:
                        os.unlink(image_for_extraction)
                    except:
                        pass
                
                # 计算相片右边界
                photo_boundary = None
                if images:
                    photo = images[0]
                    photo_x, photo_y = photo.position
                    photo_w, photo_h = photo.size
                    photo_boundary = photo_x + photo_w
                    self.logger.info(f"  相片右边界: {photo_boundary}")
                
                # 5. 识别印章文字
                self.logger.info(f"  步骤 4: 识别印章区域文字")
                
                # 5.5. 初步检测驾驶证版本（用于条形码识别）
                is_old_version_preliminary = self._detect_license_version_from_text_blocks(text_blocks)
                
                # 5.6. 识别条形码数字
                barcode_number, is_barcode_page = self._extract_barcode_number(text_blocks, is_old_version_preliminary)
                
                if barcode_number:
                    self.logger.info(f"  识别到条形码数字: {barcode_number}")
                    # 存储条形码数字（只保留第一个）
                    if not hasattr(self, '_barcode_number'):
                        self._barcode_number = barcode_number
                
                # 如果是条形码页面，只在左侧识别印章
                if is_barcode_page:
                    # 过滤出左侧的文字块（使用相片边界作为分界线）
                    if photo_boundary:
                        left_blocks = [block for block in text_blocks if block.get_center()[0] < photo_boundary]
                        seal_texts = self._extract_seal_texts_enhanced(left_blocks, has_duplicate)
                        self.logger.info(f"  条形码页面，只在左侧识别印章")
                        self.logger.info(f"  识别到印章文字: {seal_texts}")
                        all_seal_texts.extend(seal_texts)
                    else:
                        self.logger.info("  条形码页面，跳过印章识别")
                else:
                    seal_texts = self._extract_seal_texts_enhanced(text_blocks, has_duplicate)
                    self.logger.info(f"  识别到印章文字: {seal_texts}")
                    all_seal_texts.extend(seal_texts)
                
                # 6. 字段解析
                self.logger.info(f"  步骤 5: 解析驾驶证字段")
                fields = self.field_parser.parse_fields(text_blocks, has_duplicate, photo_boundary)
                fields = self.ocr_corrector.correct_fields(fields)
                self.logger.info(f"  解析到 {len(fields)} 个字段")
                all_fields.extend(fields)
            
            # 合并字段
            self.logger.info("合并所有字段...")
            merged_fields = self._merge_fields(all_fields)
            self.logger.info(f"合并后共 {len(merged_fields)} 个字段")
            
            # 检测驾驶证版本（旧版有 Valid From 字段，新版只有 Valid Period）
            is_old_version = self._detect_license_version(merged_fields)
            if is_old_version:
                self.logger.info("检测到旧版驾驶证（有 Valid From 字段）")
            else:
                self.logger.info("检测到新版驾驶证（只有 Valid Period 字段）")
            
            # 去重印章文字
            unique_seal_texts = list(set(all_seal_texts))
            self.logger.info(f"去重后印章文字: {unique_seal_texts}")
            
            # 7. 翻译字段和印章文字
            self.logger.info("步骤 6/7: 翻译字段和印章文字")
            translated_fields = self.translation_service.translate_fields(merged_fields)
            
            # 翻译印章文字
            translated_seal_texts = []
            if unique_seal_texts:
                for text in unique_seal_texts:
                    translated = self.translation_service.translate_text(text)
                    translated_seal_texts.append(translated)
                    self.logger.info(f"印章文字翻译: {text} -> {translated}")
            
            # 8. 生成文档
            self.logger.info("步骤 7/7: 生成 DOCX 文档")
            
            # 使用第一张图片获取尺寸
            img = cv2.imread(input_paths[0])
            if img is None:
                raise TranslationPipelineError(f"无法读取图片: {input_paths[0]}")
            img_height, img_width = img.shape[:2]
            
            # 获取条形码数字（如果有）
            barcode_number = getattr(self, '_barcode_number', None)
            if hasattr(self, '_barcode_number'):
                delattr(self, '_barcode_number')  # 清理临时属性
            
            # 检测页面类型（从所有文字块中检测）
            page_types = self._detect_page_types(all_text_blocks)
            has_main_page = page_types['has_main']
            has_legend_page = page_types['has_legend']
            self.logger.info(f"页面检测: 主页={has_main_page}, 副页={has_duplicate}, 准驾车型代号规定页={has_legend_page}")
            
            # 构建驾驶证数据
            license_data = LicenseData(
                fields=translated_fields,
                images=all_images,
                image_size=(img_width, img_height),
                text_blocks=all_text_blocks,
                seal_texts=translated_seal_texts,
                has_duplicate=has_duplicate,
                barcode_number=barcode_number,
                is_old_version=is_old_version,
                has_main_page=has_main_page,
                has_legend_page=has_legend_page
            )
            
            # 生成输出文件名（使用第一张图片的名称）
            output_path = self._generate_output_path(input_paths[0], output_dir)
            
            # 生成文档
            self.document_generator.generate_document(license_data, output_path)
            
            self.logger.info(f"合并翻译完成，输出保存到: {output_path}")
            
            return output_path
            
        except Exception as e:
            self.logger.error(f"合并翻译流程失败: {str(e)}")
            raise TranslationPipelineError(f"合并翻译流程失败: {str(e)}")
    
    def _generate_output_path(self, input_path: str, output_dir: str = None) -> str:
        """
        生成输出文件路径
        
        Args:
            input_path: 输入文件路径
            output_dir: 输出目录（可选）
            
        Returns:
            输出文件的完整路径
        """
        # 获取输入文件名（不含扩展名）
        input_filename = os.path.basename(input_path)
        name, _ = os.path.splitext(input_filename)
        
        # 生成时间戳（精确到分钟）
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        
        # 生成输出文件名
        output_filename = f"{name}_translation_{timestamp}.docx"
        
        # 确定输出目录
        if output_dir is None:
            output_dir = os.getcwd()
        
        # 创建输出目录（如果不存在）
        os.makedirs(output_dir, exist_ok=True)
        
        # 生成完整路径
        output_path = os.path.join(output_dir, output_filename)
        
        return output_path
    
    def _extract_seal_texts_by_position(self, text_blocks: List) -> List[str]:
        """
        根据位置识别左侧印章区域的文字
        
        印章区域特征：
        - 位于图片左侧（X坐标较小，通常 < 150）
        - 多行文字垂直排列
        - Y坐标在中间区域（250-380之间）
        - 排除英文标签（如 Address, Valid Period）
        
        Args:
            text_blocks: OCR识别的所有文字块
            
        Returns:
            印章文字列表
        """
        # 筛选左侧印章区域的文字块
        seal_blocks = []
        for block in text_blocks:
            x, y = block.get_center()
            text = block.text.strip()
            
            # 左侧印章区域：X < 150，Y在250-380之间
            # 排除纯英文标签
            if x < 150 and 250 < y < 380:
                # 检查是否包含中文字符
                if any('\u4e00' <= char <= '\u9fff' for char in text):
                    seal_blocks.append((y, block))
        
        if not seal_blocks:
            return []
        
        # 按Y坐标排序
        seal_blocks.sort(key=lambda b: b[0])
        
        # 合并文字
        combined_text = ''.join([b[1].text for b in seal_blocks])
        
        return [combined_text] if combined_text else []
    
    def _detect_duplicate_page(self, text_blocks: List) -> bool:
        """
        检测是否有副页
        
        通过查找 "Duplicate of Driving License" 或 "副页" 文字来判断
        
        Args:
            text_blocks: OCR识别的所有文字块
            
        Returns:
            True表示有副页，False表示没有
        """
        for block in text_blocks:
            text = block.text.strip()
            if "Duplicate" in text or "副页" in text:
                return True
        return False
    
    def _detect_main_page(self, text_blocks: List) -> bool:
        """
        检测是否有主页
        
        通过查找 "Driving License of the People's Republic of China" 或 
        "中华人民共和国机动车驾驶证" 文字来判断（排除副页标题）
        
        Args:
            text_blocks: OCR识别的所有文字块
            
        Returns:
            True表示有主页，False表示没有
        """
        for block in text_blocks:
            text = block.text.strip()
            # 主页标题（排除副页）
            if ("Driving License" in text or "驾驶证" in text) and "Duplicate" not in text and "副页" not in text:
                # 进一步验证：主页应该有姓名、证号等字段标签
                return True
        
        # 备用检测：检查是否有主页特有的字段标签
        main_page_labels = ["姓名", "Name", "证号", "License No", "准驾车型", "Class"]
        for block in text_blocks:
            text = block.text.strip()
            for label in main_page_labels:
                if label in text:
                    return True
        
        return False
    
    def _detect_legend_page(self, text_blocks: List) -> bool:
        """
        检测是否有准驾车型代号规定页
        
        通过查找 "Legend for Class of Vehicles" 或 "准驾车型代号规定" 文字来判断
        
        Args:
            text_blocks: OCR识别的所有文字块
            
        Returns:
            True表示有准驾车型代号规定页，False表示没有
        """
        for block in text_blocks:
            text = block.text.strip()
            if "Legend" in text or "准驾车型代号规定" in text:
                return True
            # 备用检测：检查是否有准驾车型代号规定页特有的内容
            if "代号" in text and "规定" in text:
                return True
            # 检测车型说明文字
            if any(kw in text for kw in ["大型客车", "牵引车", "城市公交车", "中型客车", "大型货车"]):
                return True
        return False
    
    def _detect_page_types(self, text_blocks: List) -> dict:
        """
        检测图片中包含的所有页面类型
        
        Args:
            text_blocks: OCR识别的所有文字块
            
        Returns:
            包含页面类型信息的字典：
            {
                'has_main': bool,      # 是否有主页
                'has_duplicate': bool, # 是否有副页
                'has_legend': bool     # 是否有准驾车型代号规定页
            }
        """
        return {
            'has_main': self._detect_main_page(text_blocks),
            'has_duplicate': self._detect_duplicate_page(text_blocks),
            'has_legend': self._detect_legend_page(text_blocks)
        }
    
    def _extract_seal_texts_enhanced(self, text_blocks: List, has_duplicate: bool = False) -> List[str]:
        """
        增强版印章文字识别
        
        策略：
        1. 优先使用"发证机关"字段的值（通常就是印章文字）
        2. 查找包含关键词的文字块，并收集其附近的文字块
        3. 如果没有识别到发证机关，使用位置识别
        4. 验证印章文字格式（应包含"公安"、"交警"等关键词）
        
        Args:
            text_blocks: OCR识别的所有文字块
            has_duplicate: 是否有副页
            
        Returns:
            印章文字列表
        """
        seal_texts = []
        
        # 策略1：查找"发证机关"相关的文字
        # 发证机关通常格式为：XX省XX市公安局交通警察支队
        
        # 第一步：找到包含关键词的核心块
        keywords = ['公安', '交警', '交通警察', '支队', '大队', '车管']
        core_blocks = []
        
        for block in text_blocks:
            text = block.text.strip()
            
            # 检查是否包含发证机关的关键词
            if any(kw in text for kw in keywords):
                # 检查是否包含足够的中文字符（至少3个）
                chinese_chars = sum(1 for c in text if '\u4e00' <= c <= '\u9fff')
                if chinese_chars >= 3:
                    core_blocks.append(block)
        
        if not core_blocks:
            return self._extract_seal_texts_by_position_fallback(text_blocks, has_duplicate)
        
        # 第二步：收集核心块附近的所有相关文字块
        # 策略：收集与核心块在同一区域（Y坐标接近，X坐标在左侧）的所有中文文字块
        issuing_authority_blocks = []
        
        # 计算核心块的Y坐标范围
        core_y_coords = [b.get_center()[1] for b in core_blocks]
        min_y = min(core_y_coords) - 200  # 向上扩展200像素（进一步增加范围）
        max_y = max(core_y_coords) + 100  # 向下扩展100像素
        
        # 计算核心块的X坐标范围
        core_x_coords = [b.get_center()[0] for b in core_blocks]
        min_x = min(core_x_coords) - 150  # 向左扩展150像素
        max_x = max(core_x_coords) + 150  # 向右扩展150像素
        
        # 需要排除的文字（字段标签和字段值）
        exclude_texts = {
            '住址', '姓名', '性别', '国籍', '记录', '出生日期', '初次领证日期',
            '准驾车型', '有效期限', '证号', '档案编号',
            'Address', 'Name', 'Sex', 'Nationality', 'Record', 'Date of Birth',
            'Date of First Issue', 'Class', 'Valid Period', 'License No',
            'File No', 'Issuing Authority'
        }
        
        # 需要排除的模式（日期、地址、姓名等）
        exclude_patterns = [
            r'\d{4}[-年./]\d{1,2}[-月./]\d{1,2}',  # 日期格式
            r'Room\s+\d+',  # 房间号
            r'No\.\s*\d+',  # 编号
            r'Long-term',  # 长期
            r'to\s+',  # 包含to的（日期范围）
            r'\d{3,}',  # 长数字串（证号、档案号等）
        ]
        
        # 需要排除的关键词（明显不是发证机关的内容）
        exclude_keywords = [
            '家属', '小区', '大厦', '楼', '室', '号', '路', '街', '巷', '村', '镇', '乡',
            # 准驾车型代号规定页面的说明文字关键词
            '汽车', '载货', '载客', '摩托车', '轮式', '三轮', '低速', '牵引车', '挂车',
            '客车', '货车', '专用', '作业', '无轨电车', '有轨电车', '自动挡', '手动挡'
        ]
        
        # 需要排除的英文单词（印章周围的英文说明）
        exclude_english_words = [
            'Except', 'for', 'public', 'security', 'traffic', 'management',
            'Traffic', 'Police', 'Detachment', 'Public', 'Security', 'Bureau',
            'Province', 'City', 'Long-term', 'Room', 'Building', 'Family', 'Courtyard',
            'of', 'the', 'and', 'in', 'to', 'from'
        ]
        
        for block in text_blocks:
            x, y = block.get_center()
            text = block.text.strip()
            
            # 检查是否在搜索区域内
            if min_x <= x <= max_x and min_y <= y <= max_y:
                # 检查是否包含中文字符
                chinese_chars = sum(1 for c in text if '\u4e00' <= c <= '\u9fff')
                if chinese_chars >= 2:
                    # 排除字段标签
                    if text not in exclude_texts:
                        # 排除包含字段标签的文本
                        is_label = any(label in text for label in exclude_texts)
                        
                        # 排除匹配排除模式的文本
                        import re
                        is_excluded_pattern = any(re.search(pattern, text) for pattern in exclude_patterns)
                        
                        # 排除包含地址关键词的文本
                        has_address_keyword = any(kw in text for kw in exclude_keywords)
                        
                        # 排除纯英文或包含英文单词的文本
                        has_english_word = any(word in text for word in exclude_english_words)
                        
                        # 包含发证机关关键词，或者是省市名称（包含"省"或"市"但不包含地址关键词）
                        has_authority_keyword = any(kw in text for kw in keywords)
                        is_location = ('省' in text or '市' in text) and not has_address_keyword
                        
                        if not is_label and not is_excluded_pattern and not has_english_word and (has_authority_keyword or is_location):
                            issuing_authority_blocks.append(block)
        
        if not issuing_authority_blocks:
            return self._extract_seal_texts_by_position_fallback(text_blocks, has_duplicate)
        
        # 第三步：按位置排序并合并
        # 分析文字块的排列方式：
        # 1. 如果X坐标变化大，Y坐标变化小 -> 水平排列（圆形印章）
        # 2. 如果X坐标变化小，Y坐标变化大 -> 垂直排列（OCR按行识别）
        
        x_coords = [b.get_center()[0] for b in issuing_authority_blocks]
        y_coords = [b.get_center()[1] for b in issuing_authority_blocks]
        
        x_range = max(x_coords) - min(x_coords) if len(x_coords) > 1 else 0
        y_range = max(y_coords) - min(y_coords) if len(y_coords) > 1 else 0
        
        # 判断排列方式
        if y_range > x_range * 1.5:  # Y范围明显大于X范围，说明是垂直排列
            # 垂直排列：直接按Y坐标从上到下排序
            issuing_authority_blocks.sort(key=lambda b: b.get_center()[1])
            combined_text = ''.join([b.text.strip() for b in issuing_authority_blocks])
        else:
            # 圆形印章文字排列规律：
            # - 上半圆：文字从右到左排列（逆时针），但阅读顺序是从左到右
            # - 下半圆：文字从左到右排列（顺时针），阅读顺序也是从左到右
            
            # 计算Y坐标的中位数作为上下半圆的分界线
            if len(y_coords) >= 2:
                y_median = (min(y_coords) + max(y_coords)) / 2
            else:
                y_median = y_coords[0] if y_coords else 0
            
            # 分离上半圆和下半圆的文字块
            upper_blocks = [b for b in issuing_authority_blocks if b.get_center()[1] < y_median]
            lower_blocks = [b for b in issuing_authority_blocks if b.get_center()[1] >= y_median]
            
            # 上半圆：按X坐标从左到右排序（阅读顺序）
            upper_blocks.sort(key=lambda b: b.get_center()[0])
            # 下半圆：按X坐标从左到右排序（阅读顺序）
            lower_blocks.sort(key=lambda b: b.get_center()[0])
            
            # 合并上半圆文字
            upper_text = ''.join([b.text.strip() for b in upper_blocks])
            
            # 合并下半圆文字
            lower_text = ''.join([b.text.strip() for b in lower_blocks])
            
            # 最终合并：上半圆 + 下半圆
            combined_text = upper_text + lower_text
        
        # 验证格式：应该包含"公安"或"交警"，且符合发证机关格式
        if '公安' in combined_text or '交警' in combined_text:
            # 进一步验证：应该包含省市信息和机关名称
            # 标准格式：XX省XX市公安局交通警察支队
            has_province = '省' in combined_text
            has_city = '市' in combined_text
            has_authority = any(kw in combined_text for kw in ['公安局', '交警', '交通警察'])
            
            if has_province and has_city and has_authority:
                # 清理可能的多余字符
                # 移除可能的空格、换行等
                cleaned_text = combined_text.replace(' ', '').replace('\n', '').replace('\r', '')
                
                # 验证长度合理（通常10-30个字符）
                if 10 <= len(cleaned_text) <= 30:
                    seal_texts.append(cleaned_text)
                    print(f"[印章识别] 结果: {cleaned_text}")
                    return seal_texts
                else:
                    pass
            else:
                pass
        else:
            pass
        
        return self._extract_seal_texts_by_position_fallback(text_blocks, has_duplicate)
    
    def _extract_seal_texts_by_position_fallback(self, text_blocks: List, has_duplicate: bool = False) -> List[str]:
        """
        备用方案：基于位置识别印章文字
        
        Args:
            text_blocks: OCR识别的所有文字块
            has_duplicate: 是否有副页
            
        Returns:
            印章文字列表
        """
    def _extract_seal_texts_by_position_fallback(self, text_blocks: List, has_duplicate: bool = False) -> List[str]:
        """
        备用方案：基于位置识别印章文字
        
        Args:
            text_blocks: OCR识别的所有文字块
            has_duplicate: 是否有副页
            
        Returns:
            印章文字列表
        """
        # 筛选左侧印章区域的文字块
        seal_blocks = []
        
        # 根据是否有副页调整X坐标范围
        x_ranges = [(0, 150)]
        if has_duplicate:
            x_ranges.append((450, 650))
        
        # 需要排除的文字（字段标签和常见字段值）
        exclude_texts = {
            '住址', '姓名', '性别', '国籍', '记录', '出生日期', '初次领证日期',
            '准驾车型', '有效期限', '证号', '档案编号', '发证机关',
            'Address', 'Name', 'Sex', 'Nationality', 'Record', 'Date of Birth',
            'Date of First Issue', 'Class', 'Valid Period', 'License No',
            'File No', 'Issuing Authority'
        }
        
        # 需要排除的内容模式（记录字段的常见内容）
        exclude_patterns = [
            '请于', '提交', '身体条件', '证明', '之后每年', '日内',
            '年审', '换证', '体检', '有效期',
            # 准驾车型代号规定页面的说明文字关键词
            '汽车', '载货', '载客', '摩托车', '轮式', '三轮', '低速', '牵引车', '挂车',
            '客车', '货车', '专用', '作业', '无轨电车', '有轨电车', '自动挡', '手动挡'
        ]
        
        for block in text_blocks:
            x, y = block.get_center()
            text = block.text.strip()
            
            # 检查是否在印章区域
            in_seal_area = False
            for x_min, x_max in x_ranges:
                # 扩大Y坐标范围：200-500（覆盖更大范围）
                if x_min < x < x_max and 200 < y < 500:
                    in_seal_area = True
                    break
            
            if in_seal_area:
                # 检查是否包含中文字符
                if any('\u4e00' <= char <= '\u9fff' for char in text):
                    # 排除字段标签
                    if text not in exclude_texts:
                        # 排除包含字段标签的文本
                        is_label = False
                        for label in exclude_texts:
                            if label in text:
                                is_label = True
                                break
                        
                        # 排除记录字段的内容模式
                        is_record_content = False
                        for pattern in exclude_patterns:
                            if pattern in text:
                                is_record_content = True
                                break
                        
                        if not is_label and not is_record_content:
                            seal_blocks.append((y, x, block))
        
        if not seal_blocks:
            return []
        
        # 按Y坐标排序
        seal_blocks.sort(key=lambda b: b[0])
        
        # 智能合并：如果Y坐标相近（差距<30），按X坐标排序后合并
        # 否则按Y坐标顺序合并
        merged_lines = []
        current_line = []
        last_y = None
        
        for y, x, block in seal_blocks:
            if last_y is None or abs(y - last_y) < 30:
                # 同一行
                current_line.append((x, block))
                last_y = y
            else:
                # 新的一行
                if current_line:
                    # 按X坐标排序
                    current_line.sort(key=lambda b: b[0])
                    line_text = ''.join([b[1].text for b in current_line])
                    merged_lines.append(line_text)
                current_line = [(x, block)]
                last_y = y
        
        # 处理最后一行
        if current_line:
            current_line.sort(key=lambda b: b[0])
            line_text = ''.join([b[1].text for b in current_line])
            merged_lines.append(line_text)
        
        # 合并所有行
        combined_text = ''.join(merged_lines)
        
        print(f"[印章识别] 位置识别结果: {combined_text}")
        
        return [combined_text] if combined_text else []
    
    def _merge_fields(self, all_fields: List) -> List:
        """
        合并多张图片的字段
        
        策略：
        1. 普通字段：保留第一个非空值
        2. 记录字段：合并所有内容
        3. 去重：相同字段名只保留一个
        
        Args:
            all_fields: 所有图片的字段列表
            
        Returns:
            合并后的字段列表
        """
        from src.models import LicenseField
        
        # 按字段名分组
        field_groups = {}
        for field in all_fields:
            field_name = field.field_name
            if field_name not in field_groups:
                field_groups[field_name] = []
            field_groups[field_name].append(field)
        
        # 合并字段
        merged_fields = []
        for field_name, fields in field_groups.items():
            if field_name == "记录":
                # 记录字段：合并所有内容
                all_records = []
                for field in fields:
                    if field.field_value and field.field_value.strip():
                        all_records.append(field.field_value.strip())
                
                if all_records:
                    # 去重并合并
                    unique_records = []
                    seen = set()
                    for record in all_records:
                        if record not in seen:
                            unique_records.append(record)
                            seen.add(record)
                    
                    merged_value = "".join(unique_records)
                    merged_field = LicenseField(
                        field_name="记录",
                        field_value=merged_value,
                        position=fields[0].position
                    )
                    merged_field.translated_value = fields[0].translated_value if hasattr(fields[0], 'translated_value') else None
                    merged_fields.append(merged_field)
                    self.logger.info(f"合并记录字段: {len(all_records)} 个片段 -> {len(merged_value)} 字符")
            else:
                # 普通字段：保留第一个非空值
                for field in fields:
                    if field.field_value and field.field_value.strip():
                        merged_fields.append(field)
                        break
        
        return merged_fields


    def _extract_barcode_number(self, text_blocks: List, is_old_version: bool = False) -> tuple:
        """
        识别条形码下方的数字
        
        新版驾驶证：
        - 条形码在"准驾车型代号规定"页面左下角
        - 页面顶部有"准驾车型代号规定"或"Legend for Class of Vehicles"标题
        
        旧版驾驶证：
        - 条形码在副页右下角
        - 副页有"Duplicate of Driving License"或"副页"标题
        
        Args:
            text_blocks: OCR识别的所有文字块
            is_old_version: 是否是旧版驾驶证
            
        Returns:
            (条形码数字, 是否是条形码页面)
            条形码数字如果未找到则返回None
        """
        # 检查是否是准驾车型代号规定页面（新版）
        is_legend_page = False
        for block in text_blocks:
            text = block.text.strip()
            if ("准驾车型代号" in text or "代号规定" in text or 
                "Legend" in text or "Class of Vehicles" in text):
                is_legend_page = True
                break
        
        # 检查是否是副页（旧版条形码在副页）
        is_duplicate_page = False
        for block in text_blocks:
            text = block.text.strip()
            if "Duplicate" in text or "副页" in text:
                is_duplicate_page = True
                break
        
        # 确定是否是条形码页面
        is_barcode_page = is_legend_page or (is_old_version and is_duplicate_page)
        
        if not is_barcode_page:
            return None, False
        
        # 查找纯数字文本块（条形码数字通常是长串数字）
        barcode_candidates = []
        
        for block in text_blocks:
            text = block.text.strip()
            # 移除空格、连字符、星号、加号等特殊字符
            clean_text = text.replace(' ', '').replace('-', '').replace('*', '').replace('[', '').replace(']', '').replace('+', '').replace('_', '')
            
            # 检查是否是纯数字且长度合理（至少6位）
            if clean_text.isdigit() and len(clean_text) >= 6:
                x, y = block.get_center()
                barcode_candidates.append({
                    'text': text,
                    'clean_text': clean_text,
                    'position': (x, y),
                    'block': block
                })
        
        if not barcode_candidates:
            return None, is_barcode_page
        
        # 选择最可能的条形码数字
        if len(barcode_candidates) == 1:
            result = barcode_candidates[0]['clean_text']
            print(f"[条形码识别] 结果: {result}")
            return result, is_barcode_page
        
        # 多个候选时，根据版本选择不同策略
        if is_old_version and is_duplicate_page:
            # 旧版：条形码在副页右下角（X大，Y大）
            for candidate in barcode_candidates:
                x, y = candidate['position']
                # 得分 = Y坐标（越大越好）+ X坐标/10（越大越好）+ 长度*10（越长越好）
                candidate['score'] = y + x/10 + len(candidate['clean_text']) * 10
            print(f"[条形码识别] 旧版驾驶证，在副页右下角查找")
        else:
            # 新版：条形码在准驾车型代号规定页面左下角（X小，Y大）
            for candidate in barcode_candidates:
                x, y = candidate['position']
                # 得分 = Y坐标（越大越好）- X坐标/10（越小越好）+ 长度*10（越长越好）
                candidate['score'] = y - x/10 + len(candidate['clean_text']) * 10
            print(f"[条形码识别] 新版驾驶证，在准驾车型代号规定页面左下角查找")
        
        best_candidate = max(barcode_candidates, key=lambda c: c['score'])
        result = best_candidate['clean_text']
        print(f"[条形码识别] 结果: {result}")
        return result, is_barcode_page

    def _rotate_image_file(self, image_path: str, angle: int) -> str:
        """
        旋转图片文件并保存到临时文件
        
        Args:
            image_path: 原始图片路径
            angle: 旋转角度（顺时针，0/90/180/270）
            
        Returns:
            旋转后的临时文件路径
        """
        import tempfile
        from PIL import Image
        
        # 打开图片
        img = Image.open(image_path)
        
        # 旋转图片（PIL 的 rotate 是逆时针，所以取负值）
        rotated_img = img.rotate(-angle, expand=True)
        
        # 保存到临时文件
        suffix = Path(image_path).suffix
        with tempfile.NamedTemporaryFile(mode='wb', suffix=suffix, delete=False) as f:
            rotated_img.save(f, format=img.format or 'PNG')
            temp_path = f.name
        
        return temp_path

    def _detect_license_version(self, fields: List) -> bool:
        """
        检测驾驶证版本（基于解析后的字段）
        
        旧版驾驶证特征：
        - 主页有 "有效期起始日期"（Valid From）字段
        - 主页有 "有效期限"（Valid For）字段
        
        新版驾驶证特征：
        - 主页只有 "有效期限"（Valid Period）字段
        - 没有 "有效期起始日期" 字段
        
        Args:
            fields: 字段列表
            
        Returns:
            True 表示旧版驾驶证，False 表示新版驾驶证
        """
        has_valid_from = False
        
        for field in fields:
            # 检查标准名称和 OCR 变体
            if field.field_name in ["有效期起始日期", "有效起始日期"]:
                has_valid_from = True
                break
        
        if has_valid_from:
            print(f"[版本检测] 检测到 '有效期起始日期' 字段，判定为旧版驾驶证")
            return True
        else:
            print(f"[版本检测] 未检测到 '有效期起始日期' 字段，判定为新版驾驶证")
            return False

    def _detect_license_version_from_text_blocks(self, text_blocks: List) -> bool:
        """
        从 OCR 文本块中初步检测驾驶证版本
        
        旧版驾驶证特征：
        - 有 "Valid From" 英文标签
        
        新版驾驶证特征：
        - 没有 "Valid From" 标签，只有 "Valid Period"
        
        Args:
            text_blocks: OCR 识别的文字块列表
            
        Returns:
            True 表示旧版驾驶证，False 表示新版驾驶证
        """
        for block in text_blocks:
            text = block.text.strip()
            # 检查是否有 "Valid From" 标签（旧版特有）
            if "Valid From" in text:
                print(f"[版本预检测] 检测到 'Valid From' 标签，判定为旧版驾驶证")
                return True
        
        print(f"[版本预检测] 未检测到 'Valid From' 标签，判定为新版驾驶证")
        return False
