"""OCR 识别服务模块"""

import logging
import requests
from typing import List, Tuple
from pathlib import Path
from PIL import Image

from .models import TextBlock
from .exceptions import OCRError
from .temp_paths import create_named_temporary_file


class OCRService:
    """OCR 识别服务，调用智谱 GLM-OCR API"""
    
    def __init__(self, api_key: str, auto_rotate: bool = True, confidence_threshold: float = 0.75):
        """
        初始化 OCR 服务
        
        Args:
            api_key: 智谱 API 密钥
            auto_rotate: 是否自动旋转图片以提高识别率
            confidence_threshold: 置信度阈值，低于此值时尝试旋转
        """
        self.api_key = api_key
        self.api_url = "https://open.bigmodel.cn/api/paas/v4/files/ocr"
        self.auto_rotate = auto_rotate
        self.confidence_threshold = confidence_threshold
        self.logger = logging.getLogger(__name__)
    
    def recognize(self, image_path: str) -> Tuple[List[TextBlock], int]:
        """
        识别图片中的文字
        
        Args:
            image_path: 图片文件路径
            
        Returns:
            (文字块列表, 旋转角度)
            旋转角度为 0 表示无需旋转，90/180/270 表示顺时针旋转角度
            
        Raises:
            OCRError: OCR 识别失败
            FileNotFoundError: 图片文件不存在
        """
        # 检查文件是否存在
        if not Path(image_path).exists():
            raise FileNotFoundError(f"图片文件不存在: {image_path}")
        
        self.logger.info(f"开始识别图片: {image_path}")
        
        if self.auto_rotate:
            # 使用自动旋转功能
            text_blocks, best_angle = self._recognize_with_rotation(image_path)
            if best_angle != 0:
                self.logger.info(f"自动旋转 {best_angle}° 后识别效果最佳")
        else:
            # 直接识别
            response = self._call_glm_ocr_api(image_path)
            text_blocks = self._parse_ocr_response(response)
            best_angle = 0
        
        self.logger.info(f"识别完成，共识别到 {len(text_blocks)} 个文字块")
        
        return text_blocks, best_angle
    
    def _call_glm_ocr_api(self, image_path: str) -> dict:
        """
        调用 GLM-OCR API
        
        Args:
            image_path: 图片文件路径
            
        Returns:
            API 响应数据
            
        Raises:
            OCRError: API 调用失败
        """
        try:
            # 准备请求头
            headers = {
                "Authorization": f"Bearer {self.api_key}"
            }
            
            # 准备文件和参数
            with open(image_path, 'rb') as f:
                files = {
                    'file': (Path(image_path).name, f, 'image/png')
                }
                
                data = {
                    'tool_type': 'hand_write',  # 手写体识别，适合驾驶证等证件
                    'language_type': 'CHN_ENG',  # 中英文混合识别
                    'probability': 'true'  # 返回置信度信息
                }
                
                # 发送请求（禁用代理，直连API）
                self.logger.debug(f"调用 GLM-OCR API: {self.api_url}")
                response = requests.post(
                    self.api_url,
                    headers=headers,
                    files=files,
                    data=data,
                    timeout=30,
                    proxies={'http': None, 'https': None}
                )
            
            # 检查响应状态
            if response.status_code != 200:
                error_msg = f"API 返回错误状态码: {response.status_code}"
                if response.status_code == 401:
                    error_msg += " - 请检查 API 密钥是否正确配置"
                elif response.status_code == 400:
                    error_msg += f" - 请求参数错误: {response.text}"
                else:
                    error_msg += f" - {response.text}"
                raise OCRError(error_msg)
            
            # 解析 JSON 响应
            response_data = response.json()
            
            # 检查响应中是否有错误
            if 'error' in response_data:
                raise OCRError(f"API 返回错误: {response_data['error']}")
            
            return response_data
            
        except requests.exceptions.Timeout:
            raise OCRError("API 调用超时，请检查网络连接")
        except requests.exceptions.ConnectionError:
            raise OCRError("无法连接到 API 服务器，请检查网络连接")
        except requests.exceptions.RequestException as e:
            raise OCRError(f"API 调用失败: {str(e)}")
        except Exception as e:
            if isinstance(e, OCRError):
                raise
            raise OCRError(f"OCR 识别过程中发生错误: {str(e)}")
    
    def _parse_ocr_response(self, response_data: dict) -> List[TextBlock]:
        """
        解析 OCR API 响应
        
        Args:
            response_data: API 响应数据
            
        Returns:
            文字块列表
            
        Raises:
            OCRError: 响应解析失败
        """
        try:
            text_blocks = []
            
            # 检查响应格式
            if 'words_result' not in response_data:
                raise OCRError("API 响应格式错误: 缺少 words_result 字段")
            
            words_result = response_data['words_result']
            
            # 解析每个文字块
            for item in words_result:
                # 提取文字内容
                if 'words' not in item:
                    self.logger.warning("跳过缺少 words 字段的项")
                    continue
                
                text = item['words']
                
                # 提取位置信息
                if 'location' not in item:
                    self.logger.warning(f"文字块 '{text}' 缺少位置信息，跳过")
                    continue
                
                location = item['location']
                left = location.get('left', 0)
                top = location.get('top', 0)
                width = location.get('width', 0)
                height = location.get('height', 0)
                
                # 构建边界框坐标 [(x1,y1), (x2,y2), (x3,y3), (x4,y4)]
                bounding_box = [
                    (left, top),                    # 左上
                    (left + width, top),            # 右上
                    (left + width, top + height),   # 右下
                    (left, top + height)            # 左下
                ]
                
                # 提取置信度信息
                confidence = 1.0  # 默认置信度
                if 'probability' in item:
                    prob = item['probability']
                    if isinstance(prob, dict) and 'average' in prob:
                        confidence = prob['average']
                    elif isinstance(prob, (int, float)):
                        confidence = float(prob)
                
                # 创建文字块对象
                text_block = TextBlock(
                    text=text,
                    bounding_box=bounding_box,
                    confidence=confidence
                )
                
                text_blocks.append(text_block)
            
            if not text_blocks:
                self.logger.warning("未识别到任何文字块")
            
            return text_blocks
            
        except KeyError as e:
            raise OCRError(f"API 响应格式错误: 缺少必需字段 {str(e)}")
        except Exception as e:
            if isinstance(e, OCRError):
                raise
            raise OCRError(f"解析 OCR 响应时发生错误: {str(e)}")
    
    def _recognize_with_rotation(self, image_path: str) -> Tuple[List[TextBlock], int]:
        """
        尝试不同角度识别图片，返回最佳结果
        
        Args:
            image_path: 图片文件路径
            
        Returns:
            (文字块列表, 最佳旋转角度)
        """
        # 先尝试原始图片
        self.logger.debug("尝试识别原始图片...")
        response = self._call_glm_ocr_api(image_path)
        original_blocks = self._parse_ocr_response(response)
        original_confidence = self._calculate_average_confidence(original_blocks)
        
        self.logger.debug(f"原始图片平均置信度: {original_confidence:.2f}")
        
        # 如果置信度足够高，直接返回
        if original_confidence >= self.confidence_threshold:
            self.logger.debug("原始图片置信度足够高，无需旋转")
            return original_blocks, 0
        
        # 尝试旋转 90°, 180°, 270°
        best_blocks = original_blocks
        best_confidence = original_confidence
        best_angle = 0
        
        for angle in [90, 180, 270]:
            self.logger.debug(f"尝试旋转 {angle}° 后识别...")
            
            # 旋转图片
            rotated_path = self._rotate_image(image_path, angle)
            
            try:
                # 识别旋转后的图片
                response = self._call_glm_ocr_api(rotated_path)
                blocks = self._parse_ocr_response(response)
                confidence = self._calculate_average_confidence(blocks)
                
                self.logger.debug(f"旋转 {angle}° 后平均置信度: {confidence:.2f}")
                
                # 更新最佳结果
                if confidence > best_confidence:
                    best_blocks = blocks
                    best_confidence = confidence
                    best_angle = angle
            finally:
                # 删除临时文件
                Path(rotated_path).unlink(missing_ok=True)
        
        self.logger.info(f"最佳旋转角度: {best_angle}°, 置信度: {best_confidence:.2f}")
        
        return best_blocks, best_angle
    
    def _rotate_image(self, image_path: str, angle: int) -> str:
        """
        旋转图片并保存到临时文件
        
        Args:
            image_path: 原始图片路径
            angle: 旋转角度（顺时针）
            
        Returns:
            旋转后的临时文件路径
        """
        # 打开图片
        img = Image.open(image_path)
        
        # 旋转图片（PIL 的 rotate 是逆时针，所以取负值）
        rotated_img = img.rotate(-angle, expand=True)
        
        # 保存到临时文件
        suffix = Path(image_path).suffix
        with create_named_temporary_file(
            mode='wb',
            suffix=suffix,
            prefix='drivers_license_ocr_rotate_',
            delete=False
        ) as f:
            rotated_img.save(f, format=img.format or 'PNG')
            temp_path = f.name
        
        return temp_path
    
    def _calculate_average_confidence(self, text_blocks: List[TextBlock]) -> float:
        """
        计算文字块的平均置信度
        
        Args:
            text_blocks: 文字块列表
            
        Returns:
            平均置信度
        """
        if not text_blocks:
            return 0.0
        
        total_confidence = sum(block.confidence for block in text_blocks)
        return total_confidence / len(text_blocks)
