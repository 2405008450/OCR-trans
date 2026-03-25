"""图像提取模块"""

import cv2
import numpy as np
from typing import List, Optional, Tuple
from .models import ExtractedImage
from .exceptions import ImageLoadError


class ImageExtractor:
    """图像提取器，从驾驶证图片中提取人像照片"""
    
    def __init__(self):
        """初始化图像提取器"""
        self._mtcnn = None
    
    @property
    def mtcnn(self):
        """延迟加载 MTCNN 检测器"""
        if self._mtcnn is None:
            from mtcnn import MTCNN
            self._mtcnn = MTCNN()
        return self._mtcnn
    
    def extract_images(self, image_path: str) -> List[ExtractedImage]:
        """
        从驾驶证图片中提取人像照片
        
        Args:
            image_path: 图片文件路径
            
        Returns:
            提取的图像列表
            
        Raises:
            ImageLoadError: 图片加载失败
        """
        # 加载图片
        image = cv2.imread(image_path)
        if image is None:
            raise ImageLoadError(f"无法加载图片: {image_path}")
        
        extracted_images = []
        
        # 提取照片
        photo = self._extract_photo(image)
        if photo:
            extracted_images.append(photo)
        
        return extracted_images
    
    def _extract_photo(self, image: np.ndarray) -> Optional[ExtractedImage]:
        """
        提取驾驶证人像照片
        
        策略：
        1. 先用 MTCNN 检测人脸位置
        2. 在人脸周围区域用边缘检测找照片边框
        3. 如果找不到边框，回退到基于人脸比例计算
        
        Args:
            image: 原始图片数组 (BGR 格式)
            
        Returns:
            提取的照片，如果未找到则返回 None
        """
        # MTCNN 需要 RGB 格式
        rgb_image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
        
        # 检测人脸
        faces = self.mtcnn.detect_faces(rgb_image)
        
        print(f"[DEBUG] MTCNN 检测到 {len(faces)} 个人脸")
        for i, face in enumerate(faces):
            print(f"[DEBUG]   人脸 {i+1}: box={face['box']}, confidence={face['confidence']:.3f}")
        
        if not faces:
            return None
        
        # 选择置信度最高的人脸
        best_face = max(faces, key=lambda f: f['confidence'])
        
        if best_face['confidence'] < 0.8:
            print(f"[DEBUG] 置信度 {best_face['confidence']:.3f} 低于阈值 0.8，跳过")
            return None
        
        fx, fy, fw, fh = best_face['box']
        keypoints = best_face['keypoints']
        
        # 尝试用边缘检测找照片边框
        photo_bounds = self._detect_photo_border(image, fx, fy, fw, fh)
        
        if photo_bounds:
            x, y, w, h = photo_bounds
            print(f"[DEBUG] 边缘检测找到照片边框: ({x}, {y}, {w}, {h})")
        else:
            # 回退到基于人脸比例计算
            x, y, w, h = self._calculate_photo_bounds(
                image, fx, fy, fw, fh, keypoints
            )
            print(f"[DEBUG] 使用比例计算: ({x}, {y}, {w}, {h})")
        
        photo_data = image[y:y+h, x:x+w]
        
        return ExtractedImage(
            image_type="photo",
            image_data=photo_data,
            position=(x, y),
            size=(w, h)
        )
    
    def _detect_photo_border(
        self,
        image: np.ndarray,
        fx: int, fy: int, fw: int, fh: int
    ) -> Optional[Tuple[int, int, int, int]]:
        """
        用边缘检测在人脸周围找照片边框
        
        Args:
            image: 原始图片
            fx, fy, fw, fh: 人脸边界框
            
        Returns:
            (x, y, w, h) 或 None
        """
        img_h, img_w = image.shape[:2]
        
        # 在人脸周围扩展搜索区域（缩小搜索范围）
        margin = int(max(fw, fh) * 0.8)
        search_x1 = max(0, fx - margin)
        search_y1 = max(0, fy - margin)
        search_x2 = min(img_w, fx + fw + margin)
        search_y2 = min(img_h, fy + fh + int(margin * 1.5))  # 下方多留一点空间
        
        roi = image[search_y1:search_y2, search_x1:search_x2]
        
        # 转灰度 + 边缘检测
        gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
        # 用自适应阈值增强边缘
        blurred = cv2.GaussianBlur(gray, (3, 3), 0)
        edges = cv2.Canny(blurred, 30, 100)
        
        # 膨胀边缘使轮廓更连续
        kernel = np.ones((2, 2), np.uint8)
        edges = cv2.dilate(edges, kernel, iterations=1)
        
        # 找轮廓
        contours, _ = cv2.findContours(edges, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
        
        # 人脸中心和边界（相对于搜索区域）
        face_cx = fx + fw // 2 - search_x1
        face_cy = fy + fh // 2 - search_y1
        face_area = fw * fh
        
        candidates = []
        
        for contour in contours:
            # 近似为多边形
            peri = cv2.arcLength(contour, True)
            approx = cv2.approxPolyDP(contour, 0.02 * peri, True)
            
            # 只要矩形（4个顶点）
            if len(approx) != 4:
                continue
            
            rx, ry, rw, rh = cv2.boundingRect(approx)
            
            # 过滤条件
            # 1. 矩形必须包含人脸中心
            if not (rx < face_cx < rx + rw and ry < face_cy < ry + rh):
                continue
            
            # 2. 宽高比接近证件照 (0.68 ~ 0.82)
            aspect = rw / rh if rh > 0 else 0
            if not (0.68 < aspect < 0.82):
                continue
            
            # 3. 面积合理（比人脸大 1.8-4 倍）
            rect_area = rw * rh
            area_ratio = rect_area / face_area
            if area_ratio < 1.8 or area_ratio > 4.0:
                continue
            
            candidates.append((rx, ry, rw, rh, rect_area, aspect))
        
        if not candidates:
            return None
        
        # 选择面积最小且宽高比最接近 0.75 的
        # 优先选小的，避免框太大
        best = min(candidates, key=lambda c: (c[4] / face_area, abs(c[5] - 0.75)))
        rx, ry, rw, rh = best[:4]
        
        return (rx + search_x1, ry + search_y1, rw, rh)
    
    def _calculate_photo_bounds(
        self,
        image: np.ndarray,
        fx: int, fy: int, fw: int, fh: int,
        keypoints: dict
    ) -> Tuple[int, int, int, int]:
        """
        根据人脸边界框计算照片截取范围
        
        根据图片分辨率动态调整比例：
        - 低分辨率图片：人脸检测框相对较小，需要更大的扩展比例
        - 高分辨率图片：人脸检测框相对较大，用较小的扩展比例
        
        Args:
            image: 原始图片
            fx, fy, fw, fh: 人脸边界框
            keypoints: 人脸关键点
            
        Returns:
            (x, y, w, h) 照片截取范围
        """
        img_height, img_width = image.shape[:2]
        
        # 根据人脸占图片的比例动态调整
        # 人脸占比越大，说明照片区域在图中占比也大，需要更小的扩展
        total_pixels = img_height * img_width
        face_area = fw * fh
        face_ratio_in_image = face_area / total_pixels
        
        # 基准：人脸占图片 1.3%（典型驾驶证扫描件）
        base_ratio = 0.013
        
        # 调整因子：占比大则扩展小，占比小则扩展大
        adjust_factor = (base_ratio / face_ratio_in_image) ** 0.4
        adjust_factor = max(0.9, min(1.15, adjust_factor))
        
        # 基础比例
        base_face_ratio_h = 0.62
        base_face_ratio_w = 0.58
        
        # 应用调整
        face_ratio_h = base_face_ratio_h / adjust_factor
        face_ratio_w = base_face_ratio_w / adjust_factor
        
        # 方法1: 基于人脸高度
        photo_h1 = int(fh / face_ratio_h)
        photo_w1 = int(photo_h1 * 0.68)  # 宽高比更小，高度更高
        
        # 方法2: 基于人脸宽度
        photo_w2 = int(fw / face_ratio_w)
        photo_h2 = int(photo_w2 / 0.68)
        
        # 取较大的估算值
        if photo_h1 * photo_w1 > photo_h2 * photo_w2:
            photo_height, photo_width = photo_h1, photo_w1
        else:
            photo_height, photo_width = photo_h2, photo_w2
        
        # 照片中心
        face_cx = fx + fw // 2
        face_cy = fy + fh // 2
        
        photo_center_x = face_cx
        # 保持适当下移
        photo_center_y = face_cy + int(fh * 0.08)
        
        # 计算左上角坐标
        x = photo_center_x - photo_width // 2
        y = photo_center_y - photo_height // 2
        
        # 确保边界在图片范围内
        x = max(0, min(x, img_width - photo_width))
        y = max(0, min(y, img_height - photo_height))
        w = min(photo_width, img_width - x)
        h = min(photo_height, img_height - y)
        
        print(f"[DEBUG] 调整因子: {adjust_factor:.2f}, 人脸比例: h={face_ratio_h:.2f}, w={face_ratio_w:.2f}")
        
        return x, y, w, h
