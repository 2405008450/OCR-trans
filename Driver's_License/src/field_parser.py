"""字段解析模块 - 使用空间规则 + 正则规则 + 键值配对优化匹配"""

from typing import List, Optional, Tuple, Callable, Dict
import re
from src.models import TextBlock, LicenseField


class FieldParser:
    """
    字段解析器，从 OCR 结果中提取驾驶证标准字段
    
    使用三层匹配策略：
    1. 空间规则：根据位置关系筛选候选值（右侧同一行 > 右上方 > 右侧接近 > 下方）
    2. 正则规则：验证候选值是否符合字段格式特征
    3. 键值配对：综合空间距离和格式匹配度选择最佳候选
    """
    
    # 驾驶证标准字段标签（中文 -> 英文）
    FIELD_LABELS = {
        "姓名": "Name",
        "性别": "Sex",
        "国籍": "Nationality",
        "出生日期": "Date of Birth",
        "住址": "Address",
        "证号": "License No.",
        "准驾车型": "Class",
        "有效期起始日期": "Valid From",
        "有效起始日期": "Valid From",  # OCR 可能识别为这个（少了一个"期"字）
        "有效期限": "Valid Period",
        "发证机关": "Issuing Authority",
        "初次领证日期": "Date of First Issue",
        "档案编号": "File No.",
        "记录": "Record"
    }
    
    # 字段正则规则：用于精确匹配字段值的格式特征
    FIELD_REGEX_PATTERNS = {
        "姓名": re.compile(r'^[\u4e00-\u9fa5]{2,6}$'),
        "性别": re.compile(r'^(男|女)$'),
        "国籍": re.compile(r'^[\u4e00-\u9fa5]{2,10}$'),
        "出生日期": re.compile(r'\d{4}[年]\d{1,2}[月]\d{1,2}|\d{4}[/\-\.]\d{2}[/\-\.]\d{2}'),
        "住址": re.compile(r'.*(省|市|区|县|镇|村|路|号|街|道).*'),
        "证号": re.compile(r'^\d{17}[\dXx]$'),
        "准驾车型": re.compile(r'^([A-E][1-3]){1,3}[DEFMNP]?$|^[DEFMNP]$'),
        "有效期起始日期": re.compile(r'\d{4}[年]\d{1,2}[月]\d{1,2}|\d{4}[/\-\.]\d{2}[/\-\.]\d{2}'),
        "有效期限": re.compile(r'.*(\d{4}[年/\-\.]\d{1,2}[月/\-\.]\d{1,2}|年|长期).*'),
        "发证机关": re.compile(r'.*(公安|交警|车管|局|支队|大队).*'),
        "初次领证日期": re.compile(r'\d{4}[年]\d{1,2}[月]\d{1,2}|\d{4}[/\-\.]\d{2}[/\-\.]\d{2}'),
        "档案编号": re.compile(r'^\d{6,20}$'),
        "记录": None
    }
    
    # 空间规则优先级定义（数值越小优先级越高）
    SPATIAL_PRIORITY = {
        "right_same_line": 1,      # 右侧同一行（最高优先级）
        "right_above": 2,           # 右上方
        "right_near_line": 3,       # 右侧接近同一行
        "below_same_column": 4,     # 下方同一列
        "below_near_column": 5      # 下方附近列
    }
    
    # 英文标签（用于识别驾驶证上已有的英文标签）
    ENGLISH_LABELS = {
        "Name": "姓名",
        "Sex": "性别",
        "Nationality": "国籍",
        "Date of Birth": "出生日期",
        "Address": "住址",
        "License No.": "证号",
        "Class": "准驾车型",
        "Valid From": "有效期起始日期",
        "Valid For": "有效期限",  # 旧版驾驶证使用 Valid For
        "Valid Period": "有效期限",
        "Issuing Authority": "发证机关",
        "Date of First Issue": "初次领证日期",
        "File No.": "档案编号",
        "Record": "记录"
    }
    
    def __init__(self):
        """初始化字段解析器"""
        pass
    
    # ==================== 正则验证方法 ====================
    
    @staticmethod
    def _regex_validate(field_name: str, value: str) -> bool:
        """
        使用正则规则验证字段值格式
        
        Args:
            field_name: 字段名称
            value: 字段值
            
        Returns:
            True表示格式正确
        """
        value = value.strip()
        if not value:
            return False
        pattern = FieldParser.FIELD_REGEX_PATTERNS.get(field_name)
        if pattern is None:
            return True  # 没有正则规则的字段默认通过
        return pattern.search(value) is not None
    
    @staticmethod
    def _validate_name(value: str) -> bool:
        """验证姓名格式：2-6个中文字符"""
        value = value.strip()
        if not value:
            return False
        if FieldParser._regex_validate("姓名", value):
            return True
        chinese_chars = sum(1 for c in value if '\u4e00' <= c <= '\u9fff')
        return 2 <= len(value) <= 6 and chinese_chars >= 2
    
    @staticmethod
    def _validate_sex(value: str) -> bool:
        """验证性别格式：男/女"""
        return FieldParser._regex_validate("性别", value)
    
    @staticmethod
    def _validate_nationality(value: str) -> bool:
        """验证国籍格式：中文国家名"""
        value = value.strip()
        if not value:
            return False
        # 排除包含地址关键词的文本
        address_keywords = ['省', '市', '区', '县', '镇', '村', '路', '号', '街', '道', '楼', '室', '房']
        if any(kw in value for kw in address_keywords):
            return False
        if FieldParser._regex_validate("国籍", value):
            return True
        chinese_chars = sum(1 for c in value if '\u4e00' <= c <= '\u9fff')
        return 2 <= len(value) <= 20 and chinese_chars >= 2
    
    @staticmethod
    def _validate_date(value: str) -> bool:
        """验证日期格式"""
        return FieldParser._regex_validate("出生日期", value)
    
    @staticmethod
    def _validate_address(value: str) -> bool:
        """验证住址格式：包含地名关键词"""
        value = value.strip()
        if not value:
            return False
        if FieldParser._regex_validate("住址", value):
            return True
        chinese_chars = sum(1 for c in value if '\u4e00' <= c <= '\u9fff')
        return len(value) >= 5 and chinese_chars >= 3
    
    @staticmethod
    def _validate_license_no(value: str) -> bool:
        """验证证号格式：18位"""
        return FieldParser._regex_validate("证号", value)
    
    @staticmethod
    def _validate_class(value: str) -> bool:
        """验证准驾车型格式"""
        value = value.strip()
        if not value:
            return False
        # 排除明显的英文标签
        if any(kw in value.lower() for kw in ['date', 'issue', 'find', 'lssue', 'dute', 'ther', 'lice']):
            return False
        if len(value) > 20:
            return False
        # 排除日期格式（有效期限的值）
        if re.search(r'\d{4}[-/年]\d{1,2}[-/月]\d{1,2}', value):
            return False
        # 排除包含"至"、"长期"等有效期限关键词
        if any(kw in value for kw in ['至', '长期', '到']):
            return False
        if FieldParser._regex_validate("准驾车型", value):
            return True
        # 兼容旧逻辑：检查是否包含有效车型代号
        valid_classes = ["A1", "A2", "A3", "B1", "B2", "C1", "C2", "C3", "C4", "C5",
                         "D", "E", "F", "M", "N", "P"]
        for cls in valid_classes:
            if len(cls) == 1:
                pattern = r'(?<![A-Za-z])' + re.escape(cls) + r'(?![a-z])'
                if re.search(pattern, value):
                    return True
            else:
                if cls in value:
                    return True
        return False
    
    @staticmethod
    def _validate_valid_period(value: str) -> bool:
        """
        验证有效期限格式
        
        有效期限的常见格式：
        - 旧版：X年（如 6年、10年）
        - 新版：YYYY-MM-DD至YYYY-MM-DD 或 YYYY年MM月DD日至YYYY年MM月DD日
        - 长期
        
        注意：纯日期格式（如 2012-06-13）不是有效期限，而是有效起始日期
        """
        value = value.strip()
        if not value:
            return False
        
        # 优先匹配 "X年" 格式（旧版有效期限）
        if re.match(r'^\d{1,2}年$', value):
            return True
        
        # 匹配 "长期"
        if value == "长期":
            return True
        
        # 匹配日期范围格式（新版有效期限）
        if "至" in value or "到" in value:
            return True
        
        # 匹配包含两个日期的格式
        date_pattern = r'\d{4}[-/年\.]\d{1,2}[-/月\.]\d{1,2}'
        dates = re.findall(date_pattern, value)
        if len(dates) >= 2:
            return True
        
        # 纯日期格式不是有效期限（这是有效起始日期）
        if re.match(r'^\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2}$', value):
            return False
        
        # 包含"有效期限"标签的混合文本不是有效值
        if "有效期限" in value or "有效起始" in value:
            return False
        
        return False
    
    @staticmethod
    def _validate_issuing_authority(value: str) -> bool:
        """验证发证机关格式"""
        value = value.strip()
        if not value:
            return False
        if FieldParser._regex_validate("发证机关", value):
            chinese_chars = sum(1 for c in value if '\u4e00' <= c <= '\u9fff')
            return len(value) >= 5 and chinese_chars >= 3
        return False
    
    @staticmethod
    def _validate_file_no(value: str) -> bool:
        """验证档案编号格式"""
        value = value.strip()
        if not value:
            return False
        if FieldParser._regex_validate("档案编号", value):
            return True
        if len(value) < 6 or len(value) > 20:
            return False
        digits = sum(1 for c in value if c.isdigit())
        if digits < 6:
            return False
        if digits / len(value) < 0.7:
            return False
        exclude_keywords = ['年', '月', '日', '请', '提交', '身体', '条件', '证明', '之后', '每年']
        if any(kw in value for kw in exclude_keywords):
            return False
        return True
    
    # 字段验证规则映射
    FIELD_VALIDATORS = {
        "姓名": lambda v: FieldParser._validate_name(v),
        "性别": lambda v: FieldParser._validate_sex(v),
        "国籍": lambda v: FieldParser._validate_nationality(v),
        "出生日期": lambda v: FieldParser._validate_date(v),
        "住址": lambda v: FieldParser._validate_address(v),
        "证号": lambda v: FieldParser._validate_license_no(v),
        "准驾车型": lambda v: FieldParser._validate_class(v),
        "有效期起始日期": lambda v: FieldParser._validate_date(v),
        "有效起始日期": lambda v: FieldParser._validate_date(v),  # OCR 变体
        "有效期限": lambda v: FieldParser._validate_valid_period(v),
        "发证机关": lambda v: FieldParser._validate_issuing_authority(v),
        "初次领证日期": lambda v: FieldParser._validate_date(v),
        "档案编号": lambda v: FieldParser._validate_file_no(v),
        "记录": lambda v: True
    }
    
    # ==================== 空间规则方法 ====================
    
    @staticmethod
    def _calculate_spatial_relation(
        label_block: TextBlock,
        candidate_block: TextBlock
    ) -> Tuple[Optional[str], float, float]:
        """
        计算候选块相对于标签块的空间关系和得分
        
        空间规则优先级：
        1. right_same_line: 右侧同一行（垂直距离 < 0.5倍标签高度）
        2. right_above: 右上方（垂直距离 < 3倍标签高度）
        3. right_near_line: 右侧接近同一行（垂直距离 < 1.5倍标签高度）
        4. below_same_column: 下方同一列（水平距离 < 标签宽度）
        5. below_near_column: 下方附近列（水平距离 < 1.5倍宽度和）
        
        Args:
            label_block: 标签块
            candidate_block: 候选值块
            
        Returns:
            (空间关系类型, 优先级得分, 距离得分)
        """
        label_x, label_y, label_w, label_h = label_block.get_rect()
        label_cx, label_cy = label_block.get_center()
        
        cand_x, cand_y, cand_w, cand_h = candidate_block.get_rect()
        cand_cx, cand_cy = candidate_block.get_center()
        
        h_dist = cand_cx - label_cx  # 水平距离（正值=右侧）
        v_dist = abs(cand_cy - label_cy)  # 垂直距离（绝对值）
        
        spatial_type = None
        priority = float('inf')
        
        # 右侧候选
        if cand_cx > label_cx:
            if v_dist < label_h * 0.5:
                spatial_type = "right_same_line"
            elif cand_cy < label_cy and v_dist < label_h * 3:
                spatial_type = "right_above"
            elif v_dist < label_h * 1.5:
                spatial_type = "right_near_line"
        
        # 下方候选（仅当未匹配到右侧规则时）
        if spatial_type is None and cand_cy > label_cy:
            h_dist_abs = abs(cand_cx - label_cx)
            max_h_offset = (label_w + cand_w) * 1.5
            if h_dist_abs < label_w and cand_y - label_y < label_h * 3:
                spatial_type = "below_same_column"
            elif h_dist_abs < max_h_offset and cand_y - label_y < label_h * 3:
                spatial_type = "below_near_column"
        
        if spatial_type:
            priority = FieldParser.SPATIAL_PRIORITY[spatial_type]
        
        distance = (h_dist ** 2 + v_dist ** 2) ** 0.5
        return spatial_type, priority, distance
    
    # ==================== 页面检测方法 ====================
    
    def _detect_page_types(self, text_blocks: List[TextBlock], page_boundary: int) -> dict:
        """检测图片中包含的页面类型"""
        page_types = {}
        for block in text_blocks:
            text = block.text
            center_x, center_y = block.get_center()
            page_side = 'left' if center_x < page_boundary else 'right'
            
            if "Driving License of the People's Republic of China" in text and "Duplicate" not in text:
                page_types[page_side] = 'main'
            elif "Duplicate of Driving License" in text or "副页" in text:
                page_types[page_side] = 'duplicate'
            elif ("Legend for Class of Vehicles" in text or 
                  "准驾车型代号规定" in text or
                  ("准驾车型" in text and "代号" in text and "规定" in text)):
                page_types[page_side] = 'legend'
        return page_types
    
    # ==================== 辅助方法 ====================
    
    def _sort_text_blocks(self, text_blocks: List[TextBlock]) -> List[TextBlock]:
        """按位置排序文字块（从上到下，从左到右）"""
        return sorted(text_blocks, key=lambda b: (b.get_center()[1], b.get_center()[0]))
    
    def _match_field_label(self, text: str) -> Optional[str]:
        """匹配中文或英文字段标签"""
        # 特殊处理：如果文本以"记录"开头，优先识别为记录字段
        # 这是为了处理记录延伸页中"记录"标签和内容合并的情况
        if text.strip().startswith("记录"):
            return "记录"
        
        # 特殊处理：如果文本以"有效期限"开头，识别为有效期限字段
        # 这是为了处理有效期限标签和值合并的情况，如 "有效期限2015-07-10至2021-07-10"
        if text.strip().startswith("有效期限"):
            return "有效期限"
        
        # 排除包含这些关键词的长文本（通常是记录内容或说明文字）
        exclude_keywords = ['请于', '办理', '提交', '身体条件', '证明', '审验', '换证', '降级', '变更', '自', '至']
        if any(kw in text for kw in exclude_keywords) and len(text) > 10:
            return None
        
        # 排除以"自"开头的文本（通常是记录内容，如"自2012年06月05日至..."）
        if text.strip().startswith("自"):
            return None
        
        for label in self.FIELD_LABELS:
            if label in text:
                return label
        text_stripped = text.strip()
        for eng_label, chi_label in self.ENGLISH_LABELS.items():
            if eng_label.lower() == text_stripped.lower() or eng_label.lower() in text_stripped.lower():
                return chi_label
        return None
    
    def _match_english_label(self, text: str) -> Optional[str]:
        """匹配英文字段标签"""
        text = text.strip()
        for english_label in self.ENGLISH_LABELS:
            if english_label.lower() in text.lower():
                return english_label
        return None
    
    def _is_english_label(self, text: str) -> bool:
        """检查文本是否是英文标签"""
        text = text.strip()
        for english_label in self.ENGLISH_LABELS:
            if english_label.lower() == text.lower():
                return True
        return False
    
    def _is_title_text(self, text: str) -> bool:
        """检查是否是标题文本"""
        title_keywords = ["中华人民共和国", "Driving License", "People's Republic", "副页"]
        return any(kw in text for kw in title_keywords)
    
    def _is_same_page(self, label_x: int, candidate_x: int, page_boundary: int, field_name: str) -> bool:
        """检查标签和候选值是否在同一页"""
        if field_name == "档案编号":
            return True
        if label_x < page_boundary and candidate_x > page_boundary + 50:
            return False
        if label_x > page_boundary + 50 and candidate_x < page_boundary:
            return False
        return True
    
    def _extract_value_from_combined_text(self, text: str, field_name: str) -> Optional[str]:
        """从合并的文本中提取字段值（当标签和值在同一个文字块中时）"""
        english_label = self.FIELD_LABELS.get(field_name)
        if english_label and text.strip().lower() == english_label.lower():
            return None
        value = text.replace(field_name, "").strip()
        if not value or len(value) < 2:
            return None
        
        # 记录字段特殊处理：不需要验证格式，直接返回
        if field_name == "记录":
            return value
        
        validator = self.FIELD_VALIDATORS.get(field_name)
        if validator and validator(value):
            return value
        
        # 验证失败时不返回值（让 _find_field_value 去找正确的值）
        return None
    
    def _extract_date_from_mixed_text(self, text: str) -> Optional[str]:
        """
        从混合文本中提取日期值
        
        处理 OCR 将日期和其他标签合并识别的情况，如 "2012-06-13有效期限"
        
        Args:
            text: 可能包含日期和其他文本的字符串
            
        Returns:
            提取的日期字符串，如果没有找到则返回 None
        """
        import re
        
        # 匹配日期格式：YYYY-MM-DD 或 YYYY/MM/DD 或 YYYY.MM.DD 或 YYYY年MM月DD日
        date_patterns = [
            r'(\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2})',  # 2012-06-13, 2012/06/13, 2012.06.13
            r'(\d{4}年\d{1,2}月\d{1,2}日?)',  # 2012年06月13日
        ]
        
        for pattern in date_patterns:
            match = re.search(pattern, text)
            if match:
                return match.group(1)
        
        return None
    
    def _extract_period_from_text(self, text: str) -> Optional[str]:
        """
        从文本中提取有效期限值
        
        有效期限的格式：X年、长期、日期范围
        
        Args:
            text: 文本
            
        Returns:
            提取的有效期限值，如果没有找到则返回 None
        """
        import re
        
        text = text.strip()
        
        # 匹配 "X年" 格式
        match = re.match(r'^(\d{1,2}年)$', text)
        if match:
            return match.group(1)
        
        # 匹配 "长期"
        if text == "长期":
            return text
        
        # 匹配日期范围
        if "至" in text or "到" in text:
            return text
        
        return None
    
    def _find_english_label_below(self, sorted_blocks, label_index, label_block, field_name):
        """查找中文标签下方的英文标签"""
        label_x, label_y, label_w, label_h = label_block.get_rect()
        label_cx, label_cy = label_block.get_center()
        expected_english = self.FIELD_LABELS.get(field_name)
        if not expected_english:
            return None
        for i in range(label_index + 1, len(sorted_blocks)):
            block = sorted_blocks[i]
            bcx, bcy = block.get_center()
            by = block.get_rect()[1]
            if bcy > label_cy and abs(bcx - label_cx) < label_w:
                english_label = self._match_english_label(block.text)
                if english_label and english_label == expected_english:
                    return english_label
                if by - label_y > label_h * 3:
                    break
        return None
    
    # ==================== 核心匹配方法（优化版） ====================
    
    def _find_field_value(
        self,
        sorted_blocks: List[TextBlock],
        label_index: int,
        label_block: TextBlock,
        used_values: dict = None,
        page_boundary: int = 500
    ) -> Tuple[Optional[str], Optional[Tuple[int, int]], Optional[int]]:
        """
        使用空间规则 + 正则规则 + 键值配对查找字段值
        
        匹配策略：
        1. 空间筛选：根据位置关系筛选候选块
        2. 格式验证：使用正则规则验证候选值格式
        3. 综合评分：优先选择格式正确且空间优先级最高的候选
        
        Args:
            sorted_blocks: 排序后的文字块列表
            label_index: 标签块的索引
            label_block: 标签块
            used_values: 已使用的值块字典
            page_boundary: 左右页分界线
            
        Returns:
            (字段值, 位置坐标, 值块索引)
        """
        if used_values is None:
            used_values = {}
        
        field_name = self._match_field_label(label_block.text)
        if not field_name:
            return None, None, None
        
        # 记录字段使用特殊合并逻辑
        if field_name == "记录":
            merged_text, merged_pos = self._merge_record_blocks_v2(
                sorted_blocks, label_index, label_block
            )
            return merged_text, merged_pos, None
        
        label_cx, label_cy = label_block.get_center()
        validator = self.FIELD_VALIDATORS.get(field_name)
        
        # 搜索范围：向前10个块 + 向后所有块
        search_range = range(max(0, label_index - 10), len(sorted_blocks))
        
        # 候选列表
        candidates = []
        
        # 日期类字段列表
        date_fields = ["有效期起始日期", "有效起始日期", "出生日期", "初次领证日期"]
        is_date_field = field_name in date_fields
        
        # 有效期限字段需要特殊处理
        is_valid_period_field = field_name == "有效期限"
        
        for i in search_range:
            if i == label_index or i in used_values:
                continue
            
            block = sorted_blocks[i]
            bcx, bcy = block.get_center()
            block_text = block.text
            
            # 对于日期类字段，检查是否是混合文本（包含日期和其他标签）
            extracted_value = None
            if is_date_field:
                # 检查文本是否包含日期格式
                extracted_value = self._extract_date_from_mixed_text(block_text)
                if extracted_value:
                    # 如果提取到日期，使用提取的日期作为候选值
                    block_text = extracted_value
            
            # 对于有效期限字段，检查是否是有效期限格式（X年、长期等）
            if is_valid_period_field:
                period_value = self._extract_period_from_text(block_text)
                if period_value:
                    extracted_value = period_value
                    block_text = period_value
                else:
                    # 如果不是有效期限格式，跳过
                    # 包括混合文本如 "2012-06-13有效期限"（这是日期+标签，不是有效期限值）
                    # 也包括纯日期格式（这是有效起始日期，不是有效期限）
                    if "有效期限" in block.text or "有效起始" in block.text:
                        continue
                    # 纯日期格式不是有效期限值
                    if re.match(r'^\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2}$', block_text.strip()):
                        continue
            
            # 跳过纯标签（但如果已经提取到值，不跳过）
            if not extracted_value:
                if self._match_field_label(block.text):
                    continue
                if self._is_english_label(block.text):
                    continue
            if self._is_title_text(block.text):
                continue
            
            # 防止跨页匹配
            if not self._is_same_page(label_cx, bcx, page_boundary, field_name):
                continue
            
            # 计算空间关系
            spatial_type, priority, distance = self._calculate_spatial_relation(
                label_block, block
            )
            
            if spatial_type is None:
                continue
            
            # 正则验证（使用可能提取的值）
            is_valid = validator(block_text) if validator else True
            
            candidates.append({
                'block': block,
                'index': i,
                'spatial_type': spatial_type,
                'priority': priority,
                'distance': distance,
                'is_valid': is_valid,
                'extracted_value': extracted_value  # 保存提取的值
            })
        
        if not candidates:
            return None, None, None
        
        # 优先选择格式正确的候选
        valid_cands = [c for c in candidates if c['is_valid']]
        invalid_cands = [c for c in candidates if not c['is_valid']]
        
        if valid_cands:
            best = min(valid_cands, key=lambda c: (c['priority'], c['distance']))
            
            # 住址字段特殊处理：查找下方的额外行（如"405号"）
            if field_name == "住址":
                merged_value, merged_pos = self._merge_address_lines(
                    sorted_blocks, best['index'], best['block'], used_values
                )
                return merged_value, merged_pos, best['index']
            
            # 如果有提取的日期值，使用提取的值
            if best.get('extracted_value'):
                return best['extracted_value'], best['block'].get_center(), best['index']
            
            return best['block'].text, best['block'].get_center(), best['index']
        
        # 降级：使用格式不正确的候选
        if invalid_cands:
            # 对关键字段额外过滤
            if field_name in ["姓名", "性别", "国籍", "准驾车型", "初次领证日期", "出生日期", "有效期起始日期"]:
                filtered = []
                for c in invalid_cands:
                    text = c.get('extracted_value') or c['block'].text.strip()
                    if field_name == "姓名":
                        # 姓名不应该太长（超过6个字符）
                        if len(text) > 6:
                            continue
                        # 姓名不应该是单个字符
                        if len(text) == 1:
                            continue
                        # 姓名不应该是性别
                        if text in ["男", "女"]:
                            continue
                        # 姓名不应该包含地址关键词
                        address_keywords = ['省', '市', '区', '县', '镇', '村', '路', '号', '街', '道', '楼', '室', '房']
                        if any(kw in text for kw in address_keywords):
                            continue
                    if field_name == "性别" and (len(text) > 2 or any(ch.isalpha() and ord(ch) < 128 for ch in text)):
                        continue
                    if field_name == "国籍":
                        # 国籍不应该包含地址关键词
                        address_keywords = ['省', '市', '区', '县', '镇', '村', '路', '号', '街', '道', '楼', '室', '房']
                        if any(kw in text for kw in address_keywords):
                            continue
                    if field_name == "准驾车型":
                        # 准驾车型不应该包含日期格式
                        if re.search(r'\d{4}[-/年]\d{1,2}[-/月]\d{1,2}', text):
                            continue
                        # 准驾车型不应该包含有效期限关键词
                        if any(kw in text for kw in ['至', '长期', '到']):
                            continue
                        # 准驾车型不应该包含印章关键词
                        if any(kw in text for kw in ['公安', '交警', '支队', '大队', '局']):
                            continue
                    if field_name in ["初次领证日期", "出生日期", "有效期起始日期"]:
                        # 日期字段不应该是准驾车型格式
                        if re.match(r'^[A-E][1-3]?[DEFMNP]?$', text):
                            continue
                        # 日期字段不应该包含印章关键词
                        if any(kw in text for kw in ['公安', '交警', '支队', '大队', '局']):
                            continue
                        # 日期字段应该包含数字
                        if not any(c.isdigit() for c in text):
                            continue
                    filtered.append(c)
                if filtered:
                    invalid_cands = filtered
                else:
                    return None, None, None
            
            best = min(invalid_cands, key=lambda c: (c['priority'], c['distance']))
            # 如果有提取的日期值，使用提取的值
            if best.get('extracted_value'):
                print(f"  [警告] 字段 {field_name} 从混合文本提取: {best['extracted_value']}")
                return best['extracted_value'], best['block'].get_center(), best['index']
            print(f"  [警告] 字段 {field_name} 选择格式不符: {best['block'].text}")
            return best['block'].text, best['block'].get_center(), best['index']
        
        return None, None, None
    
    # ==================== 主解析方法 ====================
    
    def parse_fields(self, text_blocks: List[TextBlock], has_duplicate: bool = False, photo_boundary: int = None) -> List[LicenseField]:
        """
        从文字块列表中解析驾驶证字段
        
        Args:
            text_blocks: OCR 识别的文字块列表
            has_duplicate: 是否为副页
            photo_boundary: 相片右边界的X坐标
            
        Returns:
            驾驶证字段列表
        """
        fields = []
        used_blocks = set()
        used_values = {}
        
        page_boundary = photo_boundary if photo_boundary is not None else 500
        
        page_types = self._detect_page_types(text_blocks, page_boundary)
        
        filtered_blocks = text_blocks
        if page_types.get('right') == 'legend':
            filtered_blocks = [b for b in text_blocks if b.get_center()[0] < page_boundary]
        
        sorted_blocks = self._sort_text_blocks(filtered_blocks)
        
        i = 0
        # 记录已经找到值的字段名
        found_fields = set()
        
        # 字段名变体映射（用于处理 OCR 识别变体）
        field_name_variants = {
            "有效起始日期": "有效期起始日期",  # OCR 变体 -> 标准名称
            "有效期起始日期": "有效起始日期",  # 标准名称 -> OCR 变体
        }
        
        while i < len(sorted_blocks):
            if i in used_blocks:
                i += 1
                continue
            
            block = sorted_blocks[i]
            field_name = self._match_field_label(block.text)
            
            if field_name:
                # 检查是否是英文标签，且该字段已经找到值
                is_english_label = self._is_english_label(block.text)
                if is_english_label:
                    # 检查字段名或其变体是否已经找到值
                    variant = field_name_variants.get(field_name)
                    if field_name in found_fields or (variant and variant in found_fields):
                        # 跳过已经处理过的字段的英文标签
                        i += 1
                        continue
                
                # 记录字段特殊处理
                # 检测是否是记录延伸页：文字块数量少且不是副页
                is_extension_page = len(sorted_blocks) <= 5
                
                if field_name == "记录":
                    if is_extension_page:
                        # 可能是记录延伸页，合并所有文字块的内容
                        all_texts = []
                        for j, b in enumerate(sorted_blocks):
                            text = b.text.strip()
                            # 去掉"记录"标签
                            if text.startswith("记录"):
                                text = text[2:].strip()
                            if text:
                                all_texts.append(text)
                                print(f"  [记录延伸页] 文字块 {j+1}: {text}")
                        
                        if all_texts:
                            merged_value = "".join(all_texts)
                            fields.append(LicenseField(
                                field_name="记录",
                                field_value=merged_value,
                                position=block.get_center()
                            ))
                            print(f"[OK] 记录（延伸页）: {merged_value}")
                            # 标记所有块为已使用
                            for j in range(len(sorted_blocks)):
                                used_blocks.add(j)
                            break  # 跳出循环
                    elif not has_duplicate:
                        # 不是副页也不是延伸页，跳过记录字段
                        i += 1
                        continue
                
                # 检查标签和值合并的情况
                combined_value = self._extract_value_from_combined_text(block.text, field_name)
                if combined_value:
                    fields.append(LicenseField(
                        field_name=field_name,
                        field_value=combined_value,
                        position=block.get_center()
                    ))
                    used_blocks.add(i)
                    i += 1
                    continue
                
                # 查找英文标签
                english_label = self._find_english_label_below(sorted_blocks, i, block, field_name)
                
                # 使用优化的空间+正则匹配查找字段值
                field_value, value_position, value_index = self._find_field_value(
                    sorted_blocks, i, block, used_values, page_boundary
                )
                
                if field_value and not self._is_english_label(field_value):
                    fields.append(LicenseField(
                        field_name=field_name,
                        field_value=field_value,
                        position=value_position
                    ))
                    used_blocks.add(i)
                    if value_index is not None:
                        used_values[value_index] = field_name
                    # 记录已找到值的字段
                    found_fields.add(field_name)
                    
                    print(f"[OK] {field_name}: {field_value}")
                else:
                    print(f"[X] 未找到字段值: {field_name}")
                
            i += 1
        
        fields = self._deduplicate_fields(fields)
        print(f"总共识别到 {len(fields)} 个字段")
        return fields
    
    # ==================== 去重方法 ====================
    
    def _deduplicate_fields(self, fields: List[LicenseField]) -> List[LicenseField]:
        """去重字段：优先保留格式验证通过的字段"""
        field_groups = {}
        for field in fields:
            if field.field_name not in field_groups:
                field_groups[field.field_name] = []
            field_groups[field.field_name].append(field)
        
        deduplicated = []
        for field_name, group in field_groups.items():
            if len(group) == 1:
                deduplicated.append(group[0])
            else:
                validator = self.FIELD_VALIDATORS.get(field_name)
                if validator:
                    valid_fields = [f for f in group if validator(f.field_value)]
                    if valid_fields:
                        best = min(valid_fields, key=lambda f: f.position[0] if f.position else float('inf'))
                        deduplicated.append(best)
                        print(f"去重: {field_name} 有 {len(group)} 个（{len(valid_fields)}个格式正确），保留: {best.field_value}")
                    else:
                        best = min(group, key=lambda f: f.position[0] if f.position else float('inf'))
                        deduplicated.append(best)
                        print(f"去重: {field_name} 有 {len(group)} 个（都格式不正确），保留最左侧: {best.field_value}")
                else:
                    best = min(group, key=lambda f: f.position[0] if f.position else float('inf'))
                    deduplicated.append(best)
        return deduplicated
    
    # ==================== 记录字段合并方法 ====================
    
    def _merge_address_lines(
        self,
        sorted_blocks: List[TextBlock],
        main_index: int,
        main_block: TextBlock,
        used_values: dict
    ) -> Tuple[str, Tuple[int, int]]:
        """
        合并住址字段的多行文字块
        
        有些驾驶证的住址会分成多行，例如：
        第一行：北京市朝阳区广渠路九龙山家园1楼2门
        第二行：405号
        
        Args:
            sorted_blocks: 排序后的文字块列表
            main_index: 主住址块的索引
            main_block: 主住址块
            used_values: 已使用的值块字典
            
        Returns:
            (合并后的住址, 位置坐标)
        """
        main_x, main_y, main_w, main_h = main_block.get_rect()
        main_cx, main_cy = main_block.get_center()
        main_text = main_block.text.strip()
        
        # 查找下方的额外行
        extra_lines = []
        
        # 向后搜索（下方）
        for i in range(main_index + 1, len(sorted_blocks)):
            if i in used_values:
                continue
            
            block = sorted_blocks[i]
            bcx, bcy = block.get_center()
            bx, by, bw, bh = block.get_rect()
            text = block.text.strip()
            
            # 跳过标签
            if self._match_field_label(text):
                # 不要 break，继续查找下一个块
                continue
            if self._is_english_label(text):
                continue
            
            # 垂直距离检查：在主块下方，且不超过3倍主块高度
            y_diff = bcy - main_cy
            if y_diff < 0:
                continue
            if y_diff > main_h * 3:
                break
            
            # 水平位置检查：允许在主块左侧或下方
            # 有些住址的第二行会缩进到左侧
            # 只要不是太远右侧即可（不超过主块右边界太多）
            if bcx > main_cx + main_w * 2:
                continue
            
            # 检查是否包含地址关键词
            address_keywords = ['号', '室', '房', '楼', '层', '单元', '栋', '幢']
            if any(kw in text for kw in address_keywords):
                extra_lines.append((block, bcy, i))
        
        # 如果没有找到额外行，返回原始值
        if not extra_lines:
            return main_text, main_block.get_center()
        
        # 按 Y 坐标排序
        extra_lines.sort(key=lambda x: x[1])
        
        # 合并文本
        merged_text = main_text
        for block, _, idx in extra_lines:
            merged_text += block.text.strip()
            # 标记为已使用
            used_values[idx] = "住址"
        
        return merged_text, main_block.get_center()
    
    # ==================== 记录字段合并方法 ====================
    
    def _merge_record_blocks_v2(
        self,
        sorted_blocks: List[TextBlock],
        label_index: int,
        label_block: TextBlock
    ) -> Tuple[str, Tuple[int, int]]:
        """合并"记录"字段的多个文字块"""
        label_x, label_y, label_w, label_h = label_block.get_rect()
        label_cx, label_cy = label_block.get_center()
        
        # 检查标签文本是否包含额外内容
        label_text = label_block.text.strip()
        extra_content = ""
        if "记录" in label_text and len(label_text) > 2:
            extra_content = label_text.replace("记录", "").strip()
            if extra_content:
                pass  # 标签包含额外内容
        
        record_blocks = []
        
        exclude_labels = {
            '住址', '姓名', '性别', '国籍', '出生日期', '初次领证日期',
            '准驾车型', '有效期限', '证号', '档案编号', '发证机关',
            'Address', 'Name', 'Sex', 'Nationality', 'Date of Birth',
            'Date of First Issue', 'Class', 'Valid Period', 'License No',
            'File No', 'Issuing Authority', 'Valid From'
        }
        
        seal_keywords = [
            'Traffic', 'Police', 'Detachment', 'Public', 'Security', 'Bureau',
            'Province', 'City', 'Long-term', 'Room', 'Building', 'Family', 'Courtyard',
            '交警', '公安', '支队', '大队', '分局', '派出所', '长期'
        ]
        
        # 水平范围限制：左边界 = 记录标签的左边界(label_x)
        # 右边界 = 第一行有效文字块的右边界
        left_bound = label_x
        right_bound = None  # 稍后由第一个有效文字块确定
        
        # 先找第一个有效的记录文字块，确定右边界
        for i in range(label_index + 1, len(sorted_blocks)):
            block = sorted_blocks[i]
            bcx, bcy = block.get_center()
            bx, by, bw, bh = block.get_rect()
            text = block.text.strip()
            
            if text in exclude_labels or self._is_english_label(text):
                continue
            if any(kw in text for kw in seal_keywords):
                continue
            
            is_right_side = bcx > label_x
            y_diff = bcy - label_cy
            in_y_range = -label_h * 2 <= y_diff <= label_h * 15
            
            if is_right_side and in_y_range:
                chinese_chars = sum(1 for c in text if '\u4e00' <= c <= '\u9fff')
                if chinese_chars >= 2:
                    # 用第一个有效文字块的右边界作为水平范围上限
                    right_bound = bx + bw
                    break
        
        if right_bound is None:
            # 没找到有效文字块，用标签宽度的3倍作为默认右边界
            right_bound = label_x + label_w * 3
        
        print(f"  [记录] 水平收集范围: [{left_bound}, {right_bound}]")
        
        # 向后搜索（在水平范围内）
        for i in range(label_index + 1, len(sorted_blocks)):
            block = sorted_blocks[i]
            bcx, bcy = block.get_center()
            bx, by, bw, bh = block.get_rect()
            text = block.text.strip()
            
            if text in exclude_labels:
                continue
            if self._is_english_label(text):
                continue
            if any(kw in text for kw in seal_keywords):
                continue
            
            # 水平范围检查：文字块的左边界不能超出右边界
            if bx > right_bound:
                continue
            
            is_right_side = bcx > label_x
            y_diff = bcy - label_cy
            in_y_range = -label_h * 2 <= y_diff <= label_h * 15
            not_too_far_left = bcx > label_cx - 300
            
            if is_right_side and in_y_range and not_too_far_left:
                chinese_chars = sum(1 for c in text if '\u4e00' <= c <= '\u9fff')
                if chinese_chars >= 2:
                    record_blocks.append((block, bcy, bcx))
        
        # 向前搜索（在水平范围内）
        for i in range(label_index - 1, -1, -1):
            block = sorted_blocks[i]
            bcx, bcy = block.get_center()
            bx, by, bw, bh = block.get_rect()
            text = block.text.strip()
            
            if text in exclude_labels:
                continue
            if self._is_english_label(text):
                continue
            
            # 水平范围检查
            if bx > right_bound:
                continue
            
            is_right_side = bcx > label_x
            y_diff = label_cy - bcy
            in_y_range = 0 <= y_diff <= label_h * 2
            not_too_far_left = bcx > label_cx - 300
            
            if is_right_side and in_y_range and not_too_far_left:
                chinese_chars = sum(1 for c in text if '\u4e00' <= c <= '\u9fff')
                if chinese_chars >= 2:
                    record_blocks.append((block, bcy, bcx))
        
        if not record_blocks:
            if extra_content:
                return extra_content, (label_cx, label_cy)
            return "", (label_cx, label_cy)
        
        record_blocks.sort(key=lambda x: (x[1], x[2]))
        merged_text = "".join([block.text for block, _, _ in record_blocks])
        
        if extra_content:
            merged_text = extra_content + merged_text
        
        print(f"  [记录] 合并 {len(record_blocks)} 块: {merged_text}")
        
        return merged_text, record_blocks[0][0].get_center()
    
    def _merge_record_blocks(
        self,
        sorted_blocks: List[TextBlock],
        start_index: int,
        start_block: TextBlock
    ) -> Tuple[str, Tuple[int, int]]:
        """合并"记录"字段的多个文字块（旧版兼容）"""
        record_label_block = None
        record_label_index = None
        for i in range(start_index - 1, -1, -1):
            block = sorted_blocks[i]
            if "记录" in block.text:
                record_label_block = block
                record_label_index = i
                break
        
        if not record_label_block:
            record_label_block = start_block
            record_label_index = start_index
        
        return self._merge_record_blocks_v2(sorted_blocks, record_label_index, record_label_block)
