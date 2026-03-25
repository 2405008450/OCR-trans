"""OCR 识别结果修正模块"""

import re
from typing import List
from .models import TextBlock, LicenseField


class OCRCorrector:
    """OCR 识别结果修正器，修正常见的识别错误"""
    
    # 准驾车型的有效值（根据中国驾驶证标准）
    VALID_CLASSES = {
        'A1', 'A2', 'A3',
        'B1', 'B2',
        'C1', 'C2', 'C3', 'C4', 'C5',
        'D', 'E', 'F',
        'M', 'N', 'P'
    }
    
    def __init__(self):
        """初始化修正器"""
        pass
    
    def correct_text_blocks(self, text_blocks: List[TextBlock]) -> List[TextBlock]:
        """
        修正文字块中的常见错误
        
        注意：不修正字段标签（如 "Class", "Name" 等），只修正字段值
        
        Args:
            text_blocks: OCR 识别的文字块列表
            
        Returns:
            修正后的文字块列表
        """
        corrected_blocks = []
        
        # 字段标签列表（不应该被修正）
        field_labels = {
            "Name", "Sex", "Nationality", "Date of Birth", "Address",
            "License No.", "Class", "Valid From", "Valid Period",
            "Issuing Authority", "Date of First Issue", "File No.", "Record"
        }
        
        for block in text_blocks:
            # 检查是否是字段标签，如果是则不修正
            is_label = any(label.lower() == block.text.strip().lower() for label in field_labels)
            
            if is_label:
                # 字段标签不修正
                corrected_blocks.append(block)
            else:
                # 非标签文本才进行修正
                corrected_text = self._correct_text(block.text)
                
                if corrected_text != block.text:
                    print(f"[OCR修正] {block.text} -> {corrected_text}")
                    # 创建新的文字块
                    corrected_block = TextBlock(
                        text=corrected_text,
                        bounding_box=block.bounding_box,
                        confidence=block.confidence
                    )
                    corrected_blocks.append(corrected_block)
                else:
                    corrected_blocks.append(block)
        
        return corrected_blocks
    
    def correct_fields(self, fields: List[LicenseField]) -> List[LicenseField]:
        """
        修正字段值中的常见错误
        
        Args:
            fields: 字段列表
            
        Returns:
            修正后的字段列表
        """
        for field in fields:
            if field.field_name == "准驾车型":
                corrected_value = self._correct_vehicle_class(field.field_value)
                if corrected_value != field.field_value:
                    print(f"[字段修正] 准驾车型: {field.field_value} -> {corrected_value}")
                    field.field_value = corrected_value
        
        return fields
    
    def _correct_text(self, text: str) -> str:
        """
        修正文本中的常见错误
        
        Args:
            text: 原始文本
            
        Returns:
            修正后的文本
        """
        # 修正日期格式
        text = self._correct_date(text)
        
        # 修正准驾车型
        text = self._correct_vehicle_class(text)
        
        # 修正地址中的数字混淆
        text = self._correct_address_numbers(text)
        
        return text
    
    def _correct_address_numbers(self, text: str) -> str:
        """
        修正地址中的数字混淆问题
        
        常见错误：
        - 802房 -> 862房 (0被识别成6)
        - 号802 -> 号862
        
        修正规则：
        - 在"X号YYY房"模式中，如果房号中间有6，且前后是0或8，可能是0被误识别
        - 例如：862 可能是 802 的误识别（中间的6应该是0）
        
        Args:
            text: 原始文本
            
        Returns:
            修正后的文本
        """
        # 检查是否包含地址特征
        if not any(keyword in text for keyword in ['号', '房', '路', '街', '区', '省', '市', '村', '镇']):
            return text
        
        original_text = text
        
        # 模式1：X号YYY房 - 修正房号中的数字混淆
        # 例如：39号862房 -> 39号802房
        def fix_room_number(match):
            prefix = match.group(1)  # 号
            number = match.group(2)  # 房号数字
            suffix = match.group(3)  # 房
            
            # 检查是否是 X6Y 模式（中间是6，可能是0的误识别）
            # 常见误识别：802 -> 862, 803 -> 863, 801 -> 861
            if len(number) == 3:
                first, middle, last = number[0], number[1], number[2]
                # 如果中间是6，且第一位是8，可能是0被误识别成6
                if middle == '6' and first == '8':
                    corrected_number = first + '0' + last
                    print(f"[地址房号修正] {number} -> {corrected_number}")
                    return prefix + corrected_number + suffix
            
            return match.group(0)
        
        # 匹配 "号XXX房" 模式
        text = re.sub(r'(号)(\d{3})(房)', fix_room_number, text)
        
        # 模式2：直接的 XXX房 模式（没有"号"前缀）
        def fix_standalone_room(match):
            number = match.group(1)
            suffix = match.group(2)
            
            if len(number) == 3:
                first, middle, last = number[0], number[1], number[2]
                if middle == '6' and first == '8':
                    corrected_number = first + '0' + last
                    print(f"[地址房号修正] {number} -> {corrected_number}")
                    return corrected_number + suffix
            
            return match.group(0)
        
        # 匹配 "XXX房" 模式（前面不是数字）
        text = re.sub(r'(?<!\d)(\d{3})(房)', fix_standalone_room, text)
        
        return text
    
    def _correct_date(self, text: str) -> str:
        """
        修正日期中的常见错误
        
        常见错误：
        - 1976-64-15 -> 1976-04-15 (6被识别成64)
        - 1996-05-30 (正确的不修改)
        
        Args:
            text: 原始文本
            
        Returns:
            修正后的文本
        """
        # 匹配日期格式 YYYY-MM-DD
        date_pattern = r'(\d{4})-(\d{1,3})-(\d{1,2})'
        
        def fix_date_match(match):
            year = match.group(1)
            month = match.group(2)
            day = match.group(3)
            
            # 修正月份
            month_int = int(month)
            if month_int > 12:
                # 常见错误：64 -> 04, 65 -> 05 等
                # 如果月份是6X格式，很可能是0X被误识别
                if 60 <= month_int <= 69:
                    month = f"0{month_int - 60}"
                # 如果月份是7X格式，可能是1X被误识别
                elif 70 <= month_int <= 79:
                    month = f"{month_int - 60}"
                # 其他情况，尝试取最后一位
                elif month_int >= 20:
                    last_digit = month_int % 10
                    if 1 <= last_digit <= 9:
                        month = f"0{last_digit}"
            
            # 确保月份是两位数
            month = month.zfill(2)
            
            # 修正日期
            day_int = int(day)
            if day_int > 31:
                # 如果日期大于31，可能是识别错误
                # 尝试取最后一位或两位
                if day_int < 100:
                    last_digit = day_int % 10
                    if 1 <= last_digit <= 31:
                        day = str(last_digit)
            
            # 确保日期是两位数
            day = day.zfill(2)
            
            return f"{year}-{month}-{day}"
        
        corrected = re.sub(date_pattern, fix_date_match, text)
        
        if corrected != text:
            print(f"[日期修正] {text} -> {corrected}")
        
        return corrected
    
    def _correct_vehicle_class(self, text: str) -> str:
        """
        修正准驾车型中的常见错误
        
        常见错误：
        - C1F -> CIF (数字1被识别成字母I)
        - B1D -> BID
        - A1 -> AI
        - cí -> C1 (小写c和特殊字符)
        
        Args:
            text: 原始文本
            
        Returns:
            修正后的文本
        """
        # 如果文本不像准驾车型，直接返回
        if len(text) > 10:
            return text
        
        # 如果包含中文字符，不是准驾车型，直接返回
        if any('\u4e00' <= c <= '\u9fff' for c in text):
            return text
        
        original_text = text
        
        # 特殊情况：完全错误的识别（如 "cí", "ci", "cI"）
        # 只处理非常特殊的情况：2-3个字符，看起来像C1的误识别
        if len(text) <= 3:
            text_lower = text.lower()
            # 检查是否是C1的常见误识别
            if text_lower in ['ci', 'cí', 'c1', 'cl', 'c|', 'cî', 'cì']:
                return 'C1'
            
            # 检查是否只包含字母和特殊符号（不包含中文）
            has_special = any(not c.isalnum() and c not in [' ', '-'] for c in text)
            has_chinese = any('\u4e00' <= c <= '\u9fff' for c in text)
            
            # 只有当包含特殊符号且不包含中文时，才可能是C1的误识别
            if has_special and not has_chinese and 'c' in text_lower:
                return 'C1'
        
        # 查找所有可能的准驾车型模式
        # 模式：字母 + 可能的数字/字母 + 可能的其他字母
        patterns = [
            (r'C([IO])F', r'C1F'),  # CIF -> C1F, COF -> C1F
            (r'C([IO])', r'C1'),     # CI -> C1, CO -> C1
            (r'B([IO])D', r'B1D'),   # BID -> B1D, BOD -> B1D
            (r'B([IO])', r'B1'),     # BI -> B1, BO -> B1
            (r'A([IO])', r'A1'),     # AI -> A1, AO -> A1
        ]
        
        for pattern, replacement in patterns:
            text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)
        
        # 处理组合准驾车型（如 B1D, C1F）
        # 分离并修正每个部分
        parts = []
        current = ""
        
        for char in text:
            if char.isalpha() and current and current[-1].isdigit():
                # 遇到新的字母且当前已有数字，说明是新的车型
                parts.append(current)
                current = char
            else:
                current += char
        
        if current:
            parts.append(current)
        
        # 修正每个部分
        corrected_parts = []
        for part in parts:
            # 替换常见的字母数字混淆
            part = part.replace('I', '1').replace('O', '0').replace('l', '1')
            
            # 检查是否是有效的准驾车型
            if part.upper() in self.VALID_CLASSES:
                corrected_parts.append(part.upper())
            else:
                # 尝试修正
                # 如果是两个字符且第一个是字母
                if len(part) == 2 and part[0].isalpha():
                    if part[1] in ['I', 'O', 'l']:
                        corrected = part[0].upper() + '1'
                        if corrected in self.VALID_CLASSES:
                            corrected_parts.append(corrected)
                        else:
                            corrected_parts.append(part)
                    else:
                        corrected_parts.append(part)
                else:
                    corrected_parts.append(part)
        
        result = ''.join(corrected_parts) if corrected_parts else text
        
        return result
