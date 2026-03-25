"""翻译服务模块"""

import re
from typing import List
from openai import OpenAI

from .models import LicenseField
from .exceptions import TranslationError


class TranslationService:
    """翻译服务，使用 DeepSeek API"""
    
    # 城市公安局固定译法映射表
    CITY_PSB_TRANSLATIONS = {
        '广州市公安局': 'Guangzhou Municipal Public Security Bureau',
        '深圳市公安局': 'Public Security Bureau of Shenzhen Municipality',
        '重庆市公安局': 'Chongqing Municipal Public Security Bureau',
        '成都市公安局': 'Chengdu Municipal Public Security Bureau',
        '上海市公安局': 'Shanghai Municipal Public Security Bureau',
        '北京市公安局': 'Beijing Municipal Public Security Bureau',
        '杭州市公安局': 'Hangzhou Municipal Public Security Bureau',
        '武汉市公安局': 'Wuhan Public Security Bureau',
        '南京市公安局': 'Nanjing Municipal Public Security Bureau',
        '宁波市公安局': 'Ningbo Municipal Public Security Bureau',
        '天津市公安局': 'Tianjin Municipal Public Security Bureau',
        '青岛市公安局': 'Qingdao Municipal Public Security Bureau',
        '无锡市公安局': 'Wuxi Municipal Public Security Bureau',
        '长沙市公安局': 'Changsha Public Security Bureau',
        '郑州市公安局': 'Zhengzhou Municipal Public Security Bureau',
        '福州市公安局': 'Fuzhou Municipal Public Security Bureau',
        '济南市公安局': 'Jinan Public Security Bureau',
        '合肥市公安局': 'Hefei Public Security Bureau',
        '佛山市公安局': 'Foshan Public Security Bureau',
        '苏州市公安局': 'Suzhou Municipal Public Security Bureau',
    }
    
    def __init__(self, api_key: str):
        """
        初始化翻译服务
        
        Args:
            api_key: DeepSeek API 密钥
        """
        import httpx
        # 创建完全禁用代理的 HTTP 客户端
        http_client = httpx.Client(
            proxy=None,
            mounts={"all://": httpx.HTTPTransport(proxy=None)}
        )
        self.client = OpenAI(
            api_key=api_key,
            base_url="https://api.deepseek.com",
            http_client=http_client
        )
    
    def translate_fields(self, fields: List[LicenseField]) -> List[LicenseField]:
        """
        翻译驾驶证字段值
        
        Args:
            fields: 驾驶证字段列表
            
        Returns:
            翻译后的字段列表
        """
        for field in fields:
            # 性别字段特殊处理：只能是 Male 或 Female
            if field.field_name == "性别":
                field.translated_value = self._translate_sex(field.field_value)
            # 发证机关字段特殊处理：使用固定格式翻译
            elif field.field_name == "发证机关":
                field.translated_value = self._translate_issuing_authority(field.field_value)
            else:
                # 无论是否有英文标签，都需要翻译字段值
                # translated_value 应该是字段值的翻译，而不是英文标签
                field.translated_value = self._translate_text(field.field_value)
        
        return fields
    
    def translate_text(self, text: str) -> str:
        """
        翻译单个文本（公共方法）
        
        Args:
            text: 要翻译的文本
            
        Returns:
            翻译后的文本
        """
        # 检查是否是发证机关格式（印章内容）
        if self._is_issuing_authority_format(text):
            return self._translate_issuing_authority(text)
        return self._translate_text(text)
    
    def _is_issuing_authority_format(self, text: str) -> bool:
        """
        判断文本是否是发证机关格式（印章内容）
        
        Args:
            text: 文本
            
        Returns:
            是否是发证机关格式
        """
        text = text.strip()
        # 检查是否包含发证机关的关键特征
        has_public_security = '公安局' in text
        has_traffic_police = '交通警察' in text or '交警' in text
        has_unit = '支队' in text or '大队' in text
        
        return has_public_security and has_traffic_police and has_unit
    
    def _translate_sex(self, text: str) -> str:
        """
        翻译性别字段（特殊处理）
        
        Args:
            text: 性别文本（男/女）
            
        Returns:
            翻译后的性别（Male/Female）
        """
        text = text.strip()
        
        # 标准映射
        if text == "男":
            return "Male"
        elif text == "女":
            return "Female"
        
        # 如果OCR识别错误，尝试模糊匹配
        # 包含"男"字的都认为是男性
        if "男" in text:
            return "Male"
        # 包含"女"字的都认为是女性
        if "女" in text:
            return "Female"
        
        # 如果是英文，尝试识别
        text_lower = text.lower()
        if "male" in text_lower and "female" not in text_lower:
            return "Male"
        if "female" in text_lower:
            return "Female"
        
        # 如果都不匹配，调用API翻译，但限制结果只能是Male或Female
        translated = self._call_deepseek_api(f"请将'{text}'翻译成英文性别（只能是Male或Female）")
        
        # 确保结果只能是Male或Female
        if "female" in translated.lower():
            return "Female"
        else:
            return "Male"  # 默认返回Male
    
    def _translate_issuing_authority(self, text: str) -> str:
        """
        翻译发证机关（印章内容）
        格式：XX省XX市公安局交通警察支队 -> Traffic Police Detachment of XX City Public Security Bureau, XX Province
        
        Args:
            text: 发证机关文本
            
        Returns:
            翻译后的发证机关
        """
        text = text.strip()
        
        # 尝试匹配 "XX省XX市公安局交通警察支队" 格式
        # 支持多种变体：支队、大队等
        pattern = re.compile(r'^(.+?省)?(.+?市公安局)交通警察(支队|大队)$')
        match = pattern.match(text)
        
        if match:
            province = match.group(1)  # 可能为 None
            city_psb = match.group(2)  # XX市公安局
            unit_type = match.group(3)
            
            # 去掉"省"后缀
            if province:
                province_name = province.rstrip('省')
            else:
                province_name = None
            
            # 翻译单位类型
            unit_translation = "Detachment" if unit_type == "支队" else "Brigade"
            
            # 检查是否有固定译法
            if city_psb in self.CITY_PSB_TRANSLATIONS:
                psb_translation = self.CITY_PSB_TRANSLATIONS[city_psb]
                if province_name:
                    translated_province = self._translate_place_name(province_name)
                    return f"Traffic Police {unit_translation} of {psb_translation}, {translated_province} Province"
                else:
                    return f"Traffic Police {unit_translation} of {psb_translation}"
            
            # 没有固定译法，使用原有逻辑
            city_name = city_psb.rstrip('公安局').rstrip('市')
            translated_city = self._translate_place_name(city_name)
            translated_province = self._translate_place_name(province_name) if province_name else None
            
            # 组装翻译结果
            if translated_province:
                return f"Traffic Police {unit_translation} of {translated_city} City Public Security Bureau, {translated_province} Province"
            else:
                return f"Traffic Police {unit_translation} of {translated_city} City Public Security Bureau"
        
        # 尝试匹配 "XX市公安局交通警察支队" 格式（无省份）
        pattern2 = re.compile(r'^(.+?市公安局)交通警察(支队|大队)$')
        match2 = pattern2.match(text)
        
        if match2:
            city_psb = match2.group(1)  # XX市公安局
            unit_type = match2.group(2)
            
            # 翻译单位类型
            unit_translation = "Detachment" if unit_type == "支队" else "Brigade"
            
            # 检查是否有固定译法
            if city_psb in self.CITY_PSB_TRANSLATIONS:
                psb_translation = self.CITY_PSB_TRANSLATIONS[city_psb]
                return f"Traffic Police {unit_translation} of {psb_translation}"
            
            # 没有固定译法，使用原有逻辑
            city_name = city_psb.rstrip('公安局').rstrip('市')
            translated_city = self._translate_place_name(city_name)
            return f"Traffic Police {unit_translation} of {translated_city} City Public Security Bureau"
        
        # 如果不匹配固定格式，调用API翻译
        return self._call_deepseek_api(text)
    
    def _translate_place_name(self, name: str) -> str:
        """
        翻译地名（省、市名称）
        
        Args:
            name: 中文地名
            
        Returns:
            英文地名
        """
        # 常见省市名称映射
        place_names = {
            # 直辖市
            '北京': 'Beijing', '上海': 'Shanghai', '天津': 'Tianjin', '重庆': 'Chongqing',
            # 省份
            '河北': 'Hebei', '山西': 'Shanxi', '辽宁': 'Liaoning', '吉林': 'Jilin',
            '黑龙江': 'Heilongjiang', '江苏': 'Jiangsu', '浙江': 'Zhejiang', '安徽': 'Anhui',
            '福建': 'Fujian', '江西': 'Jiangxi', '山东': 'Shandong', '河南': 'Henan',
            '湖北': 'Hubei', '湖南': 'Hunan', '广东': 'Guangdong', '海南': 'Hainan',
            '四川': 'Sichuan', '贵州': 'Guizhou', '云南': 'Yunnan', '陕西': 'Shaanxi',
            '甘肃': 'Gansu', '青海': 'Qinghai', '台湾': 'Taiwan',
            # 自治区
            '内蒙古': 'Inner Mongolia', '广西': 'Guangxi', '西藏': 'Tibet',
            '宁夏': 'Ningxia', '新疆': 'Xinjiang',
            # 特别行政区
            '香港': 'Hong Kong', '澳门': 'Macao',
        }
        
        # 先查找映射表
        if name in place_names:
            return place_names[name]
        
        # 如果不在映射表中，调用API翻译
        translated = self._call_deepseek_api(f"请将中国地名'{name}'翻译成英文拼音，只返回拼音，不要其他内容")
        # 确保首字母大写
        return translated.strip().title()
    
    def _translate_text(self, text: str) -> str:
        """
        翻译单个文本
        
        Args:
            text: 要翻译的文本
            
        Returns:
            翻译后的文本
        """
        # 特殊处理：如果是纯英文字母和数字的组合（如驾驶证类型 B1D），直接返回
        if self._is_code_format(text):
            return text
        
        # 特殊处理：如果是18位证号（身份证号/驾驶证号），直接返回
        if self._is_id_number(text):
            return text
        
        # 特殊处理：日期格式转换
        date_converted = self._convert_date_format(text)
        if date_converted != text:
            return date_converted
        
        # 判断是否包含英文
        if self._contains_english(text):
            # 如果包含斜杠分隔的中英文（如"中国/CHN"），需要翻译中文部分
            if '/' in text:
                return self._translate_mixed_format(text)
            # 其他情况提取英文部分
            return self._extract_english(text)
        else:
            # 纯中文，需要翻译
            return self._call_deepseek_api(text)
    
    def _is_id_number(self, text: str) -> bool:
        """
        判断是否是身份证号/驾驶证号（18位数字+字母）
        """
        text = text.strip()
        # 18位，前17位是数字，最后一位可能是数字或X
        if len(text) == 18:
            return text[:17].isdigit() and (text[17].isdigit() or text[17].upper() == 'X')
        return False
    
    def _is_code_format(self, text: str) -> bool:
        """
        判断是否是代码格式（如驾驶证类型 B1D）
        纯英文字母和数字的组合，长度较短
        """
        text = text.strip()
        # 只包含字母、数字，长度不超过10个字符
        return bool(re.match(r'^[A-Z0-9]+$', text) and len(text) <= 10)
    
    def _convert_date_format(self, text: str) -> str:
        """
        转换日期格式
        将 YYYY-MM-DD 格式转换为英文格式（如 "August 12, 1973"）
        处理包含"至长期"的日期范围（如 "2019-09-01至长期" -> "September 1, 2019 to long-term"）
        处理日期范围（如 "2023-10-24 to 2033-10-24" -> "October 24, 2023 to October 24, 2033"）
        
        Args:
            text: 可能包含日期的文本
            
        Returns:
            转换后的文本
        """
        # 月份映射
        months = {
            '01': 'January', '02': 'February', '03': 'March', '04': 'April',
            '05': 'May', '06': 'June', '07': 'July', '08': 'August',
            '09': 'September', '10': 'October', '11': 'November', '12': 'December'
        }
        
        # 处理日期范围（YYYY-MM-DD to YYYY-MM-DD 或 YYYY-MM-DD至YYYY-MM-DD）
        # 匹配 "2023-10-24 to 2033-10-24" 或 "2023-10-24至2033-10-24"
        date_range_match = re.match(r'(\d{4})-(\d{2})-(\d{2})\s*(?:to|至)\s*(\d{4})-(\d{2})-(\d{2})', text.strip())
        if date_range_match:
            year1, month1, day1, year2, month2, day2 = date_range_match.groups()
            month1_name = months.get(month1, month1)
            month2_name = months.get(month2, month2)
            day1 = str(int(day1))
            day2 = str(int(day2))
            return f"{month1_name} {day1}, {year1} to {month2_name} {day2}, {year2}"
        
        # 处理包含"至长期"的日期范围
        if '至长期' in text:
            # 匹配 YYYY-MM-DD至长期
            match = re.match(r'(\d{4})-(\d{2})-(\d{2})至长期', text)
            if match:
                year, month, day = match.groups()
                month_name = months.get(month, month)
                # 去掉日期前导零
                day = str(int(day))
                return f"{month_name} {day}, {year} to long-term"
        
        # 处理单个日期 YYYY-MM-DD
        match = re.match(r'^(\d{4})-(\d{2})-(\d{2})$', text.strip())
        if match:
            year, month, day = match.groups()
            month_name = months.get(month, month)
            # 去掉日期前导零
            day = str(int(day))
            return f"{month_name} {day}, {year}"
        
        # 不是日期格式，返回原文
        return text
    
    def _contains_english(self, text: str) -> bool:
        """判断文本是否包含英文字符"""
        return bool(re.search(r'[a-zA-Z]', text))
    
    def _extract_english(self, text: str) -> str:
        """从文本中提取英文部分"""
        english_parts = re.findall(r'[a-zA-Z\s]+', text)
        return ' '.join(english_parts).strip()
    
    def _translate_mixed_format(self, text: str) -> str:
        """
        翻译混合格式文本（如"中国/CHN"）
        提取中文部分进行翻译，保留英文缩写
        
        Args:
            text: 混合格式文本
            
        Returns:
            翻译后的文本（如"China/CHN"）
        """
        parts = text.split('/')
        if len(parts) == 2:
            chinese_part = parts[0].strip()
            english_part = parts[1].strip()
            
            # 翻译中文部分
            if not self._contains_english(chinese_part):
                translated = self._call_deepseek_api(chinese_part)
                return f"{translated}/{english_part}"
        
        # 如果格式不符合预期，直接翻译整个文本
        return self._call_deepseek_api(text)
    
    def _call_deepseek_api(self, text: str) -> str:
        """
        调用 DeepSeek API 进行翻译
        
        Args:
            text: 要翻译的中文文本
            
        Returns:
            翻译后的英文文本
            
        Raises:
            TranslationError: 翻译失败
        """
        try:
            response = self.client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {
                        "role": "system",
                        "content": "你是一个专业的翻译助手，专门翻译驾驶证上的文字。请将中文翻译成英文，只返回翻译结果，不要有任何解释。"
                    },
                    {
                        "role": "user",
                        "content": f"请将以下文字翻译成英文：{text}"
                    }
                ],
                temperature=0.3,
                max_tokens=100
            )
            
            translated_text = response.choices[0].message.content.strip()
            return translated_text
            
        except Exception as e:
            raise TranslationError(f"翻译失败: {str(e)}")
