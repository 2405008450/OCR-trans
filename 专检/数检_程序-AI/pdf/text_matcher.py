"""
统一的文本匹配模块
为 Word 和 PDF 提供相同的智能匹配逻辑
"""

import re
import unicodedata
from typing import List, Tuple, Optional
from difflib import SequenceMatcher


# =========================
# 文本清洗工具（从 replace_clean.py 复制）
# =========================

def clean_text_thoroughly(text: str) -> str:
    """彻底清洗文本"""
    if not text: return ""

    # 1. 统一处理：将所有智能引号、撇号变体转为标准引号
    text = text.replace(''', "'").replace(''', "'")  # 左右单引号
    text = text.replace('"', '"').replace('"', '"')  # 左右双引号
    text = text.replace('`', "'")  # 反引号
    text = text.replace('´', "'")  # 重音符

    # 2. 将全角符号转为半角（扩展版）
    text = text.replace('（', '(').replace('）', ')')
    text = text.replace('，', ',').replace('。', '.')
    text = text.replace('：', ':').replace('；', ';')
    text = text.replace('！', '!').replace('？', '?')
    text = text.replace('【', '[').replace('】', ']')
    text = text.replace('《', '<').replace('》', '>')
    text = text.replace('　', ' ')  # 全角空格

    # 3. Unicode 标准化
    text = unicodedata.normalize('NFKC', text)

    # 4. 移除隐形字符和零宽字符
    text = re.sub(r'[\u200b-\u200f\u202a-\u202e\u2060-\u206f\ufeff\u00ad\xa0]', '', text)

    # 5. 统一空格（包括各种空白字符）
    text = re.sub(r'\s+', ' ', text)

    return text.strip()


def get_alphanumeric_fingerprint(text: str) -> str:
    """提取字母数字指纹"""
    return re.sub(r'[^a-zA-Z0-9]', '', text).lower()


def is_fuzzy_match(text: str, target: str, threshold: float = 0.9) -> bool:
    """判断两段文本的相似度是否超过阈值"""
    text_clean = clean_text_thoroughly(text)
    target_clean = clean_text_thoroughly(target)
    ratio = SequenceMatcher(None, text_clean, target_clean).ratio()
    return ratio >= threshold


# =========================
# 统一的文本匹配器
# =========================

class TextMatcher:
    """
    统一的文本匹配器
    提供多种匹配策略，适用于 Word 和 PDF
    """
    
    def __init__(self):
        self.match_strategies = [
            self._exact_match,
            self._cleaned_match,
            self._no_space_match,
            self._case_insensitive_match,
            self._fuzzy_match,
            self._fingerprint_match,
            self._partial_match,
        ]
    
    def find_best_match(
        self, 
        text: str, 
        search_text: str,
        return_strategy: bool = False
    ) -> Tuple[bool, Optional[int], Optional[int], Optional[str]]:
        """
        在文本中查找最佳匹配
        
        Args:
            text: 要搜索的文本
            search_text: 要查找的文本
            return_strategy: 是否返回匹配策略名称
            
        Returns:
            (是否找到, 起始位置, 结束位置, 策略名称)
        """
        for strategy in self.match_strategies:
            result = strategy(text, search_text)
            if result[0]:  # 找到匹配
                if return_strategy:
                    return result
                else:
                    return (result[0], result[1], result[2], None)
        
        return (False, None, None, None)
    
    def _exact_match(self, text: str, search_text: str) -> Tuple[bool, Optional[int], Optional[int], str]:
        """策略1: 精确匹配"""
        if search_text in text:
            start = text.index(search_text)
            end = start + len(search_text)
            return (True, start, end, "精确匹配")
        return (False, None, None, "精确匹配")
    
    def _cleaned_match(self, text: str, search_text: str) -> Tuple[bool, Optional[int], Optional[int], str]:
        """策略2: 清洗后匹配"""
        text_clean = clean_text_thoroughly(text)
        search_clean = clean_text_thoroughly(search_text)
        
        if search_clean in text_clean:
            # 尝试在原文中找到对应位置
            # 简化：如果清洗后能找到，尝试在原文中也找到
            if search_text in text:
                start = text.index(search_text)
                end = start + len(search_text)
                return (True, start, end, "清洗后匹配")
            else:
                # 清洗后能找到，但原文找不到，返回清洗后的位置
                start = text_clean.index(search_clean)
                end = start + len(search_clean)
                return (True, start, end, "清洗后匹配(近似位置)")
        return (False, None, None, "清洗后匹配")
    
    def _no_space_match(self, text: str, search_text: str) -> Tuple[bool, Optional[int], Optional[int], str]:
        """策略3: 忽略空格匹配"""
        text_no_space = text.replace(' ', '').replace('\n', '').replace('\t', '')
        search_no_space = search_text.replace(' ', '').replace('\n', '').replace('\t', '')
        
        if search_no_space in text_no_space:
            # 尝试在原文中找到近似位置
            # 这里返回一个近似的位置
            try:
                # 找到第一个字符的位置作为起点
                first_char = next((c for c in search_text if c.strip()), None)
                if first_char and first_char in text:
                    start = text.index(first_char)
                    # 估算结束位置
                    end = start + len(search_text)
                    return (True, start, end, "忽略空格匹配")
            except:
                pass
        return (False, None, None, "忽略空格匹配")
    
    def _case_insensitive_match(self, text: str, search_text: str) -> Tuple[bool, Optional[int], Optional[int], str]:
        """策略4: 不区分大小写匹配"""
        text_lower = text.lower()
        search_lower = search_text.lower()
        
        if search_lower in text_lower:
            start = text_lower.index(search_lower)
            end = start + len(search_text)
            return (True, start, end, "不区分大小写匹配")
        return (False, None, None, "不区分大小写匹配")
    
    def _fuzzy_match(self, text: str, search_text: str) -> Tuple[bool, Optional[int], Optional[int], str]:
        """策略5: 模糊匹配（相似度 > 85%）"""
        if is_fuzzy_match(text, search_text, threshold=0.85):
            # 返回整个文本的范围
            return (True, 0, len(text), "模糊匹配")
        return (False, None, None, "模糊匹配")
    
    def _fingerprint_match(self, text: str, search_text: str) -> Tuple[bool, Optional[int], Optional[int], str]:
        """策略6: 指纹匹配（仅字母数字）"""
        fingerprint_search = get_alphanumeric_fingerprint(search_text)
        fingerprint_text = get_alphanumeric_fingerprint(text)
        
        if len(fingerprint_search) >= 3 and fingerprint_search in fingerprint_text:
            # 尝试找到近似位置
            # 这里返回整个文本的范围
            return (True, 0, len(text), "指纹匹配")
        return (False, None, None, "指纹匹配")
    
    def _partial_match(self, text: str, search_text: str) -> Tuple[bool, Optional[int], Optional[int], str]:
        """策略7: 部分匹配（搜索文本的主要部分在文本中）"""
        # 提取搜索文本中的关键词（长度 > 3 的单词）
        words = re.findall(r'\b\w{4,}\b', search_text)
        
        if not words:
            return (False, None, None, "部分匹配")
        
        # 检查是否有足够多的关键词在文本中
        matched_words = sum(1 for word in words if word.lower() in text.lower())
        
        if matched_words >= len(words) * 0.6:  # 至少60%的关键词匹配
            # 找到第一个匹配词的位置
            for word in words:
                if word.lower() in text.lower():
                    start = text.lower().index(word.lower())
                    end = start + len(search_text)
                    return (True, start, end, "部分匹配")
        
        return (False, None, None, "部分匹配")
    
    def generate_text_variants(self, text: str) -> List[str]:
        """
        生成文本的可能变体
        
        Args:
            text: 原始文本
            
        Returns:
            可能的文本变体列表
        """
        variants = [text]  # 包含原文
        
        # 1. 添加/移除空格的变体
        no_space = text.replace(" ", "")
        if no_space != text:
            variants.append(no_space)
        
        # 在数字和字母之间添加空格
        spaced = re.sub(r'(\d)([a-zA-Z])', r'\1 \2', text)
        if spaced != text:
            variants.append(spaced)
        
        # 2. 大小写变体
        if text != text.lower():
            variants.append(text.lower())
        if text != text.upper():
            variants.append(text.upper())
        if text != text.title():
            variants.append(text.title())
        
        # 3. 常见缩写和单位变体
        replacements = {
            'billion': ['bn', 'B', 'bil', ' billion'],
            'million': ['mn', 'M', 'mil', ' million'],
            'thousand': ['k', 'K', 'thou', ' thousand'],
            'USDollar': ['USD', 'US Dollar', 'US$', 'dollar', 'Dollar', ' US Dollar'],
            'yuan': ['CNY', '¥', 'RMB', 'Yuan', ' yuan'],
            'percent': ['%', 'pct'],
        }
        
        # 先生成单个替换的变体
        single_variants = []
        for key, values in replacements.items():
            if key.lower() in text.lower():
                for value in values:
                    variant = re.sub(key, value, text, flags=re.IGNORECASE)
                    if variant != text and variant not in variants:
                        single_variants.append(variant)
        
        variants.extend(single_variants)
        
        # 对于包含多个可替换项的文本，生成组合变体
        # 例如：580billionUSDollar → 580 billion US Dollar
        if len([k for k in replacements.keys() if k.lower() in text.lower()]) > 1:
            combined = text
            for key, values in replacements.items():
                if key.lower() in combined.lower():
                    # 使用第一个（通常是最标准的）替换
                    if values:
                        combined = re.sub(key, values[0] if not values[0].startswith(' ') else values[0].strip(), 
                                        combined, flags=re.IGNORECASE)
            
            # 添加带空格的版本
            combined_spaced = re.sub(r'(\d)([a-zA-Z])', r'\1 \2', combined)
            combined_spaced = re.sub(r'([a-z])([A-Z])', r'\1 \2', combined_spaced)
            
            if combined_spaced not in variants:
                variants.append(combined_spaced)
            
            # 额外尝试：在所有单词之间添加空格
            # 580billionUSDollar → 580 billion US Dollar
            fully_spaced = text
            # 先替换已知的单位
            for key, values in replacements.items():
                if key.lower() in fully_spaced.lower() and values:
                    # 使用带空格的版本
                    replacement = next((v for v in values if ' ' in v), values[0])
                    fully_spaced = re.sub(key, replacement, fully_spaced, flags=re.IGNORECASE)
            
            # 在数字和字母之间添加空格
            fully_spaced = re.sub(r'(\d)([a-zA-Z])', r'\1 \2', fully_spaced)
            # 在小写和大写之间添加空格
            fully_spaced = re.sub(r'([a-z])([A-Z])', r'\1 \2', fully_spaced)
            
            if fully_spaced not in variants:
                variants.append(fully_spaced)
        
        # 4. 移除标点符号的变体
        no_punct = re.sub(r'[^\w\s]', '', text)
        if no_punct != text and no_punct not in variants:
            variants.append(no_punct)
        
        return variants


# 全局实例
_matcher = TextMatcher()


def find_text_in_string(text: str, search_text: str) -> Tuple[bool, Optional[int], Optional[int], Optional[str]]:
    """
    在字符串中查找文本（使用所有策略）
    
    Args:
        text: 要搜索的文本
        search_text: 要查找的文本
        
    Returns:
        (是否找到, 起始位置, 结束位置, 匹配策略)
    """
    return _matcher.find_best_match(text, search_text, return_strategy=True)


def generate_search_variants(text: str) -> List[str]:
    """
    生成搜索文本的变体
    
    Args:
        text: 原始文本
        
    Returns:
        变体列表
    """
    return _matcher.generate_text_variants(text)


# 使用示例
if __name__ == "__main__":
    # 测试各种匹配策略
    test_cases = [
        ("This is 620 billion dollars", "620billion"),
        ("Document Number: TP202509", "TP202509"),
        ("The price is US Dollar 100", "USDollar"),
        ("Samsung 20%", "20%"),
    ]
    
    matcher = TextMatcher()
    
    for text, search in test_cases:
        found, start, end, strategy = matcher.find_best_match(text, search, return_strategy=True)
        print(f"\n文本: {text}")
        print(f"搜索: {search}")
        print(f"结果: {'找到' if found else '未找到'}")
        if found:
            print(f"位置: {start}-{end}")
            print(f"策略: {strategy}")
            print(f"匹配内容: {text[start:end]}")
        
        # 显示变体
        variants = matcher.generate_text_variants(search)
        print(f"生成的变体: {variants[:5]}")  # 只显示前5个
