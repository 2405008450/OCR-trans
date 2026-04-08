"""
失败案例分析工具

分析4个失败案例的原因：
1. '第二条  Article 6' - 中英混合
2. '严控标签：...' - 长中文段落
3. '（1）' - 中文括号编号
4. 'Legal & Compliance' - 特殊字符 &
5. '1–10' - 特殊破折号 (en-dash)
"""

from docx import Document
from pathlib import Path
import re

def analyze_special_characters():
    """分析特殊字符问题"""
    print("=" * 80)
    print("特殊字符分析")
    print("=" * 80)
    
    cases = [
        ('第二条  Article 6', '中英混合，双空格'),
        ('xiv. BOC Financial Technology', '长中文段落'),
        ('Level one requires developers to be proficient in using the software;', '长中文段落'),
        ('严控标签：一般为已经停止服务/更新、存在严重安全漏洞或者不再推荐使用的开源软件或某个版本，由管理方式添加严控标识。', '长中文段落'),
        ('（1）', '中文括号 vs 英文括号'),
        ('Legal & Compliance', '& 符号'),
        ('1–10', 'en-dash (–) vs hyphen (-)'),
    ]
    
    for text, desc in cases:
        print(f"\n文本: {text[:50]}...")
        print(f"描述: {desc}")
        print(f"长度: {len(text)} 字符")
        
        # 检查特殊字符
        special_chars = []
        for char in text:
            code = ord(char)
            if code > 127 or char in '&–—':
                special_chars.append(f"'{char}' (U+{code:04X})")
        
        if special_chars:
            print(f"特殊字符: {', '.join(special_chars[:10])}")
            if len(special_chars) > 10:
                print(f"  ... 还有 {len(special_chars) - 10} 个")


def suggest_fixes():
    """建议修复方案"""
    print("\n" + "=" * 80)
    print("修复建议")
    print("=" * 80)
    
    suggestions = [
        {
            'case': '第二条  Article 6',
            'problem': '中英混合文本，包含双空格',
            'solutions': [
                '1. 在 clean_text_thoroughly 中规范化空格（多个空格→单空格）',
                '2. 在匹配时忽略空格数量差异',
                '3. 使用正则表达式 r"第二条\\s+Article\\s+6"'
            ]
        },
        {
            'case': '严控标签：...',
            'problem': '长中文段落，可能被分词或格式化影响',
            'solutions': [
                '1. 检查文档中是否有隐藏字符或格式标记',
                '2. 使用部分匹配策略（前50字符）',
                '3. 检查是否跨多个 runs'
            ]
        },
        {
            'case': '（1）',
            'problem': '中文括号 (U+FF08, U+FF09) vs 英文括号 (U+0028, U+0029)',
            'solutions': [
                '1. 在 clean_text_thoroughly 中统一括号：中文→英文',
                '2. 在匹配时同时尝试两种括号',
                '3. 使用正则: r"[（(]1[）)]"'
            ]
        },
        {
            'case': 'Legal & Compliance',
            'problem': '& 符号可能被转义或替换',
            'solutions': [
                '1. 检查文档中是否为 "Legal and Compliance"',
                '2. 在匹配时同时尝试 "&" 和 "and"',
                '3. 使用正则: r"Legal\\s+(?:&|and)\\s+Compliance"'
            ]
        },
        {
            'case': '1–10',
            'problem': 'en-dash (–, U+2013) vs hyphen (-, U+002D)',
            'solutions': [
                '1. 在 clean_text_thoroughly 中统一破折号类型',
                '2. 在匹配时同时尝试 –, —, -',
                '3. 使用正则: r"1[–—-]10"'
            ]
        }
    ]
    
    for idx, item in enumerate(suggestions, 1):
        print(f"\n[案例 {idx}] {item['case'][:30]}...")
        print(f"问题: {item['problem']}")
        print("解决方案:")
        for solution in item['solutions']:
            print(f"  {solution}")


def test_normalization():
    """测试文本规范化"""
    print("\n" + "=" * 80)
    print("文本规范化测试")
    print("=" * 80)
    
    def normalize_text(text):
        """规范化文本"""
        # 统一括号
        text = text.replace('（', '(').replace('）', ')')
        text = text.replace('【', '[').replace('】', ']')
        
        # 统一破折号
        text = text.replace('–', '-').replace('—', '-')
        
        # 统一空格
        text = re.sub(r'\s+', ' ', text)
        
        # 统一 & 和 and
        text = re.sub(r'\s+&\s+', ' and ', text)
        
        return text.strip()
    
    test_cases = [
        '第二条  Article 6',
        '（1）',
        'Legal & Compliance',
        '1–10'
    ]
    
    for text in test_cases:
        normalized = normalize_text(text)
        print(f"\n原始: {text}")
        print(f"规范: {normalized}")
        if text != normalized:
            print("  ✓ 已规范化")
        else:
            print("  - 无需规范化")


def generate_fix_code():
    """生成修复代码"""
    print("\n" + "=" * 80)
    print("建议的代码修复")
    print("=" * 80)
    
    code = '''
# 在 replace/word/replace_clean.py 的 clean_text_thoroughly 函数中添加：

def clean_text_thoroughly(text: str) -> str:
    """彻底清洗文本，用于匹配对比"""
    if not text:
        return ""
    
    # 统一括号（中文→英文）
    text = text.replace('（', '(').replace('）', ')')
    text = text.replace('【', '[').replace('】', ']')
    text = text.replace('「', '"').replace('」', '"')
    
    # 统一破折号（en-dash, em-dash → hyphen）
    text = text.replace('–', '-').replace('—', '-')
    text = text.replace('‐', '-').replace('‑', '-')
    
    # 统一引号
    text = text.replace('"', '"').replace('"', '"')
    text = text.replace("'", "'").replace("'", "'")
    
    # 统一空格（多个空格→单空格）
    text = re.sub(r'\\s+', ' ', text)
    
    # 移除零宽字符
    text = re.sub(r'[\\u200b-\\u200f\\ufeff]', '', text)
    
    return text.strip()


# 在匹配时添加 & 和 and 的互换尝试：

def try_match_with_variants(old_value, full_text):
    """尝试多种变体匹配"""
    variants = [old_value]
    
    # & ↔ and
    if ' & ' in old_value:
        variants.append(old_value.replace(' & ', ' and '))
    if ' and ' in old_value:
        variants.append(old_value.replace(' and ', ' & '))
    
    # 中文括号 ↔ 英文括号
    if '（' in old_value or '）' in old_value:
        variants.append(old_value.replace('（', '(').replace('）', ')'))
    if '(' in old_value or ')' in old_value:
        variants.append(old_value.replace('(', '（').replace(')', '）'))
    
    # 破折号变体
    if '–' in old_value:
        variants.append(old_value.replace('–', '-'))
        variants.append(old_value.replace('–', '—'))
    if '-' in old_value:
        variants.append(old_value.replace('-', '–'))
    
    # 尝试所有变体
    for variant in variants:
        if variant in full_text:
            return True, variant
    
    return False, None
'''
    
    print(code)


if __name__ == "__main__":
    analyze_special_characters()
    suggest_fixes()
    test_normalization()
    generate_fix_code()
    
    print("\n" + "=" * 80)
    print("分析完成")
    print("=" * 80)
    print("\n建议：")
    print("1. 更新 clean_text_thoroughly 函数以统一特殊字符")
    print("2. 在匹配时尝试多种字符变体")
    print("3. 对于长文本，使用部分匹配或上下文匹配")
    print("4. 检查文档中的实际字符编码")
