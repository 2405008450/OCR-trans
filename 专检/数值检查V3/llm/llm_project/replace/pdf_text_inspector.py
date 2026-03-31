"""
PDF 文本检查工具
用于调试和查看 PDF 中实际存储的文本内容
"""

import fitz  # PyMuPDF
from pathlib import Path
import re


def inspect_pdf_text(pdf_path: str, search_text: str = None, page_num: int = None):
    """
    检查 PDF 中的文本内容
    
    Args:
        pdf_path: PDF 文件路径
        search_text: 要查找的文本（可选）
        page_num: 指定页码（可选，从0开始）
    """
    doc = fitz.open(pdf_path)
    
    print("=" * 80)
    print(f"PDF 文本检查: {Path(pdf_path).name}")
    print("=" * 80)
    print(f"总页数: {len(doc)}\n")
    
    # 确定要检查的页面
    if page_num is not None:
        pages_to_check = [page_num]
    else:
        pages_to_check = range(len(doc))
    
    for pnum in pages_to_check:
        if pnum < 0 or pnum >= len(doc):
            continue
        
        page = doc[pnum]
        
        print(f"\n{'=' * 80}")
        print(f"第 {pnum + 1} 页")
        print('=' * 80)
        
        # 提取文本
        text = page.get_text()
        
        if search_text:
            # 查找特定文本
            print(f"\n查找文本: '{search_text}'")
            print("-" * 80)
            
            # 精确匹配
            if search_text in text:
                print(f"✓ 找到精确匹配")
                # 显示上下文
                idx = text.find(search_text)
                start = max(0, idx - 50)
                end = min(len(text), idx + len(search_text) + 50)
                context = text[start:end]
                print(f"上下文: ...{context}...")
            else:
                print(f"✗ 未找到精确匹配")
                
                # 尝试不区分大小写
                if search_text.lower() in text.lower():
                    print(f"✓ 找到不区分大小写的匹配")
                    idx = text.lower().find(search_text.lower())
                    start = max(0, idx - 50)
                    end = min(len(text), idx + len(search_text) + 50)
                    context = text[start:end]
                    print(f"上下文: ...{context}...")
                else:
                    # 尝试移除空格后匹配
                    text_no_space = text.replace(" ", "").replace("\n", "")
                    search_no_space = search_text.replace(" ", "")
                    
                    if search_no_space in text_no_space:
                        print(f"✓ 找到忽略空格的匹配")
                        print(f"提示: 文本中可能包含额外的空格或换行符")
                    else:
                        print(f"✗ 完全未找到")
                        print(f"\n建议: 检查以下可能的原因:")
                        print(f"  1. 文本可能被分割成多个部分")
                        print(f"  2. 使用了特殊字符或编码")
                        print(f"  3. 文本在图片中（无法搜索）")
                        
                        # 显示相似的文本片段
                        print(f"\n页面中包含的数字和关键词:")
                        numbers = re.findall(r'\d+[\d\s,\.]*', text)
                        if numbers:
                            print(f"  数字: {', '.join(set(numbers[:10]))}")
                        
                        # 查找包含部分搜索文本的片段
                        words = search_text.split()
                        if len(words) > 1:
                            for word in words:
                                if len(word) > 2 and word in text:
                                    print(f"  找到部分匹配: '{word}'")
        else:
            # 显示完整文本
            print("\n页面文本内容:")
            print("-" * 80)
            print(text[:1000])  # 只显示前1000个字符
            if len(text) > 1000:
                print(f"\n... (还有 {len(text) - 1000} 个字符)")
        
        # 显示文本块信息
        blocks = page.get_text("blocks")
        print(f"\n文本块数量: {len(blocks)}")
        
        if search_text:
            print(f"\n包含搜索文本的文本块:")
            for i, block in enumerate(blocks):
                block_text = block[4] if len(block) > 4 else ""
                if search_text.lower() in block_text.lower():
                    print(f"  块 {i}: {block_text[:100]}")
    
    doc.close()


def extract_all_numbers(pdf_path: str):
    """
    提取 PDF 中所有的数字
    
    Args:
        pdf_path: PDF 文件路径
    """
    doc = fitz.open(pdf_path)
    
    print("=" * 80)
    print(f"提取所有数字: {Path(pdf_path).name}")
    print("=" * 80)
    
    all_numbers = set()
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        text = page.get_text()
        
        # 提取各种格式的数字
        patterns = [
            r'\d+',  # 整数
            r'\d+\.\d+',  # 小数
            r'\d+%',  # 百分比
            r'\d+\s*billion',  # 带单位
            r'\d+\s*million',
            r'\d+\s*yuan',
            r'\d+\s*dollar',
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            all_numbers.update(matches)
    
    doc.close()
    
    print(f"\n找到 {len(all_numbers)} 个不同的数字:")
    for num in sorted(all_numbers):
        print(f"  - {num}")


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("用法:")
        print("  python pdf_text_inspector.py <pdf文件> [搜索文本] [页码]")
        print("\n示例:")
        print("  python pdf_text_inspector.py document.pdf")
        print("  python pdf_text_inspector.py document.pdf '620billion'")
        print("  python pdf_text_inspector.py document.pdf '620billion' 0")
        sys.exit(1)
    
    pdf_path = sys.argv[1]
    search_text = sys.argv[2] if len(sys.argv) > 2 else None
    page_num = int(sys.argv[3]) if len(sys.argv) > 3 else None
    
    if not Path(pdf_path).exists():
        print(f"错误: 文件不存在 - {pdf_path}")
        sys.exit(1)
    
    inspect_pdf_text(pdf_path, search_text, page_num)
    
    if not search_text:
        print("\n" + "=" * 80)
        extract_all_numbers(pdf_path)
