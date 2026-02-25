"""
Markdown JSON 清理工具
用于清理和标准化 Markdown 代码块中的 JSON 内容
"""

import json
import re
from typing import Any

from llm.utils.json_parser import jsonParser


def clean_markdown_json(content: str) -> str:
    """
    清理 Markdown 代码块中的 JSON 内容，并返回标准 JSON 格式字符串

    处理场景：
    1. 被双引号包裹的转义字符串（如 "\"[{...}]\""）
    2. Markdown 代码块标记（```json 或 ```）
    3. 多余的空白字符和换行符
    4. 转义字符的正确处理
    5. 多层嵌套的转义字符串

    Args:
        content: 待清理的原始内容

    Returns:
        清理后的标准 JSON 字符串

    Raises:
        ValueError: 当内容不是有效的 JSON 格式时
    """
    if not isinstance(content, str):
        raise TypeError(f"Expected str, got {type(content).__name__}")

    if not content.strip():
        raise ValueError("Content is empty")

    # 步骤 1: 去除首尾空白
    cleaned = content.strip()

    # 步骤 2: 处理被双引号包裹的整个字符串（可能多层嵌套）
    max_iterations = 5  # 防止无限循环
    iteration = 0

    while cleaned.startswith('"') and cleaned.endswith('"') and iteration < max_iterations:
        try:
            # 尝试使用 json.loads 解析（会自动处理转义）
            temp = json.loads(cleaned)
            if isinstance(temp, str):
                cleaned = temp
                iteration += 1
            else:
                # 如果解析结果不是字符串，说明已经是 JSON 对象了
                break
        except json.JSONDecodeError:
            # 如果 json.loads 失败，手动去除外层引号并处理转义
            cleaned = cleaned[1:-1]
            # 处理常见的转义字符
            cleaned = cleaned.replace('\\"', '"') \
                            .replace('\\n', '\n') \
                            .replace('\\t', '\t') \
                            .replace('\\r', '\r') \
                            .replace('\\\\', '\\')
            iteration += 1

    # 步骤 3: 移除 Markdown 代码块标记
    cleaned = re.sub(r'^```(?:json|JSON)?\s*\n?', '', cleaned)  # 移除开头的 ```json 或 ```
    cleaned = re.sub(r'\n?```\s*$', '', cleaned)                # 移除结尾的 ```

    # 步骤 4: 再次去除首尾空白
    cleaned = cleaned.strip()

    # 步骤 5: 验证并格式化 JSON
    try:
        # 解析 JSON 以验证格式
        parsed_data = json.loads(cleaned)
        # 重新序列化为标准格式（美化输出）
        return json.dumps(parsed_data, ensure_ascii=False, indent=2)
    except json.JSONDecodeError as e:
        # 提供更详细的错误信息
        error_pos = max(0, e.pos - 50)
        error_end = min(len(cleaned), e.pos + 50)
        error_context = cleaned[error_pos:error_end]
        raise ValueError(
            f"Invalid JSON format at line {e.lineno}, column {e.colno}:\n"
            f"Error: {e.msg}\n"
            f"Position: {e.pos}\n"
            f"Context: ...{error_context}..."
        )


def parse_markdown_json(content: str) -> Any:
    """
    解析 Markdown 中的 JSON 内容并返回 Python 对象

    Args:
        content: 待解析的原始内容

    Returns:
        解析后的 Python 对象（list, dict 等）

    Raises:
        ValueError: 当内容不是有效的 JSON 格式时
    """
    cleaned_json = clean_markdown_json(content)
    return json.loads(cleaned_json)


def load_json_file(file_path):
    """读取并解析JSON文件，返回错误列表"""
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            content = f.read()
            print("文件内容:", content)  # 打印前200个字符看看内容

            # 2. 将文件内容（字符串）传递给解析器
            error = clean_markdown_json(content)
            # 3. 解析JSON字符串为Python对象
            parsed_data = json.loads(error)

            # 确保返回的是列表
            if isinstance(parsed_data, list):
                return parsed_data
            elif isinstance(parsed_data, dict):
                return [parsed_data]  # 单个字典转换为列表
            else:
                print(f"⚠️ 解析出的数据类型不是 dict 或 list: {type(parsed_data)}")
                return []
    except FileNotFoundError:
        print(f"错误: 文件不存在 {file_path}")
    except Exception as e:
        print(f"错误: 读取或解析文件时出错 - {e}")
        return []

# ============================================================================
# 测试用例
# ============================================================================
if __name__ == "__main__":
    body_result_path = r"C:\Users\Administrator\Desktop\project\llm\llm_project\zhengwen\output_json\文本对比结果_20260208_144950.json"
    header_result_path = r"C:\Users\Administrator\Desktop\project\llm\llm_project\yemei\output_json\文本对比结果_20260208_144950.json"
    footer_result_path = r"C:\Users\Administrator\Desktop\project\llm\llm_project\yejiao\output_json\文本对比结果_20260208_145004.json"

    # 2) 读取错误报告并解析
    print("\n正在提取正文错误报告...")
    errors = load_json_file(body_result_path)
    print(type(errors))
    print("返回json格式：",errors)
