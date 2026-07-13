import json
import re
import codecs
from typing import List


def extract_and_parse(file_path: str) -> List[dict]:
    # 1. 读取原始内容
    try:
        with codecs.open(file_path, 'r', encoding='utf-8-sig') as f:
            raw_content = f.read()
    except Exception as e:
        print(f"❌ 读取文件失败: {e}")
        return []

    return parse_json_content(raw_content)


def parse_json_content(raw_content: str) -> List[dict]:
    """从原始字符串中解析 JSON 错误列表。

    支持直接 JSON 数组、Markdown 代码块包裹、以及被序列化为字符串的 JSON。
    可用于解析 LLM 返回的原始响应文本。
    """
    if not raw_content or not raw_content.strip():
        return []

    # 2. 【核心修复】解包逻辑
    # 如果文件内容看起来像是一个被转义的字符串 (以 " 开头)，先把它还原
    content = raw_content
    if content.strip().startswith('"'):
        try:
            # json.loads 会处理掉外层的引号，并将 \\n 还原为真正的换行符
            content = json.loads(content)
            print("✅ 成功解包：检测到文件是序列化字符串，已完成解包。")
        except json.JSONDecodeError:
            print("⚠️ 尝试解包失败，将按原样处理。")

    # 2.5 移除 Markdown 代码块标记
    if content.strip().startswith('```'):
        lines = content.split('\n')
        # 找到第一个 ``` 和最后一个 ```
        start_idx = -1
        end_idx = -1
        for i, line in enumerate(lines):
            if line.strip().startswith('```'):
                if start_idx == -1:
                    start_idx = i
                else:
                    end_idx = i
        
        if start_idx != -1 and end_idx != -1:
            json_lines = lines[start_idx + 1:end_idx]
            content = '\n'.join(json_lines)
            print("✅ 移除 Markdown 代码块标记")

    # 3. 尝试直接解析 JSON（如果格式正确）
    try:
        data = json.loads(content)
        if isinstance(data, list):
            print(f"✅ 直接解析成功：提取到 {len(data)} 个对象")
            # 排序
            data.sort(key=lambda x: int(x.get("错误编号", 99999)))
            return data
    except json.JSONDecodeError as e:
        print(f"⚠️ 直接解析失败 ({e.msg})，使用栈式提取...")

    # 4. 栈式提取器：从 Markdown 中提取所有 {} 块
    # 【关键修复】在提取每个对象后，修复控制字符问题
    objects = []
    stack_depth = 0
    start_index = -1

    for i, char in enumerate(content):
        if char == '{':
            if stack_depth == 0:
                start_index = i
            stack_depth += 1
        elif char == '}':
            if stack_depth > 0:
                stack_depth -= 1
                if stack_depth == 0 and start_index != -1:
                    json_str = content[start_index: i + 1]
                    
                    # 【关键修复】尝试修复控制字符问题
                    # 方法1: 直接解析
                    try:
                        obj = json.loads(json_str)
                        objects.append(obj)
                    except json.JSONDecodeError:
                        # 方法2: 替换未转义的控制字符后再解析
                        try:
                            # 使用 repr 和 eval 来处理控制字符
                            # 这会自动转义所有控制字符
                            fixed_str = json_str.encode('unicode_escape').decode('ascii')
                            # 但这会双重转义已经转义的字符，所以需要还原
                            fixed_str = fixed_str.replace('\\\\', '\\')
                            obj = json.loads(fixed_str)
                            objects.append(obj)
                            print(f"  ✓ 修复控制字符后成功解析错误 {obj.get('错误编号', '?')}")
                        except:
                            # 方法3: 使用 strict=False（Python 3.6+）
                            try:
                                # 先手动替换常见的控制字符
                                fixed_str = json_str
                                # 不要替换已经转义的 \t \n 等
                                # 只替换未转义的制表符和换行符
                                import re
                                # 查找未转义的制表符（不在 \" 之间的 \t）
                                # 简化处理：直接替换所有制表符为 \\t
                                fixed_str = fixed_str.replace('\t', '\\t')
                                fixed_str = fixed_str.replace('\r', '\\r')
                                # 但保留 JSON 结构中的换行符
                                
                                obj = json.loads(fixed_str)
                                objects.append(obj)
                                print(f"  ✓ 手动修复控制字符后成功解析错误 {obj.get('错误编号', '?')}")
                            except Exception as final_error:
                                print(f"  ✗ 无法解析对象 (位置 {start_index}-{i}): {final_error}")
                    
                    start_index = -1

    # 5. 排序
    objects.sort(key=lambda x: int(x.get("错误编号", 99999)))

    print(f"DEBUG: 成功识别并提取了 {len(objects)} 个完整对象。")
    return objects


if __name__ == "__main__":
    body_result_path = r"C:\Users\Administrator\Desktop\中翻译规则检查\llm\llm_project\zhengwen\output_json\文本对比结果_20260305_114752.json"
    print("\n正在启动超强提取模式...")
    results = extract_and_parse(body_result_path)
    print(results)
    for res in results:
        print(res)

    if results:
        print(f"✅ 解析成功！共获取 {len(results)} 个错误条目。")
        print(f"第一个错误: ID {results[0].get('错误编号')} - {results[0].get('错误类型')}")
    else:
        print("❌ 依然没有提取到数据。请再次运行并告诉我报错信息。")