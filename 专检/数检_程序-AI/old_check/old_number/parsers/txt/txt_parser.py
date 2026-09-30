import argparse
import re
from pathlib import Path


class TextContentFormatter:
    """文本内容格式化器"""

    @staticmethod
    def format_text(raw_content: str, mode: str = "structured") -> str:
        if not raw_content:
            return ""

        # 1. 基础清理：统一换行符，确保在任何系统上读取效果一致
        # 这不会改变你的文字内容，只是规范化不可见的换行符
        content = raw_content.replace("\r\n", "\n").strip()

        # 模式 A：structured (默认) -> 100% 保留原始结构和文字
        if mode == "structured":
            return content

        # 模式 B：clean -> 压缩空行，适合给 LLM 处理，节省 Token
        if mode == "clean":
            lines = [line.strip() for line in content.splitlines() if line.strip()]
            return "\n".join(lines)

        # 模式 C：pretty -> 自动识别标题并添加分割线（仅用于演示阅读）
        if mode == "pretty":
            # 在 ## 标题前增加分割线
            content = re.sub(r"(##\s+.+)", r"\n---\n\1", content)
            return content

        return content

def pretty_print_rules(text: str):
    """专门用于控制台美化输出的函数"""
    if not text:
        return

    print("\n" + "=" * 60)
    print("      金融翻译规则指南 (已解析)      ")
    print("=" * 60 + "\n")

    # 对输出内容进行简单的视觉增强
    enhanced_text = text.replace("##", "📌").replace("**", "")
    print(enhanced_text)

    print("\n" + "=" * 60)


def parse_txt(txt_path: str, mode: str = "clean"):
    """执行 TXT 文本解析"""
    txt_file = Path(txt_path)

    if not txt_file.exists():
        print(f"❌ 错误: 文件 {txt_file} 不存在")
        return None

    print(f"🔍 正在解析文本: {txt_file.name} ...")

    try:
        # 尝试使用 utf-8 读取，如果失败则尝试 gbk (常见于中文 Windows 环境)
        try:
            with open(txt_file, "r", encoding="utf-8") as f:
                content = f.read()
        except UnicodeDecodeError:
            with open(txt_file, "r", encoding="gbk") as f:
                content = f.read()

        # 使用格式化器处理文本
        formatted_output = TextContentFormatter.format_text(content, mode=mode)

        print(f"✅ 解析成功！读取长度: {len(content)} 字符。")
        return formatted_output

    except Exception as e:
        print(f"❌ 解析过程发生错误: {e}")
        return None


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="纯 TXT 解析工具")
    # 将参数改为可选（nargs='?'），并设置默认值
    parser.add_argument("txt_path", nargs='?', help="TXT 文件路径")
    parser.add_argument("--mode", choices=["clean", "structured"], default="clean", help="提取模式")

    args = parser.parse_args()

    # --- 逻辑调整 ---
    # 如果命令行没传路径，则使用你代码里硬编码的默认路径
    target_path = args.txt_path
    if not target_path:
        target_path = r"/zhongfanyi\llm\llm_project\rule\默认规则.txt"

    # 执行解析
    txt_text = parse_txt(target_path, args.mode)
    print(txt_text)

    # # 打印结果
    # if txt_text:
    #     print("-" * 20 + " 内容预览 " + "-" * 20)
    #     print(txt_text)
    # else:
    #     print("⚠️ 未能提取到任何内容内容。")