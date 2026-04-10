import os
import sys
from pathlib import Path
from parsers.pdf.pdf_parser import parse_pdf
from parsers.txt.txt_parser import parse_txt

from utils.txt_files import write_txt_with_timestamp

PROJECT_ROOT = Path(__file__).resolve().parents[3]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from app.service.gemini_service import generate_text


class rule:
    # 文本对比函数，利用OpenAI GPT对比原文和译文
    def compare_texts(self, original_text):
        prompt = f"""
        原文：{original_text}
        
        ##【步骤】：
        1. 深入分析文档逻辑。
        2. 按照金融翻译的严谨性要求，提取格式/术语/禁忌。
        3. 输出结果。

        ##【任务要求】：
        1. 绝对忠实于原文：严禁篡改规则细节，不得凭空臆造。
        2. 零遗漏：必须覆盖文档中涉及的所有格式、用词、排版、数字和层级规定。
        3. 可执行性：输出的内容必须能够直接复制到 GPT/Claude 中作为系统提示词使用。

        # Output Format (强制输出格式)
        请不要输出任何其他无关内容并且按以下结构输出：
        (一) [分类名称]
        [规则描述]：(条目化，清晰简洁，确保无歧义)
        (二) [分类名称]
        ...      
"""

        full_response = generate_text(
            system_prompt="你是一位精通中英金融互译的资深审校专家，擅长从各类银行文件中提炼翻译规则，不要篡改原文内容和改变原文意思。",
            user_prompt=prompt,
            model=os.getenv("ZHONGFANYI_MODEL_NAME", "google/gemini-3.1-pro-preview"),
            route=os.getenv("GEMINI_ROUTE", "openrouter"),
            temperature=0,
            max_output_tokens=65536,
        )
        print(full_response, end="")
        return full_response.strip()
# 主程序
if __name__ == "__main__":
    rule_gen = rule()
    ai_rule_path = r"C:\Users\Administrator\Desktop\项目文档文件\中翻译中译英规则\1. 通用规则\银行稿件_翻译规则_通用_220307.pdf"
    ai_rule_text = parse_pdf(ai_rule_path)
    print("AI正在进行处理")
    ai_rule=rule_gen.compare_texts(ai_rule_text)
    name,path=write_txt_with_timestamp(ai_rule,r"C:\Users\Administrator\Desktop\中翻译规则检查\llm\llm_project\rule")
    print(ai_rule)
    print(path)
    ai_target_path = path
    # 执行解析
    txt_text = parse_txt(ai_target_path)
    print(txt_text)



