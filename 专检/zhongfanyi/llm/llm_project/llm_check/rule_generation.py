import os
from dotenv import load_dotenv
from openai import OpenAI
from zhongfanyi.llm.llm_project.parsers.pdf.pdf_parser import parse_pdf
from zhongfanyi.llm.llm_project.parsers.txt.txt_parser import parse_txt

from zhongfanyi.llm.llm_project.utils.txt_files import write_txt_with_timestamp

# 加载API密钥：优先从项目根目录 .env 读取（与 FastAPI 主应用一致）
# 专检/zhongfanyi/llm/llm_project/llm_check -> 项目根 fastapi-llm-demo 需 6 层 ..
_project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", ".."))
_env_path = os.path.join(_project_root, ".env")
if os.path.isfile(_env_path):
    load_dotenv(_env_path)
# 兼容 OPENROUTER_* 与旧版 API_KEY/BASE_URL
api_key = os.getenv("OPENROUTER_API_KEY") or os.getenv("API_KEY")
base_url = os.getenv("OPENROUTER_BASE_URL") or os.getenv("BASE_URL", "https://openrouter.ai/api/v1")
client = OpenAI(
    api_key=api_key,
    base_url=base_url,
)


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

        try:
            # 使用正确的API调用方式，并启用流式响应
            response = client.chat.completions.create(
                extra_headers={
                    "HTTP-Referer": "<YOUR_SITE_URL>",  # Optional. Site URL for rankings on openrouter.ai.
                    "X-Title": "<YOUR_SITE_NAME>",  # Optional. Site title for rankings on openrouter.ai.
                },
                model="google/gemini-3-pro-preview",  # 使用 OpenAI 的 google/gemini-3-pro-preview 模型
                max_tokens=65536,
                messages=[
                    {"role": "system",
                     "content": "你是一位精通中英金融互译的资深审校专家，擅长从各类银行文件中提炼翻译规则，不要篡改原文内容和改变原文意思。"},
                    {"role": "user", "content": prompt}
                ],
                temperature=0,  # 设置温度为0，确保生成的内容精确、简洁
                stream=True  # 开启流式响应
            )

            # 流式输出处理
            full_response = ""
            for chunk in response:
                if chunk.choices and chunk.choices[0].delta.content:
                    message_content = chunk.choices[0].delta.content
                    full_response += message_content
                    print(message_content, end="")  # 实时输出返回的内容

            # 返回完整的流式响应内容
            return full_response.strip()

        except Exception as e:
            print("调用 API 失败，请检查账户余额、API Key 有效性或网络环境是否正常。")
            return  # 退出程序
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



