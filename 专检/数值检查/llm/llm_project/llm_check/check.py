import os
import traceback
from pathlib import Path
from dotenv import load_dotenv
from openai import OpenAI
from llm.llm_project.parsers.body_extractor import extract_body_text
from llm.llm_project.parsers.footer_extractor import extract_footers
from llm.llm_project.parsers.header_extractor import extract_headers

# 加载API密钥（使用当前文件所在目录的 .env，确保从任何 CWD 启动都能找到）
_env_path = Path(__file__).resolve().parent / ".env"
load_dotenv(_env_path)
api_key=os.getenv("API_KEY")
base_url=os.getenv("BASE_URL")
client = OpenAI(
    api_key=api_key,
    base_url=base_url,
)

class Match:
    # 文本对比函数，利用OpenAI GPT对比原文和译文
    def compare_texts(self, original_text, translated_text):
        prompt = f"""
        原文：{original_text}
        译文（待检查）：{translated_text}

        文件类型与场景：
        - 文件类型：一般性文件 / 法律文件（合同、协议、诉讼文书等）
        - 场景：正文 / 表格 / PPT要点 / Excel要点

        【审校目标】：
        1) 逐条对照错误类型规则检查英文译文数值是否与原文中的数值一致；
        2) 不改变原文原意；不四舍五入数值；单位与位数严格按规则；

        【错误类型】（逐条执行，不要漏）：
        (一) 数值错误：
        - 检查译文全文是否与原文数值不一致，是否漏译 / 错译 / 数量级错误
        - 单位是否一致（%/公式符号等数量单位和计量单位）
        - 原文/译文大小写数值、数字（包括罗马数字和阿拉伯数字）必须完全一致
        - 零误差原则，禁止四舍五入
        (二) 标题与编号层级：
        - 禁止序号重复（如出现两个 Article 1）、顺序倒置、禁止跨越式跳号,检查章节序号（Chapter/Section）、条目序号（Article）、以及列表项序号（i/ii/iii/1/2/3）的连续性与中英映射准确性.
        (三) 时间与日期数值不对应与格式规范错误
        - 年、月、日、周期必须完全对应
        - 12/24 小时制必须时间一致
        ##请严格以json格式按顺序输出检查错误结果，不得输出任何其他内容：
        输出要求：
        1) 对于加粗斜体等这种非文本内容的格式问题在译文修改建议值加上中文括号备注如（加粗、斜体等格式）,如译文修改建议值: Chapter II(加粗)
        2) 输出的`原文数值`、`译文数值`、`原文上下文`、`译文上下文`、`替换锚点`字段的值必须严格按照原内容进行输出，不得做任何改动
        3) 若输入原文和译文有一个为空，则输出空值
        4) 若出现连续错误或者错误距离过近，必须需拆分成多个json对象错误或者整句作为json对象处理进行处理，保证`译文数值`字段为译文原内容片段。如This Contract is made in 7 counterparts, with 2 copies held by Party A and three copies held by Party B.中若7 3 three数值错误，则输出译文数值7 译文数值2 译文数值three 译文建议修改值6 译文建议修改值5 译文建议修改值five ，或者译文数值：This Contract is made in 7 counterparts, with 2 copies held by Party A and three copies held by Party B.译文建议修改值：This Contract is made in 6 counterparts, with 5 copies held by Party A and five copies held by Party B.严禁出现译文数值：7/3/three 译文建议修改值：6/5/five多个错误在一起
        输出格式示例：
                  [
                    {{
                    "错误编号": "2",
                    "错误类型": "数值错误",
                    "原文数值": "本合同一式六份",
                    "译文数值": "This Contract is made in 7 counterparts",
                    "译文修改建议值": "This Contract is made in 6 counterparts",
                    "修改理由": "违反了正文中加书名号的法律法规/议案名称等：英文需全部实词首字母大写",
                    "违反的规则": "正文中加书名号的法律法规/议案名称等：英文需全部实词首字母大写",
                    "原文上下文": "本合同一式六份",
                    "译文上下文": "This Contract is made in 7 counterparts",
                    "原文位置": "正文",
                    "译文位置": "正文",
                    "替换锚点": "made in 7 counterparts"
                    }}
                ]
        """

        try:
            # 使用正确的API调用方式，并启用流式响应
            response = client.chat.completions.create(
                extra_headers={
                    "HTTP-Referer": "<YOUR_SITE_URL>",  # Optional. Site URL for rankings on openrouter.ai.
                    "X-Title": "<YOUR_SITE_NAME>",  # Optional. Site title for rankings on openrouter.ai.
                },
                model="google/gemini-2.5-pro",  # 使用 OpenAI 的 google/gemini-3-pro-preview 模型
                max_tokens=65536,
                messages=[
                    {"role": "system",
                     "content": "你是中译英译文合规审校员，只负责依据要求对英文译文做错误类型规则符合性检查与修改建议，不要自行修正、不要补全缺失信息。"},
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
            print(f"Error occurred: {e}")
            traceback.print_exc()
            return "Error occurred during API call."


# 主程序
if __name__ == "__main__":
    # 示例文件路径
    original_path = r"C:\Users\Administrator\Desktop\project\效果\TP251117023，北京中翻译，中译英（字数2w）\原文-B251124195-Y-更新1121-附件1：中国银行股份有限公司模型风险管理办法（2025年修订）.docx"  # 请替换为原文文件路径
    translated_path = r"C:\Users\Administrator\Desktop\project\效果\TP251117023，北京中翻译，中译英（字数2w）\测试译文-清洁版-B251124195-附件1：中国银行股份有限公司模型风险管理政策（2025年修订）-.docx"  # 请替换为译文文件路径

    #处理页眉
    original_header_text=extract_headers(original_path)
    translated_header_text=extract_headers(translated_path)
    #处理页脚
    original_footer_text = extract_footers(original_path)
    translated_footer_text = extract_footers(translated_path)
    #处理正文(含脚注/表格/自动编号)
    original_body_text=extract_body_text(original_path)
    translated_body_text=extract_body_text(translated_path)
    print("======页眉===========")
    print(original_header_text)
    print(translated_header_text)
    print("======页脚===========")
    print(original_footer_text)
    print(translated_footer_text)
    print("======正文===========")
    print(original_body_text)
    print(translated_body_text)

    # # 实例化对象并进行对比
    matcher = Match()
    #正文对比
    print("======正在检查正文===========")
    body_result = matcher.compare_texts(original_body_text, translated_body_text)
    #页眉对比
    print("======正在检查页眉===========")
    header_result = matcher.compare_texts(original_header_text, translated_header_text)
    #页脚对比
    print("======正在检查页脚===========")
    footer_result = matcher.compare_texts(original_footer_text, translated_footer_text)
    print("================================")



