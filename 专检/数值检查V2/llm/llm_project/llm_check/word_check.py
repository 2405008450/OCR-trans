import os
import traceback
from dotenv import load_dotenv
from openai import OpenAI

from 数值检查1.llm.llm_project.llm_check.rule_generation import rule
from 数值检查1.llm.llm_project.parsers.word.body_extractor import extract_body_text
from 数值检查1.llm.llm_project.parsers.word.footer_extractor import extract_footers
from 数值检查1.llm.llm_project.parsers.word.header_extractor import extract_headers

# 配置代理地址
os.environ["HTTP_PROXY"] = "http://127.0.0.1:7890"
os.environ["HTTPS_PROXY"] = "http://127.0.0.1:7890"
# 加载API密钥
load_dotenv()
api_key = os.getenv("API_KEY")
base_url = os.getenv("BASE_URL")
client = OpenAI(
    api_key=api_key,
    base_url=base_url,
)


class Match:
    # 文本对比函数，利用OpenAI GPT对比原文和译文
    def word_compare_texts(self, original_text, translated_text,project_type):
        prompt = f"""
        2. 待检查文本
        译文： {original_text}
        译文： {translated_text}
        项目类型： {project_type}        当前项目类型： [请注明：通用/中翻译/资翻译 OR 玲翻译]
        
        Role: 你是一位极其严谨的专业翻译审校专家，擅长“专有名词专项检查（Proper Noun Check）”。

        Task: 请对比以下【译文】与【译文】，根据我提供的【专项检查规则】，找出译文中不符合要求的地方，并给出修改建议。
        1. 专项检查规则库
        A. 人名规则
        通用/中/资翻译： 姓前名后（如：Zhang Sanfeng）。
        玲翻译： 名前姓后（如：Sanfeng Zhang）。
        港澳台人名： 优先查官网/维基/社媒。若无，香港用粤语拼音（英音），澳门用葡文音译，台湾用威妥玛拼音。
        参考文献人名： 缩写格式为“姓, 名首字母.”（如：Ding, Y.）。
        2人：A & B
        3人：A, B & C
        4人及以上：A, B, C et al.
        B. 机构与公司
        必须以官网、官方登记名、新闻报道或Google图片为准。严禁直接拼音翻译，严禁参考百度。
        C. 地址规则（中译英）
        层/楼： 统一用 /F（如 3/F）。
        方位： “东/西/南/北”放在路名前面。
        号： “X号”放在路名之前，不加 No.，不加逗号。
        顺序： 由小到大（单元/室 > 座 > 楼层 > 路号 > 区 > 市）。
        D. 地址规则（英译中）
        香港地址： 必须查询并照搬官方中文名称，不得自译。
        外国地址： 顺序由大到小。省/市参考官方/新华社译名；区/街/楼优先查证，查不到则音译。
        
        # 工作流程:
        - Step 1: 扫描原文句子，提取所有人名/机构与公司/地址。
        - Step 2: 扫描对应译文句子，提取对应的信息。
        - Step 3: 逐一比对。若发现任何字符级的不一致（包括标点符号在数值中的用法），标记为错误。

        #输出示例：
        仅输出 JSON 数组，不得包含说明文字。格式如下：
        [
          {{
            "错误编号": "1",
            "错误类型": "人名格式、地址方位错误、拼音体系错误等",
            "原文数值": "原文提取的译文片段",
            "译文数值": "译文提取的译文片段",
            "译文修改建议值": "修正后的完整译文片段",
            "修改理由": "简述违反的具体规则（如：数量级错误/单位不符/跳号）",
            "原文上下文": "包含该数值的原文完整句",
            "译文上下文": "包含该数值的译文完整句",
            "替换锚点": "译文中需要被替换的精确字符片段"
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
                model="google/gemini-3.1-pro-preview",  # 使用 OpenAI 的 google/gemini-3-pro-preview 模型
                max_tokens=65532,
                messages=[
                    {"role": "system",
                     "content": "你是中译英以及其他语种译文合规审校员，只负责依据要求对译文做错误类型规则符合性检查与修改建议，不要自行修正、不要补全缺失信息。"},
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
            raise e


# 主程序
if __name__ == "__main__":
    # 示例文件路径
    original_path = r"C:\Users\Administrator\Desktop\项目文件\专检\数值检查1\测试文件\原文-数值检查测试文件1.docx"  # 请替换为原文文件路径
    translated_path = r"C:\Users\Administrator\Desktop\项目文件\专检\数值检查1\测试文件\译文-数值检查测试文件1.docx"  # 请替换为译文文件路径

    # 处理页眉
    original_header_text = extract_headers(original_path)
    translated_header_text = extract_headers(translated_path)
    # 处理页脚
    original_footer_text = extract_footers(original_path)
    translated_footer_text = extract_footers(translated_path)
    # 处理正文(含脚注/表格/自动编号)
    original_body_text = extract_body_text(original_path)
    translated_body_text = extract_body_text(translated_path)
    # print("======页眉===========")
    # print(original_header_text)
    # print(translated_header_text)
    # print("======页脚===========")
    # print(original_footer_text)
    # print(translated_footer_text)
    # print("======正文===========")
    # print(original_body_text)
    # print(translated_body_text)

    # # 实例化对象并进行对比
    matcher = Match()
    # 正文对比
    print("======正在检查正文===========")
    body_result = matcher.word_compare_texts(original_body_text, translated_body_text,"中翻译")
    # 页眉对比
    print("======正在检查页眉===========")
    header_result = matcher.word_compare_texts(original_header_text, translated_header_text,"中翻译")
    # 页脚对比
    print("======正在检查页脚===========")
    footer_result = matcher.word_compare_texts(original_footer_text, translated_footer_text,"中翻译")
    print("================================")



