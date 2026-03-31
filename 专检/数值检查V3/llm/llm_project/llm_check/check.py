import os
import traceback
from dotenv import load_dotenv
from openai import OpenAI

from llm.llm_project.llm_check.rule_generation import rule
from llm.llm_project.parsers.word.body_extractor import extract_body_text
from llm.llm_project.parsers.word.footer_extractor import extract_footers
from llm.llm_project.parsers.word.header_extractor import extract_headers
# 加载API密钥
load_dotenv()
api_key=os.getenv("API_KEY")
base_url=os.getenv("BASE_URL")
print(api_key)
print(base_url)
client = OpenAI(
    api_key=api_key,
    base_url=base_url,
)
class Match:
    # 文本对比函数，利用OpenAI GPT对比原文和译文
    def compare_texts(self, original_text, translated_text):
        prompt = f"""
        # 角色
        你是一名资深的文件翻译审校专家，精通中英文以及其他语种文件中的数值、金额、日期、编号层级的一致性核查。

        # 任务
        1.严格对比【原文】与【译文】，识别并提取所有数值、单位、编号、日期翻译后译文与原文不一致的错误。
        2.对照原文检查译文出现错译、漏译的情况

        # 审查规则 (执行优先级最高)：
        1. **提取纯粹性：** 提取原文和译文数值时，严禁拼接、计算或修改数字。仅当译文的数值信息与其对应的原文在含义、精确度或逻辑上不符时，才提取为错误。
        2. **数值零误差：** 严禁四舍五入，严格检查单位数值一致性（如4亿和40 million是否一致）;但是要注意文本数字兼容：原文“十四”，译文“14”或“14th”应根据语境判定。如果语境是第14条，则一致；除非译文拼写错误或数值变动（如 14 变成 24）才标记为错误。       
        3. **编号连续性：** 检查 Article, Section, (i), (1) 等层级是否跳号、重号或者连续编号混用不同符号如1. II.3. ... 。
        4. **排除非数值文本：** 禁止修改任何嵌入在单词内部的类似数字的字符（如不可将 Trustworthy 误改为 Trus2rthy），除非它是明确的编号（如 Section 2）。
        5. **原文主权原则：** "译文修改建议值"字段必须以【原文】为唯一准则。禁止引入原文中不存在的年份、公司名称或背景信息。
        6. **最小拆分原则：** 若一句话内有多个数值错误，必须拆分为多个 JSON 对象，确保 `替换锚点` 精确到具体的错误片段。

        # 工作流程:
        - Step 1: 扫描原文句子，提取所有数值/日期/编号/单位。
        - Step 2: 扫描对应译文句子，提取对应的数值信息。
        - Step 3: 安照审查规则逐一比对。若发现任何翻译后译文与原文中的数值的不一致，按照最小拆分原则标记为错误。

        # 判定标准示例 (Few-shot)：
        1. 允许的转换
        - 原文：第十四条 | 译文：Article 14 -> [一致，无需修改]
        - 原文：三 | 译文：3 -> [一致，无需修改]            
        2. 严禁的幻觉修改 (错误示范)
        - 原文：2025 ESG报告 | 译文：2025 ESG Report -> 警告：严禁将其改为 2024 ESG Report。
        - 原文：Trustworthy | 译文：Trustworthy -> 警告：严禁识别为数值错误并改为 Trus2rthy。           
        3. 必须拦截的错误 (JSON 输出)
        - 原文：2025 | 译文：2024 -> [数值错误：年份不符]
        - 原文：10.00 | 译文：10 -> [数值错误：精度丢失]
        - 原文：第一节..., 第二节..., 第三节... | 译文：Article 1..., Article 2..., Article 4... -> [层级错误：跳号]

        #输出示例：
        仅输出 JSON 数组，不得包含说明文字。若无错误，则输出 `[]`。格式如下：
        [
          {{
            "错误编号": "1",
            "错误类型": "数值错误/层级错误/日期错误",
            "原文数值": "原文提取的译文片段",
            "译文数值": "译文提取的译文片段",
            "译文修改建议值": "修正后的译文片段",
            "修改理由": "简述违反的具体规则（如：数量级错误/单位不符/跳号）",
            "原文上下文": "包含该数值的原文完整句",
            "译文上下文": "包含该数值的译文完整句",
            "替换锚点": "译文中需要被替换的精确字符片段"
          }}
        ]

        #输入数据:
        - 原文：{original_text}
        - 译文：{translated_text}
        """

        try:
            # 使用正确的API调用方式，并启用流式响应
            response = client.chat.completions.create(
                extra_headers={
                    "HTTP-Referer": "<YOUR_SITE_URL>",  # Optional. Site URL for rankings on openrouter.ai.
                    "X-Title": "<YOUR_SITE_NAME>",  # Optional. Site title for rankings on openrouter.ai.
                },
                model="deepseek-chat",  # 使用 OpenAI 的 google/gemini-3-pro-preview 模型
                max_tokens=8192,
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
                # 检测流式错误（某些 API 会在 chunk 中注入错误）
                if hasattr(chunk, 'error') and chunk.error:
                    raise RuntimeError(f"SSE stream error: {chunk.error}")

            if not full_response.strip():
                raise RuntimeError("API 返回空内容")

            # 返回完整的流式响应内容
            return full_response.strip()

        except Exception as e:
            raise e


# 主程序
if __name__ == "__main__":
    # 示例文件路径
    original_path = r"D:\Users\Administrator\Desktop\项目文件\专检\数值检查V2_zhongyingduizhao\ceishiwenjian\测试文件\2-\第二章_投标文件0317 - 技术【改合并】.docx"  # 请替换为原文文件路径
    translated_path = r"D:\Users\Administrator\Desktop\项目文件\专检\数值检查V2_zhongyingduizhao\ceishiwenjian\测试文件\2-\译文\译文第二章-投标文件0317 - 技术【改合并】.docx"  # 请替换为译文文件路径

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



