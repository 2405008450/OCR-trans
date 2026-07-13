import os
import traceback
from dotenv import load_dotenv
from openai import OpenAI

# 加载API密钥
load_dotenv()
api_key=os.getenv("API_KEY")
base_url=os.getenv("BASE_URL")
client = OpenAI(
    api_key=api_key,
    base_url=base_url,
)

class Match:
    # 文本对比函数，利用OpenAI GPT对比原文和译文
    @staticmethod
    def compare_texts(original_text, translated_text):
        prompt = f"""
        # 角色
        你是一名资深的文件翻译审校专家，精通中英文以及其他语种文件中的数值、金额、日期、编号层级的一致性核查。

        # 任务
        1.严格对比【原文】与【译文】，识别并提取所有数值、单位、编号、日期翻译后译文与原文不一致的错误。
        2.对照原文检查译文出现错译、漏译、多译的情况，以【原文】为唯一准则。
        3.若译文符合规则且无错译，严禁将其放入 JSON 列表中输出。

        # 审查规则 (执行优先级最高)：
        1. **原文主权原则：** "译文修改建议值"字段必须以【原文】为唯一准则。禁止引入原文中不存在的年份、公司名称或背景信息。
        2. **提取纯粹性：** 提取原文和译文数值时，严禁拼接、计算或修改数字。仅当译文的数值信息与其对应的原文在含义、精确度或逻辑上不符时，才提取为错误。
        3. **检查数值一致性：** 严格检查提取的单位数值翻译后是否与原文保持一致,严禁四舍五入小数点后非零的数值位;但是要注意文本数字兼容：原文“十四”，译文“14”或“14th”应根据语境判定。如果语境是第14条，则一致；除非译文拼写错误或数值变动（如 14 变成 24）才标记为错误。       
        4. **检查编号连续性：** 检查层级编号是否跳号、重号(如Article 1..., Article 1..., Article 2..., Article 4...)。
        5. **排除非数值文本：** 禁止修改任何嵌入在单词内部的类似数字的字符（如不可将 Trustworthy 误改为 Trus2rthy），除非它是明确的编号（如 Section 2）。
        6. **最小拆分原则：** 若一句话内有多个数值错误，必须拆分为多个 JSON 对象，确保 "译文数值""替换锚点"精确到具体的错误片段。

        # 工作流程:
        - Step 1: 扫描原文句子，提取所有数值/日期/编号/单位。
        - Step 2: 扫描对应译文句子，提取对应的数值信息。
        - Step 3: 安照审查规则逐一比对检查。若发现任何翻译后译文与原文中的数值的不一致，按照最小拆分原则标记为错误。

        # 判定标准示例 (Few-shot)：
        1. 允许的转换
        - 原文：第十四条 | 译文：Article 14 -> [一致，无需修改]
        - 原文：RMB108.60 billion | 译文：RMB108.6 billion -> [一致，无需修改]
        - 原文：三 | 译文：3 -> [一致，无需修改]            
        2. 严禁的幻觉修改 (错误示范)
        - 原文：2025 ESG报告 | 译文：2025 ESG Report -> 警告：严禁将其改为 2024 ESG Report。
        - 原文：Trustworthy | 译文：Trustworthy -> 警告：严禁识别为数值错误并改为 Trus2rthy。           
        3. 必须拦截的错误 (JSON 输出)
        - 原文：2025 | 译文：2024 -> [数值错误：年份不符]
        - 原文：1234.16亿 | 译文：123.42 billion -> [数值错误：精度丢失,严禁四舍五入，应为123.416 billion]
        - 原文：4亿 | 译文：40 million -> [数值错误：数据转换错误，应为400 million]
        - 原文：第一节..., 第二节..., 第三节... | 译文：Article 1..., Article 2..., Article 4... -> [层级错误：跳号](注意：多个错误要拆分,译文Article 4改为Article 3译文上下文根据所在位置单独提取)

        #输出示例：
        仅输出 JSON 数组，不得包含说明文字。若无错误，则输出 `[]`。格式如下：
            [
              {{
                "错误编号": "1",
                "原文上下文": "包含该数值的原文完整句",
                "译文上下文": "包含该数值的译文完整句",
                "原文数值": "原文提取的原文片段",
                "译文数值": "译文中提取的译文错误片段",
                "替换锚点": "译文中需要被替换的精确字符片段",
                "译文修改建议值": "修正后的提取的译文错误片段（要求这个字段的值来源于复制译文数值字段并修改其中的错误）",
                "错误类型": "数值错误",
                "修改理由": "简述违反的具体规则（如：数量级错误）",
                "违反的规则": "规则条款"
              }},
              {{
                "错误编号": "3",
                "原文上下文": "第二条  开源软件管理应遵循以下原则：",
                "译文上下文": "第二条 Article 6 Data demand management shall be in compliance with the following principles:",
                "原文数值": "第二条",
                "译文数值": "第二条 Article 6",
                "替换锚点": "第二条 Article 6",
                "译文修改建议值": "Article 2",
                "错误类型": "标题与编号层级",
                "修改理由": "译文保留了中文编号且阿拉伯数字错误（原文为二，译文误写为6）。",
                "违反的规则": "规则(二)：条目序号（Article）的连续性与中英映射准确性。"
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
                model="google/gemini-3-flash-preview",  # 使用 OpenAI 的 google/gemini-3-pro-preview 模型google/gemini-3.1-flash-lite   google/gemini-3-flash-preview
                max_tokens=65532,
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


    # # 示例文件路径
    # original_path = r"/数检/Shimizu World 2026 April_Combine（完）.docx"  # 请替换为原文文件路径
    # translated_path = r"/数检/Shimizu World 2026 April_Combine（完）.docx"  # 请替换为译文文件路径
    #
    # #处理正文(含脚注/表格/自动编号)
    # original_body_text=extract_body_text(original_path)
    # translated_body_text=extract_body_text(translated_path)
    #
    # print(original_body_text)
    # print(translated_body_text)
    #
    # # # 实例化对象并进行对比
    # # matcher = Match()
    # #正文对比
    # print("======正在检查正文===========")
    # body_result = compare_texts(original_body_text, translated_body_text)
    # print(body_result)