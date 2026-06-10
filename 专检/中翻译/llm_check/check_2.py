import os
import traceback
from dotenv import load_dotenv
from llm_check.openrouter_config import create_openrouter_client, get_model_name

from parsers.word.body_extractor1 import extract_body_text

# 配置代理地址
# os.environ["HTTP_PROXY"] = "http://127.0.0.1:7897"
# os.environ["HTTPS_PROXY"] = "http://127.0.0.1:7897"
# 加载API密钥
load_dotenv()
client = create_openrouter_client()


class Match:
    # 文本对比函数，利用OpenAI GPT对比原文和译文
    def compare_texts(self, original_text,rule):
        prompt = f"""
        原文：{original_text}
        
        ##【文件类型】：中译英或者原文译文对照文件，一句（段）原文一句（段）英文。

        ##【审校目标】：
        1) 逐条对照错误类型规则检查译文是否合规；若译文符合规则且无错译（即"译文数值"与"译文修改建议值"相同），严禁将其放入 JSON 列表中输出。
        2) 对照原文检查译文出现错译、漏译、多译的情况。
        3) 严禁四舍五入数值；不改变原文原意；单位与位数严格按规则,尤其数值错误（指换算后的实际数值不一致）。若数值经单位换算后逻辑相等（如 1亿元 = 100 million），则视为正确，严禁输出为错误；
        4) 重点检查标题与编号层级对应与编号顺序问题，如是否跳号、重号、漏译中文编号或者连续编号混用不同符号如1. II.3. ...等。        
        5) 根据原文带"<bold></bold>""<italic></italic>"词语检查翻译后对应的译文有无标签，对应的译文没有输出时也要加入标签进行包裹，加粗范围和原文保持一致；原文同时出现<bold><italic>Text</italic></bold>，译文也统保持一致。若需要修改的内容全部位于标签内，则只修改需要修改的内容并且不需要输出闭合标签。

        ##【错误类型】（逐条检查，不要漏）：{rule}

        ##请严格以json格式按顺序输出检查错误结果，不得输出任何其他内容：
        输出要求：
        1) 输出的`译文数值`、`译文上下文`、`替换锚点`字段的值必须严格按照给出的译文内容进行提取输出，不得做任何篡改；"译文修改建议值"字段必须以【原文】为唯一准则。
        2) 对于漏译、多译导致的字段可能缺少的问题，要给出上下文包含漏译、多译的位置以便进行替换修订，参考输出格式示例"错误编号": "3";禁止输出文档内容没有的值。
        3) 最小拆分原则：每个 JSON 对象仅对应一个译文句子的错误点。
        4) "<bold></bold>""<italic></italic>"带这种标签的代表加粗和斜体，输出的时候保证元素标签闭合
        
        输出格式示例：
            [
              {{
                "错误编号": "1",
                "原文上下文": "包含该数值的原文完整句",
                "译文上下文": "包含该数值的译文完整句",
                "原文数值": "原文提取的原文片段",
                "译文数值": "译文提取的译文错误片段（实际需要修改的错误片段）",
                "替换锚点": "译文中需要被替换的精确字符片段",
                "译文修改建议值": "修正后的译文片段",
                "错误类型": "数值错误",
                "修改理由": "简述违反的具体规则（如：数量级错误）",
                "违反的规则": "规则条款"
              }},
              {{
                "错误编号": "2",
                "原文上下文": "广西川金诺化工的注册资本由11,000万元增加为",
                "译文上下文": "raising its registered capital from RMB11 million to",
                "原文数值": "11,000万",
                "译文数值": "11 million",
                "替换锚点": "11 million",
                "译文修改建议值": "110 million",
                "错误类型": "数值错误",
                "修改理由": "译文数值错误，原文为11,000万元，译文写成11 million（即1,100万），数量级错误。",
                "违反的规则": "规则(一)：是否漏译 / 错译 / 数量级错误"
              }},
              {{
                "错误编号": "3",
                "原文上下文": "第二条  开源软件管理应遵循以下原则：",
                "译文上下文": "Article 2 shall be in compliance with the following principles:",
                "原文数值": "第二条  应",
                "译文数值": "Article 2 shall be",
                "替换锚点": "Article 2 shall be",
                "译文修改建议值": "Article 2 Data demand management shall be",
                "错误类型": "译文漏译开源软件管理",
                "修改理由": "译文漏译。",
                "违反的规则": "译文漏译"
              }},{{
                "错误编号": "4",
                "原文上下文": "第十一条  管理方受理开源软件引入需求",
                "译文上下文": "第十一条  The manager shall accept requests",
                "原文数值": "第十一条",
                "译文数值": "第十一条",
                "替换锚点": "第十一条",
                "译文修改建议值": "Article 11",
                "错误类型": "标题与编号层级",
                "修改理由": "译文中保留了中文编号“第十一条”。",
                "违反的规则": "规则(二)：检查标题与编号是否未翻译,译文中是否出现中文第一章/节/条...等。"
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
                model=get_model_name("google/gemini-3-flash-preview"),
                max_tokens=65536,
                messages=[
                    {"role": "system",
                     "content": "你是中译英及其他语种译文合规审校员，只负责依据要求对译文做错误类型规则和审校目标符合性检查与修改建议，不要自行修正、不要补全缺失信息。"},
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
if __name__ == '__main__':
    # 示例文件路径
    original_path = r"C:\Users\H\Desktop\测试项目\中翻译\测试文件\007 议案19-附件：合规管理及合规文化建设情况工作报告.docx"  # 请替换为原文文件路径
    translated_path = r"C:\Users\H\Desktop\测试项目\中翻译\测试文件\【批註】TP260331011_007 议案19-附件：合规管理及合规文化建设情况工作报告_edited_v2.docx"  # 请替换为译文文件路径

    #处理正文(含脚注/表格/自动编号)
    original_body_text=extract_body_text(original_path)
    translated_body_text=extract_body_text(translated_path)

    print(original_body_text)
    print(translated_body_text)

    # # 实例化对象并进行对比
    matcher = Match()
    #正文对比
    print("======正在检查正文===========")
    body_result = matcher.compare_texts(original_body_text, translated_body_text)
    print(body_result)
