import os
import sys
from pathlib import Path

from llm_check.rule_generation import rule
from parsers.txt.txt_parser import parse_txt
from parsers.word.body_extractor import extract_body_text
from parsers.word.footer_extractor import extract_footers
from parsers.word.header_extractor import extract_headers

PROJECT_ROOT = Path(__file__).resolve().parents[3]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from app.service.gemini_service import generate_text


class Match:
    # 文本对比函数，利用OpenAI GPT对比原文和译文
    def compare_texts(self, original_text,rule):
        prompt = f"""
        原文：{original_text}
        
        ##【文件类型】：中译英或者原文译文对照文件，一句（段）原文一句（段）英文。
        
        ##【任务】：你是中译英及其他语种译文合规审校员，只负责依据要求对文件中的译文做错误类型规则符合性检查与修改建议，不要自行修正、不要补全缺失信息

        ##【审校目标】：
        1) 逐条对照错误类型规则检查译文是否合规；
        2) 对照原文检查译文出现错译、漏译的情况，尤其数值错误（仅结果错误，不包括数值表达方式错误）与原文不一致的情况；
        3) 不改变原文原意；不四舍五入数值；单位与位数严格按规则；
        4) 重点检查标题与编号层级对应与编号顺序问题，如是否跳号、重号、漏译中文编号或者连续编号混用不同符号如1. II.3. ... 。

        ##【错误类型】（逐条检查，不要漏）：{rule}

        ##请严格以json格式按顺序输出检查错误结果，不得输出任何其他内容：
        输出要求：
        1) 输出的`原文数值`、`译文数值`、`原文上下文`、`译文上下文`、`替换锚点`字段的值必须严格按照给出的原内容进行输出，不得做任何篡改；"译文修改建议值"字段必须以【原文】为唯一准则。
        2) 严格保留原文的数值格式，严禁增加、减少或拼接任何数字。如果原文是数字，请将其视为不可修改的实体。
        3) 译文数值字段有单位的数值尽量带单位，如：5 working days
        4) 译文数值字段编号序号等数值符号需带上下文片段进行区分，如：iii. The output part
        5) 若输入原文和译文有一个为空，则输出空值
        6) 若出现连续错误或者错误距离过近，必须需拆分成多个json对象错误或者整句作为json对象处理进行处理，保证`译文数值`字段为译文原内容片段。 
        输出格式示例：
            [
              {{
                "错误编号": "1",
                "原文上下文": "包含该数值的原文完整句",
                "译文上下文": "包含该数值的译文完整句",
                "原文数值": "原文提取的原文片段",
                "译文数值": "译文提取的译文错误片段",
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
        """

        full_response = generate_text(
            system_prompt="你是中译英及其他语种译文合规审校员，只负责依据要求对译文做错误类型规则符合性检查与修改建议，不要自行修正、不要补全缺失信息。",
            user_prompt=prompt,
            model="google/gemini-3-flash-preview",
            route=os.getenv("GEMINI_ROUTE", "openrouter"),
            temperature=0,
            max_output_tokens=65536,
        )
        print(full_response, end="")
        return full_response.strip()

if __name__ == "__main__":
    # 示例文件路径
    original_path = r"/zhongfanyi_spilite\测试文件\更新_20260321 泰格医药2025年度可持续发展报告 V2.33-to翻译（修订标记）_Bilingual_corrected_20260325_111353.docx"  # 请替换为原文文件路径
    rule_path= r"/zhongfanyi_spilite\llm\llm_project\rule\自定义规则.txt"
    #处理正文(含脚注/表格/自动编号)
    rule_text=parse_txt(rule_path)
    original_body_text=extract_body_text(original_path)
    print("======正文===========")
    print(original_body_text)

    # # 实例化对象并进行对比
    matcher = Match()
    print("======正在检查页脚===========")
    footer_result = matcher.compare_texts(original_body_text,rule_text)
    print("================================")

