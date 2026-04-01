import os
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[3]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from app.service.gemini_service import generate_text


class Match:
    # 文本对比函数，统一走项目根目录的 Gemini/OpenRouter 封装
    def compare_texts(self, original_text, translated_text, rule):
        prompt = f"""
        原文：{original_text}
        译文（待检查）：{translated_text}

        ##【审校目标】：
        1) 逐条对照错误类型规则检查译文是否合规；
        2) 对照原文检查译文出现错译、漏译的情况
        3) 不改变原文原意；不四舍五入数值；单位与位数严格按规则；
        4) 重点检查标题与编号层级问题

        ##【错误类型】（逐条检查，不要漏）：{rule}

        ##请严格以json格式按顺序输出检查错误结果，不得输出任何其他内容：
        输出要求：
        1) 输出的`原文数值`、`译文数值`、`原文上下文`、`译文上下文`、`替换锚点`字段的值必须严格按照给出的原文内容进行输出，不得做任何篡改；"译文修改建议值"字段必须以【原文】为唯一准则。
        2) 严格保留原文的数值格式，严禁增加、减少或拼接任何数字。如果原文是数字，请将其视为不可修改的实体。
        3) 译文数值字段有单位的数值尽量带单位，如：5 working days
        4) 译文数值字段编号序号等数值符号需带上下文片段进行区分，如：iii. The output part...
        5) 若输入原文和译文有一个为空，则输出空值
        6) 最小拆分原则：多个数值错误，必须拆分为多个 JSON 对象，确保 `替换锚点` 精确到具体的错误片段。
        输出格式示例：

                [
          {{
            "错误编号": "1",
            "错误类型": "数值错误",
            "原文数值": "原文提取的译文片段",
            "译文数值": "译文提取的译文片段",
            "译文修改建议值": "修正后的译文片段",
            "修改理由": "简述违反的具体规则（如：数量级错误）",
            "违反的规则": "规则条款",
            "原文上下文": "包含该数值的原文完整句",
            "译文上下文": "包含该数值的译文完整句",
            "替换锚点": "译文中需要被替换的精确字符片段"
          }},
         {{
            "错误编号": "2",
            "错误类型": "金额四舍五入错误",
            "原文数值": "人民币 10.666 万亿元",
            "译文数值": "RMB10.67 trillion",
            "译文修改建议值": "RMB10,666 billion",
            "修改理由": "原文为 10.666，译文直接进位成 10.67 属于四舍五入。根据规则 g，不得四舍五入。应通过降级单位至 billion 来完整表达数值且满足小数点后不多于两位。",
            "违反的规则": "规则 g：不得四舍五入；规则 a：小数点后不多于两个数字。",
            "原文上下文": "总金额约为人民币 10.666 万亿元。",
            "译文上下文": "The total amount is approximately RMB10.67 trillion.",
            "替换锚点": "RMB10.67 trillion"
         }}
        ]
        """

        full_response = generate_text(
            system_prompt="你是中译英及其他语种译文合规审校员，只负责依据要求对译文做错误类型规则符合性检查与修改建议，不要自行修正、不要补全缺失信息。",
            user_prompt=prompt,
            model="google/gemini-3.1-pro-preview",
            route=os.getenv("GEMINI_ROUTE", "openrouter"),
            temperature=0,
            max_output_tokens=65532,
        )
        print(full_response, end="")
        return full_response.strip()
