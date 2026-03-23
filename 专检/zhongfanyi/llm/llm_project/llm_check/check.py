import os
import traceback
from dotenv import load_dotenv

from app.service.gemini_service import generate_text
from zhongfanyi.llm.llm_project.llm_check.rule_generation import rule
from zhongfanyi.llm.llm_project.parsers.word.body_extractor import extract_body_text
from zhongfanyi.llm.llm_project.parsers.word.footer_extractor import extract_footers
from zhongfanyi.llm.llm_project.parsers.word.header_extractor import extract_headers

# 加载API密钥：优先从项目根目录 .env 读取（与 FastAPI 主应用一致）
# 专检/zhongfanyi/llm/llm_project/llm_check -> 项目根 fastapi-llm-demo 需 6 层 ..
_project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", ".."))
_env_path = os.path.join(_project_root, ".env")
if os.path.isfile(_env_path):
    load_dotenv(_env_path)

# 兼容 OPENROUTER_* 与旧版 API_KEY/BASE_URL


class Match:
    # 文本对比函数，利用OpenAI GPT对比原文和译文
    def compare_texts(self, original_text, translated_text, rule):
        prompt = f"""
        原文：{original_text}
        译文（待检查）：{translated_text}

        ##文件类型与场景：
        - 文件类型：一般性文件 / 法律文件（合同、协议、诉讼文书等）
        - 场景：正文 / 表格 / PPT要点 / Excel要点

        ##【审校目标】：
        1) 逐条对照错误类型规则检查英文译文是否合规；
        2) 不改变原文原意；不四舍五入数值；单位与位数严格按规则；
        3) 重点检查标题与编号层级问题

        ##【错误类型】（逐条检查，不要漏）：{rule}
        
        ##请严格以json格式按顺序输出检查错误结果，不得输出任何其他内容：
        输出要求：
        1) 输出的`原文数值`、`译文数值`、`原文上下文`、`译文上下文`、`替换锚点`字段的值必须严格按照给出的原内容进行输出，不得做任何篡改
        2) 严格保留原文的数值格式，严禁增加、减少或拼接任何数字。如果原文是数字，请将其视为不可修改的实体。
        3) 译文数值字段有单位的数值尽量带单位，如：5 working days
        4) 译文数值字段编号序号等数值符号需带上下文片段进行区分，如：iii. The output part
        5) 若输入原文和译文有一个为空，则输出空值
        6) 若出现连续错误或者错误距离过近，必须需拆分成多个json对象错误或者整句作为json对象处理进行处理，保证`译文数值`字段为译文原内容片段。如This Contract is made in 7 counterparts, with 2 copies held by Party A and three copies held by Party B.中若7 3 three数值错误，则输出译文数值7 译文数值2 译文数值three 译文建议修改值6 译文建议修改值5 译文建议修改值five ，或者译文数值：This Contract is made in 7 counterparts, with 2 copies held by Party A and three copies held by Party B.译文建议修改值：This Contract is made in 6 counterparts, with 5 copies held by Party A and five copies held by Party B.
        7) 严禁出现译文数值：7/3/three 译文建议修改值：6/5/five这种多个错误在一起的情况
        输出格式示例：
                  [
                    {{
                    "错误编号": "2",
                    "错误类型": "大小写",
                    "原文数值": "《哈利·波特与魔法石》",
                    "译文数值": "harry potter and the philosopher's stone",
                    "译文修改建议值": "Harry Potter and the Philosopher's Stone",
                    "修改理由": "违反了正文中加书名号的法律法规/议案名称等：英文需全部实词首字母大写",
                    "违反的规则": "正文中加书名号的法律法规/议案名称等：英文需全部实词首字母大写",
                    "原文上下文": "我最近读了《哈利·波特与魔法石》，很喜欢。",
                    "译文上下文": "I recently read harry potter and the philosopher's stone and enjoyed it.",
                    "原文位置": "正文",
                    "译文位置": "正文",
                    "替换锚点": "Harry Potter and the Philosopher's Stone"
                    }}
                ]
        """

        try:
            full_response = generate_text(
                system_prompt="你是中译英译文合规审校员，只负责依据要求对英文译文做错误类型规则符合性检查与修改建议，不要自行修正、不要补全缺失信息。",
                user_prompt=prompt,
                model="google/gemini-3.1-pro-preview",
                route=os.getenv("GEMINI_ROUTE", "google"),
                temperature=0,
                max_output_tokens=65532,
            )
            print(full_response, end="")
            return full_response.strip()
        except Exception as e:
            raise e


# 主程序
if __name__ == "__main__":
    # 示例文件路径
    original_path = r"C:\Users\Administrator\Desktop\项目文件\专检\数值检查\测试文件\原文-中翻译规则测试文件.docx"  # 请替换为原文文件路径
    translated_path = r"C:\Users\Administrator\Desktop\项目文件\专检\数值检查\测试文件\译文-中翻译规则测试文件.docx"  # 请替换为译文文件路径

    #处理页眉
    original_header_text = extract_headers(original_path)
    translated_header_text = extract_headers(translated_path)
    #处理页脚
    original_footer_text = extract_footers(original_path)
    translated_footer_text = extract_footers(translated_path)
    #处理正文(含脚注/表格/自动编号)
    original_body_text = extract_body_text(original_path)
    translated_body_text = extract_body_text(translated_path)
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
