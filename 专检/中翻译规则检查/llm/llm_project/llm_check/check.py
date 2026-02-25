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
        1) 逐条对照错误类型规则检查英文译文是否合规；
        2) 不改变原文原意；不四舍五入数值；单位与位数严格按规则；

        【错误类型】（逐条执行，不要漏）：
        (一) 数值错误：
        - 检查译文全文是否与原文数值不一致，是否漏译 / 错译 / 数量级错误
        - 单位是否一致（%/公式符号/时间日期等数量单位和计量单位）
        - 原文/译文大小写数值、数字（包括罗马数字和阿拉伯数字）必须完全一致
        - 零误差原则，禁止四舍五入
        (二) 标题与编号层级：
        - 一级：I. II. III.（加粗；实词首字母大写）
        - 二级：i. ii. iii.（加粗；只句首字母大写）
        - 三级：1. 2. 3. / (1) (2) (3)
        - 不得沿用①②③，必须改为1) 2) 3)
        - 禁止序号重复（如出现两个 Article 1）、顺序倒置、禁止跨越式跳号,检查章节序号（Chapter/Section）、条目序号（Article）、以及列表项序号（i/ii/iii/1/2/3）的连续性与中英映射准确性.
        (三) 大小写：
        - 正文中加书名号的法律法规/议案名称等：英文需全部实词首字母大写
        - 部门名称：实词首字母大写
        - 表格表头：如非专有名词，仅第一个实词首字母大写
        - 表格名称：实词首字母大写
        - 职位：全部实词首字母大写（如 Chief Risk Officer）
        (四) 枚举结构：
        “一是…二是…三是…” → First, ...; second, ...; third, ...（不要用 firstly/secondly/thirdly）
        (五) 金额/数字/单位类数值：（million、billion等）
        - 小数点后不多于2位
        - 能用最大金额单位不用小一级；但若用最大单位导致金额变成小数且触发规则，则单位需向下降一级
        - 不得四舍五入
        - 百万及以上整数一般不得自行补零；百万以下整数不使用百万及以上单位
        - 不用 thousand 表示“千”
        (六) 日期/时间/数字拼写：
        - 日期用 “Month D, Y”（如 October 12, 2013）；“Month Y”中间不加逗号
        - am/pm 需要加点（a.m./p.m.）
        - 数字：1–10用英文单词；11+用阿拉伯数字；但若1–10与11+并列则全部用数字
        (七) 银行分支行译法：
        - 仅支行名称：按中文罗列
        - “XX分行XX部”：用“XX Department, XXXX (拼音) Branch, Bank of China”
        - “省+市+XX支行”：省市放后面（Sub-branch in City, Province），不能直接放前面
        - 原则拼音；但“XX路支行/XX广场支行/第X支行/具体地点”可意译
        (八) 缩写与冠词：
        - 银行缩写（ICBC/ABC/BOC等）前不加the；央行/部委等按规则加the
        - 缩写首次出现：全称（"缩写"），后续只用缩写
        (九) 标点与空格：
        - 英文译文不得出现全角中文符号（括号、逗号、顿号、书名号等）
        - & 前后各空一格（除固定搭配如R&D）；/ 前后不空格
        - 电话括号用英文括号且与数字间空格按规则
        - 公式：A = B ÷ (1 + C)，运算符与括号前后空一格；除A首词首字母大写外其余小写（专名除外）
        (十) 无主语句：
        - 标题可直接动词开头（如 Further improve...）
        - 正文无主语句：必须补主语或改被动；整段无主语时句型要多样，避免机械重复（连续10句不超过4句同句型）
        - 表格/Excel/PPT若为意见/规划/目标/规则等祈使句内容，可直接动词原形开头；若明确过去时应转被动
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
                    "错误类型": "大小写",
                    "原文数值": "《哈利·波特与魔法石》",
                    "译文数值": "harry potter and the philosopher's stone",
                    "译文修改建议值": "Harry Potter and the Philosopher's Stone(斜体)",
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



