import os
import traceback
from dotenv import load_dotenv
from openai import OpenAI

from parsers.word.body_extractor1 import extract_body_text

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

        try:
            # 使用正确的API调用方式，并启用流式响应
            response = client.chat.completions.create(
                extra_headers={
                    "HTTP-Referer": "<YOUR_SITE_URL>",  # Optional. Site URL for rankings on openrouter.ai.
                    "X-Title": "<YOUR_SITE_NAME>",  # Optional. Site title for rankings on openrouter.ai.
                },
                model="google/gemini-3-flash-preview",  # 使用 OpenAI 的 google/gemini-3-pro-preview 模型
                max_tokens=65536,
                messages=[
                    {"role": "system",
                     "content": "你是中译英及其他语种译文合规审校员，只负责依据要求对译文做错误类型规则符合性检查与修改建议，不要自行修正、不要补全缺失信息。"},
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

    rule="""
        (一) 数值错误：
        - 检查译文全文是否与原文数值不一致，是否漏译 / 错译 / 数量级错误
        - 单位是否一致（%/公式符号/时间日期等数量单位和计量单位）
        - 原文/译文大小写数值、数字（包括罗马数字和阿拉伯数字）必须完全一致
        - 零误差原则，禁止四舍五入
        (二) 标题与编号层级：
        - 一级：I. II. III.（实词首字母大写）
        - 二级：i. ii. iii.（只句首字母大写）
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
    """
    # 示例文件路径
    original_path = r"D:\Users\Administrator\Desktop\项目文件\专检\数值检查V2\测试文件\原文-含不可编辑_【交付翻译】20260319 艾美疫苗2025年度环境、社会及管治（ESG）报告初稿V2.12.docx"  # 请替换为原文文件路径
    translated_path = r"D:\Users\Administrator\Desktop\项目文件\专检\数值检查V2\测试文件\译文-含不可编辑_【交付翻译】20260319 艾美疫苗2025年度环境、社会及管治（ESG）报告初稿V2.12.docx"  # 请替换为译文文件路径

    # 处理正文(含脚注/表格/自动编号)
    original_body_text = extract_body_text(original_path)
    translated_body_text = extract_body_text(translated_path)

    print("======正文===========")
    print(original_body_text)
    print(translated_body_text)

    # # 实例化对象并进行对比
    matcher = Match()
    # 正文对比
    print("======正在检查正文===========")
    body_result = matcher.compare_texts(original_body_text, translated_body_text,rule)




