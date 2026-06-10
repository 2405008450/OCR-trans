import os
import traceback
from dotenv import load_dotenv
from llm_check.openrouter_config import create_openrouter_client, get_model_name

from parsers.txt.txt_parser import parse_txt
from parsers.word.body_extractor import extract_body_text

# 配置代理地址
# os.environ["HTTP_PROXY"] = "http://127.0.0.1:7897"
# os.environ["HTTPS_PROXY"] = "http://127.0.0.1:7897"
# 加载API密钥
load_dotenv()
client = create_openrouter_client()


class Match:
    # 文本对比函数，利用OpenAI GPT对比原文和译文
    def compare_texts(self, original_text, translated_text, rule):
        prompt = f"""
        原文：{original_text}
        译文（待检查）：{translated_text}

        ##【审校目标】：
        1) 逐条对照错误类型规则检查译文是否合规；若译文符合规则且无错译，严禁将其放入 JSON 列表中输出。
        2) 对照原文检查译文出现错译、漏译、多译的情况。
        3) 严禁四舍五入数值；不改变原文原意；不单位与位数严格按规则；如果原文是数字，请将其视为不可修改的实体;尤其数值错误（指换算后的实际数值不一致）。若数值经单位换算后逻辑相等（如 1亿元 = 100 million），则视为正确，严禁输出为错误。
        4) 重点检查标题与编号层级对应与编号顺序问题，如是否跳号、重号、漏译中文编号或者连续编号混用不同符号如1. II.3. ...等。
        5) 根据原文带"<bold></bold>""<italic></italic>"词语检查翻译后对应的译文有无标签，对应的译文没有输出时也要加入标签进行包裹，加粗范围和原文保持一致；原文同时出现<bold><italic>Text</italic></bold>，译文也统保持一致。若需要修改的内容全部位于标签内，则只修改需要修改的内容并且不需要输出闭合标签。
        
        ## 工作流程:
        - Step 1: 逐句扫描原文句子并找到对应译文句子,对照错误类型规则和【审校目标】根据原文句子检查对应译文句子是否合规。
        - Step 2: 将每项检查出结果判定为"正确"或者"错误"。
        - Step 3: 将 Step 2 中判定为"错误"的项输出为 JSON,且每个 JSON 对象仅对应一个译文句子的错误点。判定为"正确"的项丢弃，不得出现在最终输出中。        

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
                "原文上下文": "包含原文数值的原文上下文",
                "译文上下文": "包含译文数值的译文上下文",
                "原文数值": "原文提取的原文片段",
                "译文数值": "译文提取的译文错误片段（实际需要修改的错误片段）",
                "替换锚点": "译文中需要被替换的精确字符片段",
                "译文修改建议值": "修正后的译文片段，必须与'译文数值'在语境中完全对等，确保直接替换锚点后，译文上下文在语法、空格和单位上完全正确。例如：若锚点为'1 million'，建议值应为'10 million'，严禁只提供数字'10'。",
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
                model=get_model_name("google/gemini-3.1-pro-preview"),
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


# # 主程序
if __name__ == "__main__":
#
#     rule="""
#         (一) 数值错误：
#         - 检查译文全文是否与原文数值不一致，是否漏译 / 错译 / 数量级错误
#         - 单位是否一致（%/公式符号/时间日期等数量单位和计量单位）
#         - 原文/译文大小写数值、数字（包括罗马数字和阿拉伯数字）必须完全一致
#         - 零误差原则，禁止四舍五入
#         (二) 标题与编号层级：
#         - 一级：I. II. III.（实词首字母大写）
#         - 二级：i. ii. iii.（只句首字母大写）
#         - 三级：1. 2. 3. / (1) (2) (3)
#         - 不得沿用①②③，必须改为1) 2) 3)
#         - 禁止序号重复（如出现两个 Article 1）、顺序倒置、禁止跨越式跳号,检查章节序号（Chapter/Section）、条目序号（Article）、以及列表项序号（i/ii/iii/1/2/3）的连续性与中英映射准确性.
#         (三) 大小写：
#         - 正文中加书名号的法律法规/议案名称等：英文需全部实词首字母大写
#         - 部门名称：实词首字母大写
#         - 表格表头：如非专有名词，仅第一个实词首字母大写
#         - 表格名称：实词首字母大写
#         - 职位：全部实词首字母大写（如 Chief Risk Officer）
#         (四) 枚举结构：
#         “一是…二是…三是…” → First, ...; second, ...; third, ...（不要用 firstly/secondly/thirdly）
#         (五) 金额/数字/单位类数值：（million、billion等）
#         - 小数点后不多于2位
#         - 能用最大金额单位不用小一级；但若用最大单位导致金额变成小数且触发规则，则单位需向下降一级
#         - 不得四舍五入
#         - 百万及以上整数一般不得自行补零；百万以下整数不使用百万及以上单位
#         - 不用 thousand 表示“千”
#         (六) 日期/时间/数字拼写：
#         - 日期用 “Month D, Y”（如 October 12, 2013）；“Month Y”中间不加逗号
#         - am/pm 需要加点（a.m./p.m.）
#         - 数字：1–10用英文单词；11+用阿拉伯数字；但若1–10与11+并列则全部用数字
#         (七) 银行分支行译法：
#         - 仅支行名称：按中文罗列
#         - “XX分行XX部”：用“XX Department, XXXX (拼音) Branch, Bank of China”
#         - “省+市+XX支行”：省市放后面（Sub-branch in City, Province），不能直接放前面
#         - 原则拼音；但“XX路支行/XX广场支行/第X支行/具体地点”可意译
#         (八) 缩写与冠词：
#         - 银行缩写（ICBC/ABC/BOC等）前不加the；央行/部委等按规则加the
#         - 缩写首次出现：全称（"缩写"），后续只用缩写
#         (九) 标点与空格：
#         - 英文译文不得出现全角中文符号（括号、逗号、顿号、书名号等）
#         - & 前后各空一格（除固定搭配如R&D）；/ 前后不空格
#         - 电话括号用英文括号且与数字间空格按规则
#         - 公式：A = B ÷ (1 + C)，运算符与括号前后空一格；除A首词首字母大写外其余小写（专名除外）
#         (十) 无主语句：
#         - 标题可直接动词开头（如 Further improve...）
#         - 正文无主语句：必须补主语或改被动；整段无主语时句型要多样，避免机械重复（连续10句不超过4句同句型）
#         - 表格/Excel/PPT若为意见/规划/目标/规则等祈使句内容，可直接动词原形开头；若明确过去时应转被动
#     """
    # 示例文件路径
    original_path = r"C:\Users\H\Desktop\测试项目\中翻译\测试文件\原文—含不可编辑_01 白云电器2025年报1-管理层讨论与分析.docx"  # 请替换为原文文件路径
    translated_path = r"C:\Users\H\Desktop\测试项目\中翻译\测试文件\含不可编辑_01 白云电器2025年报1-管理层讨论与分析.docx"  # 请替换为译文文件路径
    original_body_text=extract_body_text(original_path)
    translated_body_text=extract_body_text(translated_path)
    print("======正文===========")
    print(original_body_text)

    # # 实例化对象并进行对比
    matcher = Match()
    print("======正在检查页脚===========")
    footer_result = matcher.compare_texts(original_body_text,translated_body_text)
    print("================================")




