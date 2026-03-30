import os
from pathlib import Path

from dotenv import load_dotenv

from app.service.gemini_service import generate_text
from llm.llm_project.parsers.word.footer_extractor import extract_footers
from llm.llm_project.parsers.word.header_extractor import extract_headers

_ENV_PATH = Path(__file__).with_name('.env')
load_dotenv(_ENV_PATH)


def _resolve_model_name() -> str:
    return (os.getenv("NUMBER_CHECK_MODEL") or "gemini-3.1-pro-preview").strip()


class Match:
    def compare_texts(self, original_text, translated_text):
        prompt = f"""
        # 角色
        你是一名资深的文件翻译审校专家，精通中英文件以及其他语种文件中的数值、金额、日期、编号层级一致性核查。

        # 任务
        1. 严格对比【原文】与【译文】，识别并提取所有数值、单位、编号、日期翻译后与原文不一致的错误。
        2. 对照原文检查译文中的错译、漏译情况。

        # 审查规则
        1. 严禁四舍五入，单位和数值必须精确对应。
        2. 原文数值格式必须保留，例如 10.00 不能改成 10。
        3. 编号层级必须连续且格式一致。
        4. 多个错误必须按最小粒度拆分输出。

        # 输出格式
        仅输出 JSON 数组，不得包含解释文字、Markdown 代码块或其他说明。

        必须严格使用下面这些字段名，不允许改成其他英文键名：
        - 错误编号
        - 错误类型
        - 原文数值
        - 译文数值
        - 译文修改建议值
        - 修改理由
        - 原文上下文
        - 译文上下文
        - 替换锚点

        关键要求：
        1. `译文数值` 必须是【译文中实际存在、可被程序直接定位替换】的错误片段。
        2. `译文修改建议值` 必须是【最终可直接替换进文档】的修正后译文片段，不得写成“删除……”“将……改为……”这类自然语言说明。
        3. 如果需要删除多余内容，`译文修改建议值` 必须直接给出删除后的完整正确译文片段。
        4. `原文上下文`、`译文上下文` 应尽量给出包含错误片段的完整句子或短段落。
        5. `替换锚点` 优先填写最容易在文档中唯一定位该错误的译文锚点；没有更好锚点时可与 `译文数值` 相同。
        6. 严禁输出以下字段：`error_type`、`original_text`、`translated_text`、`correction_suggestion`。

        输出示例：
        [
          {{
            "错误编号": "1",
            "错误类型": "日期不一致",
            "原文数值": "自2027年10月XX日起施行",
            "译文数值": "October XX, 2025",
            "译文修改建议值": "October XX, 2027",
            "修改理由": "原文年份为2027，译文误写为2025。",
            "原文上下文": "本办法自2027年10月XX日起施行。",
            "译文上下文": "These Measures shall come into force as of October XX, 2025.",
            "替换锚点": "October XX, 2025"
          }}
        ]

        # 输入数据
        - 原文：{original_text}
        - 译文：{translated_text}
        """

        full_response = generate_text(
            system_prompt="你是中译英以及其他语种译文合规审校员，只负责依据要求输出错误检查与修改建议。",
            user_prompt=prompt,
            model=_resolve_model_name(),
            route=os.getenv("GEMINI_ROUTE", "openrouter"),
            temperature=0,
            max_output_tokens=65532,
        )
        print(full_response, end="")
        return full_response.strip()


if __name__ == "__main__":
    original_path = r"C:\Users\Administrator\Desktop\项目文件\专检\数值检查\测试文件\原文.docx"
    translated_path = r"C:\Users\Administrator\Desktop\项目文件\专检\数值检查\测试文件\译文.docx"
    original_header_text = extract_headers(original_path)
    translated_header_text = extract_headers(translated_path)
    original_footer_text = extract_footers(original_path)
    translated_footer_text = extract_footers(translated_path)
    print(original_header_text, translated_header_text, original_footer_text, translated_footer_text)
