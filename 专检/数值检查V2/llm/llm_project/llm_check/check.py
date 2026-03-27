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
        仅输出 JSON 数组，不得包含解释文字。

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
