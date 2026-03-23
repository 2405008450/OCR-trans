import os
import sys
import traceback
from pathlib import Path
from typing import Callable, Optional

from dotenv import load_dotenv

_project_root = Path(__file__).resolve().parents[5]
load_dotenv(_project_root / ".env")

if str(_project_root) not in sys.path:
    sys.path.insert(0, str(_project_root))

from app.service.gemini_service import generate_text
from llm.llm_project.parsers.body_extractor import extract_body_text
from llm.llm_project.parsers.footer_extractor import extract_footers
from llm.llm_project.parsers.header_extractor import extract_headers


class Match:
    def __init__(self, model_name: str = "gemini-3-flash-preview", task_logger: Optional[Callable[[str], None]] = None):
        self.model_name = model_name
        self.task_logger = task_logger

    def _log(self, message: str) -> None:
        if self.task_logger:
            self.task_logger(message)

    def compare_texts(self, original_text, translated_text):
        prompt = f"""
        原文：{original_text}
        译文（待检查）：{translated_text}

        文件类型与场景：
        - 文件类型：一般性文件 / 法律文件（合同、协议、诉讼文书等）
        - 场景：正文 / 表格 / PPT要点 / Excel要点

        【审校目标】：
        1) 逐条对照错误类型规则检查英文译文数值是否与原文中的数值一致；
        2) 不改变原文原意；不四舍五入数值；单位与位数严格按规则；

        【错误类型】（逐条执行，不要漏）：
        （一）数值错误：
        - 检查译文全文是否与原文数值不一致，是否漏译 / 错译 / 数量级错误？
        - 单位是否一致（%/公式符号等数量单位和计量单位）？
        - 原文/译文大小写数值、数字（包括罗马数字和阿拉伯数字）必须完全一致
        - 零误差原则，禁止四舍五入
        （二）标题与编号层级：
        - 禁止序号重复（如出现两个 Article 1）、顺序倒置、禁止跨越式跳号，检查章节序号（Chapter/Section）、条目序号（Article），以及列表项序号（i/ii/iii/1/2/3）的连续性与中英映射准确性
        （三）时间与日期数值不对应与格式规范错误：
        - 年、月、日、周期必须完全对应
        - 12/24 小时制必须时间一致

        请严格以 JSON 格式按顺序输出检查错误结果，不得输出任何其他内容。

        输出要求：
        1) 对于加粗斜体等这类非文本内容的格式问题，在译文修改建议值加上中文括号备注如（加粗）、（斜体等格式），如译文修改建议值：Chapter II（加粗）
        2) 输出的“原文数值”“译文数值”“原文上下文”“译文上下文”“替换锚点”字段的值必须严格按照原内容进行输出，不得做任何改动
        3) 若输入原文和译文有一个为空，则输出空值
        4) 若出现连续错误或者错误距离过近，必须拆分成多个 JSON 对象错误或者整句作为 JSON 对象处理进行处理，保证“译文数值”字段为译文原内容片段。

        输出格式示例：
        [
          {{
            "错误编号": "2",
            "错误类型": "数值错误",
            "原文数值": "本合同一式六份",
            "译文数值": "This Contract is made in 7 counterparts",
            "译文修改建议值": "This Contract is made in 6 counterparts",
            "修改理由": "译文数值与原文不一致",
            "违反的规则": "数值必须与原文保持严格一致",
            "原文上下文": "本合同一式六份",
            "译文上下文": "This Contract is made in 7 counterparts",
            "原文位置": "正文",
            "译文位置": "正文",
            "替换锚点": "made in 7 counterparts"
          }}
        ]
        """

        try:
            route = os.getenv("GEMINI_ROUTE", "google")
            self._log(f"[llm] 开始调用 route={route}, model={self.model_name}")
            full_response = generate_text(
                system_prompt="你是中译英译文数字合规审校员，只负责依据要求对英文译文做数值、编号、日期与时间相关的错误检查与修改建议，不要自行修正、不要补全缺失信息。",
                user_prompt=prompt,
                model=self.model_name,
                route=route,
                temperature=0,
                max_output_tokens=65536,
                log_callback=self.task_logger,
            )
            self._log("[llm] 模型响应已返回")
            return full_response.strip()
        except Exception as exc:
            self._log(f"[llm-error] {type(exc).__name__}: {exc}")
            traceback.print_exc()
            raise RuntimeError(f"数字专检模型调用失败: {exc}") from exc


if __name__ == "__main__":
    original_path = r"C:\path\to\original.docx"
    translated_path = r"C:\path\to\translated.docx"

    original_header_text = extract_headers(original_path)
    translated_header_text = extract_headers(translated_path)
    original_footer_text = extract_footers(original_path)
    translated_footer_text = extract_footers(translated_path)
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

    matcher = Match()
    print("======正在检查正文==========")
    body_result = matcher.compare_texts(original_body_text, translated_body_text)
    print("======正在检查页眉==========")
    header_result = matcher.compare_texts(original_header_text, translated_header_text)
    print("======正在检查页脚==========")
    footer_result = matcher.compare_texts(original_footer_text, translated_footer_text)
    print("================================")
    print(body_result)
    print(header_result)
    print(footer_result)
