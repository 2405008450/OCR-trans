import os
import traceback
from pathlib import Path

from dotenv import load_dotenv
from openai import OpenAI


ROOT_ENV_FILE = Path(__file__).resolve().parents[3] / ".env"
if ROOT_ENV_FILE.exists():
    load_dotenv(ROOT_ENV_FILE, override=False)


def _resolve_api_key() -> str:
    return (
        os.getenv("API_KEY")
        or os.getenv("OPENROUTER_API_KEY")
        or os.getenv("OPENAI_API_KEY")
        or ""
    ).strip()


def _resolve_base_url() -> str:
    return (
        os.getenv("BASE_URL")
        or os.getenv("OPENROUTER_BASE_URL")
        or os.getenv("OPENAI_BASE_URL")
        or "https://openrouter.ai/api/v1"
    ).strip()


def _resolve_model_name(model_name: str | None = None) -> str:
    candidate = (
        model_name
        or os.getenv("NUMBER_CHECK_MODEL")
        or os.getenv("NUMBER_CHECK_SINGLE_MODEL")
        or "google/gemini-3-flash-preview"
    ).strip()
    return candidate if "/" in candidate else f"google/{candidate}"


class Match:
    """双文件对比函数，运行时从根目录环境中读取模型与鉴权配置。"""

    def __init__(self, model_name: str | None = None):
        self.model_name = _resolve_model_name(model_name)

    @staticmethod
    def _build_client() -> OpenAI:
        api_key = _resolve_api_key()
        if not api_key:
            raise ValueError("未配置 API_KEY / OPENROUTER_API_KEY / OPENAI_API_KEY。")
        return OpenAI(
            api_key=api_key,
            base_url=_resolve_base_url(),
        )

    def compare_texts(self, original_text, translated_text):
        prompt = f"""
        # 角色
        你是一名资深的文件翻译审校专家，精通中英文以及其他语种文件中的数值、金额、日期、编号层级的一致性核查。

        # 任务
        1.严格对比【原文】与【译文】，识别并提取所有数值、单位、编号、日期翻译后译文与原文不一致的错误。
        2.对照原文检查译文出现错译、漏译的情况

        # 审查规则 (执行优先级最高)：
        1. **提取纯粹性：** 提取原文和译文数值时，严禁拼接、计算或修改数字。仅当译文的数值信息与其对应的原文在含义、精确度或逻辑上不符时，才提取为错误。
        2. **数值零误差：** 严禁四舍五入，严格检查单位数值一致性（如4亿和40 million是否一致）;但是要注意文本数字兼容：原文“十四”，译文“14”或“14th”应根据语境判定。如果语境是第14条，则一致；除非译文拼写错误或数值变动（如 14 变成 24）才标记为错误。
        3. **编号连续性：** 检查 Article, Section, (i), (1) 等层级是否跳号、重号或者连续编号混用不同符号如1. II.3. ... 。
        4. **排除非数值文本：** 禁止修改任何嵌入在单词内部的类似数字的字符（如不可将 Trustworthy 误改为 Trus2rthy），除非它是明确的编号（如 Section 2）。
        5. **原文主权原则：** "译文修改建议值"字段必须以【原文】为唯一准则。禁止引入原文中不存在的年份、公司名称或背景信息。
        6. **最小拆分原则：** 若一句话内有多个数值错误，必须拆分为多个 JSON 对象，确保 `替换锚点` 精确到具体的错误片段。

        # 工作流程:
        - Step 1: 扫描原文句子，提取所有数值/日期/编号/单位。
        - Step 2: 扫描对应译文句子，提取对应的数值信息。
        - Step 3: 安照审查规则逐一比对。若发现任何翻译后译文与原文中的数值的不一致，按照最小拆分原则标记为错误。

        # 判定标准示例 (Few-shot)：
        1. 允许的转换
        - 原文：第十四条 | 译文：Article 14 -> [一致，无需修改]
        - 原文：三 | 译文：3 -> [一致，无需修改]
        2. 严禁的幻觉修改 (错误示范)
        - 原文：2025 ESG报告 | 译文：2025 ESG Report -> 警告：严禁将其改为 2024 ESG Report。
        - 原文：Trustworthy | 译文：Trustworthy -> 警告：严禁识别为数值错误并改为 Trus2rthy。
        3. 必须拦截的错误 (JSON 输出)
        - 原文：2025 | 译文：2024 -> [数值错误：年份不符]
        - 原文：10.00 | 译文：10 -> [数值错误：精度丢失]
        - 原文：第一节..., 第二节..., 第三节... | 译文：Article 1..., Article 2..., Article 4... -> [层级错误：跳号]

        #输出示例：
        仅输出 JSON 数组，不得包含说明文字。若无错误，则输出 `[]`。格式如下：
        [
          {{
            "错误编号": "1",
            "错误类型": "数值错误/层级错误/日期错误",
            "原文数值": "原文提取的译文片段",
            "译文数值": "译文提取的译文片段",
            "译文修改建议值": "修正后的译文片段",
            "修改理由": "简述违反的具体规则（如：数量级错误/单位不符/跳号）",
            "原文上下文": "包含该数值的原文完整句",
            "译文上下文": "包含该数值的译文完整句",
            "替换锚点": "译文中需要被替换的精确字符片段"
          }}
        ]

        #输入数据:
        - 原文：{original_text}
        - 译文：{translated_text}
        """

        try:
            client = self._build_client()
            response = client.chat.completions.create(
                extra_headers={
                    "HTTP-Referer": "local-debug",
                    "X-Title": "fastapi-llm-demo",
                },
                model=self.model_name,
                max_tokens=65532,
                messages=[
                    {
                        "role": "system",
                        "content": "你是中译英译文合规审校员，只负责依据要求对英文译文做错误类型规则符合性检查与修改建议，不要自行修正、不要补全缺失信息。",
                    },
                    {"role": "user", "content": prompt},
                ],
                temperature=0,
                stream=True,
            )

            full_response = ""
            for chunk in response:
                if chunk.choices and chunk.choices[0].delta.content:
                    message_content = chunk.choices[0].delta.content
                    full_response += message_content
                    print(message_content, end="")
            return full_response.strip()

        except Exception as e:
            print(f"Error occurred: {e}")
            traceback.print_exc()
            return "Error occurred during API call."
