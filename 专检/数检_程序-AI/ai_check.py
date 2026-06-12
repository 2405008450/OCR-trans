"""
AI 检查模块

调用 extract_values.py 的数据接口完成 AI 回填，
调用 report_generator.py 生成最终报告。
"""
import os
import re
import json
import pandas as pd
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()
client = OpenAI(api_key=os.getenv("API_KEY"), base_url=os.getenv("BASE_URL"))


# ─────────────────────────────────────────
# LLM 调用
# ─────────────────────────────────────────

def _safe_parse_json(content: str) -> dict:
    try:
        content = re.sub(r"```json|```", "", content).strip()
        match = re.search(r"\{.*\}", content, re.S)
        return json.loads(match.group() if match else content)
    except Exception:
        return {"is_correct": False, "errors": [], "source_issues": []}


_ERROR_SCHEMA = """{
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
}"""

_SOURCE_ISSUE_SCHEMA = """{
  "原文数值": "原文中存在问题的数值片段",
  "原文上下文": "包含该数值的原文上下文"
}"""


def llm_check_block(src_text: str, tgt_text: str) -> dict:
    prompt = f"""你是翻译数值审校专家。检查整块文本中的所有数值问题。

区分两类情况：
- errors：译文数值与原文不符（翻译错误），【需要修改译文】。
- source_issues：译文忠实还原了原文，但原文数值本身存在逻辑问题，【不需要修改译文】。

输出JSON：
{{
  "is_correct": true,
  "errors": [{_ERROR_SCHEMA}],
  "source_issues": [{_SOURCE_ISSUE_SCHEMA}]
}}

原文：
{src_text}

译文：
{tgt_text}
"""
    resp = client.chat.completions.create(
        model="google/gemini-3.5-flash",
        messages=[{"role": "system", "content": "只输出JSON"},
                  {"role": "user", "content": prompt}],
        temperature=0,
    )
    result = _safe_parse_json(resp.choices[0].message.content)

    if not result.get("errors"):
        result["is_correct"] = True
    result.setdefault("source_issues", [])

    return result
