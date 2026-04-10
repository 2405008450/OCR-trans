import os
import sys
from pathlib import Path
from openai import OpenAI
from register import registry

from tools.auto_lang_tools import _load_rule_silent, rule_file

PROJECT_ROOT = Path(__file__).resolve().parents[5]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from app.core.config import settings

client = OpenAI(
    api_key=settings.OPENROUTER_API_KEY,
    base_url=settings.OPENROUTER_BASE_URL,
)

@registry.register(
    name="claude",
    description="运行 AI 模型处理复杂的文本任务",
    input_schema={
        "type": "object",
        "properties": {
            "prompt": {
                "type": "string",
                "description": "输入给模型的提示词"
            }
        },
        "required": ["prompt"]
    }
)
def handle_claude(arguments):
    user_input = arguments.get("prompt", "")
    target_path = Path(__file__).parent.parent / "files" / rule_file
    rule_context = _load_rule_silent(str(target_path))

    combined_prompt = (f"""
    请根据规则检查输入文本的标点符号使用是否正确。
    规则:{rule_context}
    输入文本:{user_input}               
    ##【审校目标】：
    1) 逐条对照规则检查输入文本是否合规；
    2) 不改变输入文本内容

    ##请严格以json格式按顺序输出检查错误结果，不得输出任何其他内容：
    输出要求：
    1) 输出的`文本数值`、`文本上下文`、`替换锚点`字段的值必须严格按照输入文本原内容进行输出，不得做任何篡改
    2) 若输入文本完全符合规则，请输出空数组 []
    3) 若出现连续错误或者错误距离过近，必须需拆分成多个json对象错误或者整句作为json对象处理进行处理，
       保证`文本数值`字段为译文原内容片段。如"雨ニモマケズ）風ニモマケズ）"分成"雨ニモマケズ）"和"風ニモマケズ）"两个错误。
    输出格式示例：
              [
                {{
                "错误编号": "1",
                "错误类型": "大小写",
                "文本数值": "“花は桜木、人は武士”",
                "修改建议值": "「花は桜木、人は武士」",
                "修改理由": "引号用于表示引用部分或要求特别注意的词语",
                "违反的规则": "引号用于表示引用部分或要求特别注意的词语",
                "文本上下文": "日本人の生活“花は桜木、人は武士”",
                "文本位置": "正文",
                "替换锚点": "の生活“花は桜木、人は武士”"
                }}
            ]
                       """

                       )
    try:
        res = client.chat.completions.create(
            model=os.getenv("ZHONGFANYI_MODEL_NAME", "google/gemini-3.1-pro-preview"),
            messages=[{"role": "user", "content": combined_prompt}]
        )
        return res.choices[0].message.content
    except Exception as e:
        return f"模型调用错误: {str(e)}"
