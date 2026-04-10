"""
自动语种工具生成器
扫描 files 目录下的规则文件，自动为每个语种注册标点检查工具。
新增语种只需在 files 目录下添加对应的 .txt 规则文件即可。
"""
import os
import sys
from pathlib import Path
from openai import OpenAI
from register import registry
# from utils.txt.txt_parser import parse_txt

PROJECT_ROOT = Path(__file__).resolve().parents[5]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from app.core.config import settings

client = OpenAI(
    api_key=settings.OPENROUTER_API_KEY,
    base_url=settings.OPENROUTER_BASE_URL,
)

# 语种名称映射：文件名 -> (工具名, 描述, 函数名后缀)
LANG_MAP = {
    "arabic":      ("Arabic_check",       "检查阿拉伯语符号规则"),
    "chinese":     ("Chinese_check",      "检查中文符号规则"),
    "danish":      ("Danish_check",       "检查丹麦语符号规则"),
    "dutch":       ("Dutch_check",        "检查荷兰语符号规则"),
    "english":     ("English_check",      "检查英文文符号规则"),
    "finnish":     ("Finnish_check",      "检查芬兰语符号规则"),
    "french":      ("French_check",       "检查法语符号规则"),
    "german":      ("German_check",       "检查德语符号规则"),
    "hebrew":      ("Hebrew_check",       "检查希伯来语符号规则"),
    "indonesian":  ("Indonesian_check",   "检查印尼语符号规则"),
    "italian":     ("Italian_check",      "检查意大利语符号规则"),
    "Japanese":    ("Japanese_check",     "检查日语符号规则"),
    "korean":      ("Korean_check",       "检查韩语符号规则"),
    "malay":       ("Malay_check",        "检查马来语符号规则"),
    "norwegian":   ("Norwegian_check",    "检查挪威语符号规则"),
    "polish":      ("Polish_check",       "检查波兰语符号规则"),
    "portuguese":  ("Portuguese_check",   "检查葡萄牙语符号规则"),
    "russian":     ("Russian_check",      "检查俄语符号规则"),
    "spanish":     ("Spanish_check",      "检查西班牙语符号规则"),
    "swedish":     ("Swedish_check",      "检查瑞典语符号规则"),
    "thai":        ("Thai_check",         "检查泰语符号规则"),
    "turkish":     ("Turkish_check",      "检查土耳其语符号规则"),
    "vietnamese":  ("Vietnamese_check",   "检查越南语符号规则"),
}

def _load_rule_silent(rule_path: str) -> str:
    """静默读取规则文件，不打印日志"""
    from pathlib import Path as _P
    p = _P(rule_path)
    if not p.exists():
        return ""
    try:
        try:
            with open(p, "r", encoding="utf-8") as f:
                return f.read().strip()
        except UnicodeDecodeError:
            with open(p, "r", encoding="gbk") as f:
                return f.read().strip()
    except Exception:
        return ""

def _make_checker(rule_file: str, lang_name: str):
    """工厂函数：为指定语种创建检查函数"""
    def checker(arguments):
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
            raise f"模型调用错误: {str(e)}"
    
    checker.__name__ = f"{lang_name}_check"
    return checker

# 自动注册：扫描 files 目录，跳过已有独立工具的语种
FILES_DIR = Path(__file__).parent.parent / "files"
# 已有独立工具文件的语种（不重复注册）
SKIP_LANGS = {""}

for rule_file in sorted(FILES_DIR.glob("*.txt")):
    lang_key = rule_file.stem.lower()
    
    if lang_key in SKIP_LANGS:
        continue
    
    if lang_key not in LANG_MAP:
        continue
    
    tool_name, description = LANG_MAP[lang_key]
    
    # 检查是否已被注册（避免重复）
    if any(t["name"] == tool_name for t in registry.tools_metadata):
        continue
    
    checker_fn = _make_checker(rule_file.name, lang_key)
    
    registry.register(
        name=tool_name,
        description=description,
        input_schema={
            "type": "object",
            "properties": {
                "prompt": {
                    "type": "string",
                    "description": f"请输入要检查的{description[2:]}文本"
                }
            },
            "required": ["prompt"]
        }
    )(checker_fn)
