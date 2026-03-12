from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any, Tuple, Union


def write_json_with_timestamp(
    data: Any,
    output_dir: Union[str, os.PathLike],
    prefix: str = "文本对比结果",
    ensure_ascii: bool = False,
    indent: int = 2,
    tz_aware: bool = False,
) -> Tuple[str, str]:
    """
    将 data 写入 output_dir 目录下的“带时间戳”的 .json 文件，并返回：
    (文件名, 文件完整路径)

    文件名示例：
      文本对比结果_20260206_153012.json

    参数：
    - data: 任意可 JSON 序列化对象（dict/list/str/number...）
    - output_dir: 输出目录（不存在会自动创建）
    - prefix: 文件名前缀
    - ensure_ascii: 是否转义中文（默认 False，保留中文）
    - indent: 缩进（默认 2）
    - tz_aware: 是否使用本地时区时间（默认 False 使用系统本地；True 会用 datetime.now().astimezone()）
    """
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    now = datetime.now().astimezone() if tz_aware else datetime.now()
    ts = now.strftime("%Y%m%d_%H%M%S")

    filename = f"{prefix}_{ts}.json"
    filepath = out_dir / filename

    with filepath.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=ensure_ascii, indent=indent)

    return filename, str(filepath)


# ===== 用法示例 =====
if __name__ == "__main__":
    sample = [
        {
            "错误编号": "13",
            "错误类型": "数值",
            "原文数值": "2",
            "译文数值": "21",
            "译文修改建议值": "2",
            "修改理由": "页脚内容翻译错误，原文页码为2，译文为21。",
            "违反的规则": "(一) 总体: 准确。",
            "原文上下文": "=== 页脚内容 === 2",
            "译文上下文": "=== 页脚内容 === 21",
            "原文位置": "页脚",
            "译文位置": "Footer",
            "替换锚点": "21",
        }
    ]

    name, path = write_json_with_timestamp(sample, r"C:\Users\Administrator\Desktop\project\llm\zhongfanyi\output_json")
    print("json文件名称:", name)
    print("json文件路径:", path)
