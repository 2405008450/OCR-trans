from __future__ import annotations

import os
import json
from datetime import datetime
from pathlib import Path
from typing import Any, Tuple, Union


def write_txt_with_timestamp(
        data: Any,
        output_dir: Union[str, os.PathLike],
        prefix: str = "文本对比结果",
        as_json_pretty: bool = True,
        tz_aware: bool = False,
) -> Tuple[str, str]:
    """
    将 data 写入 output_dir 目录下的“带时间戳”的 .txt 文件。

    返回：(文件名, 文件完整路径)
    """
    # 1. 确保目录存在
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # 2. 生成时间戳和文件名
    now = datetime.now().astimezone() if tz_aware else datetime.now()
    ts = now.strftime("%Y%m%d_%H%M%S")
    filename = f"{prefix}_{ts}.txt"
    filepath = out_dir / filename

    # 3. 处理数据格式
    if as_json_pretty:
        # 将字典或列表转为美化的 JSON 字符串
        text = json.dumps(data, ensure_ascii=False, indent=4)
    else:
        # 直接转为字符串
        text = str(data)

    # 4. 写入文件 (使用 utf-8 编码防止中文乱码)
    with open(filepath, "w", encoding="utf-8") as f:
        # 写入标题头（可选，保持跟 Word 逻辑一致）
        f.write(f"=== {prefix} ===\n\n")
        f.write(text)

    return filename, str(filepath)


# ===== 用法示例 =====
if __name__ == "__main__":
    sample = {
        "错误编号": "13",
        "错误类型": "数值",
        "原文数值": "2",
        "译文数值": "21",
        "修改理由": "页脚内容翻译错误，原文页码为2，译文为21。",
        "原文位置": "页脚",
    }

    # 指定你的输出目录
    output_path = r"/zhongfanyi/llm/llm_project/rule"

    name, path = write_txt_with_timestamp(sample, output_path)

    print("TXT文件名称:", name)
    print("TXT文件路径:", path)