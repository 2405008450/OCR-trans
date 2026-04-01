import os
import sys
import time
from pathlib import Path
from typing import Dict, Optional

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


def _resolve_model_name(model_name: Optional[str] = None) -> str:
    candidate = (
        model_name
        or os.getenv("NUMBER_CHECK_SINGLE_MODEL")
        or os.getenv("NUMBER_CHECK_MODEL")
        or "google/gemini-3-flash-preview"
    ).strip()
    return candidate if "/" in candidate else f"google/{candidate}"


class Match:
    """单文件双语对照检查类。"""

    def __init__(self, model_name: Optional[str] = None):
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

    def compare_texts(self, original_text):
        prompt = f"""
        #文件类型：双语对照或者多语种对照文件

        # 角色
        你是一名资深的文件翻译审校专家，精通输入文档中英文以及其他语种文件中的数值、金额、日期、编号层级的一致性核查。

        # 任务
        1.严格对比输入文档中的原文与译文，识别并提取所有数值、单位、编号、日期翻译后译文与原文不一致的错误。
        2.对照原文检查译文出现错译、漏译的情况

        # 审查规则 (执行优先级最高)：
        1. **提取纯粹性：** 提取输入文档原文和译文数值时，严禁拼接、计算或修改数字。仅当译文的数值信息与其对应的原文在含义、精确度或逻辑上不符时，才提取为错误。
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
        - 输入文档：{original_text}
        """

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
                    "content": "你是中译英以及其他语种译文合规审校员，只负责依据要求对译文做错误类型规则符合性检查与修改建议，不要自行修正、不要补全缺失信息。",
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
            if hasattr(chunk, "error") and chunk.error:
                raise RuntimeError(f"SSE stream error: {chunk.error}")

        if not full_response.strip():
            raise RuntimeError("API 返回空内容")
        return full_response.strip()


def _compare_bilingual_with_split(matcher, text, name):
    """对单文件双语对照文本进行分块遍历检查，合并去重结果。"""
    from divide.text_splitter import _count_chars, auto_num_parts, split_text
    from parsers.json.clean_json import parse_json_content

    num_parts = auto_num_parts(text)
    chunks = split_text(text, num_parts) if num_parts > 1 else [text]
    num_parts = len(chunks)

    if num_parts <= 1:
        print(f"  [INFO] {name}文本较短（{_count_chars(text)} 字），直接对比")
        raw = None
        for attempt in range(3):
            try:
                raw = matcher.compare_texts(text)
                break
            except Exception as e:
                if attempt < 2:
                    wait = (attempt + 1) * 10
                    print(f"  [WARN] 调用失败: {e}，{wait}秒后重试...")
                    time.sleep(wait)
                else:
                    print(f"  [ERROR] 调用 API 失败（已重试 2 次）: {e}")
        if raw:
            parsed = parse_json_content(raw)
            return parsed if isinstance(parsed, list) else []
        return []

    print(f"  [INFO] {name}文本较长（{_count_chars(text)} 字），分割为 {num_parts} 块进行对比")

    all_errors = []
    seen_keys = set()

    for i, chunk in enumerate(chunks, 1):
        print(f"\n  --- {name} 第 {i}/{num_parts} 块 ---")
        print(f"      块字数: {_count_chars(chunk)} 字")

        if not chunk.strip():
            print("      [WARN] 块内容为空，跳过")
            continue

        max_retries = 2
        chunk_result = None
        for attempt in range(max_retries + 1):
            try:
                chunk_result = matcher.compare_texts(chunk)
                break
            except Exception as e:
                if attempt < max_retries:
                    wait = (attempt + 1) * 10
                    print(f"      [WARN] 第 {i} 块第 {attempt + 1} 次调用失败: {e}")
                    print(f"      [INFO] 等待 {wait} 秒后重试...")
                    time.sleep(wait)
                else:
                    print(f"      [ERROR] 第 {i} 块调用 API 失败（已重试 {max_retries} 次）: {e}")
                    print("      [WARN] 跳过该块，继续处理剩余块")
                    chunk_result = None

        chunk_count = 0
        if chunk_result:
            parsed = parse_json_content(chunk_result)
            if isinstance(parsed, list):
                chunk_count = len(parsed)
                for item in parsed:
                    dedup_key = (
                        (item.get("译文数值") or "").strip(),
                        (item.get("译文上下文") or "").strip()[:50],
                    )
                    if dedup_key not in seen_keys:
                        seen_keys.add(dedup_key)
                        all_errors.append(item)
                    else:
                        print(f"      [WARN] 去重: '{dedup_key[0]}' (重叠区域)")

        print(f"      [OK] 本块发现 {chunk_count} 个问题，累计 {len(all_errors)} 个")

    print(f"\n  [INFO] {name}合并完成: 共 {len(all_errors)} 个不重复问题")

    for idx, item in enumerate(all_errors, 1):
        item["错误编号"] = str(idx)
    return all_errors


def run_bilingual_comparison(
    file_path,
    output_json_dir=None,
    output_dirs: Optional[Dict[str, str]] = None,
    report_prefix: str = "文本对比结果",
    model_name: Optional[str] = None,
):
    """单文件双语对照检查入口。"""
    specialized_root = Path(__file__).resolve().parents[1]
    specialized_root_str = str(specialized_root)
    if specialized_root_str not in sys.path:
        sys.path.insert(0, specialized_root_str)

    from utils.json_files import write_json_with_timestamp

    path = Path(file_path)
    if not path.exists():
        print(f"[ERROR] 文件不存在: {file_path}")
        return None

    suffix = path.suffix.lower()
    print("\n--- 单文件双语对照检查 ---")
    print(f"[INFO] 文件: {path.name} ({suffix})")

    body, header, footer = "", "", ""
    try:
        if suffix == ".pdf":
            from parsers.pdf.pdf_parser import parse_pdf

            body = parse_pdf(str(path))
        elif suffix == ".xlsx":
            from parsers.excel.excel_parser import parse_excel_with_pandas

            body = parse_excel_with_pandas(str(path))
        elif suffix == ".pptx":
            from parsers.pptx.pptx_parser import parse_pptx

            body = parse_pptx(str(path))
        elif suffix in [".docx", ".doc"]:
            from parsers.word.body_extractor import extract_body_text
            from parsers.word.footer_extractor import extract_footers
            from parsers.word.header_extractor import extract_headers

            body = extract_body_text(str(path))
            header_raw = extract_headers(str(path))
            footer_raw = extract_footers(str(path))
            header = "\n".join(header_raw) if isinstance(header_raw, list) else (header_raw or "")
            footer = "\n".join(footer_raw) if isinstance(footer_raw, list) else (footer_raw or "")
        else:
            print(f"[ERROR] 不支持的文件格式: {suffix}")
            return None
    except Exception as e:
        print(f"[ERROR] 解析失败: {e}")
        return None

    report_output_dirs = {
        "正文": output_json_dir or str(specialized_root / "zhengwen" / "output_json"),
        "页眉": str(specialized_root / "yemei" / "output_json"),
        "页脚": str(specialized_root / "yejiao" / "output_json"),
    }
    if output_dirs:
        for section_name, custom_dir in output_dirs.items():
            if section_name in report_output_dirs and custom_dir:
                report_output_dirs[section_name] = custom_dir

    matcher = Match(model_name=model_name)
    parts = [
        ("正文", body, report_output_dirs["正文"]),
        ("页眉", header, report_output_dirs["页眉"]),
        ("页脚", footer, report_output_dirs["页脚"]),
    ]

    report_paths = {}
    for name, txt, out_dir in parts:
        print(f"\n====== 正在检查{name} ===========")
        if txt and txt.strip():
            try:
                res = _compare_bilingual_with_split(matcher, txt, name)
            except Exception as e:
                print(f"[ERROR] 调用 API 失败: {e}")
                return None
        else:
            res = []
            print(f"[WARN] {name}内容为空，跳过")

        _, json_path = write_json_with_timestamp(res, out_dir, prefix=report_prefix)
        report_paths[name] = json_path

    print("\n[OK] 单文件双语对照检查完成")
    for name, report_path in report_paths.items():
        print(f"  - {name}: {report_path}")

    return report_paths


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="单文件双语对照数值检查工具")
    parser.add_argument(
        "file",
        nargs="?",
        default=r"D:\Users\Administrator\Desktop\项目文件\专检\数值检查V2\ceishiwenjian\更新_20260321 泰格医药2025年度可持续发展报告 V2.33-to翻译（修订标记）_Bilingual.docx",
        help="双语对照文档路径",
    )
    parser.add_argument("--output", "-o", default=None, help="JSON 输出目录（可选）")
    parser.add_argument("--model", default=None, help="模型名称（可选）")
    args = parser.parse_args()

    if not os.path.exists(args.file):
        print(f"[ERROR] 文件不存在: {args.file}")
    else:
        run_bilingual_comparison(
            args.file,
            output_json_dir=args.output,
            model_name=args.model,
        )
