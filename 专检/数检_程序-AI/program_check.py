"""
程序检查 + AI检查 合并流程

步骤：
  1. run_number_check()  → List[Dict]  规则检查 JSON
  2. LLM 对错误行复核   → ai_map
  3. merge_ai_results()  → List[Dict]  合并 JSON
  4. generate_combined_report() 根据 JSON 生成 Excel
"""
import os
import re
import json
from dotenv import load_dotenv
from openai import OpenAI

from extract_values import run_number_check, merge_ai_results
from report_generator import generate_combined_report

load_dotenv()
client = OpenAI(api_key=os.getenv("API_KEY"), base_url=os.getenv("BASE_URL"))


_ERROR_SCHEMA = """{
  "错误编号": "1",
  "原文上下文": "包含原文数值的原文上下文",
  "译文上下文": "包含译文数值的译文上下文",
  "原文数值": "原文提取的原文片段",
  "译文数值": "译文提取的译文错误片段（实际需要修改的错误片段）",
  "替换锚点": "译文中需要被替换的精确字符片段",
  "译文修改建议值": "修正后的译文片段，必须与'译文数值'在语境中完全对等，确保直接替换锚点后，译文上下文在语法、空格和单位上完全正确。例如：若锚点为'1 million'，建议值应为'10 million'，严禁只提供数字'10'。",
  "is_source_consistent": "true或false — 译文数值是否与原文数值一致（即译文是否忠实还原了原文数值）。true表示译文忠实翻译了原文但原文本身可能有问题，false表示译文确实与原文不一致属于翻译错误",
  "错误类型": "数值错误",
  "修改理由": "简述违反的具体规则（如：数量级错误）",
  "违反的规则": "规则条款"
}"""

_SOURCE_ISSUE_SCHEMA = """{
  "原文数值": "原文中存在问题的数值片段",
  "原文上下文": "包含该数值的原文上下文"
}"""


def _safe_parse_json(content: str):
    """
    解析模型返回的 JSON 数组。
    返回 (parsed_list, status)：
      "ok"           — 正常解析（空列表也是 ok，表示模型认为无问题）
      "parse_failed" — 有内容但无法解析为合法 JSON（截断/格式错误）
      "empty_content"— 模型返回内容为空字符串
    """
    if not content or not content.strip():
        return [], "empty_content"

    # 语义嗅探：模型用自然语言说没有错误（无 JSON 对象）
    _NO_ERROR_HINTS = ["no error", "no issue", "没有发现", "未发现", "无错误",
                       "no errors found", "all correct", "符合规范"]
    _stripped_lower = content.strip().lower()
    if any(h in _stripped_lower for h in _NO_ERROR_HINTS) and '{' not in content:
        return [], "ok"

    try:
        cleaned = re.sub(r"```json|```", "", content).strip()
        match = re.search(r"\[.*\]", cleaned, re.S)
        candidate = match.group() if match else cleaned
        parsed = json.loads(candidate)
        if isinstance(parsed, list):
            return parsed, "ok"
        if isinstance(parsed, dict):
            return [parsed], "ok"
        return [], "parse_failed"
    except Exception:
        return [], "parse_failed"



_STATUS_OK         = "ok"
_STATUS_PARSE_FAIL = "parse_failed"
_STATUS_EMPTY      = "empty_content"
_STATUS_API_ERROR  = "api_error"
_STATUS_PARSE_FAIL = "parse_failed"
_STATUS_EMPTY      = "empty_content"
_STATUS_API_ERROR  = "api_error"
_DEFAULT_PENDING = {"is_correct": True, "errors": [], "source_issues": [], "_pending": True}


def _call_llm(seqs_rows: list) -> dict:
    """
    发送一批 (seq, row) 给模型，返回 {seq: result} 字典。
    seq 是调用方指定的序号，与块内位置解耦。
    """
    lines = [f"[{seq}] 原文: {r['原文']}\n[{seq}] 译文: {r['译文']}" for seq, r in seqs_rows]
    combined = "\n\n".join(lines)

    prompt = f"""你是翻译数值审校专家。以下是一批原文/译文对（共 {len(seqs_rows)} 条，用[序号]标记）。
    
##请按以下两步完成检查：
第一步：通读所有条目，建立整体语境认知——
- 理解文本所属领域（财务、工业、医疗等）
- 识别贯穿全文的数值体系（单位、量级、货币、百分比惯例等）
- 注意跨条目的数值关联（如前文提到总量，后文提到分项，两者应能对上）

第二步：基于整体语境，逐条判断每条译文数值是否正确，不要满足于只找到一个错误。每一条译文必须检查到最后一个字符，确保把**所有**翻译数值不一致的地方全部揪出来，区分两类情况——
- errors：译文数值与原文不符（数值出现错译、漏译、多译，严禁四舍五入数值，严禁序号数值和单位漏译），【需要修改译文】。一条译文中可能包含多个不同的 error，请全部列出。
- source_issues：译文忠实还原了原文，但原文数值本身存在逻辑问题，【不需要修改译文】

##要求：
特别注意中文数量单位换算：
- 万 = 10^4，例如 80万吨 = 800,000吨
- 亿 = 10^8
- 严禁直接将“80万吨”理解为 80,000 吨，必须乘以10000。
在进行数值比对前，请先写出原文的实际数值（去掉单位后的纯数字），再与译文数值对比。

##输出JSON数组，长度必须与输入条数({len(seqs_rows)})相同，每项对应同序号的条目：
[
  {{"seq": 0, "is_correct": true, "errors": [], "source_issues": []}},
  {{"seq": 1, "is_correct": false, "errors": [{_ERROR_SCHEMA}], "source_issues": [{_SOURCE_ISSUE_SCHEMA}]}}
]

待检查内容：
{combined}
"""
    try:
        resp = client.chat.completions.create(
            model="deepseek/deepseek-v4-pro",  # google/gemini-3.5-flash
            messages=[{"role": "system", "content": "只输出JSON数组"},
                      {"role": "user",   "content": prompt}],
            temperature=0,
        )
    except Exception as e:
        print(f"    ❌ 模型API调用异常: {e}")
        return {}, _STATUS_API_ERROR

    raw_content = resp.choices[0].message.content if resp.choices else ""
    parsed, status = _safe_parse_json(raw_content)
    seq_map = {item.get("seq", seqs_rows[i][0]): item for i, item in enumerate(parsed)}
    return seq_map, status


def _llm_check_block(block: list, max_retry: int = 1) -> list:
    """
    block: [(orig_idx, row), ...]
    返回与 block 等长的结果列表，明确区分三种情况：
      1. seq 在 seq_map 中(status=ok) → 用模型结果(errors=[]即真没错误)
      2. API异常/解析失败/空响应 → 补发重试；仍失败标记 _pending + _error_status
      3. 模型正常返回但漏返某条 → 单独补发；仍缺失标记 _pending
    """
    seqs_rows = [(seq, r) for seq, (_, r) in enumerate(block)]

    seq_map, status = _call_llm(seqs_rows)

    # 记录每条缺失的最后异常原因
    last_status = {seq: status for seq, _ in seqs_rows if seq not in seq_map}

    for _ in range(max_retry):
        missing = [(seq, r) for seq, r in seqs_rows if seq not in seq_map]
        if not missing:
            break
        reason_label = {
            _STATUS_API_ERROR:  "API调用异常",
            _STATUS_PARSE_FAIL: "返回无法解析",
            _STATUS_EMPTY:      "返回为空",
            _STATUS_OK:         "模型漏返部分条目",
        }.get(status, status)
        print(f"    ⚠️  {len(missing)} 条未获有效结果（{reason_label}），补发重试...")
        retry_map, status = _call_llm(missing)
        seq_map.update(retry_map)
        for seq, _ in missing:
            if seq not in retry_map:
                last_status[seq] = status

    still_missing = [seq for seq, _ in seqs_rows if seq not in seq_map]
    if still_missing:
        print(f"    ❌ {len(still_missing)} 条补发后仍无有效结果，标记为待确认: {still_missing}")
    for seq in still_missing:
        pending = _DEFAULT_PENDING.copy()
        pending["_error_status"] = last_status.get(seq, "unknown")
        seq_map[seq] = pending

    return [seq_map[seq] for seq, _ in seqs_rows]


def run(alignment_path: str,
        output_path: str = "reports/final_checked.xlsx",
        block_size: int = 20):

    # 1. 规则检查 → JSON
    print("📄 规则检查...")
    rows = run_number_check(alignment_path)
    error_count = sum(1 for r in rows if r["是否错误"] == "❗错误")
    print(f"🚨 规则错误: {error_count} / {len(rows)}")

    # 2. AI 复核（仅对错误行）
    error_rows = [(i, r) for i, r in enumerate(rows) if r["是否错误"] == "❗错误"]
    blocks = [error_rows[i:i + block_size] for i in range(0, len(error_rows), block_size)]
    print(f"🤖 AI复核，共 {len(blocks)} 块...")

    ai_map = {}
    for b_idx, block in enumerate(blocks):
        print(f"  Block {b_idx + 1}/{len(blocks)}")
        results = _llm_check_block(block)
        for pos, (idx, _) in enumerate(block):
            ai_map[idx] = results[pos]

    # 3. 合并 → JSON
    final_rows = merge_ai_results(rows, ai_map)

    # 4. JSON → Excel
    generate_combined_report(final_rows, output_path)
    print(f"✅ 完成: {output_path}")
    return final_rows


if __name__ == "__main__":
    run(
        alignment_path=r"C:\Users\H\Desktop\数检_程序-AI\测试\final_align.xlsx",
        output_path=r"C:\Users\H\Desktop\数检_程序-AI\测试\final_checked.xlsx",
    )



