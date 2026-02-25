from typing import List, Dict, Any
import json
import os
import re
import ast

from typing import List, Dict, Any, Optional
import json
import re
import ast

class jsonParser:
    def _extract_first_json_like_block(raw: str) -> Optional[str]:
        """
        提取文本中第一个完整的 JSON-like 块：{...} 或 [...]
        允许后面还有其他内容。
        """
        if not raw:
            return None

        m = re.search(r"[\{\[]", raw)
        if not m:
            return None

        start = m.start()
        open_ch = raw[start]
        close_ch = "}" if open_ch == "{" else "]"

        depth = 0
        in_str = False
        str_ch = ""
        escape = False

        for i in range(start, len(raw)):
            ch = raw[i]

            if in_str:
                if escape:
                    escape = False
                elif ch == "\\":
                    escape = True
                elif ch == str_ch:
                    in_str = False
                continue

            if ch in ("'", '"'):
                in_str = True
                str_ch = ch
                continue

            if ch == open_ch:
                depth += 1
            elif ch == close_ch:
                depth -= 1
                if depth == 0:
                    return raw[start:i + 1]

        return None


    def _split_merged_fields(one: Dict[str, Any]) -> Dict[str, Any]:
        """
        兜底：处理 '译文数值' 吞掉 '译文修改建议值:' 的情况
        """
        tv = str(one.get("译文数值") or "").strip()
        sv = str(one.get("译文修改建议值") or "").strip()

        # 允许“译 文修改建议值”等空格形态
        marker_pat = r"(译\s*文\s*修\s*改\s*建\s*议\s*值)\s*[:：]\s*"
        m = re.search(marker_pat, tv, flags=re.IGNORECASE)
        if not m:
            return one

        left = tv[:m.start()].strip()
        right = tv[m.end():].strip()

        if left:
            one["译文数值"] = re.sub(r"\s+", " ", left).strip()
        if right and not sv:
            one["译文修改建议值"] = re.sub(r"\s+", " ", right).strip()
        return one

    def _normalize_broken_json_keys(raw: str) -> str:
        """
        修复 LLM/换行导致的 JSON key 断裂问题，例如：
        "译文
          数值"  -> "译文数值"
        "译 文 修改 建 议 值" -> "译文修改建议值"

        只修 key（双引号包裹的键名），不改 value。
        """
        if not raw:
            return raw

        # 统一换行
        raw = raw.replace("\r\n", "\n").replace("\r", "\n")

        # 只在 JSON key 区域做替换： "...."\s*:
        # 把 key 内部所有空白（含换行）压掉
        def _fix_key(m: re.Match) -> str:
            key = m.group(1)
            fixed = re.sub(r"\s+", "", key)  # key 内空白全部去掉
            return f"\"{fixed}\":"

        raw = re.sub(r"\"([^\"]*?)\"\s*:", _fix_key, raw, flags=re.DOTALL)

        return raw


    def load_error_json_text(raw: str) -> List[Dict[str, Any]]:
        """
        支持：
        1) 严格 JSON 数组: [{...},{...}]
        2) 严格 JSON 单对象: {...} -> 自动包装成 list
        3) Python 字面量 list/dict（单引号） -> literal_eval
        4) 旧的“错误编号块”文本格式
        """
        raw = (raw or "").strip()
        if not raw:
            print("提示：错误报告文本为空，未找到可解析的错误列表")
            return []

        # ✅ 新增：先修复被换行/空格拆坏的 JSON key
        raw = jsonParser._normalize_broken_json_keys(raw)

        # ===== 1) 优先：提取第一个 {...} 或 [...] 块，先按 JSON 解析 =====
        block = jsonParser._extract_first_json_like_block(raw)
        if block:
            # 先尝试 JSON
            try:
                obj = json.loads(block)
                if isinstance(obj, dict):
                    obj = [obj]
                if isinstance(obj, list):
                    out = []
                    for it in obj:
                        if isinstance(it, dict):
                            out.append(jsonParser._split_merged_fields(it))
                    if out:
                        return out
            except Exception:
                # 再尝试 Python 字面量（单引号那种）
                try:
                    obj = ast.literal_eval(block)
                    if isinstance(obj, dict):
                        obj = [obj]
                    if isinstance(obj, list):
                        out = []
                        for it in obj:
                            if isinstance(it, dict):
                                out.append(jsonParser._split_merged_fields(it))
                        if out:
                            return out
                except Exception:
                    pass

        # ===== 2) 退化：解析“错误编号块”文本 =====
        cleaned = raw
        cleaned = cleaned.replace("错误\n编号", "错误编号")
        cleaned = cleaned.replace("错误\n编号:", "错误编号:")
        cleaned = cleaned.replace("编号:", "错误编号:")

        parts = re.split(r"(?=错误编号\s*[:：]\s*\d+)", cleaned)
        errors: List[Dict[str, Any]] = []

        key_map = [
            ("错误编号", r"错误\s*编号\s*[:：]\s*(.+?)(?=错误\s*类型\s*[:：]|$)"),
            ("错误类型", r"错误\s*类型\s*[:：]\s*(.+?)(?=原文\s*数值\s*[:：]|$)"),
            ("原文数值", r"原文\s*数值\s*[:：]\s*(.+?)(?=译文\s*数值\s*[:：]|$)"),
            ("译文数值", r"译文\s*数值\s*[:：]\s*(.+?)(?=译文\s*修改\s*建议\s*值\s*[:：]|修改\s*理由\s*[:：]|$)"),
            ("译文修改建议值", r"译文\s*修改\s*建议\s*值\s*[:：]\s*(.+?)(?=修改\s*理由\s*[:：]|$)"),
            ("修改理由", r"修改\s*理由\s*[:：]\s*(.+?)(?=违反\s*的\s*规则\s*[:：]|$)"),
            ("违反的规则", r"违反\s*的\s*规则\s*[:：]\s*(.+?)(?=原文\s*上下文\s*[:：]|$)"),
            ("原文上下文", r"原文\s*上下文\s*[:：]\s*(.+?)(?=译文\s*上下文\s*[:：]|$)"),
            ("译文上下文", r"译文\s*上下文\s*[:：]\s*(.+?)(?=原文\s*位置\s*[:：]|$)"),
            ("原文位置", r"原文\s*位置\s*[:：]\s*(.+?)(?=译文\s*位置\s*[:：]|$)"),
            ("译文位置", r"译文\s*位置\s*[:：]\s*(.+?)(?=替换\s*锚点\s*[:：]|$)"),
            ("替换锚点", r"替换\s*锚点\s*[:：]\s*(.+?)(?=错误\s*编号\s*[:：]|$)"),
        ]

        for p in parts:
            if "错误编号" not in p:
                continue

            one: Dict[str, Any] = {}
            for k, pat in key_map:
                mm = re.search(pat, p, flags=re.DOTALL)
                if mm:
                    v = re.sub(r"\s+", " ", mm.group(1).strip())
                    one[k] = v

            one = jsonParser._split_merged_fields(one)

            if one:
                errors.append(one)

        if not errors:
            print("提示：错误报告文本为空，未找到可解析的错误列表")
            return []

        return errors