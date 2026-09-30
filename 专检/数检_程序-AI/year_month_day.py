import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional

try:
    import pandas as pd
except Exception:  # pragma: no cover - pandas 不存在时允许模块其余功能继续使用
    pd = None


MONTH_MAP = {
    # 短缩写（不带点）
    "jan": "01", "feb": "02", "mar": "03", "apr": "04", "may": "05", "jun": "06",
    "jul": "07", "aug": "08", "sep": "09", "oct": "10", "nov": "11", "dec": "12",

    # 带点缩写（关键补充）
    "jan.": "01", "feb.": "02", "mar.": "03", "apr.": "04",
    "jun.": "06", "jul.": "07", "aug.": "08", "sep.": "09",
    "oct.": "10", "nov.": "11", "dec.": "12",

    # 完整月份
    "january": "01", "february": "02", "march": "03", "april": "04",
    "june": "06", "july": "07", "august": "08",
    "september": "09", "october": "10", "november": "11", "december": "12",
}

MONTH_REGEX = (
    r"January|February|March|April|May|June|July|August|September|"
    r"October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Oct|Nov|Dec"
)

BUILTIN_CONTEXT_MAP = {
    "year": "year",
    "years": "year",
    "yr": "year",
    "yrs": "year",
    "annual": "year",
    "年": "year",
    "month": "month",
    "months": "month",
    "monthly": "month",
    "月": "month",
    "day": "day",
    "days": "day",
    "daily": "day",
    "date": "day",
    "日": "day",
}
BUILTIN_CONTEXT_MAP.update({month_name: "month" for month_name in MONTH_MAP})

LABEL_ORDER = {"year": 0, "month": 1, "day": 2}


@dataclass
class DateContextItem:
    group_index: int
    raw_label: str
    value: str
    raw_value: str
    start: int
    end: int
    raw_text: str


@dataclass
class DateContextCompareResult:
    source_items: List[DateContextItem] = field(default_factory=list)
    target_items: List[DateContextItem] = field(default_factory=list)
    matched: bool = False
    mismatches: List[dict] = field(default_factory=list)


def _normalize_plain_number(number_text: str) -> str:
    """将数值转成不带前导零的普通字符串。"""
    try:
        return str(int(number_text))
    except (TypeError, ValueError):
        return str(number_text).strip()


def _normalize_year(year_text: str) -> str:
    year_text = str(year_text).strip()
    if len(year_text) == 2 and year_text.isdigit():
        return "20" + year_text
    return year_text


def _build_yyyymmdd(year_text: str, month_text: str, day_text: str) -> str:
    """拼成 YYYYMMDD，用于旧的归一化输出。"""
    return f"{_normalize_year(year_text)}{str(month_text).zfill(2)}{str(day_text).zfill(2)}"


def _split_alias_values(value: object) -> List[str]:
    if value is None:
        return []
    text = str(value).strip()
    if not text:
        return []
    return [part.strip() for part in re.split(r"[,，;/|、\n]+", text) if part.strip()]


def load_context_mapping_from_excel(excel_path: str, sheet_name: object = 0) -> Dict[str, str]:
    """
    从 Excel 加载上下文映射。

    支持两类常见表头：
      1. `canonical/标准/归一标签` + 若干别名列
      2. `中文/英文` 两列，默认用英文列作为归一标签，例如 年 -> year
    """
    if pd is None:
        raise ImportError("当前环境未安装 pandas，无法读取 Excel 映射表。")

    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    if df.empty:
        return {}

    normalized_cols = {str(col).strip().lower(): col for col in df.columns}
    canonical_candidates = {
        "canonical", "normalized", "standard", "标准", "标准标签", "归一标签", "canonical_label",
    }
    english_candidates = {"en", "eng", "english", "英文", "target", "tgt"}

    canonical_col = None
    for key, original_col in normalized_cols.items():
        if key in canonical_candidates:
            canonical_col = original_col
            break

    if canonical_col is None:
        for key, original_col in normalized_cols.items():
            if key in english_candidates:
                canonical_col = original_col
                break

    columns = list(df.columns)
    if canonical_col is None:
        canonical_col = columns[1] if len(columns) >= 2 else columns[0]

    alias_cols = [col for col in columns if col != canonical_col]
    mapping: Dict[str, str] = {}

    for _, row in df.fillna("").iterrows():
        canonical_value = str(row[canonical_col]).strip().lower()
        if not canonical_value:
            continue

        mapping[canonical_value] = canonical_value
        for alias_col in alias_cols:
            for alias in _split_alias_values(row[alias_col]):
                mapping[alias.lower()] = canonical_value

    return mapping


def normalize_context_label(label: str, mapping: Optional[Dict[str, str]] = None) -> str:
    """将 `年/year/june` 这类原始标签归一成统一标签。"""
    key = str(label).strip().lower()
    combined_mapping = dict(BUILTIN_CONTEXT_MAP)
    if mapping:
        combined_mapping.update({str(k).strip().lower(): str(v).strip().lower() for k, v in mapping.items()})
    return combined_mapping.get(key, key)


def _group_date_items(items: List[DateContextItem], mapping: Optional[Dict[str, str]] = None) -> List[Dict[str, DateContextItem]]:
    grouped: Dict[int, Dict[str, DateContextItem]] = {}
    for item in items:
        normalized_label = normalize_context_label(item.raw_label, mapping)
        grouped.setdefault(item.group_index, {})[normalized_label] = item

    def group_start(group: Dict[str, DateContextItem]) -> int:
        return min(item.start for item in group.values())

    return sorted(grouped.values(), key=group_start)


def extract_date_contexts(text: str) -> List[DateContextItem]:
    """
    提取日期相关的“值 + 上下文标签”。

    示例：
      - `2025年6月31日` -> year=2025, month=6, day=31
      - `June 31, 2025` -> month(raw_label=June,value=6), day=31, year=2025
    """
    if not text:
        return []

    items: List[DateContextItem] = []
    consumed: List[tuple] = []
    group_index = 0

    def is_consumed(start: int, end: int) -> bool:
        for consumed_start, consumed_end in consumed:
            if start < consumed_end and end > consumed_start:
                return True
        return False

    def mark_consumed(start: int, end: int) -> None:
        consumed.append((start, end))

    def add_group(
        match_obj: re.Match,
        year_text: str,
        month_text: str,
        day_text: str,
        year_label: str,
        month_label: str,
        day_label: str,
        raw_month_value: Optional[str] = None,
    ) -> None:
        nonlocal group_index
        start, end = match_obj.start(), match_obj.end()
        if is_consumed(start, end):
            return

        raw_text = match_obj.group(0)
        items.extend(
            [
                DateContextItem(
                    group_index=group_index,
                    raw_label=year_label,
                    value=_normalize_plain_number(_normalize_year(year_text)),
                    raw_value=year_text,
                    start=start,
                    end=end,
                    raw_text=raw_text,
                ),
                DateContextItem(
                    group_index=group_index,
                    raw_label=month_label,
                    value=_normalize_plain_number(month_text),
                    raw_value=raw_month_value or month_text,
                    start=start,
                    end=end,
                    raw_text=raw_text,
                ),
                DateContextItem(
                    group_index=group_index,
                    raw_label=day_label,
                    value=_normalize_plain_number(day_text),
                    raw_value=day_text,
                    start=start,
                    end=end,
                    raw_text=raw_text,
                ),
            ]
        )
        mark_consumed(start, end)
        group_index += 1

    # 中文/ISO 风格: 2025年6月31日, 2025-06-31, 2025/6/31
    pattern_ymd = re.compile(
        r"(?<!\d)(?P<year>\d{2,4})[年\-\./](?P<month>\d{1,2})[月\-\./](?P<day>\d{1,2})日?(?!\d)"
    )
    for match_obj in pattern_ymd.finditer(text):
        raw_text = match_obj.group(0)
        year_label = "年" if "年" in raw_text else "year"
        month_label = "月" if "月" in raw_text else "month"
        day_label = "日" if "日" in raw_text else "day"
        add_group(
            match_obj,
            match_obj.group("year"),
            match_obj.group("month"),
            match_obj.group("day"),
            year_label,
            month_label,
            day_label,
        )

    # 紧凑格式: 20250631
    pattern_compact = re.compile(r"\b(?P<year>\d{4})(?P<month>\d{2})(?P<day>\d{2})\b")
    for match_obj in pattern_compact.finditer(text):
        add_group(
            match_obj,
            match_obj.group("year"),
            match_obj.group("month"),
            match_obj.group("day"),
            "year",
            "month",
            "day",
        )

    # 英文美式: June 31, 2025 / June 31st 2025
    pattern_mdy = re.compile(
        rf"(?P<month>{MONTH_REGEX})\s+(?P<day>\d{{1,2}})(?:st|nd|rd|th)?(?:,)?\s+(?P<year>\d{{4}})",
        re.IGNORECASE,
    )
    for match_obj in pattern_mdy.finditer(text):
        month_word = match_obj.group("month")
        add_group(
            match_obj,
            match_obj.group("year"),
            MONTH_MAP[month_word.lower()],
            match_obj.group("day"),
            "year",
            month_word,
            "day",
            raw_month_value=month_word,
        )

    # 英文英式: 31 June 2025 / The 31st of June, 2025
    pattern_dmy = re.compile(
        rf"(?:the\s+)?(?P<day>\d{{1,2}})(?:st|nd|rd|th)?(?:\s+of)?\s+(?P<month>{MONTH_REGEX})(?:,)?\s+(?P<year>\d{{4}})",
        re.IGNORECASE,
    )
    for match_obj in pattern_dmy.finditer(text):
        month_word = match_obj.group("month")
        add_group(
            match_obj,
            match_obj.group("year"),
            MONTH_MAP[month_word.lower()],
            match_obj.group("day"),
            "year",
            month_word,
            "day",
            raw_month_value=month_word,
        )

    # 全数字英文日期: 31/06/2025 或 06/31/2025
    pattern_digital = re.compile(r"\b(?P<p1>\d{1,2})[/\.](?P<p2>\d{1,2})[/\.](?P<year>\d{4})\b")
    for match_obj in pattern_digital.finditer(text):
        p1 = match_obj.group("p1")
        p2 = match_obj.group("p2")
        if int(p1) > 12:
            month_text = p2
            day_text = p1
        else:
            month_text = p1
            day_text = p2
        add_group(match_obj, match_obj.group("year"), month_text, day_text, "year", "month", "day")

    items.sort(key=lambda item: (item.start, LABEL_ORDER.get(normalize_context_label(item.raw_label), 99)))
    return items


def normalize_dates(text):
    """
    将文本中的日期统一归一成 `YYYYMMDD`。
    旧接口保留，便于兼容现有调用。
    """
    if not text:
        return text

    items = extract_date_contexts(text)
    groups = _group_date_items(items)
    result_parts = []
    last_index = 0

    for group in groups:
        sample_item = next(iter(group.values()), None)
        if sample_item is None:
            continue

        year_value = group.get("year").value if group.get("year") else ""
        month_value = group.get("month").value if group.get("month") else ""
        day_value = group.get("day").value if group.get("day") else ""
        if not (year_value and month_value and day_value):
            continue

        start = sample_item.start
        end = sample_item.end
        result_parts.append(text[last_index:start])
        result_parts.append(f" {_build_yyyymmdd(year_value, month_value, day_value)} ")
        last_index = end

    result_parts.append(text[last_index:])
    return "".join(result_parts)


def compare_dates_by_context(
    source_text: str,
    target_text: str,
    mapping_excel_path: Optional[str] = None,
    sheet_name: object = 0,
) -> DateContextCompareResult:
    """
    对比两段文本中的日期语义。

    对比时不是简单按数字顺序比较，而是先提取 `year/month/day` 语义后再比。
    如果提供 Excel，可把 `年 -> year`、`月份 -> month`、`june -> month` 这类关系放进去。
    """
    extra_mapping = load_context_mapping_from_excel(mapping_excel_path, sheet_name) if mapping_excel_path else {}

    source_items = extract_date_contexts(source_text)
    target_items = extract_date_contexts(target_text)
    result = DateContextCompareResult(source_items=source_items, target_items=target_items)

    source_groups = _group_date_items(source_items, extra_mapping)
    target_groups = _group_date_items(target_items, extra_mapping)

    max_len = max(len(source_groups), len(target_groups))
    mismatches: List[dict] = []

    for index in range(max_len):
        source_group = source_groups[index] if index < len(source_groups) else {}
        target_group = target_groups[index] if index < len(target_groups) else {}

        for label in ("year", "month", "day"):
            source_item = source_group.get(label)
            target_item = target_group.get(label)
            source_value = source_item.value if source_item else "<缺失>"
            target_value = target_item.value if target_item else "<missing>"
            if source_value != target_value:
                mismatches.append(
                    {
                        "date_index": index,
                        "label": label,
                        "source": source_value,
                        "target": target_value,
                        "source_label": source_item.raw_label if source_item else "",
                        "target_label": target_item.raw_label if target_item else "",
                    }
                )

    result.matched = not mismatches and len(source_groups) == len(target_groups)
    result.mismatches = mismatches
    return result


if __name__ == '__main__':
    text1="""
    1. 中文日期表达 (Big-Endian: 年-月-日)
--------------------------------------------------------------------------------
- 标准文本：2025年6月31日
- 短文本：25年6月31日
- 全数字分隔（点）：2025.06.31 / 2025.6.31
- 全数字分隔（横杠）：2025-06-31 / 2025-6-31
- 全数字分隔（斜杠）：2025/06/31 / 2025/6/31
- 紧凑型：20250631

2. 英文日期表达 - 美式 (Middle-Endian: 月-日-年)
--------------------------------------------------------------------------------
- 标准全称：June 31, 2025
- 月份缩写：Jun 31, 2025
- 序数词带逗号：June 31st, 2025
- 全数字型：06/31/2025 (或 06-31-2025)
- 无年份表达：June 31st

3. 英文日期表达 - 英式 (Little-Endian: 日-月-年)
--------------------------------------------------------------------------------
- 标准全称：31 June 2025
- 序数词型：31st June 2025
- 带介词型：The 31st of June, 2025
- 全数字型：31/06/2025 (或 31.06.2025)

4. 国际标准 (ISO 8601)
--------------------------------------------------------------------------------
- 标准格式：2025-06-31
    """

    res1 = normalize_dates(text1)
    print(res1)

    print("\n" + "=" * 80)
    print("提取日期上下文")
    print("=" * 80)
    cn_text = """截至2025/06/31 ，公司完成披露。"""
    en_text = "As of June 31st, 2025, the Company completed the disclosure."

    test_cases = [
        # --- 中文系列 ---
        ("标准中文", "截至2025年6月31日，披露完成"),
        ("短年份中文", "25年06月31日发布报告"),
        ("中文点分隔", "日期：2025.6.31"),
        ("中文横杠", "2025-06-31期间"),
        ("中文斜杠", "2025/6/31有效"),
        ("中文紧凑", "20250631是关键节点"),

        # --- 英文美式 (Month DD, YYYY) ---
        ("美式全称", "As of June 31, 2025, completed."),
        ("美式缩写", "Reported on Jun 31, 2025."),
        ("美式序数词", "On June 31st, 2025 at noon."),
        ("美式无逗号", "Dated June 31 2025"),
        ("美式全数字", "The date is 06/31/2025."),

        # --- 英文英式 (DD Month YYYY) ---
        ("英式全称", "31 June 2025 is the deadline."),
        ("英式带the/of", "The 31st of June, 2025."),
        ("英式缩写", "31-Jun-2025"),
        ("英式数字顺序", "Deadline: 31/06/2025"),  # 重点：第一个数 > 12

        # --- 国际标准 (ISO) ---
        ("ISO标准", "The period ending 2025-06-31."),

        # --- 混合与异常情况 (抗干扰测试) ---
        ("金额干扰", "营收389.00亿元"),  # 不应识别为日期
        ("版本号干扰", "版本 v1.0.2025"),  # 不应识别
        ("百分比干扰", "利润增长了8.8%"),  # 不应识别
        ("多日期混合", "从2025/01/01到June 31, 2025"),
    ]


    def run_test(func):
        print(f"{'测试项':<15} | {'原始文本':<30} | {'归一化结果'}")
        print("-" * 80)
        for name, text in test_cases:
            res = func(text).strip()
            print(f"{name:<15} | {text:<30} | {res}")


    run_test(normalize_dates)
    print("中文提取：")
    for item in extract_date_contexts(cn_text):
        print(vars(item))

    print("英文提取：")
    for item in extract_date_contexts(en_text):
        print(vars(item))

    result = compare_dates_by_context(cn_text, en_text)
    print("对比结果:", result.matched)
    print("差异明细:", result.mismatches)
