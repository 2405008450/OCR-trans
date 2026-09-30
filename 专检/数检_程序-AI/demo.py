import re
from dataclasses import dataclass
from typing import Callable, Dict, List

from normalizer import extract_numbers
from year_month_day import MONTH_MAP, compare_dates_by_context


MONTH_WORD_PATTERN = (
    r"January|February|March|April|May|June|July|August|September|"
    r"October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Oct|Nov|Dec"
)

SMALL_NUMBER_WORDS = {
    0: "zero",
    1: "one",
    2: "two",
    3: "three",
    4: "four",
    5: "five",
    6: "six",
    7: "seven",
    8: "eight",
    9: "nine",
}

ROMAN_NUMERALS = {
    1: "I",
    2: "II",
    3: "III",
    4: "IV",
    5: "V",
    6: "VI",
    7: "VII",
    8: "VIII",
    9: "IX",
    10: "X",
}

CN_MONEY_SCALE = {
    "百": 100,
    "千": 1_000,
    "万": 10_000,
    "百万": 1_000_000,
    "千万": 10_000_000,
    "亿": 100_000_000,
}


@dataclass
class RuleCase:
    category: str
    rule: str
    cn: str
    en: str
    checker: str


@dataclass
class CheckResult:
    passed: bool
    reason: str
    extracted: Dict[str, object]


def _split_segments(text: str) -> List[str]:
    return [part.strip() for part in re.split(r"[；;]", text) if part.strip()]


def _normalize_apostrophes(text: str) -> str:
    return (
        text.replace("’", "'")
        .replace("‘", "'")
        .replace("“", '"')
        .replace("”", '"')
        .replace("″", "''")
        .replace("′", "'")
    )


def _digits_only(text: str) -> str:
    return re.sub(r"[^\d.]", "", text)


def _extract_first_plain_number(text: str) -> str:
    match_obj = re.search(r"\d[\d,]*\.?\d*", text)
    if not match_obj:
        return ""
    return match_obj.group(0).replace(",", "")


def _parse_cn_money(text: str) -> Dict[str, object]:
    match_obj = re.search(r"([\d,]+(?:\.\d+)?)\s*(百|千|万|百万|千万|亿)?(美元|人民币|元)", text)
    if not match_obj:
        return {"currency": "", "amount": None}

    raw_number = match_obj.group(1).replace(",", "")
    scale_text = match_obj.group(2) or ""
    currency_text = match_obj.group(3)
    amount = float(raw_number) * CN_MONEY_SCALE.get(scale_text, 1)
    currency = "USD" if currency_text == "美元" else "RMB"
    return {"currency": currency, "amount": amount}


def _parse_en_money(text: str) -> Dict[str, object]:
    match_obj = re.search(r"\b(USD|RMB)\s+(\d{1,3}(?:,\d{3})*(?:\.\d+)?)\b", text, re.IGNORECASE)
    if not match_obj:
        return {"currency": "", "amount": None}
    return {
        "currency": match_obj.group(1).upper(),
        "amount": float(match_obj.group(2).replace(",", "")),
    }


def _parse_cn_month_year(text: str) -> Dict[str, str]:
    match_obj = re.search(r"(?P<year>\d{4})年(?P<month>\d{1,2})月", text)
    if not match_obj:
        return {}
    return {
        "year": match_obj.group("year"),
        "month": str(int(match_obj.group("month"))),
    }


def _parse_en_month_year(text: str) -> Dict[str, str]:
    match_obj = re.search(rf"(?P<month>{MONTH_WORD_PATTERN})\s+(?P<year>\d{{4}})", text, re.IGNORECASE)
    if not match_obj:
        return {}
    return {
        "year": match_obj.group("year"),
        "month": str(int(MONTH_MAP[match_obj.group("month").lower()])),
    }


def _parse_cn_decade(text: str) -> Dict[str, str]:
    match_obj = re.search(r"(\d{1,2})世纪(\d{2})年代", text)
    if not match_obj:
        return {}
    century = int(match_obj.group(1))
    decade = match_obj.group(2)
    year_prefix = (century - 1) * 100
    return {"decade": f"{year_prefix + int(decade)}s"}


def _parse_en_decade(text: str) -> Dict[str, str]:
    match_obj = re.search(r"\bthe\s+(\d{4}s)\b", text, re.IGNORECASE)
    if not match_obj:
        return {}
    return {"decade": match_obj.group(1)}


def _parse_cn_era(text: str) -> List[Dict[str, str]]:
    result = []
    for segment in _split_segments(text):
        match_obj = re.search(r"(公元前|公元)?(\d+)年", segment)
        if not match_obj:
            continue
        era = "BC" if match_obj.group(1) == "公元前" else "AD"
        result.append({"era": era, "year": match_obj.group(2)})
    return result


def _parse_en_era(text: str) -> List[Dict[str, str]]:
    result = []
    for segment in _split_segments(text):
        bc_match = re.search(r"\b(\d+)\s*BC\b", segment, re.IGNORECASE)
        ad_match = re.search(r"\bAD\s*(\d+)\b", segment, re.IGNORECASE)
        if bc_match:
            result.append({"era": "BC", "year": bc_match.group(1)})
        elif ad_match:
            result.append({"era": "AD", "year": ad_match.group(1)})
    return result


def _check_basic_number_alignment(case: RuleCase) -> CheckResult:
    cn_numbers = extract_numbers(case.cn)
    en_numbers = extract_numbers(case.en)
    passed = cn_numbers == en_numbers
    return CheckResult(
        passed=passed,
        reason="数值顺序一致" if passed else "中英文提取出的数值序列不一致",
        extracted={"cn_numbers": cn_numbers, "en_numbers": en_numbers},
    )


def _check_digit_word_rule(case: RuleCase) -> CheckResult:
    cn_numbers = [int(float(x)) for x in extract_numbers(case.cn)]
    en_numbers = [int(float(x)) for x in extract_numbers(case.en)]
    word_hits = []
    digit_hits = []
    for value in cn_numbers:
        if value <= 9:
            word = SMALL_NUMBER_WORDS.get(value, "")
            word_hits.append(bool(re.search(rf"\b{re.escape(word)}\b", case.en, re.IGNORECASE)))
        else:
            digit_hits.append(bool(re.search(rf"\b{value}\b", case.en)))
    passed = cn_numbers == en_numbers and all(word_hits) and all(digit_hits)
    return CheckResult(
        passed=passed,
        reason="0-9 使用英文单词，10 及以上保留数字" if passed else "未同时满足小数词/大数阿拉伯数字规则",
        extracted={"cn_numbers": cn_numbers, "en_numbers": en_numbers, "word_hits": word_hits, "digit_hits": digit_hits},
    )


def _check_scientific_unit_rule(case: RuleCase) -> CheckResult:
    cn_number = _extract_first_plain_number(case.cn)
    en_number = _extract_first_plain_number(case.en)
    unit_ok = bool(re.search(r"(μg/m3|ug/m3|μg/m²|μg/m2|μg/m³)", case.en, re.IGNORECASE))
    passed = bool(cn_number and cn_number == en_number and unit_ok)
    return CheckResult(
        passed=passed,
        reason="数值和科学计量单位格式正确" if passed else "数值匹配或单位格式不符合预期",
        extracted={"cn_number": cn_number, "en_number": en_number, "unit_ok": unit_ok},
    )


def _check_heading_roman_rule(case: RuleCase) -> CheckResult:
    base_result = _check_basic_number_alignment(case)
    cn_numbers = [int(float(x)) for x in base_result.extracted["cn_numbers"]]
    roman_ok = False
    if cn_numbers:
        roman = ROMAN_NUMERALS.get(cn_numbers[0], "")
        roman_ok = bool(roman and re.search(rf"^\s*{roman}\.", case.en))
    passed = base_result.passed and roman_ok
    return CheckResult(
        passed=passed,
        reason="标题序号已转成罗马数字" if passed else "标题序号未按罗马数字呈现",
        extracted={**base_result.extracted, "roman_ok": roman_ok},
    )


def _check_percentage_point_rule(case: RuleCase) -> CheckResult:
    cn_segments = _split_segments(case.cn)
    en_segments = _split_segments(case.en)
    pairs = list(zip(cn_segments, en_segments))
    details = []
    passed = len(cn_segments) == len(en_segments) and bool(pairs)

    for cn_seg, en_seg in pairs:
        cn_number = _extract_first_plain_number(cn_seg)
        en_number = _extract_first_plain_number(en_seg)
        value = float(cn_number) if cn_number else None
        singular_ok = bool(value == 1 and re.search(r"\b1(?:\.0+)? percentage point\b", en_seg, re.IGNORECASE))
        plural_ok = bool(value is not None and value != 1 and re.search(r"\bpercentage points\b", en_seg, re.IGNORECASE))
        item_ok = bool(cn_number and cn_number == en_number and (singular_ok or plural_ok))
        details.append({"cn": cn_seg, "en": en_seg, "cn_number": cn_number, "en_number": en_number, "item_ok": item_ok})
        passed = passed and item_ok

    return CheckResult(
        passed=passed,
        reason="百分点单复数使用正确" if passed else "百分点的数值或单复数用法不符合规则",
        extracted={"pairs": details},
    )


def _check_fraction_rule(case: RuleCase) -> CheckResult:
    en_norm = case.en.lower()
    checks = [
        bool(re.search(r"\bone sixth\b", en_norm)),
        bool(re.search(r"\btwo-thirds\b", en_norm)),
        "/" not in case.en,
    ]
    passed = all(checks)
    return CheckResult(
        passed=passed,
        reason="分数已转换为英文分数表达" if passed else "分数表达仍存在直译或缺少连字符",
        extracted={"checks": checks, "cn_numbers": extract_numbers(case.cn), "en_numbers": extract_numbers(case.en)},
    )


def _check_date_rule(case: RuleCase) -> CheckResult:
    cn_segments = _split_segments(case.cn)
    en_segments = _split_segments(case.en)
    passed = len(cn_segments) == len(en_segments) and bool(cn_segments)
    details = []

    for cn_seg, en_seg in zip(cn_segments, en_segments):
        if "日" in cn_seg and re.search(r"\d{4}年\d{1,2}月\d{1,2}日", cn_seg):
            compare_result = compare_dates_by_context(cn_seg, en_seg)
            item_ok = compare_result.matched
            details.append(
                {
                    "cn": cn_seg,
                    "en": en_seg,
                    "type": "full_date",
                    "matched": compare_result.matched,
                    "mismatches": compare_result.mismatches,
                }
            )
        else:
            cn_info = _parse_cn_month_year(cn_seg)
            en_info = _parse_en_month_year(en_seg)
            item_ok = bool(cn_info and en_info and cn_info == en_info and "," not in en_seg)
            details.append({"cn": cn_seg, "en": en_seg, "type": "month_year", "cn_info": cn_info, "en_info": en_info, "matched": item_ok})
        passed = passed and item_ok

    return CheckResult(
        passed=passed,
        reason="日期格式符合“月日, 年 / 月年”规则" if passed else "日期格式或日期语义比对未通过",
        extracted={"pairs": details},
    )


def _check_decade_rule(case: RuleCase) -> CheckResult:
    cn_info = _parse_cn_decade(case.cn)
    en_info = _parse_en_decade(case.en)
    passed = bool(cn_info and en_info and cn_info == en_info)
    return CheckResult(
        passed=passed,
        reason="年代前已正确加 the" if passed else "年代表达未按 the 1930s 形式呈现",
        extracted={"cn_info": cn_info, "en_info": en_info},
    )


def _check_era_rule(case: RuleCase) -> CheckResult:
    cn_info = _parse_cn_era(case.cn)
    en_info = _parse_en_era(case.en)
    passed = cn_info == en_info and len(cn_info) == len(en_info)
    return CheckResult(
        passed=passed,
        reason="BC/AD 纪年表达正确" if passed else "纪年或公元前/公元标记不一致",
        extracted={"cn_info": cn_info, "en_info": en_info},
    )


def _check_money_rule(case: RuleCase) -> CheckResult:
    cn_info = _parse_cn_money(case.cn)
    en_info = _parse_en_money(case.en)
    thousands_ok = bool(re.search(r"\b(?:USD|RMB)\s+\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?\b", case.en))
    decimal_ok = not bool(re.search(r"\b(?:USD|RMB)\s+\d{1,3}(?:,\d{3})*\.\d{3,}\b", case.en))
    amount_match = cn_info["amount"] is not None and en_info["amount"] is not None and abs(cn_info["amount"] - en_info["amount"]) < 1e-6
    passed = amount_match and cn_info["currency"] == en_info["currency"] and thousands_ok and decimal_ok
    return CheckResult(
        passed=passed,
        reason="金额币种、千分位和小数位规则正确" if passed else "金额格式或数值换算不符合规则",
        extracted={"cn_info": cn_info, "en_info": en_info, "thousands_ok": thousands_ok, "decimal_ok": decimal_ok},
    )


def _check_unit_fullname_rule(case: RuleCase) -> CheckResult:
    cn_number = _extract_first_plain_number(case.cn)
    en_number = _extract_first_plain_number(case.en)
    unit_ok = bool(re.search(r"\bkilometers\b", case.en, re.IGNORECASE))
    passed = bool(cn_number and cn_number == en_number and unit_ok)
    return CheckResult(
        passed=passed,
        reason="数值和单位全称一致" if passed else "单位未按全称输出或数值不一致",
        extracted={"cn_number": cn_number, "en_number": en_number, "unit_ok": unit_ok},
    )


def _check_table_unit_rule(case: RuleCase) -> CheckResult:
    en_norm = _normalize_apostrophes(case.en)
    passed = "In RMB '00 Million" in en_norm
    return CheckResult(
        passed=passed,
        reason="表格单位格式正确" if passed else "表格单位格式不符合 In RMB '00 Million",
        extracted={"normalized_en": en_norm},
    )


def _check_coordinate_rule(case: RuleCase) -> CheckResult:
    cn_numbers = extract_numbers(case.cn)
    en_numbers = extract_numbers(case.en)
    en_norm = _normalize_apostrophes(case.en)
    north_ok = bool(re.search(r"\bN\b", en_norm))
    seconds_ok = "''" in en_norm
    passed = cn_numbers == en_numbers and north_ok and seconds_ok
    return CheckResult(
        passed=passed,
        reason="经纬度方向和秒符号格式正确" if passed else "经纬度数值、方向或秒符号不符合规则",
        extracted={"cn_numbers": cn_numbers, "en_numbers": en_numbers, "north_ok": north_ok, "seconds_ok": seconds_ok},
    )


def _check_free_translation_rule(case: RuleCase) -> CheckResult:
    cn_numbers = extract_numbers(case.cn)
    en_numbers = extract_numbers(case.en)
    no_digits_ok = not bool(re.search(r"\d", case.en))
    phrase_ok = "greater bay area" in case.en.lower()
    passed = cn_numbers == en_numbers and no_digits_ok and phrase_ok
    return CheckResult(
        passed=passed,
        reason="含义数字已意译且未直接保留阿拉伯数字" if passed else "含义数字仍像直译，或关键信息缺失",
        extracted={"cn_numbers": cn_numbers, "en_numbers": en_numbers, "no_digits_ok": no_digits_ok, "phrase_ok": phrase_ok},
    )


def _check_phase_rule(case: RuleCase) -> CheckResult:
    cn_numbers = [int(float(x)) for x in extract_numbers(case.cn)]
    en_numbers = [int(float(x)) for x in extract_numbers(case.en)]
    roman_ok = bool(re.search(r"\bPhase\s+[IVX]+\b", case.en))
    passed = cn_numbers == en_numbers and roman_ok
    return CheckResult(
        passed=passed,
        reason="期数已写成 Phase + 罗马数字" if passed else "期数未按 Phase + 大写罗马数字表达",
        extracted={"cn_numbers": cn_numbers, "en_numbers": en_numbers, "roman_ok": roman_ok},
    )


def _check_subscript_rule(case: RuleCase) -> CheckResult:
    pm_ok = "PM2.5" in case.cn and "PM2.5" in case.en
    note_ok = "subscript" in case.en.lower()
    passed = pm_ok and note_ok
    return CheckResult(
        passed=passed,
        reason="上下标说明保留完整" if passed else "PM2.5 或下标说明缺失",
        extracted={"pm_ok": pm_ok, "note_ok": note_ok},
    )


def _check_sentence_start_rule(case: RuleCase) -> CheckResult:
    base_result = _check_basic_number_alignment(case)
    starts_with_digit = bool(re.match(r"^\s*\d", case.en))
    passed = base_result.passed and not starts_with_digit
    return CheckResult(
        passed=passed,
        reason="英文句首未直接以数字开头" if passed else "英文句首仍以数字开头或数值不一致",
        extracted={**base_result.extracted, "starts_with_digit": starts_with_digit},
    )


CHECKERS: Dict[str, Callable[[RuleCase], CheckResult]] = {
    "digit_word": _check_digit_word_rule,
    "scientific_unit": _check_scientific_unit_rule,
    "heading_roman": _check_heading_roman_rule,
    "percentage_point": _check_percentage_point_rule,
    "fraction": _check_fraction_rule,
    "date": _check_date_rule,
    "decade": _check_decade_rule,
    "era": _check_era_rule,
    "money": _check_money_rule,
    "unit_fullname": _check_unit_fullname_rule,
    "table_unit": _check_table_unit_rule,
    "coordinate": _check_coordinate_rule,
    "free_translation": _check_free_translation_rule,
    "phase": _check_phase_rule,
    "subscript": _check_subscript_rule,
    "sentence_start": _check_sentence_start_rule,
}


RULE_CASES: List[RuleCase] = [
    RuleCase("数字/序数词", "0-9用单词，10及以上用数字", "三本书；15个人", "three books; 15 people", "digit_word"),
    RuleCase("数字/序数词", "科学数据/单位用数字", "7微克/立方米", "7 μg/m3", "scientific_unit"),
    RuleCase("标题序号", "一、二、三 译为 I, II, III", "一、背景", "I. Background", "heading_roman"),
    RuleCase("百分点", "percentage point(s) 注意单复数", "1个百分点；1.2个百分点", "1 percentage point; 1.2 percentage points", "percentage_point"),
    RuleCase("分数", "用单词，形容词加hyphen", "约占1/6；三分之二的减员", "approximately one sixth; two-thirds reduction in staff", "fraction"),
    RuleCase("日期", "月日, 年 / 月年", "1998年6月4日；1998年6月", "June 4, 1998; June 1998", "date"),
    RuleCase("日期", "年代前加 the", "20世纪30年代", "the 1930s", "decade"),
    RuleCase("日期", "纪年与公元", "公元前221年；公元210年", "221 BC; AD 210", "era"),
    RuleCase("金额", "USD/RMB + 空格 + 千分位", "1,221,000美元", "USD 1,221,000", "money"),
    RuleCase("金额", "小数点保留2位，超过则换单位", "12.215百万美元", "USD 12,215,000", "money"),
    RuleCase("单位", "全称原则，除非太长或kWh/μg/m2", "260.5千米", "260.5 kilometers", "unit_fullname"),
    RuleCase("单位", "表格单位格式", "单位：亿元", "In RMB ’00 Million", "table_unit"),
    RuleCase("经纬度", "N/E，秒用两个单引号", "北纬22°26′59″", "22°26’59’’ N", "coordinate"),
    RuleCase("含义数字", "意译，不直译数字", "“9+2”城市群", "nine mainland cities and two special administrative regions in the Greater Bay Area", "free_translation"),
    RuleCase("期数", "Phase + 大写罗马数字", "项目二期", "Project Phase II", "phase"),
    RuleCase("上下标", "PM2.5下标", "PM2.5", "PM2.5 (with subscript 2.5)", "subscript"),
    RuleCase("句首数字", "不能数字开头", "16,059家企业被纳入...", "A total of 16,059 enterprises were included...", "sentence_start"),
]


def run_rule_engine(cases: List[RuleCase]) -> List[Dict[str, object]]:
    results = []
    for index, case in enumerate(cases, start=1):
        checker = CHECKERS[case.checker]
        result = checker(case)
        results.append(
            {
                "index": index,
                "case": case,
                "result": result,
            }
        )
    return results


def print_report(results: List[Dict[str, object]]) -> None:
    passed_count = sum(1 for item in results if item["result"].passed)
    print("=" * 100)
    print("规则样例提取与判断结果")
    print("=" * 100)
    print(f"总样例数: {len(results)}")
    print(f"通过数量: {passed_count}")
    print(f"失败数量: {len(results) - passed_count}")

    for item in results:
        case: RuleCase = item["case"]
        result: CheckResult = item["result"]
        status = "PASS" if result.passed else "FAIL"
        print("\n" + "-" * 100)
        print(f"[{item['index']:02d}] {status} | {case.category} | {case.rule}")
        print(f"中文: {case.cn}")
        print(f"英文: {case.en}")
        print(f"说明: {result.reason}")
        print(f"提取: {result.extracted}")


if __name__ == "__main__":
    report = run_rule_engine(RULE_CASES)
    print_report(report)
