import unicodedata
import re
from typing import Optional
from docx import Document
from zhongfanyi.llm.llm_project.backup_copy.backup_copy import ensure_backup_copy
from zhongfanyi.llm.llm_project.llm_check.check import Match
from zhongfanyi.llm.llm_project.note.pizhu import CommentManager
from zhongfanyi.llm.llm_project.parsers.word.body_extractor import extract_body_text
from zhongfanyi.llm.llm_project.parsers.word.footer_extractor import extract_footers
from zhongfanyi.llm.llm_project.parsers.word.header_extractor import extract_headers
from zhongfanyi.llm.llm_project.replace.word.replace_clean import replace_and_comment_in_docx
from zhongfanyi.llm.llm_project.parsers.json.clean_json import extract_and_parse
from zhongfanyi.llm.llm_project.utils.json_files import write_json_with_timestamp


# =========================
# 1) 文本清洗工具
# =========================

def clean_text_thoroughly(text: str) -> str:
    """
    终极清洗：消除所有隐形字符、非标空格和格式干扰。

    Args:
        text: 待清洗的文本

    Returns:
        清洗后的标准文本
    """
    if not text:
        return ""

    # 1. Unicode 标准化 (NFKC)
    text = unicodedata.normalize('NFKC', text)

    # 2. 移除零宽字符和控制字符
    text = re.sub(r'[\u200b-‍﻿­]', '', text)

    # 3. 统一所有空白字符为标准空格
    text = re.sub(r'\s+', ' ', text)

    return text.strip()


def _normalize_spaces(text: str) -> str:
    """
    标准化空格（保留内部逻辑）

    Args:
        text: 待处理文本

    Returns:
        空格标准化后的文本
    """
    return re.sub(r'\s+', ' ', text).strip()


# =========================
# 2) 编号模式识别（修复版）
# =========================

def is_list_pattern(s: str) -> bool:
    """
    判断是否为编号模式

    支持格式：
    - 数字编号：(1), 1), 1.
    - 罗马数字：i., ii., iii., iv., v. 等
    - 字母编号：a), b), (a), (b) 等

    Args:
        s: 待判断字符串

    Returns:
        是否为编号模式
    """
    s_clean = s.strip().lower()

    # 修复后的正则表达式（移除重复的 ^）
    patterns = [
        r'^\(\d+\)$',  # (1), (2)
        r'^\d+\)$',  # 1), 2)
        r'^\d+\.$',  # 1., 2.
        r'^[ivxlcdm]+\.$',  # i., ii., iii., iv., v. 等（扩展罗马数字范围）
        r'^\([a-z]\)$',  # (a), (b)
        r'^[a-z]\)$',  # a), b)
        r'^[a-z]\.$',  # a., b.
    ]

    return any(re.match(pattern, s_clean) for pattern in patterns)


# =========================
# 3) 智能匹配模式构建（核心修复）
# =========================

def build_smart_pattern(s: str, mode: str = "balanced") -> str:
    """
    构建智能匹配模式

    Args:
        s: 待匹配字符串
        mode: 匹配模式
            - "strict": 严格匹配（完全精确）
            - "balanced": 平衡模式（数字/标点保持连续，单词间允许空格）
            - "loose": 宽松模式（字符间允许空格，但数字连续）

    Returns:
        正则表达式模式字符串
    """
    # 【关键修复】先清洗输入
    s = clean_text_thoroughly(s or "")
    if not s:
        return ""

    if mode == "strict":
        # 严格模式：完全精确匹配（转义特殊字符）
        return re.escape(s)

    elif mode == "balanced":
        # 平衡模式：数字和标点保持连续，单词间允许空格
        pieces = []
        i = 0

        while i < len(s):
            ch = s[i]

            # 跳过空格（在字符间添加可选空格）
            if ch.isspace():
                if pieces and not pieces[-1].endswith(r"\s*"):
                    pieces.append(r"\s*")
                i += 1
                continue

            # 数字序列：保持连续（包括小数点、逗号）
            if ch.isdigit():
                num_str = ""
                while i < len(s) and (s[i].isdigit() or s[i] in ".,"):
                    num_str += s[i]
                    i += 1
                pieces.append(re.escape(num_str))
                continue

            # 【关键修复】标点符号：保持连续，但不自动添加后续空格
            # 原因：编号如 "i." 后面可能紧跟内容，不应强制匹配空格
            if ch in ".,;:!?()[]{}\"'-/":
                pieces.append(re.escape(ch))
                i += 1
                # 仅在下一个字符是字母时才添加可选空格
                if i < len(s) and s[i].isalpha():
                    pieces.append(r"\s*")
                continue

            # 字母序列：单词间允许空格
            if ch.isalpha():
                word = ""
                while i < len(s) and s[i].isalpha():
                    word += s[i]
                    i += 1
                pieces.append(re.escape(word))
                # 如果后面还有内容且不是标点，添加可选空格
                if i < len(s) and s[i] not in ".,;:!?()[]{}\"'-/":
                    pieces.append(r"\s*")
                continue

            # 其他字符
            pieces.append(re.escape(ch))
            i += 1

        return "".join(pieces).strip()

    else:  # loose
        # 宽松模式：字符间允许空格，但数字保持连续
        pieces = []
        i = 0

        while i < len(s):
            ch = s[i]

            if ch.isspace():
                i += 1
                continue

            # 数字序列保持连续
            if ch.isdigit():
                num_str = ""
                while i < len(s) and (s[i].isdigit() or s[i] in ".,"):
                    num_str += s[i]
                    i += 1
                pieces.append(re.escape(num_str) + r"\s*")
                continue

            # 其他字符间允许空格
            pieces.append(re.escape(ch) + r"\s*")
            i += 1

        return "".join(pieces).strip()


# =========================
# 4) 锚点提取（增强版）
# =========================

def extract_anchor_with_target(
        context: str,
        target_value: str,
        window: int = 30
) -> Optional[str]:
    """
    从上下文中提取包含目标数值的锚点短语

    Args:
        context: 上下文文本
        target_value: 目标值（如数字、编号等）
        window: 锚点窗口大小（前后字符数）

    Returns:
        提取的锚点短语，未找到则返回 None
    """
    # 【关键修复】在最开始就清洗
    context = clean_text_thoroughly(context)
    target_value = clean_text_thoroughly(target_value)

    if not context or not target_value:
        return None

    context = _normalize_spaces(context)
    target_value = target_value.strip()

    # 步骤1：尝试严格匹配
    if target_value in context:
        idx = context.index(target_value)
        start = max(0, idx - window)
        end = min(len(context), idx + len(target_value) + window)
        anchor = context[start:end]

        # 修剪边界单词（避免截断）
        anchor = re.sub(r"^\S*\s+", "", anchor)
        anchor = re.sub(r"\s+\S*$", "", anchor)
        return anchor.strip()

    # 步骤2：【关键修复】编号模式强制使用 strict 模式
    if is_list_pattern(target_value):
        pattern = re.escape(target_value)
    else:
        pattern = build_smart_pattern(target_value, mode="balanced")

    if not pattern:
        return None

    # 步骤3：正则匹配
    match = re.search(pattern, context, flags=re.IGNORECASE)
    if not match:
        return None

    start, end = match.span()
    prefix_start = max(0, start - window)
    suffix_end = min(len(context), end + window)

    anchor = context[prefix_start:suffix_end]

    # 修剪边界单词
    anchor = re.sub(r"^\S*\s+", "", anchor)
    anchor = re.sub(r"\s+\S*$", "", anchor)

    return anchor.strip() if anchor.strip() else None


# =========================
# 5) 测试用例
# =========================

if __name__ == "__main__":
    # 示例文件路径
    original_path = r"C:\Users\Administrator\Desktop\项目文件\专检\数值检查\测试文件\B251124195-Y-更新1121-附件1：中国银行股份有限公司模型风险管理办法（2025年修订）.docx"  # 请替换为原文文件路径
    translated_path = r"C:\Users\Administrator\Desktop\项目文件\专检\数值检查\测试文件\清洁版-B251124195-附件1：中国银行股份有限公司模型风险管理政策（2025年修订）-.docx"  # 请替换为译文文件路径
    # 处理页眉
    original_header_text = extract_headers(original_path)
    translated_header_text = extract_headers(translated_path)
    # 处理页脚
    original_footer_text = extract_footers(original_path)
    translated_footer_text = extract_footers(translated_path)
    # 处理正文(含脚注/表格/自动编号)
    original_body_text = extract_body_text(original_path)
    translated_body_text = extract_body_text(translated_path)
    print("======页眉===========")
    print(original_header_text)
    print(translated_header_text)
    print("======页脚===========")
    print(original_footer_text)
    print(translated_footer_text)
    print("======正文===========")
    print(original_body_text)
    print(translated_body_text)
    # 实例化对象并进行对比
    matcher = Match()
    # 正文对比
    print("======正在检查正文===========")
    if original_body_text and translated_body_text:
        # 两个值都不为空，正常执行比较
        body_result = matcher.compare_texts(original_body_text, translated_body_text)
    else:
        # 任意一个为空，生成空结果
        body_result = {}  # 或者 body_result = []，根据你的 write_json_with_timestamp 函数期望的格式
        print("原文或译文为空，检查结果为空")

    body_result_name, body_result_path = write_json_with_timestamp(
        body_result,
        r"C:\Users\Administrator\Desktop\project\llm\llm_project\zhengwen\output_json"
    )
    # body_result = matcher.compare_texts(original_body_text, translated_body_text)
    # body_result_name, body_result_path = write_json_with_timestamp(body_result,r"C:\Users\Administrator\Desktop\project\llm\llm_project\zhengwen\output_json")
    # #页眉对比
    print("======正在检查页眉===========")
    if original_header_text and translated_header_text:
        # 两个值都不为空，正常执行比较
        header_result = matcher.compare_texts(original_header_text, translated_header_text)
    else:
        # 任意一个为空，生成空结果
        header_result = {}
        print("原文或译文为空，检查结果为空")

    header_result_name, header_result_path = write_json_with_timestamp(
        header_result,
        r"C:\Users\Administrator\Desktop\project\llm\llm_project\yemei\output_json"
    )
    # header_result = matcher.compare_texts(original_header_text, translated_header_text)
    # header_result_name, header_result_path = write_json_with_timestamp(header_result, r"C:\Users\Administrator\Desktop\project\llm\llm_project\yemei\output_json")
    # #页脚对比
    print("======正在检查页脚===========")
    if original_footer_text and translated_footer_text:
        # 两个值都不为空，正常执行比较
        footer_result = matcher.compare_texts(original_footer_text, translated_footer_text)
    else:
        # 任意一个为空，生成空结果
        footer_result = {}
        print("原文或译文为空，检查结果为空")

    footer_result_name, footer_result_path = write_json_with_timestamp(
        footer_result,
        r"C:\Users\Administrator\Desktop\project\llm\llm_project\yejiao\output_json"
    )
    # footer_result = matcher.compare_texts(original_footer_text, translated_footer_text)
    # footer_result_name, footer_result_path = write_json_with_timestamp(footer_result, r"C:\Users\Administrator\Desktop\project\llm\llm_project\yejiao\output_json")

    print("================================")
    # if not os.path.exists(error_docx_path):
    #     raise FileNotFoundError(f"错误报告文件不存在: {error_docx_path}")
    # body_result_path=r"C:\Users\Administrator\Desktop\project\llm\文本对比结果\加粗斜体测试\AI检查结果.json"
    # header_result_path=r"C:\Users\Administrator\Desktop\project\llm\llm_project\yemei\output_json\文本对比结果_20260208_144950.json"
    # footer_result_path=r"C:\Users\Administrator\Desktop\project\llm\llm_project\yejiao\output_json\文本对比结果_20260208_145004.json"

    # 1) 复制译文到 backup/
    backup_copy_path = ensure_backup_copy(translated_path)
    print(f"✅ 已复制译文副本到: {backup_copy_path}")

    # 2) 读取错误报告并解析
    print("\n正在提取解析正文错误报告...")
    body_errors = extract_and_parse(body_result_path)
    print("正文错误报告", body_errors)
    for err in body_errors:
        print(err)
    print("正文错误解析个数：", len(body_errors))

    print("\n正在提取解析页眉错误报告...")
    header_errors = extract_and_parse(header_result_path)
    print("页眉错误报告", header_errors)
    print("页眉错误解析个数：", len(header_errors))

    print("\n正在提取解析页脚错误报告...")
    footer_errors = extract_and_parse(footer_result_path)
    print("页脚错误报告", footer_errors)
    print("页脚错误解析个数：", len(footer_errors))

    # 3) 打开副本 docx
    print("正在加载文档...")
    doc = Document(backup_copy_path)

    # 4) 创建批注管理器并初始化
    print("正在初始化批注系统...")
    comment_manager = CommentManager(doc)

    # 【关键】创建初始批注以确保 comments.xml 结构完整
    if comment_manager.create_initial_comment():
        print("✓ 批注系统初始化成功\n")
    else:
        print("⚠️ 批注系统初始化失败，但将继续尝试处理\n")

    # 5) 逐条执行替换并添加批注
    print("==================== 开始处理正文错误 ====================\n")
    body_success_count = 0
    body_fail_count = 0

    for idx, e in enumerate(body_errors, 1):
        err_id = e.get("错误编号", "?")
        err_type = e.get("错误类型", "")
        old = (e.get("译文数值") or "").strip()
        new = (e.get("译文修改建议值") or "").strip()
        reason = (e.get("修改理由"),"")
        trans_context = e.get("译文上下文", "") or ""
        anchor = e.get("替换锚点", "") or ""

        if not old or not new:
            print(f"[{idx}/{len(body_errors)}] [跳过] 错误 #{err_id}: old/new 缺失")
            body_fail_count += 1
            continue

        # 执行替换并添加批注(正文)
        ok, strategy = replace_and_comment_in_docx(
            doc, old, new, reason,comment_manager,
            context=trans_context,
            anchor_text=anchor
        )

        if ok:
            print(f"[{idx}/{len(body_errors)}] [✓成功] 错误 #{err_id} ({err_type})")
            print(f"    策略: {strategy}")
            print(f"    修改理由: {reason}")
            print(f"    操作: '{old}' → '{new}' (已替换并添加批注)")
            if anchor:
                print(f"    锚点: {anchor}...")
            elif trans_context:
                print(f"    上下文: {trans_context}...")
            body_success_count += 1
        else:
            print(f"[{idx}/{len(body_errors)}] [✗失败] 错误 #{err_id} ({err_type})")
            print(f"    未找到匹配: '{old}'")
            if anchor:
                print(f"    锚点: {anchor}...")
            print(f"    上下文: {trans_context if trans_context else '无'}...")
            body_fail_count += 1
        print()

    print(f"\n==================== 正文处理完成 ====================")
    print(f"成功: {body_success_count} | 失败: {body_fail_count} | 总计: {len(body_errors)}")
    if len(body_errors) > 0:
        print(f"成功率: {body_success_count / len(body_errors) * 100:.1f}%")
    print(f"\n✅ 已保存到: {backup_copy_path}")
    print("⚠️ 原始译文文件未被修改")

    print("==================== 开始处理页眉错误 ====================\n")
    header_success_count = 0
    header_fail_count = 0



    # # 测试文本（模拟 PDF 提取的复杂格式）
    # test_context = """
    # This is a test document with various formats:
    #
    # i. First item with roman numeral
    # ii. Second item continues here
    # iii. Third item with more text
    #
    # Regular numbers:
    # (1) Parenthesized number
    # 1) Simple number with bracket
    # 1. Dotted number format
    #
    # Special cases:
    # Total: $1,234.56 million
    # Date: 2024-01-15
    # """
    #
    # # 测试用例
    # test_cases = [
    #     ("i.", "罗马数字 i."),
    #     ("ii.", "罗马数字 ii."),
    #     ("iii.", "罗马数字 iii."),
    #     ("(1)", "括号数字 (1)"),
    #     ("1)", "数字加括号 1)"),
    #     ("1.", "数字加点 1."),
    #     ("$1,234.56", "货币金额"),
    #     ("2024-01-15", "日期格式"),
    # ]
    #
    # print("=" * 60)
    # print("锚点提取测试")
    # print("=" * 60)
    #
    # for target, desc in test_cases:
    #     result = extract_anchor_with_target(test_context, target, window=40)
    #     status = "✓ 成功" if result else "✗ 失败"
    #     print(f"\n[{desc}] {status}")
    #     print(f"  目标值: '{target}'")
    #     if result:
    #         print(f"  锚点: '{result}'")
    #     else:
    #         print(f"  原因: 未匹配到 '{target}'")
    #
    # print("\n" + "=" * 60)
    # print("模式构建测试")
    # print("=" * 60)
    #
    # for target, desc in test_cases:
    #     is_list = is_list_pattern(target)
    #     pattern = build_smart_pattern(target, mode="balanced")
    #     print(f"\n[{desc}]")
    #     print(f"  原始值: '{target}'")
    #     print(f"  是否编号: {is_list}")
    #     print(f"  生成模式: {pattern}")