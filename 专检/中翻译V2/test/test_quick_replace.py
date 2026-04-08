"""
快速替换测试

直接使用内嵌的错误数据进行测试
"""

from docx import Document
from pathlib import Path
import sys

# 获取当前文件的父目录的父目录（即 "中翻译" 文件夹）
project_root = str(Path(__file__).resolve().parent.parent)

# 将项目根目录添加到 sys.path
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from replace.word.revision import RevisionManager
from replace.word.replace_revision import replace_and_revise_in_docx, flush_footnote_replacements
from backup_copy.backup_manager import ensure_backup_copy


# 测试数据 - 直接从你的 JSON 复制

# TEST_ERRORS =[{
#     "错误编号": "1",
#     "错误类型": "层级错误",
#     "原文数值": "（2）",
#     "译文数值": "（2） ii.",
#     "译文修改建议值": "ii.",
#     "修改理由": "连续编号混用不同符号及重号",
#     "原文上下文": "（2） 加速转型升级，消保数智化建设实现突破",
#     "译文上下文": "（2） ii. Accelerating transformation and upgrading, and achieving breakthroughs in the digital and intelligent development of consumer protection",
#     "替换锚点": "（2） ii."
# },
# {
#     "错误编号": "2",
#     "错误类型": "层级错误",
#     "原文数值": "（3）",
#     "译文数值": "（3） iii.",
#     "译文修改建议值": "iii.",
#     "修改理由": "连续编号混用不同符号及重号",
#     "原文上下文": "（3） 坚持系统思维，投诉治理攻坚取得扎实成效",
#     "译文上下文": "（3） iii. Adhering to systematic thinking and achieving solid results in tackling complaint management challenges",
#     "替换锚点": "（3） iii."
# },
# {
#     "错误编号": "3",
#     "错误类型": "层级错误",
#     "原文数值": "（1）",
#     "译文数值": "（1） i.",
#     "译文修改建议值": "i.",
#     "修改理由": "连续编号混用不同符号及重号",
#     "原文上下文": "（1） 消保监管评价情况",
#     "译文上下文": "（1） i. Regulatory evaluation",
#     "替换锚点": "（1） i."
# },
# {
#     "错误编号": "4",
#     "错误类型": "层级错误",
#     "原文数值": "（2）",
#     "译文数值": "（2） ii.",
#     "译文修改建议值": "ii.",
#     "修改理由": "连续编号混用不同符号及重号",
#     "原文上下文": "（2） 内部审计情况",
#     "译文上下文": "（2） ii. Internal audits",
#     "替换锚点": "（2） ii."
# }]

TEST_ERRORS = [
    {
        "错误编号": "1",
        "错误类型": "数值错误",
        "原文数值": "二〇二六年四月",
        "译文数值": "April 2026",
        "译文修改建议值": "April 2026",
        "修改理由": "日期格式符合Month Y，但译文漏译了落款处的具体日期占位符“XX日”。",
        "违反的规则": "规则(一)：检查译文全文是否与原文数值不一致，是否漏译。",
        "原文上下文": "（消费者权益保护办公室）                                                二〇二六年四月XX日",
        "译文上下文": "(Consumer Protection Office)                                                April 2026",
        "替换锚点": "April 2026"
    },
    {
        "错误编号": "2",
        "错误类型": "标题与编号层级",
        "原文数值": "（二）",
        "译文数值": "（二） ii.",
        "译文修改建议值": "ii.",
        "修改理由": "译文保留了中文全角括号序号，违反了英文译文不得出现全角中文符号的规则，且二级标题应仅为ii.。",
        "违反的规则": "规则(二)：二级：i. ii. iii.；规则(九)：英文译文不得出现全角中文符号。",
        "原文上下文": "（二） 加速转型升级，消保数智化建设实现突破",
        "译文上下文": "（二） ii. Accelerating transformation and upgrading, and achieving breakthroughs in the digital and intelligent development of consumer protection",
        "替换锚点": "（二） ii."
    },
    {
        "错误编号": "3",
        "错误类型": "标题与编号层级",
        "原文数值": "（三）",
        "译文数值": "（三） iii.",
        "译文修改建议值": "iii.",
        "修改理由": "译文保留了中文全角括号序号，违反了规则。",
        "违反的规则": "规则(二)：二级：i. ii. iii.；规则(九)：英文译文不得出现全角中文符号。",
        "原文上下文": "（三） 坚持系统思维，投诉治理攻坚取得扎实成效",
        "译文上下文": "（三） iii. Adhering to systematic thinking and achieving solid results in tackling complaint management challenges",
        "替换锚点": "（三） iii."
    },
    {
        "错误编号": "4",
        "错误类型": "数值错误",
        "原文数值": "1.3亿",
        "译文数值": "130 million",
        "译文修改建议值": "130 million",
        "修改理由": "原文为1.3亿，译文130 million数值正确，但根据规则(六)，1-10用单词，11+用数字，此处130为数字正确，但需注意全文并列关系。此处主要检查单位转换是否导致数值变动，1.3亿=130 million无误。",
        "违反的规则": "规则(一)：零误差原则。",
        "原文上下文": "超额完成1.3亿客户营销授权采集，客户信息安全权保障能力显著增强。",
        "译文上下文": "Marketing authorizations from 130 million customers were collected, surpassing the set target and significantly strengthening the protection of customers’ information security rights.",
        "替换锚点": "130 million"
    },
    {
        "错误编号": "5",
        "错误类型": "数值错误",
        "原文数值": "11.19亿",
        "译文数值": "1,119 million",
        "译文修改建议值": "1,119 million",
        "修改理由": "原文11.19亿，译文1,119 million。根据规则(五)，能用最大金额单位不用小一级，1.119 billion更优，但若为了保留小数点后两位且不四舍五入，1,119 million数值相等。此处译文数值正确。",
        "违反的规则": "规则(五)：能用最大金额单位不用小一级。",
        "原文上下文": "2025年全行累计开展金融教育宣传活动20.43万次，触达消费者超11.19亿人次",
        "译文上下文": "In 2025, the Bank conducted a total of 204,300 financial education activities, reaching more than 1,119 million consumers",
        "替换锚点": "1,119 million"
    },
    {
        "错误编号": "6",
        "错误类型": "标题与编号层级",
        "原文数值": "（一）",
        "译文数值": "（一） i.",
        "译文修改建议值": "i.",
        "修改理由": "保留了中文全角括号序号。",
        "违反的规则": "规则(二)：二级：i. ii. iii.；规则(九)：英文译文不得出现全角中文符号。",
        "原文上下文": "（一） 消保监管评价情况",
        "译文上下文": "（一） i. Regulatory evaluation",
        "替换锚点": "（一） i."
    },
    {
        "错误编号": "7",
        "错误类型": "标题与编号层级",
        "原文数值": "（二）",
        "译文数值": "（二） ii.",
        "译文修改建议值": "ii.",
        "修改理由": "保留了中文全角括号序号。",
        "违反的规则": "规则(二)：二级：i. ii. iii.；规则(九)：英文译文不得出现全角中文符号。",
        "原文上下文": "（二） 内部审计情况",
        "译文上下文": "（二） ii. Internal audits",
        "替换锚点": "（二） ii."
    },
    {
        "错误编号": "8",
        "错误类型": "标题与编号层级",
        "原文数值": "（三）",
        "译文数值": "（三） iii.",
        "译文修改建议值": "iii.",
        "修改理由": "保留了中文全角括号序号。",
        "违反的规则": "规则(二)：二级：i. ii. iii.；规则(九)：英文译文不得出现全角中文符号。",
        "原文上下文": "（三） 消保考核情况",
        "译文上下文": "（三） iii. Consumer protection appraisals",
        "替换锚点": "（三） iii."
    },
    {
        "错误编号": "9",
        "错误类型": "数值错误",
        "原文数值": "2025年",
        "译文数值": "2024年",
        "译文修改建议值": "2025",
        "修改理由": "原文为2025年，译文写成2024年，且保留了中文字符“年”。",
        "违反的规则": "规则(一)：检查译文全文是否与原文数值不一致；规则(九)：不得出现全角中文符号/字符。",
        "原文上下文": "2025年境内分行经营绩效消保指标设置“消保内外部评价”",
        "译文上下文": "In 2024, ICBC’s domestic branches set performance indicators for consumer protection",
        "替换锚点": "2024年"
    },
    {
        "错误编号": "10",
        "错误类型": "标题与编号层级",
        "原文数值": "（一）",
        "译文数值": "（一） i.",
        "译文修改建议值": "i.",
        "修改理由": "保留了中文全角括号序号。",
        "违反的规则": "规则(二)：二级：i. ii. iii.；规则(九)：英文译文不得出现全角中文符号。",
        "原文上下文": "（一） 一以贯之深化“大消保”格局",
        "译文上下文": "（一） i. Consistently deepening the \"Greater Consumer Protection\" framework",
        "替换锚点": "（一） i."
    },
    {
        "错误编号": "11",
        "错误类型": "标题与编号层级",
        "原文数值": "（二）",
        "译文数值": "（二） ii.",
        "译文修改建议值": "ii.",
        "修改理由": "保留了中文全角括号序号。",
        "违反的规则": "规则(二)：二级：i. ii. iii.；规则(九)：英文译文不得出现全角中文符号。",
        "原文上下文": "（二） 系统构建“数智消保”赋能体系",
        "译文上下文": "（二） ii. Systematically building a \"digital and intelligent consumer protection\" empowerment system",
        "替换锚点": "（二） ii."
    },
    {
        "错误编号": "12",
        "错误类型": "标题与编号层级",
        "原文数值": "（三）",
        "译文数值": "（三） iii.",
        "译文修改建议值": "iii.",
        "修改理由": "保留了中文全角括号序号。",
        "违反的规则": "规则(二)：二级：i. ii. iii.；规则(九)：英文译文不得出现全角中文符号。",
        "原文上下文": "（三） 全力推动投诉治理再上新台阶",
        "译文上下文": "（三） iii. Fully advancing complaint management to a new level",
        "替换锚点": "（三） iii."
    },
    {
        "错误编号": "13",
        "错误类型": "标题与编号层级",
        "原文数值": "（四）",
        "译文数值": "（四） iv.",
        "译文修改建议值": "iv.",
        "修改理由": "保留了中文全角括号序号。",
        "违反的规则": "规则(二)：二级：i. ii. iii.；规则(九)：英文译文不得出现全角中文符号。",
        "原文上下文": "（四） 全面提升消保助力业务发展价值创造力",
        "译文上下文": "（四） iv. Comprehensively enhancing the value contribution of consumer protection in supporting business development",
        "替换锚点": "（四） iv."
    },
    {
        "错误编号": "14",
        "错误类型": "数值错误",
        "原文数值": "69.52%",
        "译文数值": "71.07%",
        "译文修改建议值": "69.52%",
        "修改理由": "原文数值为69.52%，译文错误写成71.07%。",
        "违反的规则": "规则(一)：检查译文全文是否与原文数值不一致。",
        "原文上下文": "合计占比69.52%。",
        "译文上下文": "These issues together account for 71.07% of all complaints.",
        "替换锚点": "71.07%"
    }
]


def quick_test(doc_path: str, test_cases: list = None):
    """快速测试替换功能"""
    
    if test_cases is None:
        test_cases = TEST_ERRORS
    
    print("=" * 80)
    print("快速替换测试")
    print("=" * 80)
    
    if not Path(doc_path).exists():
        print(f"❌ 文档不存在: {doc_path}")
        return
    
    print(f"\n📂 文档: {doc_path}")
    print(f"📊 测试用例: {len(test_cases)} 个")
    
    # 创建备份
    print(f"\n📦 创建备份...")
    backup_path = ensure_backup_copy(doc_path, suffix="quicktest")
    print(f"✓ 备份: {backup_path}")

    # 自动编号静态化：把所有自动编号转成静态文本，避免改一个编号后面全乱
    try:
        from replace.word.numbering_to_static import convert_numbering_to_static, has_auto_numbering
        if has_auto_numbering(backup_path):
            print(f"\n🔢 检测到自动编号，正在转换为静态文本...")
            ok = convert_numbering_to_static(backup_path)
            if ok:
                print(f"✓ 自动编号已转为静态文本")
            else:
                print(f"⚠ 自动编号静态化失败，部分编号可能无法替换")
    except Exception as e:
        print(f"⚠ 编号静态化异常: {e}")

    # 页眉页脚域代码静态化（PAGE 等域 → 实际页码数字）
    # try:
    #     from replace.word.numbering_to_static import unlink_page_fields_xml
    #     print(f"\n📄 正在将页眉页脚域代码转为静态文本（页码等）...")
    #     ok = unlink_page_fields_xml(backup_path)
    #     if ok:
    #         print(f"✓ 页眉页脚域代码已静态化")
    #         from parsers.word.header_extractor import extract_headers
    #         from parsers.word.footer_extractor import extract_footers
    #         headers = extract_headers(backup_path)
    #         footers = extract_footers(backup_path)
    #         print(f"  静态化后页眉: {headers}")
    #         print(f"  静态化后页脚: {footers}")
    #     else:
    #         print(f"⚠ 域代码静态化失败")
    # except Exception as e:
    #     print(f"⚠ 域代码静态化异常: {e}")
    
    # 打开文档（此时自动编号已经是静态文本了）
    doc = Document(backup_path)
    doc._numbering_staticized = True  # 标记已静态化
    revision_manager = RevisionManager(doc, author="快速测试")
    
    # 执行替换
    print(f"\n🔄 开始测试...")
    print("=" * 80)
    
    results = []
    
    for idx, error in enumerate(test_cases, 1):
        old_text = (error.get("译文数值") or "").strip()
        new_text = (error.get("译文修改建议值") or "").strip()
        reason = error.get("修改理由", "")
        context = error.get("译文上下文", "")
        anchor = error.get("替换锚点", "")
        error_no = error.get("错误编号", idx)
        
        if not old_text or not new_text:
            results.append((error_no, "跳过", "缺少数据"))
            continue
        
        print(f"\n[测试 {idx}/{len(test_cases)}] 错误编号: {error_no}")
        print(f"  查找: '{old_text[:60]}...'")
        print(f"  替换: '{new_text[:60]}...'")
        
        try:
            ok, strategy = replace_and_revise_in_docx(
                doc, old_text, new_text, reason, revision_manager,
                context=context, anchor_text=anchor, region="body",
                doc_path=backup_path
            )
            
            if ok:
                results.append((error_no, "成功", strategy))
                print(f"  ✓ 成功: {strategy}")
            else:
                results.append((error_no, "失败", strategy))
                print(f"  ✗ 失败: {strategy}")
        
        except Exception as e:
            results.append((error_no, "异常", str(e)))
            print(f"  ✗ 异常: {e}")
    
    # 保存
    print(f"\n💾 保存文档...")
    doc.save(backup_path)
    
    # 执行脚注替换（必须在doc.save()之后）
    footnote_count = flush_footnote_replacements(doc, backup_path)
    if footnote_count > 0:
        print(f"✓ 脚注替换完成: {footnote_count} 处")
    
    # 统计
    print("\n" + "=" * 80)
    print("测试结果")
    print("=" * 80)
    
    success = sum(1 for _, status, _ in results if status == "成功")
    fail = sum(1 for _, status, _ in results if status == "失败")
    skip = sum(1 for _, status, _ in results if status == "跳过")
    error = sum(1 for _, status, _ in results if status == "异常")
    
    print(f"\n✓ 成功: {success}")
    print(f"✗ 失败: {fail}")
    print(f"⊘ 跳过: {skip}")
    print(f"⚠ 异常: {error}")
    print(f"━ 总计: {len(results)}")
    
    if success + fail > 0:
        print(f"\n成功率: {success / (success + fail):.1%}")
    
    # 详细结果
    print("\n详细结果:")
    for error_no, status, detail in results:
        symbol = {"成功": "✓", "失败": "✗", "跳过": "⊘", "异常": "⚠"}.get(status, "?")
        print(f"  {symbol} 错误 {error_no}: {status} - {detail[:60]}...")
    
    print(f"\n✅ 测试完成！")
    print(f"📄 结果文档: {backup_path}")
    
    return results


if __name__ == "__main__":
    # 默认文档路径 - 应该使用译文，不是原文
    DOC_PATH = r"../测试文件/译文-B260328387-关于消费者权益保护2025年工作情况与2026年工作计划的议案.docx"
    
    # 如果提供了命令行参数，使用它
    if len(sys.argv) > 1:
        DOC_PATH = sys.argv[1]
    
    print(f"使用文档: {DOC_PATH}")
    
    # 检查文档是否存在
    if not Path(DOC_PATH).exists():
        print(f"\n❌ 文档不存在: {DOC_PATH}")
        print("\n💡 提示:")
        print("  - 测试数据包含英文译文，需要使用译文文档")
        print("  - 如果文档在其他位置，请提供完整路径:")
        print(f"    python {Path(__file__).name} \"你的译文文档.docx\"")
        sys.exit(1)
    
    print()
    
    # 运行测试
    quick_test(DOC_PATH)
