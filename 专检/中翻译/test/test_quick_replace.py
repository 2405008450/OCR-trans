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
import difflib


def _fix_suggestion_overlap(context: str, anchor: str, suggestion: str) -> str:
    """检测 suggestion 与上下文的前/后重叠，返回修正后的实际替换值。"""
    if not context or not anchor or not suggestion:
        return suggestion
    if anchor not in context:
        return suggestion

    start_idx = context.find(anchor)
    end_idx = start_idx + len(anchor)
    actual_sug = suggestion

    # 前部去重
    prefix_window = context[max(0, start_idx - len(suggestion)):start_idx]
    s = difflib.SequenceMatcher(None, prefix_window, suggestion)
    match = s.find_longest_match(0, len(prefix_window), 0, len(suggestion))
    if match.size > 0 and (match.a + match.size == len(prefix_window)) and (match.b == 0):
        actual_sug = suggestion[match.size:]

    # 后部去重
    after_window = context[end_idx: end_idx + len(actual_sug)]
    s = difflib.SequenceMatcher(None, actual_sug, after_window)
    match = s.find_longest_match(0, len(actual_sug), 0, len(after_window))
    if match.size > 0 and (match.a + match.size == len(actual_sug)) and (match.b == 0):
        actual_sug = actual_sug[:match.a]

    return actual_sug


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
    "原文上下文": "☆5年\n中国工商银行股份有限公司\n董事会/董事会战略委员会会议\n议案\n关于《中国工商银行“十五五”时期\n发展规划》的议案\n(审议稿)\n二〇二六年六月",
    "译文上下文": "☆5-Year\nProposal on the<italic> 15th Five-Year Development Plan of ICBC</italic>",
    "原文数值": "中国工商银行股份有限公司\n董事会/董事会战略委员会会议\n议案\n(审议稿)\n二〇二六年六月",
    "译文数值": "☆5-Year\nProposal on the<italic> 15th Five-Year Development Plan of ICBC</italic>",
    "替换锚点": "☆5-Year\nProposal on the<italic> 15th Five-Year Development Plan of ICBC</italic>",
    "译文修改建议值": "☆5-Year\nIndustrial and Commercial Bank of China Limited\nMeeting of the Board of Directors/Strategy Committee of the Board of Directors\nProposal\nProposal on the<italic> 15th Five-Year Development Plan of ICBC</italic>\n(Draft for Deliberation)\nJune 2026",
    "错误类型": "漏译",
    "修改理由": "译文漏译了“中国工商银行股份有限公司”、“董事会/董事会战略委员会会议”、“议案”、“(审议稿)”、“二〇二六年六月”。",
    "违反的规则": "对照原文检查译文出现错译、漏译、多译的情况。"
  },
  {
    "错误编号": "2",
    "原文上下文": "董事会/董事会战略委员会：",
    "译文上下文": "<bold>Board of Directors/Strategy Committee of the Board of Directors:</bold>",
    "原文数值": "董事会/董事会战略委员会：",
    "译文数值": "<bold>Board of Directors/Strategy Committee of the Board of Directors:</bold>",
    "替换锚点": "<bold>Board of Directors/Strategy Committee of the Board of Directors:</bold>",
    "译文修改建议值": "Board of Directors/Strategy Committee of the Board of Directors:",
    "错误类型": "多译标签",
    "修改理由": "原文没有加粗，译文多加了<bold>标签。",
    "违反的规则": "根据原文带\"<bold></bold>\"\"<italic></italic>\"词语检查翻译后对应的译文有无标签，加粗范围和原文保持一致。"
  },
  {
    "错误编号": "3",
    "原文上下文": "一、《规划》稿编制过程",
    "译文上下文": "<bold>I. Preparation Process of the Plan</bold>",
    "原文数值": "一、",
    "译文数值": "I.",
    "替换锚点": "<bold>I. Preparation Process of the Plan</bold>",
    "译文修改建议值": "<bold>I Preparation Process of the Plan</bold>",
    "错误类型": "标题与编号层级",
    "修改理由": "规则要求“二、”翻译成“II”（后面没有点），因此“一、”应翻译成“I”（后面没有点），译文多了一个点。",
    "违反的规则": "标题翻译：“二、”翻译成“II”（后面没有点）"
  },
  {
    "错误编号": "4",
    "原文上下文": "二、《规划》稿基本框架",
    "译文上下文": "<bold>II. Basic Framework of the Plan</bold>",
    "原文数值": "二、",
    "译文数值": "II.",
    "替换锚点": "<bold>II. Basic Framework of the Plan</bold>",
    "译文修改建议值": "<bold>II Basic Framework of the Plan</bold>",
    "错误类型": "标题与编号层级",
    "修改理由": "规则要求“二、”翻译成“II”（后面没有点），译文多了一个点。",
    "违反的规则": "标题翻译：“二、”翻译成“II”（后面没有点）"
  },
  {
    "错误编号": "5",
    "原文上下文": "<bold>同时，</bold>提出规划落地若干支持保障措施",
    "译文上下文": "At the same time, several supporting and safeguarding measures",
    "原文数值": "<bold>同时，</bold>",
    "译文数值": "At the same time,",
    "替换锚点": "At the same time,",
    "译文修改建议值": "<bold>At the same time,</bold>",
    "错误类型": "标签缺失",
    "修改理由": "原文“同时，”有加粗标签，译文漏掉了对应的<bold>标签。",
    "违反的规则": "根据原文带\"<bold></bold>\"\"<italic></italic>\"词语检查翻译后对应的译文有无标签"
  },
  {
    "错误编号": "6",
    "原文上下文": "三、需要说明的重点问题",
    "译文上下文": "<bold>III. Key Issues Requiring Explanation</bold>",
    "原文数值": "三、",
    "译文数值": "III.",
    "替换锚点": "<bold>III. Key Issues Requiring Explanation</bold>",
    "译文修改建议值": "<bold>III Key Issues Requiring Explanation</bold>",
    "错误类型": "标题与编号层级",
    "修改理由": "规则要求“二、”翻译成“II”（后面没有点），因此“三、”应翻译成“III”（后面没有点），译文多了一个点。",
    "违反的规则": "标题翻译：“二、”翻译成“II”（后面没有点）"
  },
  {
    "错误编号": "7",
    "原文上下文": "<bold>综合化方面，</bold>以全面金融解决方案为牵引",
    "译文上下文": "In terms of integrated operation, with comprehensive financial solutions as the guide",
    "原文数值": "<bold>综合化方面，</bold>",
    "译文数值": "In terms of integrated operation,",
    "替换锚点": "In terms of integrated operation,",
    "译文修改建议值": "<bold>In terms of integrated operation,</bold>",
    "错误类型": "标签缺失",
    "修改理由": "原文“综合化方面，”有加粗标签，译文漏掉了对应的<bold>标签。",
    "违反的规则": "根据原文带\"<bold></bold>\"\"<italic></italic>\"词语检查翻译后对应的译文有无标签"
  },
  {
    "错误编号": "8",
    "原文上下文": "<bold>国际化方面，</bold>以全球一体化经营为抓手",
    "译文上下文": "In terms of internationalization, with globally integrated operation as the lever",
    "原文数值": "<bold>国际化方面，</bold>",
    "译文数值": "In terms of internationalization,",
    "替换锚点": "In terms of internationalization,",
    "译文修改建议值": "<bold>In terms of internationalization,</bold>",
    "错误类型": "标签缺失",
    "修改理由": "原文“国际化方面，”有加粗标签，译文漏掉了对应的<bold>标签。",
    "违反的规则": "根据原文带\"<bold></bold>\"\"<italic></italic>\"词语检查翻译后对应的译文有无标签"
  },
  {
    "错误编号": "9",
    "原文上下文": "<bold>数智化方面，</bold>以数据治理和AI大模型应用为牵引",
    "译文上下文": "In terms of digital-intelligent development, with data governance and the application of AI large models as the guide",
    "原文数值": "<bold>数智化方面，</bold>",
    "译文数值": "In terms of digital-intelligent development,",
    "替换锚点": "In terms of digital-intelligent development,",
    "译文修改建议值": "<bold>In terms of digital-intelligent development,</bold>",
    "错误类型": "标签缺失",
    "修改理由": "原文“数智化方面，”有加粗标签，译文漏掉了对应的<bold>标签。",
    "违反的规则": "根据原文带\"<bold></bold>\"\"<italic></italic>\"词语检查翻译后对应的译文有无标签"
  },
  {
    "错误编号": "10",
    "原文上下文": "<bold>同时，</bold>立足夯实基本盘",
    "译文上下文": "At the same time, building on consolidating its fundamentals",
    "原文数值": "<bold>同时，</bold>",
    "译文数值": "At the same time,",
    "替换锚点": "At the same time,",
    "译文修改建议值": "<bold>At the same time,</bold>",
    "错误类型": "标签缺失",
    "修改理由": "原文“同时，”有加粗标签，译文漏掉了对应的<bold>标签。",
    "违反的规则": "根据原文带\"<bold></bold>\"\"<italic></italic>\"词语检查翻译后对应的译文有无标签"
  },
  {
    "错误编号": "11",
    "原文上下文": "<bold>围绕结构优化，</bold>健全信贷投放与国民经济、区域产业相匹配的制度框架",
    "译文上下文": "Centering on structural optimization, the Bank improves the institutional framework",
    "原文数值": "<bold>围绕结构优化，</bold>",
    "译文数值": "Centering on structural optimization,",
    "替换锚点": "Centering on structural optimization,",
    "译文修改建议值": "<bold>Centering on structural optimization,</bold>",
    "错误类型": "标签缺失",
    "修改理由": "原文“围绕结构优化，”有加粗标签，译文漏掉了对应的<bold>标签。",
    "违反的规则": "根据原文带\"<bold></bold>\"\"<italic></italic>\"词语检查翻译后对应的译文有无标签"
  },
  {
    "错误编号": "12",
    "原文上下文": "<bold>围绕特色培育，</bold>完善实物贵金属全球化布局",
    "译文上下文": "Centering on the cultivation of characteristics, the Bank improves the global layout",
    "原文数值": "<bold>围绕特色培育，</bold>",
    "译文数值": "Centering on the cultivation of characteristics,",
    "替换锚点": "Centering on the cultivation of characteristics,",
    "译文修改建议值": "<bold>Centering on the cultivation of characteristics,</bold>",
    "错误类型": "标签缺失",
    "修改理由": "原文“围绕特色培育，”有加粗标签，译文漏掉了对应的<bold>标签。",
    "违反的规则": "根据原文带\"<bold></bold>\"\"<italic></italic>\"词语检查翻译后对应的译文有无标签"
  },
  {
    "错误编号": "13",
    "原文上下文": "<bold>围绕价值实现</bold>，强化市值管理",
    "译文上下文": "Centering on value realization, the Bank strengthens market value management",
    "原文数值": "<bold>围绕价值实现</bold>，",
    "译文数值": "Centering on value realization,",
    "替换锚点": "Centering on value realization,",
    "译文修改建议值": "<bold>Centering on value realization</bold>,",
    "错误类型": "标签缺失",
    "修改理由": "原文“围绕价值实现”有加粗标签，译文漏掉了对应的<bold>标签。",
    "违反的规则": "根据原文带\"<bold></bold>\"\"<italic></italic>\"词语检查翻译后对应的译文有无标签"
  },
  {
    "错误编号": "14",
    "原文上下文": "<bold>围绕风险管理，</bold>深化全面风险管理体系建设",
    "译文上下文": "Centering on risk management, the Bank deepens the construction",
    "原文数值": "<bold>围绕风险管理，</bold>",
    "译文数值": "Centering on risk management,",
    "替换锚点": "Centering on risk management,",
    "译文修改建议值": "<bold>Centering on risk management,</bold>",
    "错误类型": "标签缺失",
    "修改理由": "原文“围绕风险管理，”有加粗标签，译文漏掉了对应的<bold>标签。",
    "违反的规则": "根据原文带\"<bold></bold>\"\"<italic></italic>\"词语检查翻译后对应的译文有无标签"
  },
  {
    "错误编号": "15",
    "原文上下文": "<bold>建立“1+X+N”规划体系，</bold>以集团规划为“1”",
    "译文上下文": "A “1+X+N” planning system will be established, with the Group’s plan as the “1”",
    "原文数值": "<bold>建立“1+X+N”规划体系，</bold>",
    "译文数值": "A “1+X+N” planning system will be established,",
    "替换锚点": "A “1+X+N” planning system will be established,",
    "译文修改建议值": "<bold>A “1+X+N” planning system will be established,</bold>",
    "错误类型": "标签缺失",
    "修改理由": "原文“建立“1+X+N”规划体系，”有加粗标签，译文漏掉了对应的<bold>标签。",
    "违反的规则": "根据原文带\"<bold></bold>\"\"<italic></italic>\"词语检查翻译后对应的译文有无标签"
  },
  {
    "错误编号": "16",
    "原文上下文": "<bold>健全“规划—计划—预算—考核”闭环机制</bold>，坚持“发展有规划、年度有计划”",
    "译文上下文": "The Bank improves the closed-loop mechanism of “plan—budget plan—budget—assessment”, adheres to",
    "原文数值": "<bold>健全“规划—计划—预算—考核”闭环机制</bold>，",
    "译文数值": "The Bank improves the closed-loop mechanism of “plan—budget plan—budget—assessment”,",
    "替换锚点": "The Bank improves the closed-loop mechanism of “plan—budget plan—budget—assessment”,",
    "译文修改建议值": "<bold>The Bank improves the closed-loop mechanism of “plan—budget plan—budget—assessment”</bold>,",
    "错误类型": "标签缺失",
    "修改理由": "原文“健全“规划—计划—预算—考核”闭环机制”有加粗标签，译文漏掉了对应的<bold>标签。",
    "违反的规则": "根据原文带\"<bold></bold>\"\"<italic></italic>\"词语检查翻译后对应的译文有无标签"
  },
  {
    "错误编号": "17",
    "原文上下文": "承办人：现代金融研究院（党委深改办）",
    "译文上下文": "Modern Finance Research Institute (Office for Deepening Reform of the Party Committee)",
    "原文数值": "承办人：",
    "译文数值": "Modern Finance Research Institute (Office for Deepening Reform of the Party Committee)",
    "替换锚点": "Modern Finance Research Institute (Office for Deepening Reform of the Party Committee)",
    "译文修改建议值": "Undertaken by: Modern Finance Research Institute (Office for Deepening Reform of the Party Committee)",
    "错误类型": "漏译",
    "修改理由": "译文漏译了“承办人：”。",
    "违反的规则": "对照原文检查译文出现错译、漏译、多译的情况。"
  },
  {
    "错误编号": "18",
    "原文上下文": "（此件属商业秘密，严禁使用手机拍摄或通过互联网、微信微博等社交媒体传播、使用、处理和对外发布）",
    "译文上下文": "<bold>(This document is a commercial secret. Without the written permission of the Bank, the information related to this matter shall not be disseminated, used, processed, or disclosed in any form.)</bold>",
    "原文数值": "（此件属商业秘密，严禁使用手机拍摄或通过互联网、微信微博等社交媒体传播、使用、处理和对外发布）",
    "译文数值": "<bold>(This document is a commercial secret. Without the written permission of the Bank, the information related to this matter shall not be disseminated, used, processed, or disclosed in any form.)</bold>",
    "替换锚点": "<bold>(This document is a commercial secret. Without the written permission of the Bank, the information related to this matter shall not be disseminated, used, processed, or disclosed in any form.)</bold>",
    "译文修改建议值": "(This document is a commercial secret, and it is strictly prohibited to take photos with mobile phones or disseminate, use, process and release it externally through the Internet, WeChat, Weibo and other social media)",
    "错误类型": "错译/多译标签",
    "修改理由": "原文没有加粗标签，译文多加了<bold>标签；且译文内容与原文不符，错译了“严禁使用手机拍摄或通过互联网、微信微博等社交媒体传播、使用、处理和对外发布”。",
    "违反的规则": "对照原文检查译文出现错译、漏译、多译的情况；根据原文带\"<bold></bold>\"\"<italic></italic>\"词语检查翻译后对应的译文有无标签。"
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

    # 自动编号静态化
    try:
        from replace.word.numbering_to_static import convert_numbering_to_static, has_auto_numbering, convert_toc_to_static
        if has_auto_numbering(backup_path):
            print(f"\n🔢 检测到自动编号，正在转换为静态文本...")
            ok = convert_numbering_to_static(backup_path)
            if ok:
                print(f"✓ 自动编号已转为静态文本")
            else:
                print(f"⚠ 自动编号静态化失败，部分编号可能无法替换")
        print(f"\n📑 正在转换目录域为静态文本...")
        convert_toc_to_static(backup_path)
    except Exception as e:
        print(f"⚠ 编号/目录静态化异常: {e}")

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

        # 方案B：先修正重叠，再统一过滤
        fixed_text = _fix_suggestion_overlap(context, old_text, new_text)
        if fixed_text != new_text:
            print(f"  ⚙ 建议值修正: '{new_text}' → '{fixed_text}'")
            new_text = fixed_text
        if not new_text or new_text == old_text:
            results.append((error_no, "跳过", f"建议值无效或与原文一致，跳过"))
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
    DOC_PATH = r"../测试文件/Proposal on the 15th Five-Year Development Plan of ICBC.docx"
    
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
