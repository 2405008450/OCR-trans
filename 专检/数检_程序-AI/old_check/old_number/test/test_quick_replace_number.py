"""
快速替换测试

直接使用内嵌的错误数据进行测试
"""

from docx import Document
from pathlib import Path
import sys
import difflib

# 获取当前文件的父目录的父目录（即项目根目录）
project_root = str(Path(__file__).resolve().parent.parent)
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from revise.revision import RevisionManager
from replace.replace_revision import replace_and_revise_in_docx, flush_footnote_replacements
from backup_copy.backup_manager import ensure_backup_copy


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


# 测试数据 - 直接从 JSON 复制替换
TEST_ERRORS = [
  {
    "错误编号": "1",
    "原文上下文": "公司董事会及董事、高级管理人员保证年度报告内容的真实、准确、完整，不存在虚假记载、误导性陈述或者重大遗漏，并承担个别和连带的法律责任。",
    "译文上下文": "The Board of Directors (or the “Board”) and the directors, as well as senior management personnel of Shenzhen Senior Technology Material Co., Ltd. (hereinafter referred to as the “Company”) hereby guarantee that the contents of this annual report are true, accurate, and complete, without any false records, misleading statements, or material omissions, and assume joint and several liability for its contents.",
    "原文数值": "公司",
    "译文数值": "Shenzhen Senior Technology Material Co., Ltd.",
    "替换锚点": "Shenzhen Senior Technology Material Co., Ltd.",
    "译文修改建议值": "the Company",
    "错误类型": "错译/多译",
    "修改理由": "译文引入了原文中不存在的公司名称（Shenzhen Senior Technology Material Co., Ltd.）。",
    "违反的规则": "规则1：原文主权原则。"
  },
  {
    "错误编号": "2",
    "原文上下文": "公司经本次董事会审议通过的利润分配预案为：以274,867,523为基数，向全体股东每10股派发现金红利4元（含税），送红股0股（含税），以资本公积金向全体股东每10股转增0股。",
    "译文上下文": "The Company's profit distribution plan approved at the meeting of the Board of Directors is as follows: To distribute a cash dividend of RMB6 (tax-inclusive) per ten shares and zero bonus shares (tax-inclusive) to all shareholders based on the total share capital of 274,867,523 shares, with the conversion of capital reserves into the share capital of zero shares per ten shares.",
    "原文数值": "4",
    "译文数值": "6",
    "替换锚点": "RMB6",
    "译文修改建议值": "RMB4",
    "错误类型": "数值错误",
    "修改理由": "译文数值与原文不符（原文为4，译文为6）。",
    "违反的规则": "规则3：数值零误差。"
  },
  {
    "错误编号": "3",
    "原文上下文": "产品有效五氧化二磷含量介于过磷酸钙与重过磷酸钙之间，一般为20% P205～30%P205。",
    "译文上下文": "The effective P₂O₅ content of the product is between that of tricalcium phosphate and triple superphosphate, generally ranging from 20% P₂O₅ to 30% P₂O₅.",
    "原文数值": "20% P205",
    "译文数值": "20% P₂O₅",
    "替换锚点": "20% P₂O₅",
    "译文修改建议值": "20% P205",
    "错误类型": "数值错误",
    "修改理由": "译文将原文的P205修改为P₂O₅，违反了原文主权原则。",
    "违反的规则": "规则1：原文主权原则。"
  },
  {
    "错误编号": "4",
    "原文上下文": "产品有效五氧化二磷含量介于过磷酸钙与重过磷酸钙之间，一般为20% P205～30%P205。",
    "译文上下文": "The effective P₂O₅ content of the product is between that of tricalcium phosphate and triple superphosphate, generally ranging from 20% P₂O₅ to 30% P₂O₅.",
    "原文数值": "30%P205",
    "译文数值": "30% P₂O₅",
    "替换锚点": "30% P₂O₅",
    "译文修改建议值": "30% P205",
    "错误类型": "数值错误",
    "修改理由": "译文将原文的P205修改为P₂O₅，违反了原文主权原则。",
    "违反的规则": "规则1：原文主权原则。"
  },
  {
    "错误编号": "5",
    "原文上下文": "52%商品磷酸是P205含量为52%的高浓度湿法肥料级商品磷酸。",
    "译文上下文": "52% commercial phosphoric acid, a high-concentration wet-process fertilizer-grade phosphoric acid with a P₂O₅ content of 52%.",
    "原文数值": "P205",
    "译文数值": "P₂O₅",
    "替换锚点": "P₂O₅",
    "译文修改建议值": "P205",
    "错误类型": "数值错误",
    "修改理由": "译文将原文的P205修改为P₂O₅，违反了原文主权原则。",
    "违反的规则": "规则1：原文主权原则。"
  },
  {
    "错误编号": "6",
    "原文上下文": "1、磷酸的萃取及净化技术：完全自主开发的湿法磷酸净化技术已用于广西川金诺20万吨/年湿法净化磷酸装置，现已量产工业级、食品级的湿法净化磷酸，技术指标达到行业领先水平。",
    "译文上下文": "1. Phosphoric acid extraction and purification technology: The Company has independently developed a wet-process purified phosphoric acid technology, which has been successfully implemented at its Guangxi Chuan Jin Nuo with an annual capacity of 100,000 tonnes. Currently, the Company has achieved mass production of industrial-grade and food-grade purified wet-process phosphoric acid, with technical specifications reaching industry-leading levels.",
    "原文数值": "20万",
    "译文数值": "100,000",
    "替换锚点": "100,000 tonnes",
    "译文修改建议值": "200,000 tonnes",
    "错误类型": "数值错误",
    "修改理由": "译文数值与原文不符（原文为20万，译文为10万）。",
    "违反的规则": "规则3：数值零误差。"
  },
  {
    "错误编号": "7",
    "原文上下文": "昆明市东川区周边120公里范围内磷矿资源富集，开采品位18%-25%P205，以中低品位胶质磷矿为主。",
    "译文上下文": "The region within a 120 km radius of Dongchuan District, Kunming City, is abundant in phosphate resources, with ore grades ranging from 18% to 25% P₂O₅, primarily consisting of medium- to low-grade colloidal phosphate rock.",
    "原文数值": "18%-25%P205",
    "译文数值": "18% to 25% P₂O₅",
    "替换锚点": "18% to 25% P₂O₅",
    "译文修改建议值": "18% to 25% P205",
    "错误类型": "数值错误",
    "修改理由": "译文将原文的P205修改为P₂O₅，违反了原文主权原则。",
    "违反的规则": "规则1：原文主权原则。"
  },
  {
    "错误编号": "8",
    "原文上下文": "针对磷矿特点公司开发了独特的选矿技术，浮选出26%-33%P205精矿，满足公司磷的分级利用对矿的需求。",
    "译文上下文": "To optimize the utilization of these resources, the Company has developed proprietary beneficiation technology, enabling the flotation of phosphate concentrate with a P₂O₅ grade of 26% to 33%, thereby meeting the Company’s requirements for graded utilization of phosphate ore.",
    "原文数值": "26%-33%P205",
    "译文数值": "P₂O₅ grade of 26% to 33%",
    "替换锚点": "P₂O₅",
    "译文修改建议值": "P205",
    "错误类型": "数值错误",
    "修改理由": "译文将原文的P205修改为P₂O₅，违反了原文主权原则。",
    "违反的规则": "规则1：原文主权原则。"
  },
  {
    "错误编号": "9",
    "原文上下文": "33%P205磷精矿用于重钙二次矿，28%P205磷精矿用于半水酸生产出优质磷酸，23%P205磷精矿用于二水酸生产氢钙。",
    "译文上下文": "Phosphate concentrate with a P₂O₅ grade of 33% is used for triple superphosphate (TSP) production, 28% P₂O₅ phosphate concentrate is utilized in semi-hydrate acid production to produce high-quality phosphoric acid, and 23% P₂O₅ phosphate concentrate is applied in dihydrate acid production for the manufacturing of calcium hydrogen phosphate.",
    "原文数值": "33%P205",
    "译文数值": "P₂O₅ grade of 33%",
    "替换锚点": "P₂O₅ grade of 33%",
    "译文修改建议值": "P205 grade of 33%",
    "错误类型": "数值错误",
    "修改理由": "译文将原文的P205修改为P₂O₅，违反了原文主权原则。",
    "违反的规则": "规则1：原文主权原则。"
  },
  {
    "错误编号": "10",
    "原文上下文": "33%P205磷精矿用于重钙二次矿，28%P205磷精矿用于半水酸生产出优质磷酸，23%P205磷精矿用于二水酸生产氢钙。",
    "译文上下文": "Phosphate concentrate with a P₂O₅ grade of 33% is used for triple superphosphate (TSP) production, 28% P₂O₅ phosphate concentrate is utilized in semi-hydrate acid production to produce high-quality phosphoric acid, and 23% P₂O₅ phosphate concentrate is applied in dihydrate acid production for the manufacturing of calcium hydrogen phosphate.",
    "原文数值": "28%P205",
    "译文数值": "28% P₂O₅",
    "替换锚点": "28% P₂O₅",
    "译文修改建议值": "28% P205",
    "错误类型": "数值错误",
    "修改理由": "译文将原文的P205修改为P₂O₅，违反了原文主权原则。",
    "违反的规则": "规则1：原文主权原则。"
  },
  {
    "错误编号": "11",
    "原文上下文": "33%P205磷精矿用于重钙二次矿，28%P205磷精矿用于半水酸生产出优质磷酸，23%P205磷精矿用于二水酸生产氢钙。",
    "译文上下文": "Phosphate concentrate with a P₂O₅ grade of 33% is used for triple superphosphate (TSP) production, 28% P₂O₅ phosphate concentrate is utilized in semi-hydrate acid production to produce high-quality phosphoric acid, and 23% P₂O₅ phosphate concentrate is applied in dihydrate acid production for the manufacturing of calcium hydrogen phosphate.",
    "原文数值": "23%P205",
    "译文数值": "23% P₂O₅",
    "替换锚点": "23% P₂O₅",
    "译文修改建议值": "23% P205",
    "错误类型": "数值错误",
    "修改理由": "译文将原文的P205修改为P₂O₅，违反了原文主权原则。",
    "违反的规则": "规则1：原文主权原则。"
  },
  {
    "错误编号": "12",
    "原文上下文": "（五）关于相关利益者",
    "译文上下文": "vi. Stakeholders",
    "原文数值": "（五）",
    "译文数值": "vi.",
    "替换锚点": "vi.",
    "译文修改建议值": "v.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级错误，原文为（五），译文误写为vi.。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "13",
    "原文上下文": "1、刘甍先生：男，1970年2月生，中国国籍，无境外永久居留权，本科学历，云南省优秀民营企业家，昆明市第十届优秀企业家。",
    "译文上下文": "He is an outstanding private entrepreneur in Yunnan Province and an excellent entrepreneur in Kunming.",
    "原文数值": "第十届",
    "译文数值": "an excellent entrepreneur in Kunming",
    "替换锚点": "an excellent entrepreneur in Kunming",
    "译文修改建议值": "an excellent entrepreneur in the 10th Kunming session",
    "错误类型": "漏译",
    "修改理由": "译文漏译了原文中的数值“第十届”。",
    "违反的规则": "规则1：原文主权原则"
  },
  {
    "错误编号": "14",
    "原文上下文": "2011年8月至2017年9月15日任公司副总经理；",
    "译文上下文": "from August 2011 to September 2017, he was the Company’s Deputy General Manager;",
    "原文数值": "2017年9月15日",
    "译文数值": "September 2017",
    "替换锚点": "September 2017",
    "译文修改建议值": "September 15, 2017",
    "错误类型": "漏译",
    "修改理由": "译文漏译了原文日期中的具体日子“15日”。",
    "违反的规则": "规则1：原文主权原则"
  },
  {
    "错误编号": "15",
    "原文上下文": "2017年9月至2023年9月14日任公司监事会主席，2023年9月14日至今任公司董事。",
    "译文上下文": "from September 2017 to September 2023, he served as the Chairman of the Board of Supervisors, and since September 2023, he has been a Director of the Company.",
    "原文数值": "2023年9月14日",
    "译文数值": "September 2023",
    "替换锚点": "to September 2023",
    "译文修改建议值": "to September 14, 2023",
    "错误类型": "漏译",
    "修改理由": "译文漏译了原文日期中的具体日子“14日”。",
    "违反的规则": "规则1：原文主权原则"
  },
  {
    "错误编号": "16",
    "原文上下文": "（4）2021年，公司对广西川金诺化工进行了增资，广西川金诺化工的注册资本由11,000万元增加为55,396万元，“防城港凌沄”与“昆明凌嵘”同比例对广西川金诺化工进行同比例增资。",
    "译文上下文": "(4) In 2021, the Company increased the capital of Guangxi Chuanjinno Chemical Co., Ltd., raising its registered capital from RMB11 million to RMB55.396 million.",
    "原文数值": "11,000万元",
    "译文数值": "RMB11 million",
    "替换锚点": "RMB11 million",
    "译文修改建议值": "RMB110 million",
    "错误类型": "数值错误",
    "修改理由": "数量级转换错误，原文“11,000万元”应译为“110 million”。",
    "违反的规则": "规则3：数值零误差"
  },
  {
    "错误编号": "17",
    "原文上下文": "（4）2021年，公司对广西川金诺化工进行了增资，广西川金诺化工的注册资本由11,000万元增加为55,396万元，“防城港凌沄”与“昆明凌嵘”同比例对广西川金诺化工进行同比例增资。",
    "译文上下文": "(4) In 2021, the Company increased the capital of Guangxi Chuanjinno Chemical Co., Ltd., raising its registered capital from RMB11 million to RMB55.396 million.",
    "原文数值": "55,396万元",
    "译文数值": "RMB55.396 million",
    "替换锚点": "RMB55.396 million",
    "译文修改建议值": "RMB553.96 million",
    "错误类型": "数值错误",
    "修改理由": "数量级转换错误，原文“55,396万元”应译为“553.96 million”。",
    "违反的规则": "规则3：数值零误差"
  },
  {
    "错误编号": "18",
    "原文上下文": "广西川金诺化工有限公司\t2024年04月24日\t15,000\t2024年04月23日\t0\t连带责任保证\t\t\t\t否\t否",
    "译文上下文": "Guangxi Chuan Jin Nuo Chemical Co., Ltd.\tApril 24, 2024\t15,000\tJanuary 23, 2024\t0\tJoint and several liability guarantee\t\t\t\tNo\tNo",
    "原文数值": "2024年04月23日",
    "译文数值": "January 23, 2024",
    "替换锚点": "January 23, 2024",
    "译文修改建议值": "April 23, 2024",
    "错误类型": "日期错误",
    "修改理由": "译文月份翻译错误，原文为04月（April），译文误译为January。",
    "违反的规则": "数值零误差"
  },
  {
    "错误编号": "19",
    "原文上下文": "5万吨/年电池级磷酸铁锂正极材料前驱体材料磷酸铁及配套60万吨/年硫磺制酸项目：公司第四届董事会第三十二次会议、第四届监事会第二十一次会议审议通过了《关于使用募集资金置换预先投入募投项目及已支付发行费用自筹资金的议案》，公司独立董事发表了明确同意的独立意见。",
    "译文上下文": "50,000 tonnes/year Battery-grade Lithium Iron Phosphate Precursor Material (Iron Phosphate) and Supporting 60,000 tonnes/year Sulfuric Acid Production Project: At the 32nd Meeting of the 4th Board of Directors and the 21st Meeting of the 4th Board of Supervisors, the Proposal on the Use of Raised Funds to Replace Pre-invested Raised Projects and Self-raised Funds That Have Paid Issuance Fees was reviewed and approved.",
    "原文数值": "60万吨",
    "译文数值": "60,000 tonnes",
    "替换锚点": "60,000 tonnes",
    "译文修改建议值": "600,000 tonnes",
    "错误类型": "数值错误",
    "修改理由": "数量级错误，原文为60万（600,000），译文误译为60,000。",
    "违反的规则": "数值零误差"
  },
  {
    "错误编号": "20",
    "原文上下文": "1.重新计量设定受益计划变动额",
    "译文上下文": "6.1.1 Changes caused by re-measurements on defined benefit schemes",
    "原文数值": "1.",
    "译文数值": "6.1.1",
    "替换锚点": "6.1.1",
    "译文修改建议值": "1.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为1.，译文误写为6.1.1。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "21",
    "原文上下文": "3.其他权益工具投资公允价值变动",
    "译文上下文": "6.1.3 Changes in the fair value of investments in other equity instruments",
    "原文数值": "3.",
    "译文数值": "6.1.3",
    "替换锚点": "6.1.3",
    "译文修改建议值": "3.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为3.，译文误写为6.1.3。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "22",
    "原文上下文": "4.企业自身信用风险公允价值变动",
    "译文上下文": "6.1.4 Changes in the fair value arising from changes in own credit risk",
    "原文数值": "4.",
    "译文数值": "6.1.4",
    "替换锚点": "6.1.4",
    "译文修改建议值": "4.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为4.，译文误写为6.1.4。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "23",
    "原文上下文": "5.其他",
    "译文上下文": "6.1.5 Other",
    "原文数值": "5.",
    "译文数值": "6.1.5",
    "替换锚点": "6.1.5",
    "译文修改建议值": "5.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为5.，译文误写为6.1.5。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "24",
    "原文上下文": "2.其他债权投资公允价值变动",
    "译文上下文": "6.2.2 Changes in the fair value of investments in other debt obligations",
    "原文数值": "2.",
    "译文数值": "6.2.2",
    "替换锚点": "6.2.2",
    "译文修改建议值": "2.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为2.，译文误写为6.2.2。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "25",
    "原文上下文": "3.金融资产重分类计入其他综合收益的金额",
    "译文上下文": "6.2.3 Other comprehensive income arising from the reclassification of financial assets",
    "原文数值": "3.",
    "译文数值": "6.2.3",
    "替换锚点": "6.2.3",
    "译文修改建议值": "3.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为3.，译文误写为6.2.3。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "26",
    "原文上下文": "4.其他债权投资信用减值准备",
    "译文上下文": "6.2.4 Credit impairment allowance for investments in other debt obligations",
    "原文数值": "4.",
    "译文数值": "6.2.4",
    "替换锚点": "6.2.4",
    "译文修改建议值": "4.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为4.，译文误写为6.2.4。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "27",
    "原文上下文": "5.现金流量套期储备\t5,181,243.04\t-1,195,450.62",
    "译文上下文": "6.2.5 Reserve for cash flow hedges\t5,181,243.04\t-1,195,450.62",
    "原文数值": "5.",
    "译文数值": "6.2.5",
    "替换锚点": "6.2.5",
    "译文修改建议值": "5.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为5.，译文误写为6.2.5。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "28",
    "原文上下文": "6.外币财务报表折算差额\t-509,741.28\t",
    "译文上下文": "6.2.6 Differences arising from the translation of foreign currency-denominated financial statements\t-509,741.28\t",
    "原文数值": "6.",
    "译文数值": "6.2.6",
    "替换锚点": "6.2.6",
    "译文修改建议值": "6.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为6.，译文误写为6.2.6。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "29",
    "原文上下文": "（一）基本每股收益\t1.6510\t0.6405",
    "译文上下文": "8.1 Basic earnings per share\t1.6510\t0.6405",
    "原文数值": "（一）",
    "译文数值": "8.1",
    "替换锚点": "8.1",
    "译文修改建议值": "i.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为（一），译文误写为8.1。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "30",
    "原文上下文": "（二）稀释每股收益\t1.6510\t0.6405",
    "译文上下文": "8.2 Diluted earnings per share\t1.6510\t0.6405",
    "原文数值": "（二）",
    "译文数值": "8.2",
    "替换锚点": "8.2",
    "译文修改建议值": "ii.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为（二），译文误写为8.2。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "31",
    "原文上下文": "二、营业利润（亏损以“－”号填列）\t295,184,280.23\t87,887,434.02",
    "译文上下文": "2. Operating Profit (“-” for loss)\t295,184,280.23\t87,887,434.02",
    "原文数值": "二、",
    "译文数值": "2.",
    "替换锚点": "2.",
    "译文修改建议值": "II.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为二、，译文误写为2.。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "32",
    "原文上下文": "（一）持续经营净利润（净亏损以“－”号填列）\t271,001,842.37\t73,723,490.14",
    "译文上下文": "4.1 Net profit from continuing operations (“-” for net loss)\t271,001,842.37\t73,723,490.14",
    "原文数值": "（一）",
    "译文数值": "4.1",
    "替换锚点": "4.1",
    "译文修改建议值": "i.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为（一），译文误写为4.1。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "33",
    "原文上下文": "二、投资活动产生的现金流量：\t\t",
    "译文上下文": "2. Cash Flows from Investing Activities:\t\t",
    "原文数值": "二、",
    "译文数值": "2.",
    "替换锚点": "2.",
    "译文修改建议值": "II.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为二、，译文误写为2.。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "34",
    "原文上下文": "（一）综合收益总额\t\t\t\t\t\t\t4,671,501.76\t\t\t\t453,813,148.37\t\t458,484,650.13\t22,220,090.03\t480,704,740.16",
    "译文上下文": "3.1 Total comprehensive income\t\t\t\t\t\t\t4,671,501.76\t\t\t\t453,813,148.37\t\t458,484,650.13\t22,220,090.03\t480,704,740.16",
    "原文数值": "（一）",
    "译文数值": "3.1",
    "替换锚点": "3.1",
    "译文修改建议值": "i.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为（一），译文误写为3.1。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "35",
    "原文上下文": "（二）所有者投入和减少资本\t\t\t\t\t\t\t\t\t\t\t\t\t\t15,828,300.00\t15,828,300.00",
    "译文上下文": "3.2 Capital increased and reduced by owners\t\t\t\t\t\t\t\t\t\t\t\t\t\t15,828,300.00\t15,828,300.00",
    "原文数值": "（二）",
    "译文数值": "3.2",
    "替换锚点": "3.2",
    "译文修改建议值": "ii.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为（二），译文误写为3.2。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "36",
    "原文上下文": "1．所有者投入的普通股\t\t\t\t\t\t\t\t\t\t\t\t\t\t15,828,300.00\t15,828,300.00",
    "译文上下文": "3.2.1 Ordinary shares increased by shareholders\t\t\t\t\t\t\t\t\t\t\t\t\t\t15,828,300.00\t15,828,300.00",
    "原文数值": "1．",
    "译文数值": "3.2.1",
    "替换锚点": "3.2.1",
    "译文修改建议值": "1.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为1．，译文误写为3.2.1。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "37",
    "原文上下文": "2．其他权益工具持有者投入资本\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t",
    "译文上下文": "3.2.2 Capital increased by holders of other equity instruments\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t",
    "原文数值": "2．",
    "译文数值": "3.2.2",
    "替换锚点": "3.2.2",
    "译文修改建议值": "2.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为2．，译文误写为3.2.2。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "38",
    "原文上下文": "1．提取盈余公积\t\t\t\t\t\t\t\t\t27,100,184.24\t\t-27,100,184.24\t\t\t\t",
    "译文上下文": "3.3.1 Appropriation to surplus reserves\t\t\t\t\t\t\t\t\t27,100,184.24\t\t-27,100,184.24\t\t\t\t",
    "原文数值": "1．",
    "译文数值": "3.3.1",
    "替换锚点": "3.3.1",
    "译文修改建议值": "1.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为1．，译文误写为3.3.1。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "39",
    "原文上下文": "2．提取一般风险准备\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t",
    "译文上下文": "3.3.2 Appropriation to general reserve\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t",
    "原文数值": "2．",
    "译文数值": "3.3.2",
    "替换锚点": "3.3.2",
    "译文修改建议值": "2.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为2．，译文误写为3.3.2。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "40",
    "原文上下文": "3．对所有者（或股东）的分配\t\t\t\t\t\t\t\t\t\t\t-82,460,256.90\t\t-82,460,256.90\t-8,904,000.00\t-91,364,256.90",
    "译文上下文": "3.3.3 Appropriation to owners (or shareholders)\t\t\t\t\t\t\t\t\t\t\t-82,460,256.90\t\t-82,460,256.90\t-8,904,000.00\t-91,364,256.90",
    "原文数值": "3．",
    "译文数值": "3.3.3",
    "替换锚点": "3.3.3",
    "译文修改建议值": "3.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为3．，译文误写为3.3.3。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "41",
    "原文上下文": "4．其他\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t",
    "译文上下文": "3.2.4 Other\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t",
    "原文数值": "4．",
    "译文数值": "3.2.4",
    "替换锚点": "3.2.4",
    "译文修改建议值": "4.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为4．，译文误写为3.2.4。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "42",
    "原文上下文": "（四）所有者权益内部结转\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t",
    "译文上下文": "3.4 Transfers within owners’ equity\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t",
    "原文数值": "（四）",
    "译文数值": "3.4",
    "替换锚点": "3.4",
    "译文修改建议值": "iv.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为（四），译文误写为3.4。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "43",
    "原文上下文": "1．资本公积转增资本（或股本）\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t",
    "译文上下文": "3.4.1 Increase in capital (or share capital) from capital reserves\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t",
    "原文数值": "1．",
    "译文数值": "3.4.1",
    "替换锚点": "3.4.1",
    "译文修改建议值": "1.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为1．，译文误写为3.4.1。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "44",
    "原文上下文": "2．盈余公积转增资本（或股本）\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t",
    "译文上下文": "3.4.2 Increase in capital (or share capital) from surplus reserves\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t",
    "原文数值": "2．",
    "译文数值": "3.4.2",
    "替换锚点": "3.4.2",
    "译文修改建议值": "2.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为2．，译文误写为3.4.2。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "45",
    "原文上下文": "3．盈余公积弥补亏损\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t",
    "译文上下文": "3.4.3 Loss offset by surplus reserves\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t",
    "原文数值": "3．",
    "译文数值": "3.4.3",
    "替换锚点": "3.4.3",
    "译文修改建议值": "3.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为3．，译文误写为3.4.3。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "46",
    "原文上下文": "4．设定受益计划变动额结转留存收益\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t",
    "译文上下文": "3.4.4 Changes in defined benefit pension schemes transferred to retained earnings\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t",
    "原文数值": "4．",
    "译文数值": "3.4.4",
    "替换锚点": "3.4.4",
    "译文修改建议值": "4.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为4．，译文误写为3.4.4。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "47",
    "原文上下文": "5．其他综合收益结转留存收益\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t",
    "译文上下文": "3.4.5 Other comprehensive income transferred to retained earnings\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t",
    "原文数值": "5．",
    "译文数值": "3.4.5",
    "替换锚点": "3.4.5",
    "译文修改建议值": "5.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为5．，译文误写为3.4.5。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "48",
    "原文上下文": "6．其他\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t",
    "译文上下文": "3.4.6 Other\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t",
    "原文数值": "6．",
    "译文数值": "3.4.6",
    "替换锚点": "3.4.6",
    "译文修改建议值": "6.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为6．，译文误写为3.4.6。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "49",
    "原文上下文": "（五）专项储备\t\t\t\t\t\t\t\t1,479,017.48\t\t\t\t\t1,479,017.48\t7,975.77\t1,486,993.25",
    "译文上下文": "3.5 Specific reserve\t\t\t\t\t\t\t\t1,479,017.48\t\t\t\t\t1,479,017.48\t7,975.77\t1,486,993.25",
    "原文数值": "（五）",
    "译文数值": "3.5",
    "替换锚点": "3.5",
    "译文修改建议值": "v.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为（五），译文误写为3.5。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "50",
    "原文上下文": "2）以公允价值计量且其变动计入其他综合收益的债务工具投资",
    "译文上下文": "(2) Debt instrument investments measured at fair value through other comprehensive income (FVOCI)",
    "原文数值": "2）",
    "译文数值": "(2)",
    "替换锚点": "(2)",
    "译文修改建议值": "2)",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为2），译文误写为(2)。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "51",
    "原文上下文": "2）按照信用风险特征组合计提减值准备的组合类别及确定依据",
    "译文上下文": "(2) Categories of portfolio for which impairment provisions are made based on credit risk characteristics and the criteria for determination.",
    "原文数值": "2）",
    "译文数值": "(2)",
    "替换锚点": "(2)",
    "译文修改建议值": "2)",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为2），译文误写为(2)。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "52",
    "原文上下文": "① 应收账款（与合同资产）的组合类别及确定依据",
    "译文上下文": "a) Accounts Receivable (and Contract Assets) Portfolio Categories and Determination Basis",
    "原文数值": "①",
    "译文数值": "a)",
    "替换锚点": "a)",
    "译文修改建议值": "i.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为①，译文误写为a)。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "53",
    "原文上下文": "② 应收票据的组合类别及确定依据",
    "译文上下文": "b) Notes Receivables Portfolio Categories and Determination Basis",
    "原文数值": "②",
    "译文数值": "b)",
    "替换锚点": "b)",
    "译文修改建议值": "ii.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为②，译文误写为b)。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "54",
    "原文上下文": "③ 其他应收款的组合类别及确定依据",
    "译文上下文": "c) Other receivables portfolio categories and determination basis",
    "原文数值": "③",
    "译文数值": "c)",
    "替换锚点": "c)",
    "译文修改建议值": "iii.",
    "错误类型": "编号层级",
    "修改理由": "译文编号层级与原文不一致，原文为③，译文误写为c)。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "55",
    "原文上下文": "（1）一般原则",
    "译文上下文": "General principles",
    "原文数值": "（1）",
    "译文数值": "General principles",
    "替换锚点": "General principles",
    "译文修改建议值": "(1) General principles",
    "错误类型": "编号漏译",
    "修改理由": "译文漏译了原文的编号“（1）”。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "56",
    "原文上下文": "（2）具体方法",
    "译文上下文": "Specific method",
    "原文数值": "（2）",
    "译文数值": "Specific method",
    "替换锚点": "Specific method",
    "译文修改建议值": "(2) Specific method",
    "错误类型": "编号漏译",
    "修改理由": "译文漏译了原文的编号“（2）”。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "57",
    "原文上下文": "1）对于某一时点转让商品控制权的境内销售",
    "译文上下文": "Domestic Sales - Transfer of Control of Goods at a Point in Time",
    "原文数值": "1）",
    "译文数值": "Domestic Sales - Transfer of Control of Goods at a Point in Time",
    "替换锚点": "Domestic Sales - Transfer of Control of Goods at a Point in Time",
    "译文修改建议值": "1) Domestic Sales - Transfer of Control of Goods at a Point in Time",
    "错误类型": "编号漏译",
    "修改理由": "译文漏译了原文的编号“1）”。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "58",
    "原文上下文": "2）对于某一时点转让商品控制权的境外销售",
    "译文上下文": "Overseas Sales - Transfer of Control of Goods at a Point in Time",
    "原文数值": "2）",
    "译文数值": "Overseas Sales - Transfer of Control of Goods at a Point in Time",
    "替换锚点": "Overseas Sales - Transfer of Control of Goods at a Point in Time",
    "译文修改建议值": "2) Overseas Sales - Transfer of Control of Goods at a Point in Time",
    "错误类型": "编号漏译",
    "修改理由": "译文漏译了原文的编号“2）”。",
    "违反的规则": "规则4：编号连续性"
  },
  {
    "错误编号": "59",
    "原文上下文": "25、递延所得税资产/递延所得税负债",
    "译文上下文": "Deferred tax assets and deferred tax liabilities",
    "原文数值": "25、",
    "译文数值": "Deferred tax assets and deferred tax liabilities",
    "替换锚点": "Deferred tax assets and deferred tax liabilities",
    "译文修改建议值": "25. Deferred tax assets and deferred tax liabilities",
    "错误类型": "编号漏译",
    "修改理由": "译文漏译了原文的编号“25、”。",
    "违反的规则": "规则4：编号连续性"
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

    # 自动编号静态化 + 目录静态化
    try:
        from replace.numbering_to_static import convert_numbering_to_static, has_auto_numbering, convert_toc_to_static
        if has_auto_numbering(backup_path):
            print(f"\n🔢 检测到自动编号，正在转换为静态文本...")
            if convert_numbering_to_static(backup_path):
                print(f"✓ 自动编号已转为静态文本")
            else:
                print(f"⚠ 自动编号静态化失败，部分编号可能无法替换")
        print(f"\n📑 正在转换目录域为静态文本...")
        convert_toc_to_static(backup_path)
    except Exception as e:
        print(f"⚠ 编号/目录静态化异常: {e}")

    # 打开文档
    doc = Document(backup_path)
    doc._numbering_staticized = True
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

        # 重叠检测：修正建议值
        fixed_text = _fix_suggestion_overlap(context, old_text, new_text)
        if fixed_text != new_text:
            print(f"  ⚙ 建议值修正: '{new_text}' → '{fixed_text}'")
            new_text = fixed_text
        if not new_text or new_text == old_text:
            results.append((error_no, "跳过", f"建议值无效或与原文一致，跳过"))
            continue

        print(f"\n[测试 {idx}/{len(test_cases)}] 错误编号: {error_no}")
        print(f"  查找: '{old_text[:60]}'")
        print(f"  替换: '{new_text[:60]}'")

        try:
            ok, strategy = replace_and_revise_in_docx(
                doc, old_text, new_text, reason, revision_manager,
                context=context, anchor_text=anchor, region="body",
                doc_path=str(backup_path)
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

    # 执行脚注替换（必须在 doc.save() 之后）
    footnote_count = flush_footnote_replacements(doc, str(backup_path))
    if footnote_count > 0:
        print(f"✓ 脚注替换完成: {footnote_count} 处")

    # 统计
    print("\n" + "=" * 80)
    print("测试结果")
    print("=" * 80)

    success = sum(1 for _, status, _ in results if status == "成功")
    fail = sum(1 for _, status, _ in results if status == "失败")
    skip = sum(1 for _, status, _ in results if status == "跳过")
    error_count = sum(1 for _, status, _ in results if status == "异常")

    print(f"\n✓ 成功: {success}")
    print(f"✗ 失败: {fail}")
    print(f"⊘ 跳过: {skip}")
    print(f"⚠ 异常: {error_count}")
    print(f"━ 总计: {len(results)}")

    if success + fail > 0:
        print(f"\n成功率: {success / (success + fail):.1%}")

    print("\n详细结果:")
    for error_no, status, detail in results:
        symbol = {"成功": "✓", "失败": "✗", "跳过": "⊘", "异常": "⚠"}.get(status, "?")
        print(f"  {symbol} 错误 {error_no}: {status} - {detail[:60]}")

    print(f"\n✅ 测试完成！")
    print(f"📄 结果文档: {backup_path}")

    return results


if __name__ == "__main__":
    DOC_PATH = r"C:\Users\H\Desktop\测试项目\中翻译\测试文件\译文-含不可编辑_01 (2026-007)2025年年度报告(1).docx"

    if len(sys.argv) > 1:
        DOC_PATH = sys.argv[1]

    print(f"使用文档: {DOC_PATH}")

    if not Path(DOC_PATH).exists():
        print(f"\n❌ 文档不存在: {DOC_PATH}")
        print("\n💡 提示:")
        print("  - 需要使用译文文档")
        print("  - 如果文档在其他位置，请提供完整路径:")
        print(f"    python {Path(__file__).name} \"你的译文文档.docx\"")
        sys.exit(1)

    print()
    quick_test(DOC_PATH)
