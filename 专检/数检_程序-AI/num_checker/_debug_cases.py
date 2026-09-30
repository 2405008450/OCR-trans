"""临时调试脚本"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from num_checker._parser_core import extract

cases = [
    ("公告时间：二〇二六年三月", "Announcement date: March 2026"),
    ("Egypt Chuan Jin Nuo Chemical Co., Ltd./川金诺埃及化工有限责任公司，本公司控股孙公司",
     "Egypt Chuan Jin Nuo Chemical Co., Ltd., a holding second-tier subsidiary of the Company"),
    ("一种常见的无机酸，是中强酸。根据浓度不同分为纯磷酸、工业磷酸、稀磷酸等；根据制作工艺分为热法磷酸和湿法磷酸",
     "A common inorganic acid, it is a moderately strong acid. It is classified into pure phosphoric acid, industrial phosphoric acid, dilute phosphoric acid, and so on, depending on the concentration; and into thermal phosphoric acid and wet-process phosphoric acid depending on the manufacturing process."),
    ("五氧化二磷，也称磷酸酐，白色无定形粉末或六方晶体",
     "P₂O₅, also known as phosphorus anhydride, is a white amorphous powder or hexagonal crystal"),
    ("黄磷在空气中燃烧生成五氧化二磷，再经水或稀磷酸吸收制成的磷酸",
     "Yellow phosphorus burns in air to produce P₂O₅, which is then absorbed by water or dilute phosphoric acid to produce phosphoric acid."),
    ('简称：DCP，是一种在畜禽饲料中添加的用于补充畜禽钙、磷等两类矿物质营养元素的饲料添加剂，是目前我国畜禽养殖领域主要采用的一种\u201c钙+磷\u201d类添加剂',
     'Abbreviated as DCP, it is a feed additive used in livestock and poultry feed to supplement two types of mineral nutrients, calcium and phosphorus. It is currently one of the main "calcium + phosphorus" additives used in China\'s livestock and poultry farming industry.'),
    ("饲料级磷酸二氢钙", "Feed grade monocalcium phosphate"),
    ("简称：MCP，是一种高效、优良磷酸盐类饲料添加剂，主要用作补充动物体内的磷和钙两类矿物质营养元素",
     "Abbreviated as MCP, it is an efficient and high-quality phosphate feed additive, mainly used to supplement two types of mineral nutrients, phosphorus and calcium, in animals."),
    ("重过磷酸钙，又称三料过磷酸钙或三倍过磷酸钙，简称重钙",
     "Superphosphate, also known as triple superphosphate, an acidic, fast-acting phosphate fertilizer."),
    ("化学式为LiFePO4，是一种橄榄石结构的磷酸盐，用作锂离子电池的正极材料",
     "A phosphate with an olivine structure, used as a cathode material for lithium-ion batteries. The chemical formula is LiFePO4."),
    ("磷酸铁，又名磷酸高铁、正磷酸铁，分子式为FePO4，是一种白色、灰白色单斜晶体粉末",
     "Iron phosphate, also known as ferric phosphate or orthophosphoric acid iron, with the molecular formula FePO4, a white or gray-white monoclinic crystal powder."),
    ("使用硫酸等无机酸分解磷矿石制成的磷酸，生产工艺上可以分为二水法、半水法、无水法、半水-二水法和二水-半水法等",
     "Phosphoric acid produced by decomposing phosphate rock with inorganic acids such as sulfuric acid that can be manufactured through various processes, including dihydrate, hemihydrate, anhydrous, hemihydrate-dihydrate, and dihydrate-hemihydrate processes."),
    ("简称富钙，灰白色粉末，是用混酸(硫酸和磷酸)分解磷矿制成。产品有效五氧化二磷含量介于过磷酸钙与重过磷酸钙之间，一般为20% P205～30%P205。",
     "Commonly referred to as calcium-rich, a gray-white powder. The effective P₂O₅ content of the product is between that of tricalcium phosphate and triple superphosphate, generally ranging from 20% P₂O₅ to 30% P₂O₅."),
    ("52%商品磷酸是P205含量为52%的高浓度湿法肥料级商品磷酸",
     "52% commercial phosphoric acid, a high-concentration wet-process fertilizer-grade phosphoric acid with a P₂O₅ content of 52%."),
    ("云南省昆明市东川区铜都镇四方地工业园区",
     "Sifangdi Industrial Park, Tongdu Town, Dongchuan District, Kunming City, Yunnan Province"),
    ("云南省昆明市呈贡区乌龙街道办事处七彩云南第壹城1#办公楼（双子星·天枢）55层",
     "55/F, No. 1 Office Building (Gemini - Dubhe), Qicai Yunnan No. 1 City of Wulong Subdistrict Administrative Agency, Chenggong District, Kunming City, Yunnan Province"),
]

for src, tgt in cases:
    sv = [(t.value, t.tag, t.raw) for t in extract(src)]
    tv = [(t.value, t.tag, t.raw) for t in extract(tgt)]
    print(f"SRC: {src[:60]}")
    print(f"  src={sv}")
    print(f"  tgt={tv}")
    print()
