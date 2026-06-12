"""
test_single_replace.py — 用单条 JSON 数据测试夹逼定位+修订策略

用法：
  python test_single_replace.py --docx 译文.docx --out 输出.docx
"""
import argparse
from docx import Document
from revision import RevisionManager
from replace_revision import replace_and_revise_in_docx, build_para_cache

# ── 测试数据（直接从 full_report JSON 粘贴）──────────────────────────
TEST_ROW = {
    "excel_row": 48,
    "原文": "5,000,000.00元",
    "译文": "5,000,000.00",
    "AI修改详情": "'5,000,000.00' → 'RMB5,000,000.00'",
    "前一句": {
        "原文": "昆明精粹工程技术有限责任公司",
        "译文": "Kunming Jingcui Engineering Technology Co., Ltd."
    },
    "后一句": {
        "原文": "昆明市",
        "译文": "Kunming"
    },
}

# AI errors 数组（对应 align_body_errors.json / align_body_flat_errors.json 里的字段）
# 根据 AI修改详情 手工还原
TEST_ERRORS = [
    {
        "替换锚点":      "5,000,000.00",
        "译文数值":      "5,000,000.00",
        "译文修改建议值": "RMB5,000,000.00",
        "译文上下文":    "5,000,000.00",   # 译文本身就是这个值
        "修改理由":      "漏译货币单位'元'，导致译文缺失单位信息",
        "is_source_consistent": False,
    }
]


def run(docx_path: str, out_path: str):
    print(f"打开文档: {docx_path}")
    doc = Document(docx_path)
    rm  = RevisionManager(doc, author="数值检查")

    # 构建段落缓存
    cache = build_para_cache(doc, region="body")
    print(f"段落缓存: {len(cache)} 条")

    prev_tgt = TEST_ROW["前一句"]["译文"]
    next_tgt = TEST_ROW["后一句"]["译文"]

    for i, err in enumerate(TEST_ERRORS, 1):
        old_val  = err.get("替换锚点", "").strip() or err.get("译文数值", "").strip()
        new_val  = err.get("译文修改建议值", "").strip()
        context  = err.get("译文上下文", "").strip()
        anchor   = err.get("替换锚点", "").strip()
        reason   = err.get("修改理由", "")

        print(f"\n── 错误 {i}: '{old_val}' → '{new_val}'")
        print(f"   前一句译文: {prev_tgt}")
        print(f"   后一句译文: {next_tgt}")
        print(f"   context:   {context}")
        print(f"   anchor:    {anchor}")

        ok, strategy = replace_and_revise_in_docx(
            doc=doc,
            old_value=old_val,
            new_value=new_val,
            reason=reason,
            revision_manager=rm,
            context=context,
            anchor_text=anchor,
            region="body",
            doc_path=docx_path,
            para_cache=cache,
            prev_tgt=prev_tgt,
            next_tgt=next_tgt,
        )

        if ok:
            print(f"   ✅ 成功: {strategy}")
        else:
            print(f"   ⚠️  失败: {strategy}")

    doc.save(out_path)
    print(f"\n已保存: {out_path}")


if __name__ == "__main__":
    docx_path = r"D:\project\数检_程序-AI\测试\关于《中国工商银行“十五五”时期发展规划》的议案0606.docx"
    out_path = r"D:\project\数检_程序-AI\测试\Proposal on the 15th Five-Year Development Plan of ICBC.docx"
    run(docx_path, out_path)
