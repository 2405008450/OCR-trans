from collections import Counter
from pathlib import Path
import json

from full_content import extract_docx_in_order, SOURCE_LABELS


def build_bilingual_json(source_docx, target_docx, output_json=None):
    """
    原文 / 译文 DOCX 对照提取
    """

    print("📖 提取原文...")
    source_segments = extract_docx_in_order(source_docx)

    print("📖 提取译文...")
    target_segments = extract_docx_in_order(target_docx)

    print("=" * 70)
    print(f"原文片段: {len(source_segments)}")
    print(f"译文片段: {len(target_segments)}")
    print("=" * 70)

    # 长度检查
    max_len = max(len(source_segments), len(target_segments))

    pairs = []

    for idx in range(max_len):

        src_seg = source_segments[idx] if idx < len(source_segments) else None
        tgt_seg = target_segments[idx] if idx < len(target_segments) else None

        source_type = src_seg.source if src_seg else (
            tgt_seg.source if tgt_seg else "unknown"
        )

        item = {
            "id": idx + 1,

            # 来源
            "source": source_type,
            "source_label": SOURCE_LABELS.get(source_type, source_type),

            # 是否来源一致
            "source_match": (
                src_seg.source == tgt_seg.source
                if src_seg and tgt_seg
                else False
            ),

            # 文本
            "source_text": src_seg.text if src_seg else "",
            "target_text": tgt_seg.text if tgt_seg else "",
        }

        pairs.append(item)

    # 输出统计
    print("📊 来源统计:")
    source_counts = Counter(p["source"] for p in pairs)

    for src, count in source_counts.items():
        label = SOURCE_LABELS.get(src, src)
        print(f"    {label}: {count} 条")

    print("=" * 70)

    # 保存
    if output_json:
        output_json = Path(output_json)

        with open(output_json, "w", encoding="utf-8") as f:
            json.dump(
                pairs,
                f,
                ensure_ascii=False,
                indent=2
            )

        print(f"✅ 已保存 JSON:")
        print(output_json)

    return pairs


if __name__ == "__main__":

    source_docx = r"C:\Users\H\Desktop\数检_程序-AI\测试文件\原文-含不可编辑_01 (2026-007)2025年年度报告.docx"
    target_docx = r"C:\Users\H\Desktop\数检_程序-AI\测试文件\译文-含不可编辑_01 (2026-007)2025年年度报告(1).docx"

    output_json = r"C:\Users\H\Desktop\数检_程序-AI\output\bilingual_pairs.json"

    pairs = build_bilingual_json(
        source_docx,
        target_docx,
        output_json
    )

    print("\n📖 前 5 条预览:\n")

    for item in pairs[:5]:

        print(f"[{item['id']:>4}] "
              f"【{item['source_label']}】")

        print("原文:")
        print(item["source_text"])

        print("译文:")
        print(item["target_text"])

        print("-" * 70)
