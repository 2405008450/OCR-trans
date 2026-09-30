import pandas as pd


def load_pairs_from_excel(excel_path):
    df = pd.read_excel(excel_path)

    pairs = []

    for _, row in df.iterrows():

        src = str(row["原文"]).strip()
        tgt = str(row["译文"]).strip()

        # 跳过空行
        if not src or not tgt:
            continue

        pairs.append((src, tgt))

    return pairs


if __name__ == "__main__":
    pairs = load_pairs_from_excel(
        r"D:\project\数检_程序-AI\测试文件\output\译文-含不可编辑_01 (2026-007)2025年年度报告(1)_alignment.xlsx"
    )

    print(pairs[:3])