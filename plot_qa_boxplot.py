"""
绘制66个满足条件商品在8个季度中问答条数的箱线图。
输出：qa_quarterly_boxplot.png
"""

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager

from analyze_qa import ALL_QUARTERS, load_quarter_counts, filter_qualified, XLSX_PATH

QUARTER_LABELS = [
    "2022Q4\n(10-12月)",
    "2023Q1\n(1-3月)",
    "2023Q2\n(4-6月)",
    "2023Q3\n(7-9月)",
    "2023Q4\n(10-12月)",
    "2024Q1\n(1-3月)",
    "2024Q2\n(4-6月)",
    "2024Q3\n(7-9月)",
]


def main():
    print("读取数据...")
    counts = load_quarter_counts(XLSX_PATH)
    qualified = filter_qualified(counts)
    print(f"满足条件的商品数：{len(qualified)}")

    labels = [label for label, _, _ in ALL_QUARTERS]
    # 每个季度的数据列表，shape: (8, n_products)
    data = [
        [counts[pid].get(label, 0) for pid in qualified]
        for label in labels
    ]

    x = np.arange(1, len(labels) + 1)

    fig, ax = plt.subplots(figsize=(13, 6))

    # 处理前后底色
    ax.axvspan(0.5, 4.5, alpha=0.05, color="steelblue")
    ax.axvspan(4.5, 8.5, alpha=0.05, color="darkorange")

    # 箱线图
    bp = ax.boxplot(
        data,
        positions=x,
        widths=0.55,
        patch_artist=True,
        medianprops=dict(color="red", linewidth=2),
        whiskerprops=dict(linewidth=1.2),
        capprops=dict(linewidth=1.2),
        flierprops=dict(marker="o", markersize=3, alpha=0.5, linestyle="none"),
    )

    # 处理前蓝色、处理后橙色填充箱体
    for i, patch in enumerate(bp["boxes"]):
        patch.set_facecolor("steelblue" if i < 4 else "darkorange")
        patch.set_alpha(0.55)

    # 处理节点竖虚线
    ax.axvline(x=4.5, color="red", linestyle="--", linewidth=1.5, label="处理节点（2023-10）")

    # 区域标注
    ymax = max(max(d) for d in data)
    ax.text(2.5, ymax * 0.97, "处理前", ha="center", va="top",
            fontsize=11, color="steelblue", fontweight="bold")
    ax.text(6.5, ymax * 0.97, "处理后", ha="center", va="top",
            fontsize=11, color="darkorange", fontweight="bold")

    ax.set_xticks(x)
    ax.set_xticklabels(QUARTER_LABELS, fontsize=9)
    ax.set_xlabel("季度", fontsize=11)
    ax.set_ylabel("问答条数", fontsize=11)
    ax.set_title(
        f"各季度商品问答数量分布（箱线图）\n（{len(qualified)}个商品，2022Q4–2024Q3）",
        fontsize=13,
    )
    ax.legend(fontsize=10)
    ax.grid(axis="y", linestyle=":", alpha=0.5)

    plt.tight_layout()
    out = "qa_quarterly_boxplot.png"
    plt.savefig(out, dpi=150)
    print(f"图表已保存：{out}")


if __name__ == "__main__":
    font_manager.fontManager.addfont("/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc")
    plt.rcParams["font.sans-serif"] = ["WenQuanYi Zen Hei", "DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False
    main()
