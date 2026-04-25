"""
绘制66个满足条件商品在8个季度中的问答条数折线图。
输出：qa_quarterly_trend.png
"""

import openpyxl
from collections import defaultdict
from datetime import datetime

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np

from analyze_qa import (
    XLSX_PATH,
    ALL_QUARTERS,
    load_quarter_counts,
    filter_qualified,
)

# X轴季度标签
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

    # 构建矩阵 (n_products × 8)
    labels = [label for label, _, _ in ALL_QUARTERS]
    matrix = np.array(
        [[counts[pid].get(label, 0) for label in labels] for pid in qualified],
        dtype=float,
    )

    x = np.arange(len(labels))

    fig, ax = plt.subplots(figsize=(12, 6))

    # 各商品折线（灰色、半透明）
    for row in matrix:
        ax.plot(x, row, color="gray", alpha=0.25, linewidth=0.8)

    # 均值折线
    mean_vals = matrix.mean(axis=0)
    ax.plot(x, mean_vals, color="#1f77b4", linewidth=2.5,
            marker="o", markersize=6, label=f"均值（n={len(qualified)}）")

    # 处理节点竖虚线（post_Q1 左边，即 x=3.5）
    ax.axvline(x=3.5, color="red", linestyle="--", linewidth=1.5, label="处理节点（2023-10）")

    # 前后区域底色
    ax.axvspan(-0.5, 3.5, alpha=0.04, color="blue")
    ax.axvspan(3.5, 7.5, alpha=0.04, color="orange")

    # 区域文字标注
    ax.text(1.5, ax.get_ylim()[1] if ax.get_ylim()[1] > 0 else 1,
            "处理前", ha="center", va="top", fontsize=10, color="steelblue", alpha=0.7)
    ax.text(5.5, ax.get_ylim()[1] if ax.get_ylim()[1] > 0 else 1,
            "处理后", ha="center", va="top", fontsize=10, color="darkorange", alpha=0.7)

    ax.set_xticks(x)
    ax.set_xticklabels(QUARTER_LABELS, fontsize=9)
    ax.set_xlabel("季度", fontsize=11)
    ax.set_ylabel("问答条数", fontsize=11)
    ax.set_title("各季度商品问答数量趋势\n（66个商品，2022Q4–2024Q3）", fontsize=13)
    ax.yaxis.set_major_locator(ticker.MaxNLocator(integer=True))
    ax.legend(fontsize=10)
    ax.grid(axis="y", linestyle=":", alpha=0.5)

    # 重新调整区域标注 y 位置
    ymax = matrix.max()
    ax.texts[0].set_position((1.5, ymax * 0.97))
    ax.texts[1].set_position((5.5, ymax * 0.97))

    plt.tight_layout()
    out = "qa_quarterly_trend.png"
    plt.savefig(out, dpi=150)
    print(f"图表已保存：{out}")


if __name__ == "__main__":
    # 配置中文字体（使用系统内 WenQuanYi Zen Hei）
    from matplotlib import font_manager
    font_path = "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"
    font_manager.fontManager.addfont(font_path)
    plt.rcParams["font.sans-serif"] = ["WenQuanYi Zen Hei", "DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False
    main()
