"""
Mashgin Strategic Bundle Visualizations
=======================================

Run this script AFTER:
    mashgin_strategic_bundle_fast_csv.py

It reads the CSV files in:
    mashgin_strategic_outputs

and both displays the charts in Spyder and saves executive-ready PNG files in:
    mashgin_strategic_outputs/visualizations

Required packages:
    pip install pandas numpy matplotlib

Spyder:
    1. Put this file in the same folder as the Apriori/rule script.
    2. Run the Apriori/rule script first.
    3. Open this file in Spyder and press F5.
"""

from __future__ import annotations

from pathlib import Path
import textwrap

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import PercentFormatter


# ---------------------------------------------------------------------
# SETTINGS
# ---------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = SCRIPT_DIR / "mashgin_strategic_outputs"
VIS_DIR = OUTPUT_DIR / "visualizations"

# True = keep figures open so they appear in Spyder's Plots pane.
# False = save PNG files only.
SHOW_CHARTS_IN_SPYDER = True
SAVE_PNG_FILES = True

TOP_CATEGORY_RULES = 12
TOP_THIRD_ITEM_RULES = 10
TOP_SKU_RULES = 12
TOP_SCATTER_LABELS = 8
DPI = 220


# ---------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------
def read_csv_optional(filename: str, required_columns: list[str]) -> pd.DataFrame:
    """Read a CSV if it exists and contains the expected columns."""
    path = OUTPUT_DIR / filename

    if not path.exists():
        print(f"Skipping {filename}: file not found.")
        return pd.DataFrame()

    df = pd.read_csv(path)

    missing = [column for column in required_columns if column not in df.columns]
    if missing:
        print(f"Skipping {filename}: missing columns {missing}.")
        return pd.DataFrame()

    return df


def wrap_label(value: object, width: int = 46) -> str:
    """Wrap long rule labels without breaking words."""
    return "\n".join(
        textwrap.wrap(
            str(value),
            width=width,
            break_long_words=False,
            break_on_hyphens=False,
        )
    )


def save_figure(fig: plt.Figure, filename: str) -> None:
    """Save the figure and optionally keep it open in Spyder's Plots pane."""
    VIS_DIR.mkdir(parents=True, exist_ok=True)
    path = VIS_DIR / filename

    if SAVE_PNG_FILES:
        fig.savefig(path, dpi=DPI, bbox_inches="tight", facecolor="white")
        print(f"Saved: {path.resolve()}")

    if SHOW_CHARTS_IN_SPYDER:
        # Do not close the figure. Spyder will collect it in the Plots pane.
        fig.canvas.draw_idle()
    else:
        plt.close(fig)


def safe_numeric(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    """Convert selected columns to numeric and remove unusable rows."""
    out = df.copy()

    for column in columns:
        out[column] = pd.to_numeric(out[column], errors="coerce")

    return out.dropna(subset=columns)


def annotate_bar_metrics(
    ax: plt.Axes,
    df: pd.DataFrame,
    value_column: str,
    confidence_column: str = "confidence",
    lift_column: str = "lift",
    support_column: str = "support_count",
) -> None:
    """Add confidence, lift, and support beside horizontal bars."""
    if df.empty:
        return

    max_value = float(df[value_column].max())
    offset = max(max_value * 0.018, 0.25)

    for y_position, (_, row) in enumerate(df.iterrows()):
        annotation = (
            f"Conf. {row[confidence_column]:.1%} | "
            f"Lift {row[lift_column]:.2f} | "
            f"n={int(row[support_column]):,}"
        )
        ax.text(
            row[value_column] + offset,
            y_position,
            annotation,
            va="center",
            ha="left",
            fontsize=9.5,
        )

    ax.set_xlim(0, max_value * 1.47 if max_value > 0 else 1)


# ---------------------------------------------------------------------
# CHART 1: BASKET SIZE DISTRIBUTION
# ---------------------------------------------------------------------
def plot_basket_size_distribution() -> None:
    df = read_csv_optional(
        "basket_size_distribution.csv",
        [
            "unique_items_in_transaction",
            "transaction_count",
            "transaction_percent",
        ],
    )
    if df.empty:
        return

    df = safe_numeric(
        df,
        [
            "unique_items_in_transaction",
            "transaction_count",
            "transaction_percent",
        ],
    )
    if df.empty:
        return

    # Aggregate all baskets with five or more unique items into one executive category.
    df["basket_group"] = np.where(
        df["unique_items_in_transaction"] >= 5,
        "5+ items",
        df["unique_items_in_transaction"].astype(int).astype(str) + " item",
    )
    df.loc[df["unique_items_in_transaction"] == 1, "basket_group"] = "1 item"
    df.loc[df["unique_items_in_transaction"].between(2, 4), "basket_group"] = (
        df.loc[
            df["unique_items_in_transaction"].between(2, 4),
            "unique_items_in_transaction",
        ]
        .astype(int)
        .astype(str)
        + " items"
    )

    grouped = (
        df.groupby("basket_group", as_index=False)
        .agg(
            transaction_count=("transaction_count", "sum"),
            transaction_percent=("transaction_percent", "sum"),
        )
    )

    order = ["1 item", "2 items", "3 items", "4 items", "5+ items"]
    grouped["sort_order"] = grouped["basket_group"].map(
        {label: index for index, label in enumerate(order)}
    )
    grouped = grouped.sort_values("sort_order")

    fig, ax = plt.subplots(figsize=(10.5, 6))
    bars = ax.bar(grouped["basket_group"], grouped["transaction_percent"])

    ax.set_title(
        "How Many Unique Items Are in a Typical Basket?",
        fontsize=18,
        fontweight="bold",
        pad=18,
    )
    ax.set_xlabel("Unique items in transaction", fontsize=12)
    ax.set_ylabel("Share of transactions", fontsize=12)
    ax.yaxis.set_major_formatter(PercentFormatter(1.0))
    ax.grid(axis="y", alpha=0.25)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    for bar, percent, count in zip(
        bars,
        grouped["transaction_percent"],
        grouped["transaction_count"],
    ):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height(),
            f"{percent:.1%}\n{int(count):,} baskets",
            ha="center",
            va="bottom",
            fontsize=10,
        )

    max_percent = float(grouped["transaction_percent"].max())
    ax.set_ylim(0, max_percent * 1.22 if max_percent > 0 else 1)

    fig.tight_layout()
    save_figure(fig, "01_basket_size_distribution.png")


# ---------------------------------------------------------------------
# CHART 2: TOP CATEGORY ADD-ON OPPORTUNITIES
# ---------------------------------------------------------------------
def plot_category_addons() -> None:
    df = read_csv_optional(
        "category_addon_themes.csv",
        [
            "theme",
            "strategic_score",
            "incremental_baskets_vs_random",
            "confidence",
            "lift",
            "support_count",
        ],
    )
    if df.empty:
        return

    df = safe_numeric(
        df,
        [
            "strategic_score",
            "incremental_baskets_vs_random",
            "confidence",
            "lift",
            "support_count",
        ],
    )
    if df.empty:
        return

    top = (
        df.sort_values(
            ["strategic_score", "incremental_baskets_vs_random"],
            ascending=False,
        )
        .head(TOP_CATEGORY_RULES)
        .sort_values("incremental_baskets_vs_random", ascending=True)
        .copy()
    )
    top["display_rule"] = top["theme"].map(lambda x: wrap_label(x, 38))

    figure_height = max(6.5, 0.64 * len(top) + 2.6)
    fig, ax = plt.subplots(figsize=(13.5, figure_height))
    ax.barh(top["display_rule"], top["incremental_baskets_vs_random"])

    ax.set_title(
        "Top Category Add-On Opportunities",
        fontsize=18,
        fontweight="bold",
        pad=18,
    )
    ax.set_xlabel(
        "Observed baskets above the count expected if categories were independent",
        fontsize=11.5,
    )
    ax.set_ylabel("")
    ax.grid(axis="x", alpha=0.25)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    annotate_bar_metrics(
        ax,
        top.reset_index(drop=True),
        value_column="incremental_baskets_vs_random",
    )

    fig.tight_layout()
    save_figure(fig, "02_top_category_addon_opportunities.png")


# ---------------------------------------------------------------------
# CHART 3: CATEGORY OPPORTUNITY MAP
# ---------------------------------------------------------------------
def plot_category_opportunity_map() -> None:
    df = read_csv_optional(
        "category_addon_themes.csv",
        [
            "theme",
            "strategic_score",
            "confidence",
            "lift",
            "support_count",
            "incremental_baskets_vs_random",
        ],
    )
    if df.empty:
        return

    df = safe_numeric(
        df,
        [
            "strategic_score",
            "confidence",
            "lift",
            "support_count",
            "incremental_baskets_vs_random",
        ],
    )
    if df.empty:
        return

    # Keep enough points to show the portfolio without creating visual noise.
    plot_df = (
        df.sort_values("strategic_score", ascending=False)
        .head(35)
        .copy()
    )

    bubble_size = np.sqrt(plot_df["support_count"].clip(lower=1)) * 10

    fig, ax = plt.subplots(figsize=(11.5, 7.2))
    ax.scatter(
        plot_df["confidence"],
        plot_df["lift"],
        s=bubble_size,
        alpha=0.62,
    )

    label_df = plot_df.head(TOP_SCATTER_LABELS)
    for _, row in label_df.iterrows():
        ax.annotate(
            wrap_label(row["theme"], 28),
            (row["confidence"], row["lift"]),
            xytext=(7, 7),
            textcoords="offset points",
            fontsize=8.8,
        )

    ax.set_title(
        "Category Opportunity Map",
        fontsize=18,
        fontweight="bold",
        pad=18,
    )
    ax.set_xlabel("Confidence: share of antecedent baskets adding the item", fontsize=11)
    ax.set_ylabel("Lift: strength relative to baseline", fontsize=11)
    ax.xaxis.set_major_formatter(PercentFormatter(1.0))
    ax.grid(alpha=0.25)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    fig.text(
        0.5,
        0.01,
        "Bubble size represents rule support. Upper-right opportunities combine stronger conversion and stronger association.",
        ha="center",
        fontsize=9.5,
    )

    fig.tight_layout(rect=(0, 0.04, 1, 1))
    save_figure(fig, "03_category_opportunity_map.png")


# ---------------------------------------------------------------------
# CHART 4: THIRD-ITEM OPPORTUNITIES
# ---------------------------------------------------------------------
def plot_third_item_opportunities() -> None:
    df = read_csv_optional(
        "category_third_item_themes.csv",
        [
            "theme",
            "strategic_score",
            "incremental_baskets_vs_random",
            "confidence",
            "lift",
            "support_count",
        ],
    )
    if df.empty:
        return

    df = safe_numeric(
        df,
        [
            "strategic_score",
            "incremental_baskets_vs_random",
            "confidence",
            "lift",
            "support_count",
        ],
    )
    if df.empty:
        return

    top = (
        df.sort_values(
            ["strategic_score", "incremental_baskets_vs_random"],
            ascending=False,
        )
        .head(TOP_THIRD_ITEM_RULES)
        .sort_values("incremental_baskets_vs_random", ascending=True)
        .copy()
    )
    top["display_rule"] = top["theme"].map(lambda x: wrap_label(x, 44))

    figure_height = max(6.5, 0.72 * len(top) + 2.5)
    fig, ax = plt.subplots(figsize=(14, figure_height))
    ax.barh(top["display_rule"], top["incremental_baskets_vs_random"])

    ax.set_title(
        "Best Opportunities to Add a Third Category",
        fontsize=18,
        fontweight="bold",
        pad=18,
    )
    ax.set_xlabel(
        "Observed baskets above the count expected if categories were independent",
        fontsize=11.5,
    )
    ax.grid(axis="x", alpha=0.25)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    annotate_bar_metrics(
        ax,
        top.reset_index(drop=True),
        value_column="incremental_baskets_vs_random",
    )

    fig.tight_layout()
    save_figure(fig, "04_top_third_item_opportunities.png")


# ---------------------------------------------------------------------
# CHART 5: TOP CROSS-CATEGORY SKU BUNDLES
# ---------------------------------------------------------------------
def plot_strategic_sku_bundles() -> None:
    df = read_csv_optional(
        "strategic_cross_category_skus.csv",
        [
            "antecedent",
            "consequent",
            "strategic_score",
            "incremental_baskets_vs_random",
            "confidence",
            "lift",
            "support_count",
        ],
    )
    if df.empty:
        return

    df = safe_numeric(
        df,
        [
            "strategic_score",
            "incremental_baskets_vs_random",
            "confidence",
            "lift",
            "support_count",
        ],
    )
    if df.empty:
        return

    top = (
        df.sort_values(
            ["strategic_score", "incremental_baskets_vs_random"],
            ascending=False,
        )
        .head(TOP_SKU_RULES)
        .sort_values("incremental_baskets_vs_random", ascending=True)
        .copy()
    )

    top["display_rule"] = (
        top["antecedent"].astype(str)
        + "\n→ "
        + top["consequent"].astype(str)
    ).map(lambda x: wrap_label(x, 50))

    figure_height = max(7.5, 0.78 * len(top) + 2.6)
    fig, ax = plt.subplots(figsize=(15, figure_height))
    ax.barh(top["display_rule"], top["incremental_baskets_vs_random"])

    ax.set_title(
        "Top Cross-Category SKU Bundle Candidates",
        fontsize=18,
        fontweight="bold",
        pad=18,
    )
    ax.set_xlabel(
        "Observed baskets above the count expected if SKUs were independent",
        fontsize=11.5,
    )
    ax.grid(axis="x", alpha=0.25)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    annotate_bar_metrics(
        ax,
        top.reset_index(drop=True),
        value_column="incremental_baskets_vs_random",
    )

    fig.tight_layout()
    save_figure(fig, "05_top_cross_category_sku_bundles.png")


# ---------------------------------------------------------------------
# CHART 6: RECOMMENDATION CLASS SUMMARY
# ---------------------------------------------------------------------
def plot_recommendation_class_summary() -> None:
    df = read_csv_optional(
        "all_sku_rules.csv",
        [
            "recommendation_class",
            "incremental_baskets_vs_random",
            "support_count",
        ],
    )
    if df.empty:
        return

    df = safe_numeric(
        df,
        ["incremental_baskets_vs_random", "support_count"],
    )
    if df.empty:
        return

    summary = (
        df.groupby("recommendation_class", as_index=False)
        .agg(
            rule_count=("recommendation_class", "size"),
            total_incremental_baskets=(
                "incremental_baskets_vs_random",
                lambda series: float(series.clip(lower=0).sum()),
            ),
            total_support=("support_count", "sum"),
        )
        .sort_values("total_incremental_baskets", ascending=True)
    )

    fig, ax = plt.subplots(figsize=(11.5, 6.5))
    bars = ax.barh(
        summary["recommendation_class"],
        summary["total_incremental_baskets"],
    )

    ax.set_title(
        "Where the Rule Portfolio Creates the Most Opportunity",
        fontsize=18,
        fontweight="bold",
        pad=18,
    )
    ax.set_xlabel("Total baskets above random expectation across retained rules", fontsize=11.5)
    ax.grid(axis="x", alpha=0.25)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    max_value = float(summary["total_incremental_baskets"].max())
    ax.set_xlim(0, max_value * 1.25 if max_value > 0 else 1)

    for bar, (_, row) in zip(bars, summary.iterrows()):
        ax.text(
            bar.get_width() + max(max_value * 0.015, 0.25),
            bar.get_y() + bar.get_height() / 2,
            f"{int(row['rule_count']):,} rules",
            va="center",
            ha="left",
            fontsize=10,
        )

    fig.tight_layout()
    save_figure(fig, "06_recommendation_class_summary.png")


# ---------------------------------------------------------------------
# VISUALIZATION INDEX
# ---------------------------------------------------------------------
def save_visualization_index() -> None:
    index = pd.DataFrame(
        [
            {
                "file": "01_basket_size_distribution.png",
                "business_use": "Shows whether most transactions have room for an additional item.",
            },
            {
                "file": "02_top_category_addon_opportunities.png",
                "business_use": "Ranks broad placement and promotion themes using incremental baskets, confidence, and lift.",
            },
            {
                "file": "03_category_opportunity_map.png",
                "business_use": "Separates high-confidence rules from rules that are merely unusual.",
            },
            {
                "file": "04_top_third_item_opportunities.png",
                "business_use": "Identifies the strongest category to add when two categories are already present.",
            },
            {
                "file": "05_top_cross_category_sku_bundles.png",
                "business_use": "Provides concrete SKU-level bundle candidates for testing.",
            },
            {
                "file": "06_recommendation_class_summary.png",
                "business_use": "Shows which type of recommendation accounts for the largest total opportunity.",
            },
        ]
    )
    VIS_DIR.mkdir(parents=True, exist_ok=True)
    path = VIS_DIR / "visualization_index.csv"
    index.to_csv(path, index=False)
    print(f"Saved: {path.resolve()}")


def main() -> None:
    VIS_DIR.mkdir(parents=True, exist_ok=True)

    expected_files = [
        "basket_size_distribution.csv",
        "category_addon_themes.csv",
        "category_third_item_themes.csv",
        "strategic_cross_category_skus.csv",
        "all_sku_rules.csv",
    ]
    available_files = [
        filename for filename in expected_files
        if (OUTPUT_DIR / filename).exists()
    ]

    print(f"Visualization script folder: {SCRIPT_DIR.resolve()}")
    print(f"Looking for Apriori CSV files in: {OUTPUT_DIR.resolve()}")

    if not available_files:
        print("\nNo Apriori output CSV files were found.")
        print("Run mashgin_strategic_bundle_fast_csv.py first, then rerun this script.")
        print("Expected at least one of:")
        for filename in expected_files:
            print(f"  - {filename}")
        return

    print("\nFound:")
    for filename in available_files:
        print(f"  - {filename}")

    print("\nCreating Mashgin bundle-analysis visualizations...")
    plot_basket_size_distribution()
    plot_category_addons()
    plot_category_opportunity_map()
    plot_third_item_opportunities()
    plot_strategic_sku_bundles()
    plot_recommendation_class_summary()
    save_visualization_index()

    if SAVE_PNG_FILES:
        print(f"\nPNG files saved in: {VIS_DIR.resolve()}")

    if SHOW_CHARTS_IN_SPYDER:
        print("Charts should also appear in Spyder's Plots pane.")
        plt.show()


if __name__ == "__main__":
    main()