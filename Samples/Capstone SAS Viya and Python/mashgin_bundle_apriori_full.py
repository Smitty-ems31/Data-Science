"""
Mashgin Strategic Bundle Recommendation Analysis - fast CSV version
==================================================================

This script is designed for your Mashgin item-level export and is more useful
for business recommendations than a raw Apriori/lift-ranked table.

It creates CSV files for:
    1. category_addon_themes.csv        X category -> Y category
    2. category_third_item_themes.csv   X category + Y category -> Z category
    3. strategic_cross_category_skus.csv item X -> item Y across categories
    4. same_family_variety_skus.csv     flavor/size variety rules
    5. same_category_cross_brand_skus.csv same category but different family/brand
    6. item_support_metadata.csv        item counts with inferred category/family
    7. basket_size_distribution.csv     how many unique items per basket
    8. readme_summary.csv               key metrics and settings

Install:
    pip install pandas openpyxl numpy

Run from the same folder as the Excel file:
    python mashgin_strategic_bundle_fast_csv.py
"""

from __future__ import annotations

from collections import Counter
from itertools import combinations
from pathlib import Path
import math
import re
from typing import Iterable

import numpy as np
import pandas as pd

INPUT_XLSX = Path("2026-04_Mashgin_Items.xlsx")
SHEET_NAME = "in"
OUTPUT_DIR = Path("mashgin_strategic_outputs")

MIN_ITEM_SUPPORT_COUNT = 100
MIN_ITEM_RULE_SUPPORT_COUNT = 20
MIN_ITEM_CONFIDENCE = 0.015
MIN_ITEM_LIFT = 1.05
MAX_ITEMS_PER_BASKET_FOR_COMBOS = 20

MIN_CATEGORY_RULE_SUPPORT_COUNT = 300
MIN_CATEGORY_CONFIDENCE = 0.03
MIN_CATEGORY_LIFT = 1.05
MIN_CATEGORY_TRIPLE_SUPPORT_COUNT = 100
MIN_CATEGORY_TRIPLE_CONFIDENCE = 0.04
MIN_CATEGORY_TRIPLE_LIFT = 1.05

SUPPRESS_BUSINESS_CATEGORIES = {"Other", "Personal Care"}
REMOVE_ITEM_KEYWORDS = ["TAX", "REFUND", "VOID", "DISCOUNT", "BAG FEE"]


def clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def itemset_to_text(values: Iterable[str]) -> str:
    return " + ".join(sorted(str(x) for x in values))


def infer_category(item_name: str) -> str:
    s = item_name.lower()
    if any(k in s for k in ["toothpaste", "toothbrush", "mouthwash", "listerine", "deodorant", "shampoo", "conditioner", "soap", "tampon", " pads", "advil", "tylenol", "medicine", "bandage"]):
        return "Personal Care"
    if any(k in s for k in ["frosted flakes", "cinnamon toast crunch", "lucky charms", "cereal", "cheerios", "mini wheats", "raisin bran", "froot loops", "captain crunch", "cocoa puffs", "trix", "gm ", "kellogg"]):
        return "Cereal/Breakfast"
    if any(k in s for k in ["banana", "blueberr", "raspberr", "blackberr", "strawberr", "grapes", "apple", "orange", "clementine", "pineapple", "fruit cup", "watermelon", "solely fruit", "mango", "cut fruit"]):
        return "Fruit"
    if any(k in s for k in ["life water", "aquafina", "smartwater", "smart water", "essentia", "fiji", "dasani", "water 20 oz", "water 1 liter", "water 700 ml", "water bottle", "water gallon"]):
        return "Water"
    if any(k in s for k in ["alani", "red bull", "monster", "celsius", "rockstar", "prime energy", "bucked up", "ghost energy", "energy drink"]):
        return "Energy Drink"
    if any(k in s for k in ["pepsi", "coke", "coca", "sprite", "dr pepper", "mountain dew", "mtn dew", "soda", "poppi", "starry", "fanta", "root beer", "ginger ale", "7 up"]):
        return "Soda/Functional Soda"
    if any(k in s for k in ["propel", "gatorade", "powerade", "bodyarmor", "body armor"]):
        return "Sports Drink"
    if any(k in s for k in ["juice", "lemonade", "tea", "coffee", "latte", "frappuccino", "kombucha", "cold brew", "iced tea"]):
        return "Juice/Tea/Coffee"
    if any(k in s for k in ["milk", "creamery chug", "choc recovery"]):
        return "Milk/Dairy Drink"
    if any(k in s for k in ["yogurt", "oikos", "chobani", "clio", "siggi", "activia"]):
        return "Yogurt"
    if any(k in s for k in ["ice cream", "mochi", "popsicle", "frozen", "scoots", "gelato", "icee", "klondike", "novelty"]):
        return "Frozen Dessert"
    if any(k in s for k in ["cookie", "brownie", "brookie", "muffin", "donut", "doughnut", "bakery", "cake", "pastry", "cinnamon roll", "rice krispie treat"]):
        return "Bakery/Dessert"
    if any(k in s for k in ["chips", "doritos", "tostitos", "cheetos", "fritos", "lays", "lay's", "ruffles", "pretzel", "popcorn", "cheez it", "cheez-it", "goldfish", "takis", "pringles", "sunchips", "sun chips", "smartfood", "fl "]):
        return "Salty Snack"
    if any(k in s for k in ["candy", "chocolate", "rees", "hershey", "snicker", "m&m", "skittle", "starburst", "gummy", "sour patch", "twix", "kit kat", "gum", "mint", "nerds", "airheads", "swedish fish", "jolly rancher"]):
        return "Candy/Gum"
    if any(k in s for k in ["protein", "quest", "clif", "kind bar", "granola", "nature valley", "rxbar", "perfect bar", "barebells", "built", "think bar", "power crunch"]):
        return "Nutrition/Bar"
    if any(k in s for k in ["cheese", "sargento", "string cheese", "cheddar", "colby", "pepper jack", "mozzarella"]):
        return "Cheese"
    if any(k in s for k in ["dip", "hummus", "salsa", "guacamole", "buffalo chix", "queso"]):
        return "Dip/Spread"
    if any(k in s for k in ["sandwich", "wrap", "salad", "sushi", "burrito", "meal", "lunchable", "soup", "maruchan", "ramen", "noodle", "mac and cheese", "pizza", "grilled cheese", "panera"]):
        return "Meal/Prepared"
    if any(k in s for k in ["cracker", "crackers"]):
        return "Crackers"
    return "Other"


KNOWN_FAMILIES = [
    "Oikos Triple Zero", "Creamery Chug", "Nature Valley", "Cinnamon Toast Crunch", "Frosted Flakes", "Lucky Charms", "Alani Nu", "My Mochi", "Sargento", "Life Water", "Aquafina", "Propel", "Poppi", "Creamery", "Scoots", "Bar Clio", "Clio", "Tostitos", "Doritos", "Cheetos", "Fritos", "Lay's", "Lays", "Ruffles", "Cheez It", "Cheez-It", "Goldfish", "Quest", "Clif", "Kind", "Reese's", "Reeses", "Hershey", "Snickers", "M&M", "Skittles", "Starburst", "Dr Pepper", "Pepsi", "Coke", "Coca Cola", "Mountain Dew", "Gatorade", "Powerade", "Bodyarmor", "Red Bull", "Celsius", "Monster", "Prime", "Maruchan", "Good Humor", "Demets Flipz", "BH", "Chobani", "Barebells", "Built", "Kellogg", "GM", "Panera", "Pierre", "Heluva Good", "Solely", "Listerine", "Crest", "Kodiak", "Jif"
]


def infer_family(item_name: str) -> str:
    s = item_name.lower()
    for family in sorted(KNOWN_FAMILIES, key=len, reverse=True):
        if family.lower() in s:
            return family
    for family in ["Bananas", "Blueberries", "Raspberries", "Blackberries", "Strawberries", "Grapes", "Apples", "Oranges", "Pineapple", "Mango"]:
        if family.lower().rstrip("s") in s:
            return family
    stop = {"ss", "oz", "ml", "liter", "litre", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "12", "16", "20", "24", "1/2", "pint", "ct", "count"}
    tokens = [t for t in re.split(r"\s+", item_name) if t.lower() not in stop]
    return " ".join(tokens[:2]) if tokens else item_name


def grouped_unique_lists_fast(df: pd.DataFrame, tx_col: str, value_col: str) -> list[list[str]]:
    baskets: dict[str, set[str]] = {}
    for tx, value in df[[tx_col, value_col]].dropna().itertuples(index=False, name=None):
        baskets.setdefault(str(tx), set()).add(str(value))
    return [sorted(v) for v in baskets.values()]


def read_and_clean() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    raw = pd.read_excel(
        INPUT_XLSX,
        sheet_name=SHEET_NAME,
        header=None,
        names=["basket_number", "transaction_id", "basket_code", "local_time", "item_name"],
        engine="openpyxl",
    )
    raw["transaction_id"] = raw["transaction_id"].astype(str).str.strip()
    raw["item_name"] = raw["item_name"].map(clean_text)
    tx_parts = raw["transaction_id"].str.split("-", expand=True)
    raw["location_code"] = tx_parts[1] if tx_parts.shape[1] > 1 else np.nan

    cleaned = raw.dropna(subset=["transaction_id", "item_name"]).copy()
    cleaned = cleaned[cleaned["item_name"] != ""]
    pattern = "|".join(re.escape(x) for x in REMOVE_ITEM_KEYWORDS)
    cleaned = cleaned[~cleaned["item_name"].str.contains(pattern, case=False, na=False)]
    cleaned = cleaned.drop_duplicates(["transaction_id", "item_name"])

    n_tx = cleaned["transaction_id"].nunique()
    item_counts = cleaned.groupby("item_name")["transaction_id"].nunique().sort_values(ascending=False)
    metadata = item_counts.rename("transaction_count").reset_index()
    metadata["support"] = metadata["transaction_count"] / n_tx
    metadata["category"] = metadata["item_name"].map(infer_category)
    metadata["family"] = metadata["item_name"].map(infer_family)

    cleaned = cleaned.merge(metadata[["item_name", "category", "family"]], on="item_name", how="left")
    return raw, cleaned, metadata


def make_summary(raw: pd.DataFrame, cleaned: pd.DataFrame, metadata: pd.DataFrame) -> pd.DataFrame:
    basket_sizes = cleaned["transaction_id"].value_counts()
    rows = [
        ["Input rows", len(raw)],
        ["Cleaned transaction-item rows", len(cleaned)],
        ["Unique transactions", cleaned["transaction_id"].nunique()],
        ["Unique item names", cleaned["item_name"].nunique()],
        ["Average unique items per transaction", round(float(basket_sizes.mean()), 3)],
        ["Median unique items per transaction", round(float(basket_sizes.median()), 3)],
        ["Transactions with 1 unique item", int((basket_sizes == 1).sum())],
        ["Transactions with 2+ unique items", int((basket_sizes >= 2).sum())],
        ["Transactions with 3+ unique items", int((basket_sizes >= 3).sum())],
        ["Business note", "Use Category Addon Themes and Strategic Cross Category SKUs first; treat Same Family Variety as a separate shelf/variety insight."],
    ]
    return pd.DataFrame(rows, columns=["metric", "value"])


def basket_distribution(cleaned: pd.DataFrame) -> pd.DataFrame:
    sizes = cleaned["transaction_id"].value_counts()
    out = sizes.value_counts().sort_index().rename_axis("unique_items_in_transaction").reset_index(name="transaction_count")
    out["transaction_percent"] = out["transaction_count"] / out["transaction_count"].sum()
    return out


def category_rules(cleaned: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    n_tx = cleaned["transaction_id"].nunique()
    work = cleaned[~cleaned["category"].isin(SUPPRESS_BUSINESS_CATEGORIES)]
    baskets = grouped_unique_lists_fast(work, "transaction_id", "category")
    cat_counts = Counter(); pair_counts = Counter(); triple_counts = Counter()
    for cats in baskets:
        cat_counts.update(cats)
        if len(cats) >= 2:
            pair_counts.update(combinations(cats, 2))
        if len(cats) >= 3:
            triple_counts.update(combinations(cats, 3))

    addon_rows = []
    for (a, b), count in pair_counts.items():
        if count < MIN_CATEGORY_RULE_SUPPORT_COUNT:
            continue
        for ant, con in [(a, b), (b, a)]:
            ant_count = cat_counts[ant]; con_count = cat_counts[con]
            conf = count / ant_count; con_support = con_count / n_tx
            lift = conf / con_support if con_support else np.nan
            expected = ant_count * con_count / n_tx
            inc = count - expected
            if conf >= MIN_CATEGORY_CONFIDENCE and lift >= MIN_CATEGORY_LIFT:
                addon_rows.append({
                    "theme": f"{ant} -> {con}",
                    "offer_idea": f"Connect {ant} with {con} through placement or a small add-on deal.",
                    "plain_english": f"Among baskets with {ant}, {conf:.1%} also include {con}; {lift:.2f}x baseline; about {inc:.0f} baskets above random expectation.",
                    "antecedent_category": ant, "consequent_category": con,
                    "support_count": count, "antecedent_count_tx": ant_count, "consequent_count_tx": con_count,
                    "confidence": conf, "lift": lift, "expected_count_if_random": expected,
                    "incremental_baskets_vs_random": inc,
                    "strategic_score": max(inc, 0) * conf * math.log1p(count),
                })

    triple_rows = []
    for (a, b, c), count in triple_counts.items():
        if count < MIN_CATEGORY_TRIPLE_SUPPORT_COUNT:
            continue
        for pair, con in [((a, b), c), ((a, c), b), ((b, c), a)]:
            pair = tuple(sorted(pair))
            pair_count = pair_counts[pair]; con_count = cat_counts[con]
            conf = count / pair_count; con_support = con_count / n_tx
            lift = conf / con_support if con_support else np.nan
            expected = pair_count * con_count / n_tx
            inc = count - expected
            if conf >= MIN_CATEGORY_TRIPLE_CONFIDENCE and lift >= MIN_CATEGORY_TRIPLE_LIFT:
                ant_text = itemset_to_text(pair)
                triple_rows.append({
                    "theme": f"{ant_text} -> {con}",
                    "offer_idea": f"When a basket has {ant_text}, test a deal or placement that adds {con}.",
                    "plain_english": f"Among baskets with {ant_text}, {conf:.1%} also include {con}; {lift:.2f}x baseline; about {inc:.0f} baskets above random expectation.",
                    "antecedent_categories": ant_text, "consequent_category": con,
                    "support_count": count, "antecedent_count_tx": pair_count, "consequent_count_tx": con_count,
                    "confidence": conf, "lift": lift, "expected_count_if_random": expected,
                    "incremental_baskets_vs_random": inc,
                    "strategic_score": max(inc, 0) * conf * math.log1p(count),
                })

    addon = pd.DataFrame(addon_rows).sort_values(["strategic_score", "incremental_baskets_vs_random"], ascending=False) if addon_rows else pd.DataFrame()
    triple = pd.DataFrame(triple_rows).sort_values(["strategic_score", "incremental_baskets_vs_random"], ascending=False) if triple_rows else pd.DataFrame()
    return addon, triple


def item_rules(cleaned: pd.DataFrame) -> pd.DataFrame:
    n_tx = cleaned["transaction_id"].nunique()
    keep = cleaned.groupby("item_name")["transaction_id"].nunique()
    keep_items = keep[keep >= MIN_ITEM_SUPPORT_COUNT].index
    work = cleaned[cleaned["item_name"].isin(keep_items)]
    baskets = grouped_unique_lists_fast(work, "transaction_id", "item_name")
    meta = work[["item_name", "category", "family"]].drop_duplicates("item_name").set_index("item_name")

    item_counts = Counter(); pair_counts = Counter(); skipped = 0
    for items in baskets:
        item_counts.update(items)
        if len(items) > MAX_ITEMS_PER_BASKET_FOR_COMBOS:
            skipped += 1
            continue
        if len(items) >= 2:
            pair_counts.update(combinations(items, 2))

    rows = []
    for (a, b), count in pair_counts.items():
        if count < MIN_ITEM_RULE_SUPPORT_COUNT:
            continue
        for ant, con in [(a, b), (b, a)]:
            ant_count = item_counts[ant]; con_count = item_counts[con]
            conf = count / ant_count; con_support = con_count / n_tx
            lift = conf / con_support if con_support else np.nan
            if conf < MIN_ITEM_CONFIDENCE or lift < MIN_ITEM_LIFT:
                continue
            expected = ant_count * con_count / n_tx
            inc = count - expected
            ant_cat = meta.loc[ant, "category"]; con_cat = meta.loc[con, "category"]
            ant_family = meta.loc[ant, "family"]; con_family = meta.loc[con, "family"]
            same_cat = ant_cat == con_cat; same_family = ant_family == con_family
            suppressed = ant_cat in SUPPRESS_BUSINESS_CATEGORIES or con_cat in SUPPRESS_BUSINESS_CATEGORIES
            if same_family:
                klass = "Same family variety"; weight = 0.5
            elif not same_cat and not suppressed:
                klass = "Strategic cross-category"; weight = 3.0
            elif same_cat and not same_family and not suppressed:
                klass = "Same category cross-brand"; weight = 1.25
            elif not same_cat:
                klass = "Cross-category but review"; weight = 1.0
            else:
                klass = "Low-priority"; weight = 0.5
            rows.append({
                "recommendation_class": klass,
                "bundle_text": f"If basket has [{ant}], suggest [{con}]",
                "offer_idea": f"Test add-on placement or a small deal: {ant} + {con}.",
                "plain_english": f"Among baskets with {ant}, {conf:.1%} also include {con}; {lift:.2f}x baseline; about {inc:.1f} baskets above random expectation.",
                "antecedent": ant, "consequent": con, "target_bundle": itemset_to_text([ant, con]),
                "antecedent_category": ant_cat, "consequent_category": con_cat,
                "antecedent_family": ant_family, "consequent_family": con_family,
                "same_category": same_cat, "same_family": same_family,
                "support_count": count, "antecedent_count_tx": ant_count, "consequent_count_tx": con_count,
                "confidence": conf, "lift": lift, "expected_count_if_random": expected,
                "incremental_baskets_vs_random": inc,
                "strategic_score": max(inc, 0) * conf * math.log1p(count) * weight,
                "skipped_large_baskets": skipped,
            })
    return pd.DataFrame(rows)


def save_csv(df: pd.DataFrame, name: str) -> None:
    if df is not None and not df.empty:
        df.to_csv(OUTPUT_DIR / name, index=False)


def main() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    print("Reading and cleaning...")
    raw, cleaned, metadata = read_and_clean()
    print(f"Transactions: {cleaned['transaction_id'].nunique():,}; unique items: {cleaned['item_name'].nunique():,}")

    print("Building category themes...")
    cat_addons, cat_triples = category_rules(cleaned)

    print("Building SKU rules...")
    sku_rules = item_rules(cleaned)
    sku_rules = sku_rules.sort_values(["recommendation_class", "strategic_score", "support_count"], ascending=[True, False, False]) if not sku_rules.empty else sku_rules

    print("Saving CSV outputs...")
    save_csv(make_summary(raw, cleaned, metadata), "readme_summary.csv")
    save_csv(basket_distribution(cleaned), "basket_size_distribution.csv")
    save_csv(metadata, "item_support_metadata.csv")
    save_csv(cat_addons, "category_addon_themes.csv")
    save_csv(cat_triples, "category_third_item_themes.csv")
    if not sku_rules.empty:
        save_csv(sku_rules[sku_rules["recommendation_class"] == "Strategic cross-category"].sort_values(["strategic_score", "incremental_baskets_vs_random"], ascending=False), "strategic_cross_category_skus.csv")
        save_csv(sku_rules[sku_rules["recommendation_class"] == "Same family variety"].sort_values(["support_count", "confidence"], ascending=False), "same_family_variety_skus.csv")
        save_csv(sku_rules[sku_rules["recommendation_class"] == "Same category cross-brand"].sort_values(["strategic_score", "support_count"], ascending=False), "same_category_cross_brand_skus.csv")
        save_csv(sku_rules, "all_sku_rules.csv")

    print(f"Done. CSV files saved in: {OUTPUT_DIR.resolve()}")


if __name__ == "__main__":
    main()
