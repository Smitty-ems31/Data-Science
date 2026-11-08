

import os
import re
import glob
import shutil
import subprocess
import warnings
from pathlib import Path
from datetime import datetime

import numpy as np
import pandas as pd

import matplotlib
matplotlib.use("Agg")  # Reliable file output in Spyder/headless runs.
import matplotlib.pyplot as plt

from statsmodels.tsa.statespace.sarimax import SARIMAX
from statsmodels.tsa.seasonal import STL
from statsmodels.stats.diagnostic import acorr_ljungbox

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# =============================================================================
# USER SETTINGS
# =============================================================================

# Leave this as None to auto-detect Cstore_weekly_Aug2022_June2026*.xlsx
# in the same folder as this script or in the current working directory.
# Or set it explicitly, for example:
# INPUT_FILE = r"C:\Users\YourName\Documents\Cstore_weekly_Aug2022_June2026.xlsx"
INPUT_FILE = None

# Preferred file name if auto-detect finds more than one file.
INPUT_FILE_NAME = "Cstore_weekly_Aug2022_June2026.xlsx"

# Sheet name. If this sheet is not found, the script uses the first sheet.
SHEET_NAME = "Sheet1"

OUTPUT_DIR_NAME = "arimax_full_output"

DATE_COL = "week_start_date"
BY_COL = "rvcname"
LEGACY_COL = "legacy_ind"

# ARIMAX configuration.
# This is a practical, interpretable AR(1) + exogenous legacy_ind model.
# You can try ARIMA_ORDER=(1,0,1), but it may run longer.
ARIMA_ORDER = (1, 0, 0)
SEASONAL_ORDER = (0, 0, 0, 0)
TREND = "c"
MAX_ITER = 75

# STL annual seasonality for weekly data.
STL_PERIOD = 52

# Missing weeks inside a store's range are filled as zero activity.
# This is usually appropriate for closed/no-transaction weeks.
ADD_MISSING_WEEKS_AS_ZERO = True

# Error metrics ignore the first few fitted values to avoid ARIMA initialization noise.
METRIC_BURN_IN = 5

# Chart settings.
PLOT_DPI = 110
FULL_CHART_WIDTH = 9
FULL_CHART_HEIGHT = 8.8
COMPARISON_CHART_WIDTH = 9
COMPARISON_CHART_HEIGHT = 4.8

# Generate all full appendix decomposition/diagnostic charts.
GENERATE_FULL_APPENDIX_CHARTS = True

# Optional debug filters. Leave as None for full analysis.
# Example: DEBUG_STORES = ["East"]
DEBUG_STORES = None
DEBUG_TARGETS = None

# Try to generate PDF from DOCX if LibreOffice/soffice is available.
GENERATE_PDF = True

# Delete prior output folder at the beginning of each run.
CLEAN_OUTPUT_FOLDER = True


TARGET_ALIASES = {
    "weekly_gross_sales": ["weekly_gross_sales"],
    "weekly_transaction_count": ["weekly_transaction_count"],
    "price_adj_sales_2026": ["price_adj_sales_2026", "price_adjusted_sales_2026"],
    "pop_adj_transactions_2026": ["pop_adj_transactions_2026", "population_adjusted_transactions_2026"],
    "fully_adj_sales_2026": ["fully_adj_sales_2026", "fully_adjust_sales_2026", "fully_adjusted_sales_2026"],
}

TARGET_LABELS = {
    "weekly_gross_sales": "Weekly gross sales",
    "weekly_transaction_count": "Weekly transaction count",
    "price_adj_sales_2026": "Price-adjusted sales, 2026 terms",
    "pop_adj_transactions_2026": "Population-adjusted transactions, 2026 terms",
    "fully_adj_sales_2026": "Fully adjusted sales, 2026 terms",
}


# =============================================================================
# SETUP HELPERS
# =============================================================================


def script_dir() -> Path:
    if "__file__" in globals():
        return Path(__file__).resolve().parent
    return Path.cwd()


def find_input_workbook() -> Path:
    """Find the workbook, preferring one that contains all required adjusted columns."""
    if INPUT_FILE:
        p = Path(INPUT_FILE)
        if p.exists():
            return p
        raise FileNotFoundError(f"INPUT_FILE does not exist: {p}")

    candidates = []
    for base in [script_dir(), Path.cwd()]:
        preferred = base / INPUT_FILE_NAME
        if preferred.exists():
            candidates.append(preferred)
        candidates.extend(base.glob("Cstore_weekly_Aug2022_June2026*.xlsx"))

    # Remove duplicates and Excel temporary files.
    seen = set()
    unique = []
    for c in candidates:
        c = c.resolve()
        if c.name.startswith("~$"):
            continue
        if c not in seen:
            unique.append(c)
            seen.add(c)

    if not unique:
        raise FileNotFoundError(
            "Could not find the C-store workbook. Put the workbook in the same "
            "folder as this script or set INPUT_FILE explicitly."
        )

    def has_required_adjusted_columns(path: Path) -> bool:
        try:
            xl = pd.ExcelFile(path)
            sheet = SHEET_NAME if SHEET_NAME in xl.sheet_names else xl.sheet_names[0]
            header = pd.read_excel(path, sheet_name=sheet, nrows=0)
            cols = {str(c).strip().lower() for c in header.columns}
            required_basic = {DATE_COL.lower(), BY_COL.lower(), LEGACY_COL.lower()}
            if not required_basic.issubset(cols):
                return False
            for aliases in TARGET_ALIASES.values():
                if not any(alias.lower() in cols for alias in aliases):
                    return False
            return True
        except Exception:
            return False

    valid = [p for p in unique if has_required_adjusted_columns(p)]
    if valid:
        # Prefer the newest valid workbook, which is usually the final one.
        return sorted(valid, key=lambda p: p.stat().st_mtime, reverse=True)[0]

    # Fallback: newest candidate, but load_workbook_data will explain missing columns.
    return sorted(unique, key=lambda p: p.stat().st_mtime, reverse=True)[0]


def clean_filename(value) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", str(value)).strip("_")


def make_output_dirs(base_dir: Path):
    output_dir = base_dir / OUTPUT_DIR_NAME
    charts_dir = output_dir / "charts"
    if CLEAN_OUTPUT_FOLDER and output_dir.exists():
        shutil.rmtree(output_dir)
    charts_dir.mkdir(parents=True, exist_ok=True)
    return output_dir, charts_dir


def resolve_targets(df: pd.DataFrame):
    lookup = {c.lower(): c for c in df.columns}
    resolved = {}
    for canonical, aliases in TARGET_ALIASES.items():
        actual = None
        for alias in aliases:
            if alias.lower() in lookup:
                actual = lookup[alias.lower()]
                break
        if actual is None:
            raise KeyError(f"Missing required target column for {canonical}. Tried {aliases}")
        resolved[canonical] = actual
    return resolved


# =============================================================================
# DATA PREP
# =============================================================================


def weighted_mean(values, weights):
    v = pd.to_numeric(values, errors="coerce")
    w = pd.to_numeric(weights, errors="coerce")
    valid = v.notna() & w.notna() & (w > 0)
    if valid.any() and w[valid].sum() > 0:
        return float(np.average(v[valid], weights=w[valid]))
    if v.notna().any():
        return float(v.mean())
    return np.nan


def load_workbook_data(input_file: Path):
    xl = pd.ExcelFile(input_file)
    sheet = SHEET_NAME if SHEET_NAME in xl.sheet_names else xl.sheet_names[0]
    df = pd.read_excel(input_file, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]

    required = [DATE_COL, BY_COL, LEGACY_COL]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise KeyError(f"Workbook is missing required columns: {missing}")

    target_cols = resolve_targets(df)
    return df, target_cols, sheet


def prepare_modeling_data(raw_df: pd.DataFrame, target_cols: dict):
    df = raw_df.copy()
    df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
    df = df.dropna(subset=[DATE_COL, BY_COL])

    numeric_cols = list(target_cols.values()) + [LEGACY_COL]
    if "weekly_transaction_count" in df.columns:
        numeric_cols.append("weekly_transaction_count")

    for c in sorted(set(numeric_cols)):
        df[c] = pd.to_numeric(df[c], errors="coerce")

    rows = []
    for (rvc, week), g in df.groupby([BY_COL, DATE_COL], dropna=False):
        out = {BY_COL: rvc, DATE_COL: week}
        for canonical, actual in target_cols.items():
            out[canonical] = g[actual].sum(min_count=1)

        weight_col = target_cols.get("weekly_transaction_count")
        if weight_col and weight_col in g.columns:
            out["legacy_ind_model"] = weighted_mean(g[LEGACY_COL], g[weight_col])
        else:
            out["legacy_ind_model"] = pd.to_numeric(g[LEGACY_COL], errors="coerce").mean()
        rows.append(out)

    model_df = pd.DataFrame(rows).sort_values([BY_COL, DATE_COL]).reset_index(drop=True)

    if ADD_MISSING_WEEKS_AS_ZERO:
        completed = []
        for rvc, g in model_df.groupby(BY_COL):
            g = g.sort_values(DATE_COL).set_index(DATE_COL)
            full_index = pd.date_range(g.index.min(), g.index.max(), freq="7D")
            g = g.reindex(full_index)
            g.index.name = DATE_COL
            g[BY_COL] = rvc
            for t in target_cols.keys():
                g[t] = g[t].fillna(0.0)
            g["legacy_ind_model"] = g["legacy_ind_model"].ffill().bfill().fillna(0.0)
            completed.append(g.reset_index())
        model_df = pd.concat(completed, ignore_index=True)

    if DEBUG_STORES:
        model_df = model_df[model_df[BY_COL].isin(DEBUG_STORES)].copy()

    return model_df.sort_values([BY_COL, DATE_COL]).reset_index(drop=True)




def make_all_stores_total_data(model_df: pd.DataFrame, target_names: list):
    """Create one all-stores combined weekly time series for total trend analysis."""
    rows = []
    for week, g in model_df.groupby(DATE_COL):
        out = {DATE_COL: week, BY_COL: "All Stores Total"}
        for target in target_names:
            out[target] = g[target].sum(min_count=1)

        # Use transaction-weighted legacy share for the combined system indicator.
        if "weekly_transaction_count" in g.columns:
            w = pd.to_numeric(g["weekly_transaction_count"], errors="coerce")
            v = pd.to_numeric(g["legacy_ind_model"], errors="coerce")
            valid = v.notna() & w.notna() & (w > 0)
            if valid.any() and w[valid].sum() > 0:
                out["legacy_ind_model"] = float(np.average(v[valid], weights=w[valid]))
            else:
                out["legacy_ind_model"] = float(v.mean()) if v.notna().any() else np.nan
        else:
            out["legacy_ind_model"] = float(g["legacy_ind_model"].mean())
        rows.append(out)

    total_df = pd.DataFrame(rows).sort_values(DATE_COL).reset_index(drop=True)
    total_df["legacy_ind_model"] = total_df["legacy_ind_model"].ffill().bfill().fillna(0.0)
    return total_df

# =============================================================================
# MODELING + DECOMPOSITION
# =============================================================================


def calculate_error_metrics(actual, fitted, residuals):
    aligned = pd.concat([actual.rename("actual"), fitted.rename("fitted")], axis=1)
    aligned = aligned.replace([np.inf, -np.inf], np.nan).dropna()
    if METRIC_BURN_IN and len(aligned) > METRIC_BURN_IN:
        aligned = aligned.iloc[METRIC_BURN_IN:].copy()

    if aligned.empty:
        return {
            "MAE": np.nan, "RMSE": np.nan, "MAPE": np.nan,
            "sMAPE": np.nan, "WAPE": np.nan, "MASE": np.nan,
            "LjungBox_pvalue": np.nan,
        }

    err = aligned["actual"] - aligned["fitted"]
    abs_err = err.abs()
    mae = float(abs_err.mean())
    rmse = float(np.sqrt(np.mean(err ** 2)))

    nz = aligned["actual"].abs() > 1e-12
    mape = float((abs_err[nz] / aligned.loc[nz, "actual"].abs()).mean()) if nz.any() else np.nan

    denom = aligned["actual"].abs() + aligned["fitted"].abs()
    valid = denom > 1e-12
    smape = float((2 * abs_err[valid] / denom[valid]).mean()) if valid.any() else np.nan

    total_actual = aligned["actual"].abs().sum()
    wape = float(abs_err.sum() / total_actual) if total_actual > 1e-12 else np.nan

    y = aligned["actual"].astype(float)
    lag = STL_PERIOD if len(y) > STL_PERIOD + 5 else 1
    naive = (y - y.shift(lag)).abs().dropna()
    naive_denom = naive.mean() if len(naive) else np.nan
    mase = float(mae / naive_denom) if naive_denom and naive_denom > 1e-12 else np.nan

    resid_clean = pd.Series(residuals).replace([np.inf, -np.inf], np.nan).dropna()
    if METRIC_BURN_IN and len(resid_clean) > METRIC_BURN_IN:
        resid_clean = resid_clean.iloc[METRIC_BURN_IN:]
    lb_p = np.nan
    if len(resid_clean) > 15:
        lag_lb = min(10, max(1, len(resid_clean) // 5))
        try:
            lb = acorr_ljungbox(resid_clean, lags=[lag_lb], return_df=True)
            lb_p = float(lb["lb_pvalue"].iloc[0])
        except Exception:
            pass

    return {
        "MAE": mae,
        "RMSE": rmse,
        "MAPE": mape,
        "sMAPE": smape,
        "WAPE": wape,
        "MASE": mase,
        "LjungBox_pvalue": lb_p,
    }


def fit_models(model_df: pd.DataFrame, target_names: list):
    results = []
    fitted_all = []
    decomp_all = []

    stores = sorted(model_df[BY_COL].dropna().unique())
    for rvc in stores:
        g0 = model_df[model_df[BY_COL] == rvc].copy()
        print(f"\nStore: {rvc} ({len(g0):,} weekly rows)")

        for target in target_names:
            print(f"  Modeling {target} ...", end=" ", flush=True)
            g = g0[[DATE_COL, BY_COL, target, "legacy_ind_model"]].copy().sort_values(DATE_COL)
            series = g.set_index(DATE_COL)[target].astype(float).replace([np.inf, -np.inf], np.nan).fillna(0.0)
            exog = g.set_index(DATE_COL)[["legacy_ind_model"]].astype(float).ffill().bfill().fillna(0.0)

            inferred = pd.infer_freq(series.index)
            if inferred:
                series = series.asfreq(inferred)
                exog = exog.asfreq(inferred)

            try:
                mod = SARIMAX(
                    series,
                    exog=exog,
                    order=ARIMA_ORDER,
                    seasonal_order=SEASONAL_ORDER,
                    trend=TREND,
                    enforce_stationarity=False,
                    enforce_invertibility=False,
                )
                fit = mod.fit(disp=False, maxiter=MAX_ITER, method="lbfgs")
                fitted = fit.fittedvalues
                resid = series - fitted
                beta = float(fit.params.get("legacy_ind_model", np.nan))
                pval = float(fit.pvalues.get("legacy_ind_model", np.nan))
                aic = float(fit.aic)
                bic = float(fit.bic)
            except Exception as e:
                warnings.warn(f"Model failed for {rvc} / {target}: {e}")
                fitted = pd.Series(index=series.index, data=np.nan)
                resid = series - fitted
                beta = np.nan
                pval = np.nan
                aic = np.nan
                bic = np.nan

            metrics = calculate_error_metrics(series, fitted, resid)
            results.append({
                BY_COL: rvc,
                "target": target,
                "target_label": TARGET_LABELS[target],
                "n_obs": int(len(series)),
                "legacy_beta": beta,
                "legacy_pvalue": pval,
                "mashgin_effect_vs_legacy": -beta if not np.isnan(beta) else np.nan,
                "AIC": aic,
                "BIC": bic,
                **metrics,
            })

            fitted_all.append(pd.DataFrame({
                BY_COL: rvc,
                DATE_COL: series.index,
                "target": target,
                "target_label": TARGET_LABELS[target],
                "actual": series.values,
                "fitted": fitted.values,
                "arimax_residual": resid.values,
                "legacy_ind_model": exog["legacy_ind_model"].values,
            }))

            stl_period = STL_PERIOD if len(series) >= STL_PERIOD * 2 else max(4, min(26, len(series) // 2))
            stl = STL(series, period=stl_period, robust=True).fit()
            decomp_all.append(pd.DataFrame({
                BY_COL: rvc,
                DATE_COL: series.index,
                "target": target,
                "target_label": TARGET_LABELS[target],
                "observed": series.values,
                "trend": stl.trend.values,
                "seasonal": stl.seasonal.values,
                "stl_residual": stl.resid.values,
                "legacy_ind_model": exog["legacy_ind_model"].values,
            }))
            print("done", flush=True)

    return pd.DataFrame(results), pd.concat(fitted_all, ignore_index=True), pd.concat(decomp_all, ignore_index=True)


# =============================================================================
# CHARTING
# =============================================================================


def save_fig(fig, path: Path, dpi=PLOT_DPI):
    fig.savefig(path, dpi=dpi)
    plt.close(fig)


def plot_target_store_trends(decomp: pd.DataFrame, charts_dir: Path):
    paths = []
    for target in sorted(decomp["target"].unique(), key=lambda x: list(TARGET_LABELS).index(x)):
        d = decomp[decomp["target"] == target].copy()
        fig, ax = plt.subplots(figsize=(COMPARISON_CHART_WIDTH, COMPARISON_CHART_HEIGHT))
        for rvc, g in d.groupby(BY_COL):
            g = g.sort_values(DATE_COL)
            ax.plot(g[DATE_COL], g["trend"], label=rvc, linewidth=1.6)
        ax.set_title(f"STL trend by store: {TARGET_LABELS[target]}")
        ax.set_xlabel("Week")
        ax.set_ylabel(TARGET_LABELS[target])
        ax.grid(True, alpha=0.25)
        ax.legend(ncol=2)
        fig.tight_layout()
        path = charts_dir / f"comparison_store_trends__{clean_filename(target)}.png"
        save_fig(fig, path)
        paths.append((target, path))
    return paths


def plot_store_normalized_measure_trends(decomp: pd.DataFrame, charts_dir: Path):
    paths = []
    for rvc in sorted(decomp[BY_COL].unique()):
        d = decomp[decomp[BY_COL] == rvc].copy()
        fig, ax = plt.subplots(figsize=(COMPARISON_CHART_WIDTH, COMPARISON_CHART_HEIGHT))
        for target, g in d.groupby("target"):
            g = g.sort_values(DATE_COL)
            trend = g["trend"].astype(float)
            nonzero = trend[trend.abs() > 1e-9]
            base = nonzero.iloc[0] if len(nonzero) else np.nan
            if pd.isna(base) or abs(base) < 1e-9:
                norm = trend
            else:
                norm = trend / base * 100.0
            ax.plot(g[DATE_COL], norm, label=TARGET_LABELS[target], linewidth=1.4)
        ax.set_title(f"Normalized STL trend comparison by measure: {rvc}")
        ax.set_xlabel("Week")
        ax.set_ylabel("Index, first nonzero trend = 100")
        ax.grid(True, alpha=0.25)
        ax.legend(fontsize=7)
        fig.tight_layout()
        path = charts_dir / f"comparison_normalized_measure_trends__{clean_filename(rvc)}.png"
        save_fig(fig, path)
        paths.append((rvc, path))
    return paths


def plot_effect_bars(results: pd.DataFrame, charts_dir: Path):
    paths = []
    for target in sorted(results["target"].unique(), key=lambda x: list(TARGET_LABELS).index(x)):
        d = results[results["target"] == target].copy().sort_values(BY_COL)
        fig, ax = plt.subplots(figsize=(8.5, 4.3))
        ax.bar(d[BY_COL], d["mashgin_effect_vs_legacy"])
        ax.axhline(0, linewidth=1)
        ax.set_title(f"Model-estimated Mashgin effect vs legacy: {TARGET_LABELS[target]}")
        ax.set_xlabel("Store")
        ax.set_ylabel("Mashgin effect vs legacy")
        ax.grid(True, axis="y", alpha=0.25)
        for i, (_, row) in enumerate(d.iterrows()):
            val = row["mashgin_effect_vs_legacy"]
            if pd.notna(val):
                label = f"{val:,.0f}" if abs(val) >= 1000 else f"{val:,.2f}"
                ax.text(i, val, label, ha="center", va="bottom" if val >= 0 else "top", fontsize=7)
        fig.tight_layout()
        path = charts_dir / f"effect_mashgin_vs_legacy__{clean_filename(target)}.png"
        save_fig(fig, path)
        paths.append((target, path))
    return paths


def plot_full_analysis_charts(fitted: pd.DataFrame, decomp: pd.DataFrame, charts_dir: Path):
    paths = []
    if not GENERATE_FULL_APPENDIX_CHARTS:
        return paths

    grouped = list(decomp.groupby([BY_COL, "target"]))
    for i, ((rvc, target), d) in enumerate(grouped, start=1):
        print(f"  Chart {i}/{len(grouped)}: {rvc} / {target}", flush=True)
        d = d.sort_values(DATE_COL).copy()
        f = fitted[(fitted[BY_COL] == rvc) & (fitted["target"] == target)][
            [DATE_COL, "fitted", "arimax_residual"]
        ].copy()
        combo = d.merge(f, on=DATE_COL, how="left")

        fig, axes = plt.subplots(5, 1, figsize=(FULL_CHART_WIDTH, FULL_CHART_HEIGHT), sharex=True)
        axes[0].plot(combo[DATE_COL], combo["observed"], label="Actual", linewidth=1.1)
        axes[0].plot(combo[DATE_COL], combo["fitted"], label="ARIMAX fitted", linewidth=1.0)
        axes[0].set_title("Observed weekly time series and ARIMAX fitted values")
        axes[0].set_ylabel("Value")
        axes[0].legend(loc="upper left", ncol=2)
        axes[0].grid(True, alpha=0.2)

        axes[1].plot(combo[DATE_COL], combo["observed"], label="Observed", linewidth=0.8, alpha=0.75)
        axes[1].plot(combo[DATE_COL], combo["trend"], label="STL trend", linewidth=1.25)
        axes[1].set_title("STL trend")
        axes[1].set_ylabel("Trend")
        axes[1].legend(loc="upper left", ncol=2)
        axes[1].grid(True, alpha=0.2)

        axes[2].plot(combo[DATE_COL], combo["seasonal"], linewidth=1.0)
        axes[2].set_title("STL seasonal component")
        axes[2].set_ylabel("Seasonal")
        axes[2].grid(True, alpha=0.2)

        axes[3].plot(combo[DATE_COL], combo["stl_residual"], linewidth=1.0)
        axes[3].axhline(0, linewidth=0.8)
        axes[3].set_title("STL residual component")
        axes[3].set_ylabel("STL resid")
        axes[3].grid(True, alpha=0.2)

        axes[4].plot(combo[DATE_COL], combo["arimax_residual"], linewidth=1.0)
        axes[4].axhline(0, linewidth=0.8)
        axes[4].set_title("ARIMAX residuals")
        axes[4].set_ylabel("ARIMAX resid")
        axes[4].set_xlabel("Week")
        axes[4].grid(True, alpha=0.2)

        fig.suptitle(f"{rvc}: {TARGET_LABELS[target]}", y=0.998, fontsize=12)
        fig.tight_layout(rect=[0, 0, 1, 0.98])
        path = charts_dir / f"full_analysis__{clean_filename(rvc)}__{clean_filename(target)}.png"
        save_fig(fig, path, dpi=100)
        paths.append((rvc, target, path))

    return paths




def plot_total_fully_adjusted_trend(total_decomp: pd.DataFrame, charts_dir: Path):
    """Plot one combined all-stores fully adjusted trend line."""
    target = "fully_adj_sales_2026"
    d = total_decomp[total_decomp["target"] == target].sort_values(DATE_COL).copy()
    fig, ax = plt.subplots(figsize=(COMPARISON_CHART_WIDTH, COMPARISON_CHART_HEIGHT))
    ax.plot(d[DATE_COL], d["observed"], label="Observed all-stores total", linewidth=1.0, alpha=0.7)
    ax.plot(d[DATE_COL], d["trend"], label="All-stores STL trend", linewidth=2.2)
    ax.set_title("All-stores combined trend: Fully adjusted sales, 2026 terms")
    ax.set_xlabel("Week")
    ax.set_ylabel("Combined fully adjusted sales")
    ax.grid(True, alpha=0.25)
    ax.legend()
    fig.tight_layout()
    path = charts_dir / "all_stores_total__fully_adjusted_sales_trend.png"
    save_fig(fig, path)
    return path


def plot_total_normalized_measure_trends(total_decomp: pd.DataFrame, charts_dir: Path):
    """Plot normalized all-stores total trends across all five dependent-variable versions."""
    fig, ax = plt.subplots(figsize=(COMPARISON_CHART_WIDTH, COMPARISON_CHART_HEIGHT))
    for target, g in total_decomp.groupby("target"):
        g = g.sort_values(DATE_COL)
        trend = g["trend"].astype(float)
        nonzero = trend[trend.abs() > 1e-9]
        base = nonzero.iloc[0] if len(nonzero) else np.nan
        norm = trend if pd.isna(base) or abs(base) < 1e-9 else trend / base * 100.0
        ax.plot(g[DATE_COL], norm, label=TARGET_LABELS[target], linewidth=1.45)
    ax.set_title("All-stores total normalized STL trend comparison")
    ax.set_xlabel("Week")
    ax.set_ylabel("Index, first nonzero trend = 100")
    ax.grid(True, alpha=0.25)
    ax.legend(fontsize=7)
    fig.tight_layout()
    path = charts_dir / "all_stores_total__normalized_measure_trends.png"
    save_fig(fig, path)
    return path


def plot_total_full_fully_adjusted_chart(total_fitted: pd.DataFrame, total_decomp: pd.DataFrame, charts_dir: Path):
    """Full 5-panel analysis chart for all-stores fully adjusted sales."""
    target = "fully_adj_sales_2026"
    d = total_decomp[total_decomp["target"] == target].sort_values(DATE_COL).copy()
    f = total_fitted[total_fitted["target"] == target][[DATE_COL, "fitted", "arimax_residual"]].copy()
    combo = d.merge(f, on=DATE_COL, how="left")

    fig, axes = plt.subplots(5, 1, figsize=(FULL_CHART_WIDTH, FULL_CHART_HEIGHT), sharex=True)
    axes[0].plot(combo[DATE_COL], combo["observed"], label="Actual all-stores total", linewidth=1.1)
    axes[0].plot(combo[DATE_COL], combo["fitted"], label="ARIMAX fitted", linewidth=1.0)
    axes[0].set_title("Observed all-stores total and ARIMAX fitted values")
    axes[0].set_ylabel("Value")
    axes[0].legend(loc="upper left", ncol=2)
    axes[0].grid(True, alpha=0.2)

    axes[1].plot(combo[DATE_COL], combo["observed"], label="Observed", linewidth=0.8, alpha=0.75)
    axes[1].plot(combo[DATE_COL], combo["trend"], label="STL trend", linewidth=1.25)
    axes[1].set_title("STL trend")
    axes[1].set_ylabel("Trend")
    axes[1].legend(loc="upper left", ncol=2)
    axes[1].grid(True, alpha=0.2)

    axes[2].plot(combo[DATE_COL], combo["seasonal"], linewidth=1.0)
    axes[2].set_title("STL seasonal component")
    axes[2].set_ylabel("Seasonal")
    axes[2].grid(True, alpha=0.2)

    axes[3].plot(combo[DATE_COL], combo["stl_residual"], linewidth=1.0)
    axes[3].axhline(0, linewidth=0.8)
    axes[3].set_title("STL residual component")
    axes[3].set_ylabel("STL resid")
    axes[3].grid(True, alpha=0.2)

    axes[4].plot(combo[DATE_COL], combo["arimax_residual"], linewidth=1.0)
    axes[4].axhline(0, linewidth=0.8)
    axes[4].set_title("ARIMAX residuals")
    axes[4].set_ylabel("ARIMAX resid")
    axes[4].set_xlabel("Week")
    axes[4].grid(True, alpha=0.2)

    fig.suptitle("All Stores Total: Fully adjusted sales, 2026 terms", y=0.998, fontsize=12)
    fig.tight_layout(rect=[0, 0, 1, 0.98])
    path = charts_dir / "all_stores_total__fully_adjusted_sales_full_analysis.png"
    save_fig(fig, path, dpi=100)
    return path


def plot_fully_adjusted_store_plus_total_trend(store_decomp: pd.DataFrame, total_decomp: pd.DataFrame, charts_dir: Path):
    """Two-panel chart: store-level fully adjusted trends and single all-stores total trend."""
    target = "fully_adj_sales_2026"
    store_d = store_decomp[store_decomp["target"] == target].copy()
    total_d = total_decomp[total_decomp["target"] == target].sort_values(DATE_COL).copy()

    fig, axes = plt.subplots(2, 1, figsize=(COMPARISON_CHART_WIDTH, 7.2), sharex=True)
    for rvc, g in store_d.groupby(BY_COL):
        g = g.sort_values(DATE_COL)
        axes[0].plot(g[DATE_COL], g["trend"], label=rvc, linewidth=1.5)
    axes[0].set_title("Store-level STL trends: Fully adjusted sales, 2026 terms")
    axes[0].set_ylabel("Store trend")
    axes[0].grid(True, alpha=0.25)
    axes[0].legend(ncol=2)

    axes[1].plot(total_d[DATE_COL], total_d["observed"], label="All stores observed total", linewidth=0.9, alpha=0.65)
    axes[1].plot(total_d[DATE_COL], total_d["trend"], label="All stores total STL trend", linewidth=2.2)
    axes[1].set_title("Single combined all-stores total trend")
    axes[1].set_xlabel("Week")
    axes[1].set_ylabel("Combined total")
    axes[1].grid(True, alpha=0.25)
    axes[1].legend()

    fig.suptitle("Fully adjusted sales: stores separately and all stores combined", y=0.995)
    fig.tight_layout(rect=[0, 0, 1, 0.97])
    path = charts_dir / "all_stores_total__fully_adjusted_store_plus_total_trend.png"
    save_fig(fig, path)
    return path


def create_total_charts(total_results: pd.DataFrame, total_fitted: pd.DataFrame, total_decomp: pd.DataFrame, charts_dir: Path, store_decomp: pd.DataFrame = None):
    print("\nCreating all-stores total trend charts...", flush=True)
    out = {
        "fully_adjusted_trend": plot_total_fully_adjusted_trend(total_decomp, charts_dir),
        "normalized_measure_trends": plot_total_normalized_measure_trends(total_decomp, charts_dir),
        "fully_adjusted_full_analysis": plot_total_full_fully_adjusted_chart(total_fitted, total_decomp, charts_dir),
    }
    if store_decomp is not None:
        out["store_plus_total_fully_adjusted"] = plot_fully_adjusted_store_plus_total_trend(store_decomp, total_decomp, charts_dir)
    return out

def create_charts(results: pd.DataFrame, fitted: pd.DataFrame, decomp: pd.DataFrame, charts_dir: Path):
    print("\nCreating comparison charts...", flush=True)
    store_trends = plot_target_store_trends(decomp, charts_dir)
    measure_trends = plot_store_normalized_measure_trends(decomp, charts_dir)
    effects = plot_effect_bars(results, charts_dir)

    print("\nCreating full decomposition/diagnostic appendix charts...", flush=True)
    full = plot_full_analysis_charts(fitted, decomp, charts_dir)

    return {
        "store_trends": store_trends,
        "measure_trends": measure_trends,
        "effects": effects,
        "full": full,
    }


# =============================================================================
# DOCUMENT GENERATION
# =============================================================================


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_text(cell, text, bold=False):
    cell.text = ""
    p = cell.paragraphs[0]
    r = p.add_run(str(text))
    r.bold = bold
    r.font.size = Pt(8)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP


def fmt_num(x, kind="num"):
    if pd.isna(x):
        return "n/a"
    if kind == "p":
        return f"{x:.3f}"
    if kind == "pct":
        return f"{x:.1%}"
    if kind == "small":
        return f"{x:.2f}"
    if abs(x) >= 1000:
        return f"{x:,.0f}"
    return f"{x:,.2f}"


def add_table_from_df(doc: Document, df: pd.DataFrame, columns: list, headers: list):
    table = doc.add_table(rows=1, cols=len(columns))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"

    for i, header in enumerate(headers):
        set_cell_text(table.rows[0].cells[i], header, bold=True)
        set_cell_shading(table.rows[0].cells[i], "D9EAF7")

    for _, row in df.iterrows():
        cells = table.add_row().cells
        for i, col in enumerate(columns):
            val = row.get(col, "")
            if col in ["legacy_pvalue", "LjungBox_pvalue"]:
                txt = fmt_num(val, "p")
            elif col in ["WAPE", "MAPE", "sMAPE"]:
                txt = fmt_num(val, "pct")
            elif col == "MASE":
                txt = fmt_num(val, "small")
            elif col in ["legacy_beta", "mashgin_effect_vs_legacy", "MAE", "RMSE"]:
                txt = fmt_num(val)
            else:
                txt = str(val)
            set_cell_text(cells[i], txt)
    return table


def add_picture(doc: Document, path: Path, width=7.0):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(str(path), width=Inches(width))


def build_report(results: pd.DataFrame, charts: dict, output_dir: Path, total_results: pd.DataFrame = None, total_charts: dict = None):
    docx_path = output_dir / "Cstore_ARIMAX_Full_Analysis_Report.docx"

    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(0.55)
    sec.bottom_margin = Inches(0.55)
    sec.left_margin = Inches(0.55)
    sec.right_margin = Inches(0.55)

    doc.styles["Normal"].font.name = "Arial"
    doc.styles["Normal"].font.size = Pt(9)
    doc.styles["Heading 1"].font.name = "Arial"
    doc.styles["Heading 1"].font.size = Pt(16)
    doc.styles["Heading 2"].font.name = "Arial"
    doc.styles["Heading 2"].font.size = Pt(13)
    doc.styles["Heading 3"].font.name = "Arial"
    doc.styles["Heading 3"].font.size = Pt(11)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("C-store ARIMAX and STL Decomposition Analysis")
    r.bold = True
    r.font.size = Pt(20)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(
        "Full chart package comparing raw, price-adjusted, population-adjusted, and fully adjusted outcomes"
    ).italic = True

    doc.add_paragraph(f"Generated: {datetime.now().strftime('%B %d, %Y')}.")

    doc.add_heading("Executive summary", 1)
    doc.add_paragraph(
        "This report compares five dependent-variable versions across stores: weekly gross sales, "
        "weekly transaction count, price-adjusted sales, population-adjusted transactions, and fully "
        "adjusted sales. Each measure is modeled by store with legacy_ind as the system-change indicator, "
        "then decomposed with STL to show observed time series, trend, seasonality, and residual behavior."
    )
    doc.add_paragraph(
        "The comparison charts answer two questions: whether each store shows similar trend behavior for "
        "a given outcome, and whether the adjusted versions tell a different story than raw sales or raw "
        "transaction counts."
    )
    doc.add_paragraph(
        "Model interpretation: legacy_ind equals 1 for legacy POS and 0 for Mashgin. The reported "
        "legacy_beta is legacy level minus Mashgin level. The Mashgin effect column is therefore -legacy_beta."
    )

    if total_results is not None and total_charts is not None:
        doc.add_heading("All-stores total trend analysis", 1)
        doc.add_paragraph(
            "This section aggregates East, Pollock, South, and West into one weekly all-stores total. "
            "The fully adjusted total trend is the executive-level view that combines all four stores "
            "after applying the price and population adjustments."
        )
        doc.add_paragraph(
            "The store-level chart shows the four locations separately. The combined-total charts below "
            "show one full all-stores trend line across the full analysis period."
        )

        total_summary = total_results.copy()
        total_summary["store"] = total_summary[BY_COL]
        total_summary["measure"] = total_summary["target_label"]
        doc.add_heading("All-stores total model result summary", 2)
        add_table_from_df(
            doc,
            total_summary,
            ["store", "measure", "legacy_beta", "legacy_pvalue", "mashgin_effect_vs_legacy", "WAPE", "MASE", "LjungBox_pvalue"],
            ["Series", "Dependent variable", "Legacy beta", "p-value", "Mashgin effect", "WAPE", "MASE", "Ljung-Box p"],
        )

        doc.add_heading("All-stores total normalized trend comparison", 2)
        add_picture(doc, total_charts["normalized_measure_trends"], width=7.0)

        doc.add_heading("All-stores total fully adjusted trend", 2)
        add_picture(doc, total_charts["fully_adjusted_trend"], width=7.0)

        if "store_plus_total_fully_adjusted" in total_charts:
            doc.add_heading("Fully adjusted store trends plus all-stores total", 2)
            doc.add_paragraph(
                "The first panel repeats the store-level trend comparison. The second panel shows the "
                "single combined all-stores total trend line for the fully adjusted outcome."
            )
            add_picture(doc, total_charts["store_plus_total_fully_adjusted"], width=7.0)

        doc.add_heading("All-stores total fully adjusted decomposition and diagnostic chart", 2)
        add_picture(doc, total_charts["fully_adjusted_full_analysis"], width=7.0)

    doc.add_heading("Model result summary", 1)
    summary = results.copy()
    summary["store"] = summary[BY_COL]
    summary["measure"] = summary["target_label"]
    add_table_from_df(
        doc,
        summary,
        ["store", "measure", "legacy_beta", "legacy_pvalue", "mashgin_effect_vs_legacy", "WAPE", "MASE", "LjungBox_pvalue"],
        ["Store", "Dependent variable", "Legacy beta", "p-value", "Mashgin effect", "WAPE", "MASE", "Ljung-Box p"],
    )

    doc.add_heading("Trend comparison across stores by dependent variable", 1)
    for target, path in charts["store_trends"]:
        doc.add_heading(TARGET_LABELS[target], 2)
        add_picture(doc, path, width=7.0)

    doc.add_heading("Trend comparison across dependent variables by store", 1)
    for rvc, path in charts["measure_trends"]:
        doc.add_heading(f"{rvc}: normalized trend comparison", 2)
        add_picture(doc, path, width=7.0)

    doc.add_heading("Estimated Mashgin effect comparison", 1)
    doc.add_paragraph(
        "These charts compare the ARIMAX-estimated Mashgin effect versus the legacy period for each "
        "dependent variable and store. Negative bars indicate lower fitted Mashgin-period levels than "
        "legacy-period levels for that measure; positive bars indicate higher fitted Mashgin-period levels."
    )
    for target, path in charts["effects"]:
        doc.add_heading(TARGET_LABELS[target], 2)
        add_picture(doc, path, width=6.7)

    doc.add_heading("Full decomposition and ARIMAX diagnostic appendix", 1)
    doc.add_paragraph(
        "Each appendix chart contains actual weekly time series, ARIMAX fitted values, STL trend, "
        "STL seasonal component, STL residual, and ARIMAX residuals."
    )
    for rvc, target, path in charts["full"]:
        doc.add_heading(f"{rvc}: {TARGET_LABELS[target]}", 2)
        add_picture(doc, path, width=7.0)
        doc.add_page_break()

    doc.save(docx_path)
    return docx_path


def try_convert_to_pdf(docx_path: Path, output_dir: Path):
    if not GENERATE_PDF:
        return None

    soffice = shutil.which("soffice") or shutil.which("libreoffice") or shutil.which("soffice.exe")
    if not soffice:
        print("LibreOffice/soffice not found on PATH. Skipping PDF export.")
        return None

    try:
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(output_dir), str(docx_path)],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        pdf_path = output_dir / (docx_path.stem + ".pdf")
        if pdf_path.exists():
            return pdf_path
    except Exception as e:
        print(f"PDF export failed: {e}")
    return None


def write_readme(output_dir: Path, input_file: Path, sheet_name: str):
    txt = f"""C-store ARIMAX full output
Generated: {datetime.now().isoformat(timespec='seconds')}
Input workbook: {input_file}
Input sheet: {sheet_name}

Main files:
- Cstore_ARIMAX_Full_Analysis_Report.docx
- Cstore_ARIMAX_Full_Analysis_Report.pdf, if PDF export is available
- arimax_model_results.csv
- arimax_fitted_values.csv
- stl_decomposition_values.csv
- modeling_weekly_prepared_data.csv
- arimax_total_model_results.csv
- arimax_total_fitted_values.csv
- stl_total_decomposition_values.csv
- modeling_weekly_all_stores_total.csv
- charts/*.png

Key all-stores total charts:
- charts/all_stores_total__fully_adjusted_sales_trend.png
- charts/all_stores_total__fully_adjusted_store_plus_total_trend.png
- charts/all_stores_total__fully_adjusted_sales_full_analysis.png
- charts/all_stores_total__normalized_measure_trends.png

Model:
- SARIMAX with order={ARIMA_ORDER}, seasonal_order={SEASONAL_ORDER}, trend={TREND}
- Exogenous variable: legacy_ind_model

Interpretation:
- legacy_ind_model = 1 means legacy POS period.
- legacy_ind_model = 0 means Mashgin period.
- legacy_beta = legacy level - Mashgin level.
- mashgin_effect_vs_legacy = -legacy_beta.

Error metrics:
- WAPE is often the most business-friendly error measure in this project.
- MASE below 1 means the model beats the naive benchmark used by the script.
- Ljung-Box p <= 0.05 means residual autocorrelation likely remains.
"""
    (output_dir / "README_outputs.txt").write_text(txt, encoding="utf-8")


# =============================================================================
# MAIN
# =============================================================================


def main():
    warnings.filterwarnings("ignore")
    start = datetime.now()

    base_dir = script_dir()
    input_file = find_input_workbook()
    output_dir, charts_dir = make_output_dirs(base_dir)

    print("=" * 80)
    print("C-store full ARIMAX + STL decomposition report")
    print("=" * 80)
    print(f"Input workbook: {input_file}")
    print(f"Output folder:  {output_dir}")

    raw_df, target_cols, sheet = load_workbook_data(input_file)
    target_names = list(target_cols.keys())
    if DEBUG_TARGETS:
        target_names = [t for t in target_names if t in DEBUG_TARGETS]

    print(f"Input sheet: {sheet}")
    print("Resolved target columns:")
    for canonical, actual in target_cols.items():
        if canonical in target_names:
            print(f"  {canonical} <= {actual}")

    print("\nPreparing store-level modeling data...")
    model_df = prepare_modeling_data(raw_df, {k: v for k, v in target_cols.items() if k in target_names})
    model_df.to_csv(output_dir / "modeling_weekly_prepared_data.csv", index=False)
    print(f"Prepared store-level rows: {len(model_df):,}")

    print("\nPreparing all-stores total weekly series...")
    total_model_df = make_all_stores_total_data(model_df, target_names)
    total_model_df.to_csv(output_dir / "modeling_weekly_all_stores_total.csv", index=False)
    print(f"Prepared all-stores total rows: {len(total_model_df):,}")

    print("\nRunning store-level ARIMAX models and STL decompositions...")
    results, fitted, decomp = fit_models(model_df, target_names)
    results.to_csv(output_dir / "arimax_model_results.csv", index=False)
    fitted.to_csv(output_dir / "arimax_fitted_values.csv", index=False)
    decomp.to_csv(output_dir / "stl_decomposition_values.csv", index=False)

    print("\nRunning all-stores total ARIMAX models and STL decompositions...")
    total_results, total_fitted, total_decomp = fit_models(total_model_df, target_names)
    total_results.to_csv(output_dir / "arimax_total_model_results.csv", index=False)
    total_fitted.to_csv(output_dir / "arimax_total_fitted_values.csv", index=False)
    total_decomp.to_csv(output_dir / "stl_total_decomposition_values.csv", index=False)

    charts = create_charts(results, fitted, decomp, charts_dir)
    total_charts = create_total_charts(total_results, total_fitted, total_decomp, charts_dir, store_decomp=decomp)

    print("\nBuilding Word report...")
    docx_path = build_report(results, charts, output_dir, total_results=total_results, total_charts=total_charts)
    print(f"Word report: {docx_path}")

    pdf_path = try_convert_to_pdf(docx_path, output_dir)
    if pdf_path:
        print(f"PDF report:  {pdf_path}")

    write_readme(output_dir, input_file, sheet)

    print("\nFinished.")
    print(f"Elapsed: {datetime.now() - start}")
    print(f"Open this folder: {output_dir}")


if __name__ == "__main__":
    main()
