"""
Mashgin vs. Legacy MICROS: Overall 15-Minute Transaction Flow Analysis
======================================================================

This streamlined Spyder script produces an overall analysis only. It does not
create separate store-level result charts or store-level regression estimates.
Store remains in the regression as a control so the overall Mashgin comparison
is not distorted by differences in store volume.

What the script does
--------------------
1. Connects to SQL Server with the current Windows account.
2. Reads POS.dbo.vw_Cstore_15min_Analysis.
3. Retains comparable store/weekday/15-minute clock bins under both systems.
4. Creates clean overall descriptive charts.
5. Defines historically busy clock periods from the legacy MICROS profile.
6. Fits negative-binomial regressions for:
      a. all comparable transaction-active intervals;
      b. historically busy transaction-active intervals.
7. Exports an Excel workbook, Word report, plain-text summary, and PNG charts.

Metric interpretation
---------------------
The dependent variable is the number of transaction starts recorded in a
15-minute interval. It is an operational indicator of realized checkout
throughput: a higher value means more transactions entered the checkout process
within comparable active 15-minute periods. It does not directly measure the
elapsed duration of an individual transaction or customer wait time.


"""

from __future__ import annotations

import logging
import math
import sys
import warnings
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import statsmodels.api as sm
import statsmodels.formula.api as smf
from sqlalchemy import create_engine
from sqlalchemy.engine import URL

try:
    import pyodbc
except ImportError:
    pyodbc = None

try:
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches, Pt, RGBColor

    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# =============================================================================
# USER CONFIGURATION
# =============================================================================

SERVER = "db"
DATABASE = "db"
VIEW_NAME = "dbo.vw_Cstore_15min_Analysis"
TRUST_SERVER_CERTIFICATE = True

OUTPUT_DIR = Path.home() / "Documents" / "Mashgin_15min_Overall_Analysis_Output"

ALPHA = 0.05
CREATE_WORD_REPORT = True
SAVE_ANALYTICAL_DATA_CSV = False
REQUIRE_COMPARABLE_WEEKDAY_TIME_BINS = True

# Presentation filters. These affect charts, not the regression data, except that
# the busy-period model uses historically busy bins within the supported clock profile.
MIN_SEMESTER_INTERVALS = 1000
MIN_SEMESTER_DATES = 45
MIN_TIME_BIN_ROWS = 60
MIN_TIME_BIN_DATE_COVERAGE = 0.20
TIME_PROFILE_ROLLING_BINS = 4  # Four 15-minute bins = one-hour moving average.
LEGACY_BUSY_BIN_QUANTILE = 0.75

# A campus convenience-store operating day may continue after midnight. Moving
# midnight through 3:45 AM to the right side of the chart avoids an artificial
# break between late evening and early morning.
OPERATIONAL_DAY_BREAK_HOUR = 4
OPERATIONAL_DAY_BREAK_BIN = OPERATIONAL_DAY_BREAK_HOUR * 4

# Chart settings
DPI = 220
COLOR_LEGACY = "#6C757D"
COLOR_MASHGIN = "#1E407C"
COLOR_NAVY = "#001E44"
COLOR_LIGHT = "#E8EDF2"
COLOR_PALE_BLUE = "#E8F1F8"
COLOR_PEAK = "#F3E6B3"
COLOR_NEGATIVE = "#A33A3A"
SYSTEM_ORDER = ["Legacy MICROS", "Mashgin"]

SQL_QUERY = f"""
SELECT
      [legacy_ind]
      ,[transaction_date]
      ,[rvcname]
      ,[store_name]
      ,[semester]
      ,[season]
      ,[semester_sort]
      ,[week_index]
      ,[interval_index]
      ,[interval_start_time]
      ,[interval_end_time]
      ,[interval_label]
      ,[interval_time]
      ,[transaction_initiations_15min]
      ,[transaction_initiations_30min]
      ,[transaction_initiations_60min]
FROM {VIEW_NAME};
"""


# =============================================================================
# CONNECTION AND LOGGING
# =============================================================================


def configure_logging(output_dir: Path) -> logging.Logger:
    output_dir.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("mashgin_overall_analysis")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    console = logging.StreamHandler(sys.stdout)
    console.setFormatter(formatter)
    logger.addHandler(console)

    file_handler = logging.FileHandler(
        output_dir / "analysis_log.txt", mode="w", encoding="utf-8"
    )
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    return logger


def choose_sql_driver() -> str:
    if pyodbc is None:
        raise ImportError(
            "pyodbc is not installed in the Python environment used by Spyder. "
            "Run: pip install pyodbc"
        )
    installed = list(pyodbc.drivers())
    for driver in [
        "ODBC Driver 18 for SQL Server",
        "ODBC Driver 17 for SQL Server",
    ]:
        if driver in installed:
            return driver
    raise RuntimeError(
        "Microsoft ODBC Driver 18 or 17 for SQL Server was not found. "
        f"Installed drivers: {installed}"
    )


def create_sql_engine(logger: logging.Logger):
    driver = choose_sql_driver()
    logger.info("Using SQL Server driver: %s", driver)

    connection_string = ";".join(
        [
            f"DRIVER={{{driver}}}",
            f"SERVER={SERVER}",
            f"DATABASE={DATABASE}",
            "Trusted_Connection=yes",
            "Encrypt=yes",
            f"TrustServerCertificate={'yes' if TRUST_SERVER_CERTIFICATE else 'no'}",
        ]
    ) + ";"

    url = URL.create(
        "mssql+pyodbc",
        query={"odbc_connect": connection_string},
    )
    return create_engine(url, pool_pre_ping=True, future=True)


def load_data(engine, logger: logging.Logger) -> pd.DataFrame:
    logger.info("Reading %s from %s/%s", VIEW_NAME, SERVER, DATABASE)
    with engine.connect() as connection:
        df = pd.read_sql_query(SQL_QUERY, connection)
    logger.info("Loaded %s rows", f"{len(df):,}")
    return df


# =============================================================================
# DATA PREPARATION
# =============================================================================


def prepare_data(raw: pd.DataFrame, logger: logging.Logger) -> pd.DataFrame:
    required = [
        "legacy_ind",
        "transaction_date",
        "store_name",
        "semester",
        "season",
        "semester_sort",
        "week_index",
        "interval_index",
        "interval_label",
        "transaction_initiations_15min",
        "transaction_initiations_30min",
        "transaction_initiations_60min",
    ]
    missing = [column for column in required if column not in raw.columns]
    if missing:
        raise ValueError(f"The SQL view is missing required columns: {missing}")

    df = raw.copy()
    df["transaction_date"] = pd.to_datetime(
        df["transaction_date"], errors="coerce"
    )

    numeric_columns = [
        "legacy_ind",
        "semester_sort",
        "week_index",
        "interval_index",
        "transaction_initiations_15min",
        "transaction_initiations_30min",
        "transaction_initiations_60min",
    ]
    for column in numeric_columns:
        df[column] = pd.to_numeric(df[column], errors="coerce")

    df = df.dropna(
        subset=[
            "legacy_ind",
            "transaction_date",
            "store_name",
            "season",
            "week_index",
            "interval_index",
            "transaction_initiations_15min",
        ]
    ).copy()

    df = df[df["legacy_ind"].isin([0, 1])]
    df = df[df["transaction_initiations_15min"] >= 0]

    df["legacy_ind"] = df["legacy_ind"].astype(int)
    df["interval_index"] = df["interval_index"].astype(int)
    df["mashgin_ind"] = 1 - df["legacy_ind"]
    df["system_label"] = df["legacy_ind"].map(
        {1: "Legacy MICROS", 0: "Mashgin"}
    )
    df["weekday_number"] = df["transaction_date"].dt.weekday + 1
    df["weekday_name"] = df["transaction_date"].dt.day_name()
    df["week_index_centered"] = (
        df["week_index"] - df["week_index"].mean()
    )
    df["store_date"] = (
        df["store_name"].astype(str)
        + "|"
        + df["transaction_date"].dt.strftime("%Y-%m-%d")
    )

    # Reorder the clock for presentation so 12:00–3:45 AM follows 11:45 PM.
    df["operational_interval_index"] = np.where(
        df["interval_index"] < OPERATIONAL_DAY_BREAK_BIN,
        df["interval_index"] + 96,
        df["interval_index"],
    ).astype(int)

    # Consolidate accidental duplicate rows at the analytical grain.
    grain = [
        "legacy_ind",
        "mashgin_ind",
        "system_label",
        "transaction_date",
        "store_name",
        "semester",
        "season",
        "semester_sort",
        "week_index",
        "weekday_number",
        "weekday_name",
        "interval_index",
        "operational_interval_index",
        "interval_label",
    ]
    value_columns = [
        "transaction_initiations_15min",
        "transaction_initiations_30min",
        "transaction_initiations_60min",
    ]

    before = len(df)
    df = (
        df.groupby(grain, as_index=False, observed=True)[value_columns]
        .sum()
        .reset_index(drop=True)
    )
    if len(df) != before:
        logger.info(
            "Consolidated %s duplicate analytical rows",
            f"{before - len(df):,}",
        )

    df["week_index_centered"] = (
        df["week_index"] - df["week_index"].mean()
    )
    df["store_date"] = (
        df["store_name"].astype(str)
        + "|"
        + df["transaction_date"].dt.strftime("%Y-%m-%d")
    )

    logger.info(
        "Prepared %s rows across %d stores and %d dates",
        f"{len(df):,}",
        df["store_name"].nunique(),
        df["transaction_date"].nunique(),
    )
    return df


def retain_comparable_bins(
    df: pd.DataFrame,
    logger: logging.Logger,
) -> pd.DataFrame:
    """Keep store/weekday/clock bins represented under both systems."""
    if not REQUIRE_COMPARABLE_WEEKDAY_TIME_BINS:
        return df.copy()

    keys = ["store_name", "weekday_number", "interval_index"]
    coverage = (
        df.groupby(keys, observed=True)["legacy_ind"]
        .nunique()
        .reset_index(name="system_count")
    )
    common = coverage.loc[coverage["system_count"] == 2, keys]
    result = df.merge(common, on=keys, how="inner")

    logger.info(
        "Comparable-bin restriction retained %s of %s rows (%.1f%%)",
        f"{len(result):,}",
        f"{len(df):,}",
        100 * len(result) / len(df) if len(df) else 0,
    )
    return result


# =============================================================================
# DESCRIPTIVE TABLES AND BUSY-PERIOD DEFINITION
# =============================================================================


def summarize_counts(data: pd.DataFrame, group_columns: List[str]) -> pd.DataFrame:
    return (
        data.groupby(group_columns, observed=True)
        .agg(
            interval_rows=("transaction_initiations_15min", "size"),
            distinct_dates=("transaction_date", "nunique"),
            total_initiations=("transaction_initiations_15min", "sum"),
            mean_15min=("transaction_initiations_15min", "mean"),
            median_15min=("transaction_initiations_15min", "median"),
            p75_15min=(
                "transaction_initiations_15min",
                lambda values: values.quantile(0.75),
            ),
            p90_15min=(
                "transaction_initiations_15min",
                lambda values: values.quantile(0.90),
            ),
            maximum_15min=("transaction_initiations_15min", "max"),
        )
        .reset_index()
    )


def build_descriptive_tables(
    analysis_df: pd.DataFrame,
    logger: logging.Logger,
) -> Tuple[Dict[str, pd.DataFrame], pd.DataFrame, List[int]]:
    overall = summarize_counts(analysis_df, ["system_label", "legacy_ind"])
    overall["system_sort"] = overall["system_label"].map(
        {"Legacy MICROS": 0, "Mashgin": 1}
    )
    overall = overall.sort_values("system_sort").drop(columns="system_sort")

    semester = summarize_counts(
        analysis_df,
        ["semester", "semester_sort", "system_label", "legacy_ind"],
    ).sort_values(["semester_sort", "legacy_ind"], ascending=[True, False])

    time_profile = (
        analysis_df.groupby(
            [
                "system_label",
                "legacy_ind",
                "interval_index",
                "operational_interval_index",
                "interval_label",
            ],
            observed=True,
        )
        .agg(
            interval_rows=("transaction_initiations_15min", "size"),
            distinct_dates=("transaction_date", "nunique"),
            mean_15min=("transaction_initiations_15min", "mean"),
            median_15min=("transaction_initiations_15min", "median"),
            p90_15min=(
                "transaction_initiations_15min",
                lambda values: values.quantile(0.90),
            ),
        )
        .reset_index()
    )

    supported_bins = determine_supported_clock_bins(time_profile)
    if not supported_bins:
        raise RuntimeError(
            "No comparable, adequately represented time-of-day bins were found. "
            "Reduce MIN_TIME_BIN_ROWS or MIN_TIME_BIN_DATE_COVERAGE."
        )

    supported_original_bins = sorted(
        time_profile.loc[
            time_profile["operational_interval_index"].isin(supported_bins),
            "interval_index",
        ].unique()
    )

    # Define busy clock periods from the legacy profile only. This avoids using
    # Mashgin outcomes to decide what counts as historically busy.
    legacy_profile = time_profile.loc[
        (time_profile["legacy_ind"] == 1)
        & (time_profile["operational_interval_index"].isin(supported_bins))
    ].copy()
    busy_threshold = float(legacy_profile["mean_15min"].quantile(LEGACY_BUSY_BIN_QUANTILE))
    busy_bins = sorted(
        legacy_profile.loc[
            legacy_profile["mean_15min"] >= busy_threshold,
            "interval_index",
        ].astype(int).unique()
    )

    analysis_with_period = analysis_df.copy()
    analysis_with_period["demand_period"] = np.where(
        analysis_with_period["interval_index"].isin(busy_bins),
        "Historically busy intervals",
        "Other comparable intervals",
    )

    period_summary = summarize_counts(
        analysis_with_period,
        ["demand_period", "system_label", "legacy_ind"],
    )
    period_summary["period_sort"] = period_summary["demand_period"].map(
        {
            "Other comparable intervals": 0,
            "Historically busy intervals": 1,
        }
    )
    period_summary["system_sort"] = period_summary["system_label"].map(
        {"Legacy MICROS": 0, "Mashgin": 1}
    )
    period_summary = period_summary.sort_values(
        ["period_sort", "system_sort"]
    ).drop(columns=["period_sort", "system_sort"])

    # Daily concentration table: how much of each store-day's volume occurred
    # in historically busy bins.
    daily = (
        analysis_with_period.groupby(
            ["system_label", "legacy_ind", "store_date", "transaction_date"],
            observed=True,
        )
        .agg(
            daily_total=("transaction_initiations_15min", "sum"),
            busy_total=(
                "transaction_initiations_15min",
                lambda values: values.sum(),
            ),
        )
        .reset_index()
    )
    # Recalculate busy total correctly using a separate grouped source.
    busy_daily = (
        analysis_with_period.loc[
            analysis_with_period["demand_period"] == "Historically busy intervals"
        ]
        .groupby(["system_label", "legacy_ind", "store_date"], observed=True)[
            "transaction_initiations_15min"
        ]
        .sum()
        .rename("busy_total")
        .reset_index()
    )
    daily = daily.drop(columns="busy_total").merge(
        busy_daily,
        on=["system_label", "legacy_ind", "store_date"],
        how="left",
    )
    daily["busy_total"] = daily["busy_total"].fillna(0)
    daily["busy_share"] = daily["busy_total"] / daily["daily_total"].replace(0, np.nan)

    daily_summary = (
        daily.groupby(["system_label", "legacy_ind"], observed=True)
        .agg(
            store_days=("store_date", "nunique"),
            mean_daily_total=("daily_total", "mean"),
            median_daily_total=("daily_total", "median"),
            mean_busy_share=("busy_share", "mean"),
        )
        .reset_index()
    )

    logger.info(
        "Historically busy period definition: top %.0f%% of supported legacy clock bins; "
        "%d bins selected; legacy mean threshold %.2f starts per interval",
        (1 - LEGACY_BUSY_BIN_QUANTILE) * 100,
        len(busy_bins),
        busy_threshold,
    )

    tables = {
        "overall_summary": overall,
        "semester_summary": semester,
        "time_profile": time_profile,
        "period_summary": period_summary,
        "daily_summary": daily_summary,
    }
    return tables, analysis_with_period, busy_bins


def determine_supported_clock_bins(time_profile: pd.DataFrame) -> List[int]:
    """Return the longest consecutive operational-time sequence supported by both systems."""
    profile = time_profile.copy()
    max_dates = profile.groupby("system_label")["distinct_dates"].max().to_dict()
    profile["supported"] = profile.apply(
        lambda row: (
            row["interval_rows"] >= MIN_TIME_BIN_ROWS
            and row["distinct_dates"]
            >= max_dates.get(row["system_label"], 0) * MIN_TIME_BIN_DATE_COVERAGE
        ),
        axis=1,
    )

    pivot = profile.pivot_table(
        index="operational_interval_index",
        columns="system_label",
        values="supported",
        aggfunc="max",
        fill_value=False,
    )
    for system in SYSTEM_ORDER:
        if system not in pivot.columns:
            return []
    common = sorted(
        int(index)
        for index, row in pivot.iterrows()
        if bool(row["Legacy MICROS"]) and bool(row["Mashgin"])
    )
    if not common:
        return []

    sequences: List[List[int]] = []
    current = [common[0]]
    for value in common[1:]:
        if value == current[-1] + 1:
            current.append(value)
        else:
            sequences.append(current)
            current = [value]
    sequences.append(current)
    return max(sequences, key=len)


# =============================================================================
# NEGATIVE-BINOMIAL REGRESSION
# =============================================================================


def estimate_moment_alpha(data: pd.DataFrame) -> float:
    mean_value = float(data["transaction_initiations_15min"].mean())
    variance = float(data["transaction_initiations_15min"].var(ddof=1))
    if mean_value <= 0:
        return 1.0
    alpha = max((variance - mean_value) / (mean_value**2), 1e-8)
    return float(alpha)


def extract_effect(
    result,
    scope: str,
    formula: str,
    data: pd.DataFrame,
    model_type: str,
    notes: str = "",
) -> Dict[str, object]:
    parameter = "mashgin_ind"
    coefficient = float(result.params[parameter])
    standard_error = float(result.bse[parameter])
    p_value = float(result.pvalues[parameter])
    confidence = result.conf_int().loc[parameter]
    lower_log = float(confidence.iloc[0])
    upper_log = float(confidence.iloc[1])

    irr = math.exp(coefficient)
    lower_irr = math.exp(lower_log)
    upper_irr = math.exp(upper_log)

    legacy_mean = float(
        data.loc[data["mashgin_ind"] == 0, "transaction_initiations_15min"].mean()
    )
    mashgin_mean = float(
        data.loc[data["mashgin_ind"] == 1, "transaction_initiations_15min"].mean()
    )

    adjusted_percent = (irr - 1) * 100
    illustrative_15 = legacy_mean * adjusted_percent / 100

    return {
        "scope": scope,
        "model_type": model_type,
        "model_formula": formula,
        "interval_rows": int(len(data)),
        "distinct_store_dates": int(data["store_date"].nunique()),
        "legacy_mean": legacy_mean,
        "mashgin_mean": mashgin_mean,
        "observed_percent_difference": (
            (mashgin_mean - legacy_mean) / legacy_mean * 100
            if legacy_mean != 0
            else np.nan
        ),
        "coefficient_log_rate": coefficient,
        "standard_error": standard_error,
        "incidence_rate_ratio": irr,
        "adjusted_percent_difference": adjusted_percent,
        "ci_lower_percent": (lower_irr - 1) * 100,
        "ci_upper_percent": (upper_irr - 1) * 100,
        "p_value": p_value,
        "significant_at_0_05": p_value < ALPHA,
        "illustrative_difference_per_15min": illustrative_15,
        "illustrative_difference_per_hour": illustrative_15 * 4,
        "notes": notes,
    }


def fit_negative_binomial(
    data: pd.DataFrame,
    scope: str,
    logger: logging.Logger,
) -> Dict[str, object]:
    """Fit one overall NB2 model with store and time controls."""
    model_data = data.copy()
    if model_data["mashgin_ind"].nunique() < 2:
        raise ValueError(f"{scope} requires observations under both systems.")

    formula = (
        "transaction_initiations_15min ~ mashgin_ind "
        "+ C(store_name) "
        "+ C(weekday_number) "
        "+ C(interval_index) "
        "+ C(season) "
        "+ week_index_centered"
    )

    logger.info("Fitting negative-binomial regression: %s", scope)
    errors: List[str] = []

    for method, max_iterations in [
        ("newton", 100),
        ("bfgs", 250),
        ("lbfgs", 250),
    ]:
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                model = smf.negativebinomial(
                    formula=formula,
                    data=model_data,
                    loglike_method="nb2",
                    check_rank=False,
                )
                result = model.fit(
                    method=method,
                    maxiter=max_iterations,
                    disp=False,
                    cov_type="cluster",
                    cov_kwds={"groups": model_data["store_date"]},
                )
            converged = bool(result.mle_retvals.get("converged", True))
            if not converged:
                errors.append(f"{method}: model did not report convergence")
                continue
            return extract_effect(
                result,
                scope,
                formula,
                model_data,
                f"Negative binomial NB2 ({method}; store-date clustered SE)",
            )
        except Exception as exc:
            errors.append(f"{method}: {type(exc).__name__}: {exc}")

    # Robust fallback if maximum-likelihood estimation does not converge.
    try:
        alpha = estimate_moment_alpha(model_data)
        logger.warning(
            "NB2 maximum-likelihood model failed for %s; using GEE fallback with alpha %.6f",
            scope,
            alpha,
        )
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            gee = smf.gee(
                formula=formula,
                groups="store_date",
                data=model_data,
                family=sm.families.NegativeBinomial(alpha=alpha),
                cov_struct=sm.cov_struct.Exchangeable(),
            )
            result = gee.fit(maxiter=100)
        return extract_effect(
            result,
            scope,
            formula,
            model_data,
            "Negative binomial GEE fallback (robust SE)",
            notes="; ".join(errors),
        )
    except Exception as exc:
        errors.append(f"GEE: {type(exc).__name__}: {exc}")
        raise RuntimeError(
            f"All regression attempts failed for {scope}: {'; '.join(errors)}"
        ) from exc


def run_regressions(
    analysis_with_period: pd.DataFrame,
    logger: logging.Logger,
) -> pd.DataFrame:
    results = [
        fit_negative_binomial(
            analysis_with_period,
            "All comparable active intervals",
            logger,
        )
    ]

    busy = analysis_with_period.loc[
        analysis_with_period["demand_period"] == "Historically busy intervals"
    ].copy()
    results.append(
        fit_negative_binomial(
            busy,
            "Historically busy intervals",
            logger,
        )
    )

    regression = pd.DataFrame(results)
    regression["scope_sort"] = regression["scope"].map(
        {
            "All comparable active intervals": 0,
            "Historically busy intervals": 1,
        }
    )
    return regression.sort_values("scope_sort").drop(columns="scope_sort")


# =============================================================================
# CHARTING
# =============================================================================


def apply_chart_defaults() -> None:
    plt.rcParams.update(
        {
            "figure.facecolor": "white",
            "axes.facecolor": "white",
            "axes.edgecolor": "#8A949E",
            "axes.labelcolor": COLOR_NAVY,
            "axes.titleweight": "bold",
            "font.family": "DejaVu Sans",
            "font.size": 10.5,
            "xtick.color": "#333333",
            "ytick.color": "#333333",
            "legend.frameon": False,
            "axes.spines.top": False,
            "axes.spines.right": False,
        }
    )


def add_header(fig, title: str, subtitle: str = "") -> None:
    fig.text(0.07, 0.965, title, fontsize=17, fontweight="bold", color=COLOR_NAVY)
    if subtitle:
        fig.text(0.07, 0.925, subtitle, fontsize=10.5, color="#4F5B66")


def save_figure(fig, path: Path, bottom: float = 0.09) -> Path:
    fig.tight_layout(rect=[0.05, bottom, 0.98, 0.89])
    fig.savefig(path, dpi=DPI, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return path


def format_clock_bin(operational_bin: int) -> str:
    original = operational_bin % 96
    total_minutes = original * 15
    hour = total_minutes // 60
    minute = total_minutes % 60
    suffix = "AM" if hour < 12 else "PM"
    display_hour = hour % 12
    if display_hour == 0:
        display_hour = 12
    return f"{display_hour}:{minute:02d} {suffix}"


def plot_overall_summary(summary: pd.DataFrame, output_dir: Path) -> Path:
    lookup = summary.set_index("system_label")
    metrics = [
        ("Mean", "mean_15min"),
        ("Median", "median_15min"),
        ("90th percentile", "p90_15min"),
    ]
    y = np.arange(len(metrics))[::-1]

    legacy = np.array([lookup.loc["Legacy MICROS", column] for _, column in metrics])
    mashgin = np.array([lookup.loc["Mashgin", column] for _, column in metrics])

    fig, ax = plt.subplots(figsize=(10.5, 5.8))
    for position, legacy_value, mashgin_value in zip(y, legacy, mashgin):
        ax.plot(
            [legacy_value, mashgin_value],
            [position, position],
            color="#BCC4CB",
            linewidth=4,
            solid_capstyle="round",
            zorder=1,
        )
        ax.scatter(legacy_value, position, s=110, color=COLOR_LEGACY, zorder=3)
        ax.scatter(mashgin_value, position, s=110, color=COLOR_MASHGIN, zorder=3)
        ax.text(
            legacy_value - 0.8,
            position + 0.18,
            f"{legacy_value:.1f}",
            ha="right",
            color=COLOR_LEGACY,
            fontweight="bold",
        )
        ax.text(
            mashgin_value + 0.8,
            position + 0.18,
            f"{mashgin_value:.1f}",
            ha="left",
            color=COLOR_MASHGIN,
            fontweight="bold",
        )

    ax.set_yticks(y, [label for label, _ in metrics])
    ax.set_xlabel("Transactions initiated per transaction-active 15-minute interval")
    ax.set_xlim(left=0)
    ax.grid(axis="x", alpha=0.16)
    ax.scatter([], [], color=COLOR_LEGACY, s=80, label="Legacy MICROS")
    ax.scatter([], [], color=COLOR_MASHGIN, s=80, label="Mashgin")
    ax.legend(loc="upper right")

    add_header(
        fig,
        "Overall Transaction Flow Was Higher During the Mashgin Period",
        "Mean, median, and recurring high-demand (90th-percentile) active intervals",
    )
    return save_figure(fig, output_dir / "01_overall_transaction_flow.png")


def plot_time_profile(
    time_profile: pd.DataFrame,
    busy_bins: List[int],
    output_dir: Path,
) -> Path:
    supported = determine_supported_clock_bins(time_profile)
    data = time_profile.loc[
        time_profile["operational_interval_index"].isin(supported)
    ].copy()
    data = data.sort_values(["system_label", "operational_interval_index"])

    fig, ax = plt.subplots(figsize=(13, 6.7))

    # Shade historically busy clock bins.
    busy_operational = sorted(
        data.loc[data["interval_index"].isin(busy_bins), "operational_interval_index"].unique()
    )
    if busy_operational:
        sequences: List[List[int]] = []
        current = [busy_operational[0]]
        for value in busy_operational[1:]:
            if value == current[-1] + 1:
                current.append(value)
            else:
                sequences.append(current)
                current = [value]
        sequences.append(current)
        for sequence in sequences:
            ax.axvspan(
                sequence[0] - 0.5,
                sequence[-1] + 0.5,
                color=COLOR_PEAK,
                alpha=0.32,
                linewidth=0,
            )

    for system, color in [
        ("Legacy MICROS", COLOR_LEGACY),
        ("Mashgin", COLOR_MASHGIN),
    ]:
        subset = data.loc[data["system_label"] == system].copy()
        subset["rolling_mean"] = subset["mean_15min"].rolling(
            TIME_PROFILE_ROLLING_BINS,
            center=True,
            min_periods=1,
        ).mean()
        ax.plot(
            subset["operational_interval_index"],
            subset["mean_15min"],
            color=color,
            linewidth=1,
            alpha=0.20,
        )
        ax.plot(
            subset["operational_interval_index"],
            subset["rolling_mean"],
            color=color,
            linewidth=2.8,
            label=system,
        )

    minimum = int(data["operational_interval_index"].min())
    maximum = int(data["operational_interval_index"].max())
    ticks = list(range(minimum, maximum + 1, 8))  # Every two hours.
    ax.set_xticks(ticks, [format_clock_bin(value) for value in ticks], rotation=45, ha="right")
    ax.set_ylabel("Average transaction starts per active 15-minute interval")
    ax.set_xlabel("Time of day")
    ax.set_ylim(bottom=0)
    ax.grid(axis="y", alpha=0.16)
    ax.legend(loc="upper left")

    add_header(
        fig,
        "Transaction Flow Throughout the Operating Day",
        "Bold lines are one-hour moving averages; shaded areas are historically busy legacy clock periods",
    )
    return save_figure(fig, output_dir / "02_time_of_day_profile.png", bottom=0.13)


def plot_semester_trend(semester: pd.DataFrame, output_dir: Path) -> Path:
    data = semester.copy()
    keep = (
        (data["interval_rows"] >= MIN_SEMESTER_INTERVALS)
        & (data["distinct_dates"] >= MIN_SEMESTER_DATES)
    )
    omitted = data.loc[~keep].copy()
    data = data.loc[keep].sort_values("semester_sort")

    semester_order = (
        data[["semester", "semester_sort"]]
        .drop_duplicates()
        .sort_values("semester_sort")
    )
    positions = {
        semester_name: index
        for index, semester_name in enumerate(semester_order["semester"])
    }

    fig, ax = plt.subplots(figsize=(12, 6.2))
    for system, color in [
        ("Legacy MICROS", COLOR_LEGACY),
        ("Mashgin", COLOR_MASHGIN),
    ]:
        subset = data.loc[data["system_label"] == system].copy()
        subset["x"] = subset["semester"].map(positions)
        ax.plot(
            subset["x"],
            subset["mean_15min"],
            marker="o",
            markersize=7,
            linewidth=2.6,
            color=color,
            label=system,
        )

    ax.set_xticks(
        list(positions.values()),
        list(positions.keys()),
        rotation=40,
        ha="right",
    )
    ax.set_ylabel("Average transaction starts per active 15-minute interval")
    ax.set_ylim(bottom=0)
    ax.grid(axis="y", alpha=0.16)
    ax.legend(loc="upper left")

    add_header(
        fig,
        "Mean 15-Minute Transaction Flow by Semester",
        "Sparse transition and ramp-down cells are omitted rather than connected to the trend",
    )

    bottom = 0.12
    if not omitted.empty:
        note = "Omitted sparse cells: " + "; ".join(
            f"{row['semester']} {row['system_label']} "
            f"({int(row['interval_rows'])} intervals, {int(row['distinct_dates'])} dates)"
            for _, row in omitted.iterrows()
        )
        fig.text(0.07, 0.02, note, fontsize=8.5, color=COLOR_LEGACY)
    return save_figure(fig, output_dir / "03_semester_trend_filtered.png", bottom=bottom)


def plot_busy_period_comparison(period_summary: pd.DataFrame, output_dir: Path) -> Path:
    periods = ["Other comparable intervals", "Historically busy intervals"]
    x = np.arange(len(periods))
    width = 0.32

    fig, ax = plt.subplots(figsize=(10.8, 6.1))
    for offset, system, color in [
        (-width / 2, "Legacy MICROS", COLOR_LEGACY),
        (width / 2, "Mashgin", COLOR_MASHGIN),
    ]:
        subset = period_summary.loc[period_summary["system_label"] == system].set_index(
            "demand_period"
        )
        values = [float(subset.loc[period, "mean_15min"]) for period in periods]
        bars = ax.bar(x + offset, values, width=width, color=color, label=system)
        for bar, value in zip(bars, values):
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                value + 0.7,
                f"{value:.1f}",
                ha="center",
                va="bottom",
                fontweight="bold",
                color=color,
            )

    ax.set_xticks(x, ["Other comparable\nclock periods", "Historically busy\nclock periods"])
    ax.set_ylabel("Average transaction starts per active 15-minute interval")
    ax.set_ylim(bottom=0)
    ax.grid(axis="y", alpha=0.16)
    ax.legend(loc="upper left")

    add_header(
        fig,
        "Mashgin Handled More Transaction Starts During Historically Busy Periods",
        "Busy periods are the top quartile of supported legacy clock bins, defined before comparing Mashgin",
    )
    return save_figure(fig, output_dir / "04_busy_period_comparison.png")


def format_p_value(value: float) -> str:
    if value < 0.001:
        return "p < .001"
    return f"p = {value:.3f}"


def plot_observed_vs_adjusted(regression: pd.DataFrame, output_dir: Path) -> Path:
    overall = regression.loc[
        regression["scope"] == "All comparable active intervals"
    ].iloc[0]
    values = [
        float(overall["observed_percent_difference"]),
        float(overall["adjusted_percent_difference"]),
    ]
    labels = ["Observed mean\ndifference", "Regression-adjusted\ndifference"]
    colors = [COLOR_LIGHT, COLOR_MASHGIN]

    fig, ax = plt.subplots(figsize=(8.8, 5.8))
    bars = ax.bar(labels, values, color=colors, width=0.55)
    for index, (bar, value) in enumerate(zip(bars, values)):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            value + 0.8,
            f"{value:+.1f}%",
            ha="center",
            fontweight="bold",
            color=COLOR_NAVY if index == 0 else COLOR_MASHGIN,
            fontsize=12,
        )
    ax.axhline(0, color=COLOR_NAVY, linewidth=1)
    ax.set_ylabel("Mashgin difference in transaction initiation rate")
    ax.grid(axis="y", alpha=0.16)

    add_header(
        fig,
        "The Raw Difference Narrows After Comparable Conditions Are Aligned",
        "The regression holds constant store, weekday, clock interval, season, and the weekly trend",
    )
    return save_figure(fig, output_dir / "05_observed_vs_adjusted.png")


def plot_regression_effect(regression: pd.DataFrame, output_dir: Path) -> Path:
    data = regression.copy()
    order = [
        "All comparable active intervals",
        "Historically busy intervals",
    ]
    data["scope"] = pd.Categorical(data["scope"], categories=order, ordered=True)
    data = data.sort_values("scope")
    y = np.arange(len(data))[::-1]

    lower = float(data["ci_lower_percent"].min()) - 3
    upper = float(data["ci_upper_percent"].max()) + 10

    fig, ax = plt.subplots(figsize=(10.6, 5.7))
    ax.axvspan(lower, 0, color="#F8ECEC", alpha=0.7)
    ax.axvspan(0, upper, color=COLOR_PALE_BLUE, alpha=0.75)

    for position, (_, row) in zip(y, data.iterrows()):
        estimate = float(row["adjusted_percent_difference"])
        ci_lower = float(row["ci_lower_percent"])
        ci_upper = float(row["ci_upper_percent"])
        color = COLOR_MASHGIN if estimate >= 0 else COLOR_NEGATIVE
        ax.errorbar(
            estimate,
            position,
            xerr=[[estimate - ci_lower], [ci_upper - estimate]],
            fmt="o",
            markersize=9,
            capsize=5,
            linewidth=2.2,
            color=color,
            ecolor=color,
        )
        ax.text(
            ci_upper + 0.8,
            position,
            f"{estimate:+.1f}% ({format_p_value(float(row['p_value']))})",
            va="center",
            ha="left",
            color=color,
            fontweight="bold",
        )

    ax.axvline(0, color=COLOR_NAVY, linestyle="--", linewidth=1.2)
    ax.set_yticks(y, ["All comparable active intervals", "Historically busy intervals"])
    ax.set_xlabel("Adjusted Mashgin difference in transaction initiation rate (%)")
    ax.set_xlim(lower, upper)
    ax.grid(axis="x", alpha=0.16)

    add_header(
        fig,
        "Negative-Binomial Regression: Adjusted Mashgin Transaction-Flow Effect",
        "Points are adjusted estimates; lines are 95% confidence intervals",
    )
    fig.text(0.07, 0.025, "← Lower realized flow", color=COLOR_NEGATIVE, fontsize=9)
    fig.text(0.76, 0.025, "Higher realized flow →", color=COLOR_MASHGIN, fontsize=9)
    return save_figure(fig, output_dir / "06_regression_effect.png", bottom=0.10)


def create_charts(
    tables: Dict[str, pd.DataFrame],
    regression: pd.DataFrame,
    busy_bins: List[int],
    output_dir: Path,
    logger: logging.Logger,
) -> Dict[str, Path]:
    apply_chart_defaults()
    paths = {
        "overall": plot_overall_summary(tables["overall_summary"], output_dir),
        "time_profile": plot_time_profile(tables["time_profile"], busy_bins, output_dir),
        "semester": plot_semester_trend(tables["semester_summary"], output_dir),
        "busy": plot_busy_period_comparison(tables["period_summary"], output_dir),
        "observed_adjusted": plot_observed_vs_adjusted(regression, output_dir),
        "regression": plot_regression_effect(regression, output_dir),
    }
    logger.info("Created %d chart files", len(paths))
    return paths


# =============================================================================
# BUSINESS INTERPRETATION
# =============================================================================


def interpretation_text(
    tables: Dict[str, pd.DataFrame],
    regression: pd.DataFrame,
) -> Dict[str, str]:
    overall = tables["overall_summary"].set_index("system_label")
    legacy_mean = float(overall.loc["Legacy MICROS", "mean_15min"])
    mashgin_mean = float(overall.loc["Mashgin", "mean_15min"])
    observed = (mashgin_mean - legacy_mean) / legacy_mean * 100

    model = regression.loc[
        regression["scope"] == "All comparable active intervals"
    ].iloc[0]
    busy_model = regression.loc[
        regression["scope"] == "Historically busy intervals"
    ].iloc[0]

    adjusted = float(model["adjusted_percent_difference"])
    ci_low = float(model["ci_lower_percent"])
    ci_high = float(model["ci_upper_percent"])
    per_15 = float(model["illustrative_difference_per_15min"])
    per_hour = float(model["illustrative_difference_per_hour"])

    regression_explanation = (
        "The negative-binomial regression treats each 15-minute transaction count as a "
        "count outcome rather than as a normally distributed continuous measure. It compares "
        "Mashgin with legacy MICROS while holding constant the store, weekday, exact 15-minute "
        "clock interval, spring-versus-fall season, and the underlying weekly trend. Standard "
        "errors are clustered by store-date so that intervals from the same operating day are "
        "not treated as fully independent. The Mashgin coefficient is exponentiated into an "
        "incidence-rate ratio and then translated into the adjusted percentage difference shown "
        "in the report."
    )

    sales_explanation = (
        "A higher 15-minute initiation rate does not automatically produce higher weekly sales. "
        "Sales depend on total customer demand and the value of each purchase, not only on how "
        "many transactions can enter checkout during a busy interval. The same weekly customer "
        "volume can be processed in a more concentrated period, with less idle checkout time or "
        "greater simultaneous capacity, without creating additional visits or increasing the "
        "average check. In the broader project, adjusted weekly sales, total transactions, and "
        "basket size did not show a broad system-period increase. Taken together, the findings "
        "support an operational-capacity interpretation: Mashgin appears to have handled more "
        "transaction starts during comparable active periods, especially when demand was "
        "concentrated, but it did not by itself create new demand or increase the value of each visit."
    )

    executive = (
        f"Legacy MICROS averaged {legacy_mean:.2f} transaction starts per transaction-active "
        f"15-minute interval, compared with {mashgin_mean:.2f} under Mashgin, an observed "
        f"difference of {observed:+.1f}%. After aligning store, weekday, exact clock interval, "
        f"season, and the weekly trend, Mashgin was associated with an adjusted {adjusted:+.1f}% "
        f"difference (95% CI {ci_low:+.1f}% to {ci_high:+.1f}%; "
        f"{format_p_value(float(model['p_value']))}). At the legacy mean, that effect is "
        f"approximately {per_15:+.1f} additional transaction starts per 15 minutes, or "
        f"{per_hour:+.1f} per hour if sustained. The historically busy-period model estimated "
        f"a {float(busy_model['adjusted_percent_difference']):+.1f}% Mashgin difference."
    )

    business_definition = (
        "Transaction initiation rate is used as an operational indicator of realized checkout "
        "throughput. A higher value means that more transactions entered the checkout process "
        "during comparable active 15-minute periods. It does not prove that each individual "
        "transaction or customer wait was shorter."
    )

    return {
        "executive": executive,
        "business_definition": business_definition,
        "regression_explanation": regression_explanation,
        "sales_explanation": sales_explanation,
    }


# =============================================================================
# EXPORTS
# =============================================================================


def autosize_sheet(worksheet, dataframe: pd.DataFrame) -> None:
    fill = PatternFill("solid", fgColor="1E407C")
    font = Font(color="FFFFFF", bold=True)
    for cell in worksheet[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for column_number, column_name in enumerate(dataframe.columns, start=1):
        values = dataframe[column_name].astype(str).replace("nan", "")
        width = max([len(str(column_name))] + [len(value) for value in values.head(2000)])
        worksheet.column_dimensions[get_column_letter(column_number)].width = min(width + 2, 42)


def export_excel(
    tables: Dict[str, pd.DataFrame],
    regression: pd.DataFrame,
    busy_bins: List[int],
    output_dir: Path,
) -> Path:
    path = output_dir / "Mashgin_15min_Overall_Analysis_Results.xlsx"
    busy_table = pd.DataFrame(
        {
            "interval_index": busy_bins,
            "interval_time": [format_clock_bin(value) for value in busy_bins],
        }
    )
    exports = {
        "Overall_Summary": tables["overall_summary"],
        "Semester_Summary": tables["semester_summary"],
        "Busy_Period_Summary": tables["period_summary"],
        "Daily_Concentration": tables["daily_summary"],
        "Regression_Results": regression,
        "Busy_Clock_Bins": busy_table,
    }
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for sheet_name, dataframe in exports.items():
            safe = dataframe.copy()
            safe.to_excel(writer, sheet_name=sheet_name[:31], index=False)
            autosize_sheet(writer.sheets[sheet_name[:31]], safe)
    return path


def add_word_heading(document: Document, text: str, level: int = 1) -> None:
    heading = document.add_heading(text, level=level)
    for run in heading.runs:
        run.font.color.rgb = RGBColor(30, 64, 124)


def add_dataframe_table(document: Document, dataframe: pd.DataFrame, decimals: int = 2) -> None:
    table = document.add_table(rows=1, cols=len(dataframe.columns))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for index, column in enumerate(dataframe.columns):
        table.rows[0].cells[index].text = str(column)
    for _, row in dataframe.iterrows():
        cells = table.add_row().cells
        for index, value in enumerate(row):
            if isinstance(value, (float, np.floating)) and pd.notna(value):
                cells[index].text = f"{value:.{decimals}f}"
            else:
                cells[index].text = "" if pd.isna(value) else str(value)


def export_word_report(
    tables: Dict[str, pd.DataFrame],
    regression: pd.DataFrame,
    chart_paths: Dict[str, Path],
    interpretation: Dict[str, str],
    output_dir: Path,
) -> Optional[Path]:
    if not CREATE_WORD_REPORT or not DOCX_AVAILABLE:
        return None

    document = Document()
    styles = document.styles
    styles["Normal"].font.name = "Aptos"
    styles["Normal"].font.size = Pt(10.5)

    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("Legacy MICROS and Mashgin\nOverall 15-Minute Transaction Flow Analysis")
    run.bold = True
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor(0, 30, 68)

    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.add_run(f"Generated {datetime.now():%B %d, %Y}")

    add_word_heading(document, "Executive Summary")
    document.add_paragraph(interpretation["executive"])
    document.add_paragraph(interpretation["business_definition"])

    add_word_heading(document, "Overall Descriptive Results")
    document.add_picture(str(chart_paths["overall"]), width=Inches(6.8))
    overall_table = tables["overall_summary"][
        [
            "system_label",
            "interval_rows",
            "distinct_dates",
            "mean_15min",
            "median_15min",
            "p90_15min",
            "maximum_15min",
        ]
    ].rename(
        columns={
            "system_label": "System",
            "interval_rows": "Intervals",
            "distinct_dates": "Dates",
            "mean_15min": "Mean",
            "median_15min": "Median",
            "p90_15min": "P90",
            "maximum_15min": "Maximum",
        }
    )
    add_dataframe_table(document, overall_table)

    add_word_heading(document, "When the Difference Occurred")
    document.add_picture(str(chart_paths["time_profile"]), width=Inches(6.8))
    document.add_paragraph(
        "The shaded clock periods are defined from the legacy MICROS pattern before Mashgin is compared. "
        "This makes the busy-period comparison operationally meaningful without allowing the Mashgin result "
        "to determine which times are labeled busy."
    )
    document.add_picture(str(chart_paths["busy"]), width=Inches(6.8))
    document.add_picture(str(chart_paths["semester"]), width=Inches(6.8))

    add_word_heading(document, "What the Regression Is Doing")
    document.add_paragraph(interpretation["regression_explanation"])
    document.add_picture(str(chart_paths["observed_adjusted"]), width=Inches(6.3))
    document.add_picture(str(chart_paths["regression"]), width=Inches(6.8))

    regression_table = regression[
        [
            "scope",
            "observed_percent_difference",
            "adjusted_percent_difference",
            "ci_lower_percent",
            "ci_upper_percent",
            "p_value",
            "incidence_rate_ratio",
        ]
    ].rename(
        columns={
            "scope": "Analysis",
            "observed_percent_difference": "Observed %",
            "adjusted_percent_difference": "Adjusted %",
            "ci_lower_percent": "CI lower %",
            "ci_upper_percent": "CI upper %",
            "p_value": "P value",
            "incidence_rate_ratio": "Rate ratio",
        }
    )
    add_dataframe_table(document, regression_table, decimals=3)

    add_word_heading(document, "Why Greater Transaction Flow Does Not Necessarily Increase Sales")
    document.add_paragraph(interpretation["sales_explanation"])

    add_word_heading(document, "Interpretation Boundary")
    document.add_paragraph(
        "The source contains transaction-active 15-minute intervals. The findings therefore describe "
        "realized transaction flow when checkout activity occurred. Adding confirmed open intervals with "
        "zero transactions would support a full operating-period rate. The measure also does not directly "
        "observe individual transaction duration, queue length, or customer wait time."
    )

    path = output_dir / "Mashgin_15min_Overall_Analysis_Report.docx"
    document.save(path)
    return path


def export_text_summary(
    interpretation: Dict[str, str],
    output_dir: Path,
) -> Path:
    path = output_dir / "Mashgin_15min_Overall_Analysis_Summary.txt"
    text = (
        "MASHGIN 15-MINUTE TRANSACTION FLOW ANALYSIS\n"
        "============================================\n\n"
        "EXECUTIVE RESULT\n"
        f"{interpretation['executive']}\n\n"
        "BUSINESS MEANING\n"
        f"{interpretation['business_definition']}\n\n"
        "WHAT THE REGRESSION DOES\n"
        f"{interpretation['regression_explanation']}\n\n"
        "WHY SALES MAY NOT INCREASE\n"
        f"{interpretation['sales_explanation']}\n"
    )
    path.write_text(text, encoding="utf-8")
    return path


# =============================================================================
# MAIN
# =============================================================================


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    logger = configure_logging(OUTPUT_DIR)
    logger.info("Starting overall Mashgin 15-minute transaction-flow analysis")

    engine = None
    try:
        engine = create_sql_engine(logger)
        raw = load_data(engine, logger)
        prepared = prepare_data(raw, logger)
        analysis = retain_comparable_bins(prepared, logger)

        tables, analysis_with_period, busy_bins = build_descriptive_tables(
            analysis, logger
        )
        regression = run_regressions(analysis_with_period, logger)
        chart_paths = create_charts(
            tables, regression, busy_bins, OUTPUT_DIR, logger
        )
        interpretation = interpretation_text(tables, regression)

        excel_path = export_excel(
            tables, regression, busy_bins, OUTPUT_DIR
        )
        text_path = export_text_summary(interpretation, OUTPUT_DIR)
        word_path = export_word_report(
            tables,
            regression,
            chart_paths,
            interpretation,
            OUTPUT_DIR,
        )

        if SAVE_ANALYTICAL_DATA_CSV:
            analysis_with_period.to_csv(
                OUTPUT_DIR / "comparable_15min_analytical_data.csv",
                index=False,
            )

        logger.info("Analysis completed successfully")
        logger.info("Excel results: %s", excel_path)
        logger.info("Text summary: %s", text_path)
        if word_path:
            logger.info("Word report: %s", word_path)
        logger.info("Output folder: %s", OUTPUT_DIR)

        print("\n" + interpretation["executive"])
        print("\n" + interpretation["sales_explanation"])
        print(f"\nOutputs written to: {OUTPUT_DIR}")

    except Exception:
        logger.exception("Analysis failed")
        raise
    finally:
        if engine is not None:
            engine.dispose()


if __name__ == "__main__":
    main()
