#!/usr/bin/env python3
"""
Algo performance report generator.

Reads an Excel (.xlsx/.xls) or CSV file, prompts for algo selection, and produces:

  1. Processed data.xlsx  (in --output-dir)
     Sheet "data"           : Filtered raw rows (selected algos, required columns only).
                              allocation is scaled ×100 in this sheet.
     Sheet "processed data" : Daily aggregated per algo.
                              Columns: date, algo, sum_allocation, sum_mtm,
                                       PNL%, Absolute PNL, Base Capital

  2. {report_name}_interactive.html  (dynamic, internal use)
  3. {report_name}_client.html       (static snapshot, shareable)

Return formula:
    processed_allocation = raw_allocation × 100
    For each algo on each date:
        sum_allocation  = Σ processed_allocation  (across all users of that algo)
        sum_mtm         = Σ mtm_all               (across all users)
        daily_return    = sum_mtm / sum_allocation  (fractional)
        PNL%            = daily_return × 100
        Base Capital    = 10,000,000  (fixed, 1 crore)
        Absolute PNL    = daily_return × Base Capital

Data filters applied before any aggregation:
    1. Algo-19 broker exclusion: rows where algo == "19" AND broker is
       MasterTrust_Noren or mastertrust_dealer are dropped.
    2. 0DTE filter: for algos 1, 7, 15 only rows where dte == "0DTE"
       (case-insensitive) are retained; all other DTE rows are excluded.
    3. Date exclusion: all rows for 18 Apr 2024 are dropped entirely.

Usage:
    python report_generator.py --input /path/to/data.csv --algos 1,7,19
    python report_generator.py --input /path/to/data.xlsx --algos all --output-dir ./out
    python report_generator.py --input data.csv --report-name quant_report
"""

from __future__ import annotations

import argparse
import html
import json
import logging
import math
import os
import sys
from dataclasses import dataclass, field  # field kept for MetricBundle defaults
from datetime import datetime
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd

# --------------------------------------------------------------------------- #
# Logging
# --------------------------------------------------------------------------- #
logger = logging.getLogger("algo_report")


def configure_logging(verbose: bool = False) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s | %(levelname)-7s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        stream=sys.stdout,
    )


# --------------------------------------------------------------------------- #
# Constants – columns
# --------------------------------------------------------------------------- #
# Columns kept from the raw input for the "data" sheet
DATA_COLS = ["user_id", "alias", "mtm_all", "allocation", "server",
             "date", "broker", "algo", "index", "dte"]

# Minimum columns that must exist in the raw file
REQUIRED_COLS = ["user_id", "mtm_all", "allocation", "date", "broker", "algo"]

# Algo-19 specific broker exclusions (case-insensitive)
# Rows where algo == "19" AND broker matches any of these are dropped.
ALGO19_EXCLUDED_BROKERS: frozenset[str] = frozenset({
    "mastertrust_noren",
    "mastertrust_dealer",
})

# Algos for which ONLY 0DTE rows are retained (matched against the "dte" column)
ZERO_DTE_ALGOS: frozenset[str] = frozenset({"1", "7", "15"})

# DTE value kept for zero-DTE algos (lowercased for comparison)
ZERO_DTE_VALUE: str = "0dte"

# Dates to completely exclude from ALL calculations (format: YYYY-MM-DD)
EXCLUDED_DATES: list = ["2024-04-18"]

# Base capital (fixed, 1 crore) used for Absolute PNL calculation
BASE_CAPITAL_DEFAULT: int = 10_000_000


# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class Config:
    input_path: Path
    output_dir: Path
    rolling_window: int = 63                    # trading days (~1 quarter)
    risk_free_annual: float = 0.065          # 3-yr avg 91-day T-Bill yield (2022–2025)
    max_gap_days: int = 3                       # gap tolerance before counting a break
    allocation_scale: int = 100                 # processed_allocation = raw × scale
    algos_raw: Optional[str] = None             # None → prompt
    date_format: Optional[str] = None
    report_name: str = "Algo_performance_std"   # base name for output HTML files
    base_capital: int = BASE_CAPITAL_DEFAULT    # fixed capital for Absolute PNL

    @property
    def processed_excel_path(self) -> Path:
        return self.output_dir / "Processed data.xlsx"

    @property
    def std_interactive_path(self) -> Path:
        name = self.report_name or "Algo_performance_std"
        return self.output_dir / f"{name}_interactive.html"

    @property
    def std_client_path(self) -> Path:
        name = self.report_name or "Algo_performance_std"
        return self.output_dir / f"{name}_client.html"


# --------------------------------------------------------------------------- #
# Data loading
# --------------------------------------------------------------------------- #
@dataclass
class LoadResult:
    df: pd.DataFrame          # processed rows (allocation already scaled)
    rows_in: int
    rows_dropped_excluded_date: int   # dropped because date is in EXCLUDED_DATES
    rows_dropped_algo19_broker: int   # dropped: algo-19 mastertrust_noren/dealer rows
    rows_dropped_dte: int             # dropped: non-0DTE rows for algos 1, 7, 15
    rows_dropped_date: int            # dropped: unparseable date
    rows_dropped_mtm: int
    rows_dropped_alloc: int
    rows_dropped_dupe: int


def _read_raw(path: Path) -> pd.DataFrame:
    """Read CSV or Excel into a DataFrame."""
    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xls", ".xlsm"):
        logger.info("Reading Excel: %s", path)
        return pd.read_excel(path, dtype=str)
    logger.info("Reading CSV: %s", path)
    return pd.read_csv(path, dtype=str, low_memory=False)


def load_data(cfg: Config) -> LoadResult:
    """
    Load raw file, validate columns, coerce types, apply business filters,
    scale allocation, filter to required columns only.

    Filters applied (in order):
      1. Date exclusion  : all rows for dates in EXCLUDED_DATES are dropped entirely.
      2. Algo-19 broker  : rows where algo == "19" AND broker is mastertrust_noren or
                           mastertrust_dealer are dropped.
      3. 0DTE filter     : for algos 1, 7, 15 only rows where dte == "0DTE"
                           (case-insensitive) are retained.
      4. Standard hygiene: bad dates, null mtm, null/<=0 allocation, duplicates.

    Returns processed rows ready for aggregation.
    """
    if not cfg.input_path.exists():
        raise FileNotFoundError(f"Input file not found: {cfg.input_path}")

    raw = _read_raw(cfg.input_path)
    rows_in = len(raw)

    # --- validate required columns ---
    missing = [c for c in REQUIRED_COLS if c not in raw.columns]
    if missing:
        raise ValueError(f"Input file is missing required columns: {missing}")

    df = raw.copy()

    # --- date coercion (needed early for date-exclusion filter) ---
    if cfg.date_format:
        df["date"] = pd.to_datetime(
            df["date"], format=cfg.date_format, errors="coerce"
        ).dt.normalize()
    else:
        # Try YYYY-MM-DD first (ISO format — matches this dataset).
        # Fall back to dayfirst (DD-MM-YYYY) only for values that fail ISO parse.
        parsed_iso = pd.to_datetime(df["date"], format="%Y-%m-%d", errors="coerce")
        needs_fallback = parsed_iso.isna() & df["date"].notna()
        if needs_fallback.any():
            parsed_fallback = pd.to_datetime(
                df.loc[needs_fallback, "date"], errors="coerce", dayfirst=True
            )
            parsed_iso = parsed_iso.copy()
            parsed_iso[needs_fallback] = parsed_fallback
        df["date"] = parsed_iso.dt.normalize()

    bad_date = df["date"].isna()
    rows_dropped_date = int(bad_date.sum())
    df = df[~bad_date]

    # --- FILTER 1: drop entirely excluded dates ---
    excluded_ts = [pd.Timestamp(d) for d in EXCLUDED_DATES]
    mask_excl_date = df["date"].isin(excluded_ts)
    rows_dropped_excluded_date = int(mask_excl_date.sum())
    if rows_dropped_excluded_date:
        excl_dates_found = df.loc[mask_excl_date, "date"].dt.strftime("%d %b %Y").unique().tolist()
        logger.info(
            "Date exclusion: dropping ALL %d rows for excluded date(s): %s",
            rows_dropped_excluded_date, excl_dates_found,
        )
    df = df[~mask_excl_date].copy()

    # --- numeric coercion ---
    df["mtm_all"] = pd.to_numeric(df["mtm_all"], errors="coerce")
    df["allocation"] = pd.to_numeric(df["allocation"], errors="coerce")

    bad_mtm = df["mtm_all"].isna()
    rows_dropped_mtm = int(bad_mtm.sum())
    df = df[~bad_mtm]

    bad_alloc = df["allocation"].isna() | (df["allocation"] <= 0)
    rows_dropped_alloc = int(bad_alloc.sum())
    df = df[~bad_alloc]

    # --- algo / user_id normalization ---
    df["algo"] = df["algo"].astype(str).str.strip()
    df["user_id"] = df["user_id"].astype(str).str.strip()

    # --- FILTER 2: algo-19 broker exclusion ---
    broker_str = df["broker"].astype(str).str.strip().str.lower()
    mask_algo19 = (df["algo"] == "19") & broker_str.isin(ALGO19_EXCLUDED_BROKERS)
    rows_dropped_algo19_broker = int(mask_algo19.sum())
    if rows_dropped_algo19_broker:
        dropped_brokers = (
            df.loc[mask_algo19, "broker"].astype(str).str.strip()
            .value_counts().to_dict()
        )
        logger.info(
            "Algo-19 broker exclusion: dropping %d rows (brokers: %s)",
            rows_dropped_algo19_broker, dropped_brokers,
        )
    df = df[~mask_algo19].copy()

    # --- FILTER 3: 0DTE filter for algos 1, 7, 15 ---
    # For these algos, only rows where dte == "0DTE" (case-insensitive) are kept.
    dte_str = df["dte"].astype(str).str.strip().str.lower() if "dte" in df.columns else pd.Series("", index=df.index)
    mask_dte_algos = df["algo"].isin(ZERO_DTE_ALGOS)
    mask_non_zero_dte = mask_dte_algos & (dte_str != ZERO_DTE_VALUE)
    rows_dropped_dte = int(mask_non_zero_dte.sum())
    if rows_dropped_dte:
        logger.info(
            "0DTE filter: dropping %d non-0DTE rows for algos %s",
            rows_dropped_dte, sorted(ZERO_DTE_ALGOS, key=algo_sort_key),
        )
    df = df[~mask_non_zero_dte].copy()

    # --- de-duplicate on (user_id, algo, date) ---
    before = len(df)
    df = df.sort_values(["user_id", "algo", "date"])
    df = df.drop_duplicates(subset=["user_id", "algo", "date"], keep="first")
    rows_dropped_dupe = before - len(df)

    if rows_dropped_excluded_date:
        logger.warning("Dropped %d rows for excluded date(s)", rows_dropped_excluded_date)
    if rows_dropped_algo19_broker:
        logger.warning("Dropped %d algo-19 rows excluded by broker filter", rows_dropped_algo19_broker)
    if rows_dropped_dte:
        logger.warning("Dropped %d non-0DTE rows for algos %s", rows_dropped_dte, sorted(ZERO_DTE_ALGOS))
    if rows_dropped_date:
        logger.warning("Dropped %d rows with unparseable date", rows_dropped_date)
    if rows_dropped_mtm:
        logger.warning("Dropped %d rows with null mtm_all", rows_dropped_mtm)
    if rows_dropped_alloc:
        logger.warning("Dropped %d rows with null/<=0 allocation", rows_dropped_alloc)
    if rows_dropped_dupe:
        logger.warning("Dropped %d duplicate (user, algo, date) rows", rows_dropped_dupe)

    if df.empty:
        raise ValueError("No valid rows remain after validation.")

    # --- scale allocation × 100 ---
    df["allocation"] = df["allocation"] * cfg.allocation_scale

    # --- keep only required output columns (in defined order) ---
    keep = [c for c in DATA_COLS if c in df.columns]
    df = df[keep].copy()

    logger.info(
        "Loaded %d rows (%d in) | algos=%s | span=%s..%s",
        len(df), rows_in,
        sorted(df["algo"].unique(), key=algo_sort_key),
        df["date"].min().date(), df["date"].max().date(),
    )
    return LoadResult(
        df=df,
        rows_in=rows_in,
        rows_dropped_excluded_date=rows_dropped_excluded_date,
        rows_dropped_algo19_broker=rows_dropped_algo19_broker,
        rows_dropped_dte=rows_dropped_dte,
        rows_dropped_date=rows_dropped_date,
        rows_dropped_mtm=rows_dropped_mtm,
        rows_dropped_alloc=rows_dropped_alloc,
        rows_dropped_dupe=rows_dropped_dupe,
    )


# --------------------------------------------------------------------------- #
# Processed data (daily aggregation)
# --------------------------------------------------------------------------- #
def build_processed_data(df: pd.DataFrame, base_capital: int = BASE_CAPITAL_DEFAULT) -> pd.DataFrame:
    """
    Aggregate to (date, algo) level.

    sum_allocation = Σ (allocation already scaled) per date per algo
    sum_mtm        = Σ mtm_all                     per date per algo
    PNL%           = (sum_mtm / sum_allocation) × 100
    Base Capital   = base_capital  (fixed constant, default 1 crore = 10,000,000)
    Absolute PNL   = (sum_mtm / sum_allocation) × base_capital

    Returned columns: date, algo, sum_allocation, sum_mtm, PNL%, Absolute PNL, Base Capital
    """
    grp = (
        df.groupby(["date", "algo"], sort=True)
        .agg(sum_allocation=("allocation", "sum"), sum_mtm=("mtm_all", "sum"))
        .reset_index()
    )
    grp = grp.sort_values(["algo", "date"]).reset_index(drop=True)

    ret_frac = grp["sum_mtm"] / grp["sum_allocation"]
    grp["PNL%"] = ret_frac * 100
    grp["Base Capital"] = base_capital
    grp["Absolute PNL"] = ret_frac * base_capital

    return grp


def save_processed_excel(df_data: pd.DataFrame, df_proc: pd.DataFrame, path: Path) -> None:
    """
    Write "Processed data.xlsx" with two sheets:
      Sheet 1 "data"           : df_data  (filtered raw rows, scaled allocation)
      Sheet 2 "processed data" : df_proc  (daily aggregated)
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        # "data" sheet: format date as string for readability
        df_out = df_data.copy()
        df_out["date"] = df_out["date"].dt.strftime("%d-%m-%Y")
        df_out.to_excel(writer, sheet_name="data", index=False)

        # "processed data" sheet
        dp_out = df_proc.copy()
        dp_out["date"] = dp_out["date"].dt.strftime("%d-%m-%Y")
        dp_out.to_excel(writer, sheet_name="processed data", index=False)

    logger.info("Wrote %s  [data: %d rows | processed data: %d rows]",
                path, len(df_data), len(df_proc))


# --------------------------------------------------------------------------- #
# Calendar / segmentation
# --------------------------------------------------------------------------- #
def build_calendar(df: pd.DataFrame) -> pd.DatetimeIndex:
    return pd.DatetimeIndex(sorted(df["date"].unique()))


def trading_days_per_year(calendar: pd.DatetimeIndex) -> float:
    if len(calendar) < 2:
        return float(len(calendar))
    span_days = (calendar.max() - calendar.min()).days
    years = span_days / 365.25 if span_days > 0 else (1.0 / 365.25)
    return float(len(calendar) / years)


def split_into_segments(
    series: pd.Series, calendar: pd.DatetimeIndex, max_gap_days: int = 3
) -> list[pd.Series]:
    if series.empty:
        return []
    series = series.sort_index()
    pos = calendar.get_indexer(series.index)
    segments: list[pd.Series] = []
    start = 0
    for i in range(1, len(pos)):
        if (pos[i] - pos[i - 1] - 1) > max_gap_days:
            segments.append(series.iloc[start:i])
            start = i
    segments.append(series.iloc[start:])
    return segments


def largest_segment(segments: list[pd.Series]) -> pd.Series:
    if not segments:
        return pd.Series(dtype=float)
    return max(segments, key=len)


# --------------------------------------------------------------------------- #
# Individual metric functions  (all operate on fractional daily return Series)
# --------------------------------------------------------------------------- #
def m_cumulative_return(r: pd.Series) -> float:
    return float(r.sum()) if len(r) else float("nan")


def m_cagr(r: pd.Series, n_per_year: float) -> float:
    """
    Geometric CAGR: (1 + cumulative_return)^(N / actual_days) - 1
    Correctly accounts for compounding — consistent with the user's manual check:
        total_pnl / base_capital = cumulative return → annualised geometrically.
    """
    if len(r) == 0 or n_per_year <= 0:
        return float("nan")
    cum = float(r.sum())
    years = len(r) / n_per_year
    if years <= 0 or (1 + cum) <= 0:
        return float("nan")
    return float((1 + cum) ** (1 / years) - 1)


def m_annual_volatility(r: pd.Series, n_per_year: float) -> float:
    if len(r) < 2:
        return float("nan")
    return float(r.std(ddof=1) * math.sqrt(n_per_year))


def m_sharpe(r: pd.Series, n_per_year: float, rf_annual: float = 0.0) -> float:
    if len(r) < 2:
        return float("nan")
    sd = r.std(ddof=1)
    if sd == 0 or math.isnan(sd):
        return float("nan")
    rf_period = rf_annual / n_per_year
    return float((r.mean() - rf_period) / sd * math.sqrt(n_per_year))


def m_sortino(r: pd.Series, n_per_year: float, rf_annual: float = 0.0) -> float:
    if len(r) < 2:
        return float("nan")
    rf_period = rf_annual / n_per_year
    downside = np.minimum(r.values - 0.0, 0.0)
    dd = math.sqrt(np.mean(downside ** 2))
    if dd == 0:
        return float("nan")
    return float((r.mean() - rf_period) / dd * math.sqrt(n_per_year))


def m_equity_curve(r: pd.Series) -> pd.Series:
    return r.cumsum()


def m_max_drawdown(equity: pd.Series) -> float:
    if equity.empty:
        return float("nan")
    peak = equity.cummax()
    return float((equity - peak).min())


def m_avg_drawdown_days(equity: pd.Series) -> float:
    if equity.empty:
        return float("nan")
    peak = equity.cummax()
    underwater = equity.values < peak.values - 1e-12
    lengths: list[int] = []
    cur = 0
    for u in underwater:
        if u:
            cur += 1
        elif cur > 0:
            lengths.append(cur)
            cur = 0
    if cur > 0:
        lengths.append(cur)
    return float(np.mean(lengths)) if lengths else 0.0


def m_calmar(cagr: float, max_dd: float) -> float:
    if max_dd is None or math.isnan(max_dd) or max_dd == 0:
        return float("nan")
    return float(cagr / abs(max_dd))


def m_max_consecutive(r: pd.Series, wins: bool) -> int:
    if r.empty:
        return 0
    cond = (r.values > 0) if wins else (r.values <= 0)
    best = run = 0
    for c in cond:
        if c:
            run += 1
            best = max(best, run)
        else:
            run = 0
    return int(best)


def m_kelly_discrete(r: pd.Series) -> float:
    if r.empty:
        return float("nan")
    wins = r[r > 0]
    losses = r[r <= 0]
    n = len(r)
    w = len(wins) / n
    avg_win = wins.mean() if len(wins) else 0.0
    avg_loss = abs(losses.mean()) if len(losses) else 0.0
    if avg_loss == 0 or avg_win == 0:
        return float("nan")
    return float(w - (1 - w) / (avg_win / avg_loss))


def m_rolling_sharpe(
    r: pd.Series, window: int, n_per_year: float, rf_annual: float = 0.0
) -> pd.Series:
    if len(r) < window:
        return pd.Series(dtype=float)

    rf_daily = rf_annual / n_per_year if n_per_year > 0 else 0.0

    def _sharpe(x: np.ndarray) -> float:
        sd = np.std(x, ddof=1)
        if sd == 0:
            return np.nan
        return float((np.mean(x) - rf_daily) / sd * math.sqrt(n_per_year))

    return r.rolling(window).apply(_sharpe, raw=True)


# --------------------------------------------------------------------------- #
# Metric bundle
# --------------------------------------------------------------------------- #
@dataclass
class MetricBundle:
    n_days: int = 0
    n_segments: int = 1
    has_breaks: bool = False
    largest_segment_days: int = 0
    date_start: Optional[pd.Timestamp] = None
    date_end: Optional[pd.Timestamp] = None

    cumulative_return: float = float("nan")
    cagr: float = float("nan")
    annual_vol: float = float("nan")
    sharpe: float = float("nan")
    sortino: float = float("nan")
    max_drawdown: float = float("nan")
    avg_drawdown_days: float = float("nan")
    calmar: float = float("nan")
    max_consec_wins: int = 0
    max_consec_losses: int = 0
    kelly: float = float("nan")

    rolling_sharpe_last: float = float("nan")
    rolling_sharpe_series: list[float] = field(default_factory=list)
    rolling_dates: list[str] = field(default_factory=list)

    equity_dates: list[str] = field(default_factory=list)
    equity_vals: list[float] = field(default_factory=list)


def compute_metrics(
    series: pd.Series, calendar: pd.DatetimeIndex, cfg: Config, n_per_year: float
) -> MetricBundle:
    """
    Full metric set for a date-indexed fractional daily return series.
    Distributional metrics use ALL observations.
    Path-dependent metrics use the largest contiguous segment.
    """
    b = MetricBundle()
    series = series.dropna().sort_index()
    if series.empty:
        return b

    b.n_days = len(series)
    b.date_start = series.index.min()
    b.date_end = series.index.max()

    segments = split_into_segments(series, calendar, cfg.max_gap_days)
    b.n_segments = len(segments)
    b.has_breaks = b.n_segments > 1
    seg = largest_segment(segments)
    b.largest_segment_days = len(seg)

    # distributional (all data)
    b.cumulative_return = m_cumulative_return(series)
    b.cagr = m_cagr(series, n_per_year)
    b.annual_vol = m_annual_volatility(series, n_per_year)
    b.sharpe = m_sharpe(series, n_per_year, cfg.risk_free_annual)
    b.sortino = m_sortino(series, n_per_year, cfg.risk_free_annual)
    b.kelly = m_kelly_discrete(series)

    # path-dependent (largest segment only)
    equity = m_equity_curve(seg)
    b.max_drawdown = m_max_drawdown(equity)
    b.avg_drawdown_days = m_avg_drawdown_days(equity)
    b.calmar = m_calmar(b.cagr, b.max_drawdown)
    b.max_consec_wins = m_max_consecutive(seg, wins=True)
    b.max_consec_losses = m_max_consecutive(seg, wins=False)

    # rolling sharpe
    rs = m_rolling_sharpe(series, cfg.rolling_window, n_per_year, cfg.risk_free_annual).dropna()
    if not rs.empty:
        b.rolling_sharpe_last = float(rs.iloc[-1])
        b.rolling_sharpe_series = [float(v) for v in rs.values]
        b.rolling_dates = [pd.Timestamp(d).strftime("%Y-%m-%d") for d in rs.index]

    # additive equity curve for charting (downsampled)
    eq_full = series.cumsum()
    ds_dates, ds_vals = _downsample(list(eq_full.index), list(eq_full.values))
    b.equity_dates = [pd.Timestamp(d).strftime("%Y-%m-%d") for d in ds_dates]
    b.equity_vals = [float(v) for v in ds_vals]

    return b


# --------------------------------------------------------------------------- #
# Helper utilities
# --------------------------------------------------------------------------- #
def _na(v) -> bool:
    return v is None or (isinstance(v, float) and math.isnan(v))


def algo_sort_key(a: str) -> tuple:
    s = str(a)
    return (0, int(s), "") if s.isdigit() else (1, 0, s)


def _downsample(idx: list, vals: list, max_points: int = 260) -> tuple[list, list]:
    n = len(vals)
    if n <= max_points:
        return list(idx), list(vals)
    step = math.ceil(n / max_points)
    di = list(idx[::step])
    dv = list(vals[::step])
    if idx[-1] != di[-1]:
        di.append(idx[-1])
        dv.append(vals[-1])
    return di, dv


def series_points(s: pd.Series) -> list[list]:
    return [[pd.Timestamp(d).strftime("%Y-%m-%d"), float(v)]
            for d, v in s.dropna().sort_index().items()]


def fmt_pct(frac: float, dp: int = 2) -> str:
    if _na(frac):
        return "&mdash;"
    return f"{frac * 100:.{dp}f}%"


def fmt_return(frac: float, dp: int = 2) -> str:
    """Like fmt_pct but appends * to flag additive-return values."""
    if _na(frac):
        return "&mdash;"
    return f"{frac * 100:.{dp}f}%*"


def fmt_ratio(v: float, dp: int = 2) -> str:
    if _na(v):
        return "&mdash;"
    return f"{v:.{dp}f}"


def fmt_int(v: int) -> str:
    return str(int(v))


def fmt_days(v: float, dp: int = 1) -> str:
    if _na(v):
        return "&mdash;"
    return f"{v:.{dp}f}"


def fmt_date(ts) -> str:
    if ts is None:
        return "&mdash;"
    return pd.Timestamp(ts).strftime("%d %b %Y").upper()


def sparkline_svg(values: list[float], width: int = 130, height: int = 30) -> str:
    vals = [v for v in values if not (isinstance(v, float) and math.isnan(v))]
    if len(vals) < 2:
        return '<span class="muted">&mdash;</span>'
    lo, hi = min(vals), max(vals)
    rng = (hi - lo) or 1.0
    n = len(vals)
    pad = 2
    pts = []
    for i, v in enumerate(vals):
        x = pad + (width - 2 * pad) * i / (n - 1)
        y = pad + (height - 2 * pad) * (1 - (v - lo) / rng)
        pts.append(f"{x:.1f},{y:.1f}")
    zero_line = ""
    if lo < 0 < hi:
        zy = pad + (height - 2 * pad) * (1 - (0 - lo) / rng)
        zero_line = (f'<line x1="{pad}" y1="{zy:.1f}" x2="{width - pad}" y2="{zy:.1f}" '
                     f'class="spark-zero"/>')
    return (f'<svg class="spark" viewBox="0 0 {width} {height}" width="{width}" '
            f'height="{height}" preserveAspectRatio="none">{zero_line}'
            f'<polyline points="{" ".join(pts)}" class="spark-line"/></svg>')


# --------------------------------------------------------------------------- #
# Metric table definition
# --------------------------------------------------------------------------- #
# (key, display label, bundle attribute, formatter)
METRICS: list[tuple] = [
    ("cagr",      "CAGR",                  "cagr",              fmt_return),
    ("cumret",    "Cumulative Return",      "cumulative_return", fmt_return),
    ("sharpe",    "Sharpe",                 "sharpe",            fmt_ratio),
    ("sortino",   "Sortino",                "sortino",           fmt_ratio),
    ("calmar",    "Calmar",                 "calmar",            fmt_ratio),
    ("vol",       "Annual Volatility",      "annual_vol",        fmt_pct),
    ("maxdd",     "Max Drawdown",           "max_drawdown",      fmt_pct),
    ("avgddays",  "Avg Drawdown Days",      "avg_drawdown_days", fmt_days),
    ("maxwins",   "Max Consecutive Wins",   "max_consec_wins",   fmt_int),
    ("maxlosses", "Max Consecutive Losses", "max_consec_losses", fmt_int),
    ("kelly",     "Kelly Criterion",        "kelly",             fmt_pct),
    ("rolllast",  "Rolling Sharpe (last)",  "rolling_sharpe_last", fmt_ratio),
]

_POS_KEYS = {"cagr", "cumret", "sharpe", "sortino", "calmar"}


def _cell_class(key: str, raw) -> str:
    if key == "maxdd":
        return "neg"
    if key in _POS_KEYS:
        if _na(raw):
            return ""
        return "pos" if raw > 0 else ("neg" if raw < 0 else "")
    return ""


def _cell(algo: str, key: str, attr: str, fmt, b: MetricBundle) -> str:
    raw = getattr(b, attr)
    dv = "-Infinity" if _na(raw) else raw
    klass = _cell_class(key, raw)
    return (f"<td class='num metric-val {klass}' "
            f"data-algo='{html.escape(str(algo))}' "
            f"data-metric='{key}' data-v='{dv}'>{fmt(raw)}</td>")


# --------------------------------------------------------------------------- #
# CSS
# --------------------------------------------------------------------------- #
CSS = """
:root{
  --ink:#0f1726; --ink-2:#5b6678; --ink-3:#8a93a3;
  --line:#e6e9ef; --line-2:#f0f2f6;
  --bg:#f4f6f9; --surface:#ffffff;
  --accent:#2563eb; --accent-2:#0d9488;
  --pos:#0f8a5f; --neg:#c0392b; --warn:#9a5b14;
  --shadow:0 1px 2px rgba(16,23,38,.06),0 4px 16px rgba(16,23,38,.05);
  --mono:"SFMono-Regular",Consolas,"Liberation Mono",Menlo,monospace;
  --sans:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
}
*{box-sizing:border-box}
html{scroll-behavior:smooth}
body{margin:0;background:var(--bg);color:var(--ink);font-family:var(--sans);
  font-size:14px;line-height:1.5;-webkit-font-smoothing:antialiased}
.appbar{background:#0f1726;color:#fff;padding:18px 0}
.appbar .wrap{display:flex;align-items:flex-end;justify-content:space-between;gap:16px}
.appbar .eyebrow{color:#9fb3d1}
.appbar h1{color:#fff;margin:4px 0 0;font-size:22px;font-weight:700;letter-spacing:-.01em}
.wrap{max-width:1120px;margin:0 auto;padding:0 24px}
.eyebrow{font-size:11px;letter-spacing:.16em;text-transform:uppercase;font-weight:600;color:var(--ink-2)}
.lede{color:var(--ink-2);font-size:13.5px;max-width:70ch;margin:6px 0 0}
main{padding:26px 0 70px}
section{margin:0 0 26px}
h2{font-size:13px;font-weight:700;letter-spacing:.04em;text-transform:uppercase;
  color:var(--ink-2);margin:0 0 12px}
.toolbar{position:sticky;top:0;z-index:20;background:rgba(244,246,249,.92);
  backdrop-filter:blur(6px);border-bottom:1px solid var(--line);
  padding:10px 0;margin-bottom:22px}
.toolbar .wrap{display:flex;flex-wrap:wrap;align-items:center;gap:10px}
.search{flex:1 1 220px;min-width:160px;display:flex;align-items:center;gap:8px;
  background:var(--surface);border:1px solid var(--line);border-radius:8px;padding:7px 11px}
.search input{border:0;outline:0;width:100%;font-size:13px;background:transparent;color:var(--ink)}
.search svg{flex:none;color:var(--ink-3)}
.btn{border:1px solid var(--line);background:var(--surface);border-radius:8px;
  padding:7px 12px;font-size:12.5px;font-weight:600;color:var(--ink-2);cursor:pointer}
.btn:hover{border-color:var(--ink-3);color:var(--ink)}
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px}
.stat{background:var(--surface);border:1px solid var(--line);border-radius:12px;
  padding:14px 16px;box-shadow:var(--shadow)}
.stat .k{font-size:10.5px;letter-spacing:.08em;text-transform:uppercase;color:var(--ink-3)}
.stat .v{font-family:var(--mono);font-size:22px;font-weight:600;margin-top:4px;letter-spacing:-.02em}
.meta{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin-top:4px}
.meta .cell{background:var(--surface);border:1px solid var(--line);border-radius:10px;padding:10px 13px}
.meta .k{font-size:10px;letter-spacing:.08em;text-transform:uppercase;color:var(--ink-3)}
.meta .v{font-family:var(--mono);font-size:14px;font-weight:600;margin-top:2px}
.panel{background:var(--surface);border:1px solid var(--line);border-radius:12px;
  box-shadow:var(--shadow);overflow:hidden}
.panel-pad{padding:16px 18px}
.card{background:var(--surface);border:1px solid var(--line);border-radius:12px;
  box-shadow:var(--shadow);margin-bottom:14px;overflow:hidden}
.card-head{display:flex;align-items:center;gap:12px;padding:14px 18px;cursor:pointer;user-select:none}
.card-head:hover{background:var(--line-2)}
.card-head .chev{transition:transform .18s;color:var(--ink-3);flex:none}
.card.collapsed .chev{transform:rotate(-90deg)}
.card.collapsed .card-body{display:none}
.card-head .title{font-size:16px;font-weight:700;letter-spacing:-.01em}
.card-head .cov{font-size:12px;color:var(--ink-3);margin-left:auto;text-align:right}
.card-body{border-top:1px solid var(--line)}
.card-grid{display:grid;grid-template-columns:1.1fr .9fr;gap:0}
.card-grid > .chart-box{border-right:1px solid var(--line)}
@media(max-width:760px){.card-grid{grid-template-columns:1fr}
  .card-grid > .chart-box{border-right:0;border-bottom:1px solid var(--line)}}
.chart-box{padding:16px 18px}
.chart-box .lab{font-size:11px;letter-spacing:.06em;text-transform:uppercase;color:var(--ink-3);margin-bottom:8px}
.tbl-wrap{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:13px}
thead th{position:sticky;top:0;background:var(--surface)}
th,td{text-align:left;padding:9px 14px;border-bottom:1px solid var(--line-2)}
th{font-size:10.5px;letter-spacing:.06em;text-transform:uppercase;color:var(--ink-3);font-weight:600}
tbody tr:nth-child(even){background:#fbfcfe}
tbody tr:hover{background:#f2f6ff}
td.metric{font-weight:600;color:var(--ink)}
td.num,th.num{text-align:right;font-family:var(--mono);font-variant-numeric:tabular-nums}
th.sortable{cursor:pointer;white-space:nowrap}
th.sortable:hover{color:var(--ink)}
th.sortable .arr{opacity:.35;font-size:9px}
th.sortable.asc .arr,th.sortable.desc .arr{opacity:1}
.pos{color:var(--pos)} .neg{color:var(--neg)} .muted{color:var(--ink-3)}
.flag{display:inline-block;font-size:10.5px;font-weight:600;padding:3px 9px;border-radius:999px;
  background:#fef3e2;color:var(--warn);border:1px solid #f4d9b0}
.flag.ok{background:#e8f6ee;color:var(--pos);border-color:#cdead9}
.chart{width:100%;height:auto;display:block}
.chart .grid{stroke:#eef1f6;stroke-width:1}
.chart .zero{stroke:#c2c9d6;stroke-width:1;stroke-dasharray:3 3}
.chart .ytick{fill:var(--ink-3);font-size:10px;font-family:var(--mono);text-anchor:end}
.chart .xtick{fill:var(--ink-3);font-size:10px;font-family:var(--mono);text-anchor:middle}
.chart .legend{fill:var(--ink-2);font-size:11px;font-weight:600}
.chart-empty{color:var(--ink-3);font-size:12px;padding:30px 0;text-align:center}
.spark{display:inline-block;vertical-align:middle}
.spark-line{fill:none;stroke:var(--accent);stroke-width:1.4}
.spark-zero{stroke:#c7ccd6;stroke-width:1;stroke-dasharray:2 2}
.empty-note{color:var(--ink-3);font-size:13px;padding:16px 0}
.chart-canvas{width:100%;display:block;cursor:crosshair;touch-action:none}
.chart-hd{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:6px}
.chart-hd .lab{font-size:11px;letter-spacing:.06em;text-transform:uppercase;color:var(--ink-3)}
.chart-hint{font-size:10.5px;color:var(--ink-3)}
.ch-tip{position:fixed;z-index:50;pointer-events:none;background:#0f1726;color:#fff;
  font-size:11.5px;line-height:1.35;padding:6px 9px;border-radius:7px;box-shadow:var(--shadow);
  display:none;white-space:nowrap;font-family:var(--mono)}
.ch-tip b{font-family:var(--sans)}
.daterow{display:flex;flex-wrap:wrap;align-items:center;gap:8px;margin-top:8px;flex-basis:100%}
.daterow .lab{font-size:11px;letter-spacing:.06em;text-transform:uppercase;color:var(--ink-3)}
.daterow input[type=date]{font-family:var(--sans);font-size:13px;color:var(--ink);
  border:1px solid var(--line);border-radius:8px;padding:6px 9px;background:var(--surface)}
.daterow .preset{border:1px solid var(--line);background:var(--surface);border-radius:8px;
  padding:6px 11px;font-size:12px;font-weight:600;color:var(--ink-2);cursor:pointer}
.daterow .preset:hover{border-color:var(--ink-3);color:var(--ink)}
.daterow .preset.active{background:var(--accent);color:#fff;border-color:var(--accent)}
#rangeLabel{font-family:var(--mono);font-size:12px;color:var(--ink-2);margin-left:auto}
.toTop{position:fixed;right:22px;bottom:22px;width:40px;height:40px;border-radius:50%;
  border:1px solid var(--line);background:var(--surface);box-shadow:var(--shadow);
  cursor:pointer;display:none;align-items:center;justify-content:center;color:var(--ink-2);z-index:30}
.toTop.show{display:flex}
.warn-badge{display:inline-flex;align-items:center;gap:4px;font-size:10.5px;font-weight:600;
  padding:3px 9px;border-radius:999px;background:#fff7ed;color:#9a5b14;border:1px solid #f4d9b0;
  white-space:nowrap;cursor:default}
.method-note{background:#fffbf0;border:1px solid #f4d9b0;border-radius:12px;
  padding:14px 18px;margin-bottom:22px;font-size:12.5px}
.method-note h3{font-size:10.5px;letter-spacing:.1em;text-transform:uppercase;
  color:#9a5b14;font-weight:700;margin:0 0 10px}
.method-note ul{margin:0;padding-left:18px}
.method-note li{margin:4px 0;color:var(--ink-2);line-height:1.55}
.method-note b{color:var(--ink)}
footer{margin-top:40px;border-top:1px solid var(--line);padding-top:16px;
  font-size:12px;color:var(--ink-2)}
footer h3{font-size:11px;text-transform:uppercase;letter-spacing:.08em;margin:0 0 8px;color:var(--ink-3)}
footer ul{margin:0;padding-left:18px} footer li{margin:3px 0}
@media print{body{background:#fff}.toolbar,.toTop,.btn,.search{display:none!important}
  .card{break-inside:avoid;box-shadow:none}.card.collapsed .card-body{display:block!important}
  .appbar{background:#fff;color:#000;border-bottom:2px solid #000}
  .appbar h1,.appbar .eyebrow{color:#000}}
"""

# --------------------------------------------------------------------------- #
# JavaScript – interactive report
# --------------------------------------------------------------------------- #
JS = r"""
(function(){
  var ROOT=document.getElementById('reportData'); if(!ROOT) return;
  var DATA=JSON.parse(ROOT.textContent);
  var P=DATA.params, WIN=P.rollingWindow, MAXGAP=P.maxGapDays;
  var RF=P.rfAnnual||0;
  var CAL={}; DATA.calendar.forEach(function(d,i){ CAL[d]=i; });
  function toMs(iso){ return Date.parse(iso); }
  var DMIN=toMs(DATA.dateMin), DMAX=toMs(DATA.dateMax);
  var gFrom=DMIN, gTo=DMAX;

  /* ---- math helpers ---- */
  function mean(a){ if(!a.length) return NaN; var s=0; for(var i=0;i<a.length;i++) s+=a[i]; return s/a.length; }
  function sstd(a){ if(a.length<2) return NaN; var m=mean(a),s=0; for(var i=0;i<a.length;i++){var d=a[i]-m; s+=d*d;} return Math.sqrt(s/(a.length-1)); }
  function filt(pts,from,to){ var o=[]; for(var i=0;i<pts.length;i++){var t=toMs(pts[i][0]); if(t>=from&&t<=to) o.push(pts[i]);} return o; }
  function segs(pts){ if(!pts.length) return []; var out=[],start=0;
    for(var i=1;i<pts.length;i++){ var g=CAL[pts[i][0]]-CAL[pts[i-1][0]]-1; if(g>MAXGAP){ out.push(pts.slice(start,i)); start=i; } }
    out.push(pts.slice(start)); return out; }
  function largest(ss){ if(!ss.length) return []; var b=ss[0]; for(var i=1;i<ss.length;i++) if(ss[i].length>b.length) b=ss[i]; return b; }

  /* N = per-algo annualisation factor passed from server; rf = annual risk-free rate */
  function metrics(pts, N, rf){
    if(N==null||isNaN(N)||N<=0) N=P.nPerYear;
    if(rf==null||isNaN(rf)) rf=RF;
    var rfDaily=rf/N;
    var m={n:pts.length,from:null,to:null,segs:1,breaks:false,equity:[],rolling:[],
           cagr:NaN,cumret:NaN,sharpe:NaN,sortino:NaN,calmar:NaN,vol:NaN,
           maxdd:NaN,avgddays:0,maxwins:0,maxlosses:0,kelly:NaN,rolllast:NaN};
    if(!pts.length) return m;
    m.from=pts[0][0]; m.to=pts[pts.length-1][0];
    var r=pts.map(function(p){return p[1];});
    var mu=mean(r), sd=sstd(r);
    m.cumret=r.reduce(function(a,b){return a+b;},0);
    var years=pts.length/N;
    m.cagr=(years>0&&(1+m.cumret)>0)?Math.pow(1+m.cumret,1/years)-1:NaN;
    m.vol=(pts.length<2)?NaN:sd*Math.sqrt(N);
    m.sharpe=(pts.length<2||!(sd>0))?NaN:(mu-rfDaily)/sd*Math.sqrt(N);
    var dn=0; for(var i=0;i<r.length;i++){ if(r[i]<0) dn+=r[i]*r[i]; } dn=Math.sqrt(dn/r.length);
    m.sortino=(pts.length<2||!(dn>0))?NaN:(mu-rfDaily)/dn*Math.sqrt(N);
    var cum=0; for(var i=0;i<pts.length;i++){ cum+=pts[i][1]; m.equity.push([toMs(pts[i][0]),cum*100]); }
    var ss=segs(pts); m.segs=ss.length; m.breaks=ss.length>1;
    var seg=largest(ss), sr=seg.map(function(p){return p[1];});
    var eq=0,peak=0,maxdd=0,runs=[],cur=0;
    for(var i=0;i<sr.length;i++){ eq+=sr[i]; if(eq>peak) peak=eq; var d=eq-peak; if(d<maxdd) maxdd=d;
      if(eq<peak-1e-12){cur++;} else {if(cur>0){runs.push(cur);cur=0;}} }
    if(cur>0) runs.push(cur);
    m.maxdd=seg.length?maxdd:NaN;
    m.avgddays=runs.length?mean(runs):0;
    m.calmar=(maxdd===0||isNaN(m.cagr))?NaN:m.cagr/Math.abs(maxdd);
    var w=0,bw=0,l=0,bl=0;
    for(var i=0;i<sr.length;i++){ if(sr[i]>0){w++;if(w>bw)bw=w;l=0;}else{l++;if(l>bl)bl=l;w=0;} }
    m.maxwins=bw; m.maxlosses=bl;
    var wins=r.filter(function(x){return x>0;}), loss=r.filter(function(x){return x<=0;});
    var aw=wins.length?mean(wins):0, al=loss.length?Math.abs(mean(loss)):0;
    m.kelly=(al===0||aw===0)?NaN:(wins.length/r.length)-(1-wins.length/r.length)/(aw/al);
    if(r.length>=WIN){
      var rs=[];
      for(var i=WIN-1;i<r.length;i++){ var win=r.slice(i-WIN+1,i+1); var s=sstd(win); var v=(s>0)?(mean(win)-rfDaily)/s*Math.sqrt(N):NaN;
        if(!isNaN(v)){rs.push(v); m.rolling.push([toMs(pts[i][0]),v]);} }
      if(rs.length) m.rolllast=rs[rs.length-1];
    }
    return m;
  }

  /* ---- formatters ---- */
  function fpct(x){ return (x==null||isNaN(x))?'—':(x*100).toFixed(2)+'%'; }
  function fratio(x){ return (x==null||isNaN(x))?'—':x.toFixed(2); }
  function fint(x){ return (x==null||isNaN(x))?'0':String(Math.round(x)); }
  function fdays(x){ return (x==null||isNaN(x))?'—':x.toFixed(1); }
  var FMT={cagr:fpct,cumret:fpct,sharpe:fratio,sortino:fratio,calmar:fratio,vol:fpct,
           maxdd:fpct,avgddays:fdays,maxwins:fint,maxlosses:fint,kelly:fpct,rolllast:fratio};
  var SIGN={cagr:1,cumret:1,sharpe:1,sortino:1,calmar:1,maxdd:-1};
  function cls(k,x){ if(k in SIGN){if(SIGN[k]<0) return 'neg'; if(x>0) return 'pos'; if(x<0) return 'neg';} return ''; }
  function fdate(ms){ var d=new Date(ms); var M=['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
    return ('0'+d.getUTCDate()).slice(-2)+' '+M[d.getUTCMonth()]+' '+d.getUTCFullYear(); }

  /* ---- canvas chart ---- */
  var TIP=document.createElement('div'); TIP.className='ch-tip'; document.body.appendChild(TIP);
  function Chart(canvas){
    var c=canvas, ctx=c.getContext('2d'); var self={canvas:c,series:[],yfmt:fratio,xMin:DMIN,xMax:DMAX};
    self.setDomain=function(a,b){ if(a!=null&&b!=null&&b>a){self.xMin=a;self.xMax=b;} else {self.xMin=gFrom;self.xMax=gTo;} };
    function size(){ var w=c.clientWidth||560,h=parseInt(c.dataset.h||'220',10),dpr=window.devicePixelRatio||1;
      c.width=w*dpr; c.height=h*dpr; ctx.setTransform(dpr,0,0,dpr,1,1); return {w:w,h:h}; }
    function draw(){
      var d=size(),w=d.w,h=d.h; ctx.clearRect(0,0,w,h);
      var mL=52,mR=10,mT=12,mB=22,iw=w-mL-mR,ih=h-mT-mB;
      var from=self.xMin,to=self.xMax,span=(to-from)||1;
      var vis=[]; self.series.forEach(function(s){ s.points.forEach(function(p){ if(p[0]>=from&&p[0]<=to) vis.push(p[1]); }); });
      if(!vis.length){ctx.fillStyle='#8a93a3';ctx.font='12px sans-serif';ctx.fillText('No data in range',mL,mT+20);return;}
      var lo=Math.min.apply(null,vis),hi=Math.max.apply(null,vis); if(self.zeroBase&&lo>0)lo=0; if(self.zeroBase&&hi<0)hi=0;
      if(hi===lo){hi+=1;lo-=1;} var pad=(hi-lo)*0.08; lo-=pad; hi+=pad;
      self._map={mL:mL,mT:mT,iw:iw,ih:ih,from:from,span:span,lo:lo,hi:hi};
      function X(t){return mL+iw*(t-from)/span;} function Y(v){return mT+ih*(1-(v-lo)/(hi-lo));}
      ctx.strokeStyle='#eef1f6';ctx.lineWidth=1;ctx.fillStyle='#8a93a3';ctx.font='10px monospace';ctx.textAlign='end';
      for(var t=0;t<5;t++){var v=lo+(hi-lo)*t/4,y=Y(v);ctx.beginPath();ctx.moveTo(mL,y);ctx.lineTo(w-mR,y);ctx.stroke();ctx.fillText(self.yfmt(v),mL-5,y+3);}
      if(lo<0&&hi>0){var yz=Y(0);ctx.strokeStyle='#c2c9d6';ctx.setLineDash([3,3]);ctx.beginPath();ctx.moveTo(mL,yz);ctx.lineTo(w-mR,yz);ctx.stroke();ctx.setLineDash([]);}
      ctx.textAlign='center';
      [from,from+span/2,to].forEach(function(t){var dd=new Date(t);ctx.fillText(dd.toISOString().slice(0,10),X(t),h-7);});
      self.series.forEach(function(s){ctx.strokeStyle=s.color;ctx.lineWidth=1.7;ctx.beginPath();var started=false;
        for(var i=0;i<s.points.length;i++){var p=s.points[i];if(p[0]<from||p[0]>to) continue;var px=X(p[0]),py=Y(p[1]);
          if(!started){ctx.moveTo(px,py);started=true;}else ctx.lineTo(px,py);}ctx.stroke();});
      var lx=mL+4,ly=mT+8;ctx.textAlign='start';ctx.font='11px sans-serif';
      self.series.forEach(function(s){ctx.fillStyle=s.color;ctx.fillRect(lx,ly-7,10,3);ctx.fillStyle='#5b6678';ctx.fillText(s.name,lx+14,ly-2);lx+=24+7*s.name.length;});
    }
    self.render=function(series){self.series=series;draw();};
    self.redraw=draw;
    var dragging=false,dragX0=0,dragX1=0;
    function evX(e){var r=c.getBoundingClientRect();return (e.clientX||(e.touches&&e.touches[0].clientX))-r.left;}
    function xToMs(x){var mp=self._map;if(!mp) return null;return mp.from+(x-mp.mL)/mp.iw*mp.span;}
    c.addEventListener('mousedown',function(e){dragging=true;dragX0=dragX1=evX(e);});
    c.addEventListener('mousemove',function(e){
      var x=evX(e);
      if(dragging){dragX1=x;draw();var mp=self._map;if(mp){ctx.fillStyle='rgba(37,99,235,.12)';ctx.fillRect(Math.min(dragX0,dragX1),mp.mT,Math.abs(dragX1-dragX0),mp.ih);}return;}
      var mp=self._map;if(!mp||!self.series.length){TIP.style.display='none';return;}
      var t=xToMs(x),best=null,bestd=1e15,bestS=null;
      self.series.forEach(function(s){for(var i=0;i<s.points.length;i++){var p=s.points[i];if(p[0]<self.xMin||p[0]>self.xMax) continue;var dd=Math.abs(p[0]-t);if(dd<bestd){bestd=dd;best=p;bestS=s;}}});
      if(best){TIP.style.display='block';TIP.style.left=(e.clientX+12)+'px';TIP.style.top=(e.clientY+12)+'px';
        var rows=self.series.map(function(s){var pt=null;for(var i=0;i<s.points.length;i++)if(s.points[i][0]===best[0]){pt=s.points[i];break;}
          return pt?('<span style="color:'+s.color+'">●</span> '+s.name+': '+self.yfmt(pt[1])):'';}).filter(Boolean).join('<br>');
        TIP.innerHTML='<b>'+fdate(best[0])+'</b><br>'+rows;}
    });
    c.addEventListener('mouseleave',function(){TIP.style.display='none';if(dragging){dragging=false;draw();}});
    window.addEventListener('mouseup',function(e){
      if(!dragging)return;dragging=false;var a=xToMs(Math.min(dragX0,dragX1)),b=xToMs(Math.max(dragX0,dragX1));
      if(a!=null&&b!=null&&Math.abs(dragX1-dragX0)>6){apply(Math.max(DMIN,a),Math.min(DMAX,b));}else draw();});
    c.addEventListener('dblclick',function(){apply(DMIN,DMAX);});
    c.addEventListener('wheel',function(e){e.preventDefault();var mp=self._map;if(!mp)return;var t=xToMs(evX(e));
      var f=(e.deltaY>0)?1.2:0.8;var nf=t-(t-self.xMin)*f,nt=t+(self.xMax-t)*f;apply(Math.max(DMIN,nf),Math.min(DMAX,nt));},{passive:false});
    return self;
  }

  var CH={};
  document.querySelectorAll('canvas[data-chart]').forEach(function(cv){
    var ch=Chart(cv); ch.zeroBase=(cv.dataset.chart==='equity');
    ch.yfmt=(cv.dataset.chart==='equity')?function(v){return v.toFixed(0)+'%';}:fratio;
    CH[cv.dataset.algo+'|'+cv.dataset.chart]=ch;
  });

  /* ---- DOM updates ---- */
  var KEYS=['cagr','cumret','sharpe','sortino','calmar','vol','maxdd','avgddays','maxwins','maxlosses','kelly','rolllast'];
  function setCell(algo,key,m){
    var sel="[data-algo='"+CSS.escape(algo)+"'][data-metric='"+key+"']";
    document.querySelectorAll(sel).forEach(function(td){
      var v=m[key]; td.textContent=FMT[key](v); td.setAttribute('data-v',(v==null||isNaN(v))?'-Infinity':v);
      td.classList.remove('pos','neg'); var c=cls(key,v); if(c) td.classList.add(c);
    });
  }

  function apply(from,to){
    from=Math.round(from);to=Math.round(to);if(to<from){var t=from;from=to;to=t;}
    gFrom=from;gTo=to;
    var fi=document.getElementById('dateFrom'),ti=document.getElementById('dateTo');
    if(fi) fi.value=new Date(from).toISOString().slice(0,10);
    if(ti) ti.value=new Date(to).toISOString().slice(0,10);
    var rl=document.getElementById('rangeLabel');if(rl)rl.textContent=fdate(from)+'  –  '+fdate(to);

    DATA.algos.forEach(function(a){
      var pts=filt(a.portfolio,from,to);
      var m=metrics(pts, a.nPerYear, RF); KEYS.forEach(function(k){setCell(a.id,k,m);});
      var dmin=Infinity,dmax=-Infinity;
      if(m.equity.length){dmin=m.equity[0][0];dmax=m.equity[m.equity.length-1][0];}
      var hasData=dmin!==Infinity; if(!hasData){dmin=from;dmax=to;}
      var dcell=document.querySelector("[data-algo='"+CSS.escape(a.id)+"'][data-metric='days']");
      if(dcell){dcell.textContent=m.n;dcell.setAttribute('data-v',m.n);}
      var cov=document.querySelector(".cov[data-algo='"+CSS.escape(a.id)+"']");
      if(cov){cov.textContent=hasData?(m.n.toLocaleString()+' sessions · '+fdate(dmin)+' – '+fdate(dmax)):'no data in range';}
      /* breaks badge removed — no segment-break indicator shown */
      var eq=CH[a.id+'|equity'];if(eq){eq.setDomain(dmin,dmax);eq.render([{name:'Portfolio',color:'#2563eb',points:m.equity}]);}
      var ro=CH[a.id+'|rolling'];if(ro){ro.setDomain(dmin,dmax);ro.render([{name:'Rolling Sharpe',color:'#0d9488',points:m.rolling}]);}
    });
    var act=document.querySelector('th.sortable.asc, th.sortable.desc');if(act)act.click(),act.click();
  }

  /* ---- controls ---- */
  function wirePresets(){
    document.querySelectorAll('[data-range]').forEach(function(btn){
      btn.addEventListener('click',function(){
        document.querySelectorAll('[data-range]').forEach(function(b){b.classList.remove('active');});btn.classList.add('active');
        var k=btn.getAttribute('data-range');var to=DMAX,from=DMIN;
        if(k!=='all'){var d=new Date(DMAX);if(k==='1y')d.setFullYear(d.getFullYear()-1);else if(k==='6m')d.setMonth(d.getMonth()-6);
          else if(k==='3m')d.setMonth(d.getMonth()-3);else if(k==='ytd')d=new Date(Date.UTC(new Date(DMAX).getUTCFullYear(),0,1));from=Math.max(DMIN,d.getTime());}
        apply(from,to);
      });
    });
    var fi=document.getElementById('dateFrom'),ti=document.getElementById('dateTo');
    function fromInputs(){var a=fi.value?Date.parse(fi.value):DMIN,b=ti.value?Date.parse(ti.value):DMAX;
      document.querySelectorAll('[data-range]').forEach(function(x){x.classList.remove('active');});apply(a,b);}
    if(fi)fi.addEventListener('change',fromInputs);if(ti)ti.addEventListener('change',fromInputs);
    var rs=document.getElementById('resetRange');if(rs)rs.addEventListener('click',function(){apply(DMIN,DMAX);
      document.querySelectorAll('[data-range]').forEach(function(x){x.classList.remove('active');});
      var ab=document.querySelector("[data-range='all']");if(ab)ab.classList.add('active');});
    if(fi){fi.min=ti.min=new Date(DMIN).toISOString().slice(0,10);fi.max=ti.max=new Date(DMAX).toISOString().slice(0,10);}
  }

  document.querySelectorAll('.card-head').forEach(function(h){h.addEventListener('click',function(e){if(e.target.closest('a,button,input'))return;h.closest('.card').classList.toggle('collapsed');setTimeout(redrawAll,30);});});
  function setAll(c){document.querySelectorAll('.card').forEach(function(x){x.classList.toggle('collapsed',c);});if(!c)setTimeout(redrawAll,30);}
  var ea=document.getElementById('expandAll'),ca=document.getElementById('collapseAll');
  if(ea)ea.addEventListener('click',function(){setAll(false);}); if(ca)ca.addEventListener('click',function(){setAll(true);});
  function redrawAll(){for(var k in CH){CH[k].redraw();}}
  window.addEventListener('resize',function(){clearTimeout(window.__rz);window.__rz=setTimeout(redrawAll,150);});
  var box=document.getElementById('algoSearch');
  if(box)box.addEventListener('input',function(){var q=box.value.trim().toLowerCase();
    document.querySelectorAll('[data-algo-row]').forEach(function(el){el.style.display=(!q||el.getAttribute('data-algo-row').toLowerCase().indexOf(q)>-1)?'':'none';});});
  document.querySelectorAll('th.sortable').forEach(function(th){th.addEventListener('click',function(){
    var table=th.closest('table'),tbody=table.querySelector('tbody'),idx=Array.prototype.indexOf.call(th.parentNode.children,th);
    var asc=!th.classList.contains('asc');table.querySelectorAll('th.sortable').forEach(function(o){o.classList.remove('asc','desc');});th.classList.add(asc?'asc':'desc');
    var rows=Array.prototype.slice.call(tbody.querySelectorAll('tr'));
    rows.sort(function(a,b){var x=a.children[idx].getAttribute('data-v'),y=b.children[idx].getAttribute('data-v');var nx=parseFloat(x),ny=parseFloat(y);
      if(!isNaN(nx)&&!isNaN(ny)) return asc?nx-ny:ny-nx;x=(x||'').toLowerCase();y=(y||'').toLowerCase();return asc?(x>y?1:x<y?-1:0):(x<y?1:x>y?-1:0);});
    rows.forEach(function(r){tbody.appendChild(r);});});});
  var top=document.getElementById('toTop');
  if(top){window.addEventListener('scroll',function(){top.classList.toggle('show',window.scrollY>500);});top.addEventListener('click',function(){window.scrollTo({top:0,behavior:'smooth'});});}
  wirePresets();
  apply(DMIN,DMAX);
})();
"""

JS_STATIC = r"""
(function(){
  document.querySelectorAll('.card-head').forEach(function(h){h.addEventListener('click',function(e){if(e.target.closest('a,button,input'))return;h.closest('.card').classList.toggle('collapsed');});});
  function setAll(c){document.querySelectorAll('.card').forEach(function(x){x.classList.toggle('collapsed',c);});}
  var ea=document.getElementById('expandAll'),ca=document.getElementById('collapseAll');
  if(ea)ea.addEventListener('click',function(){setAll(false);}); if(ca)ca.addEventListener('click',function(){setAll(true);});
  var box=document.getElementById('algoSearch');
  if(box)box.addEventListener('input',function(){var q=box.value.trim().toLowerCase();
    document.querySelectorAll('[data-algo-row]').forEach(function(el){el.style.display=(!q||el.getAttribute('data-algo-row').toLowerCase().indexOf(q)>-1)?'':'none';});});
  document.querySelectorAll('th.sortable').forEach(function(th){th.addEventListener('click',function(){
    var table=th.closest('table'),tbody=table.querySelector('tbody'),idx=Array.prototype.indexOf.call(th.parentNode.children,th);
    var asc=!th.classList.contains('asc');table.querySelectorAll('th.sortable').forEach(function(o){o.classList.remove('asc','desc');});th.classList.add(asc?'asc':'desc');
    var rows=Array.prototype.slice.call(tbody.querySelectorAll('tr'));
    rows.sort(function(a,b){var x=a.children[idx].getAttribute('data-v'),y=b.children[idx].getAttribute('data-v');var nx=parseFloat(x),ny=parseFloat(y);
      if(!isNaN(nx)&&!isNaN(ny)) return asc?nx-ny:ny-nx;x=(x||'').toLowerCase();y=(y||'').toLowerCase();return asc?(x>y?1:x<y?-1:0):(x<y?1:x>y?-1:0);});
    rows.forEach(function(r){tbody.appendChild(r);});});});
  var top=document.getElementById('toTop');
  if(top){window.addEventListener('scroll',function(){top.classList.toggle('show',window.scrollY>500);});top.addEventListener('click',function(){window.scrollTo({top:0,behavior:'smooth'});});}
})();
"""

ICON_SEARCH = ('<svg width="15" height="15" viewBox="0 0 24 24" fill="none" '
               'stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="7"/>'
               '<line x1="21" y1="21" x2="16.5" y2="16.5"/></svg>')
ICON_CHEV = ('<svg class="chev" width="16" height="16" viewBox="0 0 24 24" fill="none" '
             'stroke="currentColor" stroke-width="2.2"><polyline points="6 9 12 15 18 9"/></svg>')
ICON_TOP = ('<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
            'stroke-width="2"><polyline points="6 15 12 9 18 15"/></svg>')


# --------------------------------------------------------------------------- #
# HTML helpers
# --------------------------------------------------------------------------- #
def _doc(title: str, body: str, data_json: Optional[str] = None, js: str = "") -> str:
    payload = ""
    if data_json is not None:
        safe = data_json.replace("</", "<\\/")
        payload = f"<script id='reportData' type='application/json'>{safe}</script>"
    return (
        "<!DOCTYPE html><html lang='en'><head><meta charset='utf-8'>"
        "<meta name='viewport' content='width=device-width,initial-scale=1'>"
        f"<title>{html.escape(title)}</title><style>{CSS}</style></head>"
        f"<body>{body}"
        f"<button id='toTop' class='toTop' title='Back to top'>{ICON_TOP}</button>"
        f"{payload}<script>{js}</script></body></html>"
    )


def _appbar(title: str, lede: str) -> str:
    lede_html = f"<div class='wrap'><p class='lede'>{lede}</p></div>" if lede else ""
    return (
        "<header class='appbar'><div class='wrap'><div>"
        "<div class='eyebrow'>Performance Report</div>"
        f"<h1>{html.escape(title)}</h1></div></div></header>"
        + lede_html
    )


def _meta_panel(items: list[tuple[str, str]]) -> str:
    cells = "".join(
        f"<div class='cell'><div class='k'>{html.escape(k)}</div>"
        f"<div class='v'>{v}</div></div>"
        for k, v in items
    )
    return f"<div class='meta'>{cells}</div>"


def _date_controls() -> str:
    return (
        "<div class='daterow'><span class='lab'>Range</span>"
        "<input type='date' id='dateFrom'><span class='lab'>to</span>"
        "<input type='date' id='dateTo'>"
        "<button class='preset active' data-range='all'>All</button>"
        "<button class='preset' data-range='1y'>1Y</button>"
        "<button class='preset' data-range='6m'>6M</button>"
        "<button class='preset' data-range='3m'>3M</button>"
        "<button class='preset' data-range='ytd'>YTD</button>"
        "<button class='preset' id='resetRange'>Reset</button>"
        "<span id='rangeLabel'></span></div>"
    )


def _canvas(algo: str, kind: str, height: int) -> str:
    return (f"<canvas class='chart-canvas' data-algo='{html.escape(str(algo))}' "
            f"data-chart='{kind}' data-h='{height}'></canvas>")


def _static_line_chart(dates: list[str], vals: list[float], *,
                       width: int = 560, height: int = 240,
                       pct: bool = True, color: str = "#2563eb",
                       label: str = "Portfolio") -> str:
    """Minimal static SVG line chart."""
    if len(vals) < 2:
        return '<div class="chart-empty">Not enough data to chart.</div>'
    scale = 100.0 if pct else 1.0
    unit = "%" if pct else ""
    y_dp = 0 if pct else 2
    vals_s = [v * scale for v in vals]
    floor = [0.0] if pct else []
    vmin = min(vals_s + floor)
    vmax = max(vals_s + floor)
    if vmin == vmax:
        vmax += 1.0
    pad = (vmax - vmin) * 0.08
    vmin -= pad; vmax += pad
    mL, mR, mT, mB = 52, 16, 18, 30
    iw, ih = width - mL - mR, height - mT - mB
    n = len(dates)

    def px(i: int) -> float:
        return mL + iw * i / (n - 1) if n > 1 else mL + iw / 2

    def py(v: float) -> float:
        return mT + ih * (1 - (v - vmin) / (vmax - vmin))

    parts = [f'<svg class="chart" viewBox="0 0 {width} {height}" '
             f'preserveAspectRatio="xMidYMid meet" role="img">']
    for t in range(5):
        v = vmin + (vmax - vmin) * t / 4
        y = py(v)
        parts.append(f'<line x1="{mL}" y1="{y:.1f}" x2="{width - mR}" y2="{y:.1f}" class="grid"/>')
        parts.append(f'<text x="{mL - 6}" y="{y + 3:.1f}" class="ytick">'
                     f'{v:.{y_dp}f}{unit}</text>')
    if vmin < 0 < vmax:
        y0 = py(0)
        parts.append(f'<line x1="{mL}" y1="{y0:.1f}" x2="{width - mR}" y2="{y0:.1f}" class="zero"/>')
    for i in (0, n // 2, n - 1):
        parts.append(f'<text x="{px(i):.1f}" y="{height - 8}" class="xtick">{dates[i]}</text>')
    pts_str = " ".join(f"{px(i):.1f},{py(v):.1f}" for i, v in enumerate(vals_s))
    parts.append(f'<polyline points="{pts_str}" fill="none" stroke="{color}" '
                 f'stroke-width="1.8" stroke-linejoin="round"/>')
    # legend
    parts.append(f'<rect x="{mL + 6}" y="{mT + 2}" width="10" height="3" rx="1.5" fill="{color}"/>')
    parts.append(f'<text x="{mL + 21}" y="{mT + 6}" class="legend">{html.escape(label)}</text>')
    parts.append("</svg>")
    return "".join(parts)


def _footer(load: LoadResult, cfg: Config, n_per_year: float) -> str:
    rf_pct = cfg.risk_free_annual * 100
    return (
        "<div class='wrap'><footer><h3>Formulas &amp; methodology</h3><ul>"

        # --- daily return ---
        f"<li><b>Daily return (r<sub>t</sub>)</b> = sum_mtm<sub>t</sub> / sum_allocation<sub>t</sub> &nbsp;|&nbsp; "
        f"sum_allocation = &Sigma;(raw_allocation &times; {cfg.allocation_scale}) across all users of that algo on date t. "
        "Returns are expressed as fractions; multiplied by 100 for display as %.</li>"

        # --- cumulative return ---
        "<li><b>Cumulative Return</b> = &Sigma; r<sub>t</sub> &nbsp;&nbsp;(simple additive sum, not compounded). "
        "The equity curve is the running cumulative sum of daily returns.</li>"

        # --- CAGR ---
        "<li><b>CAGR</b> = (1 + &Sigma;r<sub>t</sub>)<sup>N/days</sup> &minus; 1 &nbsp;|&nbsp; "
        "Geometric compound annualisation of the cumulative return. "
        "N = per-algo trading days/yr (distinct dates &divide; calendar years for that algo). "
        "Consistent with: total_pnl &divide; base_capital = cumulative return, annualised geometrically.</li>"

        # --- volatility ---
        f"<li><b>Annual Volatility</b> = std(r<sub>t</sub>, ddof=1) &times; &radic;N &nbsp;|&nbsp; "
        "Sample standard deviation of daily returns annualised by &radic;N.</li>"

        # --- sharpe ---
        f"<li><b>Sharpe Ratio</b> = (mean(r<sub>t</sub>) &minus; r<sub>f</sub>/N) / std(r<sub>t</sub>) &times; &radic;N "
        f"&nbsp;|&nbsp; r<sub>f</sub> = {rf_pct:.1f}% p.a. (3-yr avg 91-day Indian T-Bill, 2022&ndash;2025). "
        "Penalises both upside and downside volatility equally.</li>"

        # --- sortino ---
        f"<li><b>Sortino Ratio</b> = (mean(r<sub>t</sub>) &minus; r<sub>f</sub>/N) / DD &times; &radic;N "
        "&nbsp;|&nbsp; DD (downside deviation) = &radic;(mean(min(r<sub>t</sub>, 0)<sup>2</sup>)). "
        "Only penalises negative-return days; flat/zero days are treated as losses.</li>"

        # --- max drawdown ---
        "<li><b>Max Drawdown</b> = min(equity<sub>t</sub> &minus; peak<sub>t</sub>) "
        "where peak<sub>t</sub> = max(equity<sub>0..t</sub>). "
        "Measured on the additive equity curve; path-dependent metrics use the largest contiguous segment.</li>"

        # --- avg drawdown days ---
        "<li><b>Avg Drawdown Days</b> = mean length (trading days) of underwater episodes, "
        "where an episode starts when equity falls below its prior peak and ends when it fully recovers.</li>"

        # --- calmar ---
        "<li><b>Calmar Ratio</b> = CAGR / |Max Drawdown|. Higher is better. "
        "Uses the largest contiguous segment for the drawdown.</li>"

        # --- kelly ---
        "<li><b>Kelly Criterion</b> = W &minus; (1 &minus; W) / R &nbsp;|&nbsp; "
        "W = win rate (fraction of days with r<sub>t</sub> &gt; 0); "
        "R = avg winning return / avg |losing return|. "
        "Negative Kelly means do not size up on this setup.</li>"

        # --- rolling sharpe ---
        f"<li><b>Rolling Sharpe ({cfg.rolling_window}d window &asymp; 1 quarter):</b> "
        f"Sharpe formula (with r<sub>f</sub>) applied over a rolling {cfg.rolling_window}-trading-day window. "
        f"63 days &asymp; one calendar quarter — long enough to smooth daily noise and capture a full "
        "earnings/expiry cycle, short enough to flag a regime change within the same year. "
        f"The line begins after the first {cfg.rolling_window} days (warm-up). "
        "Stable values &gt; 1 indicate consistent risk-adjusted performance.</li>"

        # --- data quality ---
        "<li><b>Consecutive Wins/Losses</b>: longest streak of days with r &gt; 0 (wins) or "
        "r &le; 0 (losses) on the largest contiguous segment.</li>"

        "<li><b>Algo-19 broker exclusion:</b> rows where algo&nbsp;=&nbsp;19 AND broker is "
        "MasterTrust_Noren or mastertrust_dealer (case-insensitive) are dropped before "
        "any scaling, aggregation, or return calculation. No global broker filter is applied.</li>"

        "<li><b>0DTE filter:</b> for algos 1, 7, and 15 only rows where the "
        "<code>dte</code> column equals &ldquo;0DTE&rdquo; (case-insensitive) are used "
        "in calculations. All other DTE rows for those three algos are excluded. "
        "No DTE filter is applied to any other algo.</li>"

        f"<li><b>Processed data columns (Excel):</b> "
        "PNL% = (sum_mtm / sum_allocation) &times; 100 &nbsp;|&nbsp; "
        f"Base Capital = {cfg.base_capital:,} (fixed 1&nbsp;crore) &nbsp;|&nbsp; "
        "Absolute PNL = (sum_mtm / sum_allocation) &times; Base Capital. "
        "These columns show what you would have earned/lost on a 1-crore base capital.</li>"

        f"<li><b>Data quality:</b> rows in = {load.rows_in:,} | "
        f"algo-19 broker excluded = {load.rows_dropped_algo19_broker:,} | "
        f"non-0DTE rows dropped (algos 1/7/15) = {load.rows_dropped_dte:,} | "
        f"unparseable date = {load.rows_dropped_date} | "
        f"null mtm = {load.rows_dropped_mtm} | "
        f"null/&le;0 allocation = {load.rows_dropped_alloc} | "
        f"duplicate (user, algo, date) = {load.rows_dropped_dupe}.</li>"

        "</ul></footer></div>"
    )


# --------------------------------------------------------------------------- #
# Methodology / warning helpers
# --------------------------------------------------------------------------- #
def _data_years(b: MetricBundle) -> float:
    """Calendar years spanned by this algo's history."""
    if b.date_start is None or b.date_end is None or b.n_days == 0:
        return 0.0
    return (b.date_end - b.date_start).days / 365.25


def _methodology_note(cfg: Config) -> str:
    rf_pct = cfg.risk_free_annual * 100
    return (
        "<div class='method-note'>"
        "<h3>ℹ Methodology assumptions</h3>"
        "<ul>"
        f"<li><b>Risk-free rate:</b> {rf_pct:.1f}% p.a. — 3-year average of Indian 91-day"
        " T-Bill (2022&ndash;2025). Applied daily as r<sub>f&nbsp;daily</sub>"
        f" = {rf_pct:.1f}% &divide; N when computing Sharpe, Sortino, and Rolling Sharpe.</li>"
        f"<li><b>Rolling Sharpe window &mdash; {cfg.rolling_window} trading days"
        " (&asymp;&thinsp;1 calendar quarter):</b> One quarter is the shortest window that"
        " captures a full earnings/expiry cycle. It smooths day-to-day noise while still"
        " being short enough to flag a strategy regime change within the same year."
        " The Rolling Sharpe line is absent for the first"
        f" {cfg.rolling_window} trading days of each algo (warm-up period).</li>"
        "<li><b>* Return values (CAGR, Cumulative Return):</b> Marked with * to indicate"
        " these are <em>additive</em> fractional returns — daily r<sub>t</sub>"
        " = sum_mtm / sum_allocation, summed linearly for Cumulative Return and"
        " geometrically annualised for CAGR. They are <em>not</em> compounded"
        " mark-to-market returns.</li>"
        "<li><b>Short-history caveat:</b> CAGR and Calmar are unreliable for algos with"
        " &lt;&nbsp;1 year of data &mdash; the quant reference guide recommends"
        " &ge;&nbsp;3 years for Calmar. Algos flagged"
        " <span class='warn-badge'>&#9888; &lt;1&nbsp;yr</span>"
        " should be interpreted with this caveat in mind.</li>"
        "</ul></div>"
    )


# --------------------------------------------------------------------------- #
# Report renderer
# --------------------------------------------------------------------------- #
def render_report(
    per_algo: dict[str, MetricBundle],
    algo_series: dict[str, pd.Series],
    calendar: pd.DatetimeIndex,
    cfg: Config,
    load: LoadResult,
    n_per_year: float,
    date_min: pd.Timestamp,
    date_max: pd.Timestamp,
    *,
    dynamic: bool,
    data_json: Optional[str] = None,
) -> str:
    algos = sorted(per_algo.keys(), key=algo_sort_key)
    title = "Algo Performance — Aggregate Portfolio"

    if not algos:
        body = (
            _appbar(title, "No algos qualified.")
            + "<main class='wrap'><p class='empty-note'>Nothing to display.</p></main>"
            + _footer(load, cfg, n_per_year)
        )
        return _doc(title + (" (interactive)" if dynamic else ""), body,
                    data_json=data_json if dynamic else None,
                    js=JS if dynamic else JS_STATIC)

    date_ctrl = _date_controls() if dynamic else ""
    toolbar = (
        "<div class='toolbar'><div class='wrap'>"
        f"<label class='search'>{ICON_SEARCH}"
        "<input id='algoSearch' type='search' placeholder='Filter algos by ID...'></label>"
        "<button class='btn' id='expandAll'>Expand all</button>"
        "<button class='btn' id='collapseAll'>Collapse all</button>"
        + date_ctrl + "</div></div>"
    )

    def sh(label: str) -> str:
        return f"<th class='num sortable'>{label} <span class='arr'>&#9650;&#9660;</span></th>"

    sum_rows = []
    for a in algos:
        b = per_algo[a]
        data_yrs_sum = _data_years(b)
        short_sum = data_yrs_sum < 1.0 and b.n_days > 0
        warn_cell = (
            f" <span class='warn-badge' "
            f"title='Only {data_yrs_sum:.1f} yr of data — CAGR &amp; Calmar unreliable'>"
            f"&#9888; &lt;1&nbsp;yr</span>"
        ) if short_sum else ""
        sum_rows.append(
            f"<tr data-algo-row='{html.escape(str(a))}'>"
            f"<td class='metric' data-v='{html.escape(str(a))}'>"
            f"{html.escape(str(a))}{warn_cell}</td>"
            + _cell(a, "cagr",   "cagr",          fmt_return, b)
            + _cell(a, "sharpe", "sharpe",         fmt_ratio,  b)
            + _cell(a, "sortino","sortino",         fmt_ratio,  b)
            + _cell(a, "maxdd",  "max_drawdown",    fmt_pct,    b)
            + _cell(a, "calmar", "calmar",          fmt_ratio,  b)
            + _cell(a, "cumret", "cumulative_return", fmt_return, b)
            + f"<td class='num' data-algo='{html.escape(str(a))}' "
              f"data-metric='days' data-v='{b.n_days}'>{b.n_days:,}</td>"
            + "</tr>"
        )

    summary = (
        "<section><h2>Summary — All algos "
        "<span class='muted' style='text-transform:none;font-weight:400'>"
        "(click a column to sort)</span></h2>"
        "<div class='panel'><div class='tbl-wrap'><table><thead><tr>"
        "<th class='sortable'>Algo <span class='arr'>&#9650;&#9660;</span></th>"
        + sh("CAGR") + sh("Sharpe") + sh("Sortino")
        + sh("Max DD") + sh("Calmar") + sh("Cum. Return") + sh("Sessions")
        + f"</tr></thead><tbody>{''.join(sum_rows)}</tbody></table></div></div></section>"
    )

    hint = ("<span class='chart-hint'>drag to zoom &middot; double-click to reset</span>"
            if dynamic else "")
    cards = []
    for a in algos:
        b = per_algo[a]
        data_yrs_card = _data_years(b)
        short_card = data_yrs_card < 1.0 and b.n_days > 0
        warn_tag = (
            f"<span class='warn-badge' "
            f"title='Only {data_yrs_card:.1f} yr of history — CAGR &amp; Calmar unreliable; "
            f"quant guide recommends &ge; 3 yr for Calmar'>"
            f"&#9888; &lt;1&nbsp;yr data</span>"
        ) if short_card else ""
        cov_txt = (f"{b.n_days:,} sessions &middot; {fmt_date(b.date_start)} "
                   f"&ndash; {fmt_date(b.date_end)}" if b.date_start else "no coverage")

        rows = "".join(
            f"<tr><td class='metric'>{lbl}</td>"
            + _cell(a, key, attr, fmt, b)
            + "</tr>"
            for key, lbl, attr, fmt in METRICS
        )

        if dynamic:
            eq_chart = _canvas(a, "equity", 240)
            ro_chart = _canvas(a, "rolling", 150)
        else:
            eq_chart = _static_line_chart(
                b.equity_dates, b.equity_vals, pct=True, color="#2563eb", label="Portfolio")
            ro_chart = _static_line_chart(
                b.rolling_dates, b.rolling_sharpe_series,
                pct=False, color="#0d9488", label="Rolling Sharpe", height=150)

        cards.append(
            f"<div class='card' data-algo-row='{html.escape(str(a))}'>"
            f"<div class='card-head'>{ICON_CHEV}"
            f"<span class='title'>Algo {html.escape(str(a))}</span>"
            f"{warn_tag}"
            f"<span class='cov' data-algo='{html.escape(str(a))}'>{cov_txt}</span></div>"
            "<div class='card-body'><div class='card-grid'>"
            "<div class='chart-box'>"
            f"<div class='chart-hd'><span class='lab'>Cumulative return (%)</span>{hint}</div>"
            + eq_chart
            + f"<div class='chart-hd' style='margin-top:12px'>"
              f"<span class='lab'>Rolling Sharpe ({cfg.rolling_window}d window)</span>"
              "<span class='chart-hint'>line begins after first window fills</span></div>"
            + ro_chart
            + "</div>"
            "<div class='tbl-wrap'><table><thead><tr>"
            "<th>Metric</th><th class='num'>Portfolio</th></tr></thead>"
            f"<tbody>{rows}</tbody></table></div>"
            "</div></div></div>"
        )

    if dynamic:
        lede = ("Per-algo track record: daily return = sum_mtm / sum_allocation "
                "(allocation scaled ×100). "
                "Use the date range or drag a chart to recompute for any sub-period.")
    else:
        lede = ""

    footer_html = _footer(load, cfg, n_per_year) if dynamic else ""

    body = (
        _appbar(title, lede)
        + toolbar
        + "<main class='wrap'>"
        + _meta_panel([
            ("Algos reported", str(len(algos))),
            ("Data range", f"{fmt_date(date_min)} &ndash; {fmt_date(date_max)}"),
        ])
        + _methodology_note(cfg)
        + summary
        + "<section><h2>Per-algo detail</h2>" + "".join(cards) + "</section>"
        + "</main>" + footer_html
    )

    if dynamic:
        return _doc(title + " (interactive)", body, data_json=data_json, js=JS)
    return _doc(title, body, data_json=None, js=JS_STATIC)


# --------------------------------------------------------------------------- #
# Orchestration
# --------------------------------------------------------------------------- #
def resolve_algos(available: list[str], algos_raw: Optional[str]) -> list[str]:
    if algos_raw is None:
        algos_raw = prompt_for_algos(available)
    raw = (algos_raw or "").strip()
    if raw == "" or raw.lower() == "all":
        return available
    requested = {x.strip() for x in raw.replace(";", ",").split(",") if x.strip()}
    selected = [a for a in available if a in requested]
    unknown = sorted(requested - set(available), key=algo_sort_key)
    if unknown:
        logger.warning("Ignoring algos not found in data: %s", ", ".join(unknown))
    if not selected:
        logger.warning("No valid algos selected; defaulting to ALL.")
        return available
    return selected


def prompt_for_algos(available: list[str]) -> str:
    print(f"\nAlgos found in the data ({len(available)}):")
    print("  " + ", ".join(available))
    try:
        return input("Which algos to report? Comma-separated IDs, or 'all':\n> ")
    except EOFError:
        logger.warning("Non-interactive session; defaulting to ALL algos.")
        return "all"


def generate_reports(cfg: Config) -> tuple[Path, Path, Path]:
    """Full pipeline: load → filter → aggregate → Excel → compute metrics → HTML."""
    cfg.output_dir.mkdir(parents=True, exist_ok=True)

    # 1. Prompt for input path if not supplied via CLI / env
    if cfg.input_path is None:
        try:
            raw_path = input("\nEnter path to CSV file (users_filtered_all.csv):\n> ").strip()
        except EOFError:
            raise ValueError("No input file specified. Pass --input or set ALGO_REPORT_INPUT.")
        if not raw_path:
            raise ValueError("No input file specified. Pass --input or set ALGO_REPORT_INPUT.")
        object.__setattr__(cfg, "input_path", Path(raw_path).expanduser().resolve())

    load = load_data(cfg)
    df_all = load.df   # processed rows with scaled allocation

    # 2. Select algos
    available = sorted(df_all["algo"].unique().tolist(), key=algo_sort_key)
    selected = resolve_algos(available, cfg.algos_raw)
    logger.info("Reporting %d of %d algos: %s",
                len(selected), len(available), ", ".join(selected))

    df = df_all[df_all["algo"].isin(selected)].copy()

    # 3. Build processed (aggregated) data
    df_proc = build_processed_data(df, base_capital=cfg.base_capital)

    # 4. Save "Processed data.xlsx"
    save_processed_excel(df, df_proc, cfg.processed_excel_path)

    # 5. Build calendar from ALL loaded data (so gaps are detected vs the market)
    calendar = build_calendar(df_all)
    n_per_year_global = trading_days_per_year(calendar)
    logger.info("Global trading days/year (derived): %.2f", n_per_year_global)

    date_min = df_proc["date"].min()
    date_max = df_proc["date"].max()

    # 6. Compute metrics per algo from the processed (aggregated) return series.
    #    N is computed PER ALGO from that algo's own trading dates so that:
    #    - 0DTE algos (1, 7, 15) — only trade on expiry days (~54/yr after SEBI
    #      restriction) — are annualised correctly against their own trading frequency.
    #    - Carry-forward algos (8, 19, …) — trade most market days (~240/yr) — use
    #      their own denser calendar.
    #    Using a global N would inflate CAGR/Sharpe for low-frequency algos.
    per_algo: dict[str, MetricBundle] = {}
    algo_series: dict[str, pd.Series] = {}
    algo_n_per_year: dict[str, float] = {}

    for algo, grp in df_proc.groupby("algo"):
        algo = str(algo)
        grp = grp.sort_values("date").set_index("date")

        # Per-algo N: derived from this algo's own distinct trading dates
        algo_calendar = pd.DatetimeIndex(sorted(grp.index.unique()))
        n_per_year = trading_days_per_year(algo_calendar)
        algo_n_per_year[algo] = n_per_year

        # daily fractional return = sum_mtm / sum_allocation
        ret = grp["sum_mtm"] / grp["sum_allocation"]
        ret.name = algo
        algo_series[algo] = ret
        b = compute_metrics(ret, calendar, cfg, n_per_year)
        per_algo[algo] = b
        logger.info(
            "Algo %s: %d days | N/yr=%.1f | CAGR=%.2f%% | Sharpe=%.2f | MaxDD=%.2f%%",
            algo, b.n_days, n_per_year,
            b.cagr * 100 if not _na(b.cagr) else float("nan"),
            b.sharpe if not _na(b.sharpe) else float("nan"),
            b.max_drawdown * 100 if not _na(b.max_drawdown) else float("nan"),
        )

    # 7. Build JSON payload for interactive report
    payload = {
        "params": {
            "nPerYear": n_per_year_global,   # kept for reference; per-algo N used in metrics
            "rollingWindow": cfg.rolling_window,
            "maxGapDays": cfg.max_gap_days,
            "allocScale": cfg.allocation_scale,
            "rfAnnual": cfg.risk_free_annual,
        },
        "dateMin": pd.Timestamp(date_min).strftime("%Y-%m-%d"),
        "dateMax": pd.Timestamp(date_max).strftime("%Y-%m-%d"),
        "calendar": [pd.Timestamp(d).strftime("%Y-%m-%d") for d in calendar],
        "algos": [
            {
                "id": a,
                "nPerYear": algo_n_per_year[a],   # per-algo N for correct annualisation
                "portfolio": series_points(algo_series[a]),
            }
            for a in sorted(per_algo.keys(), key=algo_sort_key)
        ],
    }
    data_json = json.dumps(payload, separators=(",", ":"))

    # 8. Write interactive HTML
    interactive_html = render_report(
        per_algo, algo_series, calendar, cfg, load, n_per_year, date_min, date_max,
        dynamic=True, data_json=data_json,
    )
    cfg.std_interactive_path.write_text(interactive_html, encoding="utf-8")
    logger.info("Wrote %s", cfg.std_interactive_path)

    # 9. Write static client HTML
    client_html = render_report(
        per_algo, algo_series, calendar, cfg, load, n_per_year, date_min, date_max,
        dynamic=False,
    )
    cfg.std_client_path.write_text(client_html, encoding="utf-8")
    logger.info("Wrote %s", cfg.std_client_path)

    return cfg.std_interactive_path, cfg.std_client_path, cfg.processed_excel_path


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def prompt_for_input_path(max_tries: int = 5) -> Path:
    for _ in range(max_tries):
        try:
            raw = input("Enter path to the input file (CSV or Excel), or 'q' to quit:\n> ")
        except EOFError:
            raise ValueError("No input path provided. Pass --input or set ALGO_REPORT_INPUT.")
        raw = raw.strip().strip('"').strip("'").strip()
        if raw.lower() in {"q", "quit", "exit"}:
            raise SystemExit("Cancelled by user.")
        if not raw:
            print("  Path was empty. Try again.")
            continue
        path = Path(raw).expanduser()
        if not path.exists():
            print(f"  File not found: {path}. Try again.")
            continue
        if path.is_dir():
            print(f"  That is a folder, not a file. Include the file name.")
            continue
        return path.resolve()
    raise ValueError("No valid input path provided.")


def prompt_for_report_name(default: str = "Algo_performance_std") -> str:
    """
    Prompt the user for a base report name.
    Returns the name to use (without extension).
    """
    try:
        raw = input(
            f"\nEnter output report name (press Enter for default '{default}'):\n> "
        ).strip()
    except EOFError:
        return default
    if not raw:
        return default
    # Sanitize: remove characters that are unsafe in filenames
    safe = "".join(c for c in raw if c.isalnum() or c in "._- ")
    safe = safe.strip().replace(" ", "_")
    return safe if safe else default


def parse_args(argv: Optional[list] = None) -> Config:
    p = argparse.ArgumentParser(description="Generate algo performance HTML reports.")
    p.add_argument("--input",
                metavar="PATH", help="Path to the CSV input file.")
    p.add_argument("--algos",
        metavar="IDS", help="Comma-separated algo IDs, or 'all'.")
    p.add_argument("--output-dir", default="./reports",
        metavar="DIR", help="Directory for output files.")
    p.add_argument("--rolling-window", type=int, default=63,
        metavar="N", help="Rolling Sharpe window in trading days.")
    p.add_argument("--risk-free-annual", type=float, default=0.065,
        metavar="R", help="Annual risk-free rate for Sharpe/Sortino.")
    p.add_argument("--max-gap-days", type=int, default=3,
        metavar="N", help="Max missing trading days before a segment break.")
    p.add_argument("--date-format", default=None,
        metavar="FMT", help="Explicit date format, e.g. %%d-%%m-%%Y.")
    p.add_argument("-v", "--verbose", action="store_true",
        help="Verbose logging.")
    ns = p.parse_args(argv)

    lvl = logging.DEBUG if ns.verbose else logging.INFO
    logging.getLogger().setLevel(lvl)

    return Config(
        input_path=ns.input or os.environ.get("ALGO_REPORT_INPUT"),
        algos_raw=ns.algos or os.environ.get("ALGO_REPORT_ALGOS"),
        output_dir=Path(ns.output_dir or os.environ.get("ALGO_REPORT_OUTDIR", "./reports")),
        rolling_window=ns.rolling_window,
        risk_free_annual=ns.risk_free_annual,
        max_gap_days=ns.max_gap_days,
        date_format=ns.date_format,
    )


def main(argv: Optional[list] = None) -> int:
    cfg = parse_args(argv)
    try:
        interactive, client, excel = generate_reports(cfg)
        logger.info("Done. Reports written to %s", cfg.output_dir)
        logger.info("  Interactive : %s", cfg.std_interactive_path)
        logger.info("  Client      : %s", cfg.std_client_path)
        logger.info("  Excel       : %s", cfg.processed_excel_path)
        return 0
    except (FileNotFoundError, ValueError) as exc:
        logger.error("%s", exc)
        return 2
    except Exception as exc:
        logger.exception("Unexpected failure: %s", exc)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
