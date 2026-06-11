# Algo Performance Reports

Generates **two standardized HTML reports** plus a **verification CSV** from your
`users_filtered_all.csv`. No data leaves your machine; you run this locally.

## What it produces

Both reports are the **same Standardized report** (per algo, on a 1cr book), in two forms:

| File | Use | Interactivity |
|------|-----|---------------|
| `Algo_performance_std_interactive.html` | **Internal.** Explore the numbers yourself. | Date range From/To + presets, client-side recompute, zoomable charts, hover, search, sort, collapse. The full daily series is embedded so any sub-period can be recomputed in the browser. |
| `Algo_performance_std_client.html` | **Share with clients.** | **Static snapshot.** No date selection, no recompute, no zoom, and **no underlying data embedded**. Charts are fixed images; figures cover the full available history. (Search / sort / collapse remain, as conveniences — they don't expose data.) |

Both cover only the **algos you select** at runtime, and each algo is standardized to a
**1cr book (allocation == 100000)**, built **two ways** side-by-side: **Relay** and **Composite**.

A third file, `verification_data_used.csv`, lists the exact rows that fed every number
(see **Verification CSV** below).

> The interactive copy embeds the daily return series so it can recompute locally.
> The client copy deliberately omits that data, so recipients only see the fixed snapshot.

## Run

```bash
pip install pandas numpy
python report_generator.py
```

With no flags, the script runs interactively and asks **two** things:

1. the **CSV path** (paste the full path with the file name), then
2. **which algos** to report — it prints the algo IDs found in your data and you
   enter a comma-separated list (e.g. `1,7,8`) or `all`.

For automation, pass them directly:

```bash
python report_generator.py --input "C:\data\users_filtered_all.csv" --algos "1,7,8" --output-dir ./reports
```

Or via environment variables (nothing hardcoded):

```bash
export ALGO_REPORT_INPUT=/path/to/users_filtered_all.csv
export ALGO_REPORT_ALGOS="1,7,8"   # or "all"
export ALGO_REPORT_OUTDIR=./reports
python report_generator.py
```

### Options

| Flag | Env | Default | Meaning |
|------|-----|---------|---------|
| `--input` | `ALGO_REPORT_INPUT` | — (prompts) | Path to the CSV |
| `--algos` | `ALGO_REPORT_ALGOS` | — (prompts) | Algo IDs to report, e.g. `1,7,8`, or `all` |
| `--output-dir` | `ALGO_REPORT_OUTDIR` | `./reports` | Where the HTML is written |
| `--rolling-window` | `ALGO_REPORT_ROLLING` | `63` | Rolling Sharpe window (trading days) |
| `--primary-allocation` | `ALGO_REPORT_PRIMARY_ALLOC` | `100000` | Standard "1cr" book size |
| `--risk-free-annual` | `ALGO_REPORT_RF` | `0.0` | Risk-free rate for Sharpe/Sortino |
| `--max-gap-days` | `ALGO_REPORT_MAX_GAP` | `3` | Missing trading days tolerated before a "break" |
| `--date-format` | `ALGO_REPORT_DATEFMT` | day-first | Explicit date format, e.g. `%d-%m-%Y`. Default auto-parses day-first (DD-MM-YYYY) |
| `--exclude-broker` | `ALGO_REPORT_EXCLUDE_BROKER` | `19:MasterTrust_Noren` | Per-algo broker drops `algo:broker[,algo:broker]`. Pass `""` to disable. Case-insensitive |
| `-v` | — | off | Verbose logging |

## The interactive copy

Self-contained (CSS + JS inline, no internet needed). Whoever opens it can:

- **Pick any date range** with the From/To pickers or the **All / 1Y / 6M / 3M / YTD**
  presets. Every metric (CAGR, Sharpe, drawdown, streaks, Kelly, …) and both charts
  recompute in the browser for the chosen window — the cumulative curve re-bases to the
  period start.
- **Zoom the charts**: drag across the equity or Rolling-Sharpe chart to zoom (this also
  drives the metrics), scroll to zoom, double-click to reset.
- **Hover** for the exact value on any date, **filter** algos by ID, **sort** the
  comparison table, and **expand/collapse** per-algo cards.

Each algo's charts and coverage label use **that algo's own data range** (not a shared
global range). The Rolling-Sharpe line starts ~`rolling_window` trading days after the
data begins, because that's how long the first window takes to fill — it shares the same
time axis as the cumulative chart so the warm-up is obvious.

## The client copy

The same layout and figures, rendered **static** for sharing:

- No date controls, no recompute, no chart zoom, no hover.
- **No daily series embedded** in the file.
- Charts are fixed SVG images covering the full available history per algo.

Dates are parsed **day-first** by default (matches DD-MM-YYYY broker exports); if yours
differ, pass `--date-format`. Ranges display like `01 JAN 2024 – 29 MAY 2026`.

## Verification CSV (`verification_data_used.csv`)

So you can hand-check the figures, this file lists the **exact rows that were actually
used** — for each algo, every row that fed Relay and/or Composite on its date. Rows that
were eligible but not used (e.g. a fallback user on a day a primary user traded) and
broker-excluded rows are **not** included.

Columns:

| Column | Meaning |
|--------|---------|
| `algo`, `user_id`, `broker`, `date` | identity of the row |
| `mtm_all`, `allocation` | raw values from your file |
| `real_cap` | `allocation * 100` (true rupee capital) |
| `ret_fraction` | `mtm_all / real_cap` — the fractional daily return used in all math |
| `ret_pct` | `ret_fraction * 100` (== `mtm_all / allocation`) — easy manual check |
| `pool` | `primary` (==100k) or `fallback_xN` (an N× multiple) |
| `used_in_relay` | `True` if this user's return was the Relay pick **on that date** |
| `used_in_composite` | `True` if this user was in the Composite average **on that date** |

How to read it: filter to one `algo`. The **Composite** daily return = the average of
`ret_fraction` across the rows where `used_in_composite` is `True` that day. The **Relay**
daily return = the `ret_fraction` of the single row where `used_in_relay` is `True` that
day. Every row in the file was used by at least one of the two (so the two flags are
never both `False`).

## Excluding bad-data brokers

Some broker feeds are unreliable for specific algos. Drop them per algo:

```bash
python report_generator.py --exclude-broker "19:MasterTrust_Noren"
# multiple: --exclude-broker "19:MasterTrust_Noren,7:SomeBroker"
# disable:  --exclude-broker ""
```

Default is `19:MasterTrust_Noren` (override or disable as above). Matching is
case-insensitive and whitespace-trimmed. Excluded rows are logged, removed from **all**
metrics, and are **not** written to the verification CSV (their absence = excluded).

## Locked specification (the decisions this build encodes)

**Returns**
- One row per `(user, algo, date)`.
- `allocation` in the file = real allocation / 100  (`100000` == 1 crore).
- Daily return **fraction** = `mtm_all / (allocation * 100)` — used for all math.
- Percentages are display-only.

**Standardized report (per algo, 1cr book)**
- Primary pool: `allocation == 100000`.
- Fallback (only on dates with no exact-100k user): strict integer multiples
  (200000, 300000, …), auto-normalized by the fraction formula.
- **Relay**: greedy longest-forward-coverage — one user at a time, switch when their
  run ends; tie-break = lowest `user_id`.
- **Composite**: equal-weight average of all qualifying users each day.
- An algo with no 100k user (nor an integer multiple) is **skipped**, and logged.

**Metric conventions**
- Cumulative Return = simple sum of daily returns; equity curve is **additive**;
  drawdown measured on that additive curve.
- CAGR = mean daily return × N (simple-annualized).
- N (trading days/yr) derived from the data = distinct dates / calendar years.
- Sharpe / Sortino: risk-free = 0%; Sortino downside vs a 0 target.
- Win = return > 0; Loss = return <= 0 (**flat day counts as a loss**).
- Kelly = `W - (1-W)/R` (discrete).
- Path metrics (Calmar, Max Drawdown, Avg Drawdown Days, consecutive W/L) use the
  **largest contiguous segment** when real breaks exist; breaks are flagged.

**Data hygiene**
- Rows dropped (and logged): unparseable date, null `mtm_all`, null/<=0 `allocation`,
  duplicate `(user, algo, date)`.

## One tunable worth reviewing: `--max-gap-days`

A real algo will have occasional days where nobody traded. Treating *every* missing day
as a "break" shatters the series and makes drawdown/streak metrics unstable. This build
bridges gaps up to `max_gap_days` (default 3) and only counts a **real break** when an
algo is un-run longer than that. Raise/lower it to match how your desk operates.

## Caveats (stated assumptions)

- Fallback normalization assumes algo P&L scales **linearly** with capital (a 2cr book =
  2x a 1cr book). If lot-rounding breaks that, fallback returns are approximate.
- "Flat = loss" can deflate the average loss magnitude, which inflates Kelly. Interpret
  Kelly with that in mind.
- Pooling/averaging across users assumes the return process is comparable across them
  (stationarity). If users were moved off failing algos, survivorship can flatter an algo.
#   q u a n t _ r e p o r t _ g e n e r a t o r  
 