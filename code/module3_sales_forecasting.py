"""
Module 3: monthly series per title, 3-month weighted moving average and linear trend forecast for Q1 2023.

Source: generated and run in Better Analyst (app.betteranalyst.com), session
"Brand Salience Gap", 2026-09-03. Exported verbatim; only this header was added.
Input paths point to Better Analyst's sandbox: /home/user/workspace/1788251552534_Data_for_sales_forecasting_and_trend_line_analysis.xlsx, /home/user/workspace/outputs/sales_forecast_trends_all_titles.png.
Change them to your local data folder before running.
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pathlib import Path

source_path = "/home/user/workspace/1788251552534_Data_for_sales_forecasting_and_trend_line_analysis.xlsx"
output_path = "/home/user/workspace/outputs/sales_forecast_trends_all_titles.png"
if not Path(source_path).is_file():
    raise FileNotFoundError(f"Source file not found: {source_path}")

sheet_specs = {
    "Outlook Business": ("Outlook Business.", 12, "Month", [0, 1, 2]),
    "Outlook Traveller": ("Outlook Traveller", 12, "Month", [0, 1, 2]),
    "Outlook Money": ("Outlook Money", 12, "Month", [0, 1, 2]),
    "Outlook Hindi": ("Outlook Hindi", 26, "Week", [0, 1, 2]),
    "Outlook India": ("Outlook India", 52, "Week", [0, 1, 2]),
}
years = [2020, 2021, 2022]


def parse_indian_number(value):
    if pd.isna(value):
        return np.nan
    text = str(value).strip().replace(",", "")
    return pd.to_numeric(text, errors="coerce")


def load_title(sheet_name, issues_per_year, period_label):
    raw = pd.read_excel(source_path, sheet_name=sheet_name, header=None)
    records = []
    for year_idx, year in enumerate(years):
        base = year_idx * 4
        period_col = base
        opening_col = base + 1
        expiry_col = base + 2
        block = raw.iloc[2:, [period_col, opening_col, expiry_col]].copy()
        block.columns = ["issue", "Opening Subs", "Expiry"]
        block["issue"] = pd.to_numeric(block["issue"], errors="coerce")
        block["Opening Subs"] = block["Opening Subs"].map(parse_indian_number)
        block["Expiry"] = block["Expiry"].map(parse_indian_number)
        block = block.dropna(subset=["issue"]).copy()
        block["issue"] = block["issue"].astype(int)
        block["year"] = year
        block["month"] = np.ceil(block["issue"] * 12 / issues_per_year).astype(int).clip(1, 12)
        records.append(block[["year", "month", "Opening Subs", "Expiry"]])
    long_df = pd.concat(records, ignore_index=True)
    monthly = long_df.groupby(["year", "month"], as_index=False)[["Opening Subs", "Expiry"]].sum()
    monthly["date"] = pd.to_datetime(monthly["year"].astype(str) + "-" + monthly["month"].astype(str) + "-01")
    monthly = monthly.sort_values("date").reset_index(drop=True)
    monthly["churn_rate"] = monthly["Expiry"] / monthly["Opening Subs"]
    return monthly


title_data = {}
for title, (sheet, issues, period_label, _) in sheet_specs.items():
    title_data[title] = load_title(sheet, issues, period_label)
    if len(title_data[title]) != 36:
        raise ValueError(f"{title} did not produce 36 monthly rows; got {len(title_data[title])}")

# Build Outlook Group from the five cleaned monthly title series.
group = title_data["Outlook Business"][["date", "Opening Subs", "Expiry"]].copy()
group = group.rename(columns={"Opening Subs": "Outlook Business", "Expiry": "Business Expiry"})
for title in ["Outlook Traveller", "Outlook Money", "Outlook Hindi", "Outlook India"]:
    other = title_data[title][["date", "Opening Subs", "Expiry"]].rename(columns={"Opening Subs": title, "Expiry": f"{title} Expiry"})
    group = group.merge(other, on="date", how="inner")
gn = pd.DataFrame({"date": group["date"]})
gn["Opening Subs"] = group[["Outlook Business", "Outlook Traveller", "Outlook Money", "Outlook Hindi", "Outlook India"]].sum(axis=1)
gn["Expiry"] = group[["Business Expiry", "Outlook Traveller Expiry", "Outlook Money Expiry", "Outlook Hindi Expiry", "Outlook India Expiry"]].sum(axis=1)
gn["year"] = gn["date"].dt.year
gn["month"] = gn["date"].dt.month
gn["churn_rate"] = gn["Expiry"] / gn["Opening Subs"]
title_data["Outlook Group"] = gn


def recursive_wma(values, horizon=3):
    history = [float(x) for x in values]
    forecasts = []
    weights = np.array([1.0, 2.0, 3.0])
    for _ in range(horizon):
        pred = float(np.dot(np.array(history[-3:]), weights) / weights.sum())
        forecasts.append(pred)
        history.append(pred)
    return forecasts

results = {}
for title, data in title_data.items():
    actual = data["Opening Subs"].astype(float).to_numpy()
    x = np.arange(1, len(actual) + 1, dtype=float)
    slope, intercept = np.polyfit(x, actual, 1)
    trend_fit = intercept + slope * x
    ss_res = float(np.sum((actual - trend_fit) ** 2))
    ss_tot = float(np.sum((actual - actual.mean()) ** 2))
    r2 = 1 - ss_res / ss_tot if ss_tot else np.nan
    wma_fc = recursive_wma(actual, 3)
    trend_fc = [float(intercept + slope * k) for k in [37, 38, 39]]
    results[title] = {
        "data": data,
        "slope": float(slope),
        "intercept": float(intercept),
        "r2": float(r2),
        "trend_fit": trend_fit,
        "wma_forecast": wma_fc,
        "trend_forecast": trend_fc,
        "avg_churn_rate": float(data["churn_rate"].mean()),
        "sales_2020": float(data.loc[data["year"] == 2020, "Opening Subs"].sum()),
        "sales_2021": float(data.loc[data["year"] == 2021, "Opening Subs"].sum()),
        "sales_2022": float(data.loc[data["year"] == 2022, "Opening Subs"].sum()),
    }

print("SOURCE: Data for sales forecasting and trend line analysis.xlsx | rows used=5 source sheets | monthly output rows per series=36")
print("MONTH MAPPING: month = CEILING(issue * 12 / issues_per_year); Hindi=26 issues/year; India=52 issues/year")
print("FORECAST METHOD: recursive 3-month WMA with weights 1:2:3, most recent highest")
print("TITLE METRICS:")
for title, r in results.items():
    print(title, {"slope": round(r["slope"], 2), "intercept": round(r["intercept"], 2), "r_squared": round(r["r2"], 4), "avg_monthly_churn_rate": round(r["avg_churn_rate"], 6), "sales_2020": round(r["sales_2020"], 2), "sales_2021": round(r["sales_2021"], 2), "sales_2022": round(r["sales_2022"], 2), "wma_2023": [round(v, 2) for v in r["wma_forecast"]], "trend_2023": [round(v, 2) for v in r["trend_forecast"]]})

# Plot six small multiples: actuals, WMA history/forecast, and linear trend.
plt.rcParams.update({"figure.dpi": 120, "savefig.dpi": 200, "savefig.bbox": "tight", "font.family": "sans-serif", "font.size": 9, "axes.titlesize": 11, "axes.titleweight": "bold", "axes.spines.top": False, "axes.spines.right": False, "axes.grid": True, "axes.grid.axis": "y", "grid.color": "#E5E7EB", "axes.axisbelow": True, "legend.frameon": False})
fig, axes = plt.subplots(3, 2, figsize=(16, 14), constrained_layout=True)
axes = axes.flatten()
plot_values_all = []
for ax, (title, r) in zip(axes, results.items()):
    data = r["data"]
    actual = data["Opening Subs"].astype(float).to_numpy()
    dates = data["date"].to_numpy()
    wma_hist = pd.Series(actual).rolling(3).apply(lambda z: np.dot(z, np.array([1.0, 2.0, 3.0])) / 6, raw=True).to_numpy()
    future_dates = pd.date_range("2023-01-01", periods=3, freq="MS")
    ax.plot(dates, actual, color="#1F4E79", linewidth=1.7, label="Actual sales")
    ax.plot(dates, wma_hist, color="#E07A5F", linewidth=1.5, label="3M WMA")
    ax.plot(dates, r["trend_fit"], color="#2A9D8F", linewidth=1.5, linestyle="--", label="Linear trend")
    ax.plot(future_dates, r["wma_forecast"], color="#E07A5F", linewidth=1.8, marker="o", markersize=4, label="WMA forecast")
    ax.plot(future_dates, r["trend_forecast"], color="#2A9D8F", linewidth=1.8, linestyle="--", marker="o", markersize=4, label="Trend forecast")
    ax.axvline(pd.Timestamp("2023-01-01"), color="#6B7280", linewidth=0.9, linestyle=":")
    ax.set_title(f"{title} | Avg churn {r['avg_churn_rate']*100:.2f}%")
    ax.set_ylabel("Opening subscriptions")
    ax.tick_params(axis="x", rotation=30)
    ax.set_xlim(pd.Timestamp("2020-01-01"), pd.Timestamp("2023-03-31"))
    plot_values_all.extend(list(actual))
    plot_values_all.extend([v for v in wma_hist if not np.isnan(v)])
    plot_values_all.extend(r["wma_forecast"])
    plot_values_all.extend(r["trend_forecast"])
axes[0].legend(loc="upper left", fontsize=8, ncol=2)
fig.suptitle("Outlook Group Sales Forecasts: Actuals, 3-Month WMA, and Linear Trend", fontsize=17, fontweight="bold")
fig.text(0.5, 0.01, "Opening Subs treated as sales flow; Expiry treated as churn volume. Forecasts are Jan–Mar 2023.", ha="center", fontsize=9, color="#4B5563")
_plot_values = np.asarray(plot_values_all, dtype=float)
_n = _plot_values.size
_all_na = False if _n == 0 else bool(pd.isna(_plot_values).all())
print("PLOT_INPUT shape:", _plot_values.shape)
print("PLOT_INPUT preview:")
print(_plot_values[:5])
if _n == 0 or _all_na:
    raise SystemExit("CHART_DATA_EMPTY: plot values are empty or all-NaN. Rewrite the groupby/pivot/filter so plotted values are non-empty.")
fig.savefig(output_path, dpi=200, bbox_inches="tight")
plt.close(fig)
print(f"SAVED: {output_path}")

# Compact comparison for the requested closing summary.
print("YEAR COMPARISON:")
for title, r in results.items():
    direction = "grew" if r["sales_2022"] > r["sales_2020"] else "did not grow"
    print(title, {"2020_total": round(r["sales_2020"], 2), "2022_total": round(r["sales_2022"], 2), "change_pct": round((r["sales_2022"] / r["sales_2020"] - 1) * 100, 2), "classification": direction})
