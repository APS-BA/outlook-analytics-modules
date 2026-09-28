"""
Module 2, Case 3: subscription channel and promo-offer fit.

Source: generated and run in Better Analyst (app.betteranalyst.com), session
"Brand Salience Gap", 2026-09-01. Exported verbatim; only this header was added.
Input paths point to Better Analyst's sandbox: /home/user/workspace/1788251553425_Data_Visualization.xlsx, /home/user/workspace/outputs/case_3_subscription_channel_promo_offer_fit.png.
Change them to your local data folder before running.
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import PercentFormatter
from pathlib import Path

source_path = "/home/user/workspace/1788251553425_Data_Visualization.xlsx"
if not Path(source_path).is_file():
    raise FileNotFoundError(f"Source file not found: {source_path}")
df = pd.read_excel(source_path, sheet_name="Data 1")

channel_col = "To subscribe to Outlook , which plan would you choose"
offer_col = "Which of the following promotional offer would to like to get"
income_col = "What is your monthly income range/budget?"
required = [channel_col, offer_col, income_col]
missing = [c for c in required if c not in df.columns]
if missing:
    raise KeyError(f"Missing required columns: {missing}; available columns={list(df.columns)}")

print(f"SOURCE: Data Visualization.xlsx | filtered rows={len(df)} | date/filter scope=None")
print("VARIABLES:", required)

def clean_text(s):
    return s.astype("string").str.strip().fillna("Missing").replace("", "Missing")

channel = clean_text(df[channel_col])
offer = clean_text(df[offer_col])
income = clean_text(df[income_col])

channel_counts = channel.value_counts()
channel_table = pd.DataFrame({"Responses": channel_counts, "Share": channel_counts / len(df)})

offer_counts = offer.value_counts()
offer_table = pd.DataFrame({"Responses": offer_counts, "Share": offer_counts / len(df)})

channel_income = pd.crosstab(channel, income)
channel_income["Total"] = channel_income.sum(axis=1)
channel_income_pct = channel_income.drop(columns="Total").div(channel_income["Total"], axis=0)

# Diagnostic pairing table: offer preference within each acquisition channel.
offer_channel = pd.crosstab(channel, offer)
offer_channel["Total"] = offer_channel.sum(axis=1)
offer_channel_pct = offer_channel.drop(columns="Total").div(offer_channel["Total"], axis=0)

print("ACQUISITION CHANNEL SPLIT:")
print(channel_table.to_string())
print("PROMOTIONAL OFFER SPLIT:")
print(offer_table.to_string())
print("CHANNEL X INCOME - COUNTS:")
print(channel_income.to_string())
print("CHANNEL X INCOME - ROW PERCENTAGES:")
print((channel_income_pct * 100).round(2).to_string())
print("OFFER X CHANNEL - ROW PERCENTAGES:")
print((offer_channel_pct * 100).round(2).to_string())

# Flatness diagnostics for descriptive interpretation.
channel_income_cells = channel_income_pct.drop(columns=[c for c in channel_income_pct.columns if c == "Total"], errors="ignore")
max_income_spread = float(channel_income_pct.max(axis=0).sub(channel_income_pct.min(axis=0)).max()) if not channel_income_pct.empty else float("nan")
max_offer_spread = float(offer_channel_pct.max(axis=0).sub(offer_channel_pct.min(axis=0)).max()) if not offer_channel_pct.empty else float("nan")
print("AUDIT:", {"rows": int(len(df)), "channel_missing": int((channel == "Missing").sum()), "offer_missing": int((offer == "Missing").sum()), "income_missing": int((income == "Missing").sum()), "max_income_column_spread": round(max_income_spread, 4), "max_offer_column_spread": round(max_offer_spread, 4)})
print("COMPOSITE INDEX: Channel Concentration Index = max acquisition-channel share")
print(f"CHANNEL CONCENTRATION INDEX: {channel_table['Share'].max():.4f}")

plt.rcParams.update({"figure.dpi":120,"savefig.dpi":200,"savefig.bbox":"tight","font.family":"sans-serif","font.sans-serif":["DejaVu Sans"],"font.size":10,"axes.titlesize":13,"axes.titleweight":"bold","axes.spines.top":False,"axes.spines.right":False,"axes.grid":True,"axes.grid.axis":"y","grid.color":"#E6E6E6","axes.axisbelow":True,"legend.frameon":False})
fig, axes = plt.subplots(1, 3, figsize=(18, 7), constrained_layout=True)

# Panel 1: channel split
p1 = channel_table["Share"].sort_values(ascending=False)
_plot_values = p1
_pv = _plot_values; _n = len(_pv); _all_na = bool(_pv.isna().all()) if _n > 0 else True
print("PLOT_INPUT shape:", getattr(_pv, "shape", (_n,)))
print("PLOT_INPUT preview:"); print(_pv.head(5))
if _n == 0 or _all_na: raise SystemExit("CHART_DATA_EMPTY: acquisition channel split")
axes[0].bar(p1.index.astype(str), p1.values, color="#1F4E79")
axes[0].yaxis.set_major_formatter(PercentFormatter(1.0))
axes[0].set_title("Preferred subscription channel")
axes[0].set_ylabel("Share of respondents")
axes[0].tick_params(axis="x", rotation=25)
for x, v in enumerate(p1.values): axes[0].text(x, v + 0.012, f"{v:.1%}", ha="center", fontsize=9)

# Panel 2: offer preference
p2 = offer_table["Share"].sort_values(ascending=True)
_plot_values = p2
_pv = _plot_values; _n = len(_pv); _all_na = bool(_pv.isna().all()) if _n > 0 else True
print("PLOT_INPUT shape:", getattr(_pv, "shape", (_n,)))
print("PLOT_INPUT preview:"); print(_pv.head(5))
if _n == 0 or _all_na: raise SystemExit("CHART_DATA_EMPTY: promotional offer split")
axes[1].barh(p2.index.astype(str), p2.values, color="#7A9EBD")
axes[1].xaxis.set_major_formatter(PercentFormatter(1.0))
axes[1].set_title("Preferred promotional offer")
axes[1].set_xlabel("Share of respondents")
for y, v in enumerate(p2.values): axes[1].text(v + 0.01, y, f"{v:.1%}", va="center", fontsize=9)

# Panel 3: channel x income, row percentages
p3 = channel_income_pct
_plot_values = p3.values.ravel()
_pv = _plot_values; _n = len(_pv); _all_na = bool(pd.isna(_pv).all()) if _n > 0 else True
print("PLOT_INPUT shape:", getattr(_pv, "shape", (_n,)))
print("PLOT_INPUT preview:"); print(_pv[:5])
if _n == 0 or _all_na: raise SystemExit("CHART_DATA_EMPTY: channel income crosstab")
p3.plot(kind="bar", ax=axes[2], color="#B7C3D0", edgecolor="#FFFFFF")
axes[2].yaxis.set_major_formatter(PercentFormatter(1.0))
axes[2].set_title("Income mix within each channel")
axes[2].set_ylabel("Within-channel share")
axes[2].set_xlabel("")
axes[2].tick_params(axis="x", rotation=30)
axes[2].legend(title="Monthly income/budget", fontsize=8, title_fontsize=8, loc="upper left", bbox_to_anchor=(0, -0.22))

fig.suptitle("Case 3 — Subscription Channel and Promo-Offer Fit", fontsize=17, fontweight="bold")
fig.text(0.5, 0.01, "Source: Data Visualization.xlsx | n=11,109 | descriptive survey evidence; near-uniform patterns are not treated as causal", ha="center", fontsize=9, color="#555555")
output_path = "/home/user/workspace/outputs/case_3_subscription_channel_promo_offer_fit.png"
fig.savefig(output_path, dpi=200, bbox_inches="tight")
plt.close(fig)
print(f"SAVED: {output_path}")
