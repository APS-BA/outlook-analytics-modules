"""
Module 2, Case 4: purchase journey friction.

Source: generated and run in Better Analyst (app.betteranalyst.com), session
"Brand Salience Gap", 2026-09-01. Exported verbatim; only this header was added.
Input paths point to Better Analyst's sandbox: /home/user/workspace/1788251553425_Data_Visualization.xlsx, /home/user/workspace/outputs/case_4_purchase_journey_friction.png.
Change them to your local data folder before running.
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pathlib import Path

source_path = "/home/user/workspace/1788251553425_Data_Visualization.xlsx"
if not Path(source_path).is_file():
    raise FileNotFoundError(f"Source file not found: {source_path}")
df = pd.read_excel(source_path, sheet_name="Data 1")

visit_col = "How often do you visit the store to buy a magazine?"
time_col = "How much time do you spend on an average at store?"
accompany_col = "Who accompanies you for purchase?"
receipt_col = "How do you receive your magazines"
region_col = "Where do you stay"
buy_freq_col = "How often do you buy a magazine?"
required = [visit_col, time_col, accompany_col, receipt_col, region_col, buy_freq_col]
missing = [c for c in required if c not in df.columns]
if missing:
    raise KeyError(f"Missing required columns: {missing}; available columns={list(df.columns)}")

print(f"SOURCE: Data Visualization.xlsx | filtered rows={len(df)} | date/filter scope=None")
print("VARIABLES:", required)

def clean_text(s):
    return s.astype("string").str.strip().fillna("Missing").replace("", "Missing")

visit = clean_text(df[visit_col])
time_at_store = clean_text(df[time_col])
accompany = clean_text(df[accompany_col])
receipt = clean_text(df[receipt_col])
region = clean_text(df[region_col])
buy_freq = clean_text(df[buy_freq_col])

visit_map = {"Once in 6 months": 1, "Once in a month": 2, "Once in 2 weeks": 3, "Once in a week": 4, "Daily": 5}
time_map = {"10 minutes": 1, "20 minutes": 2, "30 minutes": 3, "An hour": 4, "More than hour": 5}
visit_score = visit.map(visit_map)
time_score = time_at_store.map(time_map)
valid = visit_score.notna() & time_score.notna()
friction_index = ((visit_score - 1) / 4 * 50) + ((time_score - 1) / 4 * 50)

analysis = pd.DataFrame({"Receipt channel": receipt, "Region": region, "Visit frequency": visit, "Time at store": time_at_store, "Accompaniment": accompany, "Buy frequency": buy_freq, "Visit score": visit_score, "Time score": time_score, "Purchase Friction Index": friction_index})
analysis_valid = analysis.loc[valid].copy()

print("VISIT FREQUENCY SPLIT:")
print(visit.value_counts().to_string())
print("TIME AT STORE SPLIT:")
print(time_at_store.value_counts().to_string())
print("RECEIPT CHANNEL SPLIT:")
print(receipt.value_counts().to_string())
print("REGION SPLIT:")
print(region.value_counts().to_string())
print("PURCHASE FREQUENCY SPLIT:")
print(buy_freq.value_counts().to_string())

receipt_summary = analysis_valid.groupby("Receipt channel", dropna=False)["Purchase Friction Index"].agg(Respondents="size", Mean_Friction="mean", Median_Friction="median").sort_values("Mean_Friction", ascending=False)
region_summary = analysis_valid.groupby("Region", dropna=False)["Purchase Friction Index"].agg(Respondents="size", Mean_Friction="mean", Median_Friction="median").sort_values("Mean_Friction", ascending=False)
receipt_region = analysis_valid.pivot_table(index="Receipt channel", columns="Region", values="Purchase Friction Index", aggfunc="mean")
receipt_region_n = analysis_valid.pivot_table(index="Receipt channel", columns="Region", values="Purchase Friction Index", aggfunc="size")

print("RECEIPT CHANNEL X FRICTION:")
print(receipt_summary.round(2).to_string())
print("REGION X FRICTION:")
print(region_summary.round(2).to_string())
print("RECEIPT CHANNEL X REGION - MEAN FRICTION INDEX:")
print(receipt_region.round(2).to_string())
print("RECEIPT CHANNEL X REGION - RESPONDENTS:")
print(receipt_region_n.to_string())
print("ACCOMPANIMENT SPLIT:")
print(accompany.value_counts().to_string())
print("AUDIT:", {"rows": int(len(df)), "valid_index_rows": int(valid.sum()), "excluded_rows": int((~valid).sum()), "visit_unmapped": int(visit_score.isna().sum()), "time_unmapped": int(time_score.isna().sum()), "overall_mean_friction": round(float(analysis_valid["Purchase Friction Index"].mean()), 2)})
print("INDEX FORMULA: PFI = 50*((visit score-1)/4) + 50*((time score-1)/4); visit and time each scored 1=lowest burden to 5=highest burden")

plt.rcParams.update({"figure.dpi":120, "savefig.dpi":200, "savefig.bbox":"tight", "font.family":"sans-serif", "font.sans-serif":["DejaVu Sans"], "font.size":10, "axes.titlesize":12, "axes.titleweight":"bold", "axes.spines.top":False, "axes.spines.right":False, "axes.grid":True, "axes.grid.axis":"y", "grid.color":"#E6E6E6", "axes.axisbelow":True, "legend.frameon":False})
fig, axes = plt.subplots(1, 3, figsize=(18, 7), constrained_layout=True)

p1 = receipt_summary["Mean_Friction"].sort_values(ascending=False)
_plot_values = p1
_pv = _plot_values; _n = len(_pv); _all_na = bool(_pv.isna().all()) if _n > 0 else True
print("PLOT_INPUT shape:", getattr(_pv, "shape", (_n,))); print("PLOT_INPUT preview:"); print(_pv.head(5))
if _n == 0 or _all_na: raise SystemExit("CHART_DATA_EMPTY: receipt channel friction")
axes[0].bar(p1.index.astype(str), p1.values, color="#1F4E79")
axes[0].set_ylim(0, 100); axes[0].set_title("Mean friction by receipt channel"); axes[0].set_ylabel("Purchase Friction Index (0–100)"); axes[0].tick_params(axis="x", rotation=25)
for x, v in enumerate(p1.values): axes[0].text(x, v + 2, f"{v:.1f}", ha="center", fontsize=9)

p2 = region_summary["Mean_Friction"].sort_values(ascending=False)
_plot_values = p2
_pv = _plot_values; _n = len(_pv); _all_na = bool(_pv.isna().all()) if _n > 0 else True
print("PLOT_INPUT shape:", getattr(_pv, "shape", (_n,))); print("PLOT_INPUT preview:"); print(_pv.head(5))
if _n == 0 or _all_na: raise SystemExit("CHART_DATA_EMPTY: region friction")
axes[1].barh(p2.index.astype(str), p2.values, color="#7A9EBD")
axes[1].set_xlim(0, 100); axes[1].set_title("Mean friction by region"); axes[1].set_xlabel("Purchase Friction Index (0–100)")
for y, v in enumerate(p2.values): axes[1].text(v + 2, y, f"{v:.1f}", va="center", fontsize=9)

p3 = receipt_region
_plot_values = p3.values.ravel()
_pv = _plot_values; _n = len(_pv); _all_na = bool(pd.isna(_pv).all()) if _n > 0 else True
print("PLOT_INPUT shape:", getattr(_pv, "shape", (_n,))); print("PLOT_INPUT preview:"); print(_pv[:5])
if _n == 0 or _all_na: raise SystemExit("CHART_DATA_EMPTY: receipt-region friction crosstab")
p3.plot(kind="bar", ax=axes[2], color=["#B7C3D0", "#7A9EBD", "#4F81BD", "#1F4E79", "#9BBB59"])
axes[2].set_ylim(0, 100); axes[2].set_title("Friction by channel and region"); axes[2].set_ylabel("Mean PFI"); axes[2].tick_params(axis="x", rotation=25)
axes[2].legend(title="Region", fontsize=8, title_fontsize=8, loc="upper left", bbox_to_anchor=(0, -0.22))

fig.suptitle("Case 4 — Purchase Journey Friction", fontsize=17, fontweight="bold")
fig.text(0.5, 0.01, "Source: Data Visualization.xlsx | n=11,109 | descriptive index; higher score means more frequent store visits and longer store time", ha="center", fontsize=9, color="#555555")
output_path = "/home/user/workspace/outputs/case_4_purchase_journey_friction.png"
fig.savefig(output_path, dpi=200, bbox_inches="tight")
plt.close(fig)
print(f"SAVED: {output_path}")
