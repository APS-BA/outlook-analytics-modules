"""
Module 2, Case 5: hidden B2B advertiser segment.

Source: generated and run in Better Analyst (app.betteranalyst.com), session
"Brand Salience Gap", 2026-09-01. Exported verbatim; only this header was added.
Input paths point to Better Analyst's sandbox: /home/user/workspace/1788251553425_Data_Visualization.xlsx, /home/user/workspace/outputs/case_5_hidden_b2b_advertiser_segment.png.
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
role_col = "How would you classify yourself?"
freq_col = "How frequently do you (or your organization) run digital ads on a monthly basis?"
platform_col = "Where do you do your magic and run digital ad campaigns? "
spend_col = "How much would you like to spend on the products from Digital channels per month?"
activity_col = "Which digital marketing activities would suit for THE OUTLOOK GROUP company?"
learn_col = "How did you learn about our company (THE OUTLOOK GROUP) website?"
required = [role_col, freq_col, platform_col, spend_col, activity_col, learn_col]
missing = [c for c in required if c not in df.columns]
if missing:
    raise KeyError(f"Missing required columns: {missing}; available columns={list(df.columns)}")
print(f"SOURCE: Data Visualization.xlsx | filtered rows={len(df)} | date/filter scope=None")
print("VARIABLES:", required)

def clean(s):
    return s.astype("string").str.strip().fillna("Missing").replace("", "Missing")

role = clean(df[role_col])
freq = clean(df[freq_col])
platform = clean(df[platform_col])
spend = clean(df[spend_col])
activity = clean(df[activity_col])
learn = clean(df[learn_col])

b2b_roles = {"Business Owner/Founder", "Digital Marketer (freelance/contract/independent)", "Digital Marketer (at an agency/organization/business/etc)"}
segment = role.map(lambda x: "B2B advertiser-relevant" if x in b2b_roles else ("Pure consumer / other" if x == "Others" else "Unclassified"))
seg_counts = segment.value_counts()
seg_table = pd.DataFrame({"Responses": seg_counts, "Share": seg_counts / len(df)})
role_table = pd.DataFrame({"Responses": role.value_counts(), "Share": role.value_counts() / len(df)})

b2b_mask = segment.eq("B2B advertiser-relevant")
b2b_platform = platform[b2b_mask].value_counts()
b2b_platform_table = pd.DataFrame({"Responses": b2b_platform, "Share within B2B": b2b_platform / b2b_mask.sum()})
b2b_freq = freq[b2b_mask].value_counts()
b2b_freq_table = pd.DataFrame({"Responses": b2b_freq, "Share within B2B": b2b_freq / b2b_mask.sum()})
b2b_spend = spend[b2b_mask].value_counts()
b2b_spend_table = pd.DataFrame({"Responses": b2b_spend, "Share within B2B": b2b_spend / b2b_mask.sum()})
b2b_activity = activity[b2b_mask].value_counts()
b2b_activity_table = pd.DataFrame({"Responses": b2b_activity, "Share within B2B": b2b_activity / b2b_mask.sum()})

# Full-sample cross-tabs requested for segment profiling.
platform_segment = pd.crosstab(platform, segment)
platform_segment["Total"] = platform_segment.sum(axis=1)
platform_segment_pct = platform_segment.drop(columns="Total").div(platform_segment["Total"], axis=0)
spend_segment = pd.crosstab(spend, segment)
spend_segment["Total"] = spend_segment.sum(axis=1)
spend_segment_pct = spend_segment.drop(columns="Total").div(spend_segment["Total"], axis=0)

print("SEGMENT SPLIT:")
print(seg_table.to_string())
print("ROLE DETAIL:")
print(role_table.to_string())
print("B2B PLATFORM PREFERENCE:")
print(b2b_platform_table.to_string())
print("B2B CAMPAIGN FREQUENCY:")
print(b2b_freq_table.to_string())
print("B2B MONTHLY DIGITAL SPEND PREFERENCE:")
print(b2b_spend_table.to_string())
print("B2B ACTIVITIES SUITABLE FOR OUTLOOK:")
print(b2b_activity_table.to_string())
print("PLATFORM X SEGMENT - COUNTS:")
print(platform_segment.to_string())
print("SPEND X SEGMENT - COUNTS:")
print(spend_segment.to_string())

b2b_share = float(b2b_mask.mean())
consumer_share = float(segment.eq("Pure consumer / other").mean())
active_campaign_share = float(freq[b2b_mask].isin(["1-5 campaigns per month", "6-10 campaigns per month", "11-20 campaigns per month", "20+ campaigns per month"]).mean())
print("AUDIT:", {"rows": int(len(df)), "b2b_rows": int(b2b_mask.sum()), "consumer_rows": int(segment.eq("Pure consumer / other").sum()), "unclassified_rows": int(segment.eq("Unclassified").sum()), "b2b_share": round(b2b_share, 4), "b2b_active_campaign_rate": round(active_campaign_share, 4)})

plt.rcParams.update({"figure.dpi":120, "savefig.dpi":200, "savefig.bbox":"tight", "font.family":"sans-serif", "font.size":10, "axes.titlesize":12, "axes.titleweight":"bold", "axes.spines.top":False, "axes.spines.right":False, "axes.grid":True, "axes.grid.axis":"y", "grid.color":"#E6E6E6", "axes.axisbelow":True, "legend.frameon":False})
fig, axes = plt.subplots(1, 3, figsize=(18, 7), constrained_layout=True)

p1 = seg_table["Share"] * 100
axes[0].bar(p1.index.astype(str), p1.values, color=["#1F4E79", "#A5A5A5", "#C9C9C9"][:len(p1)])
axes[0].set_title("Respondent segment split"); axes[0].set_ylabel("Share of respondents (%)"); axes[0].tick_params(axis="x", rotation=25)
for x, v in enumerate(p1.values): axes[0].text(x, v + 1, f"{v:.1f}%", ha="center", fontsize=9)

p2 = b2b_platform_table["Share within B2B"].sort_values(ascending=False) * 100
axes[1].barh(p2.index.astype(str), p2.values, color="#4F81BD")
axes[1].set_title("B2B preferred ad platforms"); axes[1].set_xlabel("Share within B2B segment (%)")
for y, v in enumerate(p2.values): axes[1].text(v + 1, y, f"{v:.1f}%", va="center", fontsize=9)

p3 = b2b_spend_table["Share within B2B"].sort_values(ascending=False) * 100
axes[2].bar(p3.index.astype(str), p3.values, color="#7A9EBD")
axes[2].set_title("B2B monthly digital-spend preference"); axes[2].set_ylabel("Share within B2B (%)"); axes[2].tick_params(axis="x", rotation=25)
for x, v in enumerate(p3.values): axes[2].text(x, v + 1, f"{v:.1f}%", ha="center", fontsize=9)

fig.suptitle("Case 5 — The Hidden B2B Advertiser Segment", fontsize=17, fontweight="bold")
fig.text(0.5, 0.01, "Source: Data Visualization.xlsx | B2B segment = business owners + digital marketers; descriptive survey evidence", ha="center", fontsize=9, color="#555555")
output_path = "/home/user/workspace/outputs/case_5_hidden_b2b_advertiser_segment.png"
fig.savefig(output_path, dpi=200, bbox_inches="tight")
plt.close(fig)
print(f"SAVED: {output_path}")
