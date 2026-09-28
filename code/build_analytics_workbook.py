"""
Builds the Excel workbook with the analysis tables.

Source: generated and run in Better Analyst (app.betteranalyst.com), session
"Brand Salience Gap", 2026-09-03. Exported verbatim; only this header was added.
Input paths point to Better Analyst's sandbox: /home/user/workspace/1788251552534_Data_for_sales_forecasting_and_trend_line_analysis.xlsx, /home/user/workspace/1788251553425_Data_Visualization.xlsx, /home/user/workspace/outputs/the_outlook_group_mba_business_analytics_workbook.xlsx.
Change them to your local data folder before running.
"""

import os
import json
from pathlib import Path
from copy import copy
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule

survey_path = "/home/user/workspace/1788251553425_Data_Visualization.xlsx"
forecast_path = "/home/user/workspace/1788251552534_Data_for_sales_forecasting_and_trend_line_analysis.xlsx"
output_path = "/home/user/workspace/outputs/the_outlook_group_mba_business_analytics_workbook.xlsx"
for source_path in [survey_path, forecast_path]:
    if not Path(source_path).is_file():
        raise ValueError(f"Missing source file: {source_path}")
    print(f"AUDIT: source_path={source_path}")

THEME = {"primary":"1F4E79","light":"D6E3F0","accent":"5B9BD5","green":"2E7D32","red":"C62828","gray":"666666","gold":"F57C00"}
wb = Workbook()
wb.remove(wb.active)

thin = Side(style="thin", color="D1D1D1")
medium = Side(style="medium", color=THEME["primary"])

def setup_sheet(ws, title, subtitle=None, gridlines=False):
    ws.sheet_view.showGridLines = gridlines
    ws.column_dimensions["A"].width = 3
    ws["B2"] = title
    ws["B2"].font = Font(name="Georgia", size=18, bold=True, color=THEME["primary"])
    if subtitle:
        ws["B3"] = subtitle
        ws["B3"].font = Font(name="Calibri", size=10, italic=True, color=THEME["gray"])
    ws.freeze_panes = "B5"

def style_header(ws, row, start_col, end_col):
    for c in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=THEME["primary"])
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=medium)

def style_block(ws, r1, r2, c1, c2):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(r, c)
            cell.border = Border(left=thin if c == c1 else Side(style=None), right=thin if c == c2 else Side(style=None), top=thin if r == r1 else Side(style=None), bottom=thin)
            if r > r1:
                cell.font = Font(name="Calibri", size=10, color="000000")
                cell.alignment = Alignment(vertical="top", wrap_text=True)

def write_table(ws, start_row, start_col, headers, rows, number_formats=None):
    for j, h in enumerate(headers, start=start_col):
        ws.cell(start_row, j).value = h
    style_header(ws, start_row, start_col, start_col + len(headers) - 1)
    for i, row in enumerate(rows, start=start_row + 1):
        for j, value in enumerate(row, start=start_col):
            if pd.isna(value): value = None
            if isinstance(value, (np.integer,)): value = int(value)
            if isinstance(value, (np.floating,)): value = float(value)
            ws.cell(i, j).value = value
            if number_formats and j - start_col < len(number_formats) and number_formats[j - start_col]:
                ws.cell(i, j).number_format = number_formats[j - start_col]
    style_block(ws, start_row, start_row + len(rows), start_col, start_col + len(headers) - 1)
    return start_row + len(rows)

def add_text(ws, row, text, col=2, bold=False, color="000000", size=10):
    ws.cell(row, col).value = text
    ws.cell(row, col).font = Font(name="Calibri", size=size, bold=bold, color=color)
    ws.cell(row, col).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[row].height = 32 if len(text) > 120 else 20

def add_chart(ws, title, cat_col, val_cols, header_row, first_data_row, last_data_row, anchor, colors=None):
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = title
    chart.y_axis.title = "Share / Value"
    chart.x_axis.title = "Category"
    cats = Reference(ws, min_col=cat_col, min_row=first_data_row, max_row=last_data_row)
    vals = Reference(ws, min_col=min(val_cols), max_col=max(val_cols), min_row=header_row, max_row=last_data_row)
    chart.add_data(vals, titles_from_data=True)
    chart.set_categories(cats)
    chart.height = 7.2
    chart.width = 13.5
    if colors:
        for s, color in zip(chart.series, colors):
            s.graphicalProperties.solidFill = color
            s.graphicalProperties.line.solidFill = color
    if len(chart.series) < len(val_cols):
        raise ValueError(f"Chart series count {len(chart.series)} below expected {len(val_cols)}")
    ws.add_chart(chart, anchor)
    print(f"AUDIT: chart title={title!r} series={len(chart.series)}")
    return chart

# Load survey source
survey = pd.read_excel(survey_path, sheet_name="Data 1")
print(f"AUDIT: survey rows={len(survey)} cols={len(survey.columns)}")

def clean_text(s):
    return s.astype("string").str.strip().fillna("Missing").replace("", "Missing")

def find_col(exact):
    if exact in survey.columns: return exact
    matches = [c for c in survey.columns if c.strip() == exact.strip()]
    if len(matches) != 1: raise ValueError(f"Could not resolve column {exact!r}; matches={matches}")
    return matches[0]

# Case 1
c1_recall = find_col("When you think of Indian magazine brand, which is the one that first pops up in your mind?")
c1_aware = find_col("Are you aware of Outlook magazine?")
c1_ext = find_col("Which of these extensions of outlook magazine are you aware of?")
c1_assoc = find_col("When you think about Outlook, what is the first thing that comes to your mind?")
c1_fam = find_col("How familiar are you with Outlook magazine? Rate on a 5-point scale where 1 means not at all familiar and 5 means very familiar")
c1_region = find_col("Where do you stay")
recall = clean_text(survey[c1_recall]).map(lambda x: "Outlook" if "outlook" in str(x).lower() else str(x))
aware = clean_text(survey[c1_aware]).str.lower().str.contains(r"^(yes|y|true|1|aware)$", regex=True, na=False)
region = clean_text(survey[c1_region])
fam = pd.to_numeric(clean_text(survey[c1_fam]).str.extract(r"(\d+)", expand=False), errors="coerce")
recall_counts = recall.value_counts()
c1_recall_rows = [[idx, int(v), float(v / len(survey))] for idx, v in recall_counts.items()]
c1_region_ct = pd.crosstab(region, aware)
c1_region_rows = []
for idx, r in c1_region_ct.iterrows():
    yes = int(r.get(True, 0)); no = int(r.get(False, 0)); total = yes + no
    c1_region_rows.append([idx, yes, no, total, yes / total if total else None])
c1_fam_counts = fam.value_counts().sort_index().reindex([1,2,3,4,5]).fillna(0)
c1_fam_rows = [[int(idx), int(v), float(v / fam.notna().sum())] for idx, v in c1_fam_counts.items()]

# Case 2
c2_mode = find_col("Which mode of magazines will you prefer")
c2_medium = find_col("Which is your preferred medium of keeping yourself updated?")
c2_pd = find_col("Do you prefer print magazines or digital magazines?")
c2_reason = find_col("In digital world, why will you prefer print magazine")
c2_age = find_col("In which age bracket do you fall?")
mode = clean_text(survey[c2_mode]); age = clean_text(survey[c2_age]); reason = clean_text(survey[c2_reason]); preferred_medium = clean_text(survey[c2_medium])
mode_order = ["Print", "Digital", "Digital and Print"]
c2_mode_rows = [[x, int((mode == x).sum()), float((mode == x).mean())] for x in mode_order]
age_order = ["Less than 18", "18-35", "35-50", "70 above"]
c2_age_rows = []
for a in age_order:
    mask = age == a; total = int(mask.sum())
    c2_age_rows.append([a] + [float(((mode == m) & mask).sum() / total) if total else None for m in mode_order] + [total])
c2_reason_counts = reason.value_counts()
c2_reason_rows = [[idx, int(v), float(v / len(survey))] for idx, v in c2_reason_counts.items()]

# Case 3
c3_channel = find_col("To subscribe to Outlook , which plan would you choose")
c3_offer = find_col("Which of the following promotional offer would to like to get")
c3_income = find_col("What is your monthly income range/budget?")
channel = clean_text(survey[c3_channel]); offer = clean_text(survey[c3_offer]); income = clean_text(survey[c3_income])
c3_channel_counts = channel.value_counts(); c3_offer_counts = offer.value_counts()
c3_channel_rows = [[idx, int(v), float(v / len(survey))] for idx, v in c3_channel_counts.items()]
c3_offer_rows = [[idx, int(v), float(v / len(survey))] for idx, v in c3_offer_counts.items()]
c3_income_ct = pd.crosstab(channel, income)
c3_income_rows = [[idx] + [int(c3_income_ct.loc[idx, col]) for col in c3_income_ct.columns] + [int(c3_income_ct.loc[idx].sum())] for idx in c3_income_ct.index]
c3_income_headers = ["Acquisition channel"] + list(c3_income_ct.columns) + ["Total"]
c3_pair = pd.crosstab(channel, offer).reindex(index=c3_channel_counts.index, columns=c3_offer_counts.index, fill_value=0)
c3_pair_rows = [[idx] + [float(c3_pair.loc[idx, col] / c3_pair.loc[idx].sum()) for col in c3_pair.columns] for idx in c3_pair.index]

# Case 4
c4_visit = find_col("How often do you visit the store to buy a magazine?")
c4_time = find_col("How much time do you spend on an average at store?")
c4_accomp = find_col("Who accompanies you for purchase?")
c4_receipt = find_col("How do you receive your magazines")
c4_region = find_col("Where do you stay")
c4_buy = find_col("How often do you buy a magazine?")
visit = clean_text(survey[c4_visit]); time_store = clean_text(survey[c4_time]); receipt = clean_text(survey[c4_receipt]); region4 = clean_text(survey[c4_region]); accomp = clean_text(survey[c4_accomp]); buy = clean_text(survey[c4_buy])
visit_map = {"Once in 6 months":1,"Once in a month":2,"Once in 2 weeks":3,"Once in a week":4,"Daily":5}
time_map = {"10 minutes":1,"20 minutes":2,"30 minutes":3,"An hour":4,"More than hour":5}
visit_score = visit.map(visit_map); time_score = time_store.map(time_map)
pfi = ((visit_score - 1) / 4 * 50) + ((time_score - 1) / 4 * 50)
valid4 = pfi.notna()
c4_receipt_sum = pd.DataFrame({"Respondents": pfi[valid4].groupby(receipt[valid4]).size(), "Mean PFI": pfi[valid4].groupby(receipt[valid4]).mean(), "Median PFI": pfi[valid4].groupby(receipt[valid4]).median()}).sort_values("Mean PFI", ascending=False)
c4_receipt_rows = [[idx, int(r["Respondents"]), float(r["Mean PFI"]), float(r["Median PFI"])] for idx, r in c4_receipt_sum.iterrows()]
c4_region_sum = pd.DataFrame({"Respondents": pfi[valid4].groupby(region4[valid4]).size(), "Mean PFI": pfi[valid4].groupby(region4[valid4]).mean(), "Median PFI": pfi[valid4].groupby(region4[valid4]).median()}).sort_values("Mean PFI", ascending=False)
c4_region_rows = [[idx, int(r["Respondents"]), float(r["Mean PFI"]), float(r["Median PFI"])] for idx, r in c4_region_sum.iterrows()]
c4_cross = pfi[valid4].groupby([receipt[valid4], region4[valid4]]).mean().unstack()
c4_cross_rows = [[idx] + [float(c4_cross.loc[idx, col]) if col in c4_cross.columns else None for col in c4_cross.columns] for idx in c4_cross.index]
c4_cross_headers = ["Receipt channel"] + list(c4_cross.columns)

# Case 5
c5_role = find_col("How would you classify yourself?")
c5_freq = find_col("How frequently do you (or your organization) run digital ads on a monthly basis?")
c5_platform = find_col("Where do you do your magic and run digital ad campaigns?")
c5_spend = find_col("How much would you like to spend on the products from Digital channels per month?")
c5_activity = find_col("Which digital marketing activities would suit for THE OUTLOOK GROUP company?")
c5_learn = find_col("How did you learn about our company (THE OUTLOOK GROUP) website?")
role = clean_text(survey[c5_role]); freq = clean_text(survey[c5_freq]); platform = clean_text(survey[c5_platform]); spend = clean_text(survey[c5_spend]); activity = clean_text(survey[c5_activity]); learn = clean_text(survey[c5_learn])
b2b_roles = {"Business Owner/Founder", "Digital Marketer (freelance/contract/independent)", "Digital Marketer (at an agency/organization/business/etc)"}
segment = role.map(lambda x: "B2B advertiser-relevant" if x in b2b_roles else ("Pure consumer / other" if x == "Others" else "Unclassified"))
seg_counts = segment.value_counts(); c5_seg_rows = [[idx, int(v), float(v / len(survey))] for idx, v in seg_counts.items()]
b2b = segment.eq("B2B advertiser-relevant")
def dist_rows(series):
    counts = series[b2b].value_counts()
    return [[idx, int(v), float(v / b2b.sum())] for idx, v in counts.items()]
c5_platform_rows = dist_rows(platform); c5_freq_rows = dist_rows(freq); c5_spend_rows = dist_rows(spend); c5_activity_rows = dist_rows(activity)
c5_learn_rows = dist_rows(learn)

# Forecasting source
forecast_specs = {"Outlook Business":("Outlook Business.",12),"Outlook Traveller":("Outlook Traveller",12),"Outlook Money":("Outlook Money",12),"Outlook Hindi":("Outlook Hindi",26),"Outlook India":("Outlook India",52)}
years = [2020, 2021, 2022]
def parse_num(v):
    if pd.isna(v): return np.nan
    return pd.to_numeric(str(v).strip().replace(",", ""), errors="coerce")
def load_title(sheet, issues):
    raw = pd.read_excel(forecast_path, sheet_name=sheet, header=None)
    records = []
    for yi, year in enumerate(years):
        base = yi * 4
        block = raw.iloc[2:, [base, base+1, base+2]].copy()
        block.columns = ["issue","Opening Subs","Expiry"]
        block["issue"] = pd.to_numeric(block["issue"], errors="coerce")
        block["Opening Subs"] = block["Opening Subs"].map(parse_num)
        block["Expiry"] = block["Expiry"].map(parse_num)
        block = block.dropna(subset=["issue"])
        block["issue"] = block["issue"].astype(int)
        block["year"] = year
        block["month"] = np.ceil(block["issue"] * 12 / issues).astype(int).clip(1,12)
        records.append(block[["year","month","Opening Subs","Expiry"]])
    long_df = pd.concat(records, ignore_index=True)
    monthly = long_df.groupby(["year","month"], as_index=False)[["Opening Subs","Expiry"]].sum()
    monthly["date"] = pd.to_datetime(monthly["year"].astype(str) + "-" + monthly["month"].astype(str) + "-01")
    monthly = monthly.sort_values("date").reset_index(drop=True)
    monthly["churn_rate"] = monthly["Expiry"] / monthly["Opening Subs"]
    return monthly

title_data = {title: load_title(sheet, issues) for title, (sheet, issues) in forecast_specs.items()}
for title, data in title_data.items():
    if len(data) != 36: raise ValueError(f"{title} monthly rows={len(data)}; expected 36")
group = title_data["Outlook Business"][["date","Opening Subs","Expiry"]].copy()
group = group.rename(columns={"Opening Subs":"Outlook Business","Expiry":"Business Expiry"})
for title in ["Outlook Traveller","Outlook Money","Outlook Hindi","Outlook India"]:
    other = title_data[title][["date","Opening Subs","Expiry"]].rename(columns={"Opening Subs":title,"Expiry":f"{title} Expiry"})
    group = group.merge(other, on="date", how="inner")
gn = pd.DataFrame({"date":group["date"]})
gn["Opening Subs"] = group[["Outlook Business","Outlook Traveller","Outlook Money","Outlook Hindi","Outlook India"]].sum(axis=1)
gn["Expiry"] = group[["Business Expiry","Outlook Traveller Expiry","Outlook Money Expiry","Outlook Hindi Expiry","Outlook India Expiry"]].sum(axis=1)
gn["year"] = gn["date"].dt.year
gn["month"] = gn["date"].dt.month
gn["churn_rate"] = gn["Expiry"] / gn["Opening Subs"]
title_data["Outlook Group"] = gn

def recursive_wma(values, horizon=3):
    history = [float(x) for x in values]
    out = []
    weights = np.array([1.0,2.0,3.0])
    for _ in range(horizon):
        pred = float(np.dot(np.array(history[-3:]), weights) / weights.sum())
        out.append(pred); history.append(pred)
    return out
forecast_rows = []
forecast_summary = {}
for title, data in title_data.items():
    actual = data["Opening Subs"].astype(float).to_numpy()
    x = np.arange(1, len(actual) + 1, dtype=float)
    slope, intercept = np.polyfit(x, actual, 1)
    fit = intercept + slope * x
    ss_res = float(np.sum((actual - fit) ** 2)); ss_tot = float(np.sum((actual - actual.mean()) ** 2))
    r2 = 1 - ss_res / ss_tot if ss_tot else np.nan
    wma = recursive_wma(actual, 3); trend = [float(intercept + slope * k) for k in [37,38,39]]
    forecast_summary[title] = {"slope":float(slope),"intercept":float(intercept),"r2":float(r2),"avg_churn":float(data["churn_rate"].mean()),"sales_2020":float(data.loc[data.year==2020,"Opening Subs"].sum()),"sales_2021":float(data.loc[data.year==2021,"Opening Subs"].sum()),"sales_2022":float(data.loc[data.year==2022,"Opening Subs"].sum()),"wma":wma,"trend":trend}
    for _, r in data.iterrows():
        forecast_rows.append([title, r["date"].to_pydatetime(), float(r["Opening Subs"]), float(r["Expiry"]), float(r["churn_rate"])])
    for i, dt in enumerate(pd.date_range("2023-01-01", periods=3, freq="MS")):
        forecast_rows.append([title, dt.to_pydatetime(), None, None, None])

# Overview
ov = wb.create_sheet("Overview")
setup_sheet(ov, "The Outlook Group — MBA Business Analytics Workbook", "Five survey-based business cases plus sales forecasting and trendline analysis | Generated 2026-09-03")
add_text(ov, 5, "Purpose: consolidate the five business cases and the sales forecasting analysis from this chat into one workbook. Survey cases use Data Visualization.xlsx only; forecasting uses Data for sales forecasting and trend line analysis.xlsx only.", bold=False)
add_text(ov, 7, "KEY INSIGHTS", bold=True, color=THEME["primary"], size=13)
insights = [
    "Case 1: Outlook awareness is materially higher than top-of-mind recall, creating a large salience gap; regional awareness is effectively flat.",
    "Case 2: Print, digital, and hybrid magazine preferences are almost evenly split, supporting balanced omnichannel investment.",
    "Case 3: Subscription channels and promotional offers are broadly dispersed; parallel acquisition testing is more defensible than consolidation.",
    "Case 4: Purchase Friction Index results are nearly flat by receipt channel and region; the survey does not identify a clear operational bottleneck.",
    "Case 5: 75.21% are in a business-owner or digital-marketer proxy segment, but this is a qualification opportunity, not confirmed advertising demand.",
    "Forecasting: Outlook India and Outlook Hindi grew most strongly from 2020 to 2022; Outlook Traveller was effectively flat."
]
for i, text in enumerate(insights, start=8): add_text(ov, i, "• " + text)
add_text(ov, 16, "CONTENTS", bold=True, color=THEME["primary"], size=13)
contents = ["Case 1","Case 2","Case 3","Case 4","Case 5","Forecasting","Forecast Data","Survey Data"]
for i, name in enumerate(contents, start=17):
    cell = ov.cell(i, 2); cell.value = name; cell.hyperlink = f"#'{name}'!B2"; cell.font = Font(name="Calibri", color="0563C1", underline="single")
ov.column_dimensions["B"].width = 28
for c in range(3,9): ov.column_dimensions[get_column_letter(c)].width = 18

# Case sheet builder
case1 = wb.create_sheet("Case 1"); setup_sheet(case1, "Case 1 — Brand Salience Gap", "Survey source: Data Visualization.xlsx | n=11,109")
add_text(case1,5,"Problem statement: assess whether Outlook awareness converts into strong top-of-mind recall and whether salience varies by region.")
add_text(case1,7,"Variables: top-of-mind brand recall; Outlook awareness; Outlook extension awareness; Outlook association; familiarity score; region.")
add_text(case1,8,"Composite index: Salience Gap Index = Outlook awareness rate − Outlook top-of-mind recall share = 49.38% − 17.27% = 32.1 percentage points.")
r = write_table(case1,10,2,["Brand","Responses","Share"],c1_recall_rows,[None,"#,##0","0.0%"])
add_chart(case1,"Top-of-mind recall share",2,[4],10,11,r,"F10",[THEME["primary"]])
r2 = write_table(case1,10,7,["Region","Aware","Not aware","Total","Awareness rate"],c1_region_rows,[None,"#,##0","#,##0","#,##0","0.0%"])
add_chart(case1,"Outlook awareness rate by region",7,[11],10,11,r2,"F25",[THEME["accent"]])
r3 = write_table(case1,23,2,["Familiarity score","Responses","Share"],c1_fam_rows,["0","#,##0","0.0%"])
add_chart(case1,"Familiarity distribution",2,[4],23,24,r3,"F25",[THEME["gold"]])
add_text(case1,32,"Interpretation: Outlook is first in recall only narrowly, while awareness is much higher, producing a 32.1-point salience gap. Regional awareness differs by less than one point and familiarity is almost perfectly uniform. The evidence supports a national distinctive-brand-memory priority, not a region-specific diagnosis.")
add_text(case1,34,"Recommended action: reinforce a small number of distinctive Outlook memory cues across the parent brand and extensions; measure aided awareness, unaided recall, familiarity, consideration, and recent purchase separately.")

case2 = wb.create_sheet("Case 2"); setup_sheet(case2, "Case 2 — Channel Strategy: The Even Three-Way Split", "Survey source: Data Visualization.xlsx | n=11,109")
add_text(case2,5,"Problem statement: determine whether print, digital, or hybrid magazine preference dominates and whether channel choice differs materially by age.")
add_text(case2,7,"Variables: preferred magazine mode; preferred information medium; print/digital preference; reason for print preference in a digital world; age bracket.")
add_text(case2,8,"Composite index: Channel Balance Index = 1 − largest mode share = 1 − 33.73% = 66.27%.")
r = write_table(case2,10,2,["Mode","Responses","Share"],c2_mode_rows,[None,"#,##0","0.0%"])
add_chart(case2,"Overall magazine mode split",2,[4],10,11,r,"F10",[THEME["primary"]])
r2 = write_table(case2,10,7,["Age bracket","Print","Digital","Digital and Print","Total"],c2_age_rows,[None,"0.0%","0.0%","0.0%","#,##0"])
add_chart(case2,"Mode preference by age",7,[8,9,10],10,11,r2,"F25",[THEME["primary"],THEME["accent"],THEME["gold"]])
r3 = write_table(case2,23,2,["Reason","Responses","Share"],c2_reason_rows,[None,"#,##0","0.0%"])
add_chart(case2,"Reasons for preferring print",2,[4],23,24,r3,"F25",[THEME["accent"]])
add_text(case2,32,"Interpretation: Print, digital, and hybrid preferences each represent roughly one-third of respondents. The age crosstab is also close to uniform, so the data does not support a sharp generational print-versus-digital divide or one dominant reason for print preference.")
add_text(case2,34,"Recommended action: maintain meaningful print, continue digital investment, and build integrated print-to-digital experiences. Use needs-based propositions and experiments rather than rigid age-based channel allocation.")

case3 = wb.create_sheet("Case 3"); setup_sheet(case3, "Case 3 — Subscription Channel and Promo-Offer Fit", "Survey source: Data Visualization.xlsx | n=11,109")
add_text(case3,5,"Problem statement: decide whether Outlook should consolidate subscription acquisition into one channel or invest in parallel channels with different offers.")
add_text(case3,7,"Variables: preferred subscription channel; preferred promotional offer; monthly income range/budget.")
add_text(case3,8,"Composite index: Parallel Channel Index = 1 − largest channel share = 1 − 25.68% = 74.32%.")
r = write_table(case3,10,2,["Acquisition channel","Responses","Share"],c3_channel_rows,[None,"#,##0","0.0%"])
add_chart(case3,"Acquisition channel split",2,[4],10,11,r,"F10",[THEME["primary"]])
r2 = write_table(case3,10,7,["Offer","Responses","Share"],c3_offer_rows,[None,"#,##0","0.0%"])
add_chart(case3,"Promotional offer preference",7,[9],10,11,r2,"F25",[THEME["gold"]])
r3 = write_table(case3,23,2,c3_income_headers,c3_income_rows,[None] + ["#,##0"] * (len(c3_income_headers)-1))
add_text(case3,23+len(c3_income_rows)+2,"Offer/channel diagnostic: the largest observed preference is Digital Version Free for Company Websites, Duffle Strolleys & Travelling Bags for Retailers, and Extended Subscription for Subscription Agencies. Differences are small, so these should be treated as test hypotheses.")
add_text(case3,36,"Interpretation: all four acquisition channels are close to 25%, all three offers are close to one-third, and each income group is close to one-fifth within each channel. The dataset does not support consolidation or a strong income-led channel strategy.")
add_text(case3,38,"Recommended action: maintain a parallel channel portfolio, use the website as the owned conversion channel, and run controlled offer tests with conversion, cost per acquisition, renewal, and cancellation as decision metrics.")

case4 = wb.create_sheet("Case 4"); setup_sheet(case4, "Case 4 — Purchase Journey Friction", "Survey source: Data Visualization.xlsx | n=11,109")
add_text(case4,5,"Problem statement: determine whether the magazine purchase journey creates materially different friction by receipt channel or region.")
add_text(case4,7,"Variables: store-visit frequency; time at store; purchase companion; receipt channel; region; magazine purchase frequency.")
add_text(case4,8,"Composite index: PFI = 50×(Visit Score−1)/4 + 50×(Time Score−1)/4, where higher values indicate more frequent visits and longer store time. Overall mean PFI = 50.3.")
r = write_table(case4,10,2,["Receipt channel","Respondents","Mean PFI","Median PFI"],c4_receipt_rows,[None,"#,##0","0.0","0.0"])
add_chart(case4,"Mean friction by receipt channel",2,[4],10,11,r,"F10",[THEME["primary"]])
r2 = write_table(case4,10,7,["Region","Respondents","Mean PFI","Median PFI"],c4_region_rows,[None,"#,##0","0.0","0.0"])
add_chart(case4,"Mean friction by region",7,[9],10,11,r2,"F25",[THEME["accent"]])
r3 = write_table(case4,23,2,c4_cross_headers,c4_cross_rows,[None] + ["0.0"] * (len(c4_cross_headers)-1))
add_text(case4,32,"Interpretation: mean friction is approximately 50 across Courier, Vendor, and Post, and regional means are also tightly clustered. The survey does not identify a clear operational bottleneck; longer store time may represent engagement rather than friction.")
add_text(case4,34,"Recommended action: preserve a balanced delivery portfolio, measure delivery time, on-time rate, failure, damage, complaints, and renewal by channel, and run small retail experiments such as clearer visibility, QR renewal, and simplified payment.")

case5 = wb.create_sheet("Case 5"); setup_sheet(case5, "Case 5 — The Hidden B2B Advertiser Segment", "Survey source: Data Visualization.xlsx | n=11,109")
add_text(case5,5,"Problem statement: identify a potential B2B advertising-sales audience separate from the consumer subscription funnel.")
add_text(case5,7,"Variables: role classification; digital ad frequency; ad platforms; monthly digital spend preference; suitable marketing activities; website discovery source.")
add_text(case5,8,"Composite index: B2B Opportunity Index = share classified as business owner or digital marketer = 8,355 / 11,109 = 75.21%. This is a qualification proxy, not confirmed advertiser demand.")
r = write_table(case5,10,2,["Segment","Responses","Share"],c5_seg_rows,[None,"#,##0","0.0%"])
add_chart(case5,"B2B advertiser-relevant segment",2,[4],10,11,r,"F10",[THEME["primary"]])
r2 = write_table(case5,10,7,["Platform","Responses","Share within B2B"],c5_platform_rows,[None,"#,##0","0.0%"])
add_chart(case5,"Preferred ad platforms within B2B segment",7,[9],10,11,r2,"F25",[THEME["accent"]])
r3 = write_table(case5,23,2,["Monthly spend preference","Responses","Share within B2B"],c5_spend_rows,[None,"#,##0","0.0%"])
add_chart(case5,"Monthly digital spend preference",2,[4],23,24,r3,"F40",[THEME["gold"]])
add_text(case5,48,"Interpretation: the survey surfaces a large potential advertiser-relevant audience, but role, platform, campaign frequency, spend, and activity preferences are all unusually even. There is no clear dominant platform or high-spend cluster, and the dataset should be treated as hypothesis-generating.")
add_text(case5,50,"Recommended action: launch a separate B2B sales pilot with a dedicated media-kit landing page, tiered test/mid-market/premium packages, LinkedIn-led but multi-platform positioning, and qualification questions on budget authority, organization, industry, objectives, and willingness to engage.")

# Forecasting sheet
fc = wb.create_sheet("Forecasting"); setup_sheet(fc, "Sales Forecasting and Trendline Analysis", "Source: Data for sales forecasting and trend line analysis.xlsx | 36 monthly points per series")
add_text(fc,5,"Method: monthly aggregation uses month = CEILING(issue × 12 / issues_per_year). Opening Subs is treated as sales flow; Expiry is treated as churn volume. Indian-style comma-formatted numbers were cleaned before calculation.")
summary_rows = []
for title, rsum in forecast_summary.items():
    summary_rows.append([title, rsum["sales_2020"], rsum["sales_2021"], rsum["sales_2022"], rsum["sales_2022"] / rsum["sales_2020"] - 1, rsum["slope"], rsum["intercept"], rsum["r2"], rsum["avg_churn"], rsum["wma"][0], rsum["wma"][1], rsum["wma"][2], rsum["trend"][0], rsum["trend"][1], rsum["trend"][2]])
headers = ["Title","2020 Sales","2021 Sales","2022 Sales","2020–22 Change","Slope / Month","Intercept","R-squared","Avg Monthly Churn","WMA Jan-23","WMA Feb-23","WMA Mar-23","Trend Jan-23","Trend Feb-23","Trend Mar-23"]
r = write_table(fc,7,2,headers,summary_rows,[None,"#,##0","#,##0","#,##0","0.0%","#,##0","#,##0","0.000","0.0%","#,##0","#,##0","#,##0","#,##0","#,##0","#,##0"])
add_text(fc,16,"Comparison: Outlook India, Outlook Hindi, Outlook Business, and Outlook Money grew from 2020 to 2022; Outlook Traveller was effectively flat. Outlook Group rose from 34,856,006 in 2020 to 42,238,053 in 2022, a 21.18% increase, after a sharp 2021 decline.")
add_text(fc,18,"Interpretation: India and Hindi are the strongest growth candidates. Business grew annually but has extremely low R-squared, Traveller is stable with low churn, and Money shows mild growth with low churn. Across titles, low R-squared values mean linear trendlines are directional rather than precise monthly forecasts.")
add_text(fc,20,"Recommended action: use WMA forecasts for near-term operating plans, linear trends for scenarios only, prioritize controlled growth tests for India and Hindi, protect Money’s low-churn base, stabilize Traveller, and investigate the large annual volatility before committing to major capacity or marketing shifts.")

# Forecast Data
fd = wb.create_sheet("Forecast Data"); setup_sheet(fd, "Forecast Data", "Visible monthly source-derived series and forecast rows", gridlines=True)
forecast_headers = ["Title","Month","Opening Subs / Sales Flow","Expiry / Churn Volume","Monthly Churn Rate"]
write_table(fd,5,2,forecast_headers,forecast_rows,[None,"yyyy-mm-dd","#,##0","#,##0","0.0%"])
fd.freeze_panes = "B6"
for col, width in {"B":24,"C":14,"D":24,"E":22,"F":20}.items(): fd.column_dimensions[col].width = width

# Survey Data visible source
sd = wb.create_sheet("Survey Data"); setup_sheet(sd, "Survey Data", "Source copy for traceability: Data Visualization.xlsx / Data 1", gridlines=True)
survey_headers = list(survey.columns)
for j, h in enumerate(survey_headers, start=2): sd.cell(5, j).value = str(h)
style_header(sd,5,2,1+len(survey_headers))
for i, record in enumerate(survey.to_dict("records"), start=6):
    for j, h in enumerate(survey_headers, start=2):
        v = record[h]
        if pd.isna(v): v = None
        if isinstance(v, (np.integer,)): v = int(v)
        if isinstance(v, (np.floating,)): v = float(v)
        sd.cell(i,j).value = v
sd.freeze_panes = "B6"
for j in range(2, min(1+len(survey_headers), 25) + 1): sd.column_dimensions[get_column_letter(j)].width = 18

# General formatting and widths
for ws in wb.worksheets:
    ws.sheet_view.zoomScale = 90
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None and cell.font.name is None:
                cell.font = Font(name="Calibri", size=10)
    if ws.title not in ["Survey Data","Forecast Data"]:
        for c in range(2, min(ws.max_column, 16) + 1):
            if ws.column_dimensions[get_column_letter(c)].width is None:
                ws.column_dimensions[get_column_letter(c)].width = 18

# Workbook-visible notes
for name in ["Case 1","Case 2","Case 3","Case 4","Case 5"]:
    ws = wb[name]
    ws["B3"].comment = None

# Audit and manifest
print(f"AUDIT: sheets={wb.sheetnames}")
print(f"AUDIT: survey_rows_written={len(survey)} forecast_rows_written={len(forecast_rows)}")
print(f"AUDIT: forecast_summary_titles={list(forecast_summary.keys())}")
print(f"AUDIT: key_metric=Overview!B8 and Forecasting!B8:P13")
wb.save(output_path)
exists = os.path.exists(output_path); size = os.path.getsize(output_path) if exists else 0
print(f"SAVED: {output_path} exists={exists} size_bytes={size}")
manifest = {"schemaVersion":1,"profile":"general","assertions":[
    {"id":"overview","label":"Overview exists","type":"required_sheet","sheet":"Overview"},
    {"id":"case1","label":"Case 1 exists","type":"required_sheet","sheet":"Case 1"},
    {"id":"case2","label":"Case 2 exists","type":"required_sheet","sheet":"Case 2"},
    {"id":"case3","label":"Case 3 exists","type":"required_sheet","sheet":"Case 3"},
    {"id":"case4","label":"Case 4 exists","type":"required_sheet","sheet":"Case 4"},
    {"id":"case5","label":"Case 5 exists","type":"required_sheet","sheet":"Case 5"},
    {"id":"forecast","label":"Forecasting exists","type":"required_sheet","sheet":"Forecasting"},
    {"id":"survey","label":"Survey source copy exists","type":"required_sheet","sheet":"Survey Data"},
    {"id":"forecast_data","label":"Forecast source-derived data exists","type":"required_sheet","sheet":"Forecast Data"},
    {"id":"overview_insights","label":"Overview contains key insights","type":"required_metric","value":{"sheet":"Overview","range":"B8"},"allowText":True},
    {"id":"forecast_summary","label":"Forecast summary exists","type":"required_metric","value":{"sheet":"Forecasting","range":"B8"},"allowText":True}
]}
print("__SB_WORKBOOK_SEMANTIC_MANIFEST__" + json.dumps(manifest, separators=(",",":")))
