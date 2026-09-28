# Outlook Analytics Modules

**Outlook Group live project · four business-analytics modules, each delivered as a
business report plus a structured Excel workbook**

![Excel](https://img.shields.io/badge/Excel-live%20formulas-217346)
![Python](https://img.shields.io/badge/Python-pandas%20%7C%20Matplotlib-3776AB)
![Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-survey%20automation-4285F4)

---

## Modules

| # | Module | Data | What I did | Headline |
|---|---|---|---|---|
| 1 | [Data visualisation](01-data-visualisation/) | Brand survey, 11,109 responses × 92 fields | 5 business cases (brand salience, channel strategy, subscription fit, purchase-journey friction, B2B advertisers), each with a composite index and multi-panel chart | Where groups barely differ, the report says so and treats the flatness as a finding |
| 2 | [Sales forecasting](02-sales-forecasting/) | 5 titles × 36 months (2020–22) | 3-month weighted moving average and linear trend per title; Q1 2023 forecast | Group volume −23.3% in 2021, +57.3% in 2022: net +20.7% over three years |
| 3 | [CRM & NPS](03-crm-nps/) | Own 20-question survey, 133 responses (123 valid) | Survey design, Google Form built by script, NPS and customer-lifecycle analysis, benchmarking of 6 Indian titles against 2 international subscription models | Overall NPS +11.4, ranging from +30.8 to −8.0 by title |
| 4 | [Cash flow analysis](04-cash-flow-analysis/) | Company balance sheet and P&L (no cash flow statement provided) | Derived the full cash flow statement, ratio analysis, FMCG/FMCD/BFSI comparison | FY2022 CFO −₹245.9 Cr bridged by borrowing; positive CFO from FY2023 |

## How each workbook is built

```
README → Raw data → Data dictionary → preparation and analysis sheet per case → Final results
```

All calculations are live formulas, so every number in a report traces back to
the raw data. Module 1 figures were cross-validated twice: in Python
(pandas/Matplotlib) and independently in an agentic AI analytics tool.

## Repository contents

```
01-data-visualisation/   report (.docx) + workbook (.xlsx)
02-sales-forecasting/    report (.docx) + workbook (.xlsx)
03-crm-nps/              report + workbook, plus:
    survey_design.md         the 20-question instrument and its rationale
    create_google_form.gs    Apps Script that builds the Google Form
    mail_merge.gs            Apps Script for personalised invites and reminders
    distribution_kit.md      channel-by-channel outreach templates
04-cash-flow-analysis/   report (.docx) + workbook (.xlsx)
```
