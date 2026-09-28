# Indian News & Business Magazine Reader Survey — Design Document

**Purpose:** Collect NPS + customer-lifecycle data across Indian news/business magazine
readers generally (brand-agnostic), to feed CRM Module Part 2-4: NPS calculation,
promoter/detractor/passive segmentation, and customer lifecycle (acquisition/retention/
development) analysis, benchmarked against Indian vs. international subscription
promotional practices researched separately.

**Length:** 20 questions across 6 sections (~4-5 min to complete).
**Screening:** Q1 filters out non-readers before the rest of the survey.

---

## Section 1 — Screening & Demographics

**Q1.** Do you currently read, or have you read in the past, any Indian news or business
magazine (print or digital)? *(Multiple choice: Yes / No)*
— If a respondent answers "No," later analysis should exclude them; the form itself
doesn't hard-branch (kept simple to build), so this is filtered in the response sheet.

**Q2.** What is your age bracket? *(Multiple choice: Under 18 / 18-25 / 26-35 / 36-50 /
51-65 / 65+)*

**Q3.** What is your gender? *(Multiple choice: Male / Female / Prefer not to say / Other)*

**Q4.** Which part of India do you live in? *(Multiple choice: North India / South India /
East India / West India / Central India / Outside India)*

**Q5.** Which best describes you? *(Multiple choice: Student / Salaried professional /
Business owner or self-employed / Freelancer / Retired / Other)*

## Section 2 — Customer Acquisition

**Q6.** Which Indian news/business magazine(s) do you currently read or subscribe to?
*(Checkboxes: India Today / Outlook / Business Today / The Week / Forbes India /
Frontline / Open magazine / Other — write-in)*

**Q7.** How did you first discover this magazine? *(Multiple choice: Family or friend
recommendation / Social media / Online search / Newsstand or vendor / Promotional
offer or discount / Bundled with another subscription / Other)*

**Q8.** What was the primary reason you first started reading or subscribed?
*(Multiple choice: Content quality / Price or promotional offer / Habit or brand legacy /
Convenience of access / Influenced by family, friends, or colleagues / Other)*

## Section 3 — Customer Retention

**Q9.** How long have you been reading or subscribed to this magazine?
*(Multiple choice: Less than 6 months / 6 months-1 year / 1-3 years / 3-5 years /
More than 5 years)*

**Q10.** How do you currently access it? *(Multiple choice: Print only / Digital only /
Both print and digital)*

**Q11.** How thoroughly do you typically read an issue? *(Multiple choice: Read it
cover-to-cover / Skim most articles / Only read specific sections)*

**Q12.** Have you ever seriously considered cancelling or not renewing a subscription?
*(Multiple choice: Yes / No)*

**Q13.** If yes — what was the main reason? *(Multiple choice, optional: Price increase /
Decline in content quality / Lack of time to read / Switched to another source /
Not applicable)*

## Section 4 — Customer Development

**Q14.** Have you ever subscribed to more than one magazine from the same publisher
(e.g. added a second title after starting with the first)? *(Multiple choice: Yes / No)*

**Q15.** Would you be interested in a bundled offer (e.g. print + digital + a second title
at a discount)? *(Multiple choice: Yes / No / Maybe, depends on price)*

**Q16.** Which promotional offers would most likely make you upgrade or renew?
*(Checkboxes: Extra free months on renewal / Discounted multi-title bundle / Free gift
or merchandise / Referral rewards / Exclusive extra digital content / A price-lock
guarantee against future increases)*

## Section 5 — Net Promoter Score

**Q17.** On a scale of 0-10, how likely are you to recommend this magazine to a friend
or colleague? *(Linear scale 0-10; 0 = Not at all likely, 10 = Extremely likely)*
— **This is the core NPS question.** Promoters = 9-10, Passives = 7-8, Detractors = 0-6.

**Q18.** What is the main reason for the score you gave above? *(Paragraph text, open)*

**Q19.** What one thing would most improve your experience with this magazine?
*(Multiple choice: Lower price / Greater content depth / Faster digital delivery /
More regional/local coverage / Better customer service / Other)*

## Section 6 — International Comparison

**Q20.** Have you ever subscribed to an international publication (e.g. The Economist,
TIME, The New York Times, Financial Times)? *(Multiple choice: Yes, currently /
Yes, in the past / No — if Yes, an optional write-in follow-up asks which one.)*

---

## Notes on design choices

- **Brand-agnostic framing** lets one respondent answer about whichever magazine they
actually read, so the dataset naturally segments by title without needing separate
forms per magazine.
- **NPS sits near the end** (Q17-18) after lifecycle questions warm the respondent up —
standard survey practice, reduces straight-lining.
- **Q6 (magazine read)** is the key segmentation variable for Part 2's "3-6 segment-based
players" comparison — once responses come in, filter/group by this field.
- Kept to single-select/checkbox/scale/short-text only — no matrix/grid questions —
so the Apps Script build stays simple and the form works well on mobile, which is
where most responses will come from given the distribution channels.
