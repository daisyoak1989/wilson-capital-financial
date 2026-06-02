---
name: multifamily-financial-review
description: >
  Reviews a multifamily property's monthly financial package (T12, General Ledger, Budget Comparison)
  to identify anomalies, budget variances, and items requiring owner attention. Full trend analysis
  and GL deep dives above the NOI line; lighter summary for below-the-line items. Produces a
  structured Monthly Review sheet with flagged items, budget variance highlights, and specific
  questions for the property manager.

  Use this skill whenever the user uploads or references a multifamily financial package, T12 with GL,
  monthly financial review, or property P&L. Trigger when user says: "review my property financials",
  "check the T12", "analyze the monthly numbers", "what should I ask my property manager",
  "review the budget variance", "flag anything unusual in the financials", or mentions reviewing
  Appfolio, Yardi, or RealPage exports with a T12, GL, and budget comparison.
---

# Multifamily Monthly Financial Review

## Purpose

Review a multifamily property's monthly financial package (T12, General Ledger, and Budget Comparison)
to identify anomalies, variances, and items requiring owner attention. The primary focus is on items
above the NOI line (income and operating expenses). Below-the-line items (capital expenditures, debt
service, non-operating) receive a lighter summary treatment.

The audience is a property owner reviewing their property manager's work. Frame findings as "here is
what to ask about" — not "here is what is wrong." Be diplomatic but thorough.

## House Rules (owner overrides — apply these in addition to the phases below)

These refinements come from the owner and take precedence over the generic guidance:

1. **PM questions are CURRENT-MONTH (MTD/PTD) only.** Do NOT turn YTD variances into
   questions for the property manager. YTD is summarized for context; the action
   items are month-specific.
2. **Do not question items that are favorable to / in line with budget**, even if
   they swing month-over-month (e.g., a payroll spike that is still under budget).
3. **Gain / Loss to Lease** variance vs. budget is generally **expected** (market
   repricing) — note only, do not ask.
4. **Make-Ready / turnover costs**: correlate with vacancy loss and recent move-in
   volume before flagging. High move-ins → higher make-ready is expected; present
   the linkage rather than a bare overage question.
5. **Tax, Insurance, Mortgage / Debt Service**: compare against a maintained
   portfolio reference table (across all properties) — do NOT ask the PM about them.
6. **Positive value in a contra / bad-debt account** (e.g., a positive "Bad Debt"
   line) → flag as a possible recovery or miscategorization; ask which.
7. **GL accrual-reversal check (primary expense check — do not skip):** for every
   flagged expense deviation (spike or drop), open the GL and look for accrual
   mismatches. Reversals (credits) are *normal* each month; the red flags are
   (a) a prior-month accrual reversed with **no offsetting actual** posted, or
   (b) an accrual **not booked at all**. An abnormally low/zero month for a steady
   recurring expense is the signature — verify in the GL; never assume a real
   savings.

## Prerequisites

The workbook must contain three tabs:
1. **T12** — Trailing 12-month income statement
2. **GL** (General Ledger) — Detailed transaction-level journal entries for the review period
3. **Budget Comparison** — Actual vs. budget for the current period and YTD

If any tab is missing, stop and tell the user which tab(s) are needed before proceeding.

---

## Tab Layouts (Expected Structure)

### T12 Tab

| Column | Content |
|--------|---------|
| A | GL code (e.g., 41000-000) |
| B | Line item name (indented to show hierarchy) |
| C–N (or similar) | Monthly amounts (oldest to newest, left to right). The rightmost data column is the review month. |
| Next col after months | Total (12-month sum) |
| Last col (optional) | Notes / flags |

- Header rows at top contain property name, statement type, period, and book (Accrual/Cash).
- One of the early rows (approx. row 5) contains month/year headers for each column.
- **CRITICAL: Do NOT hardcode specific month names.** The T12 is a rolling 12-month window. Read the
  header row to determine which months are present. The rightmost monthly column is always the review
  month. The column immediately to its left is the prior month. Calculate trailing averages from the
  available columns.
- Categories follow a standard chart of accounts: Income (40000s), Operating Expenses (50000s–60000s),
  Capital/Replacement (70000s), Non-Operating (80000s).
- Key subtotal rows to locate by scanning column B: Total Income, Total Operating Expenses,
  Net Operating Income / NOI, NOI After Replacements, Net Income.

### GL Tab

| Column | Content |
|--------|---------|
| A | Property code |
| B | Property Name |
| C | Date (may be serial number) |
| D | Period (may be serial number) |
| E | Person / Description |
| F | Control (journal entry ID) |
| G | Reference |
| H | Debit |
| I | Credit |
| J | Balance |

- GL codes appear as section headers in their own rows (column A has the code, column E has the account name).
- Transactions for that account follow beneath until the next GL code header.
- The GL code in column A of these headers matches the GL codes in column A of the T12.

### Budget Comparison Tab

| Column | Content |
|--------|---------|
| A | GL code |
| B | Line item name |
| C | PTD Actual |
| D | PTD Budget |
| E | Variance (Actual minus Budget) |
| F | % Variance |
| G | YTD Actual |
| H | YTD Budget |
| I | Variance (YTD) |
| J | % Variance (YTD) |
| K | Annual Budget |

- Same GL code and line item structure as T12.
- Header rows at top contain property name, report type, period, and book.

---

## Workflow

### Phase 1: Orient and Read Key Data

1. **Identify the review month.** Read the T12 header row that contains month names. The rightmost
   monthly data column is the review month. Note its column letter. The column to its left is the
   prior month. Count back from there for trailing averages. Do NOT assume specific month names.
   Always read them from the header.

2. **Read the T12 in full** using get_range_as_csv. Read in chunks if needed (500 rows at a time).
   Parse into a structured dataset: GL code, line item name, each monthly amount, and total.

3. **Identify the NOI row.** Scan column B for "Net Operating Income" or the GL code 69999-099.
   Everything above this row is above the line (the primary review focus). Everything below is
   below the line (summary treatment only).

4. **Read the Budget Comparison in full** using get_range_as_csv.

5. **Do NOT read the full GL yet.** The GL can be very large. Only read specific GL code sections
   when an anomaly is detected and a deep dive is needed.

---

### Phase 2: T12 Trend Analysis — Above the NOI Line

This phase covers ONLY items above the NOI line (Income and Operating Expenses).

For each line item above NOI, compare the review month (rightmost data column) against:
- The prior month (one column to the left)
- The trailing 3-month average (the 3 months immediately before the review month)
- The trailing 12-month average (all 12 months)

**Flag an item if ANY of these conditions are met:**

**Income Items (40000–49999 GL codes):**
- Review month vs. prior month swing > 20% AND > $1,000 absolute change → Moderate
- Review month vs. 3-month avg > 25% AND > $1,500 absolute change → Moderate
- Review month is negative for a revenue line (not a contra-revenue line) → High
- Review month is zero but 3-month avg is > $500 (revenue dropped to zero) → High
- One-Time Concessions > 150% of 3-month avg → Moderate
- Bad Debt > 200% of 3-month avg AND > $5,000 → High
- Vacancy Loss > 120% of 3-month avg AND > $5,000 → Moderate

**Operating Expense Items (50000–69999 GL codes):**
- Review month vs. prior month swing > 25% AND > $1,000 absolute change → Moderate
- Review month vs. 3-month avg > 30% AND > $1,500 absolute change → Moderate
- Review month is negative for an expense line (possible credit/reversal) → High
- Review month is zero but 3-month avg is > $500 → Moderate
- Single month > 50% of 12-month total → High
- Ad Valorem Taxes month differs > 10% from prior month → Moderate
- Insurance differs from prior month at all (should be level) → Moderate
- Management Fee / Total Income outside 2.5%–4.0% range → Moderate
- Any utility line > 150% of 3-month avg AND > $2,000 → Moderate

**Additional High-priority flags — never skip these regardless of dollar amount:**
- **Negative Other Rental Income:** Any non-primary income line (pet fees, parking, storage,
  laundry, utility reimbursements, late fees, month-to-month fees, etc.) that shows a negative
  value in the review month → High. These are structurally expected to be positive; a negative
  value almost always means a reversal, correction, or misposting.
- **Negative Expenses:** Any operating expense line (50000–69999) showing a negative value in
  the review month → High. This indicates a vendor credit, reversal, or accounting correction
  that should always be explained by the PM, regardless of size.

**Skipping rules:**
- Skip subtotal/total rows (rows where col B contains "Total" or GL code ends in -099, -098, -090, -999).
  Only flag individual line items.
- Skip items where the absolute dollar amount is immaterial (review month AND 3-month average both < $200),
  **except** for the negative Other Rental Income and negative expense flags above — those are
  never skipped on materiality grounds.
- Contra-revenue lines (concessions, vacancy loss, bad debt) are expected to be negative in the T12.
  Do not flag them as negative income. Flag them only if they spike beyond the thresholds above.

---

### Phase 3: Below-the-Line Summary

For items below the NOI line (GL codes 70000+: capital expenditures, replacements, debt service,
non-operating income/expenses):

1. **Provide a brief summary.** List each below-the-line line item, the review month amount, and the
   12-month total. No detailed trend analysis needed.

2. **Flag only large or unusual amounts.** If any single below-the-line item in the review month is
   > $5,000 AND > 200% of its 3-month average, OR > $10,000 in absolute value for a single month,
   flag it and investigate the GL for that item. Summarize what you find.

3. **Note the NOI-to-Net-Income bridge:** Total below-the-line items, so the owner can see how much
   NOI is consumed by capex, debt service, and non-operating items.

4. Debt service should be level month-to-month. Any change in debt service is worth noting in the
   summary, even if the amount doesn't cross the dollar thresholds above.

---

### Phase 4: Budget Comparison Analysis

Read the Budget Comparison tab. Analyze both PTD (period-to-date) and YTD (year-to-date) variances.

**PTD (Period-to-Date) Variances:**
- PTD unfavorable variance > 20% AND > $1,000 → Moderate
- PTD unfavorable variance > 50% AND > $2,500 → High

**YTD (Year-to-Date) Variances:**
- YTD unfavorable variance > 15% AND > $3,000 → Moderate
- YTD unfavorable variance > 30% AND > $5,000 → High

**Budget Pacing:**
- YTD actual is pacing > 15% ahead of pro-rated annual budget → Moderate
  Formula: (YTD Actual) > (Annual Budget / 12 × months elapsed × 1.15)

**Unfavorable means:**
- For Income lines: Actual < Budget (revenue shortfall)
- For Expense lines: Actual > Budget (overspending)

**Explicitly call out:**
- The top 5 largest PTD variances by absolute dollar amount (both favorable and unfavorable)
- The top 5 largest YTD variances by absolute dollar amount
- The top 5 largest PTD variances by percentage
- The top 5 largest YTD variances by percentage
- Any line item that appears in both PTD and YTD top lists (persistent issue, not just timing)

Skip subtotal rows and immaterial items (< $200).

---

### Phase 5: GL Deep Dive on Flagged Items

For every item flagged as **High** in Phase 2, Phase 3, or Phase 4:

1. Find the GL code from the T12 or Budget Comparison (column A).
2. Search the GL tab for that GL code. Read the section of the GL that corresponds to that code.
3. Identify the specific transactions causing the anomaly. Look for:
   - Unusually large single transactions
   - Journal entries that look like corrections or reversals
   - Duplicate entries
   - Vendor names that seem unusual or new
   - Transactions with descriptions suggesting one-time or non-recurring charges
4. Summarize findings for each GL code investigated.

For **Moderate** items, only go to the GL if the pattern is unclear from the T12 alone (e.g., a
spike that could be a timing issue vs. a real problem). Use judgment — the GL is large, so be
selective.

If the GL does not have enough detail to explain an anomaly, say so and recommend the owner
request backup documentation from the property manager. Do not fabricate findings.

---

### Phase 5.5: Restructure and Annotate the T12 Tab

Before creating the Monthly Review sheet, make the following structural changes to the T12 tab
and then annotate flagged rows. The goal is to make the T12 a self-contained working view —
budget context right next to the actuals, old months collapsed, and headers locked.

**Step 1: Insert Budget Columns**

Insert three new columns immediately after the last existing data column (which will become the
new column O area — insert after the review month column and before any existing Total/Notes
columns):

- **New Column O — PTD Budget:** Header = "PTD Budget". For each line item row, look up that
  row's GL code in the Budget Comparison tab and pull in the PTD Budget value (Budget Comparison
  column D). Leave blank for subtotal/total rows.

- **New Column P — YTD Actual:** Header = "YTD Actual". Pull the YTD Actual value from the
  Budget Comparison tab (Budget Comparison column G) for each matching GL code row.

- **New Column Q — YTD Budget:** Header = "YTD Budget". Pull the YTD Budget value from the
  Budget Comparison tab (Budget Comparison column H) for each matching GL code row.

Match rows by GL code (T12 column A = Budget Comparison column A). If a GL code has no match
in the Budget Comparison, leave those cells blank. Do not insert formulas — write the values
as static numbers so the T12 remains self-contained.

**Step 2: Group the First Six Months**

Group columns C through H (the six oldest monthly columns in the T12) using Excel's column
grouping feature. This collapses them into a single toggle so the owner can focus on the most
recent months by default without losing the older data. The group should be collapsible (not
hidden permanently).

**Step 3: Freeze Panes at C6**

Set a freeze pane starting at cell C6. This keeps the GL code column (A), line item name
column (B), and the header rows (rows 1–5, which contain property info and month labels)
locked in place while the owner scrolls right through the monthly columns or down through
the line items.

**Step 4: Annotate Flagged Rows**

After the structural changes above are complete, annotate each flagged line item:

- **Do NOT change any cell values.** Numbers stay exactly as they are.
- **Highlight column N** (the review month column) for every flagged row — this is the primary
  visual signal the owner sees when scanning the T12. Apply the fill to the column N cell on
  that row regardless of which specific condition triggered the flag:
  - High severity → red fill (`FF0000`)
  - Moderate severity → yellow fill (`FFFF00`)
- **Write the review note in column S** (two columns to the right of YTD Budget in column Q,
  leaving column R as a visual gap/separator). Column S header = "Review Notes". The note
  should be concise but specific:
  - What triggered the flag (e.g., "147% above 3-mo avg — $8,200 vs. $5,600 avg")
  - Severity level (High / Moderate)
  - One-line suggested question (e.g., "Ask PM: what drove the spike in HVAC repairs this month?")
- If column S already has content in that row, append on a new line rather than overwriting.
- Only annotate individual line item rows — not subtotal or total rows — **except** for the
  four summary rows described in "Summary Row Annotations" below.
- Also annotate the corresponding row in the Budget Comparison tab (same column N highlight +
  comment in Budget Comparison column P) when the same item is flagged via budget variance analysis.

**Special attention — Negative Other Rental Income:**

"Other Rental Income" refers to income line items in the 40000–49999 GL code range that are
not the primary rent roll (e.g., pet fees, parking income, storage fees, laundry income,
utility reimbursements, late fees, month-to-month fees). These are expected to be positive.
If any of these line items show a **negative value in column N** (the review month):
- Flag as **High** severity regardless of dollar amount — a negative here almost always
  indicates a reversal, correction, or misposting that warrants immediate follow-up
- Highlight column N red (`FF0000`)
- Write in column S: "NEGATIVE OTHER RENTAL INCOME — [amount]. Possible reversal or misposting.
  Ask PM: what caused this credit and has it been corrected?"
- Do not confuse these with contra-revenue lines (concessions, vacancy loss, bad debt) which
  are structurally expected to be negative.

**Special attention — Negative Expenses:**

For any operating expense line item (50000–69999 GL codes) where column N shows a **negative
value** (i.e., a credit against expense):
- Flag as **High** severity — negative expenses indicate a reversal, vendor credit, or
  accounting correction that should always be explained
- Highlight column N red (`FF0000`)
- Write in column S: "NEGATIVE EXPENSE — [line item name], [amount]. Likely a credit or
  reversal. Ask PM: what is this credit for and is it expected to recur?"
- Note: do not suppress this flag even if the dollar amount is small — the sign itself is
  the anomaly, not the magnitude.


**Summary Row Annotations — Variance Driver Comments**

In addition to the individual line-item annotations above, write summary comments in column S
for these four specific subtotal/total rows:

1. **Total Rental Income** (scan column B for "Total Rental Income" or similar)
2. **Total Income** (scan column B for "Total Income")
3. **Total Operating Expenses** (scan column B for "Total Operating Expenses")
4. **Net Operating Income** (scan column B for "Net Operating Income" or "NOI")

These comments serve a different purpose than the individual flags — they give the owner a
quick read on *what is driving the totals* relative to budget, without having to piece it
together from individual line items. Because they summarize rather than flag, use a visually
distinct format so the owner can tell at a glance that these are roll-up comments, not
anomaly flags.

**Do not highlight column N** for these rows (no red or yellow fill). The summary comment in
column S is the only annotation.

**Format — use this exact structure every time, for consistency:**

```
═══ [LINE ITEM NAME] — VARIANCE SUMMARY ═══
PTD: $[actual] vs. $[budget] ([+/-$variance], [+/-X.X%])
  Drivers: [line item 1] ([+/-$X]), [line item 2] ([+/-$X]), [line item 3] ([+/-$X])
YTD: $[actual] vs. $[budget] ([+/-$variance], [+/-X.X%])
  Drivers: [line item 1] ([+/-$X]), [line item 2] ([+/-$X]), [line item 3] ([+/-$X])
```

**Rules for populating the format:**

- **PTD line:** Pull PTD Actual (column C) and PTD Budget (column D) from the Budget
  Comparison tab for the matching row. Calculate the dollar variance and percentage.
- **YTD line:** Pull YTD Actual (column G) and YTD Budget (column H) from the Budget
  Comparison tab. Calculate the dollar variance and percentage.
- **Drivers:** For each of PTD and YTD separately, identify the top 3 individual line items
  (by absolute dollar variance) that roll up into that total and are contributing most to the
  overall variance. List them in descending order of absolute dollar impact. Use the line item
  name (not GL code) and show the signed dollar variance (positive = favorable for income,
  unfavorable for expenses; negative = the opposite). If fewer than 3 items have meaningful
  variances (> $500), list only the ones that do.
- **Sign convention:** For income rows, a positive variance means actual exceeded budget
  (favorable). For expense rows, a positive variance means actual exceeded budget (unfavorable).
  Use "+" and "-" signs to make the direction clear.
- **Formatting the cell:** Use currency format with no decimals ($#,##0) for dollar amounts
  and one decimal place for percentages. The "═══" border lines and "Drivers:" label help
  the owner visually distinguish these from the individual-flag notes above and below.

**Example** (for illustration only — never hardcode these values):

```
═══ TOTAL OPERATING EXPENSES — VARIANCE SUMMARY ═══
PTD: $52,300 vs. $48,000 (+$4,300, +9.0%)
  Drivers: Contract Services (+$2,100), Repairs & Maintenance (+$1,800), Utilities (+$900)
YTD: $198,500 vs. $192,000 (+$6,500, +3.4%)
  Drivers: Repairs & Maintenance (+$3,200), Contract Services (+$2,400), Insurance (+$1,100)
```

---

### Phase 6: Generate the Review Summary

Create a new sheet called **Monthly Review** in the workbook. Structure it as follows:

**Section 1: Executive Summary** (Rows 1–15 approx.)
- Property name and review period (read from T12 header — do NOT hardcode month names)
- Key metrics for the review month:
  - Total Income (actual)
  - Total Operating Expenses (actual)
  - NOI (actual)
  - NOI Margin (NOI / Total Income)
- Month-over-month change for each metric above (review month vs. prior month, both $ and %)
- YTD vs. Budget for Total Income, Total Expenses, NOI (pull from Budget Comparison)
- Overall assessment: One sentence on whether this month is generally on track, concerning, or
  requires immediate attention.

**Section 2: Flagged Items Table — Above the Line** (Starting approx. Row 18)

Table columns: GL Code | Line Item | Review Month | Prior Month | 3-Mo Avg | Budget (PTD) | Variance | % Var | Severity | Finding / Notes

- Sort by severity (High first), then by absolute dollar impact (largest first).
- Finding / Notes column should include:
  - What the anomaly is (e.g., "167% spike vs. 3-mo avg")
  - GL investigation findings if applicable
  - A recommended action or question for the owner

**Section 3: Below-the-Line Summary**

A concise table: GL Code | Line Item | Review Month | 12-Mo Total | Note
- Only include items with non-zero values.
- Add notes for any items that were flagged and GL-investigated in Phase 3.
- Show the NOI-to-Net-Income bridge (total below-the-line impact).

**Section 4: Budget Variance Highlights**

Two sub-tables:

*PTD Variance — Largest Items:* GL Code | Line Item | PTD Actual | PTD Budget | Variance ($) | Variance (%) | Favorable?

*YTD Variance — Largest Items:* GL Code | Line Item | YTD Actual | YTD Budget | Variance ($) | Variance (%) | Favorable?

Include the top items by both $ and % (per Phase 4 analysis). Flag any items appearing in both
PTD and YTD lists as "Persistent."

**Section 5: Questions for Property Manager**

A numbered list of specific, actionable questions the owner should ask the property manager based
on the findings. Each question should reference the specific line item, the amount, and why it
warrants follow-up. Aim for 5–10 questions, prioritized by severity.

### Formatting Guidelines for the Review Sheet

- Bold headers with dark background (#1F4E79) and white text for section headers
- Severity column: conditional formatting with red fill for High, yellow fill for Moderate
- Currency format: $#,##0 (no decimals for readability)
- Percentage format: 0.0%
- Freeze the header row of the flagged items table
- Auto-fit column widths for readability

### Formatting Guidelines for Source Tab Annotations (T12 and Budget Comparison)

- **Never alter cell values** — structural additions (inserted columns, grouping, freeze) and
  highlights/comments are the only changes made to source tabs
- **Column N is the primary highlight target** in the T12 — always apply the fill color to the
  column N cell on each flagged row, as this is the review month and the most visible column
- High severity → red fill (`FF0000`) on column N
- Moderate severity → yellow fill (`FFFF00`) on column N
- **Column S = "Review Notes"** in the T12 tab — this is where all annotation text goes.
  Column R is left blank as a visual separator between the budget columns and the notes column.
- Budget Comparison tab annotations go in **column P** of that tab (highlight the relevant
  variance cell + write the note in column P of the same row)
- Notes: plain text, no special formatting required; keep each note under 3 lines
- Annotate both the T12 and the Budget Comparison tabs when the same line item is flagged in both
- The inserted columns (O, P, Q) should have a light blue header fill to visually distinguish
  them from the original T12 monthly data columns
- Column S header ("Review Notes") should have the same dark blue header style (`1F4E79`,
  white text) to make it clearly part of the analysis layer, not the original data

---

## Important Notes

- **Never hardcode month names.** Always read the T12 header row to determine the current period.
  The T12 is a rolling window. The skill must work for any 12-month window.

- **Focus above the NOI line.** Income and operating expenses get the full trend + budget + GL
  treatment. Below-the-line items get a summary with GL dives only for large or unusual amounts.

- **Always cross-reference.** If a T12 line looks odd, check the Budget Comparison for that line.
  If flagged in both, it is higher priority.

- **Be specific in findings.** Do not just say "Repairs and Maintenance is high." Say which
  sub-line is driving it, by how much, and what the GL shows.

- **Call out large budget variances explicitly.** For the budget comparison, list the biggest
  misses by both dollar amount and percentage for PTD and YTD separately. This is a key deliverable.

- **Property taxes and insurance** are often accrued monthly at a flat rate. Any deviation from
  the flat monthly amount is worth noting, even if small.

- **Management fees** should correlate to income. Check the effective management fee rate
  (Mgmt Fee / Total Income).

- **Be specific in GL findings.** If you find the cause of a spike in the GL, name the vendor,
  the amount, and the transaction date. The owner needs enough information to ask an informed
  question of the property manager.
