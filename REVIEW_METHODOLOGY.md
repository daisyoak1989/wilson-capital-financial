# Monthly Financial Review — Methodology & House Rules

This file is the **source of truth** for how the monthly review is performed across
all Wilson Capital properties. It captures owner feedback that refines the generic
`multifamily-financial-review` skill. The analysis scripts (`analyze.py`,
`phase34.py`, `gl_dive.py`, `review_build.py`) and any future session must follow
these rules. Rules are split into **General** (apply to every property) and
**Property-specific** overrides.

---

## General rules (all properties)

### A. Scoping PM questions

1. **Questions are for the CURRENT REVIEW MONTH (MTD / PTD) only.**
   Do **not** generate questions for the property manager from YTD variances.
   YTD may be summarized for context, but action items / questions are
   month-specific. (e.g., if MTD vacancy is on track, do not ask about a YTD
   vacancy gap.)

2. **Do not raise a question for an item that is favorable to or in line with
   budget**, even if it swings materially month-over-month. Example: payroll can
   jump in a 3-paycheck month but if it is still at/under budget, it is not a
   question — note it at most.

3. **Gain / Loss to Lease** variance vs. budget is generally **expected** (market
   repricing). Do not raise it as a PM question; note only if material.

4. **Make-Ready / turnover costs**: before flagging an overage, **correlate with
   vacancy loss and recent move-in volume.** High move-in activity drives higher
   make-ready spend — that is expected, not a problem. Present the linkage
   ("make-ready up, consistent with N move-ins / falling vacancy") rather than a
   bare overage question.

5. **Tax, Insurance, and Mortgage / Debt Service**: these are tracked in a
   **separate portfolio reference table maintained across all 7 properties.**
   Compare actuals against that table yourself — do **not** ask the PM about them.
   They should be level month-to-month; a change is something to reconcile against
   the table, not a default PM question.

### A.6. Owner consulting fee (portfolio-wide)

The recurring **$3,000/month Consulting / Professional Fees** to vendor
**"Oak Real Estate Investment"** is the **owner's own consulting fee** and appears
across multiple properties (confirmed on Brio and Sommery). It is expected and
budgeted-in-substance — **never a PM question or related-party concern** on any
property. `review_build.py` suppresses it from the PM-question list portfolio-wide
(it still appears in the Section 2 flag table for transparency on non-Brio
properties).

### A.7. Auto-curation of PM questions

`review_build.py` no longer dumps one raw question per flag. For non-Brio
properties it runs every flag through a house-rule filter that: drops items
favorable to / in line with budget (A.2: within 15% **and** under $1,000 variance),
excludes tax/insurance/mortgage (A.5) and Gain/Loss to Lease (A.3), suppresses the
owner consulting fee (A.6), reframes Make-Ready with turnover context (A.4),
collapses the payroll/benefit cluster into a single pay-period note, and keeps
genuine data-quality signals — sign anomalies (B.6) and new-this-month $0 lines
(C.7). The **GL accrual-reversal confirmation (C.7) still requires the separate
`gl_dive.py` step** — the analyst runs it on the surviving flags before sending.

### B. Data-quality / categorization checks

6. **Positive value in a contra / bad-debt account** (e.g., "Bad Debt –
   Accelerated Rent" showing a positive amount) → flag it. A positive in a
   normally-negative account usually means either a **recovery** (confirm and
   possibly reclass to the recovery/income line, e.g., "Accelerated Rent") or a
   **miscategorization.** Ask which.

### C. GL accrual-reversal review  ← **primary expense check, do not skip**

7. For **every flagged expense deviation** (a spike OR a drop vs. budget, prior
   month, or trend), you **must open the GL** for that account and check for
   accrual mismatches before drawing a conclusion.

   - Normal accrual practice each month: **reverse** last month's accrual (a
     **credit** to the expense account) **and book the actual** (a **debit**).
     Credits/reversals are therefore *pervasive and normal* — a credit by itself
     is **not** a red flag.
   - **Red flags to look for:**
     - **(a) Prior-month accrual reversed (credit) but NO actual expense posted
       (no offsetting debit) that month** → expense is understated / the drop is
       artificial, not a real cost decrease.
     - **(b) Accrual not booked at all for the month** → expense missing entirely.
   - **Signatures** that should trigger the GL check: a net **credit (negative)**
     in an expense account; or an **abnormally low / zero month** for an otherwise
     consistent recurring expense (utilities, contracts, ILS, trash, etc.).
   - **Verify in the GL — never assume a low month is a genuine savings.**

   *Worked example (Brio, May 2026):* Internet Listing Services / Zillow
   (`54012-000`) — Feb & March accruals were reversed without an offsetting actual
   expense being booked, so those months understated the true ILS cost.

   > Note: the GL export currently contains **review-month transactions only**.
   > A reversal-without-offset that occurred in a *prior* month is detected from
   > the **T12 monthly trend** (an abnormally low/zero month for a steady expense),
   > then confirmed by pulling that account's GL for the month in question when
   > available. Use `gl_dive.py <code>` to inspect a section (it prints debit vs.
   > credit so reversal-vs-actual is visible).

---

## Property-specific overrides

### Brio (txbrio)

- **Consulting / Professional Fees (`58242-000`)**: the recurring **$3,000/month**
  to "Oak Real Estate Investment" is the **owner's own consulting fee** — expected
  and budgeted-in-substance. **Do not flag** as unbudgeted or as a related-party
  concern.
- Items confirmed worth calling out for Brio: Locator & Broker Referrals,
  Lease Cancellation Fee income at $0, Make-Ready (with the vacancy/move-in
  context per rule A.4).
- Open Brio data-quality items: Bad Debt – Accelerated Rent showing positive
  (rule B.6); ILS Feb/Mar accrual reversal (rule C.7).

---

## Maintained reference (to build)

- **Portfolio tax / insurance / mortgage table** across all 7 properties, used to
  self-check rule A.5 instead of asking the PM. (Not yet built — see project TODO.)
