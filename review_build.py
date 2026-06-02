"""Phase 5.5 + Phase 6: restructure & annotate T12, annotate Budget Comparison,
and generate the Monthly Review sheet.

NOTE: dollar thresholds reconstructed (skill text arrived with corrupted $ values).
NOTE: this Budget Comparison tab has NO GL-code column (col A = line item name),
so T12<->Budget matching is by line-item name, and Budget annotations use the
actual columns present (A name, B PTDact, C PTDbud, D PTDvar, E PTD%, F YTDact,
G YTDbud, H YTDvar, I YTD%, J Annual)."""

import re
import sys
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path

# Workbook path: pass the combined workbook as argv[1]; falls back to Brio May 2026.
DEFAULT_WB = Path(r"H:\.shortcut-targets-by-id\15cPT84Tcymc9b2jqyLRYJjcuEXIBOfOp\0. Business\Multi Family\!!! W.C. Portfolio\!!! Brio\Financial\202605\Brio_Financial_Analysis_202605.xlsx")
WB = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_WB
# Property-specific overrides only apply to the property they belong to.
PROPERTY_IS_BRIO = "brio" in str(WB).lower()

wb = openpyxl.load_workbook(WB)

# The combined workbook may contain text labels like "= Beginning Balance ="
# (common in the GL) that were stored with formula data_type. Excel strips
# these on open ("Removed Records: Formula"). Coerce any formula-typed cell
# back to a plain string so the re-saved workbook opens cleanly.
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for cell in row:
            if cell.data_type == "f" and isinstance(cell.value, str):
                cell.data_type = "s"

t12 = wb["T12"]
bc = wb["Budget Comparison"]

def num(v):
    return v if isinstance(v, (int, float)) else 0.0

def norm(s):
    return re.sub(r"\s+", " ", str(s).strip().lower()) if s else ""

# ---------- styles ----------
RED = PatternFill("solid", fgColor="FF0000")
YELLOW = PatternFill("solid", fgColor="FFFF00")
LIGHTBLUE = PatternFill("solid", fgColor="BDD7EE")
DARKBLUE = PatternFill("solid", fgColor="1F4E79")
WHITE_BOLD = Font(color="FFFFFF", bold=True)
WRAP = Alignment(wrap_text=True, vertical="top")

# =======================================================================
# Dynamic structure discovery — robust to row shifts month-to-month.
# Never hardcode row numbers: the T12 is a rolling window and line items
# get added/removed, so find everything by name / GL-code suffix.
# =======================================================================
def is_subtotal(code, name):
    if isinstance(code, str) and code.endswith(("-099", "-098", "-090", "-999", "-199")):
        return True
    nm = name.strip() if name else ""
    return ("Total" in nm) or nm.isupper()

def t12_find(exact_name):
    target = norm(exact_name)
    for r in range(6, t12.max_row + 1):
        if norm(t12.cell(r, 2).value) == target:
            return r
    return None

NOI_ROW = t12_find("Net Operating Income")
TOTAL_INCOME_ROW = t12_find("Total Income")
TOTAL_OPEX_ROW = t12_find("Total Operating Expenses")
NET_INCOME_ROW = t12_find("Net Income")
TOTAL_COL = 18  # original Total col O(15) shifts to R(18) after 3 budget cols inserted

# Ordered list of every T12 subtotal/total row, plus the individual line
# rows that roll up into each (the lines since the previous subtotal).
t12_subtotal_rows = []
members = {}
_pending = []
for r in range(6, t12.max_row + 1):
    nm = t12.cell(r, 2).value
    if not nm:
        continue
    if is_subtotal(t12.cell(r, 1).value, nm):
        t12_subtotal_rows.append(r)
        members[r] = _pending
        _pending = []
    else:
        _pending.append(r)

# Budget Comparison subtotal rows, keyed by normalized name (BC has no
# GL-code column, so subtotals are detected by name only).
def bc_cells(r):
    return dict(r=r, pa=num(bc.cell(r, 2).value), pb=num(bc.cell(r, 3).value),
                ya=num(bc.cell(r, 6).value), yb=num(bc.cell(r, 7).value))

bc_subtotals = {}
for r in range(6, 252):
    nm = bc.cell(r, 1).value
    if not nm:
        continue
    s = nm.strip()
    if ("Total" in s) or s.isupper() or norm(s) in ("potential rent", "net operating income", "net income"):
        bc_subtotals.setdefault(norm(s), []).append(bc_cells(r))

# =======================================================================
# Build budget lookup (by name, bucketed by section) BEFORE editing sheets
# =======================================================================
def bc_bucket(r):
    if r < 54: return "income"
    if r < 206: return "opex"
    if r < 226: return "routine"
    return "capital"

ROLLUPS = {"potential rent", "net operating income", "noi after replacements",
           "net income before allocations", "net income", "net income before debt service"}
budget_by_name = {}
for r in range(6, 252):
    name = bc.cell(r, 1).value
    if not name or "Total" in name or name.strip().isupper():
        continue
    if norm(name) in ROLLUPS:
        continue
    pa, pb, pv, pp = (num(bc.cell(r, c).value) for c in (2, 3, 4, 5))
    ya, yb, yv, yp = (num(bc.cell(r, c).value) for c in (6, 7, 8, 9))
    if pa == pb == ya == yb == 0:
        continue
    rec = dict(r=r, name=name.strip(), pa=pa, pb=pb, pv=pv, pp=pp, ya=ya, yb=yb, yv=yv, yp=yp,
               bucket=bc_bucket(r))
    budget_by_name.setdefault(norm(name), []).append(rec)

def t12_bucket(code):
    g = str(code) if code else ""
    if g[:1] == "4": return "income"
    if g[:1] in ("5", "6"): return "opex"
    if g[:5].isdigit() and 71000 <= int(g[:5]) < 71500: return "routine"
    return "capital"

def lookup_budget(name, bucket):
    cands = budget_by_name.get(norm(name))
    if not cands:
        return None
    same = [c for c in cands if c["bucket"] == bucket]
    return (same or cands)[0]

# subtotal budget figures for the exec-summary metrics (matched by name)
def bc_sub(name):
    cand = bc_subtotals.get(norm(name))
    return cand[0] if cand else dict(pa=0, pb=0, ya=0, yb=0)
SUB_BUDGET = {
    "rental": bc_sub("Total Rental Inc. - Residential"),
    "income": bc_sub("Total Income"),
    "opex":   bc_sub("Total Operating Expenses"),
    "noi":    bc_sub("Net Operating Income"),
}

# =======================================================================
# Compute T12 above-NOI flags (Phase 2 logic)  -> {t12_row: (sev, note)}
# =======================================================================
REVIEW, PRIOR = 14, 13
M3 = [11, 12, 13]
M12 = list(range(3, 15))

t12_flags = {}
for r in range(6, NOI_ROW):
    code = t12.cell(r, 1).value
    name = t12.cell(r, 2).value
    if not name or is_subtotal(code, name.strip() if name else ""):
        continue
    name = name.strip()
    rev = num(t12.cell(r, REVIEW).value); prior = num(t12.cell(r, PRIOR).value)
    a3 = sum(num(t12.cell(r, c).value) for c in M3) / 3
    total = num(t12.cell(r, 15).value)
    gl = str(code) if code else ""
    is_income = gl[:1] == "4"; is_opex = gl[:1] in ("5", "6")
    contra = any(k in name for k in ("Concession", "Vacancy Loss", "Bad Debt", "Loss To Lease",
                                     "Employee Units", "Model & Storage"))
    primary_rent = name in ("Market Rent", "Gain / Loss To Lease")
    sev = None; reasons = []
    if is_income and rev < 0 and not contra and not primary_rent:
        sev = "High"; reasons.append(f"NEGATIVE other rental income ${rev:,.0f} — possible reversal/misposting.")
    if is_opex and rev < 0:
        sev = "High"; reasons.append(f"NEGATIVE expense ${rev:,.0f} — likely credit/reversal.")
    immaterial = abs(rev) < 500 and abs(a3) < 500
    if immaterial and not reasons:
        continue
    pm = (rev - prior) / abs(prior) if prior else None
    a3s = (rev - a3) / abs(a3) if a3 else None
    if is_income and not contra:
        if pm is not None and abs(pm) > 0.20 and abs(rev - prior) > 1000:
            reasons.append(f"{pm:+.0%} vs prior (${rev:,.0f} vs ${prior:,.0f})"); sev = sev or "Moderate"
        if a3s is not None and abs(a3s) > 0.25 and abs(rev - a3) > 2500:
            reasons.append(f"{a3s:+.0%} vs 3-mo avg (${rev:,.0f} vs ${a3:,.0f})"); sev = sev or "Moderate"
        if rev == 0 and a3 > 0:
            reasons.append("revenue dropped to $0"); sev = "High"
    if "One-Time Concessions" in name and a3 and abs(rev) > 1.5*abs(a3):
        reasons.append(f"{abs(rev)/abs(a3):.0%} of 3-mo avg concessions"); sev = sev or "Moderate"
    if name == "Bad Debt - Rent" and a3 and abs(rev) > 2*abs(a3) and abs(rev) > 5000:
        reasons.append(f"bad debt {abs(rev)/abs(a3):.0%} of 3-mo avg"); sev = "High"
    if "Vacancy Loss" in name and a3 and abs(rev) > 1.2*abs(a3) and abs(rev) > 5000:
        reasons.append(f"vacancy {abs(rev)/abs(a3):.0%} of 3-mo avg"); sev = sev or "Moderate"
    # Data-quality: positive value in a contra/bad-debt account (house rule B.6)
    if "Bad Debt" in name and rev > 0:
        reasons.append(f"POSITIVE in a contra/bad-debt account (${rev:,.0f}) — confirm recovery "
                       f"or recategorize (e.g., to Accelerated Rent)"); sev = sev or "Moderate"
    if is_opex:
        if pm is not None and abs(pm) > 0.25 and abs(rev - prior) > 1000:
            reasons.append(f"{pm:+.0%} vs prior (${rev:,.0f} vs ${prior:,.0f})"); sev = sev or "Moderate"
        if a3s is not None and abs(a3s) > 0.30 and abs(rev - a3) > 2500:
            reasons.append(f"{a3s:+.0%} vs 3-mo avg (${rev:,.0f} vs ${a3:,.0f})"); sev = sev or "Moderate"
        if rev == 0 and a3 > 0:
            reasons.append("expense dropped to $0 (3-mo avg > 0)"); sev = sev or "Moderate"
        if total and abs(rev) > 0.5*abs(total):
            reasons.append(f"single month = {rev/total:.0%} of 12-mo total"); sev = "High"
        if name == "Ad Valorem Property Taxes" and prior and abs(rev-prior)/abs(prior) > 0.10:
            reasons.append(f"{(rev-prior)/prior:+.0%} vs prior"); sev = sev or "Moderate"
        if "Property Insurance" in name and prior and abs(rev-prior) > 1:
            reasons.append(f"changed vs prior (${rev:,.0f} vs ${prior:,.0f}); should be level"); sev = sev or "Moderate"
        if any(u in name for u in ("Electric","Water","Gas","Utility")) and a3 and abs(rev) > 1.5*abs(a3) and abs(rev) > 2000:
            reasons.append(f"utility {abs(rev)/abs(a3):.0%} of 3-mo avg"); sev = sev or "Moderate"
    if reasons:
        # suppress the Electric Commissions false positive (one-time item, $0 is normal)
        if name == "Electric Commissions":
            continue
        # Brio override: Consulting/Professional Fees is the owner's own $3K/mo
        # consulting fee — expected, not a concern (see REVIEW_METHODOLOGY.md).
        if PROPERTY_IS_BRIO and name == "Consulting / Professional Fees":
            continue
        t12_flags[r] = (sev or "Moderate", "; ".join(reasons))

# GL findings to enrich notes (manual from Phase 5). Property-specific — the
# analyst adds these per review; only applied to the property they belong to.
GL_NOTES = {
    "Consulting / Professional Fees": "GL: vendor 'Oak Real Estate Investment' — Feb $1,285 + Mar $3,000 + Apr $3,000 all posted in May period; unbudgeted. Ask PM: nature of engagement, why unbudgeted, why multiple months booked at once (possible related party).",
    "Locator and Broker Referrals": "GL: Competitive Edge Realty (4 leases x $1,000) + Realty Texas + accrual true-ups. Heavy broker-locator leasing. Ask PM: confirm commissions tie to signed leases.",
    "Property Insurance": "GL: single prepaid amortization entry $17,121. Running UNDER budget (favorable) — likely renewal at lower premium. Ask PM: confirm new premium/term.",
    "Paint Contractor": "GL: R&C Painting & Cleaning — many unit make-ready invoices. Over budget PTD & YTD. Ask PM: turn volume driving paint spend.",
} if PROPERTY_IS_BRIO else {}

# =======================================================================
# Compute Budget flags (Phase 4)  -> {bc_row: (sev, note)}
# =======================================================================
bc_flags = {}
# Brio: owner's own consulting fee — suppress only for Brio.
SUPPRESS = {"Consulting / Professional Fees"} if PROPERTY_IS_BRIO else set()
for recs in budget_by_name.values():
    for it in recs:
        if it["name"] in SUPPRESS:
            continue
        fl = []; sev = None
        if it["pv"] < 0 and abs(it["pp"]) > 50 and abs(it["pv"]) > 2500:
            fl.append("PTD"); sev = "High"
        elif it["pv"] < 0 and abs(it["pp"]) > 20 and abs(it["pv"]) > 1000:
            fl.append("PTD"); sev = sev or "Moderate"
        if it["yv"] < 0 and abs(it["yp"]) > 30 and abs(it["yv"]) > 10000:
            fl.append("YTD"); sev = "High"
        elif it["yv"] < 0 and abs(it["yp"]) > 15 and abs(it["yv"]) > 5000:
            fl.append("YTD"); sev = sev or "Moderate"
        if fl:
            note = (f"{'/'.join(fl)} unfavorable to budget. "
                    f"PTD ${it['pa']:,.0f} vs ${it['pb']:,.0f} (${it['pv']:,.0f}); "
                    f"YTD ${it['ya']:,.0f} vs ${it['yb']:,.0f} (${it['yv']:,.0f}).")
            bc_flags[it["r"]] = (sev, note)

# =======================================================================
# PHASE 5.5 — restructure & annotate T12
# =======================================================================
t12.insert_cols(15, 3)   # new O,P,Q ; old Total O->R(18)
hdr_row = 5
for col, label in ((15, "PTD Budget"), (16, "YTD Actual"), (17, "YTD Budget")):
    c = t12.cell(hdr_row, col, label)
    c.fill = LIGHTBLUE; c.font = Font(bold=True); c.alignment = Alignment(horizontal="center")
NOTES_COL = 19  # S
hs = t12.cell(hdr_row, NOTES_COL, "Review Notes")
hs.fill = DARKBLUE; hs.font = WHITE_BOLD; hs.alignment = Alignment(horizontal="center")
t12.column_dimensions[get_column_letter(NOTES_COL)].width = 60
for col in (15, 16, 17):
    t12.column_dimensions[get_column_letter(col)].width = 13

# fill budget values for every row: line items match by name+bucket,
# subtotals/totals match against the Budget Comparison subtotal rows by
# name. Section headers (GL code -000) carry no budget, left blank.
for r in range(6, t12.max_row + 1):
    code = t12.cell(r, 1).value
    name = t12.cell(r, 2).value
    if not name:
        continue
    nm = name.strip()
    if is_subtotal(code, nm):
        if isinstance(code, str) and code.endswith("-000"):
            continue
        cand = bc_subtotals.get(norm(nm))
        b = cand[0] if cand else None
    else:
        b = lookup_budget(nm, t12_bucket(code))
    if not b:
        continue
    for col, key in ((15, "pb"), (16, "ya"), (17, "yb")):
        cell = t12.cell(r, col, round(b[key], 2))
        cell.number_format = "#,##0"
        if is_subtotal(code, nm):
            cell.font = Font(bold=True)

# group oldest 6 months C..H (collapsed) and freeze at C6
t12.column_dimensions.group("C", "H", hidden=True)
t12.freeze_panes = "C6"

# annotate flagged rows: highlight col N + note in col S
for r, (sev, note) in t12_flags.items():
    fill = RED if sev == "High" else YELLOW
    t12.cell(r, REVIEW).fill = fill
    name = t12.cell(r, 2).value.strip()
    full = f"[{sev}] {note}"
    if name in GL_NOTES:
        full += " " + GL_NOTES[name]
    cell = t12.cell(r, NOTES_COL, full)
    cell.alignment = WRAP

# variance-summary comments on EVERY subtotal/total row (no highlight).
# Drivers = the individual line items that roll into that subtotal, ranked
# by absolute budget variance.
def member_drivers(sub_row, ptd=True):
    key = "pv" if ptd else "yv"
    found = []
    for mr_ in members.get(sub_row, []):
        nm = t12.cell(mr_, 2).value
        if not nm:
            continue
        b = lookup_budget(nm.strip(), t12_bucket(t12.cell(mr_, 1).value))
        if b:
            found.append((nm.strip(), b[key]))
    found.sort(key=lambda x: -abs(x[1]))
    return ", ".join(f"{n} ({v:+,.0f})" for n, v in found[:3] if abs(v) > 1000)

for sub_row in t12_subtotal_rows:
    code = t12.cell(sub_row, 1).value
    nm = t12.cell(sub_row, 2).value.strip()
    # section/sub-section headers (GL code ends -000, e.g. INCOME, EXPENSES,
    # HVAC) carry no totals — they're boundaries, not roll-ups. Skip.
    if isinstance(code, str) and code.endswith("-000"):
        continue
    rev = num(t12.cell(sub_row, REVIEW).value)
    tot = num(t12.cell(sub_row, TOTAL_COL).value)
    cand = bc_subtotals.get(norm(nm))
    if cand:
        s = cand[0]
        pv = s["pa"] - s["pb"]; pvp = pv / s["pb"] * 100 if s["pb"] else 0
        yv = s["ya"] - s["yb"]; yvp = yv / s["yb"] * 100 if s["yb"] else 0
        pd = member_drivers(sub_row, True); yd = member_drivers(sub_row, False)
        txt = (f"═══ {nm.upper()} — VARIANCE SUMMARY ═══\n"
               f"PTD: ${s['pa']:,.0f} vs. ${s['pb']:,.0f} ({pv:+,.0f}, {pvp:+.1f}%)"
               + (f"\n  Drivers: {pd}" if pd else "")
               + f"\nYTD: ${s['ya']:,.0f} vs. ${s['yb']:,.0f} ({yv:+,.0f}, {yvp:+.1f}%)"
               + (f"\n  Drivers: {yd}" if yd else ""))
    else:
        txt = (f"═══ {nm.upper()} — SUMMARY ═══\n"
               f"Review month ${rev:,.0f}  |  12-mo total ${tot:,.0f}\n"
               f"(no matching line in Budget Comparison)")
    cell = t12.cell(sub_row, NOTES_COL, txt)
    cell.alignment = WRAP
    cell.font = Font(italic=True, color="1F4E79")

# =======================================================================
# Annotate Budget Comparison (adapted: highlight PTD Variance col D, note in col L)
# =======================================================================
BNOTE = 12  # column L
hb = bc.cell(5, BNOTE, "Review Notes")
hb.fill = DARKBLUE; hb.font = WHITE_BOLD; hb.alignment = Alignment(horizontal="center")
bc.column_dimensions[get_column_letter(BNOTE)].width = 55
for r, (sev, note) in bc_flags.items():
    fill = RED if sev == "High" else YELLOW
    bc.cell(r, 4).fill = fill  # PTD Variance cell
    c = bc.cell(r, BNOTE, f"[{sev}] {note}")
    c.alignment = WRAP

print(f"T12 flags annotated: {len(t12_flags)}")
print(f"Budget flags annotated: {len(bc_flags)}")

# =======================================================================
# PHASE 6 — Monthly Review sheet
# =======================================================================
if "Monthly Review" in wb.sheetnames:
    del wb["Monthly Review"]
mr = wb.create_sheet("Monthly Review", 0)  # first tab
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR = lambda: (DARKBLUE, WHITE_BOLD)

def sec(ws, row, text):
    c = ws.cell(row, 1, text)
    c.fill = DARKBLUE; c.font = Font(color="FFFFFF", bold=True, size=12)
    for col in range(1, 11):
        ws.cell(row, col).fill = DARKBLUE
    return row + 1

def money(cell):
    cell.number_format = "$#,##0"
def pct(cell):
    cell.number_format = "0.0%"

REV, PRI = 14, 13
ti = num(t12.cell(TOTAL_INCOME_ROW, REV).value); ti_p = num(t12.cell(TOTAL_INCOME_ROW, PRI).value)
oe = num(t12.cell(TOTAL_OPEX_ROW, REV).value); oe_p = num(t12.cell(TOTAL_OPEX_ROW, PRI).value)
noi = num(t12.cell(NOI_ROW, REV).value); noi_p = num(t12.cell(NOI_ROW, PRI).value)
b = SUB_BUDGET

# Property name, review/prior month labels, and book — read from the T12 header
PROP = str(t12.cell(1, 1).value or "Property").strip()
REVIEW_LABEL = str(t12.cell(5, REV).value or "Review Month").strip()
PRIOR_LABEL = str(t12.cell(5, PRI).value or "Prior Month").strip()
_bk = str(t12.cell(4, 1).value or "")
BOOK = _bk.split("=", 1)[1].split(";")[0].strip() if "=" in _bk else "Accrual"

mr.cell(1, 1, f"{PROP} — Monthly Financial Review").font = Font(bold=True, size=16)
mr.cell(2, 1, f"Review Period: {REVIEW_LABEL}  ({BOOK})").font = Font(italic=True, size=11)

row = 4
row = sec(mr, row, "1. EXECUTIVE SUMMARY")
hdrs = ["Metric", REVIEW_LABEL, PRIOR_LABEL, "MoM $", "MoM %", "YTD Actual", "YTD Budget", "YTD Var $", "YTD Var %"]
for i, h in enumerate(hdrs, 1):
    c = mr.cell(row, i, h); c.fill = DARKBLUE; c.font = WHITE_BOLD
row += 1
def metric(label, cur, prev, ya, yb):
    global row
    mr.cell(row, 1, label)
    money(mr.cell(row, 2, round(cur))); money(mr.cell(row, 3, round(prev)))
    money(mr.cell(row, 4, round(cur - prev)))
    mr.cell(row, 5, (cur - prev) / abs(prev) if prev else 0); pct(mr.cell(row, 5))
    money(mr.cell(row, 6, round(ya))); money(mr.cell(row, 7, round(yb)))
    money(mr.cell(row, 8, round(ya - yb)))
    mr.cell(row, 9, (ya - yb) / abs(yb) if yb else 0); pct(mr.cell(row, 9))
    row += 1
metric("Total Income", ti, ti_p, b["income"]["ya"], b["income"]["yb"])
metric("Total Operating Expenses", oe, oe_p, b["opex"]["ya"], b["opex"]["yb"])
metric("Net Operating Income", noi, noi_p, b["noi"]["ya"], b["noi"]["yb"])
mr.cell(row, 1, "NOI Margin"); mr.cell(row, 2, noi/ti if ti else 0); pct(mr.cell(row, 2))
mr.cell(row, 3, noi_p/ti_p if ti_p else 0); pct(mr.cell(row, 3)); row += 2

# Data-driven assessment (generic across properties)
noi_ptd_var = (noi - b["noi"]["pb"]) / b["noi"]["pb"] if b["noi"]["pb"] else 0
noi_ytd_var = (b["noi"]["ya"] - b["noi"]["yb"]) / b["noi"]["yb"] if b["noi"]["yb"] else 0
n_high = sum(1 for s, _ in t12_flags.values() if s == "High")
n_mod = sum(1 for s, _ in t12_flags.values() if s == "Moderate")
_tone = ("requires attention" if (noi_ptd_var < -0.10 or n_high >= 3)
         else "on track with watch items" if (noi_ptd_var < -0.03 or n_high or n_mod)
         else "generally on track")
assess = (f"{REVIEW_LABEL}: {_tone}. NOI ${noi:,.0f} is {noi_ptd_var:+.1%} vs budget (MTD); "
          f"YTD NOI is {noi_ytd_var:+.1%} vs budget. "
          f"{n_high} high / {n_mod} moderate item(s) flagged above the NOI line — see Section 2 "
          f"and the month-focused questions in Section 5.")
mr.cell(row, 1, "Assessment:").font = Font(bold=True)
ac = mr.cell(row, 2, assess); ac.alignment = WRAP
mr.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
row += 3

# Section 2: flagged items above the line
row = sec(mr, row, "2. FLAGGED ITEMS — ABOVE THE NOI LINE")
cols2 = ["GL Code", "Line Item", REVIEW_LABEL, "Prior Mo", "3-Mo Avg", "PTD Budget", "Severity", "Finding / Question for PM"]
for i, h in enumerate(cols2, 1):
    c = mr.cell(row, i, h); c.fill = DARKBLUE; c.font = WHITE_BOLD
row += 1
mr.freeze_panes = mr.cell(row, 1)
def sev_rank(r):
    return 0 if t12_flags[r][0] == "High" else 1
for r in sorted(t12_flags, key=lambda r: (sev_rank(r), -abs(num(t12.cell(r, REV).value)))):
    sev, note = t12_flags[r]
    name = t12.cell(r, 2).value.strip()
    rev = num(t12.cell(r, REV).value); prior = num(t12.cell(r, PRI).value)
    a3 = sum(num(t12.cell(r, c).value) for c in M3) / 3
    pb = t12.cell(r, 15).value
    full = note + (" " + GL_NOTES[name] if name in GL_NOTES else "")
    mr.cell(row, 1, t12.cell(r, 1).value)
    mr.cell(row, 2, name)
    money(mr.cell(row, 3, round(rev))); money(mr.cell(row, 4, round(prior))); money(mr.cell(row, 5, round(a3)))
    if isinstance(pb, (int, float)):
        money(mr.cell(row, 6, pb))
    sc = mr.cell(row, 7, sev); sc.fill = RED if sev == "High" else YELLOW
    if sev == "High": sc.font = Font(color="FFFFFF", bold=True)
    fc = mr.cell(row, 8, full); fc.alignment = WRAP
    row += 1
row += 2

# Section 3: below the line
row = sec(mr, row, "3. BELOW-THE-LINE SUMMARY")
for i, h in enumerate(["GL Code", "Line Item", REVIEW_LABEL, "12-Mo Total", "Note"], 1):
    c = mr.cell(row, i, h); c.fill = DARKBLUE; c.font = WHITE_BOLD
row += 1
for r in range(NOI_ROW + 1, t12.max_row + 1):
    code = t12.cell(r, 1).value; name = t12.cell(r, 2).value
    if not name: continue
    nm = name.strip()
    rev = num(t12.cell(r, REV).value); tot = num(t12.cell(r, TOTAL_COL).value)
    if is_subtotal(code, nm):
        if nm in ("Total Routine Replacement Expense", "Total Capital / Renovation Expense",
                  "Total Debt Service", "NOI After Replacements", "Net Income"):
            mr.cell(row, 1, code); cc = mr.cell(row, 2, nm); cc.font = Font(bold=True)
            money(mr.cell(row, 3, round(rev))); money(mr.cell(row, 4, round(tot)))
            row += 1
        continue
    if rev == 0 and tot == 0: continue
    mr.cell(row, 1, code); mr.cell(row, 2, nm)
    money(mr.cell(row, 3, round(rev))); money(mr.cell(row, 4, round(tot)))
    if nm == "Interest Expense - 1st Mortgage":
        mr.cell(row, 5, "Debt service; varies by days-in-month, on budget. Primary driver of net loss.").alignment = WRAP
    row += 1
net_income = num(t12.cell(NET_INCOME_ROW, REV).value)
mr.cell(row, 2, "NOI-to-Net-Income bridge:").font = Font(bold=True)
mr.cell(row, 3, f"NOI ${noi:,.0f}  →  Net Income ${net_income:,.0f}  (below-line drag ${net_income-noi:,.0f}, ~all debt service)")
mr.cell(row, 3).alignment = WRAP
row += 3

# Section 4: budget highlights
row = sec(mr, row, "4. BUDGET VARIANCE HIGHLIGHTS")
allitems = [it for recs in budget_by_name.values() for it in recs]
def vtable(title, key, pct_key, n=8):
    global row
    mr.cell(row, 1, title).font = Font(bold=True); row += 1
    for i, h in enumerate(["Line Item", "Actual", "Budget", "Var $", "Var %", "Favorable?"], 1):
        c = mr.cell(row, i, h); c.fill = PatternFill("solid", fgColor="8EAADB"); c.font = Font(bold=True)
    row += 1
    a_i, b_i = (2, 3) if key == "pv" else (6, 7)
    persistent = {it["name"] for it in allitems if it["pv"] < 0 and it["yv"] < 0
                  and abs(it["pv"]) > 1000 and abs(it["yv"]) > 5000}
    for it in sorted(allitems, key=lambda x: x[key])[:n]:
        mr.cell(row, 1, it["name"])
        money(mr.cell(row, 2, round(it[a_i_key(key)[0]]))); money(mr.cell(row, 3, round(it[a_i_key(key)[1]])))
        money(mr.cell(row, 4, round(it[key])))
        mr.cell(row, 5, it[pct_key] / 100); pct(mr.cell(row, 5))
        fav = "Favorable" if it[key] > 0 else "Unfavorable"
        if it["name"] in persistent: fav += " (Persistent)"
        mr.cell(row, 6, fav)
        row += 1
    row += 1
def a_i_key(key):
    return ("pa", "pb") if key == "pv" else ("ya", "yb")
vtable("PTD — Largest Unfavorable Variances", "pv", "pp")
vtable("YTD — Largest Unfavorable Variances", "yv", "yp")

# Section 5: questions for PM
# House rules (see REVIEW_METHODOLOGY.md): questions are CURRENT-MONTH (MTD) only —
# no YTD-driven questions; skip items favorable/in-line with budget; Gain/Loss to
# Lease is expected; Make-Ready framed with vacancy/move-in context; Tax/Ins/Mortgage
# are self-checked against the portfolio reference table, not asked.
row = sec(mr, row, "5. QUESTIONS FOR PROPERTY MANAGER")
scope = ("Scope: questions cover the current review month (MTD). YTD trends are "
         "summarized in Sections 1 & 4 for context only. Tax, insurance, and mortgage "
         "are tracked against the portfolio reference table, not raised here.")
sc_cell = mr.cell(row, 2, scope); sc_cell.alignment = WRAP; sc_cell.font = Font(italic=True, color="1F4E79")
mr.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
row += 1
if PROPERTY_IS_BRIO:
    # Curated, analyst-authored questions for Brio (this review).
    questions = [
        "Locator & Broker Referrals: $13,660 this month (+646% vs 3-mo avg), $8,186 over PTD budget. Please confirm each commission ties to a signed lease (Competitive Edge Realty, 4 leases).",
        "Make-Ready / turnover (Paint Contractor, Carpets, Other Make-Ready) ran over budget this month — this looks consistent with elevated move-ins / declining vacancy. Please confirm the turn count this month so we can tie make-ready spend to move-in volume.",
        "Lease Cancellation Fee income was $0 this month vs. budget. Is this fee still being charged and collected?",
        "Bad Debt – Accelerated Rent shows a POSITIVE balance. Was this amount recovered, or should it be reclassified out of Bad Debt (e.g., into Accelerated Rent)?",
        "Internet Listing Services (Zillow, 54012-000): the Feb and March accruals were reversed with no offsetting actual expense booked, so those months understate ILS cost. Please confirm the true monthly ILS amount and rebook the missing actuals.",
    ]
else:
    # Generic auto-draft from this month's flags (High first). The analyst should
    # refine these per the house rules (REVIEW_METHODOLOGY.md) before sending.
    questions = []
    for r in sorted(t12_flags, key=lambda r: (0 if t12_flags[r][0] == "High" else 1,
                                              -abs(num(t12.cell(r, REV).value)))):
        sev, note = t12_flags[r]
        nm = t12.cell(r, 2).value.strip()
        questions.append(f"[{sev}] {nm}: {note} — please explain (and confirm via GL there is "
                         f"no accrual reversed without an offsetting actual this month).")
    if not questions:
        questions = ["No month-specific items flagged above the NOI line. Confirm no accruals were "
                     "reversed without an offsetting actual expense this month."]
for i, q in enumerate(questions, 1):
    mr.cell(row, 1, i)
    qc = mr.cell(row, 2, q); qc.alignment = WRAP
    mr.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
    row += 1

# widths
widths = {1: 14, 2: 30, 3: 13, 4: 13, 5: 13, 6: 13, 7: 12, 8: 60, 9: 12}
for col, w in widths.items():
    mr.column_dimensions[get_column_letter(col)].width = w

OUT = WB.with_name(WB.stem + "_Reviewed.xlsx")
try:
    wb.save(WB); saved = WB
except PermissionError:
    wb.save(OUT); saved = OUT
print(f"Saved: {saved}")
