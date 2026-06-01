import sys
import openpyxl
from pathlib import Path
from datetime import datetime, timedelta

WB = Path(r"H:\.shortcut-targets-by-id\15cPT84Tcymc9b2jqyLRYJjcuEXIBOfOp\0. Business\Multi Family\!!! W.C. Portfolio\!!! Brio\Financial\202605\Brio_Financial_Analysis_202605.xlsx")
wb = openpyxl.load_workbook(WB, data_only=True)
gl = wb["GL"]

def serial_to_date(v):
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, (int, float)) and v > 30000:
        return (datetime(1899, 12, 30) + timedelta(days=v)).strftime("%Y-%m-%d")
    return str(v) if v is not None else ""

codes = sys.argv[1:] if len(sys.argv) > 1 else []

# Find section header rows: column A holds a GL code like 54050-000
header_rows = []
for r in range(1, gl.max_row + 1):
    a = gl.cell(r, 1).value
    if isinstance(a, str) and len(a) >= 8 and a[5:6] == "-" and a[:5].isdigit():
        header_rows.append((r, a, gl.cell(r, 5).value))

if not codes:
    print(f"GL has {len(header_rows)} account sections. First 8 cols of row 5-8 for layout:")
    for r in range(5, 12):
        print("  r%d: " % r, [gl.cell(r, c).value for c in range(1, 12)])
    print("\nSample headers:")
    for hr in header_rows[:10]:
        print("  ", hr)
    sys.exit()

# For each requested code, print its section transactions
hr_index = {a: r for r, a, _ in header_rows}
sorted_hr = sorted(r for r, a, _ in header_rows)
for code in codes:
    if code not in hr_index:
        print(f"\n### {code}: NOT FOUND in GL")
        continue
    start = hr_index[code]
    end = next((r for r in sorted_hr if r > start), gl.max_row + 1)
    print(f"\n{'='*100}\n### GL {code}  ({gl.cell(start,5).value})   rows {start+1}..{end-1}\n{'='*100}")
    print(f"{'Date':<12}{'Period':<10}{'Description':<40}{'Control':<12}{'Debit':>12}{'Credit':>12}")
    for r in range(start + 1, end):
        date = serial_to_date(gl.cell(r, 3).value)
        period = serial_to_date(gl.cell(r, 4).value)
        desc = str(gl.cell(r, 5).value or "")[:38]
        ctrl = str(gl.cell(r, 6).value or "")[:10]
        deb = gl.cell(r, 8).value
        cred = gl.cell(r, 9).value
        deb = f"{deb:,.2f}" if isinstance(deb, (int, float)) else ""
        cred = f"{cred:,.2f}" if isinstance(cred, (int, float)) else ""
        if not (date or desc or deb or cred):
            continue
        print(f"{date:<12}{period:<10}{desc:<40}{ctrl:<12}{deb:>12}{cred:>12}")
