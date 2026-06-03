"""Step 1 of the monthly financial analysis: combine the three source
reports (12-month statement, budget comparison, general ledger) into a
single Excel workbook with standardized tab names."""

import argparse
import json
import re
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

GLCODE_RE = re.compile(r"^\s*\d{5}-\d{3}\s*$")
CODE_MAP_PATH = Path(__file__).with_name("account_codes.json")


def _norm(s):
    return re.sub(r"\s+", " ", str(s)).strip().lower() if s else ""


def has_code_column(ws):
    """True if column A already carries GL codes (e.g. '41000-000')."""
    for r in range(1, min(ws.max_row, 80) + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and GLCODE_RE.match(v):
            return True
    return False


def backfill_code_column(ws):
    """Some 12-month-statement exports omit the leading GL-code column (account
    names land in col A instead of col B). Downstream tools assume code=col A,
    name=col B. Insert a blank col A and backfill codes from the portfolio-wide
    account map so subtotal (-099/-098/...) detection and GL cross-ref work.
    Returns (inserted, filled) counts; (False, 0) if the column was already present."""
    if has_code_column(ws):
        return False, 0
    code_map = {}
    if CODE_MAP_PATH.exists():
        with open(CODE_MAP_PATH, encoding="utf-8") as fh:
            code_map = json.load(fh)
    ws.insert_cols(1)
    filled = 0
    for r in range(6, ws.max_row + 1):
        name = ws.cell(r, 2).value
        if not isinstance(name, str):
            continue
        code = code_map.get(_norm(name))
        if code:
            ws.cell(r, 1, code)
            filled += 1
    return True, filled

# Source files are detected by glob pattern — filenames embed each property's
# code (e.g. "txbrio"), so we match by shape rather than an exact name.
# (patterns tried in order, target tab name)
SOURCES = [
    (["*12_Month_Statement*", "*12*Month*Statement*", "*T12*"], "T12"),
    (["*Budget_Comparison*", "*Budget*Comparison*"], "Budget Comparison"),
    (["*GeneralLedger*", "*General*Ledger*"], "GL"),
]


def find_source(folder, patterns):
    """First .xlsx in folder matching any pattern, skipping our own outputs/temp files."""
    for pat in patterns:
        for p in sorted(folder.glob(pat)):
            if (p.suffix.lower() == ".xlsx" and "Financial_Analysis" not in p.name
                    and not p.name.startswith("~$")):
                return p
    return None


def property_name(folder):
    """Derive the property name from the path (folder named like '!!! Brio')."""
    for parent in folder.parents:
        n = parent.name
        if n.startswith("!!!") and "Portfolio" not in n:
            return n.lstrip("! ").strip()
    return None


def copy_sheet(src_ws, dst_ws):
    for row in src_ws.iter_rows():
        for cell in row:
            new_cell = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            # text labels that start with "=" (e.g. "= Beginning Balance =") must
            # stay strings, not be reinterpreted as formulas
            if isinstance(cell.value, str) and cell.value.startswith("="):
                new_cell.data_type = "s"
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = cell.number_format
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)

    for merged in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(merged))

    for col, dim in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[col].width = dim.width
        dst_ws.column_dimensions[col].hidden = dim.hidden
    for idx, dim in src_ws.row_dimensions.items():
        dst_ws.row_dimensions[idx].height = dim.height
        dst_ws.row_dimensions[idx].hidden = dim.hidden

    if src_ws.sheet_view.showGridLines is not None:
        dst_ws.sheet_view.showGridLines = src_ws.sheet_view.showGridLines
    dst_ws.freeze_panes = src_ws.freeze_panes


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("folder", help="Folder containing the three source reports")
    parser.add_argument(
        "-o", "--output", default=None,
        help="Output .xlsx path (default: <folder>/Brio_Financial_Analysis_<folder name>.xlsx)",
    )
    args = parser.parse_args()

    folder = Path(args.folder)
    if args.output:
        out = Path(args.output)
    else:
        prop = property_name(folder)
        prefix = (prop.replace(" ", "_") + "_") if prop else ""
        out = folder / f"{prefix}Financial_Analysis_{folder.name}.xlsx"

    dst_wb = openpyxl.Workbook()
    dst_wb.remove(dst_wb.active)

    for patterns, tab in SOURCES:
        src_path = find_source(folder, patterns)
        if src_path is None:
            raise FileNotFoundError(f"No source for tab {tab!r} in {folder} (patterns: {patterns})")
        src_wb = openpyxl.load_workbook(src_path, data_only=False)
        src_ws = src_wb.worksheets[0]
        dst_ws = dst_wb.create_sheet(title=tab)
        copy_sheet(src_ws, dst_ws)
        src_wb.close()
        if tab == "T12":
            backfill_code_column(dst_ws)  # silently fix exports missing the GL-code column
        print(f"  {src_path.name}  ->  tab {tab!r}  ({src_ws.max_row} rows x {src_ws.max_column} cols)")

    dst_wb.save(out)
    print(f"\nSaved: {out}")


if __name__ == "__main__":
    main()
