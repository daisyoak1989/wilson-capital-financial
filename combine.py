"""Step 1 of the monthly financial analysis: combine the three source
reports (12-month statement, budget comparison, general ledger) into a
single Excel workbook with standardized tab names."""

import argparse
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

# (source filename, target tab name)
SOURCES = [
    ("12_Month_Statement_txbrio_Accrual.xlsx", "T12"),
    ("Budget_Comparison_txbrio_Accrual (1).xlsx", "Budget Comparison"),
    ("GeneralLedger_txbrio_Accrual.xlsx", "GL"),
]


def copy_sheet(src_ws, dst_ws):
    for row in src_ws.iter_rows():
        for cell in row:
            new_cell = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
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
    out = Path(args.output) if args.output else folder / f"Brio_Financial_Analysis_{folder.name}.xlsx"

    dst_wb = openpyxl.Workbook()
    dst_wb.remove(dst_wb.active)

    for filename, tab in SOURCES:
        src_path = folder / filename
        if not src_path.exists():
            raise FileNotFoundError(src_path)
        src_wb = openpyxl.load_workbook(src_path, data_only=False)
        src_ws = src_wb.worksheets[0]
        dst_ws = dst_wb.create_sheet(title=tab)
        copy_sheet(src_ws, dst_ws)
        src_wb.close()
        print(f"  {filename}  ->  tab {tab!r}  ({src_ws.max_row} rows x {src_ws.max_column} cols)")

    dst_wb.save(out)
    print(f"\nSaved: {out}")


if __name__ == "__main__":
    main()
