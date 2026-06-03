"""Step 3 of the monthly review: turn the analyst's own notes (T12 column T,
the column immediately right of the script's 'Review Notes') into an email-ready
list. For every row that has a note in that column, emits one line:

    [A: GL code] [B: description] — [T: your note]

Usage:
    python notes_to_email.py "<combined/reviewed workbook path>"

Prints a SUBJECT line, a blank line, then the body. The Gmail draft itself is
created by the assistant from this output (this script has no email access)."""

import sys
import openpyxl
from openpyxl.utils import get_column_letter


def main():
    if len(sys.argv) < 2 or not sys.argv[1].strip():
        sys.exit('Usage: notes_to_email.py "<workbook path>"')
    path = sys.argv[1]
    wb = openpyxl.load_workbook(path, data_only=True)
    if "T12" not in wb.sheetnames:
        sys.exit("No T12 tab in workbook.")
    t12 = wb["T12"]

    # Locate the analyst-notes column: the one immediately right of the script's
    # "Review Notes" header (robust to month-to-month column shifts). Fall back
    # to column T (20) if the Review Notes header isn't found.
    notes_col = None
    for c in range(1, t12.max_column + 2):
        if str(t12.cell(5, c).value or "").strip().lower() == "review notes":
            notes_col = c + 1
            break
    if notes_col is None:
        notes_col = 20  # column T

    prop = str(t12.cell(1, 1).value or "Property").strip()
    month = str(t12.cell(5, 14).value or "").strip()

    items = []
    for r in range(6, t12.max_row + 1):
        note = t12.cell(r, notes_col).value
        if note is None or not str(note).strip():
            continue
        gl = str(t12.cell(r, 1).value or "").strip()
        desc = str(t12.cell(r, 2).value or "").strip()
        items.append((gl, desc, str(note).strip()))

    col = get_column_letter(notes_col)
    if not items:
        print(f"No notes found in column {col} of the T12 tab.")
        print(f"(Add your notes in column {col} — the column just right of 'Review "
              f"Notes' — save, then re-run.)")
        return

    subject = f"{prop} — {month} Financial Review Notes" if month else f"{prop} — Financial Review Notes"
    print(f"SUBJECT: {subject}")
    print()
    print(f"Hi [PM],")
    print()
    print(f"Following my review of the {month} financials for {prop}, a few items "
          f"I'd like to go over:")
    print()
    for i, (gl, desc, note) in enumerate(items, 1):
        print(f"{i}. {gl} — {desc}: {note}")
    print()
    print("Thanks,")
    print("Daisy")
    print()
    print(f"[{len(items)} item(s) from column {col}]", file=sys.stderr)


if __name__ == "__main__":
    main()
