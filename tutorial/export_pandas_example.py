#!/usr/bin/env python3
"""
Advanced bulk export tutorial for xlsxedit (write_dataframe).

For the beginner pandas engine (ExcelWriter / read_excel), see:
    python tutorial/pandas_tutorial.py

Run from the project folder:
    python tutorial/export_pandas_example.py

Creates tutorial/output/export_example.xlsx with three sheets — open it in Excel when done:
    Overwrite    — zebra + currency formats in Python (footer replaced)
    Insert       — rows pushed down (footer shifts to row 8)
    Style rows   — copy styles from template sample rows 7-8

Uses pandas if installed; otherwise a tiny built-in fallback (no extra install).
"""

from __future__ import annotations

import sys

# Path helps work with file and folder paths in a cross-platform way.
from pathlib import Path

# Allow ``from xlsxedit import …`` when running from a GitHub clone without
# ``pip install -e .`` (src layout: repo/src/xlsxedit).
_SRC = Path(__file__).resolve().parent.parent / "src"
if _SRC.is_dir() and str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

# Workbook is the main class: it represents one .xlsx file in memory.
from xlsxedit import Workbook

# =============================================================================
# TABLE OF CONTENTS — jump to a section by searching for its number (e.g. "§3")
# =============================================================================
# §1  Sample data (pandas or fallback)
# §2  Create workbook + three sheets
# §3  Overwrite sheet — row_styles + column_styles in Python
# §4  Insert sheet — rows push footer down
# §5  Style rows sheet — copy styles from template sample rows
# §6  Save one file + verify all sheets
# §7  Row map summary
# =============================================================================

# __file__ is the path to this script. .parent is the tutorial/ folder.
HERE = Path(__file__).resolve().parent
OUTPUT_DIR = HERE / "output"
OUTPUT = OUTPUT_DIR / "export_example.xlsx"


def section(title: str) -> None:
    """Print a visible heading so the console output is easy to follow."""
    print(f"\n--- {title} ---")


def build_report_template(ws) -> None:
    """Rows 1–4 = meta + headers; row 5 = footer. Data slot starts at A5."""
    # ws.merge_cells(cell_range: str) -> None   e.g. "A1:D1"
    ws.merge_cells("A1:D1")

    # cell.apply_style(*, bold=, font_size=, bg_color=, font_color=, …) -> None
    ws["A1"].value = "Sales report"
    ws["A1"].apply_style(bold=True, font_size=16, bg_color="FF4472C4", font_color="FFFFFFFF")

    ws["A2"].value = "Client:"
    ws["B2"].value = "Jonas Corp"
    ws["C2"].value = "Date:"
    ws["D2"].value = "2025-08-07"
    # cell.apply_date_format() -> None   sets a standard date number format
    ws["D2"].apply_date_format()

    # Row 4 = column titles (we use header=False when writing the DataFrame).
    ws["A4"].value = "Item"
    ws["B4"].value = "Qty"
    ws["C4"].value = "Unit"
    ws["D4"].value = "Line total"
    for col in ("A", "B", "C", "D"):
        ws[f"{col}4"].apply_style(bold=True, bg_color="FFD9E1F2")

    # Row 5 = footer, directly under the header block — easy to see in Excel.
    # Overwrite at A5 replaces this cell; insert at A5 pushes it down.
    ws["A5"].value = "Confidential — internal use only"
    ws["A5"].apply_style(italic=True, font_color="FF808080")


def make_dataframe():
    """Return a pandas-like object with .columns and .itertuples(index=False)."""
    try:
        import pandas as pd
    except ImportError:
        return _SimpleFrame(
            ["Item", "Qty", "Unit", "Line total"],
            [
                ("Widget", 2, 10.0, 20.0),
                ("Gadget", 1, 25.5, 25.5),
                ("Gizmo", 3, 7.25, 21.75),
            ],
        )

    return pd.DataFrame(
        {
            "Item": ["Widget", "Gadget", "Gizmo"],
            "Qty": [2, 1, 3],
            "Unit": [10.0, 25.5, 7.25],
            "Line total": [20.0, 25.5, 21.75],
        }
    )


class _SimpleFrame:
    """Minimal DataFrame stand-in so this script runs without pip install pandas."""

    def __init__(self, columns, rows):
        self.columns = columns
        self._rows = rows

    def itertuples(self, index=False):
        for row in self._rows:
            yield tuple(row)


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # -------------------------------------------------------------------------
    # §1  Sample data (pandas or fallback)
    # -------------------------------------------------------------------------
    section("Sample data")

    # write_dataframe() accepts any object with .columns and .itertuples() / .values.
    # Pandas is optional — only lxml is required for xlsxedit itself.
    df = make_dataframe()
    print(f"DataFrame columns: {list(df.columns)}")
    print(f"Row count: {len(list(df.itertuples(index=False)))}")

    # -------------------------------------------------------------------------
    # §2  Create workbook + three sheets
    # -------------------------------------------------------------------------
    section("Create workbook + three sheets")

    # Typical report layout (same on each sheet):
    #
    #   Row  1-3   Title, client, date
    #   Row  4     Column headers
    #   Row  5     Footer (right below header)
    #   Rows 5-7   Data from DataFrame (starts at A5 in overwrite / insert demos)
    #
    # Overwrite at A5 replaces the footer cell with the first data row.
    # Insert at A5 pushes the footer down by len(df) (row 5 -> row 8 for 3 rows).

    # Workbook.create() -> Workbook
    wb = Workbook.create()

    # wb.rename_worksheet(old_name: str, new_name: str) -> Worksheet
    wb.rename_worksheet("Sheet1", "Overwrite")

    # wb.add_worksheet(name: str) -> Worksheet
    wb.add_worksheet("Insert")
    wb.add_worksheet("Style rows")
    print(f"Sheets: {wb.sheetnames}")

    for name in ("Overwrite", "Insert", "Style rows"):
        build_report_template(wb[name])
    print("Built header + footer on each sheet (footer on row 5)")

    # -------------------------------------------------------------------------
    # §3  Overwrite sheet — row_styles + column_styles in Python
    # -------------------------------------------------------------------------
    section("Overwrite sheet — overwrite + inline styles")

    ws = wb["Overwrite"]

    # ws.clear_range(cell_range: str) -> int
    # Overwrite mode only — clears old cell values before rewriting (styles kept).
    ws.clear_range("A5:D10")

    # wb.write_dataframe(df, *, sheet=, at_cell=, header=, mode=, row_styles=, column_styles=) -> int
    #
    # at_cell="A5"     — first cell written (overwrites the footer we put on A5).
    # header=False     — do not write DataFrame column names (row 4 already has titles).
    # mode="overwrite" — write at fixed rows; does not shift rows below.
    #
    # row_styles — one dict per zebra stripe (row 1 white, row 2 grey, then repeats).
    # column_styles — one dict per column (A, B, C, D); use num_format for numbers/currency.
    n_overwrite = wb.write_dataframe(
        df,
        sheet="Overwrite",
        at_cell="A5",
        header=False,
        mode="overwrite",
        row_styles=[
            {"bg_color": "FFFFFFFF"},
            {"bg_color": "FFF2F2F2"},
        ],
        column_styles=[
            {},
            {"num_format": "0"},
            {"num_format": "$#,##0.00"},
            {"num_format": "$#,##0.00"},
        ],
        resize_table=False,
    )
    print(f"Wrote {n_overwrite} data rows starting at A5")
    print(f"Footer on A5 overwritten — now: {ws['A5'].value!r}")

    # -------------------------------------------------------------------------
    # §4  Insert sheet — rows push footer down
    # -------------------------------------------------------------------------
    section("Insert sheet — insert mode")

    ws = wb["Insert"]

    # mode="insert" — insert len(df) blank rows at start, then write data.
    # Footer on row 5 shifts down to row 8 (5 + 3 data rows).
    n_insert = wb.write_dataframe(
        df,
        sheet="Insert",
        at_cell="A5",
        header=False,
        mode="insert",
        row_styles=[
            {"bg_color": "FFFFFFFF"},
            {"bg_color": "FFF2F2F2"},
        ],
        column_styles=[
            {},
            {"num_format": "0"},
            {"num_format": "$#,##0.00"},
            {"num_format": "$#,##0.00"},
        ],
        resize_table=False,
    )
    print(f"Inserted {n_insert} rows at A5")
    print(f"Footer moved to A8: {ws['A8'].value!r}")

    # -------------------------------------------------------------------------
    # §5  Style rows sheet — copy styles from template sample rows
    # -------------------------------------------------------------------------
    section("Style rows sheet — template_rows")

    ws = wb["Style rows"]

    # Rows 7-8 = style samples in the sheet (design in Excel or here).
    # write_dataframe(..., template_rows=[7, 8]) copies each column's s style from these rows.
    ws["A7"].value = "sample"
    ws["A7"].apply_style(bg_color="FFFFFFFF")
    ws["B7"].apply_number_format("0")
    ws["C7"].apply_number_format("$#,##0.00")
    ws["D7"].apply_number_format("$#,##0.00")
    ws["A8"].value = "sample"
    ws["A8"].apply_style(bg_color="FFF2F2F2")
    ws["B8"].apply_number_format("0")
    ws["C8"].apply_number_format("$#,##0.00")
    ws["D8"].apply_number_format("$#,##0.00")

    wb.write_dataframe(
        df,
        sheet="Style rows",
        at_cell="A7",
        header=False,
        mode="overwrite",
        template_rows=[7, 8],
        clear_range="A7:D10",
        resize_table=False,
    )
    print("Wrote data at A7 using styles copied from rows 7-8")
    print(f"Footer on A5 unchanged: {ws['A5'].value!r}")

    # -------------------------------------------------------------------------
    # §6  Save one file + verify all sheets
    # -------------------------------------------------------------------------
    section("Save")

    # wb.save(path: str | Path) -> None
    wb.save(OUTPUT)
    print(f"Saved → {OUTPUT}")

    wb2 = Workbook.open(OUTPUT)
    assert wb2.sheetnames == ["Overwrite", "Insert", "Style rows"]

    assert wb2["Overwrite"]["A5"].value == "Widget"
    assert wb2["Overwrite"]["A5"].value != "Confidential — internal use only"

    assert wb2["Insert"]["A5"].value == "Widget"
    assert wb2["Insert"]["A8"].value == "Confidential — internal use only"

    assert wb2["Style rows"]["A7"].value == "Widget"
    assert wb2["Style rows"]["A5"].value == "Confidential — internal use only"

    print("Round-trip OK — all three sheets verified")

    # -------------------------------------------------------------------------
    # §7  Row map summary
    # -------------------------------------------------------------------------
    section("Row map summary")

    print(f"One file: {OUTPUT.name}")
    print()
    print("  Overwrite:")
    print("    Rows 1-3   header / meta")
    print("    Row  4     column titles")
    print("    Rows 5-7   data (overwrites footer that was on A5)")
    print()
    print("  Insert:")
    print("    Rows 1-4   header / meta (unchanged)")
    print("    Rows 5-7   data")
    print("    Row  8     footer (shifted down from row 5)")
    print()
    print("  Style rows:")
    print("    Rows 7-9   data (styles copied from sample rows 7-8)")
    print("    Row  5     footer (unchanged)")
    print()
    print("Open the workbook in Excel and switch tabs to compare the three approaches.")
    print()
    print("Pandas engine (write / read): python tutorial/pandas_tutorial.py")


if __name__ == "__main__":
    main()
