#!/usr/bin/env python3
"""
Beginner pandas tutorial for xlsxedit.

Shows how to use engine="xlsxedit" with pandas ExcelWriter and read_excel.

Run from the project folder:
    python tutorial/pandas_tutorial.py

Requires: pip install xlsxedit[pandas]
Creates:  tutorial/output/pandas_tutorial.xlsx
"""

from __future__ import annotations

import sys
from pathlib import Path

# Allow ``from xlsxedit import …`` when running from a GitHub clone without
# ``pip install -e .`` (src layout: repo/src/xlsxedit).
_SRC = Path(__file__).resolve().parent.parent / "src"
if _SRC.is_dir() and str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

HERE = Path(__file__).resolve().parent
OUTPUT_DIR = HERE / "output"
OUTPUT = OUTPUT_DIR / "pandas_tutorial.xlsx"
TEMPLATE = OUTPUT_DIR / "pandas_template.xlsx"


def section(title: str) -> None:
    print(f"\n--- {title} ---")


def main() -> None:
    # -------------------------------------------------------------------------
    # §1  Install
    # -------------------------------------------------------------------------
    section("Install")

    print("Need: pip install xlsxedit[pandas]")
    try:
        import pandas as pd
    except ImportError:
        raise SystemExit(
            "pandas is not installed.\n"
            "  pip install xlsxedit[pandas]\n"
            "Then run this script again."
        )
    print(f"pandas {pd.__version__} OK")

    # -------------------------------------------------------------------------
    # §2  Register the engine (once per process)
    # -------------------------------------------------------------------------
    section("Register")

    import xlsxedit.pandas_io as xpi

    # register() enables engine="xlsxedit" for ExcelWriter and read_excel.
    # Writer: pandas register_writer. Reader: patches ExcelFile._engines
    # (pandas has no public register_reader yet; a future upstream PR may add one).
    
    xpi.register()
    print('Registered — you can now use engine="xlsxedit"')


    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # -------------------------------------------------------------------------
    # §3  Write a new Excel file
    # -------------------------------------------------------------------------
    section("Write a new file")

    df = pd.DataFrame(
        {
            "Item": ["Widget", "Gadget", "Gizmo"],
            "Qty": [2, 5, 1],
        }
    )

    # need to have done xpi.register() and import xlsxedit.pandas_io as xpi
    with pd.ExcelWriter(OUTPUT, engine="xlsxedit") as writer:
        df.to_excel(writer, sheet_name="Data", index=False)

    print(f"Wrote → {OUTPUT}")

    # -------------------------------------------------------------------------
    # §4  Read it back
    # -------------------------------------------------------------------------
    section("Read it back")

    got = pd.read_excel(OUTPUT, engine="xlsxedit")
    print(got)
    assert list(got.columns) == ["Item", "Qty"]
    assert got.iloc[0, 0] == "Widget"
    assert got.iloc[1, 1] == 5
    print("Round-trip OK")

    # -------------------------------------------------------------------------
    # §5  Fill an existing template (xlsxedit's main use case)
    # -------------------------------------------------------------------------
    section("Fill a template")

    from xlsxedit import Workbook

    # Minimal template: title + column headers; data goes on row 3 (startrow=2).
    wb = Workbook.create()
    ws = wb.rename_worksheet("Sheet1", "Report")
    ws["A1"].value = "Sales report"
    ws["A2"].value = "Item"
    ws["B2"].value = "Qty"
    wb.save(TEMPLATE)
    print(f"Template → {TEMPLATE}")

    # mode="a" = open existing file; overlay = write only at startrow/startcol.
    with pd.ExcelWriter(
        TEMPLATE,
        engine="xlsxedit",
        mode="a",
        if_sheet_exists="overlay",
    ) as writer:
        df.to_excel(
            writer,
            sheet_name="Report",
            startrow=2,  # Excel row 3 (0-based)
            header=False,
            index=False,
        )

    filled = Workbook.open(TEMPLATE)
    assert filled["Report"]["A1"].value == "Sales report"  # title kept
    assert filled["Report"]["A3"].value == "Widget"
    assert filled["Report"]["B4"].value == 5
    print("Template fill OK — title preserved, data at A3")

    # -------------------------------------------------------------------------
    # §6  Done
    # -------------------------------------------------------------------------
    section("Done")

    print(f"Open in Excel: {OUTPUT}")
    print(f"Template fill: {TEMPLATE}")
    print()
    print("Next steps:")
    print("  - Docs: docs/pandas.md")
    print("  - Bulk export / styles: python tutorial/export_pandas_example.py")


if __name__ == "__main__":
    main()
