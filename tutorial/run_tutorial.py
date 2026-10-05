#!/usr/bin/env python3
"""
Tutorial script for xlsxedit.

Run from the project folder (no pip install required):
    python tutorial/run_tutorial.py

Or with an absolute path — same thing; the script adds ``../src`` to sys.path.

This creates tutorial/output/tutorial.xlsx — open it in Excel when done.
"""

from __future__ import annotations

import sys
from datetime import datetime

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
# §1  Open or create a workbook
# §2  Overview sheet — title, merge, column width
# §3  Invoice sheet — cell value types, placeholders, currency format
# §4  Search-and-replace (value_type for template placeholders)
# §5  Find by value + write beside (offset)
# §6  Formulas on Invoice
# §7  Formatting sheet — bold, colors, font size, alignment
# §8  Conditional formatting (color scale + cellIs)
# §9  Chart sheet — add_chart, chart.title
# §10 Table sheet — add_table, table.resize
# §11 Links sheet — hyperlink.url, hyperlink.location, hyperlink.display
# §12 Photos sheet — add_image, insert_image_at_placeholder, replace_image
# §13 Read-only fixture demos (console only, not saved to tutorial.xlsx)
# §14 Save to disk
# §15 Reopen and verify round-trip
# =============================================================================

# __file__ is the path to this script. .parent is the tutorial/ folder.
HERE = Path(__file__).resolve().parent
# We will save the Excel file here: tutorial/output/tutorial.xlsx
OUTPUT = HERE / "output" / "tutorial.xlsx"

# Repo-local image used for all image examples (no dependency on xlsx-inspect).
TUTORIAL_IMAGE = HERE.parent / "assets" / "coco-happy-swiss-nature.jpg"


def section(title: str) -> None:
    """Print a visible heading so the console output is easy to follow."""
    print(f"\n--- {title} ---")


def main() -> None:
    if not TUTORIAL_IMAGE.is_file():
        raise SystemExit(f"Tutorial image not found: {TUTORIAL_IMAGE}")

    # -------------------------------------------------------------------------
    # §1  Open or create a workbook
    # -------------------------------------------------------------------------
    section("Open or create a workbook")

    # Three equivalent ways to get a Workbook:
    #
    #   Workbook.open(path)  -> Workbook
    #       path: str | Path — an .xlsx file, or an unpacked OPC folder
    #
    #   Workbook(path)       -> Workbook   (same as open)
    #
    #   Workbook.create()    -> Workbook
    #   Workbook()           -> Workbook   (same as create; blank default.xlsx)
    #
    # Most real projects start from an Excel template someone designed:
    #   wb = Workbook.open("invoice-template.xlsx")
    # We use create() below to build every sheet step-by-step in this tutorial.
    # §15 reopens the saved file with Workbook.open(OUTPUT) — same API as above.

    # Quick peek: open an existing file from fixtures/ (console only).
    demo_path = HERE.parent / "fixtures" / "Book1.xlsx"
    if demo_path.is_file():
        # Workbook.open(path: str | Path) -> Workbook
        demo_wb = Workbook.open(demo_path)
        print(f"Opened {demo_path.name} — sheets: {demo_wb.sheetnames}")

    # Workbook.create() -> Workbook
    # Starts from a built-in template (like a blank Excel file) — not generating
    # XML from scratch, but loading a valid minimal package and mutating it.
    wb = Workbook.create()
    print(f"Created blank workbook — sheets: {wb.sheetnames}")

    # -------------------------------------------------------------------------
    # §2  Overview sheet — title, merge, column width
    # -------------------------------------------------------------------------
    section("Overview sheet")

    # wb.sheetnames -> list[str]
    print(f"Sheets in workbook: {wb.sheetnames}")

    # wb[name: str] -> Worksheet   select a tab by name (no rename)
    sheet1 = wb["Sheet1"]
    print(f"Selected sheet without renaming: {sheet1.name!r}")

    # wb.rename_worksheet(old_name: str, new_name: str) -> Worksheet
    overview = wb.rename_worksheet("Sheet1", "Overview")
    print(f"After rename — sheets: {wb.sheetnames}")

    # You can also select the renamed tab the same way:
    # overview = wb["Overview"]

    # ws.merge_cells(cell_range: str) -> None   e.g. "A1:F1"
    overview.merge_cells("A1:F1")

    # cell.apply_style(*, bold=, italic=, underline=, font_size=, font_name=,
    #                  font_color=, bg_color=, horizontal_align=, vertical_align=) -> None
    title = overview["A1"]
    title.value = "xlsxedit Tutorial"
    title.apply_style(
        bold=True,
        font_size=18,
        font_color="FFFFFFFF",
        bg_color="FF4472C4",
        horizontal_align="center",
        vertical_align="center",
    )

    # row_dimensions[row].height   column_dimensions[col].width
    overview.row_dimensions[1].height = 36
    overview.column_dimensions["A"].width = 28

    # Short guide to the other sheets in this workbook.
    overview["A3"].value = "Invoice — search-and-replace placeholders, dates, formulas"
    overview["A4"].value = "Formatting — fonts, colors, conditional formatting"
    overview["A5"].value = "Chart — bar chart from cell data"
    overview["A6"].value = "Table — Excel table with resize"
    overview["A7"].value = "Links — external URLs and internal sheet links"
    overview["A8"].value = "Photos — images and image search-and-replace"

    # ws.unmerge_cells(cell_range: str) -> None — undo a prior merge_cells call
    overview.merge_cells("A10:C10")
    overview["A10"].value = "temporary merge"
    overview.unmerge_cells("A10:C10")

    # -------------------------------------------------------------------------
    # §3  Invoice sheet — cell value types, placeholders, currency format
    # -------------------------------------------------------------------------
    section("Invoice sheet — layout")

    # wb.add_worksheet(name: str) -> Worksheet
    invoice = wb.add_worksheet("Invoice")
    invoice.column_dimensions["A"].width = 16
    invoice.column_dimensions["B"].width = 28

    # A small table is easier to read than scattered cells across the sheet.
    # Header row:
    invoice["A1"].value = "Field"
    invoice["A1"].apply_style(bold=True, font_size=12, bg_color="FFD9E1F2")
    invoice["B1"].value = "Value"
    invoice["B1"].apply_style(bold=True, font_size=12, bg_color="FFD9E1F2")

    # cell.value accepts: str | int | float | bool | datetime | date | None
    # The Python type you assign decides how Excel stores the cell.
    # How it *looks* (currency, date, percent) is a separate step — see apply_*_format below.
    #
    # SAR value_type (§4) is only for filling {placeholders} in templates.
    # You can always set values directly instead.

    invoice["A2"].value = "Client"
    # str — shared string placeholder (filled via SAR in §4)
    invoice["B2"].value = "{client}"

    invoice["A3"].value = "Ship date"
    # str placeholder now; §4 shows SAR → date. Direct alternative:
    # invoice["B3"].value = datetime(2025, 8, 7)
    # invoice["B3"].apply_date_format()
    invoice["B3"].value = "{ship_date}"

    invoice["A4"].value = "Item"
    # str — substring placeholder inside a longer string
    invoice["B4"].value = "Item {item} ready"

    invoice["A5"].value = "Quantity"
    # str placeholder now; §4 shows SAR → number. Direct alternative: invoice["B5"].value = 888
    invoice["B5"].value = "{qty}"

    invoice["A6"].value = "Unit price"
    # cell.value = 42          (int | float — stored as numeric <v>)
    # cell.apply_number_format(format_code: str) -> None
    invoice["B6"].value = 42
    invoice["B6"].apply_number_format("$#,##0.00")  # currency display in Excel
    print(f"Unit price: {invoice['B6'].value}, format: {invoice['B6'].style.num_format!r}")

    # -------------------------------------------------------------------------
    # §4  Search-and-replace (value_type for template placeholders)
    # -------------------------------------------------------------------------
    section("Search-and-replace on Invoice")

    # replace value_type is for template placeholders only.
    # It does the same thing you could do manually:
    #   "text"   → cell.value = str
    #   "number" → cell.value = int | float
    #   "date"   → cell.value = datetime + apply_date_format()
    # Currency / percent are not SAR value_types — set value + apply_number_format() directly (see B6).

    # wb.replace(old: str, new, *, value_type: str | None = None) -> int
    # Supported value_type values (formula cells are always skipped):
    #   "text" or None  — substring replace in string cells; new must be str
    #   "number"        — whole cell must equal old exactly; new: int | float
    #   "date"          — whole cell must equal old exactly; new: ISO str | datetime | date
    # Returns how many cells were changed.

    # value_type="text" (default) — substring replace across all sheets.
    n_client = wb.replace("{client}", "Jonas Corp", value_type="text")

    # value_type="text" — substring inside a longer string (B4: "Item {item} ready").
    n_item = invoice.replace("{item}", "Widget", value_type="text")

    # value_type="number" — whole-cell match; cell becomes a real number, not text.
    n_number = wb.replace("{qty}", 888, value_type="number")

    # value_type="date" — whole-cell match; writes Excel date serial + display format.
    n_date = wb.replace("{ship_date}", "2025-08-07", value_type="date")

    print(
        f"Replacements — text: {n_client + n_item}, "
        f"number: {n_number}, date: {n_date}"
    )
    print("B2:", invoice["B2"].value)
    print("B3:", invoice["B3"].value, "(date)")
    print("B5:", invoice["B5"].value, "(numeric)")

    # -------------------------------------------------------------------------
    # §5  Find by value + write beside (offset)
    # -------------------------------------------------------------------------
    section("Find by value and write beside")

    # Place a searchable label in free cells on Invoice (A9 was unused).
    invoice["A9"].value = "Anchor"

    # wb.find(value, *, sheet: str | None = None) -> Cell | None
    #   sheet=None  — search all sheets in workbook order
    #   sheet="…"   — search that sheet only
    # ws.find(value) / ws.findall(value) — same match rules, one sheet
    # Match: exact cell.value equality; formula cells are skipped.
    cell = wb.find("Anchor", sheet="Invoice")
    print(f"Found {cell.value!r} at {cell.address}")  # A9

    # cell.offset(cols=0, rows=0) -> Cell
    # right +1 column, next row +1 → B10
    target = cell.offset(cols=1, rows=1)
    print(f"Write target: {target.address}")
    target.value = "written via find+offset"

    # Sheet-local find returns the same address
    assert invoice.find("Anchor").address == cell.address
    assert len(wb.findall("Anchor")) == 1

    # wb.findall(value, *, sheet=None) -> list[Cell]
    # Same match rules as find; returns every hit (empty list if none).
    invoice["A11"].value = "Tag"
    invoice["C12"].value = "Tag"
    tags = invoice.findall("Tag")
    print(f"findall('Tag') → {len(tags)} cells: {[c.address for c in tags]}")
    assert len(tags) == 2
    assert [c.address for c in tags] == ["A11", "C12"]
    for c in tags:
        # one column to the right of each match
        c.offset(cols=1, rows=0).value = "via findall"

    # -------------------------------------------------------------------------
    # §6  Formulas on Invoice
    # -------------------------------------------------------------------------
    section("Formulas on Invoice")

    # cell.formula  — property: str | None  (read/write formula text, not result)
    invoice["A7"].value = "Total"
    invoice["A7"].apply_style(bold=True)
    invoice["B7"].formula = "=B5*B6"
    print("B7 formula:", invoice["B7"].formula)

    # cell.clear() -> None — remove value, formula, and style from one cell
    invoice["B8"].value = "scratch"
    invoice["B8"].clear()

    # ws.insert_columns(cols, *, at_col=… | at_cell=…) -> int
    # cols = sequence of column vectors (each top→bottom).
    # Shifts merges / CF / tables / column widths to the right.
    # Does NOT rewrite formula text or move drawing/chart anchors (same as insert_rows).
    n_cols = invoice.insert_columns([("Note", "from tutorial")], at_col="C")
    print(f"insert_columns wrote {n_cols} column(s); C1={invoice['C1'].value!r}, C2={invoice['C2'].value!r}")

    # -------------------------------------------------------------------------
    # §7  Formatting sheet — bold, colors, font size, alignment
    # -------------------------------------------------------------------------
    section("Formatting sheet — cell styles")

    formatting = wb.add_worksheet("Formatting")
    formatting.column_dimensions["A"].width = 14
    formatting.column_dimensions["B"].width = 32

    formatting["A1"].value = "Style"
    formatting["A1"].apply_style(bold=True, font_size=12, bg_color="FFD9E1F2")
    formatting["B1"].value = "Example"
    formatting["B1"].apply_style(bold=True, font_size=12, bg_color="FFD9E1F2")

    # apply_style changes the whole cell look via styles.xml (fontId, fillId on the cell).
    # Note: partial bold inside one cell (e.g. "only the middle is bolded") uses
    # shared-string <r> runs — see fixtures/TextSytle.xlsx A2. Not shown here.

    formatting["A2"].value = "Bold"
    formatting["B2"].value = "this is a bold cell"
    formatting["B2"].apply_style(bold=True)

    formatting["A3"].value = "Red text"
    formatting["B3"].value = "this is red text"
    formatting["B3"].apply_style(font_color="FF0000")

    formatting["A4"].value = "Large"
    formatting["B4"].value = "this is larger text size 16"
    formatting["B4"].apply_style(font_size=16)
    formatting.row_dimensions[4].height = 21

    formatting["A5"].value = "Yellow bg"
    formatting["B5"].value = "this has yellow background"
    formatting["B5"].apply_style(bg_color="FFFF00")

    formatting["A6"].value = "Centered"
    formatting["B6"].value = "this is centered text"
    formatting["B6"].apply_style(horizontal_align="center", vertical_align="center")
    formatting.row_dimensions[6].height = 36

    formatting["A7"].value = "Italic"
    formatting["B7"].value = "this is italic"
    formatting["B7"].apply_style(italic=True)

    formatting["A8"].value = "Underline"
    formatting["B8"].value = "this is underlined"
    formatting["B8"].apply_style(underline=True)

    formatting["A9"].value = "Arial"
    formatting["B9"].value = "this is arial text"
    formatting["B9"].apply_style(font_name="Arial")

    # -------------------------------------------------------------------------
    # §8  Conditional formatting
    # -------------------------------------------------------------------------
    section("Conditional formatting")

    # --- Color scale (gradient by value) ---
    # Colors are embedded in the rule itself — no extra style args needed.
    formatting["A11"].value = "Score"
    formatting["A11"].apply_style(bold=True, font_size=12, bg_color="FFD9E1F2")
    formatting["B11"].value = "Value"
    formatting["B11"].apply_style(bold=True, font_size=12, bg_color="FFD9E1F2")

    scores = [15, 45, 72, 91, 33]
    for i, score in enumerate(scores, start=12):
        formatting[f"A{i}"].value = f"Row {i - 11}"
        formatting[f"B{i}"].value = score

    # ws.add_color_scale_formatting(cell_range: str) -> ConditionalFormatting
    # Excel paints a red→yellow→green gradient from min to max over the range.
    formatting.add_color_scale_formatting("B12:B16")

    # --- cellIs rule (highlight when a condition is true) ---
    formatting["A18"].value = "Check"
    formatting["A18"].apply_style(bold=True, font_size=12, bg_color="FFD9E1F2")
    formatting["B18"].value = "Value"
    formatting["B18"].apply_style(bold=True, font_size=12, bg_color="FFD9E1F2")
    formatting["A19"].value = "positive"
    formatting["B19"].value = 5
    formatting["A20"].value = "negative"
    formatting["B20"].value = -1

    # ws.add_conditional_formatting(cell_range, *, operator=, formula=, bg_color=, font_color=, bold=)
    #   -> ConditionalFormatting
    # Applies to every cell in cell_range. Each cell is tested: e.g. operator="greaterThan"
    # with formula="0" means "value > 0". Matching cells get the rule's fill/font.
    # Default colors (when bg_color/font_color omitted): light-green fill C6EFCE,
    # dark-green font 006100 — same as Excel's "greater than" highlight.
    # B19 (5) should highlight; B20 (-1) stays plain.
    formatting.add_conditional_formatting("B19:B20", operator="greaterThan", formula="0")

    print(f"Conditional formatting rules: {len(formatting.conditional_formatting)}")

    # -------------------------------------------------------------------------
    # §9  Chart sheet — add_chart, chart.title
    # -------------------------------------------------------------------------
    section("Chart sheet")

    chart_ws = wb.add_worksheet("Chart")
    chart_ws["A1"].value = "Item"
    chart_ws["B1"].value = "Qty"
    chart_ws["A2"].value = "Apples"
    chart_ws["B2"].value = 3
    chart_ws["A3"].value = "Pears"
    chart_ws["B3"].value = 5

    # ws.add_chart(chart_type, *, anchor, data_range, title=, name=) -> Chart
    # chart_type: "bar" only for now. data_range must span at least two columns
    # (category column + value column, including header row).
    # anchor = top-left cell for the chart drawing (e.g. "D2").
    chart = chart_ws.add_chart("bar", anchor="D2", data_range="A1:B3", title="Sales")

    # chart.title: str | None — property read/write (updates the chart part XML)
    chart.title = "Sales updated"
    print(f"Chart title: {chart.title!r} @ {chart.anchor}")

    # -------------------------------------------------------------------------
    # §10  Table sheet — add_table, table.resize
    # -------------------------------------------------------------------------
    section("Table sheet")

    table_ws = wb.add_worksheet("Table")
    table_ws["A1"].value = "Item"
    table_ws["B1"].value = "Price"
    table_ws["A2"].value = "Widget"
    table_ws["B2"].value = 10
    table_ws["A3"].value = "Gadget"
    table_ws["B3"].value = 20
    table_ws["A4"].value = "Gizmo"
    table_ws["B4"].value = 15

    # ws.add_table(cell_range, columns, *, name=None, display_name=None, write_header=False) -> Table
    # cell_range = initial data range including headers; the header cells must already hold
    # the column names, unless write_header=True writes them.
    table = table_ws.add_table("A1:B2", ["Item", "Price"], name="TutorialTable")

    # table.resize(ref: str) -> None — grow or shrink the table range (and autoFilter).
    table.resize("A1:B4")
    print(f"Table {table.name!r} ref={table.ref!r} columns={table.columns}")

    # -------------------------------------------------------------------------
    # §11  Links sheet — hyperlink.url, hyperlink.location, hyperlink.display
    # -------------------------------------------------------------------------
    section("Links sheet")

    links = wb.add_worksheet("Links")
    links["A1"].value = "Visit"

    # cell.hyperlink.url: str | None — external URL (worksheet relationship)
    # cell.hyperlink.display: str | None — tooltip / screen-tip text in Excel
    links["A1"].hyperlink.url = "https://xlsxedit.jonasruilong.com"
    links["A1"].hyperlink.display = "Example site"
    # The link is clickable without extra styling. For blue underline in Excel:
    # links["A1"].apply_style(underline=True, font_color="0563C1")
    print(f"External hyperlink: {links['A1'].hyperlink.url!r}")

    # cell.hyperlink.location: str | None — jump to another sheet/cell (no URL relationship)
    links["B1"].value = "Table"
    links["B1"].hyperlink.location = "Table!A1"
    links["B1"].hyperlink.display = "Go to table"
    print(f"Internal hyperlink: {links['B1'].hyperlink.location!r}")

    # -------------------------------------------------------------------------
    # §12  Photos sheet — add_image, insert_image_at_placeholder, replace_image
    # -------------------------------------------------------------------------
    section("Photos sheet — images")

    photos = wb.add_worksheet("Photos")
    photos.column_dimensions["B"].width = 24
    photos["A1"].value = "Photos"
    photos["A1"].apply_style(bold=True, font_size=12, bg_color="FFD9E1F2")

    print(f"Using image: {TUTORIAL_IMAGE.name}")

    # ws.add_image(image_path, *, anchor, width=, height=, max_width=200, max_height=200, name=) -> Picture
    # Omitting width/height auto-fits using the file's real dimensions (aspect ratio kept).
    pic = photos.add_image(TUTORIAL_IMAGE, anchor="B2", name="coco")
    print(f"Added picture {pic.name!r} at {pic.anchor} ({pic.width}x{pic.height}px)")

    # wb.insert_image_at_placeholder(..., *, width=None, height=None, max_width=200, max_height=200)
    photos["D6"].value = "{logo}"
    n_img = wb.insert_image_at_placeholder("{logo}", TUTORIAL_IMAGE)
    print(f"insert_image_at_placeholder replacements: {n_img}")

    # wb.replace_image(name: str, image_path: str | Path) -> int
    n_named = wb.replace_image("coco", TUTORIAL_IMAGE)
    print(f"replace_image on named picture: {n_named}")

    for image in photos.images:
        print(f"  image {image.name!r} @ {image.anchor}")

    print(f"Sheets now: {wb.sheetnames}")

    # -------------------------------------------------------------------------
    # §13  Read-only fixture demos (console only, not saved to tutorial.xlsx)
    # -------------------------------------------------------------------------
    section("Read features from fixtures (console only)")

    # These open other workbooks from fixtures/ to show read APIs.
    # Nothing here is written into tutorial.xlsx — skip this section if you only
    # care about the output file.

    fixtures = HERE.parent / "fixtures"

    merged_path = fixtures / "MergedCells.xlsx"
    if merged_path.is_file():
        mw = Workbook.open(merged_path)
        mws = mw["Sheet1"]
        # Writing to B1 inside A1:C1 redirects to merge anchor A1.
        mws["B1"].value = "written via B1"
        print(f"MergedCells A1 after B1 write: {mws['A1'].value!r}")
        print(f"  merge ranges: {mws.merged_ranges}")

    charts_path = fixtures / "ChartsAndTables.xlsx"
    if charts_path.is_file():
        cw = Workbook.open(charts_path)
        for chart in cw["bar chart"].charts:
            print(f"  chart {chart.name!r} @ {chart.anchor}")
        for table in cw["Table"].tables:
            print(f"  table {table.name!r} {table.ref} columns={table.columns}")

    types_path = fixtures / "DifferentCellTypes.xlsx"
    if types_path.is_file():
        tw = Workbook.open(types_path)
        b5 = tw["Sheet1"]["B5"]
        print(f"  DifferentCellTypes B5: {b5.value} (is_date={b5.style.is_date})")

    text_path = fixtures / "TextSytle.xlsx"
    if text_path.is_file():
        tw = Workbook.open(text_path)
        print(f"  TextSytle A1 bold: {tw['Sheet1']['A1'].style.bold}")

    # -------------------------------------------------------------------------
    # §14  Save to disk
    # -------------------------------------------------------------------------
    section("Save")

    # wb.remove_worksheet(name, *, keep_unreachable=False) -> None — delete a tab and the parts
    # only it reached (cannot remove the last sheet)
    scratch = wb.add_worksheet("Scratch")
    scratch["A1"].value = "will be removed"
    wb.remove_worksheet("Scratch")

    # wb.save(path, *, include_orphans: bool = False) -> None
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT)
    print(f"Saved → {OUTPUT}")

    # -------------------------------------------------------------------------
    # §15  Reopen and verify round-trip
    # -------------------------------------------------------------------------
    section("Reopen and verify round-trip")

    # Workbook.open(path) — same API introduced in §1.
    wb2 = Workbook.open(OUTPUT)
    assert wb2.sheetnames == [
        "Overview", "Invoice", "Formatting", "Chart", "Table", "Links", "Photos",
    ]
    assert wb2["Invoice"]["B2"].value == "Jonas Corp"
    assert wb2["Invoice"]["B5"].value == 888
    assert wb2["Invoice"]["B6"].value == 42
    assert wb2["Invoice"]["B6"].style.num_format == "$#,##0.00"
    assert isinstance(wb2["Invoice"]["B3"].value, datetime)
    assert wb2["Invoice"]["B7"].formula == "B5*B6"
    assert wb2["Invoice"]["C1"].value == "Note"
    assert wb2["Invoice"]["C2"].value == "from tutorial"
    assert wb2["Invoice"]["A9"].value == "Anchor"
    assert wb2["Invoice"]["B10"].value == "written via find+offset"
    assert wb2["Invoice"]["A11"].value == "Tag"
    assert wb2["Invoice"]["B11"].value == "via findall"
    # insert_columns at C shifted the second Tag from C12 → D12
    assert wb2["Invoice"]["D12"].value == "Tag"
    assert wb2["Invoice"]["E12"].value == "via findall"
    assert wb2["Formatting"]["B2"].style.bold
    assert wb2["Formatting"]["B3"].style.font_color == "FFFF0000"
    assert len(wb2["Formatting"].conditional_formatting) >= 2
    assert len(wb2["Chart"].charts) >= 1
    assert wb2["Chart"].charts[0].title == "Sales updated"
    assert wb2["Table"].tables[0].ref == "A1:B4"
    assert wb2["Links"]["A1"].hyperlink.url == "https://xlsxedit.jonasruilong.com"
    assert wb2["Links"]["B1"].hyperlink.location == "Table!A1"
    assert len(wb2["Photos"].images) >= 1
    assert len(wb2["Overview"].images) == 0
    print("Round-trip OK")


if __name__ == "__main__":
    main()
