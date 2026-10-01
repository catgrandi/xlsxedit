"""Build a workbook that exercises every public write API (for manual Excel inspection).

Run ``python -m tests.smoke_workbook`` to write ``tutorial/output/api_smoke.xlsx`` and
open it in Excel. ``test_smoke_workbook`` builds the same workbook but saves to ``tmp_path``.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from xlsxedit import Workbook

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "tutorial" / "output" / "api_smoke.xlsx"
SMOKE_IMAGE = ROOT / "assets" / "coco-happy-swiss-nature.jpg"


def build_smoke_workbook() -> Workbook:
    """Return a workbook with one sheet per feature area."""
    if not SMOKE_IMAGE.is_file():
        raise FileNotFoundError(f"Smoke test image not found: {SMOKE_IMAGE}")

    wb = Workbook()

    # --- Index ---
    index = wb.rename_worksheet("Sheet1", "Index")
    index.merge_cells("A1:D1")
    index["A1"].value = "xlsxedit API smoke workbook"
    index["A1"].apply_style(bold=True, font_size=16, horizontal_align="center")
    index.row_dimensions[1].height = 28
    index.column_dimensions["A"].width = 22
    sheets_tested = [
        "Values", "Layout", "Styles", "Formulas", "SAR", "Links",
        "CF", "Chart", "Table", "Photos",
    ]
    for i, name in enumerate(sheets_tested, start=3):
        index[f"A{i}"].value = name

    # --- Values ---
    values = wb.add_worksheet("Values")
    values["A1"].value = "Type"
    values["B1"].value = "Example"
    values["A2"].value = "str"
    values["B2"].value = "hello"
    values["A3"].value = "int"
    values["B3"].value = 42
    values["A4"].value = "float"
    values["B4"].value = 3.14
    values["A5"].value = "bool"
    values["B5"].value = True
    values["A6"].value = "datetime"
    values["B6"].value = datetime(2026, 7, 9)
    values["B6"].apply_date_format()
    values["A7"].value = "currency"
    values["B7"].value = 99.5
    values["B7"].apply_number_format("$#,##0.00")
    values["A8"].value = "percent"
    values["B8"].value = 0.25
    values["B8"].apply_number_format("0.00%")
    values["A9"].value = "clear"
    values["B9"].value = "gone"
    values["B9"].clear()

    # --- Layout ---
    layout = wb.add_worksheet("Layout")
    layout.column_dimensions["B"].width = 18
    layout.row_dimensions[2].height = 24
    layout.merge_cells("A1:C1")
    layout["B1"].value = "written via B1"
    layout.unmerge_cells("A1:C1")
    layout["A3"].value = "unmerged"

    # --- Styles ---
    styles = wb.add_worksheet("Styles")
    styles["B2"].value = "bold"
    styles["B2"].apply_style(bold=True)
    styles["B3"].value = "red"
    styles["B3"].apply_style(font_color="FF0000")
    styles["B4"].value = "large"
    styles["B4"].apply_style(font_size=16)
    styles["B5"].value = "yellow bg"
    styles["B5"].apply_style(bg_color="FFFF00")
    styles["B6"].value = "centered"
    styles["B6"].apply_style(horizontal_align="center", vertical_align="center")
    styles["B7"].value = "italic"
    styles["B7"].apply_style(italic=True)
    styles["B8"].value = "underline"
    styles["B8"].apply_style(underline=True)
    styles["B9"].value = "arial"
    styles["B9"].apply_style(font_name="Arial")

    # --- Formulas ---
    formulas = wb.add_worksheet("Formulas")
    formulas["A1"].value = 10
    formulas["B1"].value = 20
    formulas["C1"].formula = "=A1+B1"

    # --- SAR ---
    sar = wb.add_worksheet("SAR")
    sar["B1"].value = "{client}"
    sar["B2"].value = "{qty}"
    sar["B3"].value = "{ship}"
    sar["B4"].value = "Item {item} ok"
    wb.replace("{client}", "Acme", value_type="text")
    wb.replace("{qty}", 7, value_type="number")
    wb.replace("{ship}", "2026-07-09", value_type="date")
    sar.replace("{item}", "Widget", value_type="text")
    sar["B5"].value = "partial old text"
    sar["B5"].replace("old", "new")

    # --- Links ---
    links = wb.add_worksheet("Links")
    links["A1"].value = "Visit"
    links["A1"].hyperlink.url = "https://xlsxedit.jonasruilong.com"
    links["A1"].hyperlink.display = "Example site"

    # --- CF ---
    cf = wb.add_worksheet("CF")
    for i, score in enumerate([10, 50, 90], start=2):
        cf[f"A{i}"].value = f"Row {i}"
        cf[f"B{i}"].value = score
    cf.add_color_scale_formatting("B2:B4")
    cf["A6"].value = 5
    cf["A7"].value = -1
    cf.add_conditional_formatting("A6:A7", operator="greaterThan", formula="0")  # A6 highlights green

    # --- Chart ---
    chart_ws = wb.add_worksheet("Chart")
    chart_ws["A1"].value = "Item"
    chart_ws["B1"].value = "Qty"
    chart_ws["A2"].value = "Apples"
    chart_ws["B2"].value = 3
    chart_ws["A3"].value = "Pears"
    chart_ws["B3"].value = 5
    chart = chart_ws.add_chart("bar", anchor="D2", data_range="A1:B3", title="Sales")
    chart.title = "Sales updated"

    # --- Table ---
    table_ws = wb.add_worksheet("Table")
    table_ws["A1"].value = "Item"
    table_ws["B1"].value = "Price"
    table_ws["A2"].value = "Widget"
    table_ws["B2"].value = 10
    table = table_ws.add_table("A1:B2", ["Item", "Price"], name="SmokeTable")
    table.resize("A1:B4")

    # --- Photos ---
    photos = wb.add_worksheet("Photos")
    photos.add_image(SMOKE_IMAGE, anchor="B2", name="autofit")
    photos.add_image(SMOKE_IMAGE, anchor="E2", width=80, height=80, name="fixed")
    photos["G6"].value = "{logo}"
    wb.insert_image_at_placeholder("{logo}", SMOKE_IMAGE)
    wb.replace_image("autofit", SMOKE_IMAGE)

    # --- Sheets API (throwaway tab removed) ---
    temp = wb.add_worksheet("TempRemove")
    temp["A1"].value = "delete me"
    wb.remove_worksheet("TempRemove")

    return wb


def main() -> None:
    wb = build_smoke_workbook()
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT)
    print(f"Saved to {OUTPUT}")
    print("Open in Excel to verify visually.")


if __name__ == "__main__":
    main()
