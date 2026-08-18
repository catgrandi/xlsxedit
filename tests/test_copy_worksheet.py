"""Tests for Workbook.copy_worksheet."""

from __future__ import annotations

from pathlib import Path

import pytest

from xlsxedit import DuplicateWorksheetError, Workbook, WorksheetNotFoundError


def test_copy_worksheet_merges_styles_and_widths():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["C"].width = 12
    ws["A1"].value = "Report"
    ws["A1"].apply_style(bold=True, font_size=16)
    ws.merge_cells("A1:C1")
    ws["A4"].value = "Alpha"

    ws2 = wb.copy_worksheet("Sheet1", "Sheet1 (2)")
    assert wb.sheetnames == ["Sheet1", "Sheet1 (2)"]
    assert ws2["A1"].value == "Report"
    assert ws2["A1"].style.bold is True
    assert "A1:C1" in ws2.merged_ranges
    assert ws2.column_dimensions["A"].width == 18
    assert ws2.column_dimensions["C"].width == 12
    assert ws2["A4"].value == "Alpha"

    ws2["A1"].value = "Report copy"
    assert ws["A1"].value == "Report"


def test_copy_worksheet_hyperlink():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A10"].value = "Website"
    ws["A10"].hyperlink.url = "https://example.com/"
    ws2 = wb.copy_worksheet("Sheet1", "Copy")
    assert ws2["A10"].hyperlink.url == "https://example.com/"
    assert ws["A10"].hyperlink.url == "https://example.com/"


def test_copy_worksheet_table_unique_name():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "Day"
    ws["B1"].value = "N"
    ws["A2"].value = "Mon"
    ws["B2"].value = 1
    t1 = ws.add_table("A1:B2", ["Day", "N"], name="Table1")
    ws2 = wb.copy_worksheet("Sheet1", "Copy")
    assert t1.name == "Table1"
    names = {t.name for t in ws2.tables}
    assert names
    assert "Table1" not in names


def test_copy_worksheet_roundtrip(tmp_path: Path):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.merge_cells("A1:C1")
    ws["A1"].value = "Title"
    wb.copy_worksheet("Sheet1", "Copy")
    out = tmp_path / "copy.xlsx"
    wb.save(out)
    wb2 = Workbook.open(out)
    assert wb2.sheetnames == ["Sheet1", "Copy"]
    assert "A1:C1" in wb2["Copy"].merged_ranges
    assert wb2["Copy"]["A1"].value == "Title"


def test_copy_worksheet_missing_source():
    wb = Workbook.create()
    with pytest.raises(WorksheetNotFoundError):
        wb.copy_worksheet("Nope", "Copy")


def test_copy_worksheet_duplicate_name():
    wb = Workbook.create()
    with pytest.raises(DuplicateWorksheetError):
        wb.copy_worksheet("Sheet1", "Sheet1")
