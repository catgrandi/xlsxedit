"""Tests for Workbook.copy_worksheet."""

from __future__ import annotations

from pathlib import Path

import pytest

from xlsxedit import DuplicateWorksheetError, Workbook, WorksheetNotFoundError
from xlsxedit.drawing import A16_NS, chart_references, drawing_parts_for_worksheet
from xlsxedit.sheet_clone import repoint_sheet_refs, sheet_prefix
from tests.conftest import INSPECT_FIXTURES
from tests.schema_validate import assert_valid_package


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


def _chart_refs(wb: Workbook, sheet: str) -> list[str]:
    root = wb[sheet].charts[0]._chart_root()
    return [elm.text for elm in chart_references(root)]


def _creation_ids(wb: Workbook, sheet: str) -> list[str]:
    root = drawing_parts_for_worksheet(wb[sheet])[0].element
    return [elm.get("id") for elm in root.iter(f"{{{A16_NS}}}creationId")]


def test_copy_worksheet_charts_plot_the_copy():
    wb = Workbook.open(INSPECT_FIXTURES["ChartsAndTables"])
    source_refs = _chart_refs(wb, "bar chart")
    assert source_refs == ["'bar chart'!$B$2", "'bar chart'!$A$3:$A$6", "'bar chart'!$B$3:$B$6"]
    wb.copy_worksheet("bar chart", "Summary")
    assert _chart_refs(wb, "Summary") == ["Summary!$B$2", "Summary!$A$3:$A$6", "Summary!$B$3:$B$6"]
    assert _chart_refs(wb, "bar chart") == source_refs
    ids = [obj.id for obj in wb["bar chart"].drawing_objects]
    assert [obj.id for obj in wb["Summary"].drawing_objects] == ids
    assert set(_creation_ids(wb, "Summary")).isdisjoint(_creation_ids(wb, "bar chart"))
    assert_valid_package(wb)


def test_copy_worksheet_keeps_references_to_other_sheets():
    wb = Workbook.open(INSPECT_FIXTURES["ChartsAndTables"])
    root = wb["bar chart"].charts[0]._chart_root()
    chart_references(root)[1].text = "'line chart'!$A$3:$A$12"
    wb.copy_worksheet("bar chart", "bar chart (2)")
    assert _chart_refs(wb, "bar chart (2)") == [
        "'bar chart (2)'!$B$2",
        "'line chart'!$A$3:$A$12",
        "'bar chart (2)'!$B$3:$B$6",
    ]


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("Sheet1!$A$1:$A$5", "'Data 2'!$A$1:$A$5"),
        ("'Sheet1'!$A$1", "'Data 2'!$A$1"),
        ("sheet1!#REF!", "'Data 2'!#REF!"),
        ("Other!$A$1", "Other!$A$1"),
        ("'Sheet1:Sheet3'!$A$1", "'Sheet1:Sheet3'!$A$1"),
        ("[1]Sheet1!$A$1", "[1]Sheet1!$A$1"),
        ("(Sheet1!$A$1,Sheet10!$B$1)", "('Data 2'!$A$1,Sheet10!$B$1)"),
    ],
)
def test_repoint_sheet_refs(text: str, expected: str):
    assert repoint_sheet_refs(text, "Sheet1", "Data 2") == expected


@pytest.mark.parametrize(
    ("name", "prefix"),
    [("Sales", "Sales"), ("Q1_2024.v2", "Q1_2024.v2"), ("A1", "'A1'"), ("R1C1", "'R1C1'"),
     ("TRUE", "'TRUE'"), ("2024", "'2024'"), ("Bob's", "'Bob''s'")],
)  # fmt: skip
def test_sheet_prefix_quotes_what_would_not_read_as_a_name(name: str, prefix: str):
    assert sheet_prefix(name) == prefix
