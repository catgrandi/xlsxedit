"""Tests for base-first APIs (layout, styles, merges, hyperlinks, CF, charts, tables)."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pytest
from lxml import etree

from xlsxedit import Workbook
from xlsxedit.drawing import CHART_NS
from xlsxedit.opc.constants import SML_NS
from tests.conftest import INSPECT_FIXTURES


def test_column_and_row_dimensions(tmp_path: Path):
    wb = Workbook.open(INSPECT_FIXTURES["CustomCellSize"])
    ws = wb["Sheet1"]
    assert ws.column_dimensions["C"].width == 32.0
    assert ws.row_dimensions[4].height == 47.0

    ws.column_dimensions["D"].width = 20.0
    ws.row_dimensions[5].height = 30.0
    out = tmp_path / "dims.xlsx"
    wb.save(out)

    wb2 = Workbook.open(out)
    assert wb2["Sheet1"].column_dimensions["D"].width == 20.0
    assert wb2["Sheet1"].row_dimensions[5].height == 30.0


def test_apply_date_format_on_create(tmp_path: Path):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "{ship}"
    wb.replace("{ship}", "2026-07-23", value_type="date")
    assert isinstance(ws["A1"].value, datetime)
    assert ws["A1"].style.is_date
    wb.save(tmp_path / "date.xlsx")
    wb2 = Workbook.open(tmp_path / "date.xlsx")
    assert isinstance(wb2["Sheet1"]["A1"].value, datetime)


def test_apply_number_format(tmp_path: Path):
    from lxml import etree

    from xlsxedit.opc.constants import SML_NS

    wb = Workbook.create()
    cell = wb["Sheet1"]["A1"]
    cell.value = 0.25
    cell.apply_number_format("0.00%")
    assert cell.style.num_format == "0.00%"
    out = tmp_path / "pct.xlsx"
    wb.save(out)

    styles = etree.fromstring(Workbook.open(out)._workbook_part.styles_part.blob)
    fonts = styles.find(f"{{{SML_NS}}}fonts")
    num_fmts = styles.find(f"{{{SML_NS}}}numFmts")
    assert num_fmts is not None
    assert list(styles).index(num_fmts) < list(styles).index(fonts)


def test_merge_and_unmerge_cells(tmp_path: Path):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.merge_cells("A1:C1")
    assert "A1:C1" in ws.merged_ranges
    ws["B1"].value = "merged"
    assert ws["A1"].value == "merged"
    ws.unmerge_cells("A1:C1")
    assert ws.merged_ranges == []
    wb.save(tmp_path / "merge.xlsx")


def _assert_hyperlinks_after_sheet_data(blob: bytes) -> None:
    ws = etree.fromstring(blob)
    ns = f"{{{SML_NS}}}"
    sheet_data = ws.find(f"{ns}sheetData")
    hyperlinks = ws.find(f"{ns}hyperlinks")
    if hyperlinks is not None:
        assert sheet_data is not None
        assert list(ws).index(hyperlinks) > list(ws).index(sheet_data)


def test_hyperlink_set_and_get(tmp_path: Path):
    wb = Workbook.create()
    cell = wb["Sheet1"]["A1"]
    cell.value = "Visit"
    cell.hyperlink.url = "https://xlsxedit.jonasruilong.com"
    cell.hyperlink.display = "Example"
    assert cell.hyperlink.url == "https://xlsxedit.jonasruilong.com"
    assert cell.hyperlink.display == "Example"
    out = tmp_path / "link.xlsx"
    wb.save(out)
    wb2 = Workbook.open(out)
    assert wb2["Sheet1"]["A1"].hyperlink.url == "https://xlsxedit.jonasruilong.com"
    _assert_hyperlinks_after_sheet_data(wb2["Sheet1"]._part.blob)


def test_hyperlink_internal_location(tmp_path: Path):
    wb = Workbook.create()
    wb.add_worksheet("Target")
    cell = wb["Sheet1"]["A1"]
    cell.value = "Go"
    cell.hyperlink.location = "Target!A1"
    cell.hyperlink.display = "Jump to Target"
    assert cell.hyperlink.url is None
    assert cell.hyperlink.location == "Target!A1"
    assert cell.hyperlink.display == "Jump to Target"
    out = tmp_path / "internal.xlsx"
    wb.save(out)
    wb2 = Workbook.open(out)
    assert wb2["Sheet1"]["A1"].hyperlink.location == "Target!A1"
    assert wb2["Sheet1"]["A1"].hyperlink.url is None
    _assert_hyperlinks_after_sheet_data(wb2["Sheet1"]._part.blob)

    wb2["Sheet1"]["A1"].hyperlink.url = "https://xlsxedit.jonasruilong.com"
    assert wb2["Sheet1"]["A1"].hyperlink.location is None
    assert wb2["Sheet1"]["A1"].hyperlink.url == "https://xlsxedit.jonasruilong.com"


def test_conditional_formatting_read_and_add(tmp_path: Path):
    wb = Workbook.open(INSPECT_FIXTURES["ConditionalFormatting"])
    ws = wb["Sheet1"]
    blocks = ws.conditional_formatting
    assert len(blocks) >= 1
    assert blocks[0].rules[0].type in ("colorScale", "dataBar", "cellIs")

    wb2 = Workbook.create()
    ws2 = wb2["Sheet1"]
    cf = ws2.add_conditional_formatting("A1:A5", operator="greaterThan", formula="0")
    assert cf.rules[0].type == "cellIs"
    out = tmp_path / "cf.xlsx"
    wb2.save(out)
    wb3 = Workbook.open(out)
    rule = wb3["Sheet1"].conditional_formatting[0].rules[0]
    assert rule._element.get("dxfId") is not None


def test_add_chart_and_set_title(tmp_path: Path):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "Item"
    ws["B1"].value = "Qty"
    ws["A2"].value = "Apples"
    ws["B2"].value = 3
    ws["A3"].value = "Pears"
    ws["B3"].value = 5
    chart = ws.add_chart("bar", anchor="D2", data_range="A1:B3", title="Sales")
    assert chart.title == "Sales"
    assert len(ws.charts) == 1
    chart.title = "Updated"
    assert chart.title == "Updated"
    out = tmp_path / "chart.xlsx"
    wb.save(out)
    wb2 = Workbook.open(out)
    assert len(wb2["Sheet1"].charts) == 1

    chart_part = wb2["Sheet1"].charts[0]._chart_part
    assert chart_part is not None
    root = etree.fromstring(chart_part.blob)
    ns = f"{{{CHART_NS}}}"
    a_ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

    for f_elm in root.iter(f"{ns}f"):
        formula = f_elm.text or ""
        if ":" in formula and "!" in formula:
            assert formula.count("!") == 1, formula

    title = root.find(f".//{ns}chart/{ns}title")
    assert title is not None
    assert etree.QName(title[0]).localname == "tx"
    assert title.find(f"{ns}tx/{ns}rich") is not None
    assert title.find(f".//{a_ns}t").text == "Updated"

    str_refs = list(root.iter(f"{ns}strRef"))
    assert len(str_refs) == 2
    assert str_refs[0].find(f"{ns}f").text == "Sheet1!$B$1"
    assert str_refs[0].find(f"{ns}strCache/{ns}ptCount").get("val") == "1"
    assert str_refs[1].find(f"{ns}f").text == "Sheet1!$A$2:$A$3"
    assert str_refs[1].find(f"{ns}strCache/{ns}ptCount").get("val") == "2"

    num_ref = root.find(f".//{ns}numRef")
    assert num_ref.find(f"{ns}f").text == "Sheet1!$B$2:$B$3"
    assert num_ref.find(f"{ns}numCache/{ns}ptCount").get("val") == "2"


def test_add_table_and_resize(tmp_path: Path):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "Item"
    ws["B1"].value = "Price"
    ws["A2"].value = "Widget"
    ws["B2"].value = 10
    table = ws.add_table("A1:B2", ["Item", "Price"], name="T1")
    assert table.name == "T1"
    assert table.ref == "A1:B2"
    table.resize("A1:B5")
    assert table.ref == "A1:B5"
    wb.save(tmp_path / "table.xlsx")
    wb2 = Workbook.open(tmp_path / "table.xlsx")
    assert wb2["Sheet1"].tables[0].ref == "A1:B5"
