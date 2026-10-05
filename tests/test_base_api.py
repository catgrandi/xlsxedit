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


def test_add_chart_anchor_to_below_from(tmp_path: Path):
    """Anchoring below the template's default ``to`` row must not invert the box."""
    from xlsxedit.drawing import XDR_NS

    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "Day"
    ws["B1"].value = "N"
    ws["A2"].value = "Mon"
    ws["B2"].value = 1
    ws.add_chart("bar", anchor="A18", data_range="A1:B2", title="Late")
    out = tmp_path / "chart_anchor.xlsx"
    wb.save(out)

    import zipfile

    with zipfile.ZipFile(out) as z:
        drawing = etree.fromstring(z.read("xl/drawings/drawing1.xml"))
    from_row = int(drawing.find(f".//{{{XDR_NS}}}from/{{{XDR_NS}}}row").text)
    to_row = int(drawing.find(f".//{{{XDR_NS}}}to/{{{XDR_NS}}}row").text)
    from_col = int(drawing.find(f".//{{{XDR_NS}}}from/{{{XDR_NS}}}col").text)
    to_col = int(drawing.find(f".//{{{XDR_NS}}}to/{{{XDR_NS}}}col").text)
    assert from_row == 17
    assert to_row > from_row
    assert to_col > from_col


def test_add_chart_default_flush_offsets(tmp_path: Path):
    from xlsxedit.drawing import XDR_NS

    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "Day"
    ws["B1"].value = "N"
    ws["A2"].value = "Mon"
    ws["B2"].value = 1
    chart = ws.add_chart("bar", anchor="C19", data_range="A1:B2", title="Flush")
    assert chart.offset_x == 0
    assert chart.offset_y == 0
    assert chart.anchor == "C19"

    out = tmp_path / "chart_flush.xlsx"
    wb.save(out)

    import zipfile

    with zipfile.ZipFile(out) as z:
        drawing = etree.fromstring(z.read("xl/drawings/drawing1.xml"))
    fr = drawing.find(f".//{{{XDR_NS}}}from")
    to = drawing.find(f".//{{{XDR_NS}}}to")
    assert int(fr.find(f"{{{XDR_NS}}}col").text) == 2
    assert int(fr.find(f"{{{XDR_NS}}}row").text) == 18
    assert int(fr.find(f"{{{XDR_NS}}}colOff").text) == 0
    assert int(fr.find(f"{{{XDR_NS}}}rowOff").text) == 0
    assert int(to.find(f"{{{XDR_NS}}}col").text) == 2 + 6
    assert int(to.find(f"{{{XDR_NS}}}row").text) == 18 + 13
    assert int(to.find(f"{{{XDR_NS}}}colOff").text) == 0
    assert int(to.find(f"{{{XDR_NS}}}rowOff").text) == 0


def test_chart_offset_post_create_roundtrip(tmp_path: Path):
    from xlsxedit.drawing import EMU_PER_PIXEL, XDR_NS

    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "Day"
    ws["B1"].value = "N"
    ws["A2"].value = "Mon"
    ws["B2"].value = 1
    chart = ws.add_chart("bar", anchor="C19", data_range="A1:B2", title="Inset")
    assert chart.offset_x == 0
    assert chart.offset_y == 0
    chart.offset_x = 10
    chart.offset_y = 8
    chart.anchor = "D20"
    assert chart.anchor == "D20"
    assert chart.offset_x == 10
    assert chart.offset_y == 8

    out = tmp_path / "chart_offset.xlsx"
    wb.save(out)
    wb2 = Workbook.open(out)
    chart2 = wb2["Sheet1"].charts[0]
    assert chart2.anchor == "D20"
    assert chart2.offset_x == 10
    assert chart2.offset_y == 8

    import zipfile

    with zipfile.ZipFile(out) as z:
        drawing = etree.fromstring(z.read("xl/drawings/drawing1.xml"))
    fr = drawing.find(f".//{{{XDR_NS}}}from")
    to = drawing.find(f".//{{{XDR_NS}}}to")
    assert int(fr.find(f"{{{XDR_NS}}}colOff").text) == 10 * EMU_PER_PIXEL
    assert int(fr.find(f"{{{XDR_NS}}}rowOff").text) == 8 * EMU_PER_PIXEL
    # span preserved when moving anchor C19→D20
    assert int(to.find(f"{{{XDR_NS}}}col").text) - int(fr.find(f"{{{XDR_NS}}}col").text) == 6
    assert int(to.find(f"{{{XDR_NS}}}row").text) - int(fr.find(f"{{{XDR_NS}}}row").text) == 13


def test_add_chart_to_anchor(tmp_path: Path):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "Day"
    ws["B1"].value = "N"
    ws["A2"].value = "Mon"
    ws["B2"].value = 1
    chart = ws.add_chart(
        "bar",
        anchor="C19",
        to_anchor="G32",
        data_range="A1:B2",
        title="Sized",
    )
    assert chart.anchor == "C19"
    assert chart.to_anchor == "G32"
    chart.to_anchor = "F30"
    assert chart.to_anchor == "F30"

    out = tmp_path / "chart_to.xlsx"
    wb.save(out)
    wb2 = Workbook.open(out)
    assert wb2["Sheet1"].charts[0].to_anchor == "F30"


def test_add_image_offset(tmp_path: Path):
    from xlsxedit.drawing import EMU_PER_PIXEL, XDR_NS

    # 1x1 PNG
    png = (
        b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01"
        b"\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDATx\x9cc\xf8\x0f\x00"
        b"\x00\x01\x01\x00\x05\x18\xd8N\x00\x00\x00\x00IEND\xaeB`\x82"
    )
    img = tmp_path / "dot.png"
    img.write_bytes(png)

    wb = Workbook.create()
    ws = wb["Sheet1"]
    pic = ws.add_image(img, anchor="H1", width=90, height=90, offset_x=5, offset_y=2)
    assert pic.offset_x == 5
    assert pic.offset_y == 2
    pic.offset_x = 7

    out = tmp_path / "img_offset.xlsx"
    wb.save(out)
    wb2 = Workbook.open(out)
    pic2 = wb2["Sheet1"].images[0]
    assert pic2.offset_x == 7
    assert pic2.offset_y == 2

    import zipfile

    with zipfile.ZipFile(out) as z:
        drawing = etree.fromstring(z.read("xl/drawings/drawing1.xml"))
    fr = drawing.find(f".//{{{XDR_NS}}}from")
    assert int(fr.find(f"{{{XDR_NS}}}colOff").text) == 7 * EMU_PER_PIXEL
    assert int(fr.find(f"{{{XDR_NS}}}rowOff").text) == 2 * EMU_PER_PIXEL


def test_add_table_and_resize(tmp_path: Path):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "Item"
    ws["B1"].value = "Price"
    ws["A2"].value = "Widget"
    ws["B2"].value = 10
    table = ws.add_table("A1:B2", ["Item", "Price"], name="Prices")
    assert table.name == "Prices"
    assert table.ref == "A1:B2"
    table.resize("A1:B5")
    assert table.ref == "A1:B5"
    wb.save(tmp_path / "table.xlsx")
    wb2 = Workbook.open(tmp_path / "table.xlsx")
    assert wb2["Sheet1"].tables[0].ref == "A1:B5"


def test_every_public_name_is_exported_from_the_package_root():
    import xlsxedit
    from xlsxedit.drawing import DrawingObject
    from xlsxedit.exceptions import TableError
    from xlsxedit.styles import Color

    assert (xlsxedit.TableError, xlsxedit.Color, xlsxedit.DrawingObject) == (
        TableError,
        Color,
        DrawingObject,
    )
    assert all(hasattr(xlsxedit, name) for name in xlsxedit.__all__)
    library_errors = {
        name
        for name, value in vars(xlsxedit.exceptions).items()
        if isinstance(value, type) and issubclass(value, xlsxedit.XlsxeditError)
    }
    assert library_errors <= set(xlsxedit.__all__)
