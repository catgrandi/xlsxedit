"""Drawing and chart parts: live trees, anchors, sizes, series caches, ids, the object walker."""

from __future__ import annotations

import io
import random
import re
import warnings
from datetime import date
from pathlib import Path

import pytest
from lxml import etree

from tests.conftest import INSPECT_FIXTURES
from tests.preservation import assert_preserved, read_pkg, worksheet_members
from tests.schema_validate import SchemaWarning, assert_valid_package
from xlsxedit import Workbook
from xlsxedit.drawing import (
    A_NS,
    CHART_NS,
    EMU_PER_PIXEL,
    XDR_NS,
    _Axis,
    _RowWalker,
    drawing_parts_for_worksheet,
    sheet_axis,
)
from xlsxedit.opc.constants import RT, SML_NS
from xlsxedit.parts import ChartPart, DrawingPart

PNG = (
    b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01"
    b"\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDATx\x9cc\xf8\x0f\x00"
    b"\x00\x01\x01\x00\x05\x18\xd8N\x00\x00\x00\x00IEND\xaeB`\x82"
)

_XDR = f"{{{XDR_NS}}}"
_A = f"{{{A_NS}}}"
_C = f"{{{CHART_NS}}}"
_A16_CREATION_ID = "{http://schemas.microsoft.com/office/drawing/2014/main}creationId"
_GUID = re.compile(r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}")


@pytest.fixture
def png(tmp_path: Path) -> Path:
    path = tmp_path / "dot.png"
    path.write_bytes(PNG)
    return path


def _reopen(wb: Workbook) -> Workbook:
    buf = io.BytesIO()
    wb.save(buf)
    return Workbook.open(io.BytesIO(buf.getvalue()))


def _drawing_xml(wb: Workbook, member: str = "xl/drawings/drawing1.xml") -> etree._Element:
    return etree.fromstring(read_pkg(wb)[member])


def _marker(anchor: etree._Element, corner: str) -> tuple[int, int, int, int]:
    marker = anchor.find(f"{_XDR}{corner}")
    return tuple(int(marker.findtext(f"{_XDR}{tag}")) for tag in ("col", "colOff", "row", "rowOff"))


def _drawing_payload(member: str) -> bool:
    """A drawing, chart or media part; their ``.rels`` are regenerated on every save."""
    return member.startswith(("xl/drawings/", "xl/charts/", "xl/media/")) and "/_rels/" not in member


def _chart_sheet(wb: Workbook, png: Path):
    ws = wb["Sheet1"]
    ws["A1"].value = "Item"
    ws["B1"].value = "Qty"
    ws["A2"].value = "Apples"
    ws["B2"].value = 3
    ws.add_chart("bar", anchor="D2", data_range="A1:B2", title="Sales")
    ws.add_image(png, anchor="H2", name="logo")
    return ws


# --- one live tree per part -------------------------------------------------


def test_registered_part_classes():
    wb = Workbook.open(INSPECT_FIXTURES["ChartsAndTables"])
    ws = wb["bar chart"]
    assert isinstance(drawing_parts_for_worksheet(ws)[0], DrawingPart)
    assert isinstance(ws.charts[0]._chart_part, ChartPart)


def test_chart_and_picture_edits_on_one_sheet_both_survive(png: Path):
    wb = Workbook.create()
    _chart_sheet(wb, png)
    wb = _reopen(wb)

    ws = wb["Sheet1"]
    ch = ws.charts[0]
    pic = ws.images[0]
    ch.anchor = "Z50"
    pic.anchor = "B40"

    ws2 = _reopen(wb)["Sheet1"]
    assert ws2.charts[0].anchor == "Z50"
    assert ws2.images[0].anchor == "B40"


def test_proxies_share_the_live_tree(png: Path):
    wb = Workbook.create()
    ws = _chart_sheet(wb, png)
    first, second = ws.images[0], ws.images[0]
    first.name = "renamed"
    assert second.name == "renamed"
    ws.charts[0].offset_x = 4
    ws.charts[0].name = "chart renamed"
    assert ws.charts[0].offset_x == 4
    ws2 = _reopen(wb)["Sheet1"]
    assert ws2.images[0].name == "renamed"
    assert (ws2.charts[0].name, ws2.charts[0].offset_x) == ("chart renamed", 4)


def test_reading_does_not_reserialise_drawing_or_chart_parts():
    path = INSPECT_FIXTURES["ChartsAndTables"]
    wb = Workbook.open(path)
    for ws in wb.worksheets:
        ws.drawing_objects
        for chart in ws.charts:
            assert chart.anchor and chart.name
            chart.title
    before, after = read_pkg(path), read_pkg(wb)
    for member in filter(_drawing_payload, before):
        assert after[member] == before[member], member


@pytest.mark.parametrize("fixture", ["ChartsAndTables", "images"])
def test_unrelated_cell_edit_keeps_drawing_and_chart_parts_byte_identical(fixture: str):
    path = INSPECT_FIXTURES[fixture]
    before = read_pkg(path)
    wb = Workbook.open(path)
    ws = wb.worksheets[0]
    ws.images, ws.charts, ws.drawing_objects  # parse the drawing part
    ws["Z99"].value = 5

    report = assert_preserved(before, wb, expected_changed={worksheet_members(before)[ws.name]})
    after = read_pkg(wb)
    for member in filter(_drawing_payload, before):
        assert member not in report.reserialised
        assert after[member] == before[member], member


def test_insert_rows_on_another_sheet_keeps_drawing_and_chart_parts_byte_identical():
    path = INSPECT_FIXTURES["ChartsAndTables"]
    before = read_pkg(path)
    wb = Workbook.open(path)
    for ws in wb.worksheets:
        ws.charts, ws.images
    wb["Table"].insert_rows([[1, 2]], at_row=40)

    # An insert also asks Excel for a full recalculation, which is in the workbook part.
    report = assert_preserved(
        before, wb, expected_changed={worksheet_members(before)["Table"], "xl/workbook.xml"}
    )
    after = read_pkg(wb)
    drawing_members = list(filter(_drawing_payload, before))
    assert drawing_members
    for member in drawing_members:
        assert member not in report.reserialised
        assert after[member] == before[member], member


def test_failed_series_formula_leaves_the_chart_unchanged(png: Path):
    wb = Workbook.create()
    ws = _chart_sheet(wb, png)
    chart = ws.charts[0]
    before = read_pkg(wb)["xl/charts/chart1.xml"]
    with pytest.raises(ValueError, match="invalid chart formula"):
        chart.set_series_formula(2, "B2:B3", ws)
    assert read_pkg(wb)["xl/charts/chart1.xml"] == before


def test_chart_title_edit_changes_only_its_chart_part():
    path = INSPECT_FIXTURES["ChartsAndTables"]
    wb = Workbook.open(path)
    chart = wb["bar chart"].charts[0]
    chart.title = "Renamed"
    assert_preserved(path, wb, expected_changed={chart.partname.lstrip("/")})
    assert _reopen(wb)["bar chart"].charts[0].title == "Renamed"


@pytest.mark.parametrize(
    "edit",
    [
        lambda chart, ws: setattr(chart, "title", "bad\x01"),
        lambda chart, ws: setattr(chart, "title", 5),
        lambda chart, ws: chart.set_series_formula(1, "'bar\x01chart'!$A$2:$A$3", ws),
        lambda chart, ws: chart.set_series_formula(1, None, ws),
    ],
)
def test_a_refused_chart_edit_changes_nothing(edit):
    path = INSPECT_FIXTURES["ChartsAndTables"]
    wb = Workbook.open(path)
    ws = wb["bar chart"]
    with pytest.raises((TypeError, ValueError)):
        edit(ws.charts[0], ws)
    assert_preserved(path, wb)


def test_clearing_a_chart_title_leaves_a_valid_chart(png: Path):
    wb = Workbook.create()
    chart = _chart_sheet(wb, png).charts[0]
    chart.title = "Named"
    assert chart.title == "Named"
    chart.title = None
    assert chart.title is None
    title = etree.fromstring(read_pkg(wb)["xl/charts/chart1.xml"]).find(f".//{_C}title")
    assert title is not None and title.find(f"{_C}tx") is None
    assert_valid_package(wb)


# --- Picture.anchor / resize ------------------------------------------------


def test_picture_anchor_translates_two_cell_to():
    wb = Workbook.open(INSPECT_FIXTURES["images"])
    pic = wb["Sheet1"].images[0]
    anchor = pic._anchor
    assert anchor.tag == f"{_XDR}twoCellAnchor"
    from_before, to_before = _marker(anchor, "from"), _marker(anchor, "to")

    pic.anchor = "D10"

    anchor = _drawing_xml(_reopen(wb)).find(f"{_XDR}twoCellAnchor")
    from_after, to_after = _marker(anchor, "from"), _marker(anchor, "to")
    assert from_after == (3, from_before[1], 9, from_before[3])
    d_col, d_row = 3 - from_before[0], 9 - from_before[2]
    assert to_after == (to_before[0] + d_col, to_before[1], to_before[2] + d_row, to_before[3])
    assert to_after[2] > from_after[2]


def test_picture_anchor_outside_the_sheet_raises_and_changes_nothing():
    wb = Workbook.open(INSPECT_FIXTURES["images"])
    pic = wb["Sheet1"].images[0]
    with pytest.raises(ValueError, match="outside the sheet"):
        pic.anchor = "A1048576"
    assert pic.anchor == "B1"


def test_two_cell_picture_resize_recomputes_to():
    wb = Workbook.open(INSPECT_FIXTURES["images"])
    ws = wb["Sheet1"]
    for col in "BCDEF":
        ws.column_dimensions[col].width = 9.140625  # 64 px
    for row in range(1, 30):
        ws.row_dimensions[row].height = 15  # 20 px
    pic = ws.images[0]
    assert _marker(pic._anchor, "from") == (1, 0, 0, 1)

    pic.width = 100
    pic.height = 50

    anchor = _drawing_xml(_reopen(wb)).find(f"{_XDR}twoCellAnchor")
    # 100 px from column B at offset 0: B is 64 px, so 36 px into C.
    # 50 px (+1 EMU) from row 1: two 20 px rows, then 10 px into row 3.
    assert _marker(anchor, "to") == (2, 36 * EMU_PER_PIXEL, 2, 10 * EMU_PER_PIXEL + 1)
    sp_pr = anchor.find(f"{_XDR}pic/{_XDR}spPr")
    assert sp_pr.find(f"{_A}ext") is None
    ext = sp_pr.find(f"{_A}xfrm/{_A}ext")
    assert (ext.get("cx"), ext.get("cy")) == (str(100 * EMU_PER_PIXEL), str(50 * EMU_PER_PIXEL))
    assert_valid_package(wb)


def test_two_cell_picture_size_without_xfrm_reads_the_anchor_box():
    wb = Workbook.open(INSPECT_FIXTURES["images"])
    ws = wb["Sheet1"]
    for col in "BCD":
        ws.column_dimensions[col].width = 9.140625
    pic = ws.images[0]
    xfrm = pic._pic.find(f"{_XDR}spPr/{_A}xfrm")
    xfrm.getparent().remove(xfrm)
    pic.width = 90

    assert pic.width == 90
    pic._pic.find(f"{_XDR}spPr").remove(pic._pic.find(f"{_XDR}spPr/{_A}xfrm"))
    assert pic.width == 90  # from xdr:from / xdr:to


def test_resize_creates_xfrm_as_first_sppr_child(png: Path):
    wb = Workbook.create()
    pic = wb["Sheet1"].add_image(png, anchor="B2", width=40, height=40)
    sp_pr = pic._pic.find(f"{_XDR}spPr")
    sp_pr.remove(sp_pr.find(f"{_A}xfrm"))

    pic.height = 70

    sp_pr = _drawing_xml(_reopen(wb)).find(f"{_XDR}twoCellAnchor/{_XDR}pic/{_XDR}spPr")
    assert etree.QName(sp_pr[0]).localname == "xfrm"
    assert [etree.QName(e).localname for e in sp_pr[0]] == ["ext"]
    assert sp_pr[0][0].get("cy") == str(70 * EMU_PER_PIXEL)


def test_resize_drops_a_bare_ext_left_by_older_releases(png: Path):
    wb = Workbook.create()
    pic = wb["Sheet1"].add_image(png, anchor="B2", width=40, height=40)
    etree.SubElement(pic._pic.find(f"{_XDR}spPr"), f"{_A}ext", cx="1", cy="1")

    pic.width = 60

    sp_pr = _drawing_xml(_reopen(wb)).find(f"{_XDR}twoCellAnchor/{_XDR}pic/{_XDR}spPr")
    assert sp_pr.find(f"{_A}ext") is None
    assert sp_pr.find(f"{_A}xfrm/{_A}ext").get("cx") == str(60 * EMU_PER_PIXEL)
    assert_valid_package(wb)


def test_one_cell_picture_resize_sets_anchor_ext_and_xfrm(png: Path):
    wb = _walker_workbook(png)
    pic = next(pic for pic in wb["Sheet1"].images if pic.name == "one")
    pic.width = 120

    anchor = _drawing_xml(_reopen(wb)).find(f"{_XDR}oneCellAnchor")
    cx = str(120 * EMU_PER_PIXEL)
    assert anchor.find(f"{_XDR}ext").get("cx") == cx
    assert anchor.find(f"{_XDR}pic/{_XDR}spPr/{_A}xfrm/{_A}ext").get("cx") == cx
    assert next(p for p in _reopen(wb)["Sheet1"].images if p.name == "one").width == 120


def test_an_added_picture_moves_with_its_cells_and_keeps_its_size(png: Path):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.column_dimensions["B"].width = 9.140625  # 64 px
    ws.row_dimensions[3].height = 30  # 40 px
    pic = ws.add_image(png, anchor="B2", width=100, height=50, offset_x=4, offset_y=2)

    anchor = _drawing_xml(_reopen(wb)).find(f"{_XDR}twoCellAnchor")
    assert anchor.get("editAs") == "oneCell"
    assert pic.placement == "oneCell"
    assert _marker(anchor, "from") == (1, 4 * EMU_PER_PIXEL, 1, 2 * EMU_PER_PIXEL)
    # 4 + 100 px from B: 64 px of B, then 40 px into C; 2 + 50 px from row 2: 20 px of
    # row 2, 30 px into the 40 px row 3 (Excel's default 15 pt row is 20 px here).
    assert _marker(anchor, "to") == (2, 40 * EMU_PER_PIXEL, 2, 32 * EMU_PER_PIXEL)
    assert (pic.width, pic.height) == (100, 50)
    size = _size(_box(ws, pic._anchor))

    ws.insert_rows([[]] * 2, at_row=3)
    assert _marker(pic._anchor, "from")[2] == 1
    assert _size(_box(ws, pic._anchor)) == size
    assert_valid_package(wb)


def test_a_picture_past_the_edge_of_the_sheet_adds_nothing(png: Path):
    wb = Workbook.create()
    before = read_pkg(wb)
    with pytest.raises(ValueError, match="would pass the edge of the sheet"):
        wb["Sheet1"].add_image(png, anchor="A1048576", width=40, height=40)
    for anchor in ("A0", "XFE1", "B1048577"):
        with pytest.raises(ValueError, match="puts a corner outside the sheet"):
            wb["Sheet1"].add_image(png, anchor=anchor)
    assert read_pkg(wb) == before


def test_placeholders_are_all_placed_before_any_is_replaced(png: Path):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["B2"].value = "{logo}"
    ws["C1048576"].value = "{logo}"
    before = read_pkg(wb)
    with pytest.raises(ValueError, match="'C1048576' would pass the edge of the sheet"):
        wb.insert_image_at_placeholder("{logo}", png, width=50, height=50)
    assert read_pkg(wb) == before
    ws["C1048576"].value = None
    assert wb.insert_image_at_placeholder("{logo}", png, width=50, height=50) == 1
    assert ws["B2"].value is None
    assert [(pic.name, pic.anchor) for pic in ws.images] == [("logo", "B2")]


def _box(ws, anchor: etree._Element) -> tuple[int, int, int, int]:
    """Left, top, right and bottom of a two-cell anchor, in EMU from the sheet's corner."""
    cols, rows = sheet_axis(ws, rows=False), sheet_axis(ws, rows=True)
    (c1, co1, r1, ro1), (c2, co2, r2, ro2) = _marker(anchor, "from"), _marker(anchor, "to")
    return cols.start(c1) + co1, rows.start(r1) + ro1, cols.start(c2) + co2, rows.start(r2) + ro2


def _size(box: tuple[int, int, int, int]) -> tuple[int, int]:
    return box[2] - box[0], box[3] - box[1]


def test_placement_reads_every_anchor_kind(png: Path):
    wb = Workbook.open(INSPECT_FIXTURES["DrawingAnchors"])
    assert wb["Anchors"].images[0].placement == "oneCell"
    assert wb["Anchors"].charts[0].placement == "twoCell"
    ws = _walker_workbook(png)["Sheet1"]
    assert {pic.name: pic.placement for pic in ws.images} == {
        "two": "twoCell",
        "one": "oneCell",
        "grouped": "twoCell",  # its group's anchor
    }
    assert {chart.name: chart.placement for chart in ws.charts} == {
        "top chart": "twoCell",
        "abs chart": "absolute",
    }


def test_placement_sets_edit_as_on_a_two_cell_anchor():
    wb = Workbook.open(INSPECT_FIXTURES["DrawingAnchors"])
    before = read_pkg(wb)
    chart = wb["Anchors"].charts[0]
    chart.placement = "absolute"
    assert chart._anchor.get("editAs") == "absolute"
    chart.placement = "oneCell"
    assert wb["Anchors"].charts[0].placement == "oneCell"
    assert_preserved(before, read_pkg(wb), expected_changed={"xl/drawings/drawing1.xml"})
    chart.placement = "twoCell"
    assert "editAs" not in chart._anchor.attrib
    assert_valid_package(wb)


def test_placement_refuses_what_an_anchor_cannot_store(png: Path):
    ws = _walker_workbook(png)["Sheet1"]
    one = next(pic for pic in ws.images if pic.name == "one")
    grouped = next(pic for pic in ws.images if pic.name == "grouped")
    with pytest.raises(ValueError, match="a oneCellAnchor is always 'oneCell'"):
        one.placement = "twoCell"
    one.placement = "oneCell"  # what it is already
    with pytest.raises(ValueError, match="part of a group"):
        grouped.placement = "oneCell"
    with pytest.raises(ValueError, match="one of 'twoCell', 'oneCell', 'absolute'"):
        ws.images[0].placement = "moveAndSize"
    with pytest.raises(TypeError):
        ws.images[0].placement = None


@pytest.mark.parametrize("placement", ["oneCell", "absolute"])
def test_moving_an_object_that_does_not_size_with_cells_keeps_its_size(placement: str):
    wb = Workbook.open(INSPECT_FIXTURES["DrawingAnchors"])
    ws = wb["Anchors"]
    chart = ws.charts[0]
    chart.placement = placement
    size = _size(_box(ws, chart._anchor))
    chart.anchor = "D2"  # over the 30 pt row 2 and the 6 pt row 4, and a wider column
    assert _marker(chart._anchor, "from")[::2] == (3, 1)
    assert _size(_box(ws, chart._anchor)) == size
    chart.offset_x = 20
    chart.offset_y = 5
    assert _size(_box(ws, chart._anchor)) == size
    assert_valid_package(wb)


def test_moving_an_object_that_sizes_with_cells_keeps_its_cells():
    wb = Workbook.open(INSPECT_FIXTURES["DrawingAnchors"])
    chart = wb["Anchors"].charts[0]
    assert _marker(chart._anchor, "from") == (1, 0, 15, 0)
    assert _marker(chart._anchor, "to") == (4, 0, 23, 0)
    chart.anchor = "D2"
    chart.offset_y = 5
    assert _marker(chart._anchor, "from") == (3, 0, 1, 5 * EMU_PER_PIXEL)
    assert _marker(chart._anchor, "to") == (6, 0, 9, 0)


def test_axis_measures_runs_and_defaults():
    axis = _Axis([(2, 3, 50), (6, 6, 0)], 10, 10)
    assert [axis.start(i) for i in range(8)] == [0, 10, 20, 70, 120, 130, 140, 140]
    assert axis.locate(25) == (2, 5)
    assert axis.locate(70) == (3, 0)  # a gridline starts the line below it
    assert axis.locate(140) == (7, 0)  # the hidden line 6 takes no room
    assert axis.locate(170) is None  # past line 9


def test_axis_anchors_past_its_last_sized_line_in_the_zero_size_lines_that_end_it():
    axis = _Axis([(4, 9, 0)], 10, 10)  # like a sheet that hides every column after D
    assert axis.locate(40) == (4, 0)
    assert axis.locate(55) == (4, 15)
    assert axis.advance(2, 35) == (4, 15)
    assert axis.advance(6, 5) == (6, 5)  # never above the line it starts from
    assert _Axis([], 10, 10).locate(100) is None  # no lines left


def _sheet_with_rows(rows: list[tuple[int, str | None, bool]], *, zero_height: bool = False):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    root = ws._part.element
    fmt = root.find(f"{{{SML_NS}}}sheetFormatPr")
    if zero_height:
        fmt.set("zeroHeight", "1")
    sheet_data = root.find(f"{{{SML_NS}}}sheetData")
    for number, height, hidden in rows:
        row = etree.SubElement(sheet_data, f"{{{SML_NS}}}row", r=str(number))
        if height is not None:
            row.set("ht", height)
        if hidden:
            row.set("hidden", "1")
    ws._invalidate_bulk_indexes()
    return wb, ws


@pytest.mark.parametrize("seed", range(40))
def test_row_walker_measures_as_the_whole_axis_does(seed: int):
    rng = random.Random(seed)
    numbers = sorted(rng.sample(range(1, 60), rng.randint(0, 25)))
    rows = [(n, rng.choice([None, "30", "6", "0", "x"]), rng.random() < 0.2) for n in numbers]
    _, ws = _sheet_with_rows(rows, zero_height=rng.random() < 0.3)
    axis, walker = sheet_axis(ws, rows=True), _RowWalker(ws)
    for _ in range(25):
        a, b = rng.randrange(70), rng.randrange(70)
        a_off, b_off, distance = rng.randrange(400000), rng.randrange(400000), rng.randrange(6000000)
        assert walker.advance(a, distance) == axis.advance(a, distance)
        assert walker.between(a, a_off, b, b_off) == axis.between(a, a_off, b, b_off)


def test_an_unreadable_height_counts_as_the_default():
    _, ws = _sheet_with_rows([(2, "", False), (3, "nan", False), (4, "-5", False)])
    assert sheet_axis(ws, rows=True).start(5) == 5 * sheet_axis(ws, rows=True).default


def test_a_picture_reaches_into_columns_hidden_past_the_used_ones(png: Path):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    root = ws._part.element
    cols = etree.Element(f"{{{SML_NS}}}cols")
    etree.SubElement(cols, f"{{{SML_NS}}}col", min="5", max="16384", hidden="1", width="8")
    root.find(f"{{{SML_NS}}}sheetData").addprevious(cols)
    pic = ws.add_image(png, anchor="D2", width=100, height=40)
    assert _marker(pic._anchor, "to")[:2] == (4, 36 * EMU_PER_PIXEL)  # D is 64 px wide
    pic.anchor = "D5"
    assert pic.width == 100
    pic.width = 64  # exactly D: the right edge lies on the last visible gridline
    assert _marker(pic._anchor, "to")[:2] == (4, 0)
    assert_valid_package(wb)


def test_rows_zero_height_hides_count_as_zero(png: Path):
    _, ws = _sheet_with_rows([(n, None, False) for n in range(1, 11)], zero_height=True)
    assert sheet_axis(ws, rows=True).start(20) == 10 * 20 * EMU_PER_PIXEL
    pic = ws.add_image(png, anchor="B9", width=20, height=80)
    assert _marker(pic._anchor, "to")[2:] == (10, 40 * EMU_PER_PIXEL)


def test_axis_keeps_the_first_of_overlapping_runs():
    axis = _Axis([(1, 4, 20), (3, 5, 30)], 10, 8)
    assert axis.runs == [(1, 4, 20), (5, 5, 30)]
    assert axis.start(6) == 10 + 4 * 20 + 30


def test_sheet_axis_reads_heights_widths_and_hidden_lines():
    wb = Workbook.open(INSPECT_FIXTURES["DrawingAnchors"])
    ws = wb["Anchors"]
    rows, cols = sheet_axis(ws, rows=True), sheet_axis(ws, rows=False)
    assert rows.default == 182880  # 14.4 pt
    assert rows.start(2) == 182880 + 30 * 12700
    assert cols.default == 64 * EMU_PER_PIXEL
    assert cols.start(2) == (64 + 144) * EMU_PER_PIXEL  # B is 20.578125 wide: 144 px
    ws.row_dimensions[3].height = 20
    ws._find_row_element(3).set("hidden", "1")
    assert sheet_axis(ws, rows=True).start(3) == 182880 + 30 * 12700


@pytest.mark.parametrize("value", [0, -5])
def test_picture_size_must_be_positive(png: Path, value: int):
    wb = Workbook.create()
    pic = wb["Sheet1"].add_image(png, anchor="B2")
    with pytest.raises(ValueError):
        pic.width = value


# --- series caches ----------------------------------------------------------


def _series_caches(wb: Workbook) -> etree._Element:
    return etree.fromstring(read_pkg(wb)["xl/charts/chart1.xml"])


def test_num_cache_holds_only_numbers_with_sparse_points():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "Item"
    ws["B1"].value = "Qty"
    labels = ["a", "b", None, "d", 5, "f"]
    values = [3, True, None, "text", date(2024, 1, 2), 2.5]
    for row, (label, value) in enumerate(zip(labels, values), start=2):
        if label is not None:
            ws[f"A{row}"].value = label
        if value is not None:
            ws[f"B{row}"].value = value
    ws["B6"].apply_date_format()
    ws["B7"].apply_number_format("0.00")
    cells_before = len(ws.cells)

    ws.add_chart("bar", anchor="D2", data_range="A1:B7")

    assert len(ws.cells) == cells_before  # reading the range creates no cells
    root = _series_caches(wb)
    num_cache = root.find(f".//{_C}numRef/{_C}numCache")
    assert num_cache.findtext(f"{_C}formatCode") == "General"
    assert num_cache.find(f"{_C}ptCount").get("val") == "6"
    points = {
        int(pt.get("idx")): (pt.findtext(f"{_C}v"), pt.get("formatCode"))
        for pt in num_cache.iter(f"{_C}pt")
    }
    assert points == {0: ("3", None), 4: ("45293", "mm-dd-yy"), 5: ("2.5", "0.00")}
    for text, _ in points.values():
        float(text)

    cat_cache = root.find(f".//{_C}cat/{_C}strRef/{_C}strCache")
    assert cat_cache.find(f"{_C}ptCount").get("val") == "6"
    assert {int(pt.get("idx")): pt.findtext(f"{_C}v") for pt in cat_cache.iter(f"{_C}pt")} == {
        0: "a",
        1: "b",
        3: "d",
        4: "5",
        5: "f",
    }
    with warnings.catch_warnings():
        warnings.simplefilter("error", SchemaWarning)
        assert_valid_package(wb)


def test_num_cache_format_code_comes_from_the_source_cells():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "Item"
    ws["B1"].value = "Share"
    for row, share in ((2, 0.25), (3, 0.5)):
        ws[f"A{row}"].value = f"item {row}"
        ws[f"B{row}"].value = share
        ws[f"B{row}"].apply_number_format("0.0%")

    ws.add_chart("bar", anchor="D2", data_range="A1:B3")

    num_cache = _series_caches(wb).find(f".//{_C}numCache")
    assert num_cache.findtext(f"{_C}formatCode") == "0.0%"
    assert all(pt.get("formatCode") is None for pt in num_cache.iter(f"{_C}pt"))


def test_blank_series_name_gives_an_empty_string_cache():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A2"].value = "x"
    ws["B2"].value = 1
    ws.add_chart("bar", anchor="D2", data_range="A1:B2")
    name_cache = _series_caches(wb).find(f".//{_C}tx/{_C}strRef/{_C}strCache")
    assert name_cache.find(f"{_C}ptCount").get("val") == "1"
    assert name_cache.find(f"{_C}pt") is None


# --- unique ids -------------------------------------------------------------


def _cnvpr_ids(root: etree._Element) -> list[int]:
    return [int(e.get("id")) for e in root.iter(f"{_XDR}cNvPr")]


def test_added_objects_get_unique_ids_and_fresh_creation_ids(png: Path):
    wb = Workbook.open(INSPECT_FIXTURES["images"])
    ws = wb["Sheet1"]
    ws["J1"].value = "k"
    ws["K1"].value = "v"
    ws["J2"].value = "x"
    ws["K2"].value = 1
    ws.add_image(png, anchor="H2")
    ws.add_image(png, anchor="H20")
    ws.add_chart("bar", anchor="M2", data_range="J1:K2")
    ws.add_chart("bar", anchor="M30", data_range="J1:K2")

    root = _drawing_xml(_reopen(wb))
    assert _cnvpr_ids(root) == [3, 4, 5, 6, 7]
    creation_ids = [e.get("id") for e in root.iter(_A16_CREATION_ID)]
    assert len(creation_ids) == 5
    assert len(set(creation_ids)) == 5
    assert all(_GUID.fullmatch(value) for value in creation_ids)
    assert [obj.id for obj in _reopen(wb)["Sheet1"].drawing_objects] == [3, 4, 5, 6, 7]


def test_first_object_on_a_new_drawing_gets_id_2(png: Path):
    wb = Workbook.create()
    wb["Sheet1"].add_image(png, anchor="B2")
    assert _cnvpr_ids(_drawing_xml(wb)) == [2]


# --- anchor walker ----------------------------------------------------------

_NS = (
    f'xmlns:xdr="{XDR_NS}" xmlns:a="{A_NS}" xmlns:c="{CHART_NS}"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
)


def _from(col: int, row: int, tag: str = "from") -> str:
    return (
        f"<xdr:{tag}><xdr:col>{col}</xdr:col><xdr:colOff>0</xdr:colOff>"
        f"<xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:{tag}>"
    )


def _two(obj: str, col: int = 1, row: int = 1) -> str:
    return (
        f"<xdr:twoCellAnchor>{_from(col, row)}{_from(col + 2, row + 4, 'to')}{obj}"
        "<xdr:clientData/></xdr:twoCellAnchor>"
    )


def _pic(id_: int, name: str, rid: str) -> str:
    return (
        f'<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{id_}" name="{name}"/><xdr:cNvPicPr/></xdr:nvPicPr>'
        f'<xdr:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
        '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>'
    )


def _sp(id_: int, name: str, tag: str = "sp") -> str:
    nv = {"sp": "nvSpPr", "cxnSp": "nvCxnSpPr"}[tag]
    c_nv = {"sp": "cNvSpPr", "cxnSp": "cNvCxnSpPr"}[tag]
    return (
        f'<xdr:{tag} macro=""><xdr:{nv}><xdr:cNvPr id="{id_}" name="{name}"/><xdr:{c_nv}/></xdr:{nv}>'
        f"<xdr:spPr/></xdr:{tag}>"
    )


def _frame(id_: int, name: str, uri: str, data: str) -> str:
    return (
        f'<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{id_}" name="{name}"/>'
        "<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>"
        '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
        f'<a:graphic><a:graphicData uri="{uri}">{data}</a:graphicData></a:graphic></xdr:graphicFrame>'
    )


def _alternate(requires: str, ns: str, choice: str, fallback: str) -> str:
    return (
        f'<mc:AlternateContent><mc:Choice xmlns:{requires}="{ns}" Requires="{requires}">{choice}</mc:Choice>'
        f"<mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent>"
    )


def _walker_workbook(png: Path) -> Workbook:
    """Sheet1's drawing holds one of every anchor and object kind the walker classifies."""
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "k"
    ws["B1"].value = "v"
    ws["A2"].value = "x"
    ws["B2"].value = 1
    ws.add_image(png, anchor="A1")
    ws.add_chart("bar", anchor="D2", data_range="A1:B2")
    part = drawing_parts_for_worksheet(ws)[0]
    image_rid = next(rel.rId for rel in part.rels if rel.reltype == RT.IMAGE)
    chart_rid = next(rel.rId for rel in part.rels if rel.reltype == RT.CHART)

    chart_uri = CHART_NS
    chart = f'<c:chart r:id="{chart_rid}"/>'
    slicer_ns = "http://schemas.microsoft.com/office/drawing/2010/slicer"
    timeline_ns = "http://schemas.microsoft.com/office/drawing/2012/timeslicer"
    chart_ex_ns = "http://schemas.microsoft.com/office/drawing/2014/chartex"
    children = "".join(
        [
            '<xdr:oneCellAnchor>' + _from(0, 0) + '<xdr:ext cx="952500" cy="952500"/>'
            + _pic(2, "one", image_rid) + "<xdr:clientData/></xdr:oneCellAnchor>",
            _two(_pic(3, "two", image_rid)),
            '<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="952500"/>'
            + _frame(4, "abs chart", chart_uri, chart) + "<xdr:clientData/></xdr:absoluteAnchor>",
            _alternate(
                "a14",
                "http://schemas.microsoft.com/office/drawing/2010/main",
                _two(_frame(5, "Region", slicer_ns, f'<sle:slicer xmlns:sle="{slicer_ns}" name="Region"/>')),
                _two(_sp(6, "Region fallback")),
            ),
            _two(
                _alternate(
                    "cx1",
                    "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex",
                    _frame(7, "Waterfall", chart_ex_ns, f'<cx:chart xmlns:cx="{chart_ex_ns}" r:id="{chart_rid}"/>'),
                    _sp(8, "Waterfall fallback"),
                )
            ),
            _alternate(
                "tsle",
                timeline_ns,
                _two(_frame(9, "Date", timeline_ns, f'<tsle:timeslicer xmlns:tsle="{timeline_ns}" name="Date"/>')),
                _two(_sp(10, "Date fallback")),
            ),
            _two(
                '<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="11" name="Group"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>'
                "<xdr:grpSpPr/>"
                + _pic(12, "grouped", image_rid)
                + _sp(13, "box")
                + _sp(14, "line", "cxnSp")
                + "</xdr:grpSp>",
                col=5,
                row=20,
            ),
            _two(_frame(15, "top chart", chart_uri, chart), col=10),
            _two(f'<xdr:contentPart r:id="{image_rid}"/>'),
        ]
    )
    root = part.element
    for child in list(root):
        root.remove(child)
    for child in list(etree.fromstring(f"<xdr:wsDr {_NS}>{children}</xdr:wsDr>")):
        root.append(child)
    part.mark_dirty()
    return _reopen(wb)


def test_drawing_objects_walk_every_anchor_group_and_alternate(png: Path):
    ws = _walker_workbook(png)["Sheet1"]
    found = [
        (obj.kind, obj.name, obj.anchor_type, obj.alternate, obj.group.name if obj.group else None)
        for obj in ws.drawing_objects
    ]
    assert found == [
        ("picture", "one", "oneCellAnchor", None, None),
        ("picture", "two", "twoCellAnchor", None, None),
        ("chart", "abs chart", "absoluteAnchor", None, None),
        ("slicer", "Region", "twoCellAnchor", "choice", None),
        ("shape", "Region fallback", "twoCellAnchor", "fallback", None),
        ("chartEx", "Waterfall", "twoCellAnchor", "choice", None),
        ("shape", "Waterfall fallback", "twoCellAnchor", "fallback", None),
        ("timeline", "Date", "twoCellAnchor", "choice", None),
        ("shape", "Date fallback", "twoCellAnchor", "fallback", None),
        ("group", "Group", "twoCellAnchor", None, None),
        ("picture", "grouped", "twoCellAnchor", None, "Group"),
        ("shape", "box", "twoCellAnchor", None, "Group"),
        ("connector", "line", "twoCellAnchor", None, "Group"),
        ("chart", "top chart", "twoCellAnchor", None, None),
        ("contentPart", None, "twoCellAnchor", None, None),
    ]
    assert [obj.id for obj in ws.drawing_objects][:3] == [2, 3, 4]
    assert ws.drawing_objects[2].anchor == ""
    assert ws.drawing_objects[10].anchor == "F21"


def test_images_and_charts_keep_their_order_and_add_what_was_hidden(png: Path):
    ws = _walker_workbook(png)["Sheet1"]
    # Direct two-cell pictures, then one-cell ones, as before; then the rest.
    assert [pic.name for pic in ws.images] == ["two", "one", "grouped"]
    assert [chart.name for chart in ws.charts] == ["top chart", "abs chart"]
    assert ws.charts[1].partname == ws.charts[0].partname


def test_replace_image_reaches_a_grouped_picture(png: Path, tmp_path: Path):
    wb = _walker_workbook(png)
    replacement = tmp_path / "other.png"
    replacement.write_bytes(PNG)
    assert wb.replace_image("grouped", replacement) == 1


def test_grouped_picture_geometry_raises(png: Path):
    pic = _walker_workbook(png)["Sheet1"].images[2]
    assert pic.name == "grouped"
    assert pic.anchor == "F21"  # the group's anchor
    for action in (
        lambda: setattr(pic, "anchor", "A1"),
        lambda: setattr(pic, "offset_x", 3),
        lambda: setattr(pic, "width", 30),
    ):
        with pytest.raises(ValueError, match="part of a group"):
            action()


def test_absolute_anchor_chart_cannot_take_a_cell(png: Path):
    chart = _walker_workbook(png)["Sheet1"].charts[1]
    assert chart.name == "abs chart"
    assert chart.anchor == "" and chart.to_anchor == ""
    with pytest.raises(ValueError, match="absolute anchor"):
        chart.anchor = "B2"
    with pytest.raises(ValueError, match="no xdr:to"):
        chart.to_anchor = "B2"


def test_absolute_anchor_picture_resizes_through_its_ext(png: Path):
    wb = _walker_workbook(png)
    ws = wb["Sheet1"]
    part = drawing_parts_for_worksheet(ws)[0]
    anchor = part.element.find(f"{_XDR}absoluteAnchor")
    frame = anchor.find(f"{_XDR}graphicFrame")
    anchor.replace(frame, etree.fromstring(f"<root {_NS}>{_pic(16, 'abs pic', 'rId1')}</root>")[0])
    part.mark_dirty()
    pic = next(p for p in ws.images if p.name == "abs pic")

    pic.width = 30

    assert anchor.find(f"{_XDR}ext").get("cx") == str(30 * EMU_PER_PIXEL)
    assert pic.width == 30
    with pytest.raises(ValueError, match="absolute anchor"):
        pic.anchor = "C3"
