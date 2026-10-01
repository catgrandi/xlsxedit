"""Drawing and chart parts: live trees, anchors, sizes, series caches, ids, the object walker."""

from __future__ import annotations

import io
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
from xlsxedit.drawing import A_NS, CHART_NS, EMU_PER_PIXEL, XDR_NS, drawing_parts_for_worksheet
from xlsxedit.opc.constants import RT
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

    report = assert_preserved(before, wb, expected_changed={worksheet_members(before)["Table"]})
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

    sp_pr = _drawing_xml(_reopen(wb)).find(f"{_XDR}oneCellAnchor/{_XDR}pic/{_XDR}spPr")
    assert etree.QName(sp_pr[0]).localname == "xfrm"
    assert [etree.QName(e).localname for e in sp_pr[0]] == ["ext"]
    assert sp_pr[0][0].get("cy") == str(70 * EMU_PER_PIXEL)


def test_resize_drops_a_bare_ext_left_by_older_releases(png: Path):
    wb = Workbook.create()
    pic = wb["Sheet1"].add_image(png, anchor="B2", width=40, height=40)
    etree.SubElement(pic._pic.find(f"{_XDR}spPr"), f"{_A}ext", cx="1", cy="1")

    pic.width = 60

    sp_pr = _drawing_xml(_reopen(wb)).find(f"{_XDR}oneCellAnchor/{_XDR}pic/{_XDR}spPr")
    assert sp_pr.find(f"{_A}ext") is None
    assert sp_pr.find(f"{_A}xfrm/{_A}ext").get("cx") == str(60 * EMU_PER_PIXEL)
    assert_valid_package(wb)


def test_one_cell_picture_resize_sets_anchor_ext_and_xfrm(png: Path):
    wb = Workbook.create()
    pic = wb["Sheet1"].add_image(png, anchor="B2", width=40, height=40)
    pic.width = 120

    anchor = _drawing_xml(_reopen(wb)).find(f"{_XDR}oneCellAnchor")
    cx = str(120 * EMU_PER_PIXEL)
    assert anchor.find(f"{_XDR}ext").get("cx") == cx
    assert anchor.find(f"{_XDR}pic/{_XDR}spPr/{_A}xfrm/{_A}ext").get("cx") == cx
    assert _reopen(wb)["Sheet1"].images[0].width == 120


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
