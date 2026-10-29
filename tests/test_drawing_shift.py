"""Drawing anchors follow inserts and deletes by their placement (issue #21).

The Excel-reference pairs in ``fixtures/excel-reference/`` record what Excel
does; ``tests/test_excel_reference.py`` compares inserts with them, and the
tests here also compare the plan-level delete that ``delete_rows`` (#22) will
call, and check the rules on ``DrawingAnchors.xlsx`` along both axes.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from lxml import etree

from tests import excel_facts as facts
from tests.conftest import INSPECT_FIXTURES
from tests.preservation import assert_preserved, read_pkg
from tests.schema_validate import assert_valid_package
from tests.test_drawing import _walker_workbook, png  # noqa: F401 (fixture)
from tests.test_excel_reference import REFERENCE, assert_matches_excel
from xlsxedit import Workbook
from xlsxedit.drawing import XDR_NS, _Axis, drawing_parts_for_worksheet, sheet_axis
from xlsxedit.drawing_shift import DrawingShift, anchor_extent
from xlsxedit.exceptions import GridOverflowError, InvalidRangeError
from xlsxedit.row_shift import _Shift

_XDR = f"{{{XDR_NS}}}"


def _pair(name: str) -> tuple[Path, Path]:
    return REFERENCE / name / "before.xlsx", REFERENCE / name / "after.xlsx"


def _anchors(wb: Workbook, sheet: str) -> dict[str, facts.Anchor]:
    return facts.drawings_by_name(facts.Book(read_pkg(wb))).get(sheet, {})


def _span(wb: Workbook, sheet: str, anchor: facts.Anchor, *, rows: bool) -> tuple[int, int]:
    """Where an anchor's two corners lie on one axis, in EMU from the sheet's edge."""
    axis = sheet_axis(wb[sheet], rows=rows)
    i, j = (2, 3) if rows else (0, 1)
    return tuple(axis.start(corner[i]) + corner[j] for corner in (anchor.start, anchor.end))


# --- Excel's recordings -----------------------------------------------------


def test_anchors_match_excel_over_the_row_heights_excel_wrote():
    """The insert-through-drawings pair differs from Excel only because Excel gives
    the inserted rows the height of the row above them (#19): measured over
    Excel's own rows, every anchor lands where Excel put it."""
    before, after = _pair("insert-rows-through-drawings")

    class OverExcelRows(DrawingShift):
        _after = sheet_axis(Workbook.open(after)["Drawings"], rows=True)

    wb = Workbook.open(before)
    OverExcelRows(wb["Drawings"], _Shift(True, 7, 3)).apply()
    assert_matches_excel(wb, after, ["drawings"])


@pytest.mark.parametrize(
    ("name", "sheet"), [("delete-rows-under-cf", "Sheet1"), ("delete-rows-at-drawing-edges", "Edges")]
)
def test_a_planned_delete_matches_excel(name: str, sheet: str):
    before, after = _pair(name)
    wb = Workbook.open(before)
    DrawingShift(wb[sheet], _Shift(True, 5, -3)).apply()
    assert_matches_excel(wb, after, ["drawings"])
    assert_valid_package(wb)


# --- DrawingAnchors.xlsx: every placement over uneven rows and columns -------


@pytest.mark.parametrize("rows", [True, False], ids=["rows", "columns"])
def test_each_placement_keeps_its_rule(rows: bool):
    wb = Workbook.open(INSPECT_FIXTURES["DrawingAnchors"])
    ws = wb["Anchors"]
    before_pkg = read_pkg(wb)
    old = _anchors(wb, "Anchors")
    old_spans = {name: _span(wb, "Anchors", a, rows=rows) for name, a in old.items()}
    if rows:
        ws.insert_rows([[]] * 2, at_row=3)
    else:
        ws.insert_columns([[]] * 2, at_col="D")
    new = _anchors(wb, "Anchors")
    assert new.keys() == old.keys()
    i = 2 if rows else 0
    for name, anchor in new.items():
        top, bottom = _span(wb, "Anchors", anchor, rows=rows)
        old_top, old_bottom = old_spans[name]
        if anchor.edit_as == "absolute":
            assert (top, bottom) == (old_top, old_bottom), name
        elif anchor.edit_as == "oneCell":
            assert bottom - top == old_bottom - old_top, name
        else:  # two-cell: each corner moves with its cell
            for corner, was in ((anchor.start, old[name].start), (anchor.end, old[name].end)):
                assert corner[i] in (was[i], was[i] + 2), name
        cross = 0 if rows else 2
        was = old[name]
        assert (anchor.start[cross], anchor.end[cross]) == (was.start[cross], was.end[cross])

    if rows:  # rows 3 and 4 go in through the shapes on rows 2 to 4, above the rest
        assert new["Move and size"].end[2:] == (5, 50800)
        assert new["Picture"].start[2:] == (8, 0)
        changed = {"xl/drawings/drawing1.xml", "xl/charts/chart1.xml"}  # the chart plots A1:A6
    else:  # columns D and E go in at the column of 'Move'
        assert new["Move and size"] == old["Move and size"]
        assert new["Move"].start[:2] == (5, 38100)
        assert new["Group"].end[:2] == (5, 2320290)
        changed = {"xl/drawings/drawing1.xml"}
    assert_preserved(
        before_pkg,
        read_pkg(wb),
        expected_changed={"xl/worksheets/sheet1.xml", "xl/workbook.xml", *changed},
    )
    assert_valid_package(wb)


def test_shift_drawings_false_leaves_every_anchor():
    wb = Workbook.open(INSPECT_FIXTURES["DrawingAnchors"])
    before = read_pkg(wb)
    wb["Anchors"].insert_rows([[]] * 2, at_row=3, shift_drawings=False)
    wb["Anchors"].insert_columns([[]], at_col="B", shift_drawings=False)
    assert_preserved(
        before,
        read_pkg(wb),
        expected_changed={"xl/worksheets/sheet1.xml", "xl/workbook.xml", "xl/charts/chart1.xml"},
    )


def test_a_chart_grows_with_the_rows_it_spans():
    wb = Workbook.open(INSPECT_FIXTURES["ChartsAndTables"])
    wb["bar chart"].insert_rows([[None], [None]], at_row=4)
    chart = _anchors(wb, "bar chart")["Chart 5"]
    assert (chart.start, chart.end) == ((2, 826880, 1, 19188), (8, 429315, 16, 160268))
    assert _anchors(wb, "line chart") == _anchors(
        Workbook.open(INSPECT_FIXTURES["ChartsAndTables"]), "line chart"
    )


# --- anchor kinds, alternate content, groups --------------------------------


def _froms(wb: Workbook) -> list[tuple[str, tuple[int, ...]]]:
    root = drawing_parts_for_worksheet(wb["Sheet1"])[0].element
    found = []
    for anchor in root.iter(f"{_XDR}twoCellAnchor", f"{_XDR}oneCellAnchor", f"{_XDR}absoluteAnchor"):
        start = anchor.find(f"{_XDR}from")
        pos, ext = anchor.find(f"{_XDR}pos"), anchor.find(f"{_XDR}ext")
        found.append(
            (
                etree.QName(anchor).localname,
                () if start is None else (int(start.findtext(f"{_XDR}col")), int(start.findtext(f"{_XDR}row"))),
                tuple(int(v) for v in pos.attrib.values()) if pos is not None else (),
                tuple(int(v) for v in ext.attrib.values()) if ext is not None else (),
            )
        )
    return found


def test_every_anchor_kind_and_branch_moves(png: Path):  # noqa: F811
    wb = _walker_workbook(png)
    old = _froms(wb)
    wb["Sheet1"].insert_rows([[]] * 2, at_row=1)
    new = _froms(wb)
    for (kind, start, pos, ext), (kind2, start2, pos2, ext2) in zip(old, new):
        assert (kind2, pos2, ext2) == (kind, pos, ext)
        assert start2 == ((start[0], start[1] + 2) if start else ())
    assert_valid_package(wb)


def test_a_planned_delete_removes_alternate_content_whole(png: Path):  # noqa: F811
    wb = _walker_workbook(png)
    ws = wb["Sheet1"]
    DrawingShift(ws, _Shift(True, 2, -4)).apply()
    names = [obj.name for obj in ws.drawing_objects]
    assert names == ["one", "abs chart", "Group", "grouped", "box", "line"]
    assert [obj.anchor for obj in ws.drawing_objects][2] == "F17"
    assert_valid_package(wb)


def test_a_planned_delete_releases_what_only_the_deleted_anchor_used():
    wb = Workbook.open(INSPECT_FIXTURES["DrawingAnchors"])
    ws = wb["Anchors"]
    DrawingShift(ws, _Shift(True, 16, -9)).apply()
    assert "Chart" not in _anchors(wb, "Anchors")
    part = drawing_parts_for_worksheet(ws)[0]
    assert [rel.reltype.rsplit("/", 1)[-1] for rel in part.rels] == ["image"]  # 'Picture' uses it
    assert [str(name) for name in wb._package.prune_unreachable()] == ["/xl/charts/chart1.xml"]
    assert_valid_package(wb)


# --- refusals ---------------------------------------------------------------


def test_an_anchor_pushed_off_the_grid_refuses_the_insert():
    wb = Workbook.open(INSPECT_FIXTURES["DrawingAnchors"])
    ws = wb["Anchors"]
    ws.images[0].anchor = "B1048573"  # three rows and a third tall: it ends in the last row
    before = read_pkg(wb)
    with pytest.raises(GridOverflowError, match="drawing object 'Picture'"):
        ws.insert_rows([[]], at_row=1)
    assert read_pkg(wb) == before


def test_a_one_cell_anchor_whose_extent_would_pass_the_last_row_refuses_the_insert(png):  # noqa: F811
    wb = _walker_workbook(png)
    ws = wb["Sheet1"]
    one = drawing_parts_for_worksheet(ws)[0].element.find(f"{_XDR}oneCellAnchor")
    one.find(f"{_XDR}from/{_XDR}row").text = "1048570"  # five 20 px rows: one row to spare
    before = read_pkg(wb)
    with pytest.raises(GridOverflowError, match="drawing object 'one'"):
        ws.insert_rows([[]] * 2, at_row=1048000)
    assert read_pkg(wb) == before
    ws.insert_rows([[]], at_row=1048000)  # now it ends in the last row
    assert one.find(f"{_XDR}from/{_XDR}row").text == "1048571"


def test_a_corner_without_its_row_refuses_the_insert():
    wb = Workbook.open(INSPECT_FIXTURES["DrawingAnchors"])
    root = drawing_parts_for_worksheet(wb["Anchors"])[0].element
    marker = root.find(f".//{_XDR}from")
    marker.remove(marker.find(f"{_XDR}row"))
    before = read_pkg(wb)
    with pytest.raises(InvalidRangeError, match="xdr:from does not name a cell"):
        wb["Anchors"].insert_rows([[]], at_row=1)
    assert read_pkg(wb) == before


def test_an_anchor_that_names_no_cell_refuses_the_insert():
    wb = Workbook.open(INSPECT_FIXTURES["DrawingAnchors"])
    root = drawing_parts_for_worksheet(wb["Anchors"])[0].element
    root.find(f".//{_XDR}to/{_XDR}row").text = "x"
    before = read_pkg(wb)
    with pytest.raises(InvalidRangeError, match="'Move and size': xdr:to does not name a cell"):
        wb["Anchors"].insert_rows([[]], at_row=1)
    assert read_pkg(wb) == before


# --- spans ------------------------------------------------------------------


def test_a_span_shift_moves_only_the_anchors_it_carries():
    """Excel's *Insert Cells* moves an object only when every column it covers
    lies in the span and it starts at or below the first new row (recorded
    2026-10-29, see ``_Shift.carries``): one the new rows cut through stays."""
    wb = Workbook.open(INSPECT_FIXTURES["DrawingAnchors"])
    ws = wb["Anchors"]
    old = _anchors(wb, "Anchors")
    DrawingShift(ws, _Shift(True, 3, 2, span=(2, 4))).apply()  # cells B3:D down two rows
    new = _anchors(wb, "Anchors")
    assert new["Move and size"] == old["Move and size"]  # rows 2-4: the new rows cut it
    assert new["Move"] == old["Move"]  # rows 2-4 as well
    assert new["Absolute"] == old["Absolute"]  # column E
    assert new["Connector"] == old["Connector"]  # columns E:G
    for name in ("Text box", "Group", "Chart"):  # wholly in B:D, below row 3
        assert new[name].start[2] == old[name].start[2] + 2, name
        assert new[name].end[2] == old[name].end[2] + 2, name
    assert new["Picture"].start[2] == old["Picture"].start[2] + 2
    top, bottom = _span(wb, "Anchors", new["Picture"], rows=True)
    old_top, old_bottom = _span(wb, "Anchors", old["Picture"], rows=True)
    assert bottom - top == old_bottom - old_top  # the rows did not move, the picture did


def test_a_one_cell_anchor_covers_the_cells_its_size_reaches():
    wb = Workbook.open(INSPECT_FIXTURES["DrawingAnchors"])
    ws = wb["Anchors"]
    root = drawing_parts_for_worksheet(ws)[0].element
    anchor = next(a for a in root if etree.QName(a).localname == "twoCellAnchor")
    one_cell = etree.fromstring(
        f'<xdr:oneCellAnchor xmlns:xdr="{XDR_NS}"><xdr:from><xdr:col>1</xdr:col>'
        "<xdr:colOff>0</xdr:colOff><xdr:row>20</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
        f'<xdr:ext cx="{3 * 609600}" cy="{2 * 190500}"/></xdr:oneCellAnchor>'
    )
    axes = {rows: sheet_axis(ws, rows=rows) for rows in (True, False)}
    markers, cells = anchor_extent(one_cell, axes.__getitem__)
    assert cells.min_col == 2 and cells.min_row == 21
    assert cells.max_col > cells.min_col and cells.max_row > cells.min_row
    assert anchor_extent(one_cell)[1] == anchor_extent(one_cell)[0]  # no axes: its from cell
    assert anchor_extent(anchor) is not None



def _one_cell_anchor(col: int, row: int, cx: int, cy: int) -> etree._Element:
    return etree.fromstring(
        f'<xdr:oneCellAnchor xmlns:xdr="{XDR_NS}"><xdr:from><xdr:col>{col}</xdr:col>'
        f"<xdr:colOff>0</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
        f'<xdr:ext cx="{cx}" cy="{cy}"/></xdr:oneCellAnchor>'
    )


def test_a_one_cell_anchor_that_ends_on_a_gridline_covers_only_its_own_lines():
    """Its far cell holds its last EMU: not the line past the gridline, nor a hidden
    column there, and an object filling the last row stays on the grid."""
    width, height = 609600, 190500
    rows = _Axis([], height, 1_048_576)
    cols_c_hidden = _Axis([(2, 2, 0)], width, 16_384)
    axes = {True: rows, False: cols_c_hidden}.__getitem__
    markers, cells = anchor_extent(_one_cell_anchor(1, 4, width, height), axes)
    assert str(markers) == str(cells) == "B5"
    markers, cells = anchor_extent(_one_cell_anchor(1, 1_048_574, width, height), axes)
    assert str(cells) == "B1048575" and markers.max_row == 1_048_575

@pytest.mark.parametrize(("name", "at_row"), [("Absolute", 1), ("Move", 3)])
def test_rows_an_insert_writes_are_visible_on_a_sheet_that_hides_unused_rows(name, at_row):
    """With ``zeroHeight``, a row the sheet does not list is hidden, but the rows an
    insert writes are listed: an absolute object keeps its place on the page and
    a move-only object keeps its size, measured over the rows as saved."""
    wb = Workbook.open(INSPECT_FIXTURES["DrawingAnchors"])
    ws = wb["Anchors"]
    ws._part.element.find(f"{{{ws._part.element.nsmap[None]}}}sheetFormatPr").set("zeroHeight", "1")
    for row in range(1, 35):
        ws[f"A{row}"].value = 1
    before = _span(wb, "Anchors", _anchors(wb, "Anchors")[name], rows=True)
    ws.insert_rows([[1]], at_row=at_row)
    after = _span(wb, "Anchors", _anchors(wb, "Anchors")[name], rows=True)
    if name == "Absolute":
        assert after == before
    else:
        assert after[1] - after[0] == before[1] - before[0]
    assert_valid_package(wb)


def test_bulk_inserts_take_shift_drawings(tmp_path: Path):
    pd = pytest.importorskip("pandas")
    import shutil

    import xlsxedit.pandas_io as xpi

    xpi.register()
    source = INSPECT_FIXTURES["DrawingAnchors"]
    frame = pd.DataFrame({"a": [1, 2], "b": [3, 4]})

    wb = Workbook.open(source)
    before = read_pkg(wb)
    wb.write_dataframe(frame, sheet="Anchors", at_cell="A2", header=False, mode="insert", shift_drawings=False)
    after = read_pkg(wb)
    drawings = [m for m in before if m.startswith("xl/drawings/drawing")]
    assert drawings and all(after[m] == before[m] for m in drawings)

    path = tmp_path / "anchors.xlsx"
    shutil.copy(source, path)
    with pd.ExcelWriter(
        path,
        engine="xlsxedit",
        mode="a",
        if_sheet_exists="overlay",
        engine_kwargs={"mode": "insert", "shift_drawings": False},
    ) as writer:
        frame.to_excel(writer, sheet_name="Anchors", startrow=1, header=False, index=False)
    saved = read_pkg(path)
    assert Workbook.open(path)["Anchors"]["A2"].value == 1
    assert all(saved[m] == before[m] for m in drawings)
