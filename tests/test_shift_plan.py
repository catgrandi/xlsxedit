"""The planner for deletes and span-restricted shifts, which no public method runs yet.

Where a rule was not obvious, the expected result was recorded from Excel for
Microsoft 365 on Windows (2026-10-19): page breaks, frozen panes and the
references in ``test_range_shift.py``; and on 2026-10-29 the drawing objects
"insert cells" moves and Excel's prompt before it unmerges cells.
"""

from __future__ import annotations

import pytest
from lxml import etree

from xlsxedit import Workbook
from xlsxedit.exceptions import InvalidRangeError
from xlsxedit.opc.constants import SML_NS
from xlsxedit.range_set import CellRange
from xlsxedit.row_shift import InsertPlan, _Planner, _Shift, rewrite_formula_refs
from xlsxedit.worksheet_order import insert_worksheet_child
from tests.schema_validate import assert_valid_package

NS = {"m": SML_NS}


def _sheet(*fragments: str):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    root = ws._part.element
    for fragment in fragments:
        views = root.find("m:sheetViews", NS)
        elm = etree.fromstring(f'<w xmlns="{SML_NS}">{fragment}</w>')[0]
        if views is not None and elm.tag == views.tag:
            root.remove(views)
        insert_worksheet_child(root, elm)
    return wb, ws


def _shift(ws, shift: _Shift) -> None:
    planner = _Planner(shift, sheet_name=ws.name, part=ws._part)
    planner.worksheet(ws._part.element)
    planner.apply()


def _attrs(ws, path: str, attr: str) -> list[str | None]:
    return [e.get(attr) for e in ws._part.element.iterfind(path, NS)]


def _rows(at: int, count: int) -> _Shift:
    return _Shift(True, at, -count)


# -- the shift itself ---------------------------------------------------------


def test_a_delete_maps_lines_and_cells_both_ways():
    shift = _Shift(True, 5, -2)
    assert [shift.move(i) for i in (4, 5, 6, 7)] == [4, None, None, 5]
    assert [shift.source(i) for i in (4, 5, 6)] == [4, 7, 8]
    assert shift.source(1_048_575) is None  # a new blank row at the bottom
    assert shift.cell(3, 9) == (3, 7)
    assert shift.cell(3, 5, back=True) == (3, 7)
    assert shift.describe() == "deleting rows 5:6"
    assert _Shift(False, 3, -2).describe() == "deleting columns C:D"


def test_an_insert_maps_lines_and_cells_both_ways():
    shift = _Shift(False, 3, 2)
    assert [shift.move(i) for i in (2, 3)] == [2, 5]
    assert [shift.source(i) for i in (2, 3, 4, 5)] == [2, None, None, 3]
    assert shift.cell(4, 9) == (6, 9)
    assert shift.cell(6, 9, back=True) == (4, 9)
    assert shift.cell(16_384, 1) is None
    assert shift.describe() == "inserting 2 column(s) at column C"


def test_a_span_moves_only_the_cells_inside_it():
    shift = _Shift(True, 5, 1, (2, 3))
    assert shift.cell(2, 5) == (2, 6)
    assert shift.cell(4, 5) == (4, 5)
    assert shift.screen is None
    assert shift.splits(CellRange.parse("A6:B6"))
    assert shift.splits(CellRange.parse("C1:D9"))
    assert not shift.splits(CellRange.parse("A1:D4"))  # above the shifted region
    assert not shift.splits(CellRange.parse("B1:C9"))
    assert shift.splits(CellRange.parse("A1:D9"))
    assert _Shift(True, 5, 1, (1, 16_384)).span is None


def test_tears_tells_which_areas_a_shift_moves_only_part_of():
    rows_7 = _Shift(True, 7, 1, (1, 4))
    assert rows_7.tears(CellRange.parse("A6:A8"))  # the inserted row cuts through it
    assert not rows_7.tears(CellRange.parse("A7:A8"))
    assert rows_7.tears(CellRange.parse("A9:F9"))  # across the span's edge
    assert not rows_7.tears(CellRange.parse("A1:F6"))  # above the inserted row
    assert rows_7.tears(CellRange.parse("C:F"))  # over every row, but its cells in C:D move
    assert rows_7.tears(CellRange.parse("A:A"))  # the inserted row cuts through it
    assert _Shift(True, 7, -1, (1, 4)).tears(CellRange.parse("A6:A8"))
    assert not _Shift(True, 7, -2, (1, 4)).tears(CellRange.parse("A7:A8"))  # deleted whole
    assert _Shift(True, 7, -2, (1, 4)).tears(CellRange.parse("A8:A9"))
    assert _Shift(True, 7, 1).tears(CellRange.parse("A6:A8"))


@pytest.mark.parametrize(
    ("cells", "shift", "moves"),
    [
        ("A12:B14", _Shift(True, 7, 1, (1, 4)), True),
        ("C12:F14", _Shift(True, 7, 1, (1, 4)), False),  # across the span's edge
        ("A12:B14", _Shift(True, 7, 1, (2, 4)), False),  # starts left of the span
        ("A5:B9", _Shift(True, 7, 1, (1, 4)), False),  # the inserted row cuts through it
        ("A7:B9", _Shift(True, 7, 1, (1, 4)), True),
        ("A9:B9", _Shift(True, 7, -1, (1, 4)), True),
        ("A7:B9", _Shift(True, 7, -1, (1, 4)), False),
    ],
)
def test_carries_holds_for_the_drawing_objects_excel_moves(cells, shift, moves):
    """Recorded 2026-10-29: shapes over these cells, "insert cells A7:D7" (B7:D7)."""
    assert shift.carries(CellRange.parse(cells)) is moves


@pytest.mark.parametrize(
    ("args", "error"),
    [
        ((True, 5, 0), ValueError),
        ((True, 0, 1), ValueError),
        ((True, 1_048_575, -3), InvalidRangeError),
        ((False, 16_384, -2), InvalidRangeError),
        ((True, 5, 1, (0, 3)), InvalidRangeError),
        ((True, 5, 1, (4, 3)), InvalidRangeError),
        ((False, 5, 1, (1, 1_048_577)), InvalidRangeError),
        ((True, 1_048_578, 1), InvalidRangeError),
        ((False, 16_385, -1), InvalidRangeError),
    ],
)
def test_a_shift_off_the_grid_is_refused(args, error):
    with pytest.raises(error):
        _Shift(*args)


def test_an_insert_just_past_the_grid_moves_nothing():
    wb, ws = _sheet(_FROZEN, '<mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells>')
    before = etree.tostring(ws._part.element)
    for shift in (_Shift(True, 1_048_577, 1), _Shift(False, 16_385, 1)):
        _shift(ws, shift)
    assert etree.tostring(ws._part.element) == before


def test_insert_plan_carries_out_whole_line_inserts_only():
    ws = Workbook.create()["Sheet1"]
    with pytest.raises(ValueError, match="inserts of whole rows or columns only"):
        InsertPlan(ws, _Shift(True, 5, -1))
    with pytest.raises(ValueError, match="inserts of whole rows or columns only"):
        InsertPlan(ws, _Shift(True, 5, 1, (2, 3)))


# -- ranges on the sheet ------------------------------------------------------


def test_a_delete_drops_shrinks_and_moves_merges_formats_and_validations():
    wb, ws = _sheet(
        '<mergeCells count="4"><mergeCell ref="A5:B6"/><mergeCell ref="A4:B8"/>'
        '<mergeCell ref="D7:E7"/><mergeCell ref="G6:G7"/></mergeCells>',
        '<conditionalFormatting sqref="A5:A6 C2:C9"><cfRule type="expression" priority="1">'
        "<formula>TRUE</formula></cfRule></conditionalFormatting>",
        '<conditionalFormatting sqref="H5:H6"><cfRule type="expression" priority="2">'
        "<formula>TRUE</formula></cfRule></conditionalFormatting>",
        '<dataValidations count="2"><dataValidation sqref="B6"/><dataValidation sqref="B7 5:5 E:E"/>'
        "</dataValidations>",
    )
    _shift(ws, _rows(5, 2))
    assert ws.merged_ranges == ["A4:B6", "D5:E5"]  # G6:G7 shrinks to one cell: unmerged
    assert _attrs(ws, "m:mergeCells", "count") == ["2"]
    assert _attrs(ws, "m:conditionalFormatting", "sqref") == ["C2:C7"]
    assert _attrs(ws, "m:dataValidations/m:dataValidation", "sqref") == ["B5 E:E"]
    assert _attrs(ws, "m:dataValidations", "count") == ["1"]
    assert_valid_package(wb)


def test_a_delete_removes_a_hyperlink_and_its_relationship():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["B5"].hyperlink.url = "https://example.com/"
    ws["B9"].hyperlink.url = "https://example.org/"
    _shift(ws, _rows(4, 3))
    assert _attrs(ws, "m:hyperlinks/m:hyperlink", "ref") == ["B6"]
    targets = [rel.target_ref for rel in ws._part.rels]
    assert targets == ["https://example.org/"]


def test_deleting_columns_moves_the_filter_columns_and_drops_deleted_ones():
    wb, ws = _sheet(
        '<autoFilter ref="A1:E10"><filterColumn colId="1"><filters><filter val="x"/></filters>'
        '</filterColumn><filterColumn colId="2"><filters><filter val="y"/></filters></filterColumn>'
        '<filterColumn colId="4"><filters><filter val="z"/></filters></filterColumn></autoFilter>'
    )
    _shift(ws, _Shift(False, 3, -2))
    assert _attrs(ws, "m:autoFilter", "ref") == ["A1:C10"]
    assert _attrs(ws, "m:autoFilter/m:filterColumn", "colId") == ["1", "2"]
    assert_valid_package(wb)


def test_deleting_the_first_filtered_columns_renumbers_the_rest():
    _, ws = _sheet(
        '<autoFilter ref="C1:F10"><filterColumn colId="0"/><filterColumn colId="2"/>'
        '<filterColumn colId="3"/></autoFilter>'
    )
    _shift(ws, _Shift(False, 2, -3))  # columns B:D
    assert _attrs(ws, "m:autoFilter", "ref") == ["B1:C10"]
    assert _attrs(ws, "m:autoFilter/m:filterColumn", "colId") == ["0", "1"]


def test_filter_columns_pushed_off_the_grid_go_even_when_the_range_is_clamped():
    _, ws = _sheet('<autoFilter ref="XFA1:XFD1"><filterColumn colId="1"/><filterColumn colId="3"/></autoFilter>')
    _shift(ws, _Shift(False, 16_382, 1))  # one column at XFB
    assert _attrs(ws, "m:autoFilter", "ref") == ["XFA1:XFD1"]
    assert _attrs(ws, "m:autoFilter/m:filterColumn", "colId") == ["2"]


def test_an_insert_away_from_a_filter_keeps_its_filter_columns():
    wb, ws = _sheet('<autoFilter ref="A1:C10"><filterColumn colId="1"/><filterColumn colId="7"/></autoFilter>')
    ws.insert_columns([[None]], at_col="Z")
    assert _attrs(ws, "m:autoFilter/m:filterColumn", "colId") == ["1", "7"]


def test_a_sort_condition_in_deleted_columns_goes():
    _, ws = _sheet(
        '<autoFilter ref="A1:E10"><sortState ref="A2:E10"><sortCondition ref="B2:B10"/>'
        '<sortCondition ref="E2:E10"/></sortState></autoFilter>'
    )
    _shift(ws, _Shift(False, 2, -1))
    assert _attrs(ws, "m:autoFilter/m:sortState", "ref") == ["A2:D10"]
    assert _attrs(ws, "m:autoFilter/m:sortState/m:sortCondition", "ref") == ["D2:D10"]


def test_deleting_a_scenario_input_cell_is_refused():
    _, ws = _sheet(
        '<scenarios><scenario name="s"><inputCells r="B6" val="1"/></scenario></scenarios>'
    )
    with pytest.raises(InvalidRangeError, match="deleting rows 5:6 would delete scenario inputCells"):
        _shift(ws, _rows(5, 2))


# -- whole-line structures ----------------------------------------------------

_BREAKS = (
    '<rowBreaks count="3" manualBreakCount="3"><brk id="5" max="16383" man="1"/>'
    '<brk id="10" max="16383" man="1"/><brk id="20" max="16383" man="1"/></rowBreaks>'
)


@pytest.mark.parametrize(
    ("at", "count", "ids"),
    [
        (2, 2, ["3", "8", "18"]),
        (1, 6, ["4", "14"]),
        (4, 3, ["7", "17"]),
        (6, 1, ["9", "19"]),
        (5, 1, ["4", "9", "19"]),
        (8, 4, ["5", "16"]),
        (11, 1, ["5", "19"]),
    ],
)
def test_page_breaks_follow_the_row_after_them_as_in_excel(at, count, ids):
    _, ws = _sheet(_BREAKS)
    _shift(ws, _rows(at, count))
    assert _attrs(ws, "m:rowBreaks/m:brk", "id") == ids
    assert _attrs(ws, "m:rowBreaks", "manualBreakCount") == [str(len(ids))]


def test_a_break_brought_to_the_top_row_goes():
    _, ws = _sheet('<rowBreaks count="1" manualBreakCount="1"><brk id="6" man="1"/></rowBreaks>')
    _shift(ws, _rows(1, 6))
    assert ws._part.element.find("m:rowBreaks", NS) is None


_FROZEN = (
    '<sheetViews><sheetView workbookViewId="0"><pane xSplit="2" ySplit="4" topLeftCell="C5" '
    'activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="D7" '
    'sqref="D7:E8"/></sheetView></sheetViews>'
)


@pytest.mark.parametrize(
    ("shift", "split", "top_left"),
    [
        (_Shift(True, 2, -2), "2", "C3"),
        (_Shift(True, 1, -6), "1", "C2"),
        (_Shift(True, 4, -3), "3", "C4"),
        (_Shift(True, 6, -1), "4", "C5"),
        (_Shift(True, 5, -1), "4", "C5"),
        (_Shift(True, 2, 1), "5", "C6"),
        (_Shift(True, 5, 1), "4", "C5"),
    ],
)
def test_a_frozen_split_shrinks_with_deleted_frozen_rows_as_in_excel(shift, split, top_left):
    _, ws = _sheet(_FROZEN)
    _shift(ws, shift)
    assert _attrs(ws, "m:sheetViews/m:sheetView/m:pane", "ySplit") == [split]
    assert _attrs(ws, "m:sheetViews/m:sheetView/m:pane", "topLeftCell") == [top_left]


@pytest.mark.parametrize(
    ("shift", "view_top", "split", "pane_top"),
    [
        (_Shift(True, 1, -1), "A2", "4", "A6"),
        (_Shift(True, 2, -2), "A2", "3", "A5"),
    ],
)
def test_a_frozen_pane_counts_from_where_the_view_top_went(shift, view_top, split, pane_top):
    _, ws = _sheet(
        '<sheetViews><sheetView workbookViewId="0" topLeftCell="A3"><pane ySplit="4" '
        'topLeftCell="A7" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
    )
    _shift(ws, shift)
    assert _attrs(ws, "m:sheetViews/m:sheetView", "topLeftCell") == [view_top]
    assert _attrs(ws, "m:sheetViews/m:sheetView/m:pane", "ySplit") == [split]
    assert _attrs(ws, "m:sheetViews/m:sheetView/m:pane", "topLeftCell") == [pane_top]


def test_selection_areas_that_merge_keep_the_active_area():
    _, ws = _sheet(
        '<sheetViews><sheetView workbookViewId="0"><selection activeCell="B9" activeCellId="2" '
        'sqref="A5 A6 B9"/></sheetView></sheetViews>'
    )
    _shift(ws, _rows(5, 2))
    assert _attrs(ws, "m:sheetViews/m:sheetView/m:selection", "sqref") == ["A5 B7"]
    assert _attrs(ws, "m:sheetViews/m:sheetView/m:selection", "activeCellId") == ["1"]
    assert _attrs(ws, "m:sheetViews/m:sheetView/m:selection", "activeCell") == ["B7"]


def test_a_span_shift_that_would_split_a_merge_is_refused():
    _, ws = _sheet('<mergeCells count="2"><mergeCell ref="B7:C8"/><mergeCell ref="A5:C6"/></mergeCells>')
    with pytest.raises(
        InvalidRangeError, match="in columns B:C would move only part of merged cells A5:C6"
    ) as err:
        _shift(ws, _Shift(True, 5, 1, (2, 3)))
    assert "(Excel: This operation will cause some merged cells to unmerge." in str(err.value)
    assert ws.merged_ranges == ["B7:C8", "A5:C6"]


def test_a_merge_over_every_row_across_the_span_edge_is_refused_too():
    _, ws = _sheet('<mergeCells count="1"><mergeCell ref="A1:B1048576"/></mergeCells>')
    with pytest.raises(InvalidRangeError, match="would move only part of merged cells A1:B1048576"):
        _shift(ws, _Shift(True, 5, 1, (1, 1)))


def test_deleted_frozen_columns_shrink_the_split():
    _, ws = _sheet(_FROZEN)
    _shift(ws, _Shift(False, 1, -2))
    assert _attrs(ws, "m:sheetViews/m:sheetView/m:pane", "xSplit") == ["1"]
    assert _attrs(ws, "m:sheetViews/m:sheetView/m:pane", "topLeftCell") == ["B5"]


def test_a_deleted_active_cell_is_pinned_to_the_band():
    _, ws = _sheet(_FROZEN)
    _shift(ws, _rows(7, 3))
    assert _attrs(ws, "m:sheetViews/m:sheetView/m:selection", "activeCell") == ["D7"]
    assert _attrs(ws, "m:sheetViews/m:sheetView/m:selection", "sqref") == ["D7:E7"]


def test_deleting_columns_shrinks_and_drops_column_widths_and_row_spans():
    wb, ws = _sheet(
        '<cols><col min="1" max="2" width="9"/><col min="3" max="3" width="5"/>'
        '<col min="4" max="6" width="7"/><col min="8" max="16384" width="3"/></cols>'
    )
    sheet_data = ws._part.element.find("m:sheetData", NS)
    for r, spans in (("1", "1:5"), ("2", "3:4"), ("3", "2:2 6:9")):
        etree.SubElement(sheet_data, f"{{{SML_NS}}}row", r=r, spans=spans)
    _shift(ws, _Shift(False, 3, -2))
    assert _attrs(ws, "m:cols/m:col", "min") == ["1", "3", "6"]
    assert _attrs(ws, "m:cols/m:col", "max") == ["2", "4", "16384"]  # the last still reaches XFD
    assert _attrs(ws, "m:sheetData/m:row", "spans") == ["1:3", None, "2:2 4:7"]
    assert_valid_package(wb)


def test_a_span_shift_leaves_whole_line_structures_alone():
    _, ws = _sheet(
        _BREAKS,
        _FROZEN,
        '<cols><col min="2" max="3" width="9"/></cols>',
        '<conditionalFormatting sqref="B7:C8 A7"><cfRule type="expression" priority="1">'
        "<formula>TRUE</formula></cfRule></conditionalFormatting>",
    )
    _shift(ws, _Shift(True, 5, -1, (2, 3)))
    assert _attrs(ws, "m:rowBreaks/m:brk", "id") == ["5", "10", "20"]
    assert _attrs(ws, "m:sheetViews/m:sheetView/m:pane", "ySplit") == ["4"]
    assert _attrs(ws, "m:sheetViews/m:sheetView/m:pane", "topLeftCell") == ["C5"]
    assert _attrs(ws, "m:conditionalFormatting", "sqref") == ["B6:C7 A7"]


# -- references in text -------------------------------------------------------


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("SUM(Sheet1!$A$5:$A$6)", "SUM(Sheet1!#REF!)"),
        ("SUM(Sheet1!$A$4:$A$8)+Sheet1!B7", "SUM(Sheet1!$A$4:$A$6)+Sheet1!B5"),
        ("SUM(A5:A6)+A9", "SUM(#REF!)+A7"),
        ("Sheet1!$5:$6", "Sheet1!#REF!"),
        ("Sheet1!$A:$A", "Sheet1!$A:$A"),
        ("Other!A5+SUM(Sheet1:Other!A5)", "Other!A5+SUM(Sheet1:Other!A5)"),
    ],
)
def test_rewrite_formula_refs_on_a_delete(text, expected):
    assert rewrite_formula_refs(text, "Sheet1", _rows(5, 2).area, unqualified=True) == expected
