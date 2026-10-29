"""``shift_cells``: Excel's "insert cells" and "delete cells", which no public method runs yet.

The expected results and Excel's messages were recorded from Excel 16.0.20527
for Microsoft 365 on Windows (2026-10-29). Each case ran ``Range.Insert`` or
``Range.Delete`` with a shift direction on a scratch workbook through COM,
with alerts on so that Excel's prompt could be read and answered Cancel, and
the saved workbook was read back. On ``TableWithContentBelow.xlsx``,
``ListRows.Add`` and "insert cells A7:D7, shift down" left the sheet the same.
"""

from __future__ import annotations

import re

import pytest
from lxml import etree

from xlsxedit import FormulaGroupError, GridOverflowError, InvalidRangeError, TableError, Workbook
from xlsxedit.drawing import _iter_anchors, drawing_parts_for_worksheet
from xlsxedit.opc.constants import SML_NS
from xlsxedit.oxml.address import index_to_col
from xlsxedit.range_set import CellRange
from xlsxedit.row_shift import CellShiftPlan, _anchor_extent, _Shift, shift_cells
from xlsxedit.worksheet_order import insert_worksheet_child
from tests.conftest import INSPECT_FIXTURES
from tests.preservation import assert_preserved, check_consistency, read_pkg
from tests.schema_validate import assert_valid_package
from tests.test_formula_hygiene import _plant
from tests.test_insert_rows import _add_defined_name, _defined_name_text

NS = {"m": SML_NS}

EXCEL_TABLE = "This won't work because it would move cells in a table on your worksheet."
EXCEL_MERGE = "This operation will cause some merged cells to unmerge."
EXCEL_ARRAY = "You can't change part of an array."
EXCEL_DATA_TABLE = "Can't change part of a data table."
EXCEL_GRID = "Microsoft Excel can't insert new cells because it would push non-empty cells off the end"
EXCEL_OBJECTS = "Can't push objects off the sheet."

# what a shift that moves cells always changes, and the calcChain it drops
WORKBOOK_PARTS = {"xl/workbook.xml", "xl/_rels/workbook.xml.rels"}
CALC_CHAIN = {"xl/calcChain.xml"}


def _cells(ref: str, direction: str) -> _Shift:
    """The shift of Excel's ``Range(ref).Insert`` (down, right) or ``.Delete`` (up, left)."""
    area = CellRange.parse(ref)
    if direction in ("down", "up"):
        count = area.max_row - area.min_row + 1
        span = (area.min_col, area.max_col)
        return _Shift(True, area.min_row, count if direction == "down" else -count, span)
    count = area.max_col - area.min_col + 1
    span = (area.min_row, area.max_row)
    return _Shift(False, area.min_col, count if direction == "right" else -count, span)


def _address(col: int, row: int) -> str:
    return f"{index_to_col(col - 1)}{row}"


def _filled(ref: str = "A1:H20"):
    """A sheet whose cells in ``ref`` hold their own address, as ``"vA1"``."""
    wb = Workbook.create()
    ws = wb["Sheet1"]
    area = CellRange.parse(ref)
    columns = range(area.min_col, area.max_col + 1)
    rows = [[f"v{_address(c, r)}" for c in columns] for r in range(area.min_row, area.max_row + 1)]
    ws.write_rows(rows, at_cell=_address(area.min_col, area.min_row))
    return wb, ws


def _values(ws) -> dict[str, str]:
    """Each cell with a value, by address."""
    out = {}
    for c in ws._part.element.iterfind("m:sheetData/m:row/m:c", NS):
        value = ws[c.get("r")].value
        if value is not None:
            out[c.get("r")] = value
    return out


def _moved(ws) -> dict[str, str]:
    """The cells of a ``_filled`` sheet that no longer hold their own address."""
    return {a: v for a, v in _values(ws).items() if isinstance(v, str) and v != f"v{a}"}


def _table(ws, ref: str, name: str):
    area = CellRange.parse(ref)
    header = [ws[_address(c, area.min_row)].value for c in range(area.min_col, area.max_col + 1)]
    return ws.add_table(ref, header, name=name)


def _refused(wb, ws, shift: _Shift, error: type[Exception], excel: str) -> None:
    before = read_pkg(wb)
    with pytest.raises(error, match=re.escape(excel)):
        shift_cells(ws, shift)
    assert_preserved(before, read_pkg(wb))


def _rows(ws) -> dict[int, list[str]]:
    return {
        int(row.get("r")): [c.get("r") for c in row.iterfind("m:c", NS)]
        for row in ws._part.element.iterfind("m:sheetData/m:row", NS)
    }


# -- the fixture the issue names ----------------------------------------------


def test_cells_below_a_table_move_down_within_its_columns_as_in_excel():
    """The note and the merged title below the table move; the content right of
    the table stays, and so does the chart, which covers columns A:E."""
    wb = Workbook.open(INSPECT_FIXTURES["TableWithContentBelow"])
    ws = wb["Report"]
    before = read_pkg(wb)
    beside = {a: v for a, v in _values(ws).items() if a[0] in "FG"}
    shift_cells(ws, _cells("A7:D7", "down"))
    after = read_pkg(wb)
    assert_preserved(
        before,
        after,
        expected_changed={"xl/worksheets/sheet1.xml", *WORKBOOK_PARTS},
        expected_removed=CALC_CHAIN,
    )
    assert check_consistency(after) == []
    assert_valid_package(wb)
    values = _values(ws)
    assert values["A8"] == "Figures in thousands of units."
    assert values["A9"] == "Quarterly sales by region"
    assert "A7" not in values
    assert {a: v for a, v in values.items() if a[0] in "FG"} == beside
    assert ws.merged_ranges == ["A9:D9"]
    assert [t.ref for t in ws.tables] == ["A1:D6"]
    rows = _rows(ws)
    assert 7 not in rows  # left with nothing, as Excel leaves it out
    assert rows[8] == ["A8"] and rows[9] == ["A9", "B9", "C9", "D9"]
    assert ws._part.element.find("m:dimension", NS).get("ref") == "A1:G9"


def test_the_fixture_chart_moves_only_with_a_band_over_every_column_it_covers():
    """Excel moved the chart for "insert cells A7:E7" and not for A7:D7: its
    ``xdr:to`` marker reaches column F at offset 0, so it covers A:E."""
    ws = Workbook.open(INSPECT_FIXTURES["TableWithContentBelow"])["Report"]
    (anchor, _), = _iter_anchors(drawing_parts_for_worksheet(ws)[0].element, None)
    markers, cells = _anchor_extent(anchor)
    assert (str(markers), str(cells)) == ("A10:F24", "A10:E23")
    assert not _cells("A7:D7", "down").carries(cells)
    assert _cells("A7:E7", "down").carries(cells)


# -- moving the cells ---------------------------------------------------------


def test_only_the_cells_in_the_band_move_down_into_rows_made_for_them():
    wb, ws = _filled()
    ws.row_dimensions[8].height = 30
    shift_cells(ws, _cells("A7:D7", "down"))
    moved = _moved(ws)
    assert moved["A8"] == "vA7" and moved["D21"] == "vD20"
    assert len(moved) == 4 * 14  # A7:D20, each one row down
    assert all(f"{c}7" not in _values(ws) for c in "ABCD")
    assert _values(ws)["E7"] == "vE7"
    assert ws.row_dimensions[8].height == 30  # rows keep their height where they are
    assert _rows(ws)[21] == ["A21", "B21", "C21", "D21"]
    assert_valid_package(wb)


def test_cells_moving_into_a_row_go_between_its_other_cells_in_column_order():
    wb, ws = _filled("A1:F3")
    for address in ("B5", "E5"):
        ws[address].value = address
    shift_cells(ws, _cells("C2:D3", "down"))  # C2:D3 land on rows 4:5
    assert _rows(ws)[5] == ["B5", "C5", "D5", "E5"]
    assert _rows(ws)[4] == ["C4", "D4"]
    assert _values(ws)["C5"] == "vC3"


def test_a_delete_drops_the_band_cells_and_moves_the_rest_up():
    wb, ws = _filled()
    shift_cells(ws, _cells("B5:C6", "up"))
    values = _values(ws)
    assert values["B5"] == "vB7" and values["C18"] == "vC20"
    assert "B19" not in values and "B20" not in values
    assert values["A5"] == "vA5" and values["D20"] == "vD20"
    assert_valid_package(wb)


@pytest.mark.parametrize(
    ("ref", "direction", "address", "value"),
    [("E5:E8", "right", "F5", "vE5"), ("C5:D8", "left", "C5", "vE5")],
)
def test_a_column_shift_moves_cells_within_their_rows(ref, direction, address, value):
    wb, ws = _filled()
    shift_cells(ws, _cells(ref, direction))
    assert _values(ws)[address] == value
    assert _values(ws)["E4"] == "vE4" and _values(ws)["E9"] == "vE9"
    assert_valid_package(wb)


def test_a_row_left_with_only_its_own_attributes_stays():
    wb, ws = _filled("A1:A3")
    ws.row_dimensions[2].height = 25
    shift_cells(ws, _cells("A2:A2", "down"))
    assert _rows(ws)[2] == []  # its height is its own
    assert ws.row_dimensions[2].height == 25
    assert _rows(ws)[4] == ["A4"]


def test_cells_and_rows_that_only_imply_their_address_get_it_written():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    sheet_data = ws._part.element.find("m:sheetData", NS)
    rows = '<row r="2"><c r="A2"><v>1</v></c><c><v>2</v></c></row><row><c r="A3"><v>3</v></c></row>'
    sheet_data.extend(etree.fromstring(f'<w xmlns="{SML_NS}">{rows}</w>'))
    shift_cells(ws, _cells("A2:A2", "down"))
    assert _rows(ws) == {2: ["B2"], 3: ["A3"], 4: ["A4"]}
    assert (ws["B2"].value, ws["A3"].value, ws["A4"].value) == (2, 1, 3)


def test_a_column_shift_writes_the_row_numbers_the_xml_only_implied():
    """Else removing the row it empties would renumber the rows after it."""
    wb = Workbook.create()
    ws = wb["Sheet1"]
    rows = "<row><c><v>1</v></c></row><row><c><v>2</v></c></row><row><c><v>3</v></c></row>"
    ws._part.element.find("m:sheetData", NS).extend(etree.fromstring(f'<w xmlns="{SML_NS}">{rows}</w>'))
    shift_cells(ws, _cells("A1:A1", "left"))
    rows = ws._part.element.iterfind("m:sheetData/m:row", NS)
    assert [(row.get("r"), row.findtext("m:c/m:v", namespaces=NS)) for row in rows] == [("2", "2"), ("3", "3")]


@pytest.mark.parametrize("address", ["1A", "A0"])
def test_an_unreadable_cell_address_anywhere_on_the_sheet_refuses_the_shift(unchecked_workbooks, address):
    """Refreshing ``<dimension>`` after the move reads every address."""
    wb, ws = _filled("A1:B6")
    ws._part.element.find("m:sheetData/m:row/m:c", NS).set("r", address)
    _refused(wb, ws, _cells("A5:A5", "down"), InvalidRangeError, f"invalid cell address {address!r}")


@pytest.mark.parametrize(("table", "ref"), [("F1:G6", "A3:D3"), ("A1:B4", "A6:B6")])
def test_a_table_beside_or_above_the_shift_is_left_alone(unchecked_workbooks, table, ref):
    """Even when its ``autoFilter`` range cannot be read."""
    wb, ws = _filled()
    (part,) = [_table(ws, table, "Sales")._part]
    part.element.find("m:autoFilter", NS).set("ref", table.split(":")[0] + ":Z")
    name = str(part.partname).lstrip("/")
    before = read_pkg(wb)
    shift_cells(ws, _cells(ref, "down"))
    assert read_pkg(wb)[name] == before[name]


def test_a_data_table_deleted_with_its_input_cell_goes_with_it():
    wb, ws = _filled()
    _plant(ws["B10"], {}, '<f t="dataTable" ref="B10:B12" dt2D="0" dtr="0" r1="A15"/><v>0</v>')
    shift_cells(ws, _cells("A9:B16", "up"))
    assert not list(ws._part.element.iter(f"{{{SML_NS}}}f"))
    assert _values(ws)["A9"] == "vA17"


def test_a_cell_with_only_a_style_pushed_off_the_grid_goes():
    wb, ws = _filled("A1:B2")
    ws["B1048576"].apply_style(bold=True)
    shift_cells(ws, _cells("A2:D2", "down"))
    assert ws._find_cell_element("B1048576") is None
    assert _values(ws)["B3"] == "vB2"


def test_a_shift_needs_a_span():
    ws = Workbook.create()["Sheet1"]
    with pytest.raises(ValueError, match="span"):
        CellShiftPlan(ws, _Shift(True, 5, 1))


# -- Excel's refusals ---------------------------------------------------------


@pytest.mark.parametrize(
    ("table", "ref", "direction"),
    [
        ("C10:F12", "A7:D7", "down"),  # across the band's edge
        ("A1:D6", "A4:B4", "down"),  # cut through, narrower than the table
        ("A1:D6", "A4:F4", "down"),  # cut through, wider than the table
        ("G7:H10", "E5:E8", "right"),
        ("C10:F12", "A7:D7", "up"),
        ("A1:D6", "A4:F4", "up"),
    ],
)
def test_a_shift_that_moves_only_part_of_a_table_is_refused_as_in_excel(table, ref, direction):
    wb, ws = _filled()
    _table(ws, table, "Sales")
    _refused(wb, ws, _cells(ref, direction), TableError, EXCEL_TABLE)


@pytest.mark.parametrize(
    ("merge", "ref", "direction"),
    [
        ("C10:F10", "A7:D7", "down"),
        ("C6:F8", "A7:D7", "down"),
        ("G8:H9", "E5:E8", "right"),
        ("C10:F10", "A7:D7", "up"),
    ],
)
def test_a_shift_across_a_merge_is_refused_where_excel_asks_to_unmerge(merge, ref, direction):
    wb, ws = _filled()
    ws.merge_cells(merge)
    _refused(wb, ws, _cells(ref, direction), InvalidRangeError, EXCEL_MERGE)


_ARRAY_A9 = '<f t="array" ref="A9:F9">A1:F1</f><v>0</v>'
_ARRAY_A6 = '<f t="array" ref="A6:A8">B1:B3</f><v>0</v>'
_DATA_TABLE = '<f t="dataTable" ref="B10:B12" dt2D="0" dtr="0" r1="B1"/><v>0</v>'


@pytest.mark.parametrize(
    ("cell", "inner", "ref", "direction", "excel"),
    [
        ("A9", _ARRAY_A9, "A7:D7", "down", EXCEL_ARRAY),
        ("A6", _ARRAY_A6, "A7:D7", "down", EXCEL_ARRAY),
        ("A6", _ARRAY_A6, "A7:D7", "up", EXCEL_ARRAY),
        ("A1", '<f t="array" ref="A1:B1048576">C1:D1048576</f><v>0</v>', "A5:A5", "down", EXCEL_ARRAY),
        # results B10:B12; Excel's table A9:B12 includes the formula row and input column
        ("B10", _DATA_TABLE, "A7:A7", "down", EXCEL_DATA_TABLE),
    ],
)
def test_a_shift_that_changes_part_of_an_array_or_data_table_is_refused_as_in_excel(
    cell, inner, ref, direction, excel
):
    wb, ws = _filled()
    _plant(ws[cell], {}, inner)
    _refused(wb, ws, _cells(ref, direction), FormulaGroupError, excel)


@pytest.mark.parametrize(("ref", "direction"), [("A7:D7", "down"), ("A6:D6", "up")])
def test_a_shift_that_splits_a_shared_formula_is_refused(ref, direction):
    """Excel turns the moved (or remaining) cells into separate formulas."""
    wb, ws = _filled()
    _plant(ws["A6"], {}, '<f t="shared" ref="A6:A8" si="0">B1</f><v>0</v>')
    for address in ("A7", "A8"):
        _plant(ws[address], {}, '<f t="shared" si="0"/><v>0</v>')
    _refused(wb, ws, _cells(ref, direction), FormulaGroupError, "shared formula range A6:A8 of A6")


@pytest.mark.parametrize(
    ("cell", "ref", "direction"), [("B1048576", "A7:D7", "down"), ("XFD6", "E5:E8", "right")]
)
def test_content_pushed_off_the_grid_is_refused_as_in_excel(cell, ref, direction):
    wb, ws = _filled()
    ws[cell].value = "last"
    _refused(wb, ws, _cells(ref, direction), GridOverflowError, EXCEL_GRID)


def test_content_off_the_band_at_the_last_row_does_not_block():
    wb, ws = _filled()
    ws["F1048576"].value = "last"
    shift_cells(ws, _cells("A7:D7", "down"))
    assert ws["F1048576"].value == "last"


def _drawing_at_the_bottom(name: str):
    """DrawingAnchors with the object ``name`` moved to rows 1048570:1048575."""
    wb = Workbook.open(INSPECT_FIXTURES["DrawingAnchors"])
    ws = wb["Anchors"]
    drawing = drawing_parts_for_worksheet(ws)[0]
    for anchor, _ in _iter_anchors(drawing.element, None):
        if anchor.find(f".//*[@name={name!r}]") is not None:
            for marker, row in (("from", 1_048_569), ("to", 1_048_574)):
                anchor.find(f"{{*}}{marker}/{{*}}row").text = str(row)
            drawing.mark_dirty()
            return wb, ws
    raise AssertionError(name)


@pytest.mark.parametrize(("name", "band"), [("Move and size", "B7:B16"), ("Absolute", "E7:G16")])
def test_an_object_pushed_off_the_sheet_is_refused_as_in_excel_whatever_its_placement(name, band):
    wb, ws = _drawing_at_the_bottom(name)
    _refused(wb, ws, _cells(band, "down"), GridOverflowError, EXCEL_OBJECTS)


def test_an_absolute_object_in_the_shifted_region_does_not_block():
    """Excel leaves it where it is on the page; its anchor is not touched here."""
    wb = Workbook.open(INSPECT_FIXTURES["DrawingAnchors"])
    ws = wb["Anchors"]
    for address in ("E3", "F3", "G9"):
        ws[address].value = 1
    before = read_pkg(wb)
    shift_cells(ws, _cells("E1:G1", "down"))
    assert ws["E4"].value == 1 and ws["G10"].value == 1
    assert_preserved(before, read_pkg(wb), expected_changed={"xl/worksheets/sheet1.xml", "xl/workbook.xml"})


def test_notes_in_the_shifted_region_do_not_block():
    """Excel moves a note with its cell whatever its placement (the default is
    free floating); xlsxedit moves no comments, as for inserted rows."""
    wb = Workbook.open(INSPECT_FIXTURES["Preservation"])
    ws = wb["Report"]
    before = read_pkg(wb)
    shift_cells(ws, _cells("A2:A2", "down"))
    report = assert_preserved(
        before,
        read_pkg(wb),
        expected_changed={"xl/worksheets/sheet2.xml", *WORKBOOK_PARTS},
        expected_removed=CALC_CHAIN,
    )
    assert not report.changed & {"xl/comments1.xml", "xl/drawings/vmlDrawing1.vml"}
    rows = {r.get("r"): r for r in ws._part.element.iterfind("m:sheetData/m:row", NS)}
    assert [n for n, r in rows.items() if r.get("hidden")] == ["8", "9", "10", "13"]  # stay put


@pytest.mark.parametrize(
    ("ref", "direction", "lost"),
    [
        ("A1:D1", "up", "the header row"),
        ("A2:D6", "up", "every data row"),
        ("C1:C6", "right", "add columns to"),
        ("C1:C6", "left", "delete columns of"),
    ],
)
def test_table_changes_excel_makes_and_xlsxedit_does_not_are_refused(ref, direction, lost):
    wb, ws = _filled()
    _table(ws, "A1:D6", "Sales")
    _refused(wb, ws, _cells(ref, direction), TableError, lost)


def test_deleting_a_totals_row_is_refused():
    wb, ws = _filled()
    table = _table(ws, "A1:D6", "Sales")
    table.totals_row_count = 1
    _refused(wb, ws, _cells("A6:D6", "up"), TableError, "the totals row")


# -- what moves with the cells ------------------------------------------------


@pytest.mark.parametrize(
    ("table", "ref", "direction", "expected"),
    [
        ("A10:B12", "A7:D7", "down", "A11:B13"),
        ("C2:F4", "A7:D7", "down", "C2:F4"),
        ("A1:D6", "A4:D4", "down", "A1:D7"),
        ("A3:D6", "A3:D3", "down", "A4:D7"),
        ("A3:D6", "A3:F3", "down", "A4:D7"),
        ("A1:D6", "A4:D4", "up", "A1:D5"),
        ("A10:D12", "A2:F3", "up", "A8:D10"),
        ("B1:D6", "A1:A6", "right", "C1:E6"),
    ],
)
def test_tables_move_grow_and_shrink_as_in_excel(table, ref, direction, expected):
    wb, ws = _filled()
    _table(ws, table, "Sales")
    shift_cells(ws, _cells(ref, direction))
    (moved,) = ws.tables
    assert moved.ref == expected
    assert_valid_package(wb)


def test_rows_inserted_at_a_totals_row_join_the_data_as_in_excel():
    wb, ws = _filled()
    table = _table(ws, "A1:D6", "Sales")
    table.totals_row_count = 1
    shift_cells(ws, _cells("A6:D6", "down"))
    assert table.ref == "A1:D7"
    assert table._element.find("m:autoFilter", NS).get("ref") == "A1:D6"


@pytest.mark.parametrize(
    ("merge", "ref", "direction", "expected"),
    [
        ("A6:B8", "A7:D7", "down", "A6:B9"),
        ("A7:B8", "A7:D7", "down", "A8:B9"),
        ("A10:B11", "A7:D7", "down", "A11:B12"),
        ("C3:F3", "A7:D7", "down", "C3:F3"),
        ("A6:B8", "A7:D7", "up", "A6:B7"),
    ],
)
def test_merges_in_the_band_move_grow_and_shrink_as_in_excel(merge, ref, direction, expected):
    wb, ws = _filled()
    ws.merge_cells(merge)
    shift_cells(ws, _cells(ref, direction))
    assert ws.merged_ranges == [expected]


def test_formats_across_the_band_split_and_their_x14_twins_with_them():
    """Excel splits a format's ranges in horizontal strips and keeps its formula."""
    wb = Workbook.open(INSPECT_FIXTURES["ConditionalFormattingRich"])
    ws = wb["Scores"]
    before = read_pkg(wb)
    shift_cells(ws, _cells("A5:D5", "down"))
    after = read_pkg(wb)
    assert_preserved(before, after, expected_changed={"xl/worksheets/sheet1.xml", "xl/workbook.xml"})
    assert check_consistency(after) == []
    root = ws._part.element
    assert [cf.get("sqref") for cf in root.iterfind("m:conditionalFormatting", NS)] == [
        "A2:E4 E5:E11 A5:D12",
        "C2:C12 E2:E11",
        "B2:B12",
        "I2:I8",
    ]
    assert [f.text for f in root.iterfind("m:conditionalFormatting/m:cfRule/m:formula", NS)][0] == "$B2>0"
    assert [s.text for s in root.iter("{*}sqref")] == ["D2:D12", "B2:B12"]
    assert [t.ref for t in ws.tables] == ["G1:I8"]
    assert_valid_package(wb)


@pytest.mark.parametrize(
    ("sqref", "ref", "direction", "expected", "formula"),
    [
        ("A5:F10", "A7:D7", "down", "A5:F6 E7:F10 A7:D11", "ISTEXT(A5)"),
        ("A5:F10", "C7:D7", "down", "A5:F6 A7:B10 E7:F10 C7:D11", "ISTEXT(A5)"),
        ("C9:F12", "A7:D7", "down", "E9:F12 C10:D13", "ISTEXT(C9)"),
        ("A5:F10", "A7:D7", "up", "A5:F6 E7:F10 A7:D9", "ISTEXT(A5)"),
        ("C3:H10", "E5:E8", "right", "C3:H4 C9:H10 C5:I8", "ISTEXT(C3)"),
        ("A5:D10", "A7:D7", "down", "A5:D11", "ISTEXT(A5)"),
        ("C:F", "A7:D7", "down", None, "ISTEXT(C1)"),  # over every row: unchanged
    ],
)
def test_formats_and_validations_across_the_band_split_as_in_excel(sqref, ref, direction, expected, formula):
    wb, ws = _filled()
    first = CellRange.parse(sqref)
    anchor = str(CellRange.from_bounds(first.min_col, first.min_row, first.min_col, first.min_row))
    ws.add_conditional_formatting(sqref, operator="notEqual", formula=f"ISTEXT({anchor})")
    ws.add_data_validation(sqref, type="textLength", operator="lessThanOrEqual", formula1="100")
    root = ws._part.element
    paths = ("m:conditionalFormatting", "m:dataValidations/m:dataValidation")
    before = [[e.get("sqref") for e in root.iterfind(path, NS)] for path in paths]
    shift_cells(ws, _cells(ref, direction))
    after = [[e.get("sqref") for e in root.iterfind(path, NS)] for path in paths]
    assert after == (before if expected is None else [[expected], [expected]])
    assert root.findtext("m:conditionalFormatting/m:cfRule/m:formula", namespaces=NS) == formula


@pytest.mark.parametrize(
    ("link", "ref", "direction", "expected"),
    [
        ("A9:F9", "A7:D7", "down", "A9:F9"),
        ("A9:B9", "A7:D7", "down", "A10:B10"),
        ("A9:F9", "A7:D7", "up", "A9:F9"),
    ],
)
def test_hyperlinks_in_the_band_move_and_those_across_it_stay_as_in_excel(link, ref, direction, expected):
    wb, ws = _filled()
    links = f'<hyperlinks xmlns="{SML_NS}"><hyperlink ref="{link}" location="Sheet1!A1"/></hyperlinks>'
    insert_worksheet_child(ws._part.element, etree.fromstring(links))
    shift_cells(ws, _cells(ref, direction))
    assert [h.get("ref") for h in ws._part.element.iterfind("m:hyperlinks/m:hyperlink", NS)] == [expected]


def test_formula_ranges_in_the_band_move_whole():
    wb, ws = _filled()
    _plant(ws["A9"], {}, '<f t="array" ref="A9:B10">C1:D2</f><v>0</v>')
    _plant(ws["C12"], {}, '<f t="dataTable" ref="C12:C14" dt2D="0" dtr="0" r1="B1"/><v>0</v>')
    _plant(ws["B16"], {}, '<f t="shared" ref="B16:C16" si="0">A1</f><v>0</v>')
    _plant(ws["C16"], {}, '<f t="shared" si="0"/><v>0</v>')
    shift_cells(ws, _cells("A7:D7", "down"))
    refs = {f.getparent().get("r"): f.attrib for f in ws._part.element.iter(f"{{{SML_NS}}}f")}
    assert refs["A10"]["ref"] == "A10:B11"
    assert (refs["C13"]["ref"], refs["C13"]["r1"]) == ("C13:C15", "B1")
    assert refs["B17"]["ref"] == "B17:C17" and "C17" in refs


def test_references_elsewhere_move_only_within_the_band():
    """As Excel moved the references in cell formulas: one across the band stays."""
    wb, ws = _filled()
    _add_defined_name(wb, "Inside", "Sheet1!$A$9:$B$10")
    _add_defined_name(wb, "Across", "Sheet1!$A$9:$F$10")
    shift_cells(ws, _cells("A7:D7", "down"))
    assert _defined_name_text(wb, "Inside") == "Sheet1!$A$10:$B$11"
    assert _defined_name_text(wb, "Across") == "Sheet1!$A$9:$F$10"
    assert_valid_package(wb)
