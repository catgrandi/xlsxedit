"""``CellRange.shift`` on deletes and span-restricted shifts, against Excel.

The tables were recorded from Excel for Microsoft 365 on Windows (2026-10-19):
each reference sat in a formula, Excel carried out the operation, and the
tables hold the reference Excel wrote back. ``#REF!`` means Excel turned the
reference into a reference error, which ``shift`` reports as ``None``.
"""

from __future__ import annotations

from dataclasses import replace

import pytest

from xlsxedit.range_set import COLUMNS, RANGE, CellRange, SheetRangeSet, shift_bounds


def _as_excel_writes(area: CellRange | None) -> str:
    """Excel writes a range over every row as whole columns (``A1:A1048576`` as ``A:A``)."""
    if area is None:
        return "#REF!"
    if area.kind == RANGE and area.spans_all_rows:
        area = replace(area, kind=COLUMNS, text=None)
    return str(area)


def _check(ref: str, results: str, ops) -> None:
    area = CellRange.parse(ref)
    for op, expected in zip(ops, results.split(), strict=True):
        assert _as_excel_writes(area.shift(*op)) == expected, (ref, op)


# insert 1 row at 5, delete rows 5:6, delete rows 3:5, delete rows 1048575:1048576
WHOLE_ROWS = [(True, 5, 1), (True, 5, -2), (True, 3, -3), (True, 1_048_575, -2)]
# insert 1 column at C, delete columns C:D
WHOLE_COLS = [(False, 3, 1), (False, 3, -2)]
# insert cells B5:C5 shifting down, delete cells B5:C6 shifting up,
# insert cells C3:C5 shifting right, delete cells C3:D5 shifting left
SPANS = [(True, 5, 1, (2, 3)), (True, 5, -2, (2, 3)), (False, 3, 1, (3, 5)), (False, 3, -2, (3, 5))]

_WHOLE_ROWS_TABLE = [
    ("A5",                "A6                #REF!             #REF!             A5"),
    ("$A$5",              "$A$6              #REF!             #REF!             $A$5"),
    ("A4:A8",             "A4:A9             A4:A6             A3:A5             A4:A8"),
    ("A5:A6",             "A6:A7             #REF!             A3:A3             A5:A6"),
    ("A3:A5",             "A3:A6             A3:A4             #REF!             A3:A5"),
    ("A6:A9",             "A7:A10            A5:A7             A3:A6             A6:A9"),
    ("A7:A9",             "A8:A10            A5:A7             A4:A6             A7:A9"),
    ("A2:A12",            "A2:A13            A2:A10            A2:A9             A2:A12"),
    ("A:A",               "A:A               A:A               A:A               A:A"),
    ("A1:A1048576",       "A:A               A:A               A:A               A:A"),
    ("5:10",              "6:11              5:8               3:7               5:10"),
    ("5:6",               "6:7               #REF!             3:3               5:6"),
    ("$5:$6",             "$6:$7             #REF!             $3:$3             $5:$6"),
    ("4:4",               "4:4               4:4               #REF!             4:4"),
    ("A1048570:A1048576", "A1048571:A1048576 A1048568:A1048576 A1048567:A1048576 A1048570:A1048576"),
    ("B1:B10",            "B1:B11            B1:B8             B1:B7             B1:B10"),
    ("A1:D10",            "A1:D11            A1:D8             A1:D7             A1:D10"),
    ("B1:E10",            "B1:E11            B1:E8             B1:E7             B1:E10"),
    ("C3",                "C3                C3                #REF!             C3"),
    ("B7",                "B8                B5                B4                B7"),
    ("D7",                "D8                D5                D4                D7"),
    ("B:B",               "B:B               B:B               B:B               B:B"),
    ("5:5",               "6:6               #REF!             #REF!             5:5"),
    ("B5:C5",             "B6:C6             #REF!             #REF!             B5:C5"),
    ("B4:C6",             "B4:C7             B4:C4             B3:C3             B4:C6"),
    ("A5:B5",             "A6:B6             #REF!             #REF!             A5:B5"),
    ("B6:C9",             "B7:C10            B5:C7             B3:C6             B6:C9"),
    ("B5:B6",             "B6:B7             #REF!             B3:B3             B5:B6"),
    ("C2:C3",             "C2:C3             C2:C3             C2:C2             C2:C3"),
    ("B1:C1048576",       "B:C               B:C               B:C               B:C"),
    ("B3:C20",            "B3:C21            B3:C18            B3:C17            B3:C20"),
    ("D5:E8",             "D6:E9             D5:E6             D3:E5             D5:E8"),
    ("C5:D8",             "C6:D9             C5:D6             C3:D5             C5:D8"),
    ("B2:H2",             "B2:H2             B2:H2             B2:H2             B2:H2"),
    ("E3",                "E3                E3                #REF!             E3"),
    ("E5:G5",             "E6:G6             #REF!             #REF!             E5:G5"),
    ("C:D",               "C:D               C:D               C:D               C:D"),
    ("B:H",               "B:H               B:H               B:H               B:H"),
    ("D:D",               "D:D               D:D               D:D               D:D"),
    ("A1048576",          "#REF!             A1048574          A1048573          #REF!"),
    ("A1048575:A1048576", "A1048576:A1048576 A1048573:A1048576 A1048572:A1048576 #REF!"),
    ("A5:A1048576",       "A6:A1048576       A5:A1048576       A3:A1048576       A5:A1048576"),
    ("A1048576:A1048576", "#REF!             A1048574:A1048576 A1048573:A1048576 #REF!"),
    ("1048570:1048576",   "1048571:1048576   1048568:1048576   1048567:1048576   1048570:1048576"),
    ("A5:D5",             "A6:D6             #REF!             #REF!             A5:D5"),
    ("B5:D5",             "B6:D6             #REF!             #REF!             B5:D5"),
    ("A4:B5",             "A4:B6             A4:B4             #REF!             A4:B5"),
    ("A5:C6",             "A6:C7             #REF!             A3:C3             A5:C6"),
    ("XFD1",              "XFD1              XFD1              XFD1              XFD1"),
    ("XFA1:XFD1",         "XFA1:XFD1         XFA1:XFD1         XFA1:XFD1         XFA1:XFD1"),
    ("C1:XFD1",           "C1:XFD1           C1:XFD1           C1:XFD1           C1:XFD1"),
    ("A6:B7",             "A7:B8             A5:B5             A3:B4             A6:B7"),
    ("A5:A5",             "A6:A6             #REF!             #REF!             A5:A5"),
]

_WHOLE_COLS_TABLE = [
    ("A5",                "A5                A5"),
    ("$A$5",              "$A$5              $A$5"),
    ("A4:A8",             "A4:A8             A4:A8"),
    ("A5:A6",             "A5:A6             A5:A6"),
    ("A3:A5",             "A3:A5             A3:A5"),
    ("A6:A9",             "A6:A9             A6:A9"),
    ("A7:A9",             "A7:A9             A7:A9"),
    ("A2:A12",            "A2:A12            A2:A12"),
    ("A:A",               "A:A               A:A"),
    ("A1:A1048576",       "A:A               A:A"),
    ("5:10",              "5:10              5:10"),
    ("5:6",               "5:6               5:6"),
    ("$5:$6",             "$5:$6             $5:$6"),
    ("4:4",               "4:4               4:4"),
    ("A1048570:A1048576", "A1048570:A1048576 A1048570:A1048576"),
    ("B1:B10",            "B1:B10            B1:B10"),
    ("A1:D10",            "A1:E10            A1:B10"),
    ("B1:E10",            "B1:F10            B1:C10"),
    ("C3",                "D3                #REF!"),
    ("B7",                "B7                B7"),
    ("D7",                "E7                #REF!"),
    ("B:B",               "B:B               B:B"),
    ("5:5",               "5:5               5:5"),
    ("B5:C5",             "B5:D5             B5:B5"),
    ("B4:C6",             "B4:D6             B4:B6"),
    ("A5:B5",             "A5:B5             A5:B5"),
    ("B6:C9",             "B6:D9             B6:B9"),
    ("B5:B6",             "B5:B6             B5:B6"),
    ("C2:C3",             "D2:D3             #REF!"),
    ("B1:C1048576",       "B:D               B:B"),
    ("B3:C20",            "B3:D20            B3:B20"),
    ("D5:E8",             "E5:F8             C5:C8"),
    ("C5:D8",             "D5:E8             #REF!"),
    ("B2:H2",             "B2:I2             B2:F2"),
    ("E3",                "F3                C3"),
    ("E5:G5",             "F5:H5             C5:E5"),
    ("C:D",               "D:E               #REF!"),
    ("B:H",               "B:I               B:F"),
    ("D:D",               "E:E               #REF!"),
    ("A1048576",          "A1048576          A1048576"),
    ("A1048575:A1048576", "A1048575:A1048576 A1048575:A1048576"),
    ("A5:A1048576",       "A5:A1048576       A5:A1048576"),
    ("A1048576:A1048576", "A1048576:A1048576 A1048576:A1048576"),
    ("1048570:1048576",   "1048570:1048576   1048570:1048576"),
    ("A5:D5",             "A5:E5             A5:B5"),
    ("B5:D5",             "B5:E5             B5:B5"),
    ("A4:B5",             "A4:B5             A4:B5"),
    ("A5:C6",             "A5:D6             A5:B6"),
    ("XFD1",              "#REF!             XFB1"),
    ("XFA1:XFD1",         "XFB1:XFD1         XEY1:XFD1"),
    ("C1:XFD1",           "D1:XFD1           C1:XFD1"),
    ("A6:B7",             "A6:B7             A6:B7"),
    ("A5:A5",             "A5:A5             A5:A5"),
]

_SPANS_TABLE = [
    ("A5",                "A5                A5                A5                A5"),
    ("$A$5",              "$A$5              $A$5              $A$5              $A$5"),
    ("A4:A8",             "A4:A8             A4:A8             A4:A8             A4:A8"),
    ("A5:A6",             "A5:A6             A5:A6             A5:A6             A5:A6"),
    ("A3:A5",             "A3:A5             A3:A5             A3:A5             A3:A5"),
    ("A6:A9",             "A6:A9             A6:A9             A6:A9             A6:A9"),
    ("A7:A9",             "A7:A9             A7:A9             A7:A9             A7:A9"),
    ("A2:A12",            "A2:A12            A2:A12            A2:A12            A2:A12"),
    ("A:A",               "A:A               A:A               A:A               A:A"),
    ("A1:A1048576",       "A:A               A:A               A:A               A:A"),
    ("5:10",              "5:10              5:10              5:10              5:10"),
    ("5:6",               "5:6               5:6               5:6               5:6"),
    ("$5:$6",             "$5:$6             $5:$6             $5:$6             $5:$6"),
    ("4:4",               "4:4               4:4               4:4               4:4"),
    ("A1048570:A1048576", "A1048570:A1048576 A1048570:A1048576 A1048570:A1048576 A1048570:A1048576"),
    ("B1:B10",            "B1:B11            B1:B8             B1:B10            B1:B10"),
    ("A1:D10",            "A1:D10            A1:D10            A1:D10            A1:D10"),
    ("B1:E10",            "B1:E10            B1:E10            B1:E10            B1:E10"),
    ("C3",                "C3                C3                D3                #REF!"),
    ("B7",                "B8                B5                B7                B7"),
    ("D7",                "D7                D7                D7                D7"),
    ("B:B",               "B:B               B:B               B:B               B:B"),
    ("5:5",               "5:5               5:5               5:5               5:5"),
    ("B5:C5",             "B6:C6             #REF!             B5:D5             B5:B5"),
    ("B4:C6",             "B4:C7             B4:C4             B4:C6             B4:C6"),
    ("A5:B5",             "A5:B5             A5:A5             A5:B5             A5:B5"),
    ("B6:C9",             "B7:C10            B5:C7             B6:C9             B6:C9"),
    ("B5:B6",             "B6:B7             #REF!             B5:B6             B5:B6"),
    ("C2:C3",             "C2:C3             C2:C3             C2:C3             C2:C2"),
    ("B1:C1048576",       "B:C               B:C               B:C               B:C"),
    ("B3:C20",            "B3:C21            B3:C18            B3:C20            B3:C20"),
    ("D5:E8",             "D5:E8             D5:E8             D5:E8             D5:E8"),
    ("C5:D8",             "C5:D8             C5:D8             C5:D8             C6:D8"),
    ("B2:H2",             "B2:H2             B2:H2             B2:H2             B2:H2"),
    ("E3",                "E3                E3                F3                C3"),
    ("E5:G5",             "E5:G5             E5:G5             F5:H5             C5:E5"),
    ("C:D",               "C:D               C:D               C:D               C:D"),
    ("B:H",               "B:H               B:H               B:H               B:H"),
    ("D:D",               "D:D               D:D               D:D               D:D"),
    ("A1048576",          "A1048576          A1048576          A1048576          A1048576"),
    ("A1048575:A1048576", "A1048575:A1048576 A1048575:A1048576 A1048575:A1048576 A1048575:A1048576"),
    ("A5:A1048576",       "A5:A1048576       A5:A1048576       A5:A1048576       A5:A1048576"),
    ("A1048576:A1048576", "A1048576:A1048576 A1048576:A1048576 A1048576:A1048576 A1048576:A1048576"),
    ("1048570:1048576",   "1048570:1048576   1048570:1048576   1048570:1048576   1048570:1048576"),
    ("A5:D5",             "A5:D5             A5:D5             A5:E5             A5:B5"),
    ("B5:D5",             "B5:D5             D5:D5             B5:E5             B5:B5"),
    ("A4:B5",             "A4:B5             A4:B5             A4:B5             A4:B5"),
    ("A5:C6",             "A5:C6             A5:A6             A5:C6             A5:C6"),
    ("XFD1",              "XFD1              XFD1              XFD1              XFD1"),
    ("XFA1:XFD1",         "XFA1:XFD1         XFA1:XFD1         XFA1:XFD1         XFA1:XFD1"),
    ("C1:XFD1",           "C1:XFD1           C1:XFD1           C1:XFD1           C1:XFD1"),
    ("A6:B7",             "A6:B7             A6:B7             A6:B7             A6:B7"),
    ("A5:A5",             "A5:A5             A5:A5             A5:A5             A5:A5"),
]


@pytest.mark.parametrize(("ref", "results"), _WHOLE_ROWS_TABLE)
def test_whole_row_inserts_and_deletes_match_excel(ref: str, results: str):
    _check(ref, results, WHOLE_ROWS)


@pytest.mark.parametrize(("ref", "results"), _WHOLE_COLS_TABLE)
def test_whole_column_inserts_and_deletes_match_excel(ref: str, results: str):
    _check(ref, results, WHOLE_COLS)


@pytest.mark.parametrize(("ref", "results"), _SPANS_TABLE)
def test_span_restricted_shifts_match_excel(ref: str, results: str):
    _check(ref, results, SPANS)


@pytest.mark.parametrize(
    ("lo", "hi", "at", "delta", "sticky", "expected"),
    [
        (5, 9, 3, 2, True, (7, 11)),
        (5, 9, 7, 2, True, (5, 11)),
        (5, 9, 10, 2, True, (5, 9)),
        (5, 9, 5, -5, True, None),
        (5, 9, 4, -2, True, (4, 7)),
        (5, 9, 7, -5, True, (5, 6)),
        (5, 10, 6, -2, True, (5, 8)),
        (8, 10, 3, -3, True, (5, 7)),
        (5, 10, 1, -3, True, (2, 7)),
        (10, 10, 1, -3, False, (7, 7)),
        (8, 10, 1, -3, False, (5, 7)),
        (8, 10, 1, -3, True, (5, 7)),
        (8, 20, 1, -3, True, (5, 20)),  # ends on the last line: keeps ending there
        (20, 20, 1, -3, True, (17, 20)),
        (20, 20, 1, -3, False, (17, 17)),
        (19, 20, 19, -2, True, None),
        (18, 20, 19, 3, True, (18, 20)),
        (19, 20, 18, 3, True, None),
    ],
)
def test_shift_bounds(lo, hi, at, delta, sticky, expected):
    assert shift_bounds(lo, hi, at, delta, 20, sticky=sticky) == expected


def test_a_cell_moves_off_the_last_row_but_an_area_reaching_it_stays():
    assert str(CellRange.parse("B1048576").shift(True, 5, -2)) == "B1048574"
    assert str(CellRange.parse("B1048576:B1048576").shift(True, 5, -2)) == "B1048574:B1048576"
    assert str(CellRange.parse("XFD3").shift(False, 2, -1)) == "XFC3"
    assert str(CellRange.parse("XFA3:XFD3").shift(False, 2, -1)) == "XEZ3:XFD3"


def test_a_shift_keeps_dollars_and_an_unchanged_area_keeps_its_text():
    area = CellRange.parse("$b$7:c$9")
    assert area.shift(True, 1, -2) is not area
    assert str(area.shift(True, 1, -2)) == "$B$5:C$7"
    assert area.shift(True, 10, -2) is area
    assert area.shift(True, 1, -2, (4, 6)) is area  # outside the span


def test_a_span_delete_trims_an_area_inside_the_deleted_rows():
    # delete cells B5:C6, shifting up: the cells of the area in B:C go
    assert str(CellRange.parse("A5:C6").shift(True, 5, -2, (2, 3))) == "A5:A6"
    assert str(CellRange.parse("5:5").shift(True, 5, -2, (1, 3))) == "D5:XFD5"
    assert CellRange.parse("A5:D5").shift(True, 5, -2, (2, 3)) == CellRange.parse("A5:D5")
    assert CellRange.parse("B5:C5").shift(True, 5, -2, (2, 3)) is None


@pytest.mark.parametrize(
    ("text", "cols", "rows", "expected"),
    [
        ("B2", 1, 3, "C5"),
        ("$B2", 1, 3, "$B5"),
        ("B$2", 1, 3, "C$2"),
        ("$B$2", 1, 3, "$B$2"),
        ("A1:B2", 2, 2, "C3:D4"),
        ("$A1:B$2", 2, 2, "$A$2:D3"),  # bounds cross: each axis is re-ordered
        ("A$1:A5", 0, -2, "A$1:A3"),
        ("A$1:A5", 0, -5, None),
        ("B:C", 1, 5, "C:D"),
        ("$B:C", 1, 5, "$B:D"),
        ("3:4", 9, -1, "2:3"),
        ("A1", -1, 0, None),
        ("A1", 0, -1, None),
        ("XFD1048576", 1, 0, None),
        ("A:A", -1, 0, None),
    ],
)
def test_translate_moves_relative_bounds_as_a_copy_does(text, cols, rows, expected):
    moved = CellRange.parse(text).translate(cols, rows)
    assert (None if moved is None else str(moved)) == expected


def test_translate_reorders_bounds_that_cross():
    moved = CellRange.parse("A$3:A5").translate(0, -4)
    assert (moved.min_row, moved.max_row) == (1, 3)
    assert str(moved) == "A1:A$3"


def test_top_left_is_the_corner_of_the_box_around_every_area():
    assert SheetRangeSet.parse("J5:J8 H10:H12").top_left == (8, 5)
    assert SheetRangeSet.parse("C:C B7").top_left == (2, 1)
    assert SheetRangeSet.parse("B2").top_left == (2, 2)
    assert SheetRangeSet(()).top_left is None


def test_a_range_set_drops_deleted_areas_and_keeps_its_text_when_nothing_moves():
    ranges = SheetRangeSet.parse("A1:A2  $C$5:$D$6 E9")
    assert ranges.shift(True, 20, -2) is ranges
    assert str(ranges.shift(True, 20, -2)) == "A1:A2  $C$5:$D$6 E9"
    assert str(ranges.shift(True, 5, -2)) == "A1:A2 E7"
    assert str(ranges.shift(True, 5, -2, (5, 5))) == "A1:A2 $C$5:$D$6 E7"
    assert not SheetRangeSet.parse("C5:D6").shift(True, 5, -2)
