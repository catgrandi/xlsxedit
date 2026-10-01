"""``CellRange`` and ``SheetRangeSet``: every ``ST_Ref`` form, shifting, clamping."""

from __future__ import annotations

import pytest

from xlsxedit.exceptions import InvalidRangeError
from xlsxedit.range_set import (
    CELL,
    COLUMNS,
    RANGE,
    ROWS,
    CellRange,
    SheetRangeSet,
    column_number,
)


@pytest.mark.parametrize(
    ("text", "bounds", "kind"),
    [
        ("B2", (2, 2, 2, 2), CELL),
        ("$B$2", (2, 2, 2, 2), CELL),
        ("b2", (2, 2, 2, 2), CELL),
        ("A1:C3", (1, 1, 3, 3), RANGE),
        ("$A$10:$C$20", (1, 10, 3, 20), RANGE),
        ("A$1:$C3", (1, 1, 3, 3), RANGE),
        ("C3:A1", (1, 1, 3, 3), RANGE),
        ("A5:A5", (1, 5, 1, 5), RANGE),
        ("A:A", (1, 1, 1, 1_048_576), COLUMNS),
        ("$A:$C", (1, 1, 3, 1_048_576), COLUMNS),
        ("10:10", (1, 10, 16_384, 10), ROWS),
        ("$1:$3", (1, 1, 16_384, 3), ROWS),
        ("A1:A1048576", (1, 1, 1, 1_048_576), RANGE),
        ("XFD1048576", (16_384, 1_048_576, 16_384, 1_048_576), CELL),
    ],
)
def test_parse_every_st_ref_form(text: str, bounds: tuple[int, int, int, int], kind: str):
    area = CellRange.parse(text)
    assert (area.min_col, area.min_row, area.max_col, area.max_row) == bounds
    assert area.kind == kind
    assert str(area) == text


@pytest.mark.parametrize(
    "text",
    ["", "A", "1", "A1:B", "A:1", "A1:B2:C3", "XFE1", "A0", "A1048577", "$$A1", "A1 B2", "#REF!"],
)
def test_parse_rejects_malformed_and_off_grid_refs(text: str):
    with pytest.raises(InvalidRangeError):
        CellRange.parse(text)


@pytest.mark.parametrize(
    ("text", "at", "expected"),
    [
        ("A1:A10", 1, "A3:A12"),  # at the first row: moves
        ("A1:A10", 10, "A1:A12"),  # at the last row: grows
        ("A1:A10", 11, "A1:A10"),  # below: unchanged
        ("A1:A10", 5, "A1:A12"),  # straddling: grows
        ("$A$10:$C$20", 5, "$A$12:$C$22"),
        ("A$1:$C3", 2, "A$1:$C5"),
        ("B7", 7, "B9"),
        ("B7", 8, "B7"),
        ("3:5", 3, "5:7"),
        ("$3:$5", 4, "$3:$7"),
        ("A:C", 1, "A:C"),  # whole columns ignore row insertion
        ("A1:A1048576", 1, "A1:A1048576"),  # so does Excel's whole-column idiom
        ("A5:A1048576", 3, "A7:A1048576"),  # clamped at the last row
        ("A1048574:B1048576", 1, "A1048576:B1048576"),
    ],
)
def test_shift_rows(text: str, at: int, expected: str):
    assert str(CellRange.parse(text).shift_rows(at, 2)) == expected


@pytest.mark.parametrize(
    ("text", "at", "expected"),
    [
        ("A1:C1", 2, "A1:E1"),
        ("C1", 2, "E1"),
        ("A1:B1", 3, "A1:B1"),
        ("$B:$D", 1, "$D:$F"),
        ("5:5", 1, "5:5"),  # whole rows ignore column insertion
        ("A5:XFD5", 2, "A5:XFD5"),
        ("XFA1:XFD1", 1, "XFC1:XFD1"),
    ],
)
def test_shift_cols(text: str, at: int, expected: str):
    assert str(CellRange.parse(text).shift_cols(at, 2)) == expected


def test_areas_pushed_off_the_grid_are_dropped():
    assert CellRange.parse("A1048576").shift_rows(1, 1) is None
    assert CellRange.parse("XFD1").shift_cols(1, 1) is None
    assert CellRange.parse("1048576:1048576").shift_rows(5, 1) is None


def test_an_unchanged_area_keeps_its_text():
    area = CellRange.parse("$b$2:c3")
    assert area.shift_rows(9, 1) is area
    assert str(area.shift_rows(9, 1)) == "$b$2:c3"
    assert str(area.shift_rows(1, 1)) == "$B$3:C4"


def test_extend_to_row():
    assert str(CellRange.parse("C2").extend_to_row(10)) == "C2:C10"
    assert str(CellRange.parse("$C$2:$C$5").extend_to_row(9)) == "$C$2:$C$9"
    assert str(CellRange.parse("C2:C9").extend_to_row(5)) == "C2:C9"
    assert str(CellRange.parse("C2").extend_to_row(2_000_000)) == "C2:C1048576"
    assert str(CellRange.parse("A:A").extend_to_row(5)) == "A:A"


def test_intersects_and_contains():
    a = CellRange.parse("B2:D4")
    assert a.intersects(CellRange.parse("D4:E9"))
    assert not a.intersects(CellRange.parse("E1:E9"))
    assert a.intersects(CellRange.parse("C:C"))
    assert a.intersects(CellRange.parse("3:3"))
    assert a.contains(3, 3) and not a.contains(5, 3)


def test_from_bounds():
    assert str(CellRange.from_bounds(1, 2, 1, 2)) == "A2"
    assert str(CellRange.from_bounds(1, 2, 3, 4)) == "A2:C4"
    with pytest.raises(InvalidRangeError):
        CellRange.from_bounds(1, 1, 16_385, 1)


def test_range_set_parses_and_keeps_the_original_text():
    text = "A1:A1048576  $C$4:$D$9 E:E 7:8"
    ranges = SheetRangeSet.parse(text)
    assert len(ranges) == 4
    assert ranges.shift_rows(20, 1) is ranges
    assert str(ranges.shift_rows(20, 1)) == text


def test_range_set_shifts_each_area_and_drops_those_pushed_off():
    ranges = SheetRangeSet.parse("A1:A1048576 $C$4:$D$9 E:E 7:8 B1048576")
    assert str(ranges.shift_rows(5, 2)) == "A1:A1048576 $C$4:$D$11 E:E 9:10"
    assert str(ranges.shift_cols(1, 1)) == "B1:B1048576 $D$4:$E$9 F:F 7:8 C1048576"
    assert not SheetRangeSet.parse("A1048576").shift_rows(1, 1)


def test_range_set_rejects_any_malformed_area():
    with pytest.raises(InvalidRangeError, match="A1:B"):
        SheetRangeSet.parse("C3 A1:B")


def test_range_set_intersects():
    ranges = SheetRangeSet.parse("A1 C3:D4")
    assert ranges.intersects("D4")
    assert ranges.intersects(CellRange.parse("B:C"))
    assert not ranges.intersects(SheetRangeSet.parse("B1 E5"))


@pytest.mark.parametrize(
    ("text", "bounds", "rendered_after_shift"),
    [
        ("A01", (1, 1, 1, 1), "A3"),
        ("A0001:B02", (1, 1, 2, 2), "A3:B4"),
        ("xfd:a", (1, 1, 16_384, 1_048_576), "xfd:a"),  # whole columns: unchanged
        ("1048576:1", (1, 1, 16_384, 1_048_576), "1048576:1"),  # every row: unchanged
        ("AAA1", (703, 1, 703, 1), "AAA3"),
        ("$C3:A$1", (1, 1, 3, 3), "A$3:$C5"),
    ],
)
def test_more_st_ref_forms(text: str, bounds: tuple[int, int, int, int], rendered_after_shift: str):
    area = CellRange.parse(text)
    assert (area.min_col, area.min_row, area.max_col, area.max_row) == bounds
    assert str(area) == text
    assert str(area.shift_rows(1, 2)) == rendered_after_shift


@pytest.mark.parametrize("text", ["XFE:XFE", "1:1048577", "0:1", "A1:XFE1"])
def test_off_grid_whole_forms_are_rejected(text: str):
    with pytest.raises(InvalidRangeError):
        CellRange.parse(text)


def test_column_number():
    assert column_number("A") == 1
    assert column_number("xfd") == 16_384
    for bad in ("", "B1", "B ", "$B", "XFE", "ABCD"):
        with pytest.raises(InvalidRangeError):
            column_number(bad)


def test_edges_and_corners_intersect():
    a = CellRange.parse("B2:D4")
    assert a.intersects(CellRange.parse("D4"))
    assert a.intersects(CellRange.parse("E1:E9")) is False
    assert a.intersects(CellRange.parse("A1:B2"))


def test_range_set_splits_on_xml_whitespace_only():
    assert len(SheetRangeSet.parse("A1\tB2\r\nC3")) == 3
    with pytest.raises(InvalidRangeError):
        SheetRangeSet.parse("B1\u00a0C3")


def test_range_set_extend_to_row():
    ranges = SheetRangeSet.parse("C2 $D$2:$D$5 A:A")
    assert str(ranges.extend_to_row(9)) == "C2:C9 $D$2:$D$9 A:A"
    assert ranges.extend_to_row(1) is ranges
    assert str(CellRange.parse("3:5").extend_to_row(10)) == "3:10"
