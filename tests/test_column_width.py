"""Setting one column's width where a ``<col>`` covers several columns.

The ``<col>`` is split around the column, as Excel splits it. ``EXCEL`` records
what Excel 16.0 (build 20527, 64-bit, Windows 11 10.0.26200) wrote on
2026-11-04: each workbook was built with xlsxedit, with ``<col min="1" max="1"
width="12" customWidth="1"/>`` and one ``<col min="3" max="5" .../>``, opened in
Excel, given ``Columns(column).ColumnWidth = 20`` and saved. Excel stored the
20 as ``20.546875``, from its font metrics; xlsxedit stores the value given, so
the column set is compared without its width.
"""

from __future__ import annotations

import io

import pytest
from lxml import etree

from xlsxedit import Workbook
from xlsxedit.opc.constants import SML_NS
from tests.preservation import assert_preserved, read_pkg
from tests.schema_validate import assert_valid_package

_COLS = f"{{{SML_NS}}}cols"
_COL = f"{{{SML_NS}}}col"
SHEET_XML = "xl/worksheets/sheet1.xml"
A = {"min": "1", "max": "1", "width": "12", "customWidth": "1"}


def _col(low: int, high: int, **attrs: str) -> dict[str, str]:
    return {"min": str(low), "max": str(high), **attrs}


# case -> (attributes of <col min="3" max="5">, column set, the <col>s Excel wrote),
# with "set" for the width Excel gave the column set. style="1" is the sheet's green fill.
EXCEL: dict[str, tuple[dict[str, str], str, list[dict[str, str]]]] = {
    "last column": (
        {"width": "9", "style": "1", "customWidth": "1"},
        "E",
        [
            A,
            _col(3, 4, width="9", style="1", customWidth="1"),
            _col(5, 5, width="set", style="1", customWidth="1"),
        ],
    ),
    "middle column": (
        {"width": "9", "style": "1", "customWidth": "1"},
        "D",
        [
            A,
            _col(3, 3, width="9", style="1", customWidth="1"),
            _col(4, 4, width="set", style="1", customWidth="1"),
            _col(5, 5, width="9", style="1", customWidth="1"),
        ],
    ),
    "first column": (
        {"width": "9", "style": "1", "customWidth": "1"},
        "C",
        [
            A,
            _col(3, 3, width="set", style="1", customWidth="1"),
            _col(4, 5, width="9", style="1", customWidth="1"),
        ],
    ),
    "bestFit and outlineLevel": (
        {"width": "9", "style": "1", "bestFit": "1", "customWidth": "1", "outlineLevel": "1"},
        "D",
        [
            A,
            _col(3, 3, width="9", style="1", bestFit="1", customWidth="1", outlineLevel="1"),
            _col(4, 4, width="set", style="1", customWidth="1", outlineLevel="1"),
            _col(5, 5, width="9", style="1", bestFit="1", customWidth="1", outlineLevel="1"),
        ],
    ),
    "no customWidth": (
        {"width": "9", "style": "1"},
        "D",
        [
            A,
            _col(3, 3, width="9", style="1"),
            _col(4, 4, width="set", style="1", customWidth="1"),
            _col(5, 5, width="9", style="1"),
        ],
    ),
    "no style": (
        {"width": "9", "customWidth": "1"},
        "D",
        [
            A,
            _col(3, 3, width="9", customWidth="1"),
            _col(4, 4, width="set", customWidth="1"),
            _col(5, 5, width="9", customWidth="1"),
        ],
    ),
}


def _book(attrs: dict[str, str]) -> Workbook:
    """The probe workbook: A1 set, column A 12 wide, then ``<col min="3" max="5">``."""
    wb = Workbook.create()
    ws = wb["Sheet1"]
    green = wb.styles.allocate_cell_style(bg_color="FF00FF00")
    assert green == 1
    ws["A1"].value = 1
    ws.column_dimensions["A"].width = 12
    etree.SubElement(ws._part.element.find(_COLS), _COL, min="3", max="5", **attrs)
    ws._part.mark_dirty()
    return wb


def _cols(wb: Workbook, column_set: str | None = None, width: str = "20") -> list[dict[str, str]]:
    cols = []
    for col in wb["Sheet1"]._part.element.find(_COLS).iterfind(_COL):
        attrs = dict(col.attrib)
        if column_set is not None and attrs.get("width") == width and attrs["min"] == attrs["max"]:
            attrs["width"] = "set"
        cols.append(attrs)
    return cols


@pytest.mark.parametrize("case", EXCEL)
def test_the_covering_col_is_split_as_excel_splits_it(case: str):
    attrs, column, excel = EXCEL[case]
    wb = _book(attrs)
    wb["Sheet1"].column_dimensions[column].width = 20
    assert _cols(wb, column) == excel
    assert_valid_package(wb)


def test_a_hidden_column_keeps_hidden_where_excel_unhides_it():
    """Excel dropped ``hidden`` from D and kept it on C and E. Nothing in the
    API hides a column, so xlsxedit keeps it rather than lose it for good."""
    wb = _book({"width": "9", "style": "1", "hidden": "1", "customWidth": "1"})
    wb["Sheet1"].column_dimensions["D"].width = 20
    assert _cols(wb, "D") == [
        A,
        _col(3, 3, width="9", style="1", hidden="1", customWidth="1"),
        _col(4, 4, width="set", style="1", hidden="1", customWidth="1"),  # Excel: no hidden
        _col(5, 5, width="9", style="1", hidden="1", customWidth="1"),
    ]


def test_the_neighbours_keep_their_width_and_every_column_its_style():
    wb = _book({"width": "9", "style": "1", "customWidth": "1"})
    ws = wb["Sheet1"]
    ws.column_dimensions["D"].width = 20
    back = Workbook.open(io.BytesIO(_saved(wb)))["Sheet1"]
    widths = {c: back.column_dimensions[c].width for c in "CDE"}
    assert widths == {"C": 9.0, "D": 20.0, "E": 9.0}
    for column in "CDE":
        assert back._workbook.styles.effective_xf_index(back, f"{column}7") == 1


def test_first_writes_after_the_split_take_each_columns_style():
    wb = _book({"width": "9", "style": "1", "customWidth": "1"})
    ws = wb["Sheet1"]
    ws["C2"].value = 1  # reads the styled <col> ranges into the cache
    ws.column_dimensions["D"].width = 20
    for column in "CDEF":
        ws[f"{column}3"].value = 1
    assert [ws[f"{c}3"].style_index for c in "CDEF"] == ["1", "1", "1", None]


def test_a_width_the_covering_col_already_has_changes_nothing():
    wb = _book({"width": "9", "style": "1", "customWidth": "1"})
    saved = _saved(wb)
    wb = Workbook.open(io.BytesIO(saved))
    wb["Sheet1"].column_dimensions["D"].width = 9
    assert not wb["Sheet1"]._part.is_dirty
    report = assert_preserved(saved, wb)
    assert not report.reserialised


def test_a_single_column_col_drops_bestfit_when_its_width_changes():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.column_dimensions["E"].width = 9
    ws._part.element.find(_COLS).find(_COL).set("bestFit", "1")
    ws.column_dimensions["E"].width = 9  # the width it has: nothing changes
    assert _cols(wb) == [_col(5, 5, width="9", customWidth="1", bestFit="1")]
    ws.column_dimensions["E"].width = 20
    assert _cols(wb) == [_col(5, 5, width="20", customWidth="1")]


def test_only_the_first_of_overlapping_cols_keeps_the_column():
    wb = _book({"width": "9", "style": "1", "customWidth": "1"})
    ws = wb["Sheet1"]
    etree.SubElement(ws._part.element.find(_COLS), _COL, min="4", max="6", width="30")
    ws.column_dimensions["D"].width = 20
    assert [c for c in _cols(wb) if int(c["min"]) <= 4 <= int(c["max"])] == [
        _col(4, 4, width="20", style="1", customWidth="1")
    ]
    assert _cols(wb)[-1] == _col(5, 6, width="30")


def _saved(wb: Workbook) -> bytes:
    pkg = read_pkg(wb)
    buf = io.BytesIO()
    wb.save(buf)
    assert read_pkg(buf.getvalue()) == pkg
    return buf.getvalue()
