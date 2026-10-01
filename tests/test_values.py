"""Unit tests for Worksheet.values (no pandas required)."""

from __future__ import annotations

from xlsxedit import Workbook


def test_values_dense_and_empty_as_blank():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "h"
    ws["B1"].value = 1
    ws["A2"].value = "r"
    grid = ws.values()
    assert grid == [["h", 1], ["r", ""]]


def test_values_sparse_padding_and_row_gaps():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["C3"].value = "c3"
    ws["E5"].value = "e5"
    grid = ws.values()
    # Excel rows 3,4,5 → three output rows; cols A..E padded
    assert len(grid) == 3
    assert grid[0] == ["", "", "c3", "", ""]
    assert grid[1] == ["", "", "", "", ""]
    assert grid[2] == ["", "", "", "", "e5"]


def test_values_max_rows():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "a"
    ws["A2"].value = "b"
    ws["A3"].value = "c"
    assert ws.values(max_rows=2) == [["a"], ["b"]]


def test_values_formula_cache():
    from xlsxedit.cell import _ensure_v

    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = 1
    ws["B1"].formula = "=A1+1"
    _ensure_v(ws["B1"]._element).text = "42"
    assert ws["B1"].has_formula
    assert ws.values() == [[1, 42]]

    # Uncached formula → ""; keep a trailing value so the empty is not trimmed.
    ws["B1"].formula = "=A1*2"
    # clear cached <v> by rewriting formula on a fresh middle cell
    ws["B1"].clear()
    ws["B1"].formula = "=A1*2"
    ws["C1"].value = "x"
    grid = ws.values()
    assert grid[0] == [1, "", "x"]


def test_values_empty_sheet():
    wb = Workbook.create()
    assert wb["Sheet1"].values() == []


def _set_formula_error(cell, formula: str, error: str) -> None:
    """Plant a formula cell with Excel error cache (``t=\"e\"``)."""
    from xlsxedit.cell import _ensure_v

    cell.formula = formula
    cell._element.set("t", "e")
    _ensure_v(cell._element).text = error


def test_formula_error_ref_in_value_and_grid():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "ok"
    _set_formula_error(ws["B1"], "=Z99", "#REF!")
    assert ws["B1"].value == "#REF!"
    assert ws.values() == [["ok", "#REF!"]]


def test_formula_error_value_and_na():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    _set_formula_error(ws["A1"], "=1/0", "#DIV/0!")
    _set_formula_error(ws["B1"], "=NA()", "#N/A")
    assert ws["A1"].value == "#DIV/0!"
    assert ws["B1"].value == "#N/A"
    assert ws.values() == [["#DIV/0!", "#N/A"]]


def test_formula_cache_scientific_and_float():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    from xlsxedit.cell import _ensure_v

    ws["A1"].value = 0
    ws["A1"].formula = "=1E3"
    _ensure_v(ws["A1"]._element).text = "1E3"
    ws["B1"].value = 0
    ws["B1"].formula = "=1.5"
    _ensure_v(ws["B1"]._element).text = "1.5"
    assert ws["A1"].value == 1000.0
    assert ws["B1"].value == 1.5
    assert ws.values() == [[1000.0, 1.5]]


def test_non_formula_junk_v_does_not_raise():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    from xlsxedit.cell import _ensure_v

    # Numeric-typed cell (no t) with non-numeric <v> must not raise.
    cell = ws["A1"]
    _ensure_v(cell._element).text = "#REF!"
    assert cell.value == "#REF!"


def test_errors_as_nan_flag():
    import math

    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "ok"
    _set_formula_error(ws["B1"], "=Z99", "#REF!")

    assert ws.values() == [["ok", "#REF!"]]
    grid = ws.values(errors_as_nan=True)
    assert grid[0][0] == "ok"
    assert isinstance(grid[0][1], float) and math.isnan(grid[0][1])


def test_literal_hash_ref_string_not_nan():
    """Shared-string ``#REF!`` is not an Excel error type — stay a string."""
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "#REF!"
    assert ws.values(errors_as_nan=True) == [["#REF!"]]


def test_merge_non_anchor_leftover_blanked():
    from lxml import etree

    from xlsxedit.cell import _ensure_v
    from xlsxedit.opc.constants import SML_NS

    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "v"
    ws.merge_cells("A1:A3")
    ws["B1"].value = 1
    ws["B2"].value = 2
    ws["B3"].value = 3

    # Plant a leftover value on non-anchor A2 (bypass merge write redirect).
    row2 = ws._ensure_row(2)
    c = etree.SubElement(row2, f"{{{SML_NS}}}c")
    c.set("r", "A2")
    row2.insert(0, c)  # ahead of B2, in column order
    _ensure_v(c).text = "999"

    assert ws.values() == [["v", 1], ["", 2], ["", 3]]
