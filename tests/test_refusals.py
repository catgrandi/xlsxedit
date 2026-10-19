"""A refused cell or bulk write changes no package member."""

from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

import pytest

from xlsxedit import FormulaGroupError, GridOverflowError, InvalidRangeError, Workbook
from xlsxedit.bulk import write_cell_value
from tests.conftest import INSPECT_FIXTURES
from tests.preservation import assert_preserved, read_pkg
from tests.test_formula_hygiene import RANGE_FORMULAS, _plant

SIMPLE = INSPECT_FIXTURES["SimpleFormula"]
NUL = chr(0)

BAD_VALUES = {
    "unsupported type": (object(), TypeError),
    "text XML cannot hold": ("a" + NUL, ValueError),
    "aware datetime": (datetime(2024, 1, 1, tzinfo=timezone.utc), TypeError),
}


def _simple():
    """SimpleFormula, whose sheet has a calcChain, a normal formula in C3 and a
    shared-formula master in C4 that C5:C7 derive from."""
    wb = Workbook.open(SIMPLE)
    return wb, wb.worksheets[0]


# --- Cell.value and bulk.write_cell_value ---------------------------------------------


@pytest.mark.parametrize("bad", BAD_VALUES)
@pytest.mark.parametrize("address", ["C3", "C4"])
def test_refused_cell_value_on_a_formula_cell_changes_nothing(address: str, bad: str):
    value, error = BAD_VALUES[bad]
    wb, ws = _simple()
    before = read_pkg(wb)
    with pytest.raises(error):
        ws[address].value = value
    assert_preserved(before, read_pkg(wb))


@pytest.mark.parametrize("bad", BAD_VALUES)
def test_refused_bulk_cell_value_on_a_formula_cell_changes_nothing(bad: str):
    value, error = BAD_VALUES[bad]
    wb, ws = _simple()
    before = read_pkg(wb)
    with pytest.raises(error):
        write_cell_value(ws._find_cell_element("C4"), value, ws.shared_strings, {}, worksheet=ws)
    assert_preserved(before, read_pkg(wb))


@pytest.mark.parametrize("write", ["cell", "bulk"])
def test_string_without_a_shared_strings_part_changes_nothing(write: str):
    wb, ws = _simple()
    wb._shared_strings = None
    before = read_pkg(wb)
    with pytest.raises(ValueError, match="no shared strings part"):
        if write == "cell":
            ws["C4"].value = "text"
        else:
            write_cell_value(ws._find_cell_element("C4"), "text", None, {}, worksheet=ws)
    assert_preserved(before, read_pkg(wb))


@pytest.mark.parametrize("index", ["-1", "99", "x"])
def test_string_over_a_corrupt_shared_string_index_adds_a_string(index: str):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "kept"
    _plant(ws["B1"], {"t": "s"}, f"<v>{index}</v>")

    ws["B1"].value = "new"
    assert (ws["A1"].value, ws["B1"].value) == ("kept", "new")


@pytest.mark.parametrize("index", ["-1", "99", "x"])
def test_add_table_writes_a_header_over_a_corrupt_shared_string_index(index: str):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "old"
    _plant(ws["B1"], {"t": "s"}, f"<v>{index}</v>")

    ws.add_table("A1:B2", ["Name", "Qty"], write_header=True)
    assert (ws["A1"].value, ws["B1"].value) == ("Name", "Qty")


# --- Bulk overwrites ---------------------------------------------------------------------


def _array_sheet():
    """SimpleFormula with an array formula over E3:E4 and a merge E3:F3."""
    wb, ws = _simple()
    _plant(ws["E3"], {}, RANGE_FORMULAS["array"].replace("B1:B2", "E3:E4") + "<v>2</v>")
    _plant(ws["E4"], {}, "<v>4</v>")
    ws.merge_cells("E3:F3")
    return wb, ws


BULK_REFUSALS = {
    "anchor after written cells": (
        lambda ws: ws.write_rows([["x", "y"], ["z", 0]], at_cell="D2"),
        FormulaGroupError,
    ),
    "None over the anchor": (lambda ws: ws.write_rows([["x", None]], at_cell="D3"), FormulaGroupError),
    "styled anchor": (
        lambda ws: ws.write_rows([[0]], at_cell="E3", row_styles=[{"bold": True}]),
        FormulaGroupError,
    ),
    "value a merge sends to the anchor": (
        lambda ws: ws.write_rows([["x", "y"]], at_cell="F3"),
        FormulaGroupError,
    ),
    "unsupported value in a later row": (
        lambda ws: ws.write_rows([["a"], [object()]], at_cell="H1"),
        TypeError,
    ),
    "text XML cannot hold in a later row": (
        lambda ws: ws.write_rows([["a"], ["b" + NUL]], at_cell="H1"),
        ValueError,
    ),
    "aware datetime": (
        lambda ws: ws.write_rows([[1, datetime(2024, 1, 1, tzinfo=timezone.utc)]], at_cell="H1"),
        TypeError,
    ),
    "string without a shared strings part": (
        lambda ws: (setattr(ws._workbook, "_shared_strings", None), ws.write_rows([[1, "a"]], at_cell="H1")),
        ValueError,
    ),
    "past the last row": (lambda ws: ws.write_rows([["a"], ["b"]], at_row=1048576), GridOverflowError),
    "past column XFD": (lambda ws: ws.write_rows([["a", "b"]], at_cell="XFD1"), GridOverflowError),
    "row zero": (lambda ws: ws.write_rows([["a"]], at_row=0), InvalidRangeError),
    "write_rows_at": (
        lambda ws: ws._write_rows_at([["a"], [object()]], start_row=1, start_col_idx=7),
        TypeError,
    ),
    "write_dataframe after clear_range": (
        lambda ws: ws._workbook.write_dataframe(
            [("a",), (object(),)], at_cell="H1", header=False, clear_range="A3:B4"
        ),
        TypeError,
    ),
    "write_dataframe over the anchor": (
        lambda ws: ws._workbook.write_dataframe(
            [("x", "y"), ("z", 0)], at_cell="D2", header=False, clear_range="A3:B4"
        ),
        FormulaGroupError,
    ),
    "clear_range over the anchor": (lambda ws: ws.clear_range("A2:E3"), FormulaGroupError),
}


@pytest.mark.parametrize("case", BULK_REFUSALS)
def test_refused_bulk_write_changes_nothing(case: str):
    write, error = BULK_REFUSALS[case]
    wb, ws = _array_sheet()
    before = read_pkg(wb)
    with pytest.raises(error):
        write(ws)
    assert_preserved(before, read_pkg(wb))


def test_refused_pandas_overwrite_with_clear_range_changes_nothing(tmp_path: Path):
    pd = pytest.importorskip("pandas")
    from xlsxedit.pandas_io import register

    register()
    wb, _ws = _array_sheet()
    path = tmp_path / "template.xlsx"
    wb.save(path)
    before = read_pkg(path)

    df = pd.DataFrame({"D": ["x"], "E": [0]})
    with pytest.raises(FormulaGroupError):
        with pd.ExcelWriter(
            path,
            engine="xlsxedit",
            mode="a",
            if_sheet_exists="overlay",
            engine_kwargs={"clear_range": "A3:B4"},
        ) as writer:
            df.to_excel(writer, sheet_name="Sheet1", startrow=1, startcol=3, index=False)
    assert_preserved(before, path)


def test_bulk_write_beside_the_anchor_still_writes():
    wb, ws = _array_sheet()
    assert ws.write_rows([["x", None, None, "y"]], at_cell="D2") == 1
    assert (ws["D2"].value, ws["G2"].value) == ("x", "y")
    assert ws["E3"].formula_type == "array"
