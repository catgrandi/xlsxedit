"""Values Excel cannot hold: NaN writes an empty cell, and infinity is refused.

Excel repairs a workbook whose cell stores ``<v>nan</v>`` or ``<v>inf</v>``
(recorded with Excel 16 on 2026-11-04: it does not open such a file without
repairing it). The suite guard checks every saved workbook for such numbers
(``cell-value`` in ``tests/preservation.py``).
"""

from __future__ import annotations

import io

import pytest

from xlsxedit import Workbook
from tests.preservation import assert_preserved, read_pkg
from tests.schema_validate import assert_valid_package

NAN = float("nan")
INF = float("inf")


def _sheet():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.write_rows([[1, 2, 3], ["a", "b", "c"]], at_row=1)
    return wb, ws


def _reopened(wb):
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return Workbook.open(buf)["Sheet1"]


@pytest.mark.parametrize("nan", [NAN, -NAN], ids=["nan", "negative nan"])
def test_nan_clears_a_cell_as_none_does(nan: float):
    wb, ws = _sheet()
    ws["A1"].value = nan
    ws["D1"].value = nan  # a cell the sheet does not have: none is created
    assert ws["A1"].value is None
    assert ws._find_cell_element("D1") is None
    assert _reopened(wb)["A1"].value is None
    assert_valid_package(wb)


def test_numpy_nan_writes_an_empty_cell():
    np = pytest.importorskip("numpy")
    wb, ws = _sheet()
    ws["B1"].value = np.float64("nan")
    ws.write_rows([[np.float64("nan"), np.float64(4.5)]], at_cell="A3")
    assert (ws["B1"].value, ws["A3"].value, ws["B3"].value) == (None, None, 4.5)


def test_bulk_writes_write_nan_as_an_empty_cell():
    wb, ws = _sheet()
    assert ws.write_rows([[NAN, 5, NAN], NAN], at_row=1) == 2
    assert [ws[a].value for a in ("A1", "B1", "C1", "A2", "B2")] == [None, 5, None, None, "b"]
    ws.insert_rows([[7, NAN]], at_row=1)
    assert (ws["A1"].value, ws["B1"].value) == (7, None)
    ws.insert_columns([[NAN, 8]], at_col="A")
    assert (ws["A1"].value, ws["A2"].value) == (None, 8)
    assert_valid_package(wb)


def test_nan_in_a_string_column_writes_an_empty_cell_and_infinity_text():
    wb, ws = _sheet()
    ws.write_rows([[NAN, INF, -INF]], at_row=5, string_columns={0, 1, 2})
    assert [ws[a].value for a in ("A5", "B5", "C5")] == [None, "inf", "-inf"]


def test_nan_aimed_at_a_merged_cell_is_not_sent_to_the_merge_s_top_left_cell():
    wb, ws = _sheet()
    ws.merge_cells("A1:B1")
    ws.write_rows([[NAN]], at_cell="B1")
    assert ws["A1"].value == 1
    ws.write_rows([[5]], at_cell="B1")
    assert ws["A1"].value == 5


def test_write_dataframe_writes_nan_as_an_empty_cell():
    pd = pytest.importorskip("pandas")
    np = pytest.importorskip("numpy")
    wb, ws = _sheet()
    frame = pd.DataFrame({"x": [1.5, np.nan], "y": ["a", np.nan]})
    assert wb.write_dataframe(frame, at_cell="A4") == 2
    assert [ws[a].value for a in ("A4", "B4", "A5", "B5", "A6", "B6")] == ["x", "y", 1.5, "a", None, None]
    assert_valid_package(wb)


INFINITE_WRITES = {
    "Cell.value": lambda ws, v: setattr(ws["A1"], "value", v),
    "Cell.value on a new cell": lambda ws, v: setattr(ws["F9"], "value", v),
    "write_rows, in a later row": lambda ws, v: ws.write_rows([[1, 2], [3, v]], at_row=1),
    "insert_rows": lambda ws, v: ws.insert_rows([[1, v]], at_row=1),
    "insert_columns": lambda ws, v: ws.insert_columns([[1, v]], at_col="B"),
    "write_dataframe": lambda ws, v: ws._workbook.write_dataframe(
        [(1.0,), (v,)], at_cell="A1", header=False, clear_range="A1:C2"
    ),
}


@pytest.mark.parametrize("value", [INF, -INF], ids=["inf", "-inf"])
@pytest.mark.parametrize("case", INFINITE_WRITES)
def test_infinity_is_refused_before_anything_changes(case: str, value: float):
    wb, ws = _sheet()
    before = read_pkg(wb)
    with pytest.raises(ValueError, match="Excel has no infinite numbers"):
        INFINITE_WRITES[case](ws, value)
    assert_preserved(before, read_pkg(wb))
