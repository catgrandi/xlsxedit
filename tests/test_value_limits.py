"""Values Excel cannot hold: NaN, and pandas' NaT and NA, write an empty cell, while infinity and text over
32,767 characters are refused.

Excel repairs a workbook whose cell stores ``<v>nan</v>`` or ``<v>inf</v>``, or a
text longer than 32,767 UTF-16 code units: 16,383 emoji and a letter open, 16,384
emoji do not (recorded with Excel 16 on 2026-11-04: it does not open such a file
without repairing it). The suite guard checks every saved workbook for such
values (``cell-value`` in ``tests/preservation.py``).
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


MISSING_WRITES = {  # each writes a missing value into A1 and 5 into the second cell named
    "Cell.value": (lambda ws, v: setattr(ws["A1"], "value", v), None),
    "write_rows": (lambda ws, v: ws.write_rows([[v, 5]], at_row=1), "B1"),
    "write_rows, string column": (lambda ws, v: ws.write_rows([[v, 5]], at_row=1, string_columns={0}), "B1"),
    "insert_rows": (lambda ws, v: ws.insert_rows([[v, 5]], at_row=1), "B1"),
    "insert_columns": (lambda ws, v: ws.insert_columns([[v, 5]], at_col="A"), "A2"),
}


@pytest.mark.parametrize("write", list(MISSING_WRITES))
@pytest.mark.parametrize("missing", ["NaT", "NA"])
def test_pandas_nat_and_na_write_an_empty_cell_as_nan_does(missing: str, write: str):
    pd = pytest.importorskip("pandas")
    wb, ws = _sheet()
    action, five = MISSING_WRITES[write]
    action(ws, getattr(pd, missing))
    assert ws["A1"].value is None
    if five is not None:
        assert ws[five].value == 5
    ws["F9"].value = getattr(pd, missing)  # a cell the sheet does not have: none is created
    assert ws._find_cell_element("F9") is None
    assert_valid_package(wb)


def test_write_dataframe_writes_nat_and_na_as_empty_cells():
    pd = pytest.importorskip("pandas")
    wb, ws = _sheet()
    frame = pd.DataFrame(
        {
            "d": pd.to_datetime(["2024-01-01", None]),
            "f": pd.array([1.5, None], dtype="Float64"),
            "s": pd.array(["a", None], dtype="string"),
        }
    )
    assert wb.write_dataframe(frame, at_cell="A4") == 2
    assert [ws[a].value for a in ("B5", "C5")] == [1.5, "a"]
    assert ws["A5"].value is not None
    assert [ws[a].value for a in ("A6", "B6", "C6")] == [None, None, None]
    assert_valid_package(wb)


def test_write_dataframe_writes_nullable_integer_and_boolean_columns():
    pd = pytest.importorskip("pandas")
    wb, ws = _sheet()
    frame = pd.DataFrame({"i": pd.array([1, None], dtype="Int64"), "b": pd.array([True, None], dtype="boolean")})
    assert wb.write_dataframe(frame, at_cell="A4") == 2
    assert [ws[a].value for a in ("A5", "B5", "A6", "B6")] == [1, True, None, None]
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


@pytest.mark.parametrize("value", [10**400, -(10**309)], ids=["10**400", "-10**309"])
@pytest.mark.parametrize("case", INFINITE_WRITES)
def test_an_integer_past_the_largest_double_is_refused(case: str, value: int):
    """Excel stores numbers as doubles; such an integer reads as infinity."""
    wb, ws = _sheet()
    before = read_pkg(wb)
    with pytest.raises(ValueError, match="largest number Excel holds"):
        INFINITE_WRITES[case](ws, value)
    assert_preserved(before, read_pkg(wb))


def test_the_largest_integer_a_double_holds_is_written():
    wb, ws = _sheet()
    largest = int(1.7976931348623157e308)
    ws["A1"].value = largest
    ws.write_rows([[largest]], at_cell="B5")
    ws.write_rows([[10**400]], at_cell="C5", string_columns={2})  # as text
    reopened = _reopened(wb)
    assert reopened["A1"].value == reopened["B5"].value == largest
    assert reopened["C5"].value == str(10**400)


EMOJI = "\N{GRINNING FACE}"
LONGEST = {"32767 letters": "a" * 32_767, "16383 emoji and a letter": EMOJI * 16_383 + "a"}
TOO_LONG = {"32768 letters": "a" * 32_768, "16384 emoji": EMOJI * 16_384}


@pytest.mark.parametrize("text", LONGEST.values(), ids=LONGEST)
def test_text_of_32767_utf16_code_units_is_written(text: str):
    wb, ws = _sheet()
    ws["A1"].value = text
    ws.write_rows([[text]], at_cell="A3")
    reopened = _reopened(wb)
    assert reopened["A1"].value == reopened["A3"].value == text


TOO_LONG_WRITES = {
    "Cell.value": lambda ws, t: setattr(ws["A1"], "value", t),
    "Cell.value on a new cell": lambda ws, t: setattr(ws["F9"], "value", t),
    "write_rows, in a later row": lambda ws, t: ws.write_rows([["x"], [t]], at_row=1),
    "write_rows, a string column": lambda ws, t: ws.write_rows(
        [[1], [_Text(t)]], at_row=1, string_columns={0}
    ),
    "insert_rows": lambda ws, t: ws.insert_rows([["x", t]], at_row=1),
    "insert_columns": lambda ws, t: ws.insert_columns([["x", t]], at_col="B"),
    "write_dataframe": lambda ws, t: ws._workbook.write_dataframe(
        [("x",), (t,)], at_cell="A1", header=False, clear_range="A1:C2"
    ),
}


class _Text:
    """A value a string column writes as its ``str()``."""

    def __init__(self, text: str):
        self._text = text

    def __str__(self) -> str:
        return self._text


@pytest.mark.parametrize("text", TOO_LONG.values(), ids=TOO_LONG)
@pytest.mark.parametrize("case", TOO_LONG_WRITES)
def test_text_over_32767_utf16_code_units_is_refused_before_anything_changes(case: str, text: str):
    wb, ws = _sheet()
    before = read_pkg(wb)
    with pytest.raises(ValueError, match="32,768 characters is longer than the 32,767"):
        TOO_LONG_WRITES[case](ws, text)
    assert_preserved(before, read_pkg(wb))
