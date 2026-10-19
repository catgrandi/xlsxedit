"""A refused cell write changes no package member."""

from __future__ import annotations

from datetime import datetime, timezone

import pytest

from xlsxedit import Workbook
from xlsxedit.bulk import write_cell_value
from tests.conftest import INSPECT_FIXTURES
from tests.preservation import assert_preserved, read_pkg
from tests.test_formula_hygiene import _plant

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
