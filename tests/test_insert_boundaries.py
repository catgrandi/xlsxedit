"""Insert boundaries on real fixtures (issue #5).

Each case inserts at the header row (first column), inside the body, at the
last data row (column) and just below (right of) it, then pins where the
tables, merges, conditional formats and data validations end up, which package
members changed, and that the result is schema-valid and consistent.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field

import pytest

from xlsxedit import Workbook
from xlsxedit.opc.constants import SML_NS
from tests.conftest import INSPECT_FIXTURES
from tests.preservation import assert_preserved, check_consistency, read_pkg
from tests.schema_validate import assert_valid_package
from tests.test_insert_rows import _add_list_validation

XM = "http://schemas.microsoft.com/office/excel/2006/main"
NS = {"m": SML_NS, "xm": XM}
POSITIONS = ("header", "body", "last", "below")


def _shared_ref(ws) -> str:
    return ws._part.element.find(".//m:f[@t='shared'][@ref]", NS).get("ref")


def _dv(ws) -> str:
    return ws._part.element.find("m:dataValidations/m:dataValidation", NS).get("sqref")


def _merges(ws) -> tuple[str, ...]:
    return tuple(ws.merged_ranges)


def _cf(ws) -> tuple[str, ...]:
    return tuple(cf.cell_range for cf in ws.conditional_formatting)


def _x14(ws) -> tuple[str, ...]:
    return tuple(e.text for e in ws._part.element.iterfind(".//xm:sqref", NS))


def _table(ws) -> str:
    (table,) = ws.tables
    assert table._element.find("m:autoFilter", NS).get("ref") == table.ref
    return table.ref


GETTERS: dict[str, Callable] = {
    "shared": _shared_ref,
    "dv": _dv,
    "merges": _merges,
    "cf": _cf,
    "x14": _x14,
    "table": _table,
}


@dataclass
class Case:
    fixture: str
    sheet: str
    member: str
    setup: Callable
    at: dict[str, int | str]
    expect: dict[str, tuple]
    changed: frozenset[str] = frozenset()
    changed_unless_below: frozenset[str] = frozenset()
    removed: frozenset[str] = frozenset()
    positions: tuple[str, ...] = field(default=POSITIONS)


def _simple_formula(ws) -> None:
    _add_list_validation(ws, "C3:C7")
    ws.merge_cells("D3:D7")
    ws.merge_cells("A10:C10")


def _conditional_formatting(ws) -> None:
    _add_list_validation(ws, "E2:E7")
    ws.merge_cells("K2:K7")
    ws.merge_cells("A9:I9")


def _table_sheet(ws) -> None:
    _add_list_validation(ws, "B2:B11")
    ws.merge_cells("E2:E11")


def _nothing(ws) -> None:
    pass


_CALC_CHAIN = {
    "changed": frozenset({"xl/workbook.xml", "xl/_rels/workbook.xml.rels"}),
    "removed": frozenset({"xl/calcChain.xml"}),
}


def _cf_columns(*columns: str, rows: str = "2:7") -> tuple[str, ...]:
    first, last = rows.split(":")
    return tuple(f"{c}{first}:{c}{last}" for c in columns)


ROW_CASES = {
    "SimpleFormula": Case(
        "SimpleFormula",
        "Sheet1",
        "xl/worksheets/sheet1.xml",
        _simple_formula,
        at={"header": 2, "body": 5, "last": 8, "below": 9},
        expect={
            "shared": ("C5:C8", "C4:C8", "C4:C7", "C4:C7"),
            "dv": ("C4:C8", "C3:C8", "C3:C7", "C3:C7"),
            "merges": (
                ("D4:D8", "A11:C11"),
                ("D3:D8", "A11:C11"),
                ("D3:D7", "A11:C11"),
                ("D3:D7", "A11:C11"),
            ),
        },
        **_CALC_CHAIN,
    ),
    "ConditionalFormatting": Case(
        "ConditionalFormatting",
        "Sheet1",
        "xl/worksheets/sheet1.xml",
        _conditional_formatting,
        at={"header": 1, "body": 4, "last": 7, "below": 8},
        expect={
            "cf": (
                _cf_columns(*"ACEGI", rows="3:8"),
                _cf_columns(*"ACEGI", rows="2:8"),
                _cf_columns(*"ACEGI", rows="2:8"),
                _cf_columns(*"ACEGI"),
            ),
            "x14": (("C3:C8",), ("C2:C8",), ("C2:C8",), ("C2:C7",)),
            "dv": ("E3:E8", "E2:E8", "E2:E8", "E2:E7"),
            "merges": (
                ("K3:K8", "A10:I10"),
                ("K2:K8", "A10:I10"),
                ("K2:K8", "A10:I10"),
                ("K2:K7", "A10:I10"),
            ),
        },
        changed=frozenset({"xl/workbook.xml"}),
    ),
    "ChartsAndTables-table": Case(
        "ChartsAndTables",
        "Table",
        "xl/worksheets/sheet3.xml",
        _table_sheet,
        at={"header": 1, "body": 3, "last": 11, "below": 12},
        expect={
            "table": ("A2:C12", "A1:C12", "A1:C12", "A1:C11"),
            "dv": ("B3:B12", "B2:B12", "B2:B12", "B2:B11"),
            "merges": (("E3:E12",), ("E2:E12",), ("E2:E12",), ("E2:E11",)),
        },
        changed=frozenset({"xl/workbook.xml"}),
        changed_unless_below=frozenset({"xl/tables/table1.xml"}),
    ),
    "ChartsAndTables-merge": Case(
        "ChartsAndTables",
        "bar chart",
        "xl/worksheets/sheet1.xml",
        _nothing,
        at={"header": 1, "body": 3, "last": 6, "below": 7},
        expect={"merges": (("A2:B2",), ("A1:B1",), ("A1:B1",), ("A1:B1",))},
        changed=frozenset({"xl/workbook.xml"}),
    ),
}

COLUMN_CASES = {
    "SimpleFormula": Case(
        "SimpleFormula",
        "Sheet1",
        "xl/worksheets/sheet1.xml",
        _simple_formula,
        at={"header": "A", "body": "B", "last": "C", "below": "D"},
        expect={
            "shared": ("D4:D7", "D4:D7", "D4:D7", "C4:C7"),
            "dv": ("D3:D7", "D3:D7", "D3:D7", "C3:C7"),
            "merges": (
                ("E3:E7", "B10:D10"),
                ("E3:E7", "A10:D10"),
                ("E3:E7", "A10:D10"),
                ("E3:E7", "A10:C10"),
            ),
        },
        **_CALC_CHAIN,
    ),
    "ConditionalFormatting": Case(
        "ConditionalFormatting",
        "Sheet1",
        "xl/worksheets/sheet1.xml",
        _conditional_formatting,
        at={"header": "A", "body": "E", "last": "I", "below": "J"},
        expect={
            "cf": (
                _cf_columns(*"BDFHJ"),
                _cf_columns(*"ACFHJ"),
                _cf_columns(*"ACEGJ"),
                _cf_columns(*"ACEGI"),
            ),
            "x14": (("D2:D7",), ("C2:C7",), ("C2:C7",), ("C2:C7",)),
            "dv": ("F2:F7", "F2:F7", "E2:E7", "E2:E7"),
            "merges": (
                ("L2:L7", "B9:J9"),
                ("L2:L7", "A9:J9"),
                ("L2:L7", "A9:J9"),
                ("L2:L7", "A9:I9"),
            ),
        },
        changed=frozenset({"xl/workbook.xml"}),
    ),
    # Inside the table a column insert also needs a new tableColumn: issue #7.
    "ChartsAndTables-table": Case(
        "ChartsAndTables",
        "Table",
        "xl/worksheets/sheet3.xml",
        _table_sheet,
        at={"header": "A", "below": "D"},
        expect={
            "table": ("B1:D11", "A1:C11"),
            "dv": ("C2:C11", "B2:B11"),
            "merges": (("F2:F11",), ("F2:F11",)),
        },
        changed=frozenset({"xl/workbook.xml"}),
        changed_unless_below=frozenset({"xl/tables/table1.xml"}),
        positions=("header", "below"),
    ),
    "ChartsAndTables-merge": Case(
        "ChartsAndTables",
        "bar chart",
        "xl/worksheets/sheet1.xml",
        _nothing,
        at={"header": "A", "body": "B", "last": "D", "below": "E"},
        expect={"merges": (("B1:C1",), ("A1:C1",), ("A1:B1",), ("A1:B1",))},
        changed=frozenset({"xl/workbook.xml"}),
    ),
}


def _params(cases: dict[str, Case]):
    return [
        pytest.param(case, position, id=f"{name}-{position}")
        for name, case in cases.items()
        for position in case.positions
    ]


def _check(case: Case, position: str, insert: Callable) -> None:
    wb = Workbook.open(INSPECT_FIXTURES[case.fixture])
    ws = wb[case.sheet]
    case.setup(ws)
    before = read_pkg(wb)
    insert(ws, case.at[position])
    after = read_pkg(wb)

    changed = {case.member, *case.changed}
    if position != "below":
        changed |= case.changed_unless_below
    assert_preserved(before, after, expected_changed=changed, expected_removed=case.removed)
    index = case.positions.index(position)
    for name, values in case.expect.items():
        assert GETTERS[name](ws) == values[index], name
    check_consistency(after)
    assert_valid_package(wb)


@pytest.mark.parametrize(("case", "position"), _params(ROW_CASES))
def test_insert_rows_boundaries(case: Case, position: str):
    _check(case, position, lambda ws, at: ws.insert_rows([[1, 2]], at_row=at))


@pytest.mark.parametrize(("case", "position"), _params(COLUMN_CASES))
def test_insert_columns_boundaries(case: Case, position: str):
    _check(case, position, lambda ws, at: ws.insert_columns([[1, 2]], at_col=at))
