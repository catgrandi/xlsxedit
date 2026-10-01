"""Table part integrity: ``write_dataframe`` resizing.

Every test reads the table parts back from the saved package; the suite guard
in ``tests/conftest.py`` also runs ``check_consistency`` on every workbook
built here.
"""

from __future__ import annotations

import typing

import pytest
from lxml import etree

from xlsxedit import Table, Workbook
from xlsxedit.exceptions import TableError
from xlsxedit.opc.constants import SML_NS
from tests.conftest import INSPECT_FIXTURES
from tests.preservation import assert_preserved, check_consistency, read_pkg

# Sheet "Table" holds Table2 (id 2) over A1:C11: Item, Price, Quantity.
TABLES = INSPECT_FIXTURES["ChartsAndTables"]
TABLE = "xl/tables/table1.xml"
SHEET = "xl/worksheets/sheet3.xml"
FIXTURE_COLUMNS = [("1", "Item"), ("2", "Price"), ("3", "Quantity")]


def _sml(localname: str) -> str:
    return f"{{{SML_NS}}}{localname}"


class Frame:
    """A pandas-like frame: ``columns`` and ``itertuples``."""

    def __init__(self, columns, rows):
        self.columns = columns
        self._rows = rows

    def itertuples(self, index=False):
        return (tuple(row) for row in self._rows)


def _saved_tables(wb: Workbook) -> dict[str, etree._Element]:
    """The table parts of the saved package, by member name."""
    pkg = read_pkg(wb)
    return {m: etree.fromstring(pkg[m]) for m in sorted(pkg) if m.startswith("xl/tables/")}


def _summary(table: etree._Element) -> dict:
    """Ref, autoFilter ref and columns of a saved table part."""
    af = table.find(_sml("autoFilter"))
    cols = table.find(_sml("tableColumns"))
    return {
        "ref": table.get("ref"),
        "autoFilter": None if af is None else af.get("ref"),
        "count": cols.get("count"),
        "columns": [(c.get("id"), c.get("name")) for c in cols.iterchildren(_sml("tableColumn"))],
    }


def _identity(table: etree._Element) -> tuple[str | None, str | None, str | None]:
    return table.get("id"), table.get("name"), table.get("displayName")


def _fixture_table(ref: str = "A1:C11", columns=FIXTURE_COLUMNS) -> dict:
    return {"ref": ref, "autoFilter": ref, "count": str(len(columns)), "columns": columns}


def _fill(ws, at_cell: str, columns: list[str], rows: int) -> None:
    """Write a header row of ``columns`` and ``rows`` data rows at ``at_cell``."""
    data = [[f"{name}{i}" for name in columns] for i in range(1, rows + 1)]
    ws.write_rows([columns, *data], at_cell=at_cell)


def _side_by_side() -> Workbook:
    """Sheet1 with table Left over A1:B3 and table Right over D1:E3."""
    wb = Workbook.create()
    ws = wb["Sheet1"]
    _fill(ws, "A1", ["Item", "Qty"], rows=2)
    _fill(ws, "D1", ["Day", "N"], rows=2)
    ws.add_table("A1:B3", ["Item", "Qty"], name="Left")
    ws.add_table("D1:E3", ["Day", "N"], name="Right")
    return wb


# --- write_dataframe ---------------------------------------------------------


def test_write_dataframe_resizes_only_the_table_it_writes_into():
    wb = _side_by_side()
    wb.write_dataframe([("Mon", 1), ("Tue", 2), ("Wed", 3)], at_cell="D2", header=False)
    tables = _saved_tables(wb)
    left, right = tables["xl/tables/table1.xml"], tables["xl/tables/table2.xml"]
    assert _summary(left) == _fixture_table("A1:B3", [("1", "Item"), ("2", "Qty")])
    assert _summary(right) == _fixture_table("D1:E4", [("1", "Day"), ("2", "N")])


def test_write_dataframe_leaves_a_table_it_misses_byte_for_byte():
    before = read_pkg(TABLES)
    wb = Workbook.open(TABLES)
    wb.write_dataframe([("a", 1), ("b", 2)], sheet="Table", at_cell="F1", header=False)
    after = read_pkg(wb)
    assert after[TABLE] == before[TABLE]
    assert_preserved(before, after, expected_changed={SHEET, "xl/sharedStrings.xml"})


def test_write_dataframe_body_rows_keep_the_header_and_end_the_table():
    wb = Workbook.open(TABLES)
    rows = [("Pen", 1.5, 4), ("Ink", 3.0, 2)]
    assert wb.write_dataframe(rows, sheet="Table", at_cell="A2", header=False) == 2
    table = _saved_tables(wb)[TABLE]
    assert _identity(table) == ("2", "Table2", "Table2")
    assert _summary(table) == _fixture_table("A1:C3")


def test_write_dataframe_header_row_must_hold_the_column_names():
    wb = Workbook.open(TABLES)
    wb.write_dataframe(Frame(["Item", "Price", "Quantity"], [("Pen", 1.5, 4)]), sheet="Table")
    assert _summary(_saved_tables(wb)[TABLE]) == _fixture_table("A1:C2")

    before = read_pkg(wb)
    renamed = Frame(["item", "Price", "Qty"], [("Pen", 1.5, 4)])
    with pytest.raises(
        TableError, match="A1 holds 'item', expected 'Item'; C1 holds 'Qty', expected 'Quantity'"
    ):
        wb.write_dataframe(renamed, sheet="Table")
    assert_preserved(before, read_pkg(wb))


def test_write_dataframe_header_without_rows_must_hold_the_column_names():
    wb = Workbook.open(TABLES)
    before = read_pkg(wb)
    with pytest.raises(
        TableError, match="overwrite the header of table 'Table2' .*A1 holds 'item', expected 'Item'"
    ):
        wb.write_dataframe(Frame(["item", "Price", "Quantity"], []), sheet="Table")
    assert_preserved(before, read_pkg(wb))

    assert wb.write_dataframe(Frame(["Item", "Price", "Quantity"], []), sheet="Table") == 0
    wb.write_dataframe(Frame(["a", "b", "c"], []), sheet="Table", at_cell="A5")  # a body row
    assert _summary(_saved_tables(wb)[TABLE]) == _fixture_table()


def test_write_dataframe_body_rows_from_the_header_row_must_start_with_the_names():
    wb = Workbook.open(TABLES)
    before = read_pkg(wb)
    with pytest.raises(TableError, match="would become the header row of table 'Table2'"):
        wb.write_dataframe([("Pen", 1.5, 4), ("Ink", 3.0, 2)], sheet="Table", header=False)
    names = ("Item", "Price", "Quantity")
    with pytest.raises(TableError, match="would leave it without a data row"):
        wb.write_dataframe([names], sheet="Table", header=False)
    assert_preserved(before, read_pkg(wb))

    wb.write_dataframe([names, ("Pen", 1.5, 4)], sheet="Table", header=False)
    assert _summary(_saved_tables(wb)[TABLE]) == _fixture_table("A1:C2")


@pytest.mark.parametrize(
    ("at_cell", "rows"),
    [
        ("A2", [("Pen", 1.5)]),
        ("B2", [("Pen", 1.5, 4)]),
        ("A2", [("Pen", 1.5, 4, "extra")]),
    ],
    ids=["narrower", "shifted", "wider"],
)
def test_write_dataframe_refuses_rows_that_miss_the_table_columns(at_cell, rows):
    wb = Workbook.open(TABLES)
    before = read_pkg(wb)
    with pytest.raises(
        TableError, match=r"does not span exactly the columns of table 'Table2' \(A1:C11\)"
    ):
        wb.write_dataframe(rows, sheet="Table", at_cell=at_cell, header=False, clear_range="A2:C5")
    assert_preserved(before, read_pkg(wb))

    wb.write_dataframe(rows, sheet="Table", at_cell=at_cell, header=False, resize_table=False)
    assert read_pkg(wb)[TABLE] == before[TABLE]


def test_write_dataframe_refuses_rows_that_overlap_two_tables():
    wb = _side_by_side()
    before = read_pkg(wb)
    with pytest.raises(TableError, match="overlaps tables 'Left', 'Right'; write within one table"):
        wb.write_dataframe([("a", 1, "x", "b", 2)], at_cell="A2", header=False)
    assert_preserved(before, read_pkg(wb))


def test_write_dataframe_never_grows_a_table_over_another():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    _fill(ws, "A1", ["Item", "Qty"], rows=1)
    _fill(ws, "A4", ["Day", "N"], rows=1)
    ws.add_table("A1:B2", ["Item", "Qty"], name="Upper")
    ws.add_table("A4:B5", ["Day", "N"], name="Middle")
    before = read_pkg(wb)
    over_middle = r"would overlap table 'Middle' \(A4:B5\)"
    with pytest.raises(TableError, match=f"resizing table 'Upper' to A1:B8 {over_middle}"):
        wb.write_dataframe([("x", 9)], at_cell="A8", header=False, table="Upper")
    rows = [(f"r{i}", i) for i in range(6)]
    with pytest.raises(TableError, match="overlaps tables 'Upper', 'Middle'"):
        wb.write_dataframe(rows, at_cell="A2", header=False)
    with pytest.raises(TableError, match=f"resizing table 'Upper' to A1:B7 {over_middle}"):
        wb.write_dataframe(rows, at_cell="A2", header=False, table="Upper")
    assert_preserved(before, read_pkg(wb))


def test_write_dataframe_table_grows_a_table_from_below():
    wb = Workbook.open(TABLES)
    rows = [("Pen", 1.5, 4), ("Ink", 3.0, 2)]
    wb.write_dataframe(rows, sheet="Table", at_cell="A12", header=False)
    assert _summary(_saved_tables(wb)[TABLE]) == _fixture_table("A1:C11")
    wb.write_dataframe(rows, sheet="Table", at_cell="A12", header=False, table="Table2")
    assert _summary(_saved_tables(wb)[TABLE]) == _fixture_table("A1:C13")


@pytest.mark.parametrize(
    "pick", [lambda ws: "TABLE2", lambda ws: ws.tables[0]], ids=["name", "object"]
)
def test_write_dataframe_table_takes_a_name_or_a_table(pick):
    wb = Workbook.open(TABLES)
    table = pick(wb["Table"])
    wb.write_dataframe([("Pen", 1.5, 4)], sheet="Table", at_cell="A12", header=False, table=table)
    assert _summary(_saved_tables(wb)[TABLE]) == _fixture_table("A1:C12")


def test_write_dataframe_table_must_name_a_table_on_the_sheet():
    wb = Workbook.open(TABLES)
    before = read_pkg(wb)
    rows = [("Pen", 1.5, 4)]
    with pytest.raises(TableError, match="the workbook has no table named 'Nope'"):
        wb.write_dataframe(rows, sheet="Table", at_cell="A12", header=False, table="Nope")
    with pytest.raises(TableError, match="table 'Table2' is on sheet 'Table', not 'bar chart'"):
        wb.write_dataframe(rows, sheet="bar chart", at_cell="A12", header=False, table="Table2")
    other = Workbook.open(TABLES)["Table"].tables[0]
    with pytest.raises(TableError, match="is not a table on sheet 'Table'"):
        wb.write_dataframe(rows, sheet="Table", at_cell="A12", header=False, table=other)
    with pytest.raises(TypeError, match="table must be a table name or a Table"):
        wb.write_dataframe(rows, sheet="Table", at_cell="A12", header=False, table=2)
    for kwargs in ({"mode": "insert"}, {"resize_table": False}):
        with pytest.raises(ValueError, match="needs mode='overwrite' and resize_table=True"):
            wb.write_dataframe(rows, sheet="Table", at_cell="A12", table="Table2", **kwargs)
    assert_preserved(before, read_pkg(wb))


def test_write_dataframe_type_hints_resolve():
    assert typing.get_type_hints(Workbook.write_dataframe)["table"] == str | Table | None


def test_write_dataframe_refuses_a_table_whose_columns_already_disagree(unchecked_workbooks):
    wb = Workbook.open(TABLES)
    part = wb["Table"].tables[0]._part
    part._blob = part.blob.replace(b'<tableColumns count="3">', b'<tableColumns count="4">')
    with pytest.raises(TableError, match=r"has 3 tableColumn elements \(count=4\)"):
        wb.write_dataframe([("Pen", 1.5, 4)], sheet="Table", at_cell="A2", header=False)
    violations = check_consistency(wb, waive={"table-columns"})
    assert [v.detail for v in violations] == ["tableColumns/@count=4 but 3 tableColumn children"]
