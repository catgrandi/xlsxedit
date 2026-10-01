"""Table part integrity: ``write_dataframe`` resizing, ``add_table`` naming and
validation, ``copy_worksheet`` clones, and ``tableColumns`` on column insert.

Every test reads the table parts back from the saved package; the suite guard
in ``tests/conftest.py`` also runs ``check_consistency`` on every workbook
built here.
"""

from __future__ import annotations

import re
import typing

import pytest
from lxml import etree

from xlsxedit import Table, Workbook
from xlsxedit.exceptions import InvalidRangeError, TableError
from xlsxedit.opc.constants import SML_NS
from tests.conftest import INSPECT_FIXTURES
from tests.preservation import assert_preserved, check_consistency, read_pkg
from tests.schema_validate import assert_valid_package

# Sheet "Table" holds Table2 (id 2) over A1:C11: Item, Price, Quantity.
TABLES = INSPECT_FIXTURES["ChartsAndTables"]
TABLE = "xl/tables/table1.xml"
SHEET = "xl/worksheets/sheet3.xml"
FIXTURE_UID = "{ADEF1FC0-C6FE-3447-9A1E-67E59ED4BB8C}"
FIXTURE_COLUMNS = [("1", "Item"), ("2", "Price"), ("3", "Quantity")]

_XR_UID = "{http://schemas.microsoft.com/office/spreadsheetml/2014/revision}uid"
_XR3_UID = "{http://schemas.microsoft.com/office/spreadsheetml/2016/revision3}uid"
_GUID = re.compile(r"\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}")


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


def _add_defined_name(wb: Workbook, name: str) -> None:
    wb_elm = wb._workbook_part.element
    block = wb_elm.find(_sml("definedNames"))
    if block is None:
        block = etree.Element(_sml("definedNames"))
        wb_elm.insert(wb._workbook_child_insert_index(wb_elm, "definedNames"), block)
    etree.SubElement(block, _sml("definedName"), name=name).text = f"'{wb.sheetnames[0]}'!$A$1"


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


# --- add_table: names and revision uids ---------------------------------------


def test_add_table_names_each_table_after_its_workbook_wide_id():
    wb = Workbook.create()
    other = wb.add_worksheet("Other")
    for ws in (wb["Sheet1"], other):
        _fill(ws, "A1", ["x", "y"], rows=1)
        ws.add_table("A1:B2", ["x", "y"])
    tables = _saved_tables(wb)
    assert [_identity(t) for t in tables.values()] == [
        ("1", "Table1", "Table1"),
        ("2", "Table2", "Table2"),
    ]


def test_add_table_default_name_skips_names_in_use():
    wb = Workbook.open(TABLES)
    _add_defined_name(wb, "TABLE4")
    ws = wb.add_worksheet("New")
    _fill(ws, "A1", ["x", "y"], rows=1)
    _fill(ws, "D1", ["p", "q"], rows=1)
    assert ws.add_table("A1:B2", ["x", "y"]).name == "Table3"
    assert ws.add_table("D1:E2", ["p", "q"]).name == "Table5"
    new = list(_saved_tables(wb).values())[1:]
    assert [_identity(t) for t in new] == [("3", "Table3", "Table3"), ("4", "Table5", "Table5")]


def test_add_table_gives_each_table_fresh_revision_uids():
    tables = _saved_tables(_side_by_side())
    uids = []
    for table in tables.values():
        uid = table.get(_XR_UID)
        assert _GUID.fullmatch(uid)
        assert table.find(_sml("autoFilter")).get(_XR_UID) == uid  # Excel's convention
        uids.append(uid)
    assert len(set(uids)) == 2
    assert FIXTURE_UID not in uids


@pytest.mark.parametrize(
    "name",
    [
        "My Table",
        "T1",
        "xfd1048576",
        "R1C1",
        "RC",
        "r",
        "C3",
        "1st",
        "Sales-2024",
        "Sales!",
        "",
        "S" * 256,
    ],
)
def test_add_table_rejects_invalid_names(name):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    _fill(ws, "A1", ["x", "y"], rows=1)
    with pytest.raises(TableError, match=f"^table name {re.escape(repr(name))} "):
        ws.add_table("A1:B2", ["x", "y"], name=name)
    with pytest.raises(TableError, match=f"^table display_name {re.escape(repr(name))} "):
        ws.add_table("A1:B2", ["x", "y"], name="Valid", display_name=name)
    assert ws.tables == []


@pytest.mark.parametrize(
    "name",
    [
        "Sales",
        "_Sales",
        "\\Sales",
        "Sales.2024",
        "Data2024",
        "R2D2",
        "Tabelle_Ä",
        "S" * 255,
        "YTD2024",  # column YTD lies past XFD, so this is no cell reference
        "XFE1",
    ],
)
def test_add_table_accepts_valid_names(name):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    _fill(ws, "A1", ["x", "y"], rows=1)
    ws.add_table("A1:B2", ["x", "y"], name=name)
    assert _identity(_saved_tables(wb)[TABLE]) == ("1", name, name)


def test_add_table_rejects_names_already_in_use():
    wb = Workbook.open(TABLES)
    _add_defined_name(wb, "Sales")
    ws = wb.add_worksheet("New")
    _fill(ws, "A1", ["x", "y"], rows=1)
    in_use = "is already used by a table or defined name in this workbook"
    with pytest.raises(TableError, match=f"table name 'table2' {in_use}"):
        ws.add_table("A1:B2", ["x", "y"], name="table2")
    with pytest.raises(TableError, match=f"table display_name 'TABLE2' {in_use}"):
        ws.add_table("A1:B2", ["x", "y"], name="Fresh", display_name="TABLE2")
    with pytest.raises(TableError, match=f"table name 'SALES' {in_use}"):
        ws.add_table("A1:B2", ["x", "y"], name="SALES")
    assert ws.tables == []


# --- add_table: columns, range, overlap and header cells -----------------------


@pytest.mark.parametrize(
    ("columns", "message"),
    [
        ([], "columns must not be empty"),
        (["x", ""], "column names must be non-empty strings, got ''"),
        (["x", 2], "column names must be non-empty strings, got 2"),
        (["Qty", "qty"], "duplicate column name 'qty'"),
        (["x", "y", "z"], "table range A1:B2 is 2 columns wide but 3 column names were given"),
    ],
)
def test_add_table_rejects_bad_column_names(columns, message):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    _fill(ws, "A1", ["x", "y"], rows=1)
    with pytest.raises(TableError, match=re.escape(message)):
        ws.add_table("A1:B2", columns)
    assert ws.tables == []


@pytest.mark.parametrize(
    ("cell_range", "error"),
    [
        ("A1:B1", TableError),
        ("B2:A1", InvalidRangeError),
        ("A0:B2", InvalidRangeError),
        ("XFD1:XFE2", InvalidRangeError),
        ("A1048576:B1048577", InvalidRangeError),
        ("$A$1:$B$2", InvalidRangeError),
    ],
)
def test_add_table_rejects_ranges_that_cannot_hold_a_table(cell_range, error):
    ws = Workbook.create()["Sheet1"]
    with pytest.raises(error):
        ws.add_table(cell_range, ["x", "y"], write_header=True)
    assert ws.tables == []


def test_add_table_normalises_the_range():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.add_table("a1:b3", ["x", "y"], write_header=True)
    assert _summary(_saved_tables(wb)[TABLE])["ref"] == "A1:B3"


def test_add_table_rejects_a_range_that_overlaps_a_table_on_the_sheet():
    wb = _side_by_side()
    ws = wb["Sheet1"]
    before = read_pkg(wb)
    with pytest.raises(TableError, match=r"table range B3:C5 overlaps table 'Left' \(A1:B3\)"):
        ws.add_table("B3:C5", ["y", "z"], write_header=True)
    assert_preserved(before, read_pkg(wb))

    other = wb.add_worksheet("Other")
    other.add_table("A1:B3", ["x", "y"], write_header=True)
    assert len(_saved_tables(wb)) == 3


def test_add_table_header_cells_must_hold_the_column_names():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "Item"
    ws["B1"].value = 7
    ws["C1"].formula = '"Qty"'
    before = read_pkg(wb)
    with pytest.raises(TableError) as excinfo:
        ws.add_table("A1:D2", ["Item", "Price", "Qty", "Note"])
    assert str(excinfo.value) == (
        "the header cells of table range A1:D2 do not match the column names: "
        "B1 holds 7, expected 'Price'; C1 holds a formula, expected 'Qty'; "
        "D1 is empty, expected 'Note'; write the header first or pass write_header=True"
    )
    assert_preserved(before, read_pkg(wb))


def test_add_table_write_header_writes_the_names_in_column_order():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["B1"].formula = "1+1"
    ws["D1"].value = "kept"
    ws.add_table("A1:C2", ["x", "y", "z"], write_header=True)
    sheet = etree.fromstring(read_pkg(wb)["xl/worksheets/sheet1.xml"])
    row = sheet.find(f"{_sml('sheetData')}/{_sml('row')}")
    assert [c.get("r") for c in row] == ["A1", "B1", "C1", "D1"]
    assert [ws[a].value for a in ("A1", "B1", "C1", "D1")] == ["x", "y", "z", "kept"]
    assert not ws["B1"].has_formula
    assert_valid_package(wb)


# --- copy_worksheet ------------------------------------------------------------


def test_copy_worksheet_gives_every_cloned_table_its_own_id_name_and_uid():
    wb = _side_by_side()
    wb.copy_worksheet("Sheet1", "Copy")
    tables = list(_saved_tables(wb).values())
    assert [_identity(t) for t in tables] == [
        ("1", "Left", "Left"),
        ("2", "Right", "Right"),
        ("3", "Table3", "Table3"),
        ("4", "Table4", "Table4"),
    ]
    uids = [t.get(_XR_UID) for t in tables]
    assert len(set(uids)) == 4
    for table, uid in zip(tables, uids):
        assert table.find(_sml("autoFilter")).get(_XR_UID) == uid


def test_copy_worksheet_refreshes_the_column_uids_of_a_cloned_table():
    before = read_pkg(TABLES)
    wb = Workbook.open(TABLES)
    wb.copy_worksheet("Table", "Table copy")
    pkg = read_pkg(wb)
    assert pkg[TABLE] == before[TABLE]
    source, clone = (etree.fromstring(pkg[m]) for m in (TABLE, "xl/tables/table2.xml"))
    assert _identity(clone) == ("3", "Table3", "Table3")
    assert _summary(clone) == _fixture_table()
    assert clone.get(_XR_UID) not in (None, FIXTURE_UID)
    assert clone.find(_sml("autoFilter")).get(_XR_UID) == clone.get(_XR_UID)
    source_uids = {c.get(_XR3_UID) for c in source.iter(_sml("tableColumn"))}
    clone_uids = {c.get(_XR3_UID) for c in clone.iter(_sml("tableColumn"))}
    assert len(clone_uids) == 3
    assert not source_uids & clone_uids


# --- insert_columns --------------------------------------------------------------


def test_insert_columns_inside_a_table_adds_table_columns_in_place():
    wb = Workbook.open(TABLES)
    wb["Table"].insert_columns([[None, "a"], [None, 1]], at_col="B")
    table = _saved_tables(wb)[TABLE]
    columns = [("1", "Item"), ("4", "Column1"), ("5", "Column2"), ("2", "Price"), ("3", "Quantity")]
    assert _summary(table) == _fixture_table("A1:E11", columns)
    uids = {c.get("name"): c.get(_XR3_UID) for c in table.iter(_sml("tableColumn"))}
    assert uids["Item"] == "{54273001-98B0-DA43-A786-292A1776F894}"
    assert uids["Column1"] is None
    assert_valid_package(wb)


def test_insert_columns_names_new_table_columns_like_excel():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    _fill(ws, "A1", ["Column1", "column3"], rows=1)
    ws.add_table("A1:B2", ["Column1", "column3"])
    ws.insert_columns([[None], [None], [None]], at_col="B")
    names = [name for _, name in _summary(_saved_tables(wb)[TABLE])["columns"]]
    assert names == ["Column1", "Column2", "Column4", "Column5", "column3"]


@pytest.mark.parametrize(("at_col", "ref"), [("A", "B1:D11"), ("D", "A1:C11")])
def test_insert_columns_at_a_table_edge_moves_or_misses_it(at_col, ref):
    wb = Workbook.open(TABLES)
    wb["Table"].insert_columns([[None]], at_col=at_col)
    assert _summary(_saved_tables(wb)[TABLE]) == _fixture_table(ref)


def test_insert_columns_moves_filter_columns_with_their_columns():
    wb = Workbook.open(TABLES)
    part = wb["Table"].tables[0]._part
    filters = (
        b'<filterColumn colId="0"><filters><filter val="Pen"/></filters></filterColumn>'
        b'<filterColumn colId="2"><filters><filter val="3"/></filters></filterColumn>'
    )
    part._blob = part.blob.replace(
        b'4BB8C}"/><tableColumns', b'4BB8C}">' + filters + b"</autoFilter><tableColumns"
    )
    wb["Table"].insert_columns([[None]], at_col="B")
    table = _saved_tables(wb)[TABLE]
    assert [fc.get("colId") for fc in table.iter(_sml("filterColumn"))] == ["0", "3"]


def test_insert_columns_refuses_a_table_bound_to_a_query():
    wb = Workbook.open(TABLES)
    part = wb["Table"].tables[0]._part
    part._blob = part.blob.replace(b'name="Price"/>', b'name="Price" queryTableFieldId="2"/>')
    bound = part.blob
    with pytest.raises(
        TableError, match=r"cannot insert columns inside table 'Table2' \(A1:C11\)"
    ):
        wb["Table"].insert_columns([[None]], at_col="B")
    assert part.blob == bound

    wb = Workbook.open(TABLES)
    part = wb["Table"].tables[0]._part
    part._blob = part.blob.replace(b'name="Price"/>', b'name="Price" queryTableFieldId="2"/>')
    wb["Table"].insert_columns([[None]], at_col="A")
    assert _summary(_saved_tables(wb)[TABLE])["ref"] == "B1:D11"
