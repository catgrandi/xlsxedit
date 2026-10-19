"""Tables: the ``Table`` model over the live table part, ``Table.resize`` and the
setters, ``write_dataframe`` resizing, ``add_table`` naming and validation,
``copy_worksheet`` clones, and tables under row and column inserts.

Every test reads the table parts back from the saved package; the suite guard
in ``tests/conftest.py`` also runs ``check_consistency`` on every workbook
built here.
"""

from __future__ import annotations

import io
import re
import typing
import zipfile

import pytest
from lxml import etree

from xlsxedit import Table, TableColumn, TableStyle, Workbook
from xlsxedit.exceptions import FormulaGroupError, InvalidRangeError, TableError
from xlsxedit.opc.constants import SML_NS
from xlsxedit.parts import TablePart
from tests.conftest import INSPECT_FIXTURES
from tests.preservation import assert_preserved, canon, check_consistency, read_pkg
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


def _edited_tables(*edits: tuple[bytes, bytes], member: str = TABLE) -> Workbook:
    """The ``ChartsAndTables`` fixture with each ``(old, new)`` replaced in ``member``."""
    return _edited_package(*((member, old, new) for old, new in edits))


def _edited_package(*edits: tuple[str, bytes, bytes], added: dict | None = None) -> Workbook:
    """The ``ChartsAndTables`` fixture with each ``(member, old, new)`` replacement
    made and the ``added`` members written."""
    pkg = dict(read_pkg(TABLES))
    for member, old, new in edits:
        assert pkg[member].count(old) == 1, old
        pkg[member] = pkg[member].replace(old, new)
    pkg.update(added or {})
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        for member, blob in pkg.items():
            z.writestr(member, blob)
    buf.seek(0)
    return Workbook.open(buf)


_FIXTURE_FILTER = f'<autoFilter ref="A1:C11" xr:uid="{FIXTURE_UID}"/>'.encode()

# Table2 with a totals row on row 11, a descending sort on Price, a dxf and
# cell style, and a label, a totals function and a calculated column.
_RICH = (
    (TABLE, b'totalsRowShown="0"', b'totalsRowCount="1" headerRowDxfId="0" dataCellStyle="Normal"'),
    (
        TABLE,
        _FIXTURE_FILTER,
        f'<autoFilter ref="A1:C10" xr:uid="{FIXTURE_UID}">'.encode()
        + b'<sortState ref="A2:C10"><sortCondition descending="1" ref="B2:B10"/></sortState>'
        + b"</autoFilter>",
    ),
    (TABLE, b'name="Item"/>', b'name="Item" uniqueName="Item" totalsRowLabel="Total"/>'),
    (TABLE, b'name="Price"/>', b'name="Price" totalsRowFunction="sum" dataDxfId="0"/>'),
    (
        TABLE,
        b'name="Quantity"/>',
        b'name="Quantity" totalsRowFunction="custom">'
        b"<calculatedColumnFormula>Table2[[#This Row],[Price]]*2</calculatedColumnFormula>"
        b"<totalsRowFormula>SUM(Table2[Quantity])</totalsRowFormula></tableColumn>",
    ),
    ("xl/styles.xml", b'<dxfs count="0"/>', b'<dxfs count="1"><dxf><font><b/></font></dxf></dxfs>'),
)


def _with_workbook_part(member: str, content_type: str, reltype: str, xml: bytes) -> Workbook:
    """The fixture with ``member`` added and related from the workbook part."""
    override = f'<Override PartName="/{member}" ContentType="{content_type}"/>'
    rel = f'<Relationship Id="rId99" Type="{reltype}" Target="{member[3:]}"/>'
    return _edited_package(
        ("[Content_Types].xml", b"</Types>", override.encode() + b"</Types>"),
        ("xl/_rels/workbook.xml.rels", b"</Relationships>", rel.encode() + b"</Relationships>"),
        added={member: xml},
    )


def _rich() -> Workbook:
    return _edited_package(*_RICH)


def _replaced(blob: bytes, *edits: tuple[bytes, bytes]) -> bytes:
    for old, new in edits:
        assert blob.count(old) == 1, old
        blob = blob.replace(old, new)
    return blob


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
        TableError,
        match="overwrite the header of table 'Table2' .*A1 holds 'item', expected 'Item'",
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
    wb = _edited_tables((b'<tableColumns count="3">', b'<tableColumns count="4">'))
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
    ],
)
def test_add_table_rejects_ranges_that_cannot_hold_a_table(cell_range, error):
    ws = Workbook.create()["Sheet1"]
    with pytest.raises(error):
        ws.add_table(cell_range, ["x", "y"], write_header=True)
    assert ws.tables == []


@pytest.mark.parametrize("cell_range", ["a1:b3", "$A$1:$B$3"])
def test_add_table_normalises_the_range(cell_range):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.add_table(cell_range, ["x", "y"], write_header=True)
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


@pytest.mark.parametrize("kind", ["array", "dataTable"])
def test_add_table_write_header_refuses_a_range_formula_anchor_before_writing(kind):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "old"
    anchor = ws._get_or_create_cell_element("B1")
    etree.SubElement(anchor, _sml("f"), t=kind, ref="B1:B3").text = "ROW(1:3)"
    before = read_pkg(wb)
    with pytest.raises(FormulaGroupError, match="B1 anchors"):
        ws.add_table("A1:C4", ["x", "y", "z"], write_header=True)
    assert ws.tables == []
    assert_preserved(before, read_pkg(wb))


@pytest.mark.parametrize("write_header", [False, True])
def test_add_table_refuses_a_column_name_xml_cannot_store(write_header):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "old"
    before = read_pkg(wb)
    with pytest.raises(TableError, match="XML cannot store"):
        ws.add_table("A1:B2", ["x", "y\x01"], write_header=write_header)
    assert_preserved(before, read_pkg(wb))


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
    filters = (
        b'<filterColumn colId="0"><filters><filter val="Pen"/></filters></filterColumn>'
        b'<filterColumn colId="2"><filters><filter val="3"/></filters></filterColumn>'
    )
    wb = _edited_tables(
        (b'4BB8C}"/><tableColumns', b'4BB8C}">' + filters + b"</autoFilter><tableColumns")
    )
    wb["Table"].insert_columns([[None]], at_col="B")
    table = _saved_tables(wb)[TABLE]
    assert [fc.get("colId") for fc in table.iter(_sml("filterColumn"))] == ["0", "3"]


def test_insert_columns_refuses_a_table_bound_to_a_query():
    bind = (b'name="Price"/>', b'name="Price" queryTableFieldId="2"/>')
    wb = _edited_tables(bind)
    before = read_pkg(wb)
    with pytest.raises(TableError, match=r"cannot insert columns inside table 'Table2' \(A1:C11\)"):
        wb["Table"].insert_columns([[None]], at_col="B")
    # The refusal comes from the plan, before any cell on the sheet moves.
    assert_preserved(before, read_pkg(wb))

    wb = _edited_tables(bind)
    wb["Table"].insert_columns([[None]], at_col="A")
    assert _summary(_saved_tables(wb)[TABLE])["ref"] == "B1:D11"


# --- the Table model ------------------------------------------------------------


def test_tables_are_live_proxies_over_their_table_parts():
    wb = Workbook.open(TABLES)
    ws = wb["Table"]
    table = ws.tables[0]
    assert ws.tables[0] is table
    assert type(table._part) is TablePart
    ws.insert_rows([["Pen", 1.5, 4]], at_row=5)
    assert table.ref == "A1:C12"
    assert _summary(_saved_tables(wb)[TABLE]) == _fixture_table("A1:C12")


def test_add_table_returns_the_proxy_ws_tables_hands_out():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    table = ws.add_table("A1:B2", ["x", "y"], write_header=True)
    assert ws.tables == [table] and ws.tables[0] is table
    assert type(table._part) is TablePart


def test_reading_a_table_and_writing_what_it_holds_leaves_its_part_byte_for_byte():
    before = read_pkg(TABLES)
    wb = Workbook.open(TABLES)
    table = wb["Table"].tables[0]
    for attr in ("name", "display_name", "ref", "columns", "column_objects", "style", "data_range"):
        getattr(table, attr)
    table.resize("A1:C11")
    table.resize("a1:c11", columns=["Item", "Price", "Quantity"])
    table.display_name = "Table2"
    table.header_row_count = 1
    table.totals_row_count = 0
    table.totals_row_shown = False
    table.style = table.style
    table.data_dxf_id = None
    after = read_pkg(wb)
    assert after[TABLE] == before[TABLE]
    assert_preserved(before, after)


def test_table_reads_the_fixture_and_the_schema_defaults():
    table = Workbook.open(TABLES)["Table"].tables[0]
    assert (table.id, table.name, table.display_name, table.ref) == (
        2,
        "Table2",
        "Table2",
        "A1:C11",
    )
    assert table.columns == ["Item", "Price", "Quantity"]
    assert (table.header_row_count, table.totals_row_count, table.totals_row_shown) == (1, 0, False)
    assert table.style == TableStyle("TableStyleMedium2", False, False, True, False)
    assert table.data_dxf_id is table.header_row_dxf_id is table.table_border_dxf_id is None
    assert table.data_cell_style is None
    assert (table.header_range, table.data_range, table.totals_range) == ("A1:C1", "A2:C11", None)
    assert (table.first_data_row, table.last_data_row) == (2, 11)
    item = table.column_objects[0]
    assert isinstance(item, TableColumn)
    assert (item.id, item.name, item.totals_row_function) == (1, "Item", "none")
    assert item.calculated_column_formula is item.totals_row_formula is item.unique_name is None


def test_table_reads_every_formatting_and_column_attribute():
    table = _rich()["Table"].tables[0]
    assert (table.totals_row_count, table.totals_row_shown) == (1, True)
    assert (table.header_row_dxf_id, table.data_dxf_id, table.data_cell_style) == (
        0,
        None,
        "Normal",
    )
    assert (table.data_range, table.totals_range, table.last_data_row) == ("A2:C10", "A11:C11", 10)
    item, price, quantity = table.column_objects
    assert (item.unique_name, item.totals_row_label) == ("Item", "Total")
    assert (price.totals_row_function, price.data_dxf_id) == ("sum", 0)
    assert quantity.totals_row_function == "custom"
    assert quantity.calculated_column_formula == "Table2[[#This Row],[Price]]*2"
    assert quantity.totals_row_formula == "SUM(Table2[Quantity])"
    assert item.query_table_field_id is None


def test_a_table_without_a_header_row_has_no_header_range():
    wb = _edited_tables(
        (b'totalsRowShown="0"', b'headerRowCount="0" totalsRowShown="0"'), (_FIXTURE_FILTER, b"")
    )
    table = wb["Table"].tables[0]
    assert (table.header_range, table.data_range, table.first_data_row) == (None, "A1:C11", 1)


def test_table_reports_a_malformed_attribute():
    table = _edited_tables((b'totalsRowShown="0"', b'totalsRowShown="0" totalsRowCount="x"'))
    with pytest.raises(TableError, match="table 'Table2' has a malformed totalsRowCount 'x'"):
        table["Table"].tables[0].totals_row_count


# --- Table.resize ---------------------------------------------------------------


def test_resize_keeps_the_totals_row_out_of_the_filter_and_everything_else_as_it_was():
    wb = _rich()
    before = read_pkg(wb)
    wb["Table"].tables[0].resize("A1:C12")
    after = read_pkg(wb)
    assert_preserved(before, after, expected_changed={TABLE})
    expected = _replaced(
        before[TABLE],
        (b'ref="A1:C11"', b'ref="A1:C12"'),
        (b'<autoFilter ref="A1:C10"', b'<autoFilter ref="A1:C11"'),
        (b'<sortState ref="A2:C10"', b'<sortState ref="A2:C11"'),
        (b'ref="B2:B10"', b'ref="B2:B11"'),
    )
    # xr:uid, xr3:uid, mc:Ignorable, dxf ids, styles and column formulas are untouched.
    assert canon(TABLE, after[TABLE]) == canon(TABLE, expected)
    assert after[TABLE].startswith(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n')
    assert_valid_package(wb)


def test_resize_clamps_the_sort_state_into_the_new_data_rows():
    wb = _rich()
    wb["Table"].tables[0].resize("A1:C5")
    table = _saved_tables(wb)[TABLE]
    af = table.find(_sml("autoFilter"))
    assert af.get("ref") == "A1:C4"
    assert af.find(_sml("sortState")).get("ref") == "A2:C4"
    assert af.find(f"{_sml('sortState')}/{_sml('sortCondition')}").get("ref") == "B2:B4"


def test_resize_drops_a_sort_condition_whose_column_leaves_the_table():
    wb = _rich()
    wb["Table"].tables[0].resize("A1:A12", columns=["Item"])
    af = _saved_tables(wb)[TABLE].find(_sml("autoFilter"))
    assert af.get("ref") == "A1:A11"
    assert af.find(_sml("sortState")) is None


def test_resize_without_a_header_row_removes_the_auto_filter():
    wb = _edited_tables((b'totalsRowShown="0"', b'headerRowCount="0" totalsRowShown="0"'))
    wb["Table"].tables[0].resize("A1:C12")
    table = _saved_tables(wb)[TABLE]
    assert table.get("ref") == "A1:C12"
    assert table.find(_sml("autoFilter")) is None


def test_resize_adds_columns_named_by_columns_and_writes_empty_headers():
    wb = Workbook.open(TABLES)
    ws = wb["Table"]
    ws["E1"].value = "Note"
    ws.tables[0].resize("A1:E11", columns=["Item", "Price", "Quantity", "Total", "Note"])
    summary = _summary(_saved_tables(wb)[TABLE])
    assert summary == _fixture_table("A1:E11", [*FIXTURE_COLUMNS, ("4", "Total"), ("5", "Note")])
    assert [ws[a].value for a in ("D1", "E1")] == ["Total", "Note"]


def test_resize_drops_columns_and_moves_their_filters():
    filters = (
        b'<filterColumn colId="0"><filters><filter val="Pen"/></filters></filterColumn>'
        b'<filterColumn colId="2"><filters><filter val="3"/></filters></filterColumn>'
    )
    wb = _edited_tables((_FIXTURE_FILTER, _FIXTURE_FILTER[:-2] + b">" + filters + b"</autoFilter>"))
    wb["Table"].tables[0].resize("B1:C11", columns=["Price", "Quantity"])
    table = _saved_tables(wb)[TABLE]
    assert _summary(table) == _fixture_table("B1:C11", FIXTURE_COLUMNS[1:])
    assert [fc.get("colId") for fc in table.iter(_sml("filterColumn"))] == ["1"]


@pytest.mark.parametrize(
    ("ref", "columns", "message"),
    [
        ("A2:C12", None, r"would move its header row from row 1 to row 2"),
        ("A1:C1", None, r"resizing table 'Table2' to A1:C1 would leave it without a data row"),
        ("A1:D11", None, r"would change its columns \(A1:C11\); pass columns="),
        (
            "A1:C11",
            ["Item", "Cost", "Quantity"],
            r"named 'Price', not 'Cost'; resize does not rename",
        ),
        ("A1:D11", ["Item", "Price", "Quantity", "price"], r"duplicate column name 'price'"),
        ("A1:D11", ["Item", "Price"], r"2 column names were given for a range 4 columns wide"),
        (
            "A1:D11",
            ["Item", "Price", "Quantity", "Unit"],
            r"header cell D1 of new column 'Unit' holds 'x'",
        ),
    ],
    ids=[
        "header-moves",
        "no-data-row",
        "no-columns",
        "rename",
        "duplicate",
        "width",
        "header-cell",
    ],
)
def test_resize_refuses_without_changing_anything(ref, columns, message):
    wb = Workbook.open(TABLES)
    ws = wb["Table"]
    ws["D1"].value = "x"
    before = read_pkg(wb)
    with pytest.raises(TableError, match=message):
        ws.tables[0].resize(ref, columns=columns)
    assert_preserved(before, read_pkg(wb))


def test_resize_refuses_a_range_that_misses_the_table_or_overlaps_another():
    wb = _side_by_side()
    left = wb["Sheet1"].tables[0]
    before = read_pkg(wb)
    with pytest.raises(TableError, match=r"to G1:H3 would not overlap its current range A1:B3"):
        left.resize("G1:H3", columns=["Item", "Qty"])
    with pytest.raises(
        TableError, match=r"resizing table 'Left' to A1:D3 would overlap table 'Right' \(D1:E3\)"
    ):
        left.resize("A1:D3", columns=["Item", "Qty", "x", "Day"])
    with pytest.raises(InvalidRangeError):
        left.resize("B3:A1")
    with pytest.raises(TypeError):
        left.resize(None)
    assert_preserved(before, read_pkg(wb))


@pytest.mark.parametrize(
    ("ref", "columns"),
    [("A1:D11", ["Item", "Price", "Quantity", "New"]), ("A1:B11", ["Item", "Price"])],
    ids=["add", "drop"],
)
def test_resize_refuses_to_change_the_columns_of_a_table_bound_to_a_query(ref, columns):
    wb = _edited_tables((b'name="Price"/>', b'name="Price" queryTableFieldId="2"/>'))
    table = wb["Table"].tables[0]
    before = read_pkg(wb)
    with pytest.raises(TableError, match=r"would add or drop columns of a table bound to a query"):
        table.resize(ref, columns=columns)
    assert_preserved(before, read_pkg(wb))
    table.resize("A1:C12")
    assert table.ref == "A1:C12"


_SLICER_CACHE = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    b'<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
    b'xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" '
    b'name="Slicer_Quantity" sourceName="Quantity"><extLst>'
    b'<ext uri="{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}">'
    b'<x15:tableSlicerCache tableId="2" column="3"/></ext></extLst></slicerCacheDefinition>'
)


def test_resize_refuses_to_drop_a_column_a_slicer_filters_on():
    wb = _with_workbook_part(
        "xl/slicerCaches/slicerCache1.xml",
        "application/vnd.ms-excel.slicerCache+xml",
        "http://schemas.microsoft.com/office/2007/relationships/slicerCache",
        _SLICER_CACHE,
    )
    table = wb["Table"].tables[0]
    before = read_pkg(wb)
    with pytest.raises(
        TableError,
        match=r"would drop column 'Quantity', which the slicer cache "
        r"/xl/slicerCaches/slicerCache1.xml filters on",
    ):
        table.resize("A1:B11", columns=["Item", "Price"])
    assert_preserved(before, read_pkg(wb))
    table.resize("B1:C11", columns=["Price", "Quantity"])
    assert table.columns == ["Price", "Quantity"]


def test_resize_refuses_to_drop_a_column_a_formula_names():
    wb = Workbook.open(TABLES)
    ws = wb["Table"]
    ws["E1"].formula = "SUM(Table2[Quantity])"
    before = read_pkg(wb)
    with pytest.raises(
        TableError, match=r"would drop column 'Quantity', which cell E1 of sheet 'Table'"
    ):
        ws.tables[0].resize("A1:B11", columns=["Item", "Price"])
    assert_preserved(before, read_pkg(wb))
    ws.tables[0].resize("B1:C11", columns=["Price", "Quantity"])
    assert ws.tables[0].columns == ["Price", "Quantity"]


# --- Table setters ----------------------------------------------------------------


def test_display_name_renames_a_table_no_formula_names():
    wb = Workbook.open(TABLES)
    table = wb["Table"].tables[0]
    table.display_name = "Sales"
    assert _identity(_saved_tables(wb)[TABLE]) == ("2", "Table2", "Sales")


@pytest.mark.parametrize(
    ("value", "message"),
    [
        ("A1", "looks like a cell reference"),
        ("1st", "must start with a letter"),
        ("Chart", "is already used by a table or defined name"),
        ("", "must be a non-empty string"),
    ],
)
def test_display_name_refuses_an_invalid_or_used_name(value, message):
    wb = Workbook.open(TABLES)
    _add_defined_name(wb, "chart")
    before = read_pkg(wb)
    with pytest.raises(TableError, match=message):
        wb["Table"].tables[0].display_name = value
    assert_preserved(before, read_pkg(wb))


def test_display_name_refuses_to_rename_a_table_a_formula_names():
    wb = Workbook.open(TABLES)
    wb["Table"]["E1"].formula = "ROWS(table2)"
    before = read_pkg(wb)
    with pytest.raises(
        TableError, match=r"would break cell E1 of sheet 'Table', whose formula 'ROWS\(table2\)'"
    ):
        wb["Table"].tables[0].display_name = "Sales"
    assert_preserved(before, read_pkg(wb))
    wb["Table"].tables[0].display_name = "TABLE2"  # only the case changes

    wb = Workbook.open(TABLES)
    _add_defined_name(wb, "Total")
    names = wb._workbook_part.element.find(_sml("definedNames"))
    names.find(f"{_sml('definedName')}[@name='Total']").text = "SUM(Table2[Price])"
    with pytest.raises(TableError, match="would break defined name 'Total'"):
        wb["Table"].tables[0].display_name = "Sales"


_PIVOT_CACHE = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    b'<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    b'recordCount="0"><cacheSource type="worksheet"><worksheetSource name="table2"/>'
    b'</cacheSource><cacheFields count="0"/></pivotCacheDefinition>'
)


def test_display_name_refuses_to_rename_a_table_a_pivot_cache_names():
    wb = _with_workbook_part(
        "xl/pivotCache/pivotCacheDefinition1.xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition",
        _PIVOT_CACHE,
    )
    before = read_pkg(wb)
    with pytest.raises(
        TableError, match=r"would break /xl/pivotCache/pivotCacheDefinition1.xml, which names it"
    ):
        wb["Table"].tables[0].display_name = "Sales"
    assert_preserved(before, read_pkg(wb))


def test_display_name_ignores_the_name_in_parts_that_cannot_name_a_table():
    wb = Workbook.open(TABLES)  # its theme names the font "Calibri"
    ws = wb.add_worksheet("New")
    table = ws.add_table("A1:B2", ["x", "y"], name="Calibri", write_header=True)
    table.display_name = "Renamed"
    assert table.display_name == "Renamed"


def test_display_name_reads_only_the_sheets_that_mention_the_name():
    wb = Workbook.open(TABLES, large=True)
    wb["Table"].tables[0].display_name = "Sales"
    assert wb["bar chart"]._part.unparsed_blob is not None
    assert _identity(_saved_tables(wb)[TABLE])[2] == "Sales"


def test_header_row_count_turns_the_header_into_data_and_back():
    wb = Workbook.open(TABLES)
    table = wb["Table"].tables[0]
    table.header_row_count = 0
    saved = _saved_tables(wb)[TABLE]
    assert saved.get("headerRowCount") == "0" and saved.find(_sml("autoFilter")) is None
    assert (table.ref, table.data_range, table.header_range) == ("A1:C11", "A1:C11", None)
    table.header_row_count = 1
    saved = _saved_tables(wb)[TABLE]
    assert saved.get("headerRowCount") is None and table.header_range == "A1:C1"


def test_header_row_count_refuses_a_first_row_without_the_column_names():
    wb = _edited_tables(
        (b'totalsRowShown="0"', b'headerRowCount="0" totalsRowShown="0"'), (_FIXTURE_FILTER, b"")
    )
    ws = wb["Table"]
    ws["B1"].value = "Cost"
    before = read_pkg(wb)
    with pytest.raises(
        TableError, match="B1 holds 'Cost', expected 'Price'; write them there first"
    ):
        ws.tables[0].header_row_count = 1
    with pytest.raises(TableError, match="must be 0 or 1, got 2"):
        ws.tables[0].header_row_count = 2
    with pytest.raises(TypeError):
        ws.tables[0].header_row_count = True
    assert_preserved(before, read_pkg(wb))


def test_header_row_count_moves_the_sort_state_with_the_first_data_row():
    sort_state = b'<sortState ref="A1:C11"><sortCondition ref="B1:B11"/></sortState>'
    wb = _edited_tables(
        (b'totalsRowShown="0"', b'headerRowCount="0" totalsRowShown="0"'),
        (_FIXTURE_FILTER, sort_state),
    )
    wb["Table"].tables[0].header_row_count = 1
    saved = _saved_tables(wb)[TABLE].find(_sml("sortState"))
    assert (saved.get("ref"), saved[0].get("ref")) == ("A2:C11", "B2:B11")


def test_totals_row_count_turns_the_last_data_row_into_the_totals_row():
    wb = Workbook.open(TABLES)
    table = wb["Table"].tables[0]
    table.totals_row_count = 1
    saved = _saved_tables(wb)[TABLE]
    assert (saved.get("totalsRowCount"), saved.get("totalsRowShown")) == ("1", None)
    assert saved.find(_sml("autoFilter")).get("ref") == "A1:C10"
    assert (table.totals_range, table.totals_row_shown) == ("A11:C11", True)
    with pytest.raises(TableError, match="shows its totals row; set totals_row_count to 0 first"):
        table.totals_row_shown = False
    table.totals_row_count = 0
    table.totals_row_shown = False
    assert _summary(_saved_tables(wb)[TABLE]) == _fixture_table()


def test_totals_row_count_moves_the_sort_state_with_the_last_data_row():
    wb = _rich()
    wb["Table"].tables[0].totals_row_count = 0
    af = _saved_tables(wb)[TABLE].find(_sml("autoFilter"))
    assert (af.get("ref"), af.find(_sml("sortState")).get("ref")) == ("A1:C11", "A2:C11")


def test_totals_row_count_refuses_to_leave_no_data_row():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    table = ws.add_table("A1:B2", ["x", "y"], write_header=True)
    before = read_pkg(wb)
    with pytest.raises(
        TableError,
        match="setting totals_row_count of table 'Table1' to 1 would leave it without a data row",
    ):
        table.totals_row_count = 1
    assert_preserved(before, read_pkg(wb))


def test_style_writes_table_style_info():
    wb = Workbook.open(TABLES)
    table = wb["Table"].tables[0]
    table.style = TableStyle("TableStyleLight9", show_first_column=True, show_row_stripes=False)
    info = _saved_tables(wb)[TABLE].find(_sml("tableStyleInfo"))
    assert dict(info.attrib) == {
        "name": "TableStyleLight9",
        "showFirstColumn": "1",
        "showLastColumn": "0",
        "showRowStripes": "0",
        "showColumnStripes": "0",
    }
    table.style = None
    assert _saved_tables(wb)[TABLE].find(_sml("tableStyleInfo")) is None
    table.style = TableStyle(name=None)
    assert table.style == TableStyle(name=None)
    assert_valid_package(wb)


def test_style_refuses_an_unknown_style_name():
    wb = Workbook.open(TABLES)
    before = read_pkg(wb)
    with pytest.raises(TableError, match="'TableStyleMedium29' is neither a built-in table style"):
        wb["Table"].tables[0].style = TableStyle("TableStyleMedium29")
    with pytest.raises(TypeError):
        wb["Table"].tables[0].style = "TableStyleMedium2"
    assert_preserved(before, read_pkg(wb))


def test_dxf_ids_must_name_a_dxf_of_the_stylesheet():
    wb = Workbook.open(TABLES)
    table = wb["Table"].tables[0]
    before = read_pkg(wb)
    with pytest.raises(TableError, match="dxf id 0 is not in styles.xml, which has 0 dxf entries"):
        table.header_row_dxf_id = 0
    with pytest.raises(TypeError):
        table.header_row_dxf_id = True
    assert_preserved(before, read_pkg(wb))
    dxf = wb.styles.ensure_dxf(bold=True)
    table.header_row_dxf_id = dxf
    table.table_border_dxf_id = dxf
    saved = _saved_tables(wb)[TABLE]
    assert (saved.get("headerRowDxfId"), saved.get("tableBorderDxfId")) == ("0", "0")
    table.table_border_dxf_id = None
    assert _saved_tables(wb)[TABLE].get("tableBorderDxfId") is None


# --- tables under row and column inserts -----------------------------------------


@pytest.mark.parametrize("at_row", [2, 5, 11], ids=["first-data-row", "inside", "totals-row"])
def test_insert_rows_grows_the_data_rows_above_the_totals_row(at_row):
    wb = _rich()
    wb["Table"].insert_rows([[None]], at_row=at_row)
    table = _saved_tables(wb)[TABLE]
    af = table.find(_sml("autoFilter"))
    sort_state = af.find(_sml("sortState"))
    assert (table.get("ref"), af.get("ref")) == ("A1:C12", "A1:C11")
    assert (sort_state.get("ref"), sort_state[0].get("ref")) == ("A2:C11", "B2:B11")


def test_inserts_refuse_a_malformed_totals_row_count_before_anything_moves():
    wb = _edited_tables((b'totalsRowShown="0"', b'totalsRowCount="one"'))
    ws = wb["Table"]
    before = read_pkg(wb)
    with pytest.raises(TableError, match="malformed totalsRowCount 'one'"):
        ws.insert_rows([[None]], at_row=5)
    with pytest.raises(TableError, match="malformed totalsRowCount 'one'"):
        ws.insert_columns([[None]], at_col="B")
    assert_preserved(before, read_pkg(wb))


def test_insert_columns_moves_the_sort_state_with_its_columns():
    wb = _rich()
    wb["Table"].insert_columns([["Total"]], at_col="B")
    sort_state = _saved_tables(wb)[TABLE].find(f"{_sml('autoFilter')}/{_sml('sortState')}")
    assert (sort_state.get("ref"), sort_state[0].get("ref")) == ("A2:D10", "C2:C10")


def test_insert_columns_names_a_joining_column_after_its_header_value():
    wb = Workbook.open(TABLES)
    ws = wb["Table"]
    ws.insert_columns([["Total", 1.5], [None, 2]], at_col="B")
    names = [name for _, name in _summary(_saved_tables(wb)[TABLE])["columns"]]
    assert names == ["Item", "Total", "Column1", "Price", "Quantity"]
    assert [ws[a].value for a in ("A1", "B1", "C1", "D1")] == ["Item", "Total", "Column1", "Price"]


def test_insert_columns_writes_generated_names_into_a_header_row_it_does_not_reach():
    wb = Workbook.open(TABLES)
    ws = wb["Table"]
    ws.insert_columns([[1, 2]], at_cell="B5")
    assert (ws["B1"].value, ws["B5"].value) == ("Column1", 1)
    assert ws.tables[0].columns == ["Item", "Column1", "Price", "Quantity"]


def test_insert_columns_names_a_joining_column_in_large_mode():
    wb = Workbook.open(TABLES, large=True)
    wb["Table"].insert_columns([[None]], at_col="C")
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    ws = Workbook.open(buf)["Table"]
    assert [ws[a].value for a in ("B1", "C1", "D1")] == ["Price", "Column1", "Quantity"]


@pytest.mark.parametrize(
    "cols",
    [[["price"]], [[""]], [[5]], [["Total"], ["TOTAL"]]],
    ids=["used-name", "empty-string", "number", "repeated"],
)
def test_insert_columns_refuses_a_header_value_that_cannot_name_the_column(cols):
    wb = Workbook.open(TABLES)
    before = read_pkg(wb)
    with pytest.raises(TableError, match=r"joining table 'Table2' would hold"):
        wb["Table"].insert_columns(cols, at_col="B")
    assert_preserved(before, read_pkg(wb))


def test_insert_columns_into_a_table_without_a_header_row_writes_no_name():
    wb = _edited_tables(
        (b'totalsRowShown="0"', b'headerRowCount="0" totalsRowShown="0"'), (_FIXTURE_FILTER, b"")
    )
    ws = wb["Table"]
    ws.insert_columns([[7]], at_col="B")
    assert ws["B1"].value == 7
    assert ws.tables[0].columns == ["Item", "Column1", "Price", "Quantity"]


def test_write_dataframe_refuses_to_resize_a_table_with_a_totals_row():
    wb = _rich()
    before = read_pkg(wb)
    with pytest.raises(TableError, match="table 'Table2' has a totals row"):
        wb.write_dataframe([("Pen", 1.5, 4)], sheet="Table", at_cell="A2", header=False)
    assert_preserved(before, read_pkg(wb))
    wb.write_dataframe(
        [("Pen", 1.5, 4)], sheet="Table", at_cell="A2", header=False, resize_table=False
    )
    assert _saved_tables(wb)[TABLE].get("ref") == "A1:C11"
