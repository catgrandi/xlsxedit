"""Text written into a workbook without a shared-strings part adds the part.

Excel writes no ``xl/sharedStrings.xml`` for a workbook whose cells hold no
text, as for ``DrawingAnchors.xlsx``. The first string an API writes there adds
the part, its workbook relationship and its content type; a refused write adds
nothing.
"""

from __future__ import annotations

import io
import shutil
import zipfile
from collections.abc import Callable
from datetime import datetime, timezone
from pathlib import Path

import pytest

from xlsxedit import FormulaGroupError, GridOverflowError, TableError, Workbook
from xlsxedit.bulk import write_cell_value
from xlsxedit.opc.constants import CT, RT
from tests.conftest import INSPECT_FIXTURES
from tests.preservation import assert_preserved, content_types, read_pkg
from tests.schema_validate import assert_valid_package
from tests.test_formula_hygiene import RANGE_FORMULAS, _plant

BOOK = INSPECT_FIXTURES["DrawingAnchors"]
SHEET = "Anchors"
SHEET_XML = "xl/worksheets/sheet1.xml"
SHEET_RELS = "xl/worksheets/_rels/sheet1.xml.rels"
SST = "xl/sharedStrings.xml"
WORKBOOK_XML = "xl/workbook.xml"
WORKBOOK_RELS = "xl/_rels/workbook.xml.rels"
TABLE_XML = "xl/tables/table1.xml"


def _reopen(source) -> Workbook:
    return Workbook.open(io.BytesIO(source if isinstance(source, bytes) else _saved(source)))


def _saved(wb: Workbook) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _sst_members(pkg: dict[str, bytes]) -> list[str]:
    return sorted(m for m, ct in content_types(pkg).items() if ct == CT.SHARED_STRINGS)


def test_the_fixture_has_no_shared_strings_part():
    assert _sst_members(read_pkg(BOOK)) == []
    assert Workbook.open(BOOK).shared_strings is None


# --- every API that writes a cell value -----------------------------------------------


def _in_memory(edit: Callable) -> Callable[[Path, object], None]:
    def write(path: Path, value) -> None:
        wb = Workbook.open(path)
        edit(wb[SHEET], value)
        wb.save(path)

    return write


def _pandas(**engine_kwargs) -> Callable[[Path, object], None]:
    def write(path: Path, value) -> None:
        pd = pytest.importorskip("pandas")
        from xlsxedit.pandas_io import register

        register()
        frame = pd.DataFrame([[value, 1]])
        with pd.ExcelWriter(
            path, engine="xlsxedit", mode="a", if_sheet_exists="overlay", engine_kwargs=engine_kwargs
        ) as writer:
            frame.to_excel(writer, sheet_name=SHEET, startrow=1, index=False, header=False)

    return write


# name -> (write ``value`` into a copy of BOOK at ``path``, the cell it lands in)
WRITES: dict[str, tuple[Callable[[Path, object], None], str]] = {
    "Cell.value": (_in_memory(lambda ws, v: setattr(ws["A1"], "value", v)), "A1"),
    "write_rows": (_in_memory(lambda ws, v: ws.write_rows([[1, v]], at_cell="A1")), "B1"),
    "write_dataframe": (
        _in_memory(lambda ws, v: ws._workbook.write_dataframe([(v, 1)], at_cell="A2", header=False)),
        "A2",
    ),
    "insert_rows": (_in_memory(lambda ws, v: ws.insert_rows([[v, 1]], at_row=2)), "A2"),
    "insert_columns": (_in_memory(lambda ws, v: ws.insert_columns([[1, v]], at_col="B")), "B2"),
    "pandas engine": (_pandas(), "A2"),
    "pandas engine, insert mode": (_pandas(mode="insert"), "A2"),
}


def _written(tmp_path: Path, write: str, value) -> Path:
    path = tmp_path / f"{type(value).__name__}.xlsx"
    shutil.copy(BOOK, path)
    WRITES[write][0](path, value)
    return path


@pytest.mark.parametrize("write", WRITES)
def test_text_is_read_back_after_a_save_and_the_package_validates(tmp_path: Path, write: str):
    path = _written(tmp_path, write, "seven")
    assert Workbook.open(path)[SHEET][WRITES[write][1]].value == "seven"
    assert _sst_members(read_pkg(path)) == [SST]
    assert_valid_package(path)


@pytest.mark.parametrize("write", WRITES)
def test_text_changes_what_a_number_changes_and_adds_the_part(tmp_path: Path, write: str):
    numbers = _written(tmp_path, write, 7)
    text = _written(tmp_path, write, "seven")
    assert_preserved(
        numbers, text, expected_changed={SHEET_XML, WORKBOOK_RELS}, expected_added={SST}
    )


@pytest.mark.parametrize("large", [False, True], ids=["normal", "large"])
def test_a_cell_write_leaves_every_excel_written_member_as_it_was(large: bool):
    wb = Workbook.open(BOOK, large=large)
    wb[SHEET]["A1"].value = "text"
    report = assert_preserved(
        BOOK, wb, expected_changed={SHEET_XML, WORKBOOK_RELS}, expected_added={SST}
    )
    assert report.reserialised == {"[Content_Types].xml"}  # rebuilt with the part's Override
    assert _reopen(wb)[SHEET]["A1"].value == "text"


def test_string_columns_write_a_number_as_text():
    wb = Workbook.open(BOOK)
    wb[SHEET].write_rows([[7]], at_cell="A1", string_columns={0})
    assert _reopen(wb)[SHEET]["A1"].value == "7"


def test_the_part_is_added_as_excel_adds_it():
    wb = Workbook.open(BOOK)
    wb[SHEET]["A1"].value = " padded "
    pkg = read_pkg(wb)
    assert content_types(pkg)[SST] == CT.SHARED_STRINGS
    rels = zipfile.ZipFile(io.BytesIO(_saved(wb))).read(WORKBOOK_RELS)
    assert f'Type="{RT.SHARED_STRINGS}" Target="sharedStrings.xml"'.encode() in rels
    assert pkg[SST].endswith(
        b'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" '
        b'uniqueCount="1"><si><t xml:space="preserve"> padded </t></si></sst>'
    )


def test_every_write_shares_one_part():
    wb = Workbook.open(BOOK)
    ws = wb[SHEET]
    ws["A1"].value = "one"
    table = wb.shared_strings
    ws.write_rows([["two", "one"]], at_cell="A2")
    ws.insert_rows([["three"]], at_row=1)
    assert wb.shared_strings is table
    assert len(wb._workbook_part.rels.parts_with_reltype(RT.SHARED_STRINGS)) == 1
    pkg = read_pkg(wb)
    assert _sst_members(pkg) == [SST]
    back = _reopen(wb)[SHEET]
    assert [back[a].value for a in ("A1", "A2", "A3", "B3")] == ["three", "one", "two", "one"]


def test_the_part_does_not_take_an_orphans_name(tmp_path: Path):
    path = tmp_path / "orphan.xlsx"
    shutil.copy(BOOK, path)
    with zipfile.ZipFile(path, "a") as zf:
        zf.writestr(SST, b"orphan cargo")
    wb = Workbook.open(path)
    assert wb.orphan_partnames == ("/xl/sharedStrings.xml",)
    wb[SHEET]["A1"].value = "text"
    out = tmp_path / "out.xlsx"
    wb.save(out, include_orphans=True)
    pkg = read_pkg(out)
    assert pkg[SST] == b"orphan cargo"
    assert _sst_members(pkg) == ["xl/sharedStrings1.xml"]
    assert Workbook.open(out)[SHEET]["A1"].value == "text"


# --- tables ----------------------------------------------------------------------------


def _book_with_a_table() -> bytes:
    """BOOK with table ``H1:I3`` whose header cells hold inline strings, so it still
    has no shared strings part."""
    wb = Workbook.open(BOOK)
    ws = wb[SHEET]
    for address, text in (("H1", "Name"), ("I1", "Qty")):
        _plant(ws[address], {"t": "inlineStr"}, f"<is><t>{text}</t></is>")
    ws.write_rows([[1, 2], [3, 4]], at_cell="H2")
    ws.add_table("H1:I3", ["Name", "Qty"])
    assert wb.shared_strings is None
    return _saved(wb)


TABLE_WRITES = {
    "add_table(write_header=True)": (
        lambda ws: ws.add_table("K1:L3", ["One", "Two"], write_header=True),
        {"K1": "One", "L1": "Two"},
        {SHEET_XML, SHEET_RELS},
        {"xl/tables/table2.xml"},
    ),
    "Table.resize": (
        lambda ws: ws.tables[0].resize("H1:J3", columns=["Name", "Qty", "Price"]),
        {"J1": "Price"},
        {SHEET_XML, TABLE_XML},
        set(),
    ),
    "insert_columns naming a joining column": (
        lambda ws: ws.insert_columns([[None, 5, 6]], at_col="I"),
        {"I1": "Column1"},
        {SHEET_XML, TABLE_XML, WORKBOOK_XML},
        set(),
    ),
}


@pytest.mark.parametrize("write", TABLE_WRITES)
def test_a_header_cell_a_table_api_writes_adds_the_part(write: str):
    edit, headers, changed, added = TABLE_WRITES[write]
    before = _book_with_a_table()
    wb = _reopen(before)
    edit(wb[SHEET])
    assert_preserved(
        before, wb, expected_changed={*changed, WORKBOOK_RELS}, expected_added={*added, SST}
    )
    saved = _saved(wb)
    assert_valid_package(saved)
    back = _reopen(saved)[SHEET]
    assert {address: back[address].value for address in headers} == headers


# --- search and replace ----------------------------------------------------------------


def _inline_text_book() -> Workbook:
    wb = Workbook.open(BOOK)
    _plant(wb[SHEET]["A1"], {"t": "inlineStr"}, "<is><t>{{name}} was here</t></is>")
    _plant(wb[SHEET]["A2"], {"t": "inlineStr"}, "<is><t>{{name}}</t></is>")
    return wb


def test_replace_in_inline_strings_needs_no_part():
    wb = _inline_text_book()
    assert wb.replace("{{name}}", "Ada") == 2
    assert wb.shared_strings is None
    back = _reopen(wb)[SHEET]
    assert (back["A1"].value, back["A2"].value) == ("Ada was here", "Ada")


def test_whole_cell_replace_with_text_adds_the_part():
    wb = _inline_text_book()
    assert wb.replace("{{name}}", "Ada", value_type="number") == 1
    pkg = read_pkg(wb)
    assert _sst_members(pkg) == [SST]
    back = _reopen(wb)[SHEET]
    assert (back["A1"].value, back["A2"].value) == ("{{name}} was here", "Ada")


# --- refused writes --------------------------------------------------------------------


def _array_anchor(ws) -> None:
    """An array formula over E3:E4, whose anchor no write may overwrite."""
    _plant(ws["E3"], {}, RANGE_FORMULAS["array"].replace("B1:B2", "E3:E4") + "<v>2</v>")
    _plant(ws["E4"], {}, "<v>4</v>")


def _write_dataframe(rows, **kwargs):
    return lambda ws: ws._workbook.write_dataframe(rows, header=False, **kwargs)


AWARE = datetime(2024, 1, 1, tzinfo=timezone.utc)
REFUSALS: dict[str, tuple[Callable, type[BaseException]]] = {
    "Cell.value over an array formula's anchor": (
        lambda ws: setattr(ws["E3"], "value", "text"),
        FormulaGroupError,
    ),
    "Cell.value with text XML cannot hold": (
        lambda ws: setattr(ws["A1"], "value", "a\x00"),
        ValueError,
    ),
    "write_rows with an unsupported value after text": (
        lambda ws: ws.write_rows([["text"], [object()]], at_cell="A1"),
        TypeError,
    ),
    "write_rows with text XML cannot hold after text": (
        lambda ws: ws.write_rows([["text", "b\x01"]], at_cell="A1"),
        ValueError,
    ),
    "write_rows over an array formula's anchor after text": (
        lambda ws: ws.write_rows([["text", "more"]], at_cell="D3"),
        FormulaGroupError,
    ),
    "write_dataframe with an aware datetime after text": (
        _write_dataframe([("text",), (AWARE,)], at_cell="A1"),
        TypeError,
    ),
    "write_dataframe with clear_range over an array formula's anchor": (
        _write_dataframe([("text",)], at_cell="A1", clear_range="E3:E3"),
        FormulaGroupError,
    ),
    "insert_rows with an unsupported value after text": (
        lambda ws: ws.insert_rows([["text", object()]], at_row=2),
        TypeError,
    ),
    "insert_rows past the last row": (
        lambda ws: ws.insert_rows([["text"], ["more"]], at_row=1048576),
        GridOverflowError,
    ),
    "insert_columns with an unsupported value after text": (
        lambda ws: ws.insert_columns([["text", object()]], at_col="B"),
        TypeError,
    ),
    "add_table(write_header=True) over an array formula's anchor": (
        lambda ws: ws.add_table("E3:F5", ["One", "Two"], write_header=True),
        FormulaGroupError,
    ),
    "add_table(write_header=True) with a name Excel rejects": (
        lambda ws: ws.add_table("K1:L3", ["One", "Two"], name="1st", write_header=True),
        TableError,
    ),
    "bulk write_cell_value with text XML cannot hold": (
        lambda ws: write_cell_value(
            ws._get_or_create_cell_element("A1"), "a\x00", None, {}, worksheet=ws
        ),
        ValueError,
    ),
}


@pytest.mark.parametrize("case", REFUSALS)
def test_a_refused_text_write_adds_no_part(case: str):
    write, error = REFUSALS[case]
    wb = Workbook.open(BOOK)
    _array_anchor(wb[SHEET])
    before = read_pkg(wb)
    with pytest.raises(error):
        write(wb[SHEET])
    assert wb.shared_strings is None
    assert_preserved(before, read_pkg(wb))


def test_a_refused_table_resize_adds_no_part():
    before = _book_with_a_table()
    wb = _reopen(before)
    with pytest.raises(TableError):
        wb[SHEET].tables[0].resize("H1:J1", columns=["Name", "Qty", "Price"])  # no data row
    assert wb.shared_strings is None
    assert_preserved(before, wb)


def test_a_refused_pandas_write_adds_no_part(tmp_path: Path):
    pd = pytest.importorskip("pandas")
    from xlsxedit.pandas_io import register

    register()
    wb = Workbook.open(BOOK)
    _array_anchor(wb[SHEET])
    path = tmp_path / "anchor.xlsx"
    wb.save(path)
    before = read_pkg(path)
    with pytest.raises(FormulaGroupError):
        with pd.ExcelWriter(
            path, engine="xlsxedit", mode="a", if_sheet_exists="overlay"
        ) as writer:
            pd.DataFrame([["text", "more"]]).to_excel(
                writer, sheet_name=SHEET, startrow=2, startcol=3, index=False, header=False
            )
    assert_preserved(before, path)
