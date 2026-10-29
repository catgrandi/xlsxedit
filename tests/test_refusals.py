"""A refused cell write, bulk write or sheet copy changes no package member."""

from __future__ import annotations

import io
from datetime import datetime, timezone
from pathlib import Path

import pytest

from xlsxedit import FormulaGroupError, GridOverflowError, InvalidColorError, InvalidRangeError, Workbook
from xlsxedit.bulk import write_cell_value
from tests.conftest import INSPECT_FIXTURES
from tests.preservation import assert_preserved, read_pkg
from tests.test_formula_hygiene import RANGE_FORMULAS, _plant

SIMPLE = INSPECT_FIXTURES["SimpleFormula"]
CHARTS = INSPECT_FIXTURES["ChartsAndTables"]
NUL = chr(0)
_SML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

BAD_VALUES = {
    "unsupported type": (object(), TypeError),
    "text XML cannot hold": ("a" + NUL, ValueError),
    "aware datetime": (datetime(2024, 1, 1, tzinfo=timezone.utc), TypeError),
    "integer too long for str()": (10**5000, ValueError),
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


def test_strings_added_to_a_large_workbook_keep_their_indexes():
    wb = Workbook.open(INSPECT_FIXTURES["TestBook"], large=True)
    ws = wb.worksheets[0]
    ws["Z50"].value = "first new"
    ws["Z51"].value = "second new"
    ws.write_rows([["third new", "fourth new"]], at_cell="Z60")

    reopened = Workbook.open(io.BytesIO(_saved(wb))).worksheets[0]
    assert [reopened[a].value for a in ("Z50", "Z51", "Z60", "AA60")] == [
        "first new",
        "second new",
        "third new",
        "fourth new",
    ]


@pytest.mark.parametrize("count", ["", "1e3", "-"])
def test_a_shared_strings_count_that_is_no_number_is_rewritten(count: str):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "kept"
    wb.shared_strings._element.set("count", count)

    ws["A2"].value = "added"
    assert (ws["A1"].value, ws["A2"].value) == ("kept", "added")
    assert wb.shared_strings._element.get("count") == "2"


@pytest.mark.parametrize("index", [" 0 ", "00", "+0"])
def test_string_over_an_index_in_another_form_leaves_the_shared_string_alone(index: str):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "kept"
    _plant(ws["B1"], {"t": "s"}, f"<v>{index}</v>")

    ws["B1"].value = "new"
    assert (ws["A1"].value, ws["B1"].value) == ("kept", "new")


def test_a_refused_style_builds_nothing_in_the_stylesheet():
    wb = Workbook.create()
    before = read_pkg(_saved(wb))
    with pytest.raises(InvalidColorError):
        wb["Sheet1"]["B2"].apply_style(bold=True, bg_color="nope")
    with pytest.raises(InvalidColorError):
        wb.styles.allocate_cell_style(num_format="0.00", font_color="nope")
    assert read_pkg(_saved(wb)) == before


@pytest.mark.parametrize("link", [None, "https://example.com/original"])
@pytest.mark.parametrize("field", ["location", "url"])
def test_a_refused_hyperlink_keeps_the_cell_s_link(field, link):
    wb = Workbook.create()
    cell = wb["Sheet1"]["C3"]
    if link is not None:
        cell.hyperlink.url = link
    before = read_pkg(_saved(wb))
    with pytest.raises(ValueError):
        setattr(cell.hyperlink, field, "Sheet1!A1" + chr(1))
    assert read_pkg(_saved(wb)) == before
    assert cell.hyperlink.url == link


def test_numpy_scalars_are_written_as_numbers():
    np = pytest.importorskip("numpy")
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = np.float64(2.5)
    ws.write_rows([[np.float64(1.5), np.float64(0.5)]], at_cell="A2")
    assert [ws[a]._v_text() for a in ("A1", "A2", "B2")] == ["2.5", "1.5", "0.5"]


def _saved(wb) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


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
    "integer too long for str() in a later row": (
        lambda ws: ws.write_rows([[1, 2], [10**5000]], at_cell="H1"),
        ValueError,
    ),
    "template_rows that are not rows, with clear_range": (
        lambda ws: ws._workbook.write_dataframe(
            [("a",)], at_cell="H1", header=False, clear_range="A3:B4", template_rows=2.0
        ),
        TypeError,
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


def test_bulk_write_on_a_sheet_with_whole_column_and_row_merges():
    from lxml import etree

    from xlsxedit.opc.constants import SML_NS

    wb, ws = _array_sheet()
    block = ws._part.element.find(f"{{{SML_NS}}}mergeCells")
    for ref in ("J:K", "20:20"):
        etree.SubElement(block, f"{{{SML_NS}}}mergeCell").set("ref", ref)
    ws._invalidate_merge_map()

    assert ws.write_rows([[1, 2]], at_row=5) == 1
    with pytest.raises(FormulaGroupError):
        ws.write_rows([["x", "y"]], at_cell="F3")


@pytest.mark.parametrize("extra_merges", [0, 5], ids=["merge by merge", "cell by cell"])
def test_value_a_merge_sends_to_the_anchor_is_refused(extra_merges: int):
    wb, ws = _array_sheet()
    for row in range(10, 10 + extra_merges):
        ws.merge_cells(f"A{row}:B{row}")
    before = read_pkg(wb)
    with pytest.raises(FormulaGroupError):
        ws.write_rows([["x"]], at_cell="F3")
    assert_preserved(before, read_pkg(wb))
    assert ws.write_rows([[None, "y"]], at_cell="F3") == 1  # None is not sent on


def test_rows_without_values_write_no_cells():
    wb, ws = _array_sheet()
    dimension = ws._part.element.find(f"{{{_SML}}}dimension").get("ref")
    assert ws.write_rows([[]], at_row=5) == 1
    assert ws.write_rows([[], []], at_cell="C5") == 2
    assert ws._part.element.find(f"{{{_SML}}}dimension").get("ref") == dimension


def test_bulk_write_beside_the_anchor_still_writes():
    wb, ws = _array_sheet()
    assert ws.write_rows([["x", None, None, "y"]], at_cell="D2") == 1
    assert (ws["D2"].value, ws["G2"].value) == ("x", "y")
    assert ws["E3"].formula_type == "array"


# --- copy_worksheet ----------------------------------------------------------------------


def test_copy_worksheet_that_fails_while_cloning_changes_nothing(monkeypatch):
    import xlsxedit.sheet_clone as sheet_clone

    wb = Workbook.open(CHARTS)
    before = read_pkg(wb)
    clone_part = sheet_clone._clone_part
    cloned = []

    def fail_on_the_second_part(package, source, *args):
        if cloned:
            raise RuntimeError("injected")
        cloned.append(source)
        return clone_part(package, source, *args)

    monkeypatch.setattr(sheet_clone, "_clone_part", fail_on_the_second_part)
    with pytest.raises(RuntimeError, match="injected"):
        wb.copy_worksheet("bar chart", "Copy")
    assert_preserved(before, read_pkg(wb))

    monkeypatch.setattr(sheet_clone, "_clone_part", clone_part)
    copy = wb.copy_worksheet("bar chart", "Copy")
    assert copy._part.partname == "/xl/worksheets/sheet4.xml"
    assert len(copy.charts) == 1


@pytest.mark.parametrize(
    "corrupt",
    [
        lambda wb: wb._workbook_part.sheet_elements[0].set("sheetId", "x"),
        lambda wb: _add_defined_name(wb, "x"),
    ],
    ids=["sheetId", "localSheetId"],
)
def test_copy_worksheet_refused_by_workbook_xml_changes_nothing(corrupt):
    wb = Workbook.open(CHARTS)
    corrupt(wb)
    before = read_pkg(wb)
    with pytest.raises(ValueError):
        wb.copy_worksheet("bar chart", "Copy")
    assert_preserved(before, read_pkg(wb))


def _add_defined_name(wb, local_sheet_id: str) -> None:
    from lxml import etree

    from xlsxedit.opc.constants import SML_NS

    wb_elm = wb._workbook_part.element
    block = wb_elm.find(f"{{{SML_NS}}}definedNames")
    if block is None:
        block = etree.Element(f"{{{SML_NS}}}definedNames")
        wb_elm.find(f"{{{SML_NS}}}sheets").addnext(block)
    for local in ("0", local_sheet_id):
        name = etree.SubElement(block, f"{{{SML_NS}}}definedName")
        name.set("name", f"Local{local}")
        name.set("localSheetId", local)
        name.text = "'bar chart'!$A$1"
