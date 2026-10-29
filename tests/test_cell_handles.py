"""``ws[address]`` reads a cell without creating it (#44).

A cell the sheet has no ``<c>`` for reads as empty. Its first write creates
the row and the cell in order, in the style Excel gives a cell typed there,
and a write that is refused or writes nothing leaves the sheet as it was.
"""

from __future__ import annotations

import io

import pytest
from lxml import etree

from xlsxedit import FormulaGroupError, GridOverflowError, InvalidColorError, Workbook
from xlsxedit.opc.constants import SML_NS
from tests.conftest import BOOK1
from tests.preservation import assert_preserved, check_consistency, read_pkg

SHEET = "Sheet1"
_ROW = f"{{{SML_NS}}}row"
_C = f"{{{SML_NS}}}c"


def _layout(ws) -> list[tuple[int, list[str]]]:
    sheet_data = ws._part.element.find(f"{{{SML_NS}}}sheetData")
    return [
        (int(row.get("r")), [c.get("r") for c in row.iterchildren(_C)])
        for row in sheet_data.iterchildren(_ROW)
    ]


def _saved(wb: Workbook) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _styled_book() -> tuple[Workbook, int, int]:
    """A saved and reopened workbook whose column D has a green, percent style
    and whose row 5 a yellow one set with ``customFormat``."""
    wb = Workbook.create()
    ws = wb[SHEET]
    styles = wb.styles
    column = styles.allocate_cell_style(bg_color="FF00FF00", num_format="0.00%")
    row = styles.allocate_cell_style(bg_color="FFFFFF00")
    ws.column_dimensions["D"].width = 12
    ws._part.element.find(f"{{{SML_NS}}}cols")[0].set("style", str(column))
    row_elm = ws._ensure_row(5)
    row_elm.set("s", str(row))
    row_elm.set("customFormat", "1")
    ws["A1"].value = "title"
    return Workbook.open(io.BytesIO(_saved(wb))), column, row


def test_reading_a_missing_cell_changes_nothing():
    """Recorded in Excel 16.0.20527.20088 on 2026-10-29: an empty ``<c>`` without
    ``s`` in a column or row that has a style shows in style 0, not in that
    style, and Excel saves it back as it is. A cell that a read created would
    therefore have dropped its column's or row's formatting in Excel."""
    wb, column, row = _styled_book()
    source = _saved(wb)
    ws = wb[SHEET]
    layout = _layout(ws)

    for address, xf in (("D3", column), ("C5", row), ("D5", row), ("ZZ999", 0)):
        cell = ws[address]
        assert cell.address == address
        assert (cell.value, cell.formula, cell.formula_type, cell.data_type) == (None,) * 4
        assert (cell.style_index, cell.has_formula, cell.is_string) == (None, False, False)
        assert cell.style.style_index == str(xf)
        assert cell.hyperlink.url is None
        assert cell.offset(cols=1, rows=1).value is None
    assert ws["d3"].style.num_format == "0.00%"
    assert _layout(ws) == layout
    assert not ws._part.is_dirty
    assert read_pkg(_saved(wb)) == read_pkg(source)


def test_a_missing_cell_reads_what_any_write_puts_there():
    wb = Workbook.create()
    ws = wb[SHEET]
    held = ws["C3"]
    assert held.value is None
    ws["C3"].value = 5
    assert held.value == 5
    ws.write_rows([[6]], at_cell="C3")
    assert held.value == 6
    held.value = 7
    assert ws["C3"].value == 7
    assert _layout(ws) == [(3, ["C3"])]


def test_first_writes_create_rows_and_cells_in_order():
    wb = Workbook.create()
    ws = wb[SHEET]
    held = [ws[address] for address in ("C5", "A5", "B2", "AA5", "B9")]
    for n, cell in enumerate(held):
        cell.value = n
    ws["Z5"].formula = "A5+1"
    ws.write_rows([["bulk"]], at_cell="B5")
    ws["B5"].apply_style(bold=True)
    ws["C9"].apply_number_format("0.0")

    assert _layout(ws) == [(2, ["B2"]), (5, ["A5", "B5", "C5", "Z5", "AA5"]), (9, ["B9", "C9"])]
    assert [cell.value for cell in held] == [0, 1, 2, 3, 4]
    assert ws["B5"].value == "bulk" and ws["B5"].style.bold
    assert ws["Z5"].formula == "A5+1"
    assert ws._part.is_dirty
    check_consistency(Workbook.open(io.BytesIO(_saved(wb))))


def test_a_first_write_takes_the_style_excel_gives_a_typed_cell():
    """Recorded in Excel 16.0.20527.20088 on 2026-10-29: typing into a cell with
    no ``<c>`` gives it its row's style when the row sets ``customFormat``,
    else its column's; typing into an empty ``<c>`` without ``s`` keeps it without."""
    wb, column, row = _styled_book()
    ws = wb[SHEET]
    ws._get_or_create_cell_element("D7")

    ws["D3"].value = 7
    ws["C5"].value = 8
    ws["D5"].value = 9
    ws["B2"].value = 1
    ws["D7"].value = 10
    ws["D8"].apply_style(bold=True)
    ws["D9"].style.num_format = "0.0"

    assert [ws[a].style_index for a in ("D3", "C5", "D5", "B2", "D7")] == [
        str(column), str(row), str(row), None, None
    ]
    assert ws["D8"].style.bold and ws["D8"].style.bg_color == "FF00FF00"
    assert ws["D8"].style.num_format == "0.00%"
    assert ws["D9"].style.num_format == "0.0" and ws["D9"].style.bg_color == "FF00FF00"
    assert ws["D9"].value is None and ws._find_cell_element("D9") is not None


_NOTHING_WRITTEN = {
    "None": lambda cell: setattr(cell, "value", None),
    "no formula": lambda cell: setattr(cell, "formula", None),
    "clear": lambda cell: cell.clear(),
    "replace": lambda cell: cell.replace("a", "b"),
    "no style": lambda cell: cell.apply_style(),
}


@pytest.mark.parametrize("write", _NOTHING_WRITTEN.values(), ids=_NOTHING_WRITTEN)
def test_a_write_of_nothing_creates_no_cell(write):
    source = read_pkg(BOOK1)
    wb = Workbook.open(BOOK1)
    ws = wb[SHEET]
    write(ws["ZZ999"])
    assert ws._find_row_element(999) is None
    assert not ws._part.is_dirty
    assert read_pkg(_saved(wb)) == source


def _array_book() -> bytes:
    """An array formula over D2:E3 whose cells E2 and E3 have no ``<c>``."""
    wb = Workbook.create()
    ws = wb[SHEET]
    f = etree.SubElement(ws._get_or_create_cell_element("D2"), f"{{{SML_NS}}}f")
    f.attrib.update({"t": "array", "ref": "D2:E3"})
    f.text = "A2:B3*2"
    ws["A1"].value = "title"
    return _saved(wb)


_REFUSED = {
    "text XML cannot hold": (lambda ws: setattr(ws["B7"], "value", "a\x01"), ValueError),
    "unsupported type": (lambda ws: setattr(ws["B7"], "value", object()), TypeError),
    "empty formula": (lambda ws: setattr(ws["B7"], "formula", " "), ValueError),
    "formula in an array range": (lambda ws: setattr(ws["E3"], "formula", "1"), FormulaGroupError),
    "bad colour": (lambda ws: ws["B7"].apply_style(bg_color="nope"), InvalidColorError),
    "bad number format": (lambda ws: ws["B7"].apply_number_format(5), TypeError),
    "past the grid": (lambda ws: ws.write_rows([[1, 2]], at_cell="XFD7"), GridOverflowError),
}


@pytest.mark.parametrize("case", _REFUSED)
def test_a_refused_write_leaves_no_cell_behind(case: str):
    write, error = _REFUSED[case]
    source = _array_book()
    wb = Workbook.open(io.BytesIO(source))
    ws = wb[SHEET]
    with pytest.raises(error):
        write(ws)
    assert ws._find_cell_element("E3") is None
    assert ws._find_row_element(7) is None
    assert not ws._part.is_dirty
    assert_preserved(read_pkg(source), read_pkg(_saved(wb)))


def _index_state(ws) -> tuple:
    return (
        dict(ws._row_index),
        dict(ws._cell_index),
        set(ws._cell_indexed_rows),
        list(ws._row_numbers or ()),
        ws._bulk_max_row_num,
    )


def test_writes_that_leave_nothing_restore_the_row_index(monkeypatch):
    from tests.test_row_placement import _Walks

    wb = Workbook.create()
    ws = wb[SHEET]
    for n in range(10, 1001, 10):
        ws[f"A{n}"].value = n
    ws["B25"].value = 25  # placed by bisecting the sorted row numbers
    assert ws._row_numbers is not None
    state = _index_state(ws)
    layout = _layout(ws)

    for address in ("B35", "B5000", "C20", "D25"):
        ws[address].value = None
        ws[address].apply_style()
        with pytest.raises(ValueError):
            ws[address].value = "a\x01"
    assert _index_state(ws) == state
    assert _layout(ws) == layout

    walks = _Walks(monkeypatch)
    for n, address in enumerate(("B35", "B5000", "C20", "D25", "B36", "B4999")):
        ws[address].value = n
    assert walks.counts["scan"] == 0 and walks.counts["index"] == 0
    rows = dict(_layout(ws))
    assert list(rows) == sorted(rows)
    assert (rows[20], rows[25], rows[35], rows[36], rows[4999], rows[5000]) == (
        ["A20", "C20"], ["B25", "D25"], ["B35"], ["B36"], ["B4999"], ["B5000"]
    )
    check_consistency(Workbook.open(io.BytesIO(_saved(wb))))


def _values(ws) -> dict[str, object]:
    return {cell.address: cell.value for cell in ws.iter_cells()}


@pytest.mark.parametrize("read_first", [False, True], ids=["unread", "read"])
def test_a_held_missing_cell_moves_with_inserted_rows_and_columns(read_first: bool):
    wb = Workbook.create()
    ws = wb[SHEET]
    ws["A5"].value = "keep"
    ws["C1"].value = "x"
    below, right, filled = ws["A6"], ws["D1"], ws["B5"]
    ws["B5"].value = 1  # through another handle, before ``filled`` binds
    if read_first:
        assert filled.value == 1
    ws.insert_rows([["new"]], at_cell="A2")
    ws.insert_columns([["new"]], at_cell="B1")

    assert (below.address, right.address, filled.address) == ("A7", "E1", "C6")
    assert filled.value == 1
    below.value = "below"
    right.value = "right"
    assert _values(ws) == {
        "A2": "new", "B1": "new", "D1": "x", "E1": "right", "A6": "keep", "C6": 1, "A7": "below",
    }
    check_consistency(Workbook.open(io.BytesIO(_saved(wb))))


def test_a_held_missing_cell_pushed_off_the_grid_raises_on_use():
    wb = Workbook.create()
    ws = wb[SHEET]
    held = ws["A1048576"]
    ws.insert_rows([["new"]], at_row=1)
    assert repr(held) == "<Cell 'A1048576'>"
    for use in (lambda: held.value, lambda: held.address, lambda: setattr(held, "value", 1)):
        with pytest.raises(GridOverflowError, match="moved cell A1048576 past the grid"):
            use()
    assert ws._find_row_element(1048576) is None


def _fresh_style(wb: Workbook, address: str) -> int:
    """The style a cache-free copy of ``wb`` gives ``address``."""
    copy = Workbook.open(io.BytesIO(_saved(wb)))
    return copy.styles.effective_xf_index(copy[SHEET], address)


def test_new_cells_take_column_styles_as_the_columns_change():
    wb = Workbook.create()
    ws = wb[SHEET]
    green = wb.styles.allocate_cell_style(bg_color="FF00FF00")
    ws.column_dimensions["B"].width = 10
    cols = ws._part.element.find(f"{{{SML_NS}}}cols")
    etree.SubElement(cols, f"{{{SML_NS}}}col", min="3", max="5", width="9", style=str(green))
    ws["C1"].value = 1
    cache = ws._column_styles
    for row in range(2, 50):
        ws[f"D{row}"].value = row
    assert ws._column_styles is cache  # read once for all of them
    assert ws["D49"].style_index == str(green)

    def write(address: str) -> None:
        expected = _fresh_style(wb, address)
        ws[address].value = 1
        assert int(ws[address].style_index or 0) == expected, address

    etree.SubElement(cols, f"{{{SML_NS}}}col", min="8", max="8", width="9", style=str(green))
    write("H1")
    ws.insert_columns([["new"]], at_col="A")  # C:E moves to D:F, H to I
    write("F2")
    write("I2")
    ws.column_dimensions["E"].width = 20  # rewrites the <col> covering E
    write("E3")
    write("G3")
    assert (ws["I1"].style_index, ws["I2"].style_index) == (str(green), str(green))
