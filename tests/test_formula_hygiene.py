"""``Cell.formula`` setter hygiene: stored form, stale caches, and formula groups."""

from __future__ import annotations

import io

import pytest
from lxml import etree

from xlsxedit import FormulaGroupError, Workbook, XlsxeditError
from xlsxedit.opc.constants import SML_NS
from tests.conftest import INSPECT_FIXTURES
from tests.preservation import (
    CellState,
    assert_preserved,
    check_consistency,
    read_pkg,
    worksheet_diff,
)

_C = f"{{{SML_NS}}}c"
_F = f"{{{SML_NS}}}f"
_V = f"{{{SML_NS}}}v"
_CALC_PR = f"{{{SML_NS}}}calcPr"
SHEET = "xl/worksheets/sheet1.xml"
SIMPLE = INSPECT_FIXTURES["SimpleFormula"]


def _cell(pkg: dict[str, bytes], address: str) -> etree._Element:
    return next(c for c in etree.fromstring(pkg[SHEET]).iter(_C) if c.get("r") == address)


def _plant(cell, attrs: dict[str, str], inner: str) -> None:
    """Replace a cell's type and children with raw XML, keeping ``r`` and ``s``."""
    c = cell.worksheet._get_or_create_cell_element(cell.address)
    for child in list(c):
        c.remove(child)
    for name in [n for n in c.attrib if n not in ("r", "s")]:
        del c.attrib[name]
    c.attrib.update(attrs)
    for child in etree.fromstring(f'<c xmlns="{SML_NS}">{inner}</c>'):
        c.append(child)


@pytest.mark.parametrize(("expr", "stored"), [("=A1*2", "A1*2"), ("SUM(A1:B1)", "SUM(A1:B1)")])
def test_formula_is_saved_without_a_leading_equals(expr: str, stored: str):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = 10
    ws["A2"].formula = expr
    assert ws["A2"].formula == stored

    buf = io.BytesIO()
    wb.save(buf)
    assert f'<c r="A2"><f>{stored}</f></c>'.encode() in read_pkg(buf.getvalue())[SHEET]
    buf.seek(0)
    assert Workbook.open(buf)["Sheet1"]["A2"].formula == stored


STALE = {
    "shared string": ({"t": "s"}, "<v>0</v>"),
    "inline string": ({"t": "inlineStr"}, "<is><t>text</t></is>"),
    "boolean": ({"t": "b"}, "<v>1</v>"),
    "error": ({"t": "e"}, "<v>#N/A</v>"),
    "number": ({}, "<v>5</v>"),
    "cached formula result": ({"t": "str"}, '<f>B1&amp;"x"</f><v>abx</v>'),
}


@pytest.mark.parametrize("before", STALE)
def test_formula_drops_the_stale_type_and_cached_value(before: str):
    wb = Workbook.create()
    cell = wb["Sheet1"]["A1"]
    _plant(cell, *STALE[before])

    cell.formula = "=B1*2"

    assert cell.data_type is None
    assert [child.tag for child in cell._element] == [_F]
    assert cell.formula == "B1*2"
    assert cell.formula_type == "normal"
    assert cell.value is None


def test_formula_sets_full_calc_on_load():
    wb = Workbook.create()

    def full_calc_on_load() -> str | None:
        calc_pr = wb._workbook_part.element.find(_CALC_PR)
        return None if calc_pr is None else calc_pr.get("fullCalcOnLoad")

    assert full_calc_on_load() is None
    wb["Sheet1"]["A1"].formula = "1+1"
    assert full_calc_on_load() == "1"


@pytest.mark.parametrize(
    ("address", "kind"), [("A3", None), ("C3", "normal"), ("C4", "shared"), ("C5", "shared")]
)
def test_formula_type_names_the_kind_of_formula(address: str, kind: str | None):
    assert Workbook.open(SIMPLE).worksheets[0][address].formula_type == kind


def test_formula_type_is_read_only():
    cell = Workbook.open(SIMPLE).worksheets[0]["C3"]
    with pytest.raises(AttributeError):
        cell.formula_type = "array"


def test_shared_follower_formula_reads_as_none():
    cell = Workbook.open(SIMPLE).worksheets[0]["C5"]
    assert cell.has_formula
    assert cell.formula is None


def test_formula_on_a_shared_follower_gives_it_a_formula_of_its_own():
    before = read_pkg(SIMPLE)
    wb = Workbook.open(SIMPLE)
    wb.worksheets[0]["C5"].formula = "=A5+B5"
    after = read_pkg(wb)

    check_consistency(after)
    assert_preserved(before, after, expected_changed={SHEET, "xl/workbook.xml"})
    diff = worksheet_diff(before[SHEET], after[SHEET])
    assert diff.changed == {
        "C5": (CellState(None, None, "", "684"), CellState(None, None, "A5+B5", None))
    }
    assert not (diff.added or diff.removed or diff.rows or diff.children or diff.root)
    assert dict(_cell(after, "C5").find(_F).attrib) == {}
    master = _cell(after, "C4").find(_F)
    assert dict(master.attrib) == {"t": "shared", "ref": "C4:C7", "si": "0"}
    assert master.text == "A4*B4"


def test_formula_on_a_shared_master_with_followers_is_refused():
    before = read_pkg(SIMPLE)
    wb = Workbook.open(SIMPLE)
    with pytest.raises(FormulaGroupError, match="C4 holds shared formula si=0, which 3 other"):
        wb.worksheets[0]["C4"].formula = "A4+B4"
    assert_preserved(before, read_pkg(wb))


def test_formula_on_a_shared_master_without_followers_replaces_it():
    wb = Workbook.open(SIMPLE)
    ws = wb.worksheets[0]
    for row in (5, 6, 7):
        ws[f"C{row}"].formula = f"A{row}*2"
    ws["C4"].formula = "A4*3"
    after = read_pkg(wb)

    check_consistency(after)
    master = _cell(after, "C4").find(_F)
    assert dict(master.attrib) == {}
    assert master.text == "A4*3"


def test_formula_group_error_is_a_value_error():
    assert issubclass(FormulaGroupError, XlsxeditError)
    assert issubclass(FormulaGroupError, ValueError)


RANGE_FORMULAS = {
    "array": '<f t="array" ref="B1:B2">A1:A2*2</f>',
    "dataTable": '<f t="dataTable" ref="B1:B2" dt2D="0" dtr="0" r1="A1"/>',
}

PARTIAL_EDITS = {
    "formula on the anchor": lambda ws: setattr(ws["B1"], "formula", "A1*3"),
    "formula on another cell": lambda ws: setattr(ws["B2"], "formula", "A2*3"),
    "formula removal": lambda ws: setattr(ws["B1"], "formula", None),
    "value": lambda ws: setattr(ws["B1"], "value", 0),
    "clear": lambda ws: ws["B1"].clear(),
    "clear_range": lambda ws: ws.clear_range("B1:B2"),
    "write_rows": lambda ws: ws.write_rows([[0]], at_cell="B1"),
    # Excel refuses these too: "You can't change part of an array." and "Can't change
    # part of a data table." (recorded with Excel 16 on 2026-11-04: neither a value
    # written into part of an array nor ClearContents there changes it)
    "value on another cell": lambda ws: setattr(ws["B2"], "value", 0),
    "None on another cell": lambda ws: setattr(ws["B2"], "value", None),
    "clear on another cell": lambda ws: ws["B2"].clear(),
    "clear_range of another cell": lambda ws: ws.clear_range("A2:C5"),
    "write_rows over another cell": lambda ws: ws.write_rows([[5, 6]], at_cell="A2"),
    "write_rows of None over another cell": lambda ws: ws.write_rows([[None]], at_cell="B2"),
    "write_dataframe over another cell": lambda ws: ws._workbook.write_dataframe(
        [(0,)], at_cell="B2", header=False, clear_range="C1:C2"
    ),
    "add_table header over another cell": lambda ws: ws.add_table(
        "B2:C3", ["x", "y"], write_header=True
    ),
}


@pytest.mark.parametrize("edit", PARTIAL_EDITS)
@pytest.mark.parametrize("kind", RANGE_FORMULAS)
def test_partial_edit_of_a_range_formula_is_refused(kind: str, edit: str):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = 1
    ws["A2"].value = 2
    _plant(ws["B1"], {}, f"{RANGE_FORMULAS[kind]}<v>2</v>")
    _plant(ws["B2"], {}, "<v>4</v>")
    before = read_pkg(wb)

    with pytest.raises(FormulaGroupError, match=f"{kind} formula .*B1:B2"):
        PARTIAL_EDITS[edit](ws)
    assert_preserved(before, read_pkg(wb))


def _array_over_b1_b2():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = 1
    ws["A2"].value = 2
    _plant(ws["B1"], {}, f"{RANGE_FORMULAS['array']}<v>2</v>")
    _plant(ws["B2"], {}, "<v>4</v>")
    return wb, ws


def test_a_value_a_merge_sends_to_part_of_an_array_is_refused():
    wb, ws = _array_over_b1_b2()
    ws.merge_cells("B2:C2")
    before = read_pkg(wb)
    with pytest.raises(FormulaGroupError, match="B2 lies in the array formula anchored at B1"):
        ws.write_rows([["x"]], at_cell="C2")
    assert_preserved(before, read_pkg(wb))


def test_an_insert_value_a_merge_sends_to_part_of_an_array_is_refused():
    wb, ws = _array_over_b1_b2()
    ws.merge_cells("B2:C3")  # its top-left cell lies in the array; the insert grows it
    before = read_pkg(wb)
    with pytest.raises(FormulaGroupError, match="B2 lies in the array formula anchored at B1"):
        ws.insert_rows([[None, "x"]], at_cell="B3")
    assert_preserved(before, read_pkg(wb))


def test_range_formulas_are_found_once_however_the_sheet_changes():
    """Finding them takes a pass over the sheet, so merges, clears, writes and inserts,
    which add none, keep what was found; an insert moves the formulas it finds."""
    wb, ws = _array_over_b1_b2()
    ws["C1"].value = 1
    found = ws._range_formulas
    assert [str(area) for _, _, area in ws._range_formula_areas()] == ["B1:B2"]
    ws.merge_cells("E1:F1")
    ws.clear_range("C1:C2")
    ws.write_rows([[1, 2]], at_cell="E5")
    ws.insert_rows([["x"]], at_row=1)
    assert ws._range_formulas is found
    assert [(anchor, str(area)) for _, anchor, area in ws._range_formula_areas()] == [("B2", "B2:B3")]
    with pytest.raises(FormulaGroupError):
        ws["B3"].value = 0


def test_writes_beside_an_array_and_into_a_single_cell_array_are_accepted():
    wb, ws = _array_over_b1_b2()
    _plant(ws["D1"], {}, '<f t="array" ref="D1">SUM(A1:A2)</f><v>3</v>')
    ws["C2"].value = 1
    ws.write_rows([[5]], at_cell="A2")
    ws.clear_range("C1:D2")
    assert (ws["A2"].value, ws["B2"].value, ws["C2"].value, ws["D1"].value) == (5, 4, None, None)


def test_part_of_an_array_follows_an_insert():
    wb, ws = _array_over_b1_b2()
    ws.insert_rows([["new"]], at_row=1)  # the array moves to B2:B3
    ws["B1"].value = "free now"
    with pytest.raises(FormulaGroupError, match="B3 lies in the array formula anchored at B2 over B2:B3"):
        ws["B3"].value = 0


def test_typed_replace_and_placeholders_skip_the_cells_of_an_array():
    from tests.conftest import ASSETS

    wb, ws = _array_over_b1_b2()
    ws["C5"].value = 4
    assert wb.replace(4, 9, value_type="number") == 1
    assert (ws["B2"].value, ws["C5"].value) == (4, 9)
    _plant(ws["B2"], {"t": "s"}, f"<v>{wb.shared_strings.add('{logo}')}</v>")
    assert wb.insert_image_at_placeholder("{logo}", ASSETS / "XlsxEdit-Demo.jpg") == 0
    assert ws["B2"].value == "{logo}"


def test_clear_range_refuses_an_anchor_before_clearing_any_cell():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = 1
    _plant(ws["B1"], {}, f"{RANGE_FORMULAS['array']}<v>2</v>")
    _plant(ws["B2"], {}, "<v>4</v>")
    before = read_pkg(wb)

    with pytest.raises(FormulaGroupError):
        ws.clear_range("A1:B2")
    assert_preserved(before, read_pkg(wb))


def test_formula_on_a_single_cell_array_keeps_it_an_array_formula():
    wb = Workbook.create()
    cell = wb["Sheet1"]["C1"]
    _plant(cell, {}, '<f t="array" ref="C1">SUM(A1:A2*B1:B2)</f><v>5</v>')

    cell.formula = "=SUM(A1:A3*B1:B3)"
    f = cell._element.find(_F)
    assert dict(f.attrib) == {"t": "array", "ref": "C1"}
    assert f.text == "SUM(A1:A3*B1:B3)"
    assert cell._element.find(_V) is None

    cell.value = 7
    assert cell.formula_type is None
    assert cell.value == 7


@pytest.mark.parametrize("expr", ["", "=", "  "])
def test_empty_formula_is_rejected(expr: str):
    cell = Workbook.create()["Sheet1"]["A1"]
    with pytest.raises(ValueError, match="empty"):
        cell.formula = expr
    assert not cell.has_formula


@pytest.mark.parametrize("address", ["A3", "C3", "C5"])
def test_formula_text_xml_cannot_hold_leaves_the_cell_unchanged(address: str):
    before = read_pkg(SIMPLE)
    wb = Workbook.open(SIMPLE)
    with pytest.raises(ValueError, match="XML compatible"):
        wb.worksheets[0][address].formula = "A5\x00"
    assert_preserved(before, read_pkg(wb))


def test_non_string_formula_is_rejected():
    with pytest.raises(TypeError, match="str or None"):
        Workbook.create()["Sheet1"]["A1"].formula = 5
