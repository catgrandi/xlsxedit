"""``tests/excel_facts.py`` tells apart what Excel shows differently."""

from __future__ import annotations

from collections.abc import Callable

from lxml import etree

from tests import excel_facts as facts
from tests.conftest import FIXTURES, INSPECT_FIXTURES
from tests.preservation import read_pkg

DELETED = FIXTURES / "excel-reference" / "delete-rows-under-cf" / "after.xlsx"
SHEET1 = "xl/worksheets/sheet1.xml"
X14_CF = f"{{{facts.X14}}}conditionalFormatting"


def _changed(path, member: str, edit: Callable[[etree._Element], None]) -> facts.Book:
    pkg = dict(read_pkg(path))
    root = etree.fromstring(pkg[member])
    edit(root)
    pkg[member] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return facts.Book(pkg)


def _x14_block(root: etree._Element, area: str) -> etree._Element:
    (block,) = [b for b in root.iter(X14_CF) if b.findtext(f"{{{facts.XM}}}sqref") == area]
    return block


def test_an_x14_twin_keeps_its_own_range():
    def widen(root):
        _x14_block(root, "G2:G7").find(f"{{{facts.XM}}}sqref").text = "G2:G10"

    excel = facts.conditional_formats(facts.Book(DELETED))
    assert facts.conditional_formats(_changed(DELETED, SHEET1, widen)) != excel


def test_an_x14_twin_without_its_base_rule_is_listed():
    def drop_base(root):
        (base,) = root.iterfind(f"{facts.sml('conditionalFormatting')}[@sqref='G2:G7']")
        root.remove(base)

    rules = facts.conditional_formats(_changed(DELETED, SHEET1, drop_base))["Sheet1"]
    assert [r.sqref for r in rules if r.source == "x14 twin without base rule"] == ["G2:G7"]


def test_a_boolean_is_not_a_number():
    def to_boolean(root):
        cell = root.find(f".//{facts.sml('c')}[@r='A2']")
        cell.set("t", "b")

    path = INSPECT_FIXTURES["Formulas"]
    before = facts.cells(facts.Book(path), "Calc!A2")
    after = facts.cells(_changed(path, SHEET1, to_boolean), "Calc!A2")
    assert before == {"A2": facts.CellFact("n", 1.0, None)}
    assert after == {"A2": facts.CellFact("b", True, None)}


def test_a_data_table_input_cell_counts():
    def move_input(root):
        root.find(f".//{facts.sml('f')}[@t='dataTable']").set("r1", "I3")

    path = INSPECT_FIXTURES["Formulas"]
    excel = facts.cells(facts.Book(path), "Calc!L3")
    assert excel["L3"].formula.options == "dt2D=0,dtr=0,r1=I2"
    assert facts.cells(_changed(path, SHEET1, move_input), "Calc!L3") != excel
