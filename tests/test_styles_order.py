"""Stylesheet and cell-format children are created in schema order."""

from __future__ import annotations

from pathlib import Path

import pytest
from lxml import etree

from xlsxedit import Workbook
from xlsxedit.opc.constants import SML_NS
from xlsxedit.styles import Styles
from tests.conftest import FIXTURES
from tests.schema_validate import assert_valid_package

FIXTURE_PATHS = sorted(FIXTURES.glob("*.xlsx"))


def _names(elm: etree._Element) -> list[str]:
    return [etree.QName(child).localname for child in elm]


def _stylesheet(wb: Workbook) -> etree._Element:
    return wb._workbook_part.styles_part.element


@pytest.mark.parametrize(
    "path", [None, *FIXTURE_PATHS], ids=["default", *(p.stem for p in FIXTURE_PATHS)]
)
def test_cf_rule_on_a_stylesheet_without_dxfs_stays_valid(path: Path | None):
    """Regression: a missing ``<dxfs>`` was appended after ``tableStyles``/``extLst``."""
    wb = Workbook.create() if path is None else Workbook.open(path)
    stylesheet = _stylesheet(wb)
    stylesheet.remove(stylesheet.find(f"{{{SML_NS}}}dxfs"))

    wb.worksheets[0].add_conditional_formatting("A1", operator="greaterThan", formula="0")

    names = _names(stylesheet)
    assert names.index("cellStyles") < names.index("dxfs") < names.index("tableStyles")
    assert_valid_package(wb)


def test_alignment_goes_before_existing_protection():
    """Regression: ``clone_xf`` appended ``<alignment>`` after ``<protection>``."""
    wb = Workbook.create()
    xf = _stylesheet(wb).find(f"{{{SML_NS}}}cellXfs")[0]
    etree.SubElement(xf, f"{{{SML_NS}}}protection").set("locked", "0")
    cell = wb["Sheet1"]["A1"]
    cell.value = 1
    cell.apply_style(horizontal_align="center")

    new_xf = _stylesheet(wb).find(f"{{{SML_NS}}}cellXfs")[-1]
    assert _names(new_xf) == ["alignment", "protection"]
    assert_valid_package(wb)


def test_missing_fills_and_cell_xfs_are_created_in_place():
    stylesheet = etree.fromstring(
        f'<styleSheet xmlns="{SML_NS}"><numFmts count="0"/><cellStyleXfs count="0"/>'
        '<cellStyles count="0"/><dxfs count="0"/><extLst/></styleSheet>'
    )
    styles = Styles(stylesheet)
    styles._fills_element()
    styles._cell_xfs_element()
    assert _names(stylesheet) == [
        "numFmts",
        "fills",
        "cellStyleXfs",
        "cellXfs",
        "cellStyles",
        "dxfs",
        "extLst",
    ]


def test_misplaced_num_fmts_moves_to_the_front():
    wb = Workbook.create()
    stylesheet = _stylesheet(wb)
    etree.SubElement(stylesheet, f"{{{SML_NS}}}numFmts").set("count", "0")
    cell = wb["Sheet1"]["A1"]
    cell.value = 1.5
    cell.apply_number_format("0.000")

    assert _names(stylesheet)[0] == "numFmts"
    assert_valid_package(wb)


def test_misplaced_num_fmts_moves_to_the_front_without_fonts():
    stylesheet = etree.fromstring(
        f'<styleSheet xmlns="{SML_NS}"><fills count="0"/><numFmts count="0"/></styleSheet>'
    )
    Styles(stylesheet)._ensure_num_fmts_block()
    assert _names(stylesheet) == ["numFmts", "fills"]
