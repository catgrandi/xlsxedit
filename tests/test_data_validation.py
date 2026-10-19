"""Data validations: reading base and x14 rules from ``ws.data_validations``.

There is no data-validation fixture yet, so the tests build their rules into
workbooks and read them back from the saved package.
"""

from __future__ import annotations

import io

import pytest
from lxml import etree

from xlsxedit import DataValidation, InvalidRangeError, Workbook
from xlsxedit.opc.constants import SML_NS
from xlsxedit.worksheet_order import insert_worksheet_child
from tests.preservation import check_consistency, read_pkg

X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
XM = "http://schemas.microsoft.com/office/excel/2006/main"
NS = {"m": SML_NS, "x14": X14, "xm": XM}
SHEET = "xl/worksheets/sheet1.xml"

BASE_RULES = """
<dataValidations count="3">
  <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1"
      errorTitle="Oops" error="Pick from the list" promptTitle="Status" prompt="Choose a status"
      sqref="B2:B10"><formula1>"Open,Closed"</formula1></dataValidation>
  <dataValidation type="decimal" errorStyle="warning" operator="notBetween" allowBlank="true"
      sqref="C2:C10 E2"><formula1>0.5</formula1><formula2>$H$1</formula2></dataValidation>
  <dataValidation sqref="D2"/>
</dataValidations>
"""

X14_RULES = """
<extLst>
  <ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}">
    <x14:dataValidations count="1">
      <x14:dataValidation type="list" allowBlank="1" showErrorMessage="1" prompt="From Lists">
        <x14:formula1><xm:f>Lists!$A$1:$A$5</xm:f></x14:formula1>
        <xm:sqref>F2:F10</xm:sqref>
      </x14:dataValidation>
    </x14:dataValidations>
  </ext>
  <ext uri="{00000000-0000-0000-0000-000000000000}">
    <x14:dataValidations count="1">
      <x14:dataValidation type="whole"><xm:sqref>G2</xm:sqref></x14:dataValidation>
    </x14:dataValidations>
  </ext>
</extLst>
"""


def _insert_xml(ws, fragment: str) -> etree._Element:
    namespaces = f'xmlns="{SML_NS}" xmlns:x14="{X14}" xmlns:xm="{XM}"'
    elm = etree.fromstring(f"<w {namespaces}>{fragment}</w>")[0]
    insert_worksheet_child(ws._part.element, elm)
    return elm


def _book(*fragments: str):
    wb = Workbook.create()
    wb.add_worksheet("Lists")
    ws = wb["Sheet1"]
    for fragment in fragments:
        _insert_xml(ws, fragment)
    return wb, ws


def _reopened(wb: Workbook) -> Workbook:
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return Workbook.open(buf)


# -- reading ------------------------------------------------------------------


def test_reads_base_rules_then_x14_rules():
    wb, ws = _book(BASE_RULES, X14_RULES)
    rules = ws.data_validations
    assert [str(dv.sqref) for dv in rules] == ["B2:B10", "C2:C10 E2", "D2", "F2:F10"]
    assert [dv.is_x14 for dv in rules] == [False, False, False, True]
    assert all(isinstance(dv, DataValidation) for dv in rules)
    assert repr(rules[0]) == "<DataValidation list 'B2:B10'>"


def test_reads_every_attribute_of_a_base_rule():
    wb, ws = _book(BASE_RULES)
    listed, ranged, _ = _reopened(wb)["Sheet1"].data_validations
    assert (listed.type, listed.operator, listed.error_style) == ("list", "between", "stop")
    assert (listed.allow_blank, listed.show_input_message, listed.show_error_message) == (
        True,
        True,
        True,
    )
    assert (listed.error_title, listed.error) == ("Oops", "Pick from the list")
    assert (listed.prompt_title, listed.prompt) == ("Status", "Choose a status")
    assert (listed.formula1, listed.formula2) == ('"Open,Closed"', None)
    assert (ranged.type, ranged.operator) == ("decimal", "notBetween")
    assert ranged.error_style == "warning"
    assert (ranged.allow_blank, ranged.show_input_message) == (True, False)
    assert (ranged.formula1, ranged.formula2) == ("0.5", "$H$1")
    assert [(r.min_col, r.min_row, r.max_col, r.max_row) for r in ranged.sqref] == [
        (3, 2, 3, 10),
        (5, 2, 5, 2),
    ]


def test_absent_attributes_read_as_schema_defaults():
    wb, ws = _book(BASE_RULES)
    bare = ws.data_validations[2]
    assert (bare.type, bare.operator, bare.error_style) == ("none", "between", "stop")
    assert (bare.allow_blank, bare.show_input_message, bare.show_error_message) == (
        False,
        False,
        False,
    )
    assert (bare.error_title, bare.error, bare.prompt_title, bare.prompt) == (None,) * 4
    assert (bare.formula1, bare.formula2) == (None, None)


def test_reads_x14_rules_of_the_validation_extension_only():
    wb, ws = _book(X14_RULES)
    (rule,) = _reopened(wb)["Sheet1"].data_validations
    assert rule.is_x14
    assert str(rule.sqref) == "F2:F10"
    assert (rule.type, rule.formula1, rule.formula2) == ("list", "Lists!$A$1:$A$5", None)
    assert (rule.allow_blank, rule.show_input_message, rule.show_error_message) == (
        True,
        False,
        True,
    )
    assert rule.prompt == "From Lists"


def test_rules_compare_by_element():
    wb, ws = _book(BASE_RULES)
    first = ws.data_validations[0]
    assert first == ws.data_validations[0]
    assert first != ws.data_validations[1]
    assert first in ws.data_validations
    assert len({*ws.data_validations, *ws.data_validations}) == 3


def test_a_sheet_without_validations_has_none():
    assert Workbook.create()["Sheet1"].data_validations == []


def test_a_malformed_stored_sqref_raises_on_read(unchecked_workbooks):
    wb, ws = _book('<dataValidations count="1"><dataValidation sqref="B2:"/></dataValidations>')
    with pytest.raises(InvalidRangeError):
        ws.data_validations[0].sqref
    check_consistency(read_pkg(wb), waive={"grid-bounds"})
