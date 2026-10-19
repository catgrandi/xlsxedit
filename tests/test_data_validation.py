"""Data validations: reading base and x14 rules, ``add_data_validation``,
``remove_data_validation`` and the settable ``sqref``.

There is no data-validation fixture yet, so the tests build their rules into
workbooks and assert on the saved XML. Every refused edit is checked to leave
the saved package unchanged.
"""

from __future__ import annotations

import enum
import io
import re

import pytest
from lxml import etree

from xlsxedit import (
    CellRange,
    DataValidation,
    DataValidationError,
    InvalidRangeError,
    SheetRangeSet,
    Workbook,
)
from xlsxedit.opc.constants import SML_NS
from xlsxedit.worksheet_order import insert_worksheet_child
from tests.conftest import INSPECT_FIXTURES
from tests.preservation import assert_preserved, check_consistency, child_order, read_pkg
from tests.schema_validate import assert_valid_package

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


def _saved_sheet(wb: Workbook) -> etree._Element:
    return etree.fromstring(read_pkg(wb)[SHEET])


def _bare(elm: etree._Element) -> str:
    """``elm`` serialized without namespace declarations."""
    return re.sub(r' xmlns(:\w+)?="[^"]*"', "", etree.tostring(elm, encoding=str, with_tail=False))


def _saved_rules(wb: Workbook) -> list[str]:
    return [_bare(e) for e in _saved_sheet(wb).iterfind("m:dataValidations/m:dataValidation", NS)]


def _reopened(wb: Workbook) -> Workbook:
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return Workbook.open(buf)


def _refused(wb: Workbook, error: type[Exception], match: str, call) -> None:
    before = read_pkg(wb)
    with pytest.raises(error, match=match):
        call()
    assert_preserved(before, read_pkg(wb))


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
    assert (listed.type, listed.operator, listed.error_style) == ("list", None, "stop")
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
    assert (bare.type, bare.operator, bare.error_style) == ("none", None, "stop")
    assert (bare.allow_blank, bare.show_input_message, bare.show_error_message) == (
        False,
        False,
        False,
    )
    assert (bare.error_title, bare.error, bare.prompt_title, bare.prompt) == (None,) * 4
    assert (bare.formula1, bare.formula2) == (None, None)


def test_operator_defaults_to_between_for_the_types_that_compare():
    wb, ws = _book(
        '<dataValidations count="2">'
        '<dataValidation type="whole" sqref="A1"><formula1>1</formula1><formula2>2</formula2>'
        "</dataValidation>"
        '<dataValidation type="list" operator="equal" sqref="A2"><formula1>"a"</formula1>'
        "</dataValidation></dataValidations>"
    )
    assert [dv.operator for dv in ws.data_validations] == ["between", None]


def test_reads_a_formula_excel_wraps_in_alternate_content():
    mc = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    x12ac = "http://schemas.microsoft.com/office/spreadsheetml/2011/1/ac"
    wb, ws = _book(
        '<dataValidations count="1"><dataValidation type="list" allowBlank="1" sqref="B2">'
        f'<mc:AlternateContent xmlns:mc="{mc}" xmlns:x12ac="{x12ac}">'
        '<mc:Choice Requires="x12ac"><x12ac:list>"a,b",c</x12ac:list></mc:Choice>'
        '<mc:Fallback><formula1>"a,b,c"</formula1></mc:Fallback>'
        "</mc:AlternateContent></dataValidation></dataValidations>"
    )
    (rule,) = _reopened(wb)["Sheet1"].data_validations
    assert (rule.type, rule.formula1, rule.formula2) == ("list", '"a,b,c"', None)


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


def test_a_malformed_stored_sqref_raises_on_read_and_refuses_an_add(unchecked_workbooks):
    wb, ws = _book('<dataValidations count="1"><dataValidation sqref="B2:"/></dataValidations>')
    with pytest.raises(InvalidRangeError):
        ws.data_validations[0].sqref
    _refused(
        wb,
        InvalidRangeError,
        "invalid range",
        lambda: ws.add_data_validation("Z9", type="list", formula1='"a"'),
    )
    check_consistency(read_pkg(wb), waive={"grid-bounds"})


# -- adding -------------------------------------------------------------------


def test_add_writes_a_list_rule_as_excel_does():
    wb, ws = _book()
    before = read_pkg(wb)
    dv = ws.add_data_validation("B2:B10", type="list", formula1='"Yes,No"', prompt="Pick one")
    after = read_pkg(wb)
    assert_preserved(before, after, expected_changed={SHEET})
    assert _saved_rules(wb) == [
        '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" '
        'prompt="Pick one" sqref="B2:B10"><formula1>"Yes,No"</formula1></dataValidation>'
    ]
    assert _saved_sheet(wb).find("m:dataValidations", NS).get("count") == "1"
    assert dv in ws.data_validations and not dv.is_x14
    assert_valid_package(wb)


def test_add_places_the_container_in_schema_order_and_changes_only_that_sheet():
    path = INSPECT_FIXTURES["ConditionalFormatting"]
    wb = Workbook.open(path)
    wb["Sheet1"].add_data_validation("H2:H7", type="whole", operator="greaterThan", formula1=0)
    after = read_pkg(wb)
    assert_preserved(read_pkg(path), after, expected_changed={SHEET})
    order = child_order(after[SHEET])
    assert order[order.index("conditionalFormatting") :] == (
        ("conditionalFormatting",) * 5 + ("dataValidations", "pageMargins", "extLst")
    )
    assert_valid_package(wb)


def test_each_add_appends_to_the_one_container_and_counts():
    wb, ws = _book(BASE_RULES, X14_RULES)
    ws.add_data_validation("J2", type="custom", formula1="LEN(J2)<5")
    ws.add_data_validation("K2", type="none", prompt="Free text")
    root = _saved_sheet(wb)
    (block,) = root.findall("m:dataValidations", NS)
    assert block.get("count") == "5"
    assert [dv.get("sqref") for dv in block] == ["B2:B10", "C2:C10 E2", "D2", "J2", "K2"]
    assert [str(dv.sqref) for dv in ws.data_validations][-3:] == ["J2", "K2", "F2:F10"]
    assert_valid_package(wb)


@pytest.mark.parametrize(
    ("given", "written"),
    [
        ("b2:b10", "B2:B10"),
        ("$B$2:$B$10", "B2:B10"),
        ("B10:B2", "B2:B10"),
        ("B:B", "B1:B1048576"),
        ("3:4", "A3:XFD4"),
        ("C5:C5", "C5"),
        (" B2  D4:E5\tB2 ", "B2 D4:E5"),
        (CellRange.parse("D4:E5"), "D4:E5"),
        (SheetRangeSet.parse("B2 $D$4"), "B2 D4"),
    ],
)
def test_add_writes_sqref_as_excel_does(given, written):
    wb, ws = _book()
    dv = ws.add_data_validation(given, type="list", formula1='"a,b"')
    assert str(dv.sqref) == written
    assert _saved_sheet(wb).find("m:dataValidations/m:dataValidation", NS).get("sqref") == written


@pytest.mark.parametrize(
    ("options", "attributes", "formulas"),
    [
        (
            {"type": "whole", "formula1": 1, "formula2": "=10"},
            'type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1"',
            ("1", "10"),
        ),
        (
            {
                "type": "decimal",
                "operator": "lessThan",
                "formula1": 0.25,
                "error_style": "information",
            },
            'type="decimal" errorStyle="information" operator="lessThan" allowBlank="1" '
            'showInputMessage="1" showErrorMessage="1"',
            ("0.25", None),
        ),
        (
            {"type": "time", "operator": "between", "formula1": 0.25, "formula2": 0.75},
            'type="time" allowBlank="1" showInputMessage="1" showErrorMessage="1"',
            ("0.25", "0.75"),
        ),
        (
            {"type": "date", "operator": "greaterThan", "formula1": "DATE(2026,1,1)"},
            'type="date" operator="greaterThan" allowBlank="1" showInputMessage="1" '
            'showErrorMessage="1"',
            ("DATE(2026,1,1)", None),
        ),
        (
            {"type": "textLength", "operator": "lessThanOrEqual", "formula1": 10**6},
            'type="textLength" operator="lessThanOrEqual" allowBlank="1" showInputMessage="1" '
            'showErrorMessage="1"',
            ("1000000", None),
        ),
        (
            {"type": "none", "prompt": "Note", "allow_blank": False},
            'showInputMessage="1" showErrorMessage="1" prompt="Note"',
            (None, None),
        ),
        (
            {
                "type": "list",
                "formula1": "=$H$1:$H$3",
                "allow_blank": False,
                "show_input_message": False,
                "show_error_message": False,
                "error_style": "stop",
                "error_title": "Error",
                "error": "Bad value",
                "prompt_title": "Pick",
            },
            'type="list" errorTitle="Error" error="Bad value" promptTitle="Pick"',
            ("$H$1:$H$3", None),
        ),
    ],
)
def test_add_writes_only_attributes_that_differ_from_the_schema(options, attributes, formulas):
    wb, ws = _book()
    ws.add_data_validation("B2", **options)
    expected = "".join(
        f"<{name}>{text}</{name}>"
        for name, text in zip(("formula1", "formula2"), formulas)
        if text is not None
    )
    tail = f">{expected}</dataValidation>" if expected else "/>"
    assert _saved_rules(wb) == [f'<dataValidation {attributes} sqref="B2"{tail}']
    rule = _reopened(wb)["Sheet1"].data_validations[0]
    assert (rule.formula1, rule.formula2) == formulas
    assert rule.type == options["type"]
    assert_valid_package(wb)


@pytest.mark.parametrize(
    ("formula1", "written"),
    [
        ('"' + "x" * 255 + '"', None),
        ('"Hi!,Bye"', None),
        ('"say ""hi!"""', None),
        ("MyList", None),
        ("$H$1:$H$3", None),
        ("Sheet1!$H$1:$H$3", "$H$1:$H$3"),
        ("'Sheet1'!$H$1:$H$3", "$H$1:$H$3"),
        ("sheet1!H1:H3", "H1:H3"),
        ("Sheet1!H1:Sheet1!H3", "H1:H3"),
        ("$H$1:Sheet1!$H$3", "$H$1:$H$3"),
        ("'Sheet1'!H1:'Sheet1'!H3", "H1:H3"),
        ("OFFSET(Sheet1!$H$1,0,0,COUNTA('Sheet1'!$H:$H),1)", "OFFSET($H$1,0,0,COUNTA($H:$H),1)"),
        ('INDIRECT("Lists!A1:A3")', None),
    ],
)
def test_add_accepts_list_sources_on_the_same_sheet(formula1, written):
    wb, ws = _book()
    dv = ws.add_data_validation("B2", type="list", formula1=formula1)
    assert dv.formula1 == (written or formula1)


@pytest.mark.parametrize(
    ("formula1", "written"),
    [
        ("A1<>#REF!", None),
        ("ISERROR(A1)=#div/0!", None),
        ("AND(A1<>#NULL!,#N/A=A1)", None),
        ("Sheet1!#REF!=A1", "#REF!=A1"),
        ('A1&"Sheet1!x"&Sheet1!B1="a"', 'A1&"Sheet1!x"&B1="a"'),
    ],
)
def test_error_values_and_strings_are_not_sheet_references(formula1, written):
    wb, ws = _book()
    dv = ws.add_data_validation("B2", type="custom", formula1=formula1)
    assert dv.formula1 == (written or formula1)


def test_sheet_names_match_without_case_but_not_folded():
    wb = Workbook.create()
    wb.rename_worksheet("Sheet1", "Straße")
    wb.add_worksheet("STRASSE")
    ws = wb["Straße"]
    _refused(
        wb,
        DataValidationError,
        "another sheet",
        lambda: ws.add_data_validation("B2", type="list", formula1="STRASSE!$A$1:$A$3"),
    )
    dv = ws.add_data_validation("B2", type="list", formula1="'STRAßE'!$A$1:$A$3")
    assert dv.formula1 == "$A$1:$A$3"


def test_formulas_name_the_sheet_as_it_is_called_now():
    wb, ws = _book()
    wb.rename_worksheet("Sheet1", "Data")
    stale = ws
    _refused(
        wb,
        DataValidationError,
        "another sheet",
        lambda: stale.add_data_validation("B2", type="list", formula1="Sheet1!$H$1:$H$3"),
    )
    dv = stale.add_data_validation("B2", type="list", formula1="Data!$H$1:$H$3")
    assert dv.formula1 == "$H$1:$H$3"
    assert [str(v.sqref) for v in wb["Data"].data_validations] == ["B2"]


def test_a_removed_sheet_refuses_every_edit():
    wb, ws = _book()
    lists = wb["Lists"]
    dv = lists.add_data_validation("A1", type="list", formula1='"a"')
    wb.remove_worksheet("Lists")
    for call in (
        lambda: lists.add_data_validation("B1", type="list", formula1='"a"'),
        lambda: setattr(dv, "sqref", "B1"),
        lambda: lists.remove_data_validation(dv),
    ):
        _refused(wb, DataValidationError, "no longer", call)


class _Level(enum.IntEnum):
    LOW = 3


class _Float(float):
    def __repr__(self) -> str:
        return f"_Float({float(self)!r})"


@pytest.mark.parametrize(
    ("formula1", "written"),
    [
        (" =A1+1 ", "A1+1"),
        ("=\tA1+1", "A1+1"),
        (_Level.LOW, "3"),
        (_Float(0.25), "0.25"),
        (-0.5, "-0.5"),
        (1e20, "1e+20"),
        (10**20, "100000000000000000000"),
    ],
)
def test_add_writes_formulas_as_excel_reads_them(formula1, written):
    wb, ws = _book()
    ws.add_data_validation("B2", type="decimal", operator="greaterThan", formula1=formula1)
    assert _saved_rules(wb)[0].endswith(f"<formula1>{written}</formula1></dataValidation>")


def test_add_accepts_messages_at_excels_limits():
    wb, ws = _book()
    dv = ws.add_data_validation(
        "B2",
        type="list",
        formula1='"a,b"',
        prompt_title="t" * 32,
        prompt="p" * 255,
        error_title="\U0001f600" * 16,
        error="Line one\nline two & <three>",
    )
    rule = _reopened(wb)["Sheet1"].data_validations[0]
    assert (rule.prompt_title, rule.prompt) == ("t" * 32, "p" * 255)
    assert (rule.error_title, rule.error) == ("\U0001f600" * 16, "Line one\nline two & <three>")
    assert dv.error == rule.error
    assert_valid_package(wb)


_LIST = {"sqref": "B2", "type": "list", "formula1": '"a,b"'}


@pytest.mark.parametrize(
    ("options", "error", "match"),
    [
        ({"sqref": "B2:"}, InvalidRangeError, "invalid range"),
        ({"sqref": " "}, InvalidRangeError, "names no cells"),
        ({"sqref": "A1048577"}, InvalidRangeError, "outside"),
        ({"sqref": "XFE1"}, InvalidRangeError, "past XFD"),
        ({"sqref": "B2,B3"}, InvalidRangeError, "invalid range"),
        ({"sqref": ["B2"]}, TypeError, "sqref must be"),
        ({"sqref": SheetRangeSet(["B2"])}, TypeError, "CellRange areas"),
        ({"type": "lists"}, DataValidationError, "type must be one of"),
        ({"type": None}, TypeError, "type must be a str"),
        (
            {"type": "whole", "operator": "bigger", "formula1": 1},
            DataValidationError,
            "operator must be one of",
        ),
        ({"operator": "equal"}, DataValidationError, "operator applies to"),
        ({"type": "custom", "operator": "between"}, DataValidationError, "operator applies to"),
        ({"error_style": "Stop"}, DataValidationError, "error_style must be one of"),
        ({"formula1": None}, DataValidationError, "needs formula1"),
        ({"type": "whole", "formula1": 1}, DataValidationError, "'between' needs formula2"),
        (
            {"type": "date", "operator": "notBetween", "formula1": 1},
            DataValidationError,
            "'notBetween' needs formula2",
        ),
        (
            {"type": "whole", "operator": "equal", "formula1": 1, "formula2": 2},
            DataValidationError,
            "formula2 applies only",
        ),
        ({"formula2": "1"}, DataValidationError, "formula2 applies only"),
        ({"type": "none"}, DataValidationError, "takes no formula"),
        ({"type": "custom", "formula1": "="}, DataValidationError, "formula1 is empty"),
        ({"type": "custom", "formula1": "  "}, DataValidationError, "formula1 is empty"),
        ({"type": "custom", "formula1": " = "}, DataValidationError, "formula1 is empty"),
        ({"type": "custom", "formula1": "==A1"}, DataValidationError, "more than one '='"),
        ({"type": "custom", "formula1": " = =A1"}, DataValidationError, "more than one '='"),
        ({"formula1": True}, TypeError, "formula1 must be"),
        ({"formula1": ["a", "b"]}, TypeError, "formula1 must be"),
        (
            {"type": "decimal", "operator": "lessThan", "formula1": float("nan")},
            DataValidationError,
            "not a number",
        ),
        (
            {"type": "decimal", "operator": "lessThan", "formula1": float("-inf")},
            DataValidationError,
            "not a number",
        ),
        (
            {"type": "whole", "operator": "lessThan", "formula1": 10**400},
            DataValidationError,
            "not a number",
        ),
        ({"formula1": '"a\x01b"'}, DataValidationError, "XML cannot store"),
        ({"type": "custom", "formula1": "A1+" * 3000 + "1"}, DataValidationError, "8192"),
        ({"formula1": '"' + "x" * 256 + '"'}, DataValidationError, "255"),
        ({"formula1": '"' + '""' * 256 + '"'}, DataValidationError, "255"),
        ({"formula1": ' "' + "x" * 256 + '"\n'}, DataValidationError, "255"),
        ({"formula1": "Lists!$A$1:$A$5"}, DataValidationError, "another sheet"),
        ({"formula1": "='Lists'!$A$1:$A$5"}, DataValidationError, "another sheet"),
        ({"formula1": "'Other ''quoted'' sheet'!A1"}, DataValidationError, "another sheet"),
        ({"formula1": "[1]Sheet1!$A$1:$A$5"}, DataValidationError, "another sheet"),
        ({"formula1": "'[Book.xlsx]Sheet1'!$A$1"}, DataValidationError, "another sheet"),
        ({"formula1": "Sheet1:Lists!$A$1"}, DataValidationError, "another sheet"),
        ({"formula1": "Lists:Sheet1!$A$1"}, DataValidationError, "another sheet"),
        ({"formula1": "Sheet1!H:Sheet1!H"}, DataValidationError, "another sheet"),
        ({"type": "custom", "formula1": "SUM(Jan:Sheet1!B2)<9"}, DataValidationError, "sheet"),
        ({"type": "custom", "formula1": "SUM(A:Sheet1!B2)>0"}, DataValidationError, "sheet"),
        ({"formula1": "x'Sheet1'!A1"}, DataValidationError, "another sheet"),
        ({"formula1": "$Sheet1!A1"}, DataValidationError, "another sheet"),
        ({"formula1": "Sheet1!Sheet1!A1"}, DataValidationError, "another sheet"),
        ({"type": "custom", "formula1": "Other#Sheet1!A1"}, DataValidationError, "another sheet"),
        ({"formula1": "Sheet1!A1:Lists!A2"}, DataValidationError, "another sheet"),
        ({"type": "custom", "formula1": "A1=#REF!+Lists!A1"}, DataValidationError, "another sheet"),
        (
            {"type": "custom", "formula1": 'COUNTIF(Lists!A:A,"x!")>0'},
            DataValidationError,
            "another sheet",
        ),
        (
            {"type": "whole", "operator": "between", "formula1": 1, "formula2": "Lists!$A$1"},
            DataValidationError,
            "formula2 .* another sheet",
        ),
        ({"type": "custom", "formula1": "LEN(A1)!"}, DataValidationError, "another sheet"),
        ({"allow_blank": 1}, TypeError, "allow_blank must be a bool"),
        ({"show_input_message": None}, TypeError, "show_input_message must be a bool"),
        ({"show_error_message": "yes"}, TypeError, "show_error_message must be a bool"),
        ({"prompt": 5}, TypeError, "prompt must be a str"),
        ({"error_style": None}, TypeError, "error_style must be a str"),
        ({"prompt_title": "x" * 33}, DataValidationError, "32"),
        ({"error_title": "\U0001f600" * 17}, DataValidationError, "32"),
        ({"prompt": "x" * 256}, DataValidationError, "255"),
        ({"error": "x" * 256}, DataValidationError, "255"),
        ({"error": "bad\x00"}, DataValidationError, "XML cannot store"),
        ({"prompt": "lone \ud800"}, DataValidationError, "XML cannot store"),
    ],
)
def test_add_refuses_without_changing_anything(options, error, match):
    wb, ws = _book(BASE_RULES)
    kwargs = {**_LIST, **options}
    _refused(wb, error, match, lambda: ws.add_data_validation(**kwargs))
    assert len(ws.data_validations) == 3


@pytest.mark.parametrize("sqref", ["B5", "A1:C2", "E1:E5", "F:F", "Z1 F10"])
def test_add_refuses_a_range_that_overlaps_another_validation(sqref):
    wb, ws = _book(BASE_RULES, X14_RULES)
    _refused(
        wb,
        DataValidationError,
        "overlaps",
        lambda: ws.add_data_validation(sqref, type="list", formula1='"a"'),
    )


def test_rules_of_other_extensions_do_not_block_an_add():
    wb, ws = _book(BASE_RULES, X14_RULES)
    ws.add_data_validation("G:G", type="list", formula1='"a"')
    assert str(ws.data_validations[3].sqref) == "G1:G1048576"


def test_add_needs_a_type():
    wb, ws = _book()
    with pytest.raises(TypeError):
        ws.add_data_validation("B2", formula1='"a"')


def test_a_rule_read_back_adds_an_equal_rule():
    wb, ws = _book(BASE_RULES)
    copy = wb["Lists"]
    names = (
        "type",
        "formula1",
        "formula2",
        "operator",
        "allow_blank",
        "show_input_message",
        "show_error_message",
        "error_style",
        "error_title",
        "error",
        "prompt_title",
        "prompt",
    )

    def fields(dv: DataValidation) -> list:
        return [str(dv.sqref), *(getattr(dv, n) for n in names)]

    for rule in ws.data_validations:
        copy.add_data_validation(str(rule.sqref), **{n: getattr(rule, n) for n in names})
    copied = _reopened(wb)["Lists"].data_validations
    assert [fields(dv) for dv in copied] == [fields(dv) for dv in ws.data_validations]
    assert_valid_package(wb)


# -- sqref --------------------------------------------------------------------


def test_setting_sqref_rewrites_only_that_attribute():
    wb, ws = _book(BASE_RULES)
    before = read_pkg(wb)
    rule = ws.data_validations[0]
    rule.sqref = "b2:b20 $H$5 B2:B3"
    assert_preserved(before, read_pkg(wb), expected_changed={SHEET})
    assert str(rule.sqref) == "B2:B20 H5 B2:B3"
    assert _saved_rules(wb)[0].endswith(
        'sqref="B2:B20 H5 B2:B3"><formula1>"Open,Closed"</formula1></dataValidation>'
    )
    rule.sqref = SheetRangeSet.parse("A1")
    assert str(ws.data_validations[0].sqref) == "A1"


def test_setting_sqref_may_overlap_its_own_cells_only():
    wb, ws = _book(BASE_RULES)
    rule = ws.data_validations[0]
    rule.sqref = "B1:B11"
    _refused(wb, DataValidationError, "overlaps", lambda: setattr(rule, "sqref", "B2:D2"))
    assert str(rule.sqref) == "B1:B11"


@pytest.mark.parametrize(
    ("value", "error", "match"),
    [
        ("B2:", InvalidRangeError, "invalid range"),
        ("", InvalidRangeError, "names no cells"),
        (5, TypeError, "sqref must be"),
        (None, TypeError, "sqref must be"),
    ],
)
def test_setting_a_bad_sqref_changes_nothing(value, error, match):
    wb, ws = _book(BASE_RULES)
    rule = ws.data_validations[0]
    _refused(wb, error, match, lambda: setattr(rule, "sqref", value))


def test_x14_rules_are_read_only():
    wb, ws = _book(BASE_RULES, X14_RULES)
    x14 = ws.data_validations[-1]
    _refused(wb, DataValidationError, "read-only", lambda: setattr(x14, "sqref", "Z1"))
    _refused(wb, DataValidationError, "read-only", lambda: ws.remove_data_validation(x14))
    assert str(x14.sqref) == "F2:F10"


# -- removing -----------------------------------------------------------------


def test_remove_drops_the_rule_and_recounts():
    wb, ws = _book(BASE_RULES, X14_RULES)
    before = read_pkg(wb)
    middle = ws.data_validations[1]
    ws.remove_data_validation(middle)
    assert_preserved(before, read_pkg(wb), expected_changed={SHEET})
    block = _saved_sheet(wb).find("m:dataValidations", NS)
    assert block.get("count") == "2"
    assert [dv.get("sqref") for dv in block] == ["B2:B10", "D2"]
    assert middle not in ws.data_validations
    assert len(ws.data_validations) == 3


def test_removing_the_last_rule_removes_the_container():
    wb, ws = _book()
    before = read_pkg(wb)
    dv = ws.add_data_validation("B2", type="list", formula1='"a"')
    ws.remove_data_validation(dv)
    report = assert_preserved(before, read_pkg(wb))
    assert SHEET in report.reserialised  # an edited part is rewritten even when restored


def test_remove_refuses_a_rule_that_is_not_on_the_sheet():
    wb, ws = _book(BASE_RULES)
    other = wb["Lists"]
    elsewhere = other.add_data_validation("A1", type="list", formula1='"a"')
    rule = ws.data_validations[0]
    _refused(
        wb,
        DataValidationError,
        "not on worksheet 'Sheet1'",
        lambda: ws.remove_data_validation(elsewhere),
    )
    ws.remove_data_validation(rule)
    _refused(wb, DataValidationError, "not on worksheet", lambda: ws.remove_data_validation(rule))
    _refused(
        wb, DataValidationError, "no longer on its worksheet", lambda: setattr(rule, "sqref", "Z1")
    )
    _refused(wb, TypeError, "expected a DataValidation", lambda: ws.remove_data_validation("B2"))


# -- other edits --------------------------------------------------------------


def test_an_added_rule_moves_with_an_insert():
    wb, ws = _book()
    dv = ws.add_data_validation("B5:B6", type="list", formula1='"a"')
    ws.insert_rows([[1]], at_row=2)
    assert str(dv.sqref) == "B6:B7"
    assert_valid_package(wb)


def test_a_copied_sheet_keeps_its_rules():
    wb, ws = _book(BASE_RULES)
    copy = wb.copy_worksheet("Sheet1", "Copy")
    copy.remove_data_validation(copy.data_validations[0])
    assert [str(dv.sqref) for dv in ws.data_validations] == ["B2:B10", "C2:C10 E2", "D2"]
    assert [str(dv.sqref) for dv in copy.data_validations] == ["C2:C10 E2", "D2"]
