"""Conditional-format, validation and x14 formula text on insert and delete (issue #12).

Excel evaluates these formulas relative to the top-left cell of their areas.
When rows move, Excel adjusts the formula of every cell and splits the format
where cells come to need different formulas; xlsxedit keeps one formula, the
one Excel writes for the part that holds the new top-left cell. The table
below was recorded from Excel for Microsoft 365 on Windows (2026-10-19).
"""

from __future__ import annotations

import io

import pytest
from lxml import etree

from xlsxedit import Workbook
from xlsxedit.exceptions import GridOverflowError
from xlsxedit.opc.constants import SML_NS
from xlsxedit.row_shift import _Planner, _Shift
from xlsxedit.worksheet_order import insert_worksheet_child
from tests.preservation import assert_preserved
from tests.schema_validate import assert_valid_package

X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
XM = "http://schemas.microsoft.com/office/excel/2006/main"
NS = {"m": SML_NS, "x14": X14, "xm": XM}

_FORMATS = [  # (sqref, formula of an expression rule)
    ("B2:B20", "B12>0"),
    ("C10:C20", "$C10>10"),
    ("D3:D20", "D1>0"),
    ("E5:E10", "E5>0"),
    ("F2:F20", "$G$12>0"),
    ("J5:J8 H10:H12", "J5>0"),  # relative to H5, the corner of the box around both areas
    ("L10:L12 K5:K8", "L10>0"),
    ("M2:M20", "SUM(N4:N8)>0"),
    ("P5:P20", "P5>0"),
]
_VALIDATIONS = [("Q2:Q20", "list", "$R$10:$R$15"), ("S2:S20", "custom", "S12>0"), ("T5:T10", "custom", "T5>0")]

_EXCEL = {
    "insert row 5": (
        _Shift(True, 5, 1),
        "B13>0 $C11>10 D1>0 E6>0 $G$13>0 J6>0 L11>0 SUM(N4:N9)>0 P6>0 $R$11:$R$16 S13>0 T6>0",
    ),
    "insert row 2": (
        _Shift(True, 2, 1),
        "B13>0 $C11>10 D1>0 E6>0 $G$13>0 J6>0 L11>0 SUM(N5:N9)>0 P6>0 $R$11:$R$16 S13>0 T6>0",
    ),
    "insert row 7": (
        _Shift(True, 7, 1),
        "B13>0 $C11>10 D1>0 E5>0 $G$13>0 J5>0 L11>0 SUM(N4:N9)>0 P5>0 $R$11:$R$16 S13>0 T5>0",
    ),
    "insert row 11": (
        _Shift(True, 11, 1),
        "B13>0 $C10>10 D1>0 E5>0 $G$13>0 J5>0 L10>0 SUM(N4:N8)>0 P5>0 $R$10:$R$16 S13>0 T5>0",
    ),
    "delete rows 3:6": (
        _Shift(True, 3, -4),
        "B8>0 $C6>10 #REF!>0 E3>0 $G$8>0 J3>0 L8>0 SUM(N3:N4)>0 P3>0 $R$6:$R$11 S8>0 T3>0",
    ),
    "delete rows 4:6": (
        _Shift(True, 4, -3),
        "B9>0 $C7>10 D1>0 E4>0 $G$9>0 J4>0 L9>0 SUM(N4:N5)>0 P4>0 $R$7:$R$12 S9>0 T4>0",
    ),
    "delete rows 6:7": (
        _Shift(True, 6, -2),
        "B10>0 $C8>10 D1>0 E5>0 $G$10>0 J5>0 L8>0 SUM(N4:N6)>0 P5>0 $R$8:$R$13 S10>0 T5>0",
    ),
}


def _add(ws, fragment: str) -> None:
    root = etree.fromstring(f'<w xmlns="{SML_NS}" xmlns:x14="{X14}" xmlns:xm="{XM}">{fragment}</w>')
    insert_worksheet_child(ws._part.element, root[0])


def _rules_sheet():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    for priority, (sqref, formula) in enumerate(_FORMATS, 1):
        _add(
            ws,
            f'<conditionalFormatting sqref="{sqref}"><cfRule type="expression" priority="{priority}">'
            f"<formula>{formula.replace('>', '&gt;')}</formula></cfRule></conditionalFormatting>",
        )
    items = "".join(
        f'<dataValidation type="{kind}" allowBlank="1" sqref="{sqref}"><formula1>'
        f"{formula.replace('>', '&gt;')}</formula1></dataValidation>"
        for sqref, kind, formula in _VALIDATIONS
    )
    _add(ws, f'<dataValidations count="{len(_VALIDATIONS)}">{items}</dataValidations>')
    return wb, ws


def _formulas(ws) -> list[str]:
    root = ws._part.element
    return [f.text for f in root.iterfind("m:conditionalFormatting/m:cfRule/m:formula", NS)] + [
        f.text for f in root.iterfind("m:dataValidations/m:dataValidation/m:formula1", NS)
    ]


def _texts(ws, path: str) -> list[str]:
    return [e.text for e in ws._part.element.iterfind(path, NS)]


def _saved(wb) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


@pytest.mark.parametrize("operation", list(_EXCEL))
def test_format_and_validation_formulas_move_as_excel_moves_them(operation: str):
    shift, expected = _EXCEL[operation]
    wb, ws = _rules_sheet()
    planner = _Planner(shift, sheet_name=ws.name, part=ws._part)
    planner.worksheet(ws._part.element)
    planner.apply()
    assert _formulas(ws) == expected.split()
    assert_valid_package(wb)


def test_insert_rows_and_columns_rewrite_format_and_validation_formulas():
    wb, ws = _rules_sheet()
    ws.insert_rows([[None]], at_row=5)
    assert _formulas(ws) == _EXCEL["insert row 5"][1].split()
    ws.insert_columns([[None]], at_col="C")
    assert _formulas(ws)[:3] == ["B13>0", "$D11>10", "E1>0"]
    assert _formulas(ws)[-3:] == ["$S$11:$S$16", "T13>0", "U6>0"]
    assert_valid_package(wb)


# A format over every row (column) that the shift leaves in place: recorded from Excel.
_WHOLE_LINES = [
    ("A:A", "A1<>Other!A1", _Shift(True, 1, 1), "A1<>Other!A1"),
    ("B:B", "B1>$C$1", _Shift(True, 1, 1), "B1>$C$2"),
    ("B:B", "B1>$C$1", _Shift(True, 1, 2), "B1>$C$3"),
    ("E:E", "SUM(Other!E1:E3)>0", _Shift(True, 1, 1), "SUM(Other!E1:E3)>0"),
    ("1:1", "A1<>Other!A1", _Shift(True, 1, 1), "A2<>Other!A1"),
    ("1:1", "A1<>Other!A1", _Shift(False, 1, 1), "A1<>Other!A1"),
    ("D2:D10", "D1048576<>Other!D1", _Shift(True, 1, 1), "#REF!<>Other!D1"),
    ("F1:F10", "F1<>Other!F1", _Shift(True, 1, 1), "F2<>Other!F1"),
    ("A:A", "A1<>Other!A1", _Shift(True, 1, -1), "A1<>Other!A2"),
    ("B:B", "B1>$C$1", _Shift(True, 1, -1), "B1>#REF!"),
    ("E:E", "SUM(Other!E1:E3)>0", _Shift(True, 1, -1), "SUM(Other!E2:E4)>0"),
    ("D2:D10", "D1048576<>Other!D1", _Shift(True, 1, -1), "D1048575<>Other!D1"),
    ("A:A", "A1<>Other!A1", _Shift(True, 1, -2), "A1<>Other!A3"),
    ("D2:D10", "D1048576<>Other!D1", _Shift(True, 1, -2), "#REF!<>Other!D2"),
]


@pytest.mark.parametrize(("sqref", "formula", "shift", "expected"), _WHOLE_LINES)
def test_formats_at_the_edge_of_the_grid_move_as_excel_moves_them(sqref, formula, shift, expected):
    wb = Workbook.create()
    wb.add_worksheet("Other")
    ws = wb["Sheet1"]
    _add(
        ws,
        f'<conditionalFormatting sqref="{sqref}"><cfRule type="expression" priority="1">'
        f"<formula>{formula.replace('<', '&lt;').replace('>', '&gt;')}</formula></cfRule>"
        "</conditionalFormatting>",
    )
    planner = _Planner(shift, sheet_name=ws.name, part=ws._part)
    planner.worksheet(ws._part.element)
    planner.apply()
    assert _texts(ws, ".//m:cfRule/m:formula") == [expected]


def test_inserting_at_the_top_keeps_a_whole_column_comparison_with_another_sheet():
    wb = Workbook.create()
    wb.add_worksheet("Other")
    ws = wb["Sheet1"]
    _add(
        ws,
        '<conditionalFormatting sqref="A:A"><cfRule type="expression" priority="1">'
        "<formula>A1&lt;&gt;Other!A1</formula></cfRule></conditionalFormatting>",
    )
    _add(
        ws,
        '<dataValidations count="1"><dataValidation type="custom" sqref="A:A">'
        "<formula1>A1&lt;&gt;[1]Data!A1</formula1></dataValidation></dataValidations>",
    )
    ws.insert_rows([[None]], at_row=1)
    assert _texts(ws, ".//m:cfRule/m:formula") == ["A1<>Other!A1"]
    assert _texts(ws, ".//m:formula1") == ["A1<>[1]Data!A1"]


def test_a_whole_column_format_keeps_pointing_at_its_own_cells():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    _add(
        ws,
        '<conditionalFormatting sqref="A:A"><cfRule type="expression" priority="1">'
        "<formula>A1&gt;$B$1</formula></cfRule></conditionalFormatting>",
    )
    ws.insert_rows([[None]], at_row=1)
    assert _texts(ws, ".//m:cfRule/m:formula") == ["A1>$B$2"]


def test_string_literals_and_structured_references_in_formulas_stay():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    _add(
        ws,
        '<conditionalFormatting sqref="B10:B20"><cfRule type="expression" priority="1">'
        '<formula>ISNUMBER(SEARCH("B10",B10))</formula><formula>B10&gt;SUM(Sales[B10])</formula>'
        "</cfRule></conditionalFormatting>",
    )
    ws.insert_rows([[None]], at_row=5)
    assert _texts(ws, ".//m:cfRule/m:formula") == ['ISNUMBER(SEARCH("B10",B11))', "B11>SUM(Sales[B10])"]


_X14_FORMULAS = (
    '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings>'
    '<x14:conditionalFormatting><x14:cfRule type="expression" priority="1" '
    'id="{00000000-0000-0000-0000-000000000001}"><xm:f>Other!$A$1+$B$10&gt;C10</xm:f><x14:dxf/>'
    "</x14:cfRule><xm:sqref>C10:C20</xm:sqref></x14:conditionalFormatting>"
    "</x14:conditionalFormattings></ext>"
    '<ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"><x14:dataValidations count="1">'
    '<x14:dataValidation type="custom"><x14:formula1><xm:f>COUNTIF(Sheet1!$H$10:$H$12,D10)&gt;0</xm:f>'
    "</x14:formula1><xm:sqref>D10</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>"
)


def test_x14_formulas_with_several_references_move():
    wb = Workbook.create()
    wb.add_worksheet("Other")
    ws = wb["Sheet1"]
    _add(ws, _X14_FORMULAS)
    ws.insert_rows([[None]], at_row=5)
    assert _texts(ws, ".//xm:f") == ["Other!$A$1+$B$11>C11", "COUNTIF(Sheet1!$H$11:$H$13,D11)>0"]
    assert _texts(ws, ".//xm:sqref") == ["C11:C21", "D11"]
    wb["Other"].insert_rows([[None]], at_row=1)
    assert _texts(ws, ".//xm:f") == ["Other!$A$2+$B$11>C11", "COUNTIF(Sheet1!$H$11:$H$13,D11)>0"]


def test_another_sheets_formats_and_validations_follow_the_shifted_sheet():
    wb = Workbook.create()
    summary = wb.add_worksheet("Summary")
    _add(
        summary,
        '<conditionalFormatting sqref="B2:B5"><cfRule type="expression" priority="1">'
        "<formula>Sheet1!B10&gt;B2</formula></cfRule></conditionalFormatting>",
    )
    _add(
        summary,
        '<dataValidations count="1"><dataValidation type="list" sqref="C2">'
        "<formula1>Sheet1!$A$10:$A$12</formula1></dataValidation></dataValidations>",
    )
    _add(summary, _X14_FORMULAS.replace("Other!", "Sheet1!"))
    wb["Sheet1"].insert_rows([[None]], at_row=5)
    assert _texts(summary, ".//m:cfRule/m:formula") == ["Sheet1!B11>B2"]
    assert _texts(summary, ".//m:formula1") == ["Sheet1!$A$11:$A$13"]
    assert _texts(summary, ".//xm:f") == ["Sheet1!$A$1+$B$10>C10", "COUNTIF(Sheet1!$H$11:$H$13,D10)>0"]
    assert _texts(summary, ".//xm:sqref") == ["C10:C20", "D10"]


def test_a_refused_insert_leaves_every_formula_alone():
    wb, ws = _rules_sheet()
    wb.add_worksheet("Other")
    _add(ws, _X14_FORMULAS)
    ws["A1048576"].value = "last"
    before = _saved(wb)
    with pytest.raises(GridOverflowError):
        ws.insert_rows([[None]], at_row=5)
    assert_preserved(before, _saved(wb))
