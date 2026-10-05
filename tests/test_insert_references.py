"""What ``insert_rows`` and ``insert_columns`` move besides the cells (issue #5)."""

from __future__ import annotations

import io
from datetime import datetime, timedelta, timezone
from decimal import Decimal

import pytest
from lxml import etree

from xlsxedit import Workbook
from xlsxedit.exceptions import GridOverflowError, InvalidRangeError
from xlsxedit.opc.constants import SML_NS
from xlsxedit.row_shift import rewrite_formula_refs
from xlsxedit.worksheet_order import insert_worksheet_child
from tests.conftest import INSPECT_FIXTURES
from tests.preservation import assert_preserved, read_pkg
from tests.schema_validate import assert_valid_package
from tests.test_insert_rows import _add_defined_name, _defined_name_text

X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
XM = "http://schemas.microsoft.com/office/excel/2006/main"
NS = {"m": SML_NS, "x14": X14, "xm": XM}


def _xml(fragment: str) -> etree._Element:
    root = etree.fromstring(f'<w xmlns="{SML_NS}" xmlns:x14="{X14}" xmlns:xm="{XM}">{fragment}</w>')
    return root[0]


def _add(ws, fragment: str) -> etree._Element:
    elm = _xml(fragment)
    insert_worksheet_child(ws._part.element, elm)
    return elm


def _attr(ws, path: str, attr: str) -> str | None:
    elm = ws._part.element.find(path, NS)
    return None if elm is None else elm.get(attr)


def _texts(ws, path: str) -> list[str]:
    return [e.text for e in ws._part.element.iterfind(path, NS)]


def _saved(wb: Workbook) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# -- defined names and the reference tokenizer --------------------------------


def test_defined_names_never_match_inside_sheet_or_function_names():
    wb = Workbook.create()
    data = wb.rename_worksheet("Sheet1", "Data2024")
    q1 = wb.add_worksheet("Q1")
    _add_defined_name(wb, "_xlnm.Print_Area", "Data2024!$A$1:$C$20", local_sheet_id=0)
    _add_defined_name(wb, "Calc", "LOG10(Data2024!$B$7)+DAYS360($A$1,$A$9)", local_sheet_id=0)
    _add_defined_name(wb, "Here", "Data2024!$A$10")
    _add_defined_name(wb, "There", "'Q1'!$A$10")

    data.insert_rows([[1]], at_row=5)
    assert _defined_name_text(wb, "_xlnm.Print_Area") == "Data2024!$A$1:$C$21"
    assert _defined_name_text(wb, "Calc") == "LOG10(Data2024!$B$8)+DAYS360($A$1,$A$10)"
    assert _defined_name_text(wb, "Here") == "Data2024!$A$11"
    assert _defined_name_text(wb, "There") == "'Q1'!$A$10"

    q1.insert_rows([[1]], at_row=5)
    assert _defined_name_text(wb, "There") == "'Q1'!$A$11"
    assert _defined_name_text(wb, "Here") == "Data2024!$A$11"
    assert _defined_name_text(wb, "Calc") == "LOG10(Data2024!$B$8)+DAYS360($A$1,$A$10)"


def test_print_titles_whole_rows_and_columns_move_on_their_axis():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    _add_defined_name(wb, "_xlnm.Print_Titles", "Sheet1!$A:$B,Sheet1!$1:$2", local_sheet_id=0)
    ws.insert_rows([[1]], at_row=1)
    assert _defined_name_text(wb, "_xlnm.Print_Titles") == "Sheet1!$A:$B,Sheet1!$2:$3"
    ws.insert_columns([[1]], at_col="A")
    assert _defined_name_text(wb, "_xlnm.Print_Titles") == "Sheet1!$B:$C,Sheet1!$2:$3"


def test_defined_name_local_to_the_sheet_moves_unqualified_references():
    wb = Workbook.create()
    wb.add_worksheet("Other")
    _add_defined_name(wb, "Local", "SUM($A$5:$A$9)", local_sheet_id=0)
    _add_defined_name(wb, "Elsewhere", "SUM($A$5:$A$9)", local_sheet_id=1)
    wb["Sheet1"].insert_rows([[1]], at_row=6)
    assert _defined_name_text(wb, "Local") == "SUM($A$5:$A$10)"
    assert _defined_name_text(wb, "Elsewhere") == "SUM($A$5:$A$9)"


def test_defined_name_pushed_off_the_grid_becomes_a_ref_error():
    wb = Workbook.create()
    _add_defined_name(wb, "Last", "Sheet1!$A$1048576")
    wb["Sheet1"].insert_rows([[1]], at_row=1)
    assert _defined_name_text(wb, "Last") == "Sheet1!#REF!"


def test_names_pointing_at_quoted_quarter_sheets_survive_an_insert():
    wb = Workbook.create()
    summary = wb.rename_worksheet("Sheet1", "Summary")
    for quarter in ("Q1", "Q2", "Sales Q1"):
        wb.add_worksheet(quarter)
    _add_defined_name(wb, "YearTotal", "'Q1'!Total+'Q2'!Total", local_sheet_id=0)
    _add_defined_name(wb, "SalesTotal", "'Sales Q1'!Total*2", local_sheet_id=0)
    summary.insert_rows([[1]], at_row=1)
    summary.insert_columns([[1]], at_col="A")
    assert _defined_name_text(wb, "YearTotal") == "'Q1'!Total+'Q2'!Total"
    assert _defined_name_text(wb, "SalesTotal") == "'Sales Q1'!Total*2"


def _row_insert(at: int):
    return lambda area: area.shift_rows(at, 1)


@pytest.mark.parametrize(
    ("text", "unqualified", "expected"),
    [
        ("'Q1'!Total+'Q2'!Total", True, "'Q1'!Total+'Q2'!Total"),
        ("'Sales Q1'!Total*2+A1", True, "'Sales Q1'!Total*2+A2"),
        ("'FY2024 Q1'!Revenue", True, "'FY2024 Q1'!Revenue"),
        ("'[1]Sales Q1'!Total", True, "'[1]Sales Q1'!Total"),
        ("'Bob''s Q1'!Total+A1", True, "'Bob''s Q1'!Total+A2"),
        ("'Notes Summary!A1 x'!Total", False, "'Notes Summary!A1 x'!Total"),
        ("'Summary'!A1+'Q1'!$A$1", False, "'Summary'!A2+'Q1'!$A$1"),
        ("SUM(Sales[Sales '[EUR'] FY24])+A1", True, "SUM(Sales[Sales '[EUR'] FY24])+A2"),
        ("Table1[[#This Row],[Sales '[EUR'] FY24]]", True, "Table1[[#This Row],[Sales '[EUR'] FY24]]"),
        ("Table1[#All]+Table1[[Col1]:[Col2]]", True, "Table1[#All]+Table1[[Col1]:[Col2]]"),
        ("IF(FY24?,\\Q1,0)+$A$1", True, "IF(FY24?,\\Q1,0)+$A$2"),
        ("Summary!FY24?*2+Summary!$A$1", False, "Summary!FY24?*2+Summary!$A$2"),
        ("SUM('Q1:Q4'!A1)+A1", True, "SUM('Q1:Q4'!A1)+A2"),
    ],
)
def test_rewrite_formula_refs_reads_whole_tokens(text: str, unqualified: bool, expected: str):
    assert rewrite_formula_refs(text, "Summary", _row_insert(1), unqualified=unqualified) == expected


@pytest.mark.parametrize(
    "text",
    ["SUM(Summary:Q4!A5)", "SUM('Summary:Q4'!A5)", "[1]Summary!A5", "'[1]Summary'!A5"],
)
def test_spans_and_other_books_naming_the_sheet_stay(text: str):
    assert rewrite_formula_refs(text, "Summary", _row_insert(1), unqualified=False) == text


# -- tolerant references, refusals and drops ----------------------------------


def test_absolute_and_whole_column_references_move():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    _add(ws, '<mergeCells count="1"><mergeCell ref="$A$10:$B$10"/></mergeCells>')
    _add(ws, '<conditionalFormatting sqref="$A$10:$C$20 F:F"><cfRule type="expression" priority="1"><formula>TRUE</formula></cfRule></conditionalFormatting>')
    _add(ws, '<dataValidations count="1"><dataValidation type="whole" sqref="A:A B10 3:3"><formula1>0</formula1></dataValidation></dataValidations>')
    ws.insert_rows([[1], [2]], at_row=5)
    assert ws.merged_ranges == ["$A$12:$B$12"]
    assert _attr(ws, "m:conditionalFormatting", "sqref") == "$A$12:$C$22 F:F"
    assert _attr(ws, "m:dataValidations/m:dataValidation", "sqref") == "A:A B12 3:3"
    assert_valid_package(wb)


@pytest.mark.parametrize(
    "fragment",
    [
        '<mergeCells count="1"><mergeCell ref="A1:B"/></mergeCells>',
        '<conditionalFormatting sqref="C3 A1:"><cfRule type="expression" priority="1"><formula>TRUE</formula></cfRule></conditionalFormatting>',
        '<dataValidations count="1"><dataValidation sqref="ZZZZ9"/></dataValidations>',
        '<hyperlinks><hyperlink ref="A1:" display="x"/></hyperlinks>',
    ],
    ids=["merge", "cf", "dv", "hyperlink"],
)
def test_a_malformed_reference_refuses_the_insert_before_changing_anything(
    fragment: str, unchecked_workbooks
):
    wb = Workbook.open(INSPECT_FIXTURES["SimpleFormula"])
    ws = wb["Sheet1"]
    _add(ws, fragment)
    before = _saved(wb)
    with pytest.raises(InvalidRangeError):
        ws.insert_rows([[1]], at_row=3)
    with pytest.raises(InvalidRangeError):
        ws.insert_columns([[1]], at_col="B")
    assert_preserved(before, _saved(wb))


def test_a_column_span_it_cannot_read_refuses_insert_columns(unchecked_workbooks):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    _add(ws, '<cols><col min="a" max="2" width="9"/></cols>')
    before = _saved(wb)
    with pytest.raises(InvalidRangeError):
        ws.insert_columns([[1]], at_col="A")
    assert_preserved(before, _saved(wb))


def test_content_pushed_past_the_last_row_refuses_the_insert():
    wb = Workbook.open(INSPECT_FIXTURES["ConditionalFormatting"])
    ws = wb["Sheet1"]
    ws["A1048576"].value = 1
    before = _saved(wb)
    with pytest.raises(GridOverflowError, match="cell A1048576 past row 1048576"):
        ws.insert_rows([[1]], at_row=3)
    assert_preserved(before, _saved(wb))


def test_content_pushed_past_column_xfd_refuses_the_insert():
    wb = Workbook.open(INSPECT_FIXTURES["ConditionalFormatting"])
    ws = wb["Sheet1"]
    ws["XFD3"].value = "edge"
    before = _saved(wb)
    with pytest.raises(GridOverflowError, match="cell XFD3 past column XFD"):
        ws.insert_columns([[1]], at_col="B")
    assert_preserved(before, _saved(wb))


def test_a_table_pushed_off_the_grid_refuses_the_insert(unchecked_workbooks):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1048575"].value = "x"
    ws["B1048575"].value = "y"
    ws.add_table("A1048575:B1048576", ["x", "y"])
    ws["A1048575"].value = None
    ws["B1048575"].value = None
    with pytest.raises(GridOverflowError, match="table"):
        ws.insert_rows([[1], [2], [3]], at_row=5)


@pytest.mark.parametrize(
    ("method", "kwargs"),
    [
        ("insert_rows", {"rows": [[1]], "at_row": 5}),
        ("insert_columns", {"cols": [[1]], "at_col": "A"}),
    ],
)
def test_a_table_partly_pushed_off_the_grid_refuses_the_insert(method, kwargs, unchecked_workbooks):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    corner = "A1048575" if method == "insert_rows" else "XFC1"
    far = "B1048576" if method == "insert_rows" else "XFD2"
    ws[corner].value = "x"
    other = "B1048575" if method == "insert_rows" else "XFD1"
    ws[other].value = "y"
    ws.add_table(f"{corner}:{far}", ["x", "y"])
    ws[corner].value = None
    ws[other].value = None
    before = _saved(wb)
    with pytest.raises(GridOverflowError, match="table"):
        getattr(ws, method)(**kwargs)
    assert_preserved(before, _saved(wb))


@pytest.mark.parametrize(
    ("method", "kwargs"),
    [
        ("insert_rows", {"rows": [[1], [2]], "at_row": 1_048_576}),
        ("insert_rows", {"rows": [[1, 2]], "at_cell": "XFD1"}),
        ("insert_columns", {"cols": [[1], [2]], "at_col": "XFD"}),
        ("insert_columns", {"cols": [[1, 2]], "at_cell": "A1048576"}),
    ],
)
def test_writing_past_the_grid_is_refused(method: str, kwargs: dict):
    wb = Workbook.create()
    with pytest.raises(GridOverflowError, match="would write past"):
        getattr(wb["Sheet1"], method)(**kwargs)


@pytest.mark.parametrize(
    ("method", "kwargs"),
    [
        ("insert_rows", {"rows": [[1]], "at_row": 0}),
        ("insert_columns", {"cols": [[1]], "at_col": -1}),
    ],
)
def test_an_insert_point_off_the_grid_is_refused(method, kwargs):
    from xlsxedit import InvalidRangeError as PublicInvalidRangeError

    with pytest.raises(PublicInvalidRangeError):
        getattr(Workbook.create()["Sheet1"], method)(**kwargs)


def test_grid_overflow_error_is_public():
    import xlsxedit

    assert xlsxedit.GridOverflowError is GridOverflowError
    assert issubclass(GridOverflowError, InvalidRangeError)


def test_blank_cells_and_areas_pushed_off_the_grid_are_dropped():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["B1048576"]._element.set("s", "0")
    _add(ws, '<conditionalFormatting sqref="A1048576 A1"><cfRule type="expression" priority="1"><formula>TRUE</formula></cfRule></conditionalFormatting>')
    _add(ws, '<conditionalFormatting sqref="D1048576"><cfRule type="expression" priority="2"><formula>TRUE</formula></cfRule></conditionalFormatting>')
    _add(ws, '<dataValidations count="2"><dataValidation sqref="C1048575:C1048576"/><dataValidation sqref="E1048576"/></dataValidations>')
    _add(ws, '<mergeCells count="1"><mergeCell ref="F1048575:F1048576"/></mergeCells>')
    ws.insert_rows([[1]], at_row=1)
    assert ws._find_cell_element("B1048576") is None
    assert _attr(ws, "m:conditionalFormatting", "sqref") == "A2"
    assert len(ws.conditional_formatting) == 1
    assert _attr(ws, "m:dataValidations", "count") == "1"
    assert _attr(ws, "m:dataValidations/m:dataValidation", "sqref") == "C1048576:C1048576"
    assert ws.merged_ranges == []
    assert ws._part.element.find("m:mergeCells", NS) is None
    assert_valid_package(wb)


def test_a_hyperlink_pushed_off_the_grid_releases_its_relationship():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1048576"].hyperlink.url = "https://example.com/"
    ws["A1048576"].value = None
    assert len(ws._part.rels._rels) == 1
    ws.insert_rows([[1]], at_row=1)
    assert ws._part.element.find("m:hyperlinks", NS) is None
    assert len(ws._part.rels._rels) == 0


def test_column_widths_are_clamped_at_xfd():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    _add(ws, '<cols><col min="2" max="3" width="20" customWidth="1"/><col min="5" max="16384" width="9" customWidth="1"/></cols>')
    ws.insert_columns([[1]], at_col="C")
    assert [(c.get("min"), c.get("max")) for c in ws._part.element.iterfind("m:cols/m:col", NS)] == [
        ("2", "4"),
        ("6", "16384"),
    ]
    assert_valid_package(wb)


def test_a_column_span_pushed_off_the_grid_is_dropped():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    _add(ws, '<cols><col min="1" max="1" width="20" customWidth="1"/><col min="16384" max="16384" width="9" customWidth="1"/></cols>')
    ws.insert_columns([[1]], at_col="B")
    assert [(c.get("min"), c.get("max")) for c in ws._part.element.iterfind("m:cols/m:col", NS)] == [
        ("1", "1")
    ]
    assert_valid_package(wb)


def test_inserting_blank_rows_and_columns():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "head"
    assert ws.insert_rows([[], []], at_row=1) == 2
    assert ws["A3"].value == "head" and ws["A1"].value is None
    assert ws.insert_columns([[]], at_col="A") == 1
    assert ws["B3"].value == "head"
    assert_valid_package(wb)


@pytest.mark.parametrize(
    ("method", "kwargs", "error"),
    [
        ("insert_rows", {"rows": [["ok", Decimal("1.5")]], "at_row": 5}, TypeError),
        ("insert_rows", {"rows": [["bad \x01 text"]], "at_row": 5}, ValueError),
        ("insert_columns", {"cols": [[1, object()]], "at_col": "B"}, TypeError),
        ("insert_columns", {"cols": [["\ufffe"]], "at_col": "B"}, ValueError),
        ("insert_rows", {"rows": [[1]], "at_row": 2.0}, TypeError),
        ("insert_columns", {"cols": [[1]], "at_col": "B1"}, InvalidRangeError),
        ("insert_columns", {"cols": [[1]], "at_col": "B "}, InvalidRangeError),
    ],
)
def test_a_value_or_argument_it_cannot_use_refuses_before_changing_anything(method, kwargs, error):
    wb = Workbook.open(INSPECT_FIXTURES["SimpleFormula"])
    before = _saved(wb)
    with pytest.raises(error):
        getattr(wb["Sheet1"], method)(**kwargs)
    assert_preserved(before, _saved(wb))


def test_strings_need_a_shared_strings_part_before_anything_changes():
    wb = Workbook.open(INSPECT_FIXTURES["SimpleFormula"])
    wb._shared_strings = None
    before = _saved(wb)
    with pytest.raises(ValueError, match="no shared strings part"):
        wb["Sheet1"].insert_rows([["text"]], at_row=3)
    assert_preserved(before, _saved(wb))


@pytest.mark.parametrize("method", ["insert_rows", "insert_columns"])
@pytest.mark.parametrize(
    ("value", "error"),
    [
        (datetime(2024, 1, 1, tzinfo=timezone.utc), TypeError),
        (datetime(2024, 1, 1, tzinfo=timezone(timedelta(hours=2))), TypeError),
    ],
)
def test_a_timezone_aware_datetime_refuses_before_changing_anything(method, value, error):
    wb = Workbook.open(INSPECT_FIXTURES["SimpleFormula"])
    before = _saved(wb)
    ws = wb["Sheet1"]
    with pytest.raises(error):
        if method == "insert_rows":
            ws.insert_rows([["ok", value]], at_row=5)
        else:
            ws.insert_columns([["ok", value]], at_col="B")
    assert_preserved(before, _saved(wb))


def test_pandas_nat_refuses_before_changing_anything():
    pd = pytest.importorskip("pandas")
    wb = Workbook.open(INSPECT_FIXTURES["SimpleFormula"])
    before = _saved(wb)
    with pytest.raises(ValueError, match="use None"):
        wb["Sheet1"].insert_rows([["ok", pd.NaT]], at_row=5)
    assert_preserved(before, _saved(wb))


@pytest.mark.parametrize("method", ["insert_rows", "insert_columns"])
@pytest.mark.parametrize(
    "styles",
    [
        {"row_styles": [{"bg_color": "notacolour"}]},
        {"column_styles": [{"font_color": "#12"}]},
        {"column_styles": [{"font_size": "big"}]},
        {"row_styles": [{}, {"bg_color": 12}]},
        {"template_rows": 3, "column_styles": [{"font_color": "nope"}]},
        {"column_styles": [{"num_format": 5}]},
        {"column_styles": [{"num_format": "0\x01"}]},
        {"row_styles": [{"font_name": 7}]},
        {"row_styles": [{"horizontal_align": "middle"}]},
        {"column_styles": [{"vertical_align": 1}]},
    ],
)
def test_a_bad_inline_style_refuses_before_changing_anything(method, styles):
    wb = Workbook.open(INSPECT_FIXTURES["SimpleFormula"])
    before = _saved(wb)
    ws = wb["Sheet1"]
    with pytest.raises((ValueError, TypeError)):
        if method == "insert_rows":
            ws.insert_rows([["ok", 1]], at_row=5, **styles)
        else:
            ws.insert_columns([["ok", 1]], at_col="B", **styles)
    assert_preserved(before, _saved(wb))


# -- formula ranges and recalculation -----------------------------------------


def test_data_table_ranges_and_input_cells_move():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    f = etree.SubElement(ws["C3"]._element, f"{{{SML_NS}}}f")
    f.attrib.update({"t": "dataTable", "ref": "C3:D9", "dt2D": "1", "dtr": "1", "r1": "A1", "r2": "B12"})
    ws.insert_rows([[None], [None]], at_row=5)
    assert dict(f.attrib) == {
        "t": "dataTable",
        "ref": "C3:D11",
        "dt2D": "1",
        "dtr": "1",
        "r1": "A1",
        "r2": "B14",
    }
    ws.insert_columns([[None]], at_col="B")
    assert (f.get("ref"), f.get("r1"), f.get("r2")) == ("D3:E11", "A1", "C14")


def test_a_data_table_input_cell_pushed_off_refuses_the_insert():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    f = etree.SubElement(ws["C3"]._element, f"{{{SML_NS}}}f")
    f.attrib.update({"t": "dataTable", "ref": "C3:D9", "dt2D": "0", "dtr": "0", "r1": "A1048576"})
    before = _saved(wb)
    with pytest.raises(GridOverflowError, match="dataTable"):
        ws.insert_rows([[None]], at_row=5)
    assert_preserved(before, _saved(wb))


def test_shared_formula_range_moves_with_its_followers():
    wb = Workbook.open(INSPECT_FIXTURES["SimpleFormula"])
    ws = wb["Sheet1"]
    ws.insert_columns([[None]], at_col="C")
    master = ws._find_cell_element("D4").find(f"{{{SML_NS}}}f")
    assert (master.get("ref"), master.text) == ("D4:D7", "A4*B4")
    ws.insert_rows([[None]], at_row=6)
    assert master.get("ref") == "D4:D8"
    assert ws._find_cell_element("D8").find(f"{{{SML_NS}}}f").get("si") == "0"


@pytest.mark.parametrize("method", ["insert_rows", "insert_columns"])
def test_both_entry_points_drop_the_calc_chain_and_request_a_full_recalc(method: str):
    wb = Workbook.open(INSPECT_FIXTURES["SimpleFormula"])
    before = read_pkg(wb)
    ws = wb["Sheet1"]
    if method == "insert_rows":
        ws.insert_rows([[1]], at_row=5)
    else:
        ws.insert_columns([[1]], at_col="B")
    after = read_pkg(wb)
    assert_preserved(
        before,
        after,
        expected_changed={"xl/worksheets/sheet1.xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels"},
        expected_removed={"xl/calcChain.xml"},
    )
    assert b'fullCalcOnLoad="1"' in after["xl/workbook.xml"]


def test_formula_text_is_not_rewritten():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    etree.SubElement(ws["A1"]._element, f"{{{SML_NS}}}f").text = "B10*2"
    _add(ws, '<conditionalFormatting sqref="C1"><cfRule type="expression" priority="1"><formula>$B$10&gt;0</formula></cfRule></conditionalFormatting>')
    _add(ws, '<dataValidations count="1"><dataValidation type="list" sqref="D1"><formula1>$B$10:$B$12</formula1></dataValidation></dataValidations>')
    ws.insert_rows([[None]], at_row=5)
    assert ws["A1"]._element.find(f"{{{SML_NS}}}f").text == "B10*2"
    assert _texts(ws, ".//m:cfRule/m:formula") == ["$B$10>0"]
    assert _texts(ws, ".//m:formula1") == ["$B$10:$B$12"]


# -- x14 extensions and single-reference text ---------------------------------


def test_x14_conditional_formatting_moves_with_its_base_rule_on_column_insert():
    wb = Workbook.open(INSPECT_FIXTURES["ConditionalFormatting"])
    ws = wb["Sheet1"]
    ws.insert_columns([[None]], at_col="B")
    assert [cf.cell_range for cf in ws.conditional_formatting] == [
        "A2:A7",
        "D2:D7",
        "F2:F7",
        "H2:H7",
        "J2:J7",
    ]
    assert _texts(ws, ".//xm:sqref") == ["D2:D7"]


_X14_EXT = """
<extLst>
  <ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}">
    <x14:dataValidations count="3">
      <x14:dataValidation type="list">
        <x14:formula1><xm:f>Lists!$A$1:$A$5</xm:f></x14:formula1>
        <xm:sqref>B10:B20</xm:sqref>
      </x14:dataValidation>
      <x14:dataValidation type="list">
        <x14:formula1><xm:f>Sheet1!$H$10:$H$12</xm:f></x14:formula1>
        <xm:sqref>C10</xm:sqref>
      </x14:dataValidation>
      <x14:dataValidation type="list">
        <x14:formula1><xm:f>Lists!$A$1:$A$5</xm:f></x14:formula1>
        <xm:sqref>D1048576</xm:sqref>
      </x14:dataValidation>
    </x14:dataValidations>
  </ext>
  <ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
    <x14:sparklineGroups>
      <x14:sparklineGroup displayEmptyCellsAs="gap">
        <x14:colorSeries rgb="FF376092"/>
        <x14:sparklines>
          <x14:sparkline><xm:f>Sheet1!A10:E10</xm:f><xm:sqref>F10</xm:sqref></x14:sparkline>
          <x14:sparkline><xm:f>Sheet1!A3:E3</xm:f><xm:sqref>F3</xm:sqref></x14:sparkline>
        </x14:sparklines>
      </x14:sparklineGroup>
    </x14:sparklineGroups>
  </ext>
</extLst>
"""


def test_x14_validations_and_sparklines_move():
    wb = Workbook.create()
    wb.add_worksheet("Lists")
    ws = wb["Sheet1"]
    _add(ws, _X14_EXT)
    ws.insert_rows([[1], [2]], at_row=5)
    assert _texts(ws, ".//x14:dataValidation/xm:sqref") == ["B12:B22", "C12"]
    assert _attr(ws, ".//x14:dataValidations", "count") == "2"
    assert _texts(ws, ".//x14:dataValidation//xm:f") == ["Lists!$A$1:$A$5", "Sheet1!$H$12:$H$14"]
    assert _texts(ws, ".//x14:sparkline/xm:f") == ["Sheet1!A12:E12", "Sheet1!A3:E3"]
    assert _texts(ws, ".//x14:sparkline/xm:sqref") == ["F12", "F3"]
    assert_valid_package(wb)


def test_another_sheets_sparklines_follow_the_shifted_sheet():
    wb = Workbook.create()
    summary = wb.add_worksheet("Summary")
    _add(summary, _X14_EXT.replace("Sheet1!$H", "Summary!$H"))
    wb["Sheet1"].insert_columns([[1]], at_col="B")
    assert _texts(summary, ".//x14:sparkline/xm:f") == ["Sheet1!A10:F10", "Sheet1!A3:F3"]
    assert _texts(summary, ".//x14:sparkline/xm:sqref") == ["F10", "F3"]
    assert _texts(summary, ".//x14:dataValidation//xm:f")[1] == "Summary!$H$10:$H$12"


def test_a_sparkline_group_left_without_sparklines_is_removed():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    _add(
        ws,
        '<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"><x14:sparklineGroups>'
        '<x14:sparklineGroup><x14:colorSeries rgb="FF376092"/><x14:sparklines>'
        "<x14:sparkline><xm:f>Sheet1!A1:E1</xm:f><xm:sqref>F1048576</xm:sqref></x14:sparkline>"
        "</x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>",
    )
    ws.insert_rows([[1]], at_row=1)
    assert ws._part.element.find("m:extLst", NS) is None


def test_a_malformed_xm_sqref_refuses_the_insert(unchecked_workbooks):
    wb = Workbook.open(INSPECT_FIXTURES["SimpleFormula"])
    ws = wb["Sheet1"]
    _add(
        ws,
        '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"><x14:dataValidations count="1">'
        "<x14:dataValidation><xm:sqref>B10:</xm:sqref></x14:dataValidation>"
        "</x14:dataValidations></ext></extLst>",
    )
    before = _saved(wb)
    with pytest.raises(InvalidRangeError):
        ws.insert_rows([[1]], at_row=3)
    assert_preserved(before, _saved(wb))


def test_conditional_format_value_references_move_with_their_x14_twin():
    wb = Workbook.create()
    wb.add_worksheet("Other")
    ws = wb["Sheet1"]
    _add(
        ws,
        '<conditionalFormatting sqref="C2:C20"><cfRule type="dataBar" priority="1"><dataBar>'
        '<cfvo type="formula" val="$B$10"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>'
        '<extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">'
        "<x14:id>{00000000-0000-0000-0000-000000000001}</x14:id></ext></extLst>"
        "</cfRule></conditionalFormatting>",
    )
    _add(
        ws,
        '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings>'
        "<x14:conditionalFormatting>"
        '<x14:cfRule type="dataBar" id="{00000000-0000-0000-0000-000000000001}"><x14:dataBar>'
        '<x14:cfvo type="formula"><xm:f>$B$10</xm:f></x14:cfvo><x14:cfvo type="autoMax"/>'
        "</x14:dataBar></x14:cfRule>"
        '<x14:cfRule type="cellIs" priority="2" operator="greaterThan" id="{00000000-0000-0000-0000-000000000002}">'
        "<xm:f>Other!$A$5</xm:f><x14:dxf/></x14:cfRule>"
        "<xm:sqref>C2:C20</xm:sqref></x14:conditionalFormatting>"
        "</x14:conditionalFormattings></ext></extLst>",
    )
    ws.insert_rows([[1]], at_row=5)
    assert _attr(ws, ".//m:cfvo", "val") == "$B$11"
    assert _texts(ws, ".//xm:f") == ["$B$11", "Other!$A$5"]
    wb["Other"].insert_rows([[1]], at_row=1)
    assert _texts(ws, ".//xm:f") == ["$B$11", "Other!$A$6"]


def test_another_sheets_cfvo_moves_with_its_x14_twin():
    wb = Workbook.create()
    other = wb.add_worksheet("Other")
    ws = wb["Sheet1"]
    _add(
        ws,
        '<conditionalFormatting sqref="C2:C20"><cfRule type="dataBar" priority="1"><dataBar>'
        '<cfvo type="formula" val="Other!$B$10"/><cfvo type="max"/><color rgb="FF638EC6"/>'
        "</dataBar></cfRule></conditionalFormatting>",
    )
    _add(
        ws,
        '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings>'
        '<x14:conditionalFormatting><x14:cfRule type="dataBar" id="{00000000-0000-0000-0000-000000000003}">'
        '<x14:dataBar><x14:cfvo type="formula"><xm:f>Other!$B$10</xm:f></x14:cfvo>'
        '<x14:cfvo type="autoMax"/></x14:dataBar></x14:cfRule><xm:sqref>C2:C20</xm:sqref>'
        "</x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>",
    )
    other.insert_rows([[1]], at_row=1)
    assert _attr(ws, ".//m:cfvo", "val") == "Other!$B$11"
    assert _texts(ws, ".//xm:f") == ["Other!$B$11"]


def test_hyperlink_locations_follow_the_shifted_sheet():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    other = wb.add_worksheet("Other")
    ws["A1"].hyperlink.location = "A10"
    other["A1"].hyperlink.location = "Sheet1!A10"
    other["A2"].hyperlink.location = "A10"
    ws.insert_rows([[1]], at_row=5)
    assert ws["A1"].hyperlink.location == "A11"
    assert other["A1"].hyperlink.location == "Sheet1!A11"
    assert other["A2"].hyperlink.location == "A10"


def test_external_hyperlink_locations_stay():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    other = wb.add_worksheet("Other")

    def external(sheet, address: str, location: str) -> None:
        sheet[address].hyperlink.url = "Budget.xlsx"
        sheet._part.element.find(f".//m:hyperlink[@ref='{address}']", NS).set("location", location)

    external(ws, "B2", "A10")
    external(ws, "B3", "Sheet1!A10")
    external(other, "B2", "Sheet1!A10")
    other["B4"].hyperlink.location = "Sheet1!A10"
    ws.insert_rows([[1]], at_row=5)
    ws.insert_columns([[1]], at_col="A")
    locations = sorted(
        link.get("location")
        for sheet in (ws, other)
        for link in sheet._part.element.iterfind(".//m:hyperlink", NS)
    )
    assert locations == ["A10", "Sheet1!A10", "Sheet1!A10", "Sheet1!B11"]


def test_large_workbooks_stay_lazy_for_sheets_that_cannot_refer_to_the_shifted_one():
    wb = Workbook.open(INSPECT_FIXTURES["ChartsAndTables"], large=True)
    wb["Table"].insert_rows([[1]], at_row=3)
    assert wb["bar chart"]._part.unparsed_blob is not None
    assert wb["Table"]._part.unparsed_blob is None


def test_large_workbooks_rewrite_deferred_sheets_that_refer_to_the_shifted_one():
    wb = Workbook.create()
    summary = wb.add_worksheet("Summary")
    _add(summary, _X14_EXT.replace("Sheet1!$H", "Summary!$H"))
    summary["A1"].hyperlink.location = "Sheet1!A10"
    buf = io.BytesIO(_saved(wb))
    wb = Workbook.open(buf, large=True)
    assert wb["Summary"]._part.unparsed_blob is not None
    wb["Sheet1"].insert_rows([[1]], at_row=5)
    summary = wb["Summary"]
    assert _texts(summary, ".//x14:sparkline/xm:f") == ["Sheet1!A11:E11", "Sheet1!A3:E3"]
    assert summary["A1"].hyperlink.location == "Sheet1!A11"


# -- filters, page breaks and views -------------------------------------------


def _filtered_sheet() -> Workbook:
    wb = Workbook.create()
    ws = wb["Sheet1"]
    _add(
        ws,
        '<autoFilter ref="A1:D20"><filterColumn colId="2"><filters><filter val="x"/></filters>'
        '</filterColumn><sortState ref="A2:D20"><sortCondition ref="C2:C20"/></sortState></autoFilter>',
    )
    return wb


def test_sheet_auto_filter_and_its_sort_state_move_on_row_insert():
    wb = _filtered_sheet()
    ws = wb["Sheet1"]
    ws.insert_rows([[1]], at_row=5)
    assert _attr(ws, "m:autoFilter", "ref") == "A1:D21"
    assert _attr(ws, "m:autoFilter/m:sortState", "ref") == "A2:D21"
    assert _attr(ws, ".//m:sortCondition", "ref") == "C2:C21"
    assert_valid_package(wb)


@pytest.mark.parametrize(
    ("at_col", "ref", "col_id", "condition"),
    [("A", "B1:E20", "2", "D2:D20"), ("B", "A1:E20", "3", "D2:D20"), ("D", "A1:E20", "2", "C2:C20")],
)
def test_sheet_auto_filter_columns_keep_their_filters_on_column_insert(
    at_col: str, ref: str, col_id: str, condition: str
):
    wb = _filtered_sheet()
    ws = wb["Sheet1"]
    ws.insert_columns([[1]], at_col=at_col)
    assert _attr(ws, "m:autoFilter", "ref") == ref
    assert _attr(ws, "m:autoFilter/m:filterColumn", "colId") == col_id
    assert _attr(ws, ".//m:sortCondition", "ref") == condition


def test_a_sort_state_keeps_its_filter_when_its_last_condition_is_pushed_off():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    _add(
        ws,
        '<autoFilter ref="A1:D1048576"><sortState ref="A2:D1048576">'
        '<sortCondition ref="C1048576"/></sortState></autoFilter>',
    )
    ws.insert_rows([[1]], at_row=1)
    assert _attr(ws, "m:autoFilter", "ref") == "A1:D1048576"
    assert _attr(ws, "m:autoFilter/m:sortState", "ref") == "A3:D1048576"
    assert ws._part.element.find(".//m:sortCondition", NS) is None


@pytest.mark.parametrize(("at_row", "ids"), [(5, ["6", "22"]), (6, ["4", "22"]), (21, ["4", "22"]), (22, ["4", "20"])])
def test_row_breaks_move_with_the_row_after_them(at_row: int, ids: list[str]):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    _add(
        ws,
        '<rowBreaks count="2" manualBreakCount="2"><brk id="4" max="16383" man="1"/>'
        '<brk id="20" max="16383" man="1"/></rowBreaks>',
    )
    ws.insert_rows([[1], [2]], at_row=at_row)
    assert [b.get("id") for b in ws._part.element.iterfind("m:rowBreaks/m:brk", NS)] == ids
    assert_valid_package(wb)


def test_column_breaks_move_and_breaks_pushed_off_are_dropped():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    _add(
        ws,
        '<colBreaks count="2" manualBreakCount="2"><brk id="3" max="1048575" man="1"/>'
        '<brk id="16383" max="1048575" man="1"/></colBreaks>',
    )
    ws.insert_columns([[1]], at_col="B")
    block = ws._part.element.find("m:colBreaks", NS)
    assert [b.get("id") for b in block] == ["4"]
    assert (block.get("count"), block.get("manualBreakCount")) == ("1", "1")
    assert_valid_package(wb)


def _frozen_view(ws, top: str | None) -> None:
    """Rows top..top+2 frozen (top defaults to row 1); the bottom pane starts right below."""
    views = ws._part.element.find("m:sheetViews", NS)
    ws._part.element.remove(views)
    first = 1 if top is None else int(top[1:])
    top_attr = "" if top is None else f' topLeftCell="{top}"'
    _add(
        ws,
        f'<sheetViews><sheetView workbookViewId="0"{top_attr}>'
        f'<pane ySplit="3" topLeftCell="A{first + 3}" activePane="bottomLeft" state="frozen"/>'
        '<selection pane="bottomLeft" activeCell="B10" sqref="B10:C12 A1"/>'
        "</sheetView></sheetViews>",
    )


@pytest.mark.parametrize(
    ("top", "at_row", "view_top", "split", "pane_top"),
    [
        (None, 1, None, "5", "A6"),  # inside the frozen rows: the split grows
        (None, 2, None, "5", "A6"),
        (None, 4, None, "3", "A4"),  # at the first unfrozen row: the new rows stay in sight
        (None, 11, None, "3", "A4"),
        ("A5", 1, "A7", "3", "A10"),  # above a view frozen while scrolled
        ("A5", 5, "A5", "5", "A10"),
        ("A5", 8, "A5", "3", "A8"),
    ],
)
def test_frozen_panes_keep_the_same_rows_frozen(top, at_row, view_top, split, pane_top):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    _frozen_view(ws, top)
    ws.insert_rows([[1], [2]], at_row=at_row)
    assert _attr(ws, ".//m:sheetView", "topLeftCell") == view_top
    assert _attr(ws, ".//m:pane", "ySplit") == split
    assert _attr(ws, ".//m:pane", "topLeftCell") == pane_top
    assert_valid_package(wb)


@pytest.mark.parametrize(
    ("at_row", "active", "sqref"),
    [(1, "B12", "B12:C14 A3"), (4, "B12", "B12:C14 A1"), (11, "B10", "B10:C14 A1")],
)
def test_selections_follow_their_cells(at_row, active, sqref):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    _frozen_view(ws, None)
    ws.insert_rows([[1], [2]], at_row=at_row)
    assert _attr(ws, ".//m:selection", "activeCell") == active
    assert _attr(ws, ".//m:selection", "sqref") == sqref


def test_frozen_columns_grow_on_column_insert():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    views = ws._part.element.find("m:sheetViews", NS)
    ws._part.element.remove(views)
    _add(
        ws,
        '<sheetViews><sheetView workbookViewId="0">'
        '<pane xSplit="2" topLeftCell="C1" activePane="topRight" state="frozen"/>'
        "</sheetView></sheetViews>",
    )
    ws.insert_columns([[1]], at_col="B")
    assert (_attr(ws, ".//m:pane", "xSplit"), _attr(ws, ".//m:pane", "topLeftCell")) == ("3", "D1")
    ws.insert_columns([[1]], at_col="D")
    assert (_attr(ws, ".//m:pane", "xSplit"), _attr(ws, ".//m:pane", "topLeftCell")) == ("3", "D1")


def test_a_selection_at_the_last_row_stays_on_the_grid():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    view = ws._part.element.find("m:sheetViews/m:sheetView", NS)
    view.append(_xml('<selection activeCell="C1048576" sqref="C1048575:C1048576 A1048576"/>'))
    ws.insert_rows([[1]], at_row=1)
    assert _attr(ws, ".//m:selection", "activeCell") == "C1048576"
    assert _attr(ws, ".//m:selection", "sqref") == "C1048576:C1048576 A1048576"
    assert_valid_package(wb)


def test_protected_ranges_ignored_errors_and_cell_watches_move():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    _add(ws, '<protectedRanges><protectedRange name="p" sqref="B5:C9"/></protectedRanges>')
    _add(ws, '<cellWatches><cellWatch r="D7"/></cellWatches>')
    _add(ws, '<ignoredErrors><ignoredError sqref="E2:E9" numberStoredAsText="1"/></ignoredErrors>')
    ws.insert_rows([[1]], at_row=6)
    assert _attr(ws, ".//m:protectedRange", "sqref") == "B5:C10"
    assert _attr(ws, ".//m:cellWatch", "r") == "D8"
    assert _attr(ws, ".//m:ignoredError", "sqref") == "E2:E10"
    assert_valid_package(wb)


def test_custom_views_scenarios_and_smart_tags_move():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    _add(
        ws,
        '<customSheetViews><customSheetView guid="{11111111-1111-1111-1111-111111111111}"'
        ' topLeftCell="A3"><selection activeCell="B9" sqref="B9"/>'
        '<rowBreaks count="1" manualBreakCount="1"><brk id="1048575" max="16383" man="1"/></rowBreaks>'
        '<autoFilter ref="A2:C9"/></customSheetView></customSheetViews>',
    )
    _add(ws, '<smartTags><cellSmartTags r="C7"><cellSmartTag type="0"/></cellSmartTags></smartTags>')
    _add(
        ws,
        '<scenarios current="0" sqref="D5"><scenario name="s" count="1">'
        '<inputCells r="D5" val="1"/></scenario></scenarios>',
    )
    ws.insert_rows([[1]], at_row=2)
    view = ws._part.element.find(".//m:customSheetView", NS)
    assert view is not None and view.get("topLeftCell") == "A4"
    assert _attr(ws, ".//m:customSheetView/m:selection", "sqref") == "B10"
    assert view.find("m:rowBreaks", NS) is None  # its only break moved off the grid
    assert _attr(ws, ".//m:customSheetView/m:autoFilter", "ref") == "A3:C10"
    assert _attr(ws, ".//m:cellSmartTags", "r") == "C8"
    assert (_attr(ws, "m:scenarios", "sqref"), _attr(ws, ".//m:inputCells", "r")) == ("D6", "D6")
    assert_valid_package(wb)


def test_a_scenario_input_cell_pushed_off_refuses_the_insert():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    _add(
        ws,
        '<scenarios><scenario name="s" count="1"><inputCells r="D1048576" val="1"/>'
        "</scenario></scenarios>",
    )
    before = _saved(wb)
    with pytest.raises(GridOverflowError, match="inputCells"):
        ws.insert_rows([[1]], at_row=1)
    assert_preserved(before, _saved(wb))


@pytest.mark.parametrize(
    "fragment",
    [
        '<autoFilter ref="A1:Q"/>',
        '<protectedRanges><protectedRange name="p" sqref="B5:"/></protectedRanges>',
        '<rowBreaks count="1"><brk id="x" max="16383" man="1"/></rowBreaks>',
    ],
    ids=["autofilter", "protected-range", "row-break"],
)
def test_malformed_filter_break_and_protection_refs_refuse_the_insert(
    fragment: str, unchecked_workbooks
):
    wb = Workbook.open(INSPECT_FIXTURES["SimpleFormula"])
    ws = wb["Sheet1"]
    _add(ws, fragment)
    before = _saved(wb)
    with pytest.raises(InvalidRangeError):
        ws.insert_rows([[1]], at_row=3)
    assert_preserved(before, _saved(wb))


# -- rows after insert_columns ------------------------------------------------


def test_row_spans_cover_the_cells_after_insert_columns():
    wb = Workbook.open(INSPECT_FIXTURES["SimpleFormula"])
    ws = wb["Sheet1"]
    ws.insert_columns([[None, 1, 2, 3]], at_col="B")
    assert {r.get("r"): r.get("spans") for r in ws._part.element.iterfind(".//m:row", NS)} == {
        "1": None,
        "2": "1:4",
        "3": "1:4",
        "4": "1:4",
        "5": "1:4",
        "6": "1:4",
        "7": "1:4",
        "8": "1:4",
    }
    ws.insert_columns([[None, 1]], at_col="F")
    spans = {r.get("r"): r.get("spans") for r in ws._part.element.iterfind(".//m:row", NS)}
    assert (spans["2"], spans["3"]) == ("1:6", "1:4")


def _cell_order(ws) -> list[list[str]]:
    return [
        [c.get("r") for c in row.iterfind("m:c", NS)]
        for row in ws._part.element.iterfind("m:sheetData/m:row", NS)
    ]


def test_insert_columns_keeps_cells_in_column_order():
    wb = Workbook.open(INSPECT_FIXTURES["SimpleFormula"])
    ws = wb["Sheet1"]
    ws.insert_columns([["new", 1, 2]], at_col="B")
    for cells in _cell_order(ws):
        letters = [c.rstrip("0123456789") for c in cells]
        assert letters == sorted(letters, key=lambda s: (len(s), s)), cells
    assert ws["B1"].value == "new" and ws["C2"].value == "Quantity"


def test_insert_columns_keeps_cells_before_a_row_extension():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.write_rows([[1, None, 3], [4, None, 6]], at_row=1)
    for row in ws._part.element.iterfind("m:sheetData/m:row", NS):
        row.append(_xml('<extLst><ext uri="{X}"><xm:f>A1</xm:f></ext></extLst>'))
    ws.insert_columns([["b1", "b2"]], at_col="B")
    for row in ws._part.element.iterfind("m:sheetData/m:row", NS):
        assert [etree.QName(child).localname for child in row] == ["c", "c", "c", "c", "extLst"]
    assert_valid_package(wb)


def _rowless(ws) -> None:
    sheet_data = ws._part.element.find("m:sheetData", NS)
    for value in (1, 2):
        row = etree.SubElement(sheet_data, f"{{{SML_NS}}}row")
        cell = etree.SubElement(row, f"{{{SML_NS}}}c")
        etree.SubElement(cell, f"{{{SML_NS}}}v").text = str(value)


@pytest.mark.parametrize(
    ("method", "kwargs", "rows"),
    [
        ("insert_columns", {"cols": [[10, 20]], "at_col": "A"}, [("1", ["A1", "B1"]), ("2", ["A2", "B2"])]),
        ("insert_rows", {"rows": [[10]], "at_row": 1}, [("1", ["A1"]), ("2", ["A2"]), ("3", ["A3"])]),
    ],
)
def test_rows_and_cells_without_r_are_numbered(method, kwargs, rows):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    _rowless(ws)
    getattr(ws, method)(**kwargs)
    found = [
        (row.get("r"), [c.get("r") for c in row.iterfind("m:c", NS)])
        for row in ws._part.element.iterfind("m:sheetData/m:row", NS)
    ]
    assert found == rows


# -- template rows ------------------------------------------------------------


def test_template_row_at_the_insert_row_styles_the_new_rows():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A3"].value = "styled"
    ws["A3"].apply_style(bold=True)
    ws["B3"].apply_style(italic=True)
    styles = (ws["A3"]._element.get("s"), ws["B3"]._element.get("s"))
    ws.insert_rows([["new", 1], ["new", 2]], at_row=3, template_rows=3)
    assert (ws["A3"]._element.get("s"), ws["B3"]._element.get("s")) == styles
    assert (ws["A4"]._element.get("s"), ws["B4"]._element.get("s")) == styles
    assert ws["A5"].value == "styled"


def test_template_row_for_insert_columns_is_read_before_the_shift():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["B1"].apply_style(bold=True)
    style = ws["B1"]._element.get("s")
    ws.insert_columns([("new", "x")], at_col="B", template_rows=1)
    assert ws["B1"]._element.get("s") == style
    assert ws["B2"]._element.get("s") == style
    assert ws["C1"]._element.get("s") == style


def test_template_rows_name_rows_as_they_were_before_the_insert():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A9"].apply_style(bold=True)
    ws["A10"].apply_style(italic=True)
    italic = ws["A10"]._element.get("s")
    ws.merge_cells("B10:C10")
    ws.insert_rows([["new"]], at_row=5, template_rows=10)
    assert ws["A5"]._element.get("s") == italic
    assert "B5:C5" in ws.merged_ranges and "B11:C11" in ws.merged_ranges


def test_a_template_merge_overlapping_a_grown_merge_refuses_the_insert():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.merge_cells("A2:C2")
    ws.merge_cells("B5:B8")
    before = _saved(wb)
    with pytest.raises(ValueError, match="overlaps existing 'B5:B9'"):
        ws.insert_rows([["x"]], at_row=6, template_rows=2)
    assert_preserved(before, _saved(wb))


# -- values aimed at merged cells ---------------------------------------------


def _has_value(ws, address: str) -> bool:
    c = ws._find_cell_element(address)
    return c is not None and any(etree.QName(child).localname in ("v", "f", "is") for child in c)


@pytest.mark.parametrize("row_styles", [None, [{}], [{"bold": True}, {}]])
def test_bulk_writes_into_a_grown_merge_go_to_its_anchor(row_styles):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A3"].value = "label"
    ws.merge_cells("A3:A6")
    ws.insert_rows([[None, 1], [None, 2]], at_row=5, row_styles=row_styles)
    assert ws.merged_ranges == ["A3:A8"]
    assert ws["A3"].value == "label"
    assert not _has_value(ws, "A5") and not _has_value(ws, "A6")
    assert (ws["B5"].value, ws["B6"].value) == (1, 2)
    ws.insert_rows([["relabelled", 3]], at_row=4, row_styles=row_styles)
    assert ws["A3"].value == "relabelled"
    assert not _has_value(ws, "A4")


def test_insert_columns_into_a_horizontal_merge_writes_its_anchor():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["B2"].value = "title"
    ws.merge_cells("B2:D2")
    ws.insert_columns([("top", "inside", "below")], at_col="C")
    assert ws.merged_ranges == ["B2:E2"]
    assert (ws["C1"].value, ws["B2"].value, ws["C3"].value) == ("top", "inside", "below")
    assert not _has_value(ws, "C2")


@pytest.mark.parametrize("row_styles", [None, [{}], [{"bold": True}, {}]])
def test_a_write_that_fills_the_anchor_keeps_its_own_value(row_styles):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.merge_cells("A10:C10")
    ws.merge_cells("A11:C11")
    ws.write_rows([["a", "b", "c"], ["d", None, "f"]], at_row=10, row_styles=row_styles)
    assert (ws["A10"].value, ws["A11"].value) == ("a", "d")
    assert not any(_has_value(ws, a) for a in ("B10", "C10", "B11", "C11"))


def test_several_values_for_one_outside_anchor_leave_the_last():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A2"].value = "Block"
    ws.merge_cells("A2:B4")
    ws.insert_rows([["x", "y"]], at_row=3)
    assert ws["A2"].value == "y"
    assert not _has_value(ws, "A3") and not _has_value(ws, "B3")


def test_writing_none_at_the_anchor_clears_it_and_drops_covered_values():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "old"
    ws.merge_cells("A1:B1")
    ws.write_rows([[None, "x"]], at_row=1)
    assert ws["A1"].value is None
    assert not _has_value(ws, "B1")


def test_template_row_merges_exist_before_the_new_rows_are_written():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.merge_cells("A2:C2")
    ws.insert_rows([["x", "y", "z"]], at_row=6, template_rows=2)
    assert "A6:C6" in ws.merged_ranges
    assert ws["A6"].value == "x"
    assert not _has_value(ws, "B6") and not _has_value(ws, "C6")


def test_a_routed_value_creates_the_anchor_in_column_order():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["C1"].value = "right"
    ws.merge_cells("A1:B4")
    ws.insert_rows([["x"]], at_row=3)
    assert ws["A1"].value == "x"
    assert _cell_order(ws)[0] == ["A1", "C1"]


def test_a_routed_anchor_lands_before_a_row_extension():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.write_rows([[1, None, 3], [4, None, 6]], at_row=1)
    for row in ws._part.element.iterfind("m:sheetData/m:row", NS):
        row.append(_xml('<extLst><ext uri="{X}"><xm:f>A1</xm:f></ext></extLst>'))
    ws.merge_cells("E2:F3")
    ws.insert_rows([[None, None, None, None, None, "routed"]], at_row=3)
    assert ws["E2"].value == "routed"  # the anchor of E2:F4, created in row 2
    row2 = ws._part.element.find("m:sheetData/m:row[@r='2']", NS)
    assert [etree.QName(child).localname for child in row2] == ["c", "c", "c", "c", "extLst"]
    assert_valid_package(wb)


def test_bulk_writes_route_through_whole_column_and_large_merges():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    _add(ws, '<mergeCells count="2"><mergeCell ref="D:E"/><mergeCell ref="G1:H3000"/></mergeCells>')
    assert ws._merge_index._large  # neither merge is indexed cell by cell
    ws.insert_rows([[1, 2, 3, 4, 5, 6, 7, 8]], at_row=3)
    assert ws.merged_ranges == ["D:E", "G1:H3001"]
    assert (ws["D1"].value, ws["G1"].value) == (5, 8)
    assert not any(_has_value(ws, a) for a in ("D3", "E3", "G3", "H3"))


@pytest.mark.parametrize("merge", ["H:I", "3:4"])
def test_whole_column_and_row_merges_work_with_template_rows(merge: str):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.merge_cells("B1:C1")
    ws._part.element.find("m:mergeCells", NS).append(_xml(f'<mergeCell ref="{merge}"/>'))
    ws._invalidate_merge_map()
    ws.insert_rows([["a", "b"]], at_row=2, template_rows=1)
    assert "B2:C2" in ws.merged_ranges
    assert_valid_package(wb)
