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
