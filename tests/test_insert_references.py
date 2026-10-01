"""What ``insert_rows`` and ``insert_columns`` move besides the cells (issue #5)."""

from __future__ import annotations

import pytest

from xlsxedit import Workbook
from xlsxedit.row_shift import rewrite_formula_refs
from tests.test_insert_rows import _add_defined_name, _defined_name_text


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
