"""insert_rows shifts hyperlinks, dataValidation, and definedNames."""

from __future__ import annotations

from lxml import etree

from xlsxedit import Workbook
from xlsxedit.opc.constants import SML_NS
from xlsxedit.row_shift import shift_dollar_ref
from xlsxedit.worksheet_order import insert_worksheet_child

_DV = f"{{{SML_NS}}}dataValidation"
_DVS = f"{{{SML_NS}}}dataValidations"
_DEFINED_NAMES = f"{{{SML_NS}}}definedNames"
_DEFINED_NAME = f"{{{SML_NS}}}definedName"
_SHEETS = f"{{{SML_NS}}}sheets"


def _add_list_validation(ws, sqref: str, formula: str = '"yes,no"') -> None:
    dvs = etree.Element(_DVS)
    dvs.set("count", "1")
    dv = etree.SubElement(dvs, _DV)
    dv.set("type", "list")
    dv.set("sqref", sqref)
    f1 = etree.SubElement(dv, f"{{{SML_NS}}}formula1")
    f1.text = formula
    insert_worksheet_child(ws._part.element, dvs)


def _dv_sqref(ws) -> str | None:
    dv = ws._part.element.find(f".//{_DV}")
    return None if dv is None else dv.get("sqref")


def _add_defined_name(wb, name: str, formula: str, local_sheet_id: int | None = None) -> None:
    root = wb._workbook_part.element
    block = root.find(_DEFINED_NAMES)
    if block is None:
        block = etree.Element(_DEFINED_NAMES)
        sheets = root.find(_SHEETS)
        if sheets is not None:
            sheets.addnext(block)
        else:
            root.append(block)
    elm = etree.SubElement(block, _DEFINED_NAME)
    elm.set("name", name)
    if local_sheet_id is not None:
        elm.set("localSheetId", str(local_sheet_id))
    elm.text = formula


def _defined_name_text(wb, name: str) -> str | None:
    block = wb._workbook_part.element.find(_DEFINED_NAMES)
    if block is None:
        return None
    for elm in block.findall(_DEFINED_NAME):
        if elm.get("name") == name:
            return elm.text
    return None


def test_shift_dollar_ref_print_area_and_cell():
    assert shift_dollar_ref("$A$10", 5, 2) == "$A$12"
    assert shift_dollar_ref("$A$1:$C$10", 5, 2) == "$A$1:$C$12"
    assert shift_dollar_ref("$A$10", 12, 2) == "$A$10"


def test_insert_rows_shifts_hyperlink():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A10"].value = "Website"
    ws["A10"].hyperlink.url = "https://example.com/"
    ws.insert_rows([["new1"], ["new2"]], at_row=5)
    assert ws["A12"].value == "Website"
    assert ws["A12"].hyperlink.url == "https://example.com/"
    assert ws["A10"].value is None
    assert ws["A10"].hyperlink.url is None


def test_insert_rows_shifts_data_validation():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A10"].value = "Choose:"
    _add_list_validation(ws, "B10")
    ws.insert_rows([["new1"], ["new2"]], at_row=5)
    assert ws["A12"].value == "Choose:"
    assert _dv_sqref(ws) == "B12"


def test_insert_rows_shifts_defined_names():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A10"].value = "Target"
    wb.add_worksheet("Sheet2")
    _add_defined_name(wb, "TargetCell", "Sheet1!$A$10")
    _add_defined_name(wb, "_xlnm.Print_Area", "Sheet1!$A$1:$C$10", local_sheet_id=0)
    _add_defined_name(wb, "Other", "Sheet2!$A$10")
    ws.insert_rows([["new1"], ["new2"]], at_row=5)
    assert ws["A12"].value == "Target"
    assert _defined_name_text(wb, "TargetCell") == "Sheet1!$A$12"
    assert _defined_name_text(wb, "_xlnm.Print_Area") == "Sheet1!$A$1:$C$12"
    assert _defined_name_text(wb, "Other") == "Sheet2!$A$10"


def test_insert_rows_below_named_cell_leaves_defined_name():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A10"].value = "Target"
    _add_defined_name(wb, "TargetCell", "Sheet1!$A$10")
    ws.insert_rows([["x"], ["y"]], at_row=12)
    assert ws["A10"].value == "Target"
    assert _defined_name_text(wb, "TargetCell") == "Sheet1!$A$10"


def test_insert_rows_copies_template_row_merges():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A5"].value = "band"
    ws.merge_cells("A5:C5")
    ws["A8"].value = "Below"
    ws.insert_rows([["row A"], ["row B"]], at_row=6, template_rows=5)
    assert "A5:C5" in ws.merged_ranges
    assert "A6:C6" in ws.merged_ranges
    assert "A7:C7" in ws.merged_ranges
    assert ws["A10"].value == "Below"
