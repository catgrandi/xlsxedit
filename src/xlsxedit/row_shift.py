"""Shift row/column numbers in worksheet range references."""

from __future__ import annotations

import re
from collections.abc import Callable

from lxml import etree

from xlsxedit.exceptions import InvalidRangeError
from xlsxedit.merge import parse_range
from xlsxedit.opc.constants import OFFICE_REL_NS, SML_NS
from xlsxedit.oxml.address import col_to_index, index_to_col, join_address, split_address
from xlsxedit.oxml.parser import parse_xml, serialize_xml
from xlsxedit.range_set import CellRange

_ROW = f"{{{SML_NS}}}row"
_C = f"{{{SML_NS}}}c"
_COLS = f"{{{SML_NS}}}cols"
_COL = f"{{{SML_NS}}}col"
_MERGE_CELLS = f"{{{SML_NS}}}mergeCells"
_MERGE_CELL = f"{{{SML_NS}}}mergeCell"
_CF = f"{{{SML_NS}}}conditionalFormatting"
_TABLE_PARTS = f"{{{SML_NS}}}tableParts"
_TABLE_PART = f"{{{SML_NS}}}tablePart"
_AUTO_FILTER = f"{{{SML_NS}}}autoFilter"
_HYPERLINKS = f"{{{SML_NS}}}hyperlinks"
_HYPERLINK = f"{{{SML_NS}}}hyperlink"
_DATA_VALIDATIONS = f"{{{SML_NS}}}dataValidations"
_DATA_VALIDATION = f"{{{SML_NS}}}dataValidation"
_DEFINED_NAMES = f"{{{SML_NS}}}definedNames"
_DEFINED_NAME = f"{{{SML_NS}}}definedName"

_SHEETS = f"{{{SML_NS}}}sheets"
_SHEET = f"{{{SML_NS}}}sheet"

# A reference in formula text: a cell, a range, whole columns or whole rows.
_REF = (
    r"(?:\$?[A-Za-z]{1,3}\$?[0-9]+(?::\$?[A-Za-z]{1,3}\$?[0-9]+)?"
    r"|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}"
    r"|\$?[0-9]+:\$?[0-9]+)"
)
# A sheet prefix: quoted, or an unquoted name, or an unquoted 3D span.
_SHEET_NAME = r"(?:'(?:[^']|'')+'|[^\W\d][\w.]*(?::[^\W\d][\w.]*)?)"
# A reference ends where no name (which may hold ? and \), call, structured
# reference or sheet prefix continues.
_REF_END = r"(?![\w.(\[!$?\\])"
_FORMULA_TOKEN_RE = re.compile(
    rf"""
    (?P<string>"(?:[^"]|"")*")
  | (?P<bracket>\[(?:'.|[^\]'])*\])  # ' escapes a bracket inside a structured reference
  | (?P<sheet>{_SHEET_NAME})!(?P<qref>{_REF}{_REF_END}|\#REF!)
  | '(?:[^']|'')*'  # a quoted prefix to a name, never read inside
  | (?P<ref>{_REF}){_REF_END}
  | [\w.$?\\]+
  | .
    """,
    re.VERBOSE | re.DOTALL,
)


def shift_row_number(row: int, at_row: int, delta: int) -> int:
    if row >= at_row:
        return row + delta
    return row


def shift_col_index(col_idx: int, at_col: int, delta: int) -> int:
    """Shift a 0-based column index when inserting at ``at_col``."""
    if col_idx >= at_col:
        return col_idx + delta
    return col_idx


def shift_range_ref(ref: str, at_row: int, delta: int) -> str:
    """Shift row numbers in a range at or below ``at_row``."""
    c1, r1, c2, r2 = parse_range(ref)
    if r2 < at_row:
        return ref
    new_r1 = shift_row_number(r1, at_row, delta)
    new_r2 = shift_row_number(r2, at_row, delta)
    if c1 == c2 and new_r1 == new_r2:
        return join_address(c1, new_r1)
    return f"{join_address(c1, new_r1)}:{join_address(c2, new_r2)}"


def shift_range_ref_cols(ref: str, at_col: int, delta: int) -> str:
    """Shift column letters in a range at or to the right of ``at_col`` (0-based)."""
    c1, r1, c2, r2 = parse_range(ref)
    c1i, c2i = col_to_index(c1), col_to_index(c2)
    if c2i < at_col:
        return ref
    new_c1i = shift_col_index(c1i, at_col, delta)
    new_c2i = shift_col_index(c2i, at_col, delta)
    new_c1, new_c2 = index_to_col(new_c1i), index_to_col(new_c2i)
    if new_c1 == new_c2 and r1 == r2:
        return join_address(new_c1, r1)
    return f"{join_address(new_c1, r1)}:{join_address(new_c2, r2)}"


def shift_sqref(sqref: str, at_row: int, delta: int) -> str:
    return " ".join(shift_range_ref(part, at_row, delta) for part in sqref.split())


def shift_sqref_cols(sqref: str, at_col: int, delta: int) -> str:
    return " ".join(shift_range_ref_cols(part, at_col, delta) for part in sqref.split())


def _shift_attr_ref(ref: str, at: int, delta: int, *, cols: bool) -> str:
    bare = ref.replace("$", "")
    try:
        if cols:
            return shift_range_ref_cols(bare, at, delta)
        return shift_range_ref(bare, at, delta)
    except InvalidRangeError:
        return ref


def _shift_sqref_attr(sqref: str, at: int, delta: int, *, cols: bool) -> str:
    return " ".join(_shift_attr_ref(part, at, delta, cols=cols) for part in sqref.split())


def _shift_hyperlinks_and_validations(
    ws_element: etree._Element, at: int, delta: int, *, cols: bool
) -> None:
    block = ws_element.find(_HYPERLINKS)
    if block is not None:
        for hl in block.findall(_HYPERLINK):
            ref = hl.get("ref")
            if ref:
                hl.set("ref", _shift_sqref_attr(ref, at, delta, cols=cols))
    dvs = ws_element.find(_DATA_VALIDATIONS)
    if dvs is not None:
        for dv in dvs.findall(_DATA_VALIDATION):
            sqref = dv.get("sqref")
            if sqref:
                dv.set("sqref", _shift_sqref_attr(sqref, at, delta, cols=cols))


def shift_sheet_row_references(ws_element: etree._Element, at_row: int, delta: int) -> None:
    """Shift merge, CF, hyperlink, and dataValidation references after row insertion."""
    merge_block = ws_element.find(_MERGE_CELLS)
    if merge_block is not None:
        for merge_elm in merge_block.findall(_MERGE_CELL):
            ref = merge_elm.get("ref")
            if ref:
                merge_elm.set("ref", shift_range_ref(ref, at_row, delta))

    for cf_elm in ws_element.findall(_CF):
        sqref = cf_elm.get("sqref")
        if sqref:
            cf_elm.set("sqref", shift_sqref(sqref, at_row, delta))

    _shift_hyperlinks_and_validations(ws_element, at_row, delta, cols=False)


def shift_sheet_col_references(ws_element: etree._Element, at_col: int, delta: int) -> None:
    """Shift merge, CF, hyperlink, and dataValidation references after column insertion."""
    merge_block = ws_element.find(_MERGE_CELLS)
    if merge_block is not None:
        for merge_elm in merge_block.findall(_MERGE_CELL):
            ref = merge_elm.get("ref")
            if ref:
                merge_elm.set("ref", shift_range_ref_cols(ref, at_col, delta))

    for cf_elm in ws_element.findall(_CF):
        sqref = cf_elm.get("sqref")
        if sqref:
            cf_elm.set("sqref", shift_sqref_cols(sqref, at_col, delta))

    _shift_hyperlinks_and_validations(ws_element, at_col, delta, cols=True)


def shift_table_parts(worksheet, at_row: int, delta: int) -> None:
    """Shift table ``ref`` / ``autoFilter`` for tables linked from the worksheet."""
    tp = worksheet._part.element.find(_TABLE_PARTS)
    if tp is None:
        return
    for rel_elm in tp.findall(_TABLE_PART):
        r_id = rel_elm.get(f"{{{OFFICE_REL_NS}}}id")
        if r_id is None:
            continue
        try:
            part = worksheet._part.rels[r_id].target_part
        except KeyError:
            continue
        table_elm = parse_xml(part.blob)
        ref = table_elm.get("ref")
        if ref:
            table_elm.set("ref", shift_range_ref(ref, at_row, delta))
        af = table_elm.find(_AUTO_FILTER)
        if af is not None and af.get("ref"):
            af.set("ref", shift_range_ref(af.get("ref"), at_row, delta))
        part._blob = serialize_xml(table_elm)


def shift_table_parts_cols(worksheet, at_col: int, delta: int) -> None:
    """Shift table ``ref`` / ``autoFilter`` columns after column insertion."""
    tp = worksheet._part.element.find(_TABLE_PARTS)
    if tp is None:
        return
    for rel_elm in tp.findall(_TABLE_PART):
        r_id = rel_elm.get(f"{{{OFFICE_REL_NS}}}id")
        if r_id is None:
            continue
        try:
            part = worksheet._part.rels[r_id].target_part
        except KeyError:
            continue
        table_elm = parse_xml(part.blob)
        ref = table_elm.get("ref")
        if ref:
            table_elm.set("ref", shift_range_ref_cols(ref, at_col, delta))
        af = table_elm.find(_AUTO_FILTER)
        if af is not None and af.get("ref"):
            af.set("ref", shift_range_ref_cols(af.get("ref"), at_col, delta))
        part._blob = serialize_xml(table_elm)


def shift_sheet_data_rows(sheet_data: etree._Element, at_row: int, delta: int) -> None:
    """Bump ``r`` on rows and cells with row >= ``at_row``, then re-sort."""
    rows = list(sheet_data.findall(_ROW))
    for row_elm in rows:
        row_num = int(row_elm.get("r", "0"))
        if row_num < at_row:
            continue
        new_row = row_num + delta
        row_elm.set("r", str(new_row))
        for c_elm in row_elm.findall(_C):
            addr = c_elm.get("r")
            if not addr:
                continue
            col, _ = split_address(addr)
            c_elm.set("r", join_address(col, new_row))
    for row_elm in rows:
        sheet_data.remove(row_elm)
    rows.sort(key=lambda r: int(r.get("r", "0")))
    for row_elm in rows:
        sheet_data.append(row_elm)


def shift_sheet_data_cols(sheet_data: etree._Element, at_col: int, delta: int) -> None:
    """Bump column letters on cells with col >= ``at_col`` (0-based), re-sort within rows."""
    for row_elm in sheet_data.findall(_ROW):
        row_num = int(row_elm.get("r", "0"))
        cells = list(row_elm.findall(_C))
        for c_elm in cells:
            addr = c_elm.get("r")
            if not addr:
                continue
            col, r = split_address(addr)
            col_idx = col_to_index(col)
            if col_idx < at_col:
                continue
            c_elm.set("r", join_address(index_to_col(col_idx + delta), r if r else row_num))
        for c_elm in cells:
            row_elm.remove(c_elm)
        cells.sort(key=lambda c: col_to_index(split_address(c.get("r", "A1"))[0]))
        for c_elm in cells:
            row_elm.append(c_elm)


def shift_col_dimensions(ws_element: etree._Element, at_col: int, delta: int) -> None:
    """Shift ``<col min/max>`` (1-based) for columns at or to the right of ``at_col``."""
    cols = ws_element.find(_COLS)
    if cols is None:
        return
    at_excel = at_col + 1
    for col_elm in list(cols.findall(_COL)):
        min_c = int(col_elm.get("min", "0"))
        max_c = int(col_elm.get("max", "0"))
        if max_c < at_excel:
            continue
        new_min = min_c + delta if min_c >= at_excel else min_c
        new_max = max_c + delta if max_c >= at_excel else max_c
        col_elm.set("min", str(new_min))
        col_elm.set("max", str(new_max))


def _same_sheet(prefix: str, sheet_name: str | None) -> bool:
    """Whether a reference prefix names exactly ``sheet_name`` (not a 3D span or another book)."""
    if sheet_name is None:
        return False
    if prefix.startswith("'"):
        name = prefix[1:-1].replace("''", "'")
        if ":" in name or "[" in name:
            return False
    elif ":" in prefix:
        return False
    else:
        name = prefix
    return name.casefold() == sheet_name.casefold()


def _move_ref_text(ref: str, move: Callable[[CellRange], CellRange | None]) -> str:
    try:
        area = CellRange.parse(ref)
    except InvalidRangeError:
        return ref  # past the grid, so a name rather than a cell
    moved = move(area)
    return "#REF!" if moved is None else str(moved)


def rewrite_formula_refs(
    text: str,
    sheet_name: str,
    move: Callable[[CellRange], CellRange | None],
    *,
    unqualified: bool,
) -> str:
    """Move the references to ``sheet_name`` in formula ``text``.

    Rewrites references qualified with that sheet, and, when ``unqualified``,
    references without a sheet prefix. A reference is a whole token: never
    part of a sheet name (``Data2024!``), a function name (``LOG10(``,
    ``DAYS360(``), a defined name or a structured reference. String literals,
    other sheets, 3D spans and external workbooks are left alone. A reference
    pushed off the grid becomes ``#REF!``.
    """
    out: list[str] = []
    after_bracket = False
    for m in _FORMULA_TOKEN_RE.finditer(text):
        token = m.group(0)
        qref, ref = m.group("qref"), m.group("ref")
        if qref is not None and qref != "#REF!" and not after_bracket:
            if _same_sheet(m.group("sheet"), sheet_name):
                token = f"{m.group('sheet')}!{_move_ref_text(qref, move)}"
        elif ref is not None and unqualified and text[m.start() - 1 : m.start()] not in ("!", "]"):
            token = _move_ref_text(ref, move)
        out.append(token)
        after_bracket = m.group("bracket") is not None
    return "".join(out)


def _local_sheet_index(wb_elm: etree._Element, sheet_name: str) -> int | None:
    """Position of ``sheet_name`` among ``<sheets>``, the index ``localSheetId`` uses."""
    for index, sheet in enumerate(wb_elm.iterfind(f"{_SHEETS}/{_SHEET}")):
        if sheet.get("name") == sheet_name:
            return index
    return None


def shift_dollar_ref(ref: str, at_row: int, delta: int) -> str:
    """Shift rows in an A1 ref, preserving ``$``; an unparseable ref is returned as is."""
    return _move_ref_text(ref, lambda area: area.shift_rows(at_row, delta))


def shift_dollar_ref_cols(ref: str, at_col: int, delta: int) -> str:
    """Shift columns in an A1 ref, preserving ``$`` (``at_col`` is 0-based)."""
    return _move_ref_text(ref, lambda area: area.shift_cols(at_col + 1, delta))


def shift_defined_names(workbook, sheet_name: str, at_row: int, delta: int) -> None:
    """Shift A1 rows in definedNames that refer to ``sheet_name``."""
    _shift_defined_names(workbook, sheet_name, lambda area: area.shift_rows(at_row, delta))


def shift_defined_names_cols(workbook, sheet_name: str, at_col: int, delta: int) -> None:
    """Shift A1 columns in definedNames that refer to ``sheet_name``."""
    _shift_defined_names(workbook, sheet_name, lambda area: area.shift_cols(at_col + 1, delta))


def _shift_defined_names(
    workbook, sheet_name: str, move: Callable[[CellRange], CellRange | None]
) -> None:
    wb_elm = workbook._workbook_part.element
    block = wb_elm.find(_DEFINED_NAMES)
    if block is None:
        return
    index = _local_sheet_index(wb_elm, sheet_name)
    for elm in block.findall(_DEFINED_NAME):
        if not elm.text:
            continue
        local_id = elm.get("localSheetId")
        local = index is not None and local_id is not None and local_id == str(index)
        elm.text = rewrite_formula_refs(elm.text, sheet_name, move, unqualified=local)
