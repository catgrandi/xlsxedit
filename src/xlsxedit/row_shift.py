"""Shift row/column numbers in worksheet range references."""

from __future__ import annotations

from lxml import etree

from xlsxedit.merge import parse_range
from xlsxedit.opc.constants import OFFICE_REL_NS, SML_NS
from xlsxedit.oxml.address import col_to_index, index_to_col, join_address, split_address
from xlsxedit.oxml.parser import parse_xml, serialize_xml

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


def shift_sheet_row_references(ws_element: etree._Element, at_row: int, delta: int) -> None:
    """Shift merge and CF references after row insertion."""
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


def shift_sheet_col_references(ws_element: etree._Element, at_col: int, delta: int) -> None:
    """Shift merge and CF references after column insertion."""
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
