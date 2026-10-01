"""Shift row/column numbers in worksheet range references."""

from __future__ import annotations

import re

from lxml import etree

from xlsxedit.exceptions import InvalidRangeError
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
_HYPERLINKS = f"{{{SML_NS}}}hyperlinks"
_HYPERLINK = f"{{{SML_NS}}}hyperlink"
_DATA_VALIDATIONS = f"{{{SML_NS}}}dataValidations"
_DATA_VALIDATION = f"{{{SML_NS}}}dataValidation"
_DEFINED_NAMES = f"{{{SML_NS}}}definedNames"
_DEFINED_NAME = f"{{{SML_NS}}}definedName"

_CELL_TOKEN_RE = re.compile(r"^(\$?)([A-Za-z]+)(\$?)(\d+)$")
_A1_RANGE_RE = re.compile(
    r"\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?"
)
_QUALIFIED_REF_RE = re.compile(
    r"(?P<sheet>'(?:[^']|'')+'|[^'!\s]+)"
    r"!"
    r"(?P<ref>\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)"
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


_TABLE_COLUMNS = f"{{{SML_NS}}}tableColumns"
_TABLE_COLUMN = f"{{{SML_NS}}}tableColumn"
_FILTER_COLUMN = f"{{{SML_NS}}}filterColumn"


def _inner_column_offset(ref: str | None, at_col: int) -> int | None:
    """Where ``at_col`` falls in table ``ref`` when inserting there widens the table.

    That is the case strictly inside the table: an insertion at its first
    column moves the whole table right, and one past its last column misses it.
    """
    if not ref:
        return None
    c1, _, c2, _ = parse_range(ref)
    first, last = col_to_index(c1), col_to_index(c2)
    return at_col - first if first < at_col <= last else None


def _insert_table_columns(table_elm: etree._Element, offset: int, delta: int) -> None:
    """Add ``delta`` ``tableColumn``s before the one at ``offset``, like Excel names them.

    Each gets the next free ``id`` and the first unused ``ColumnN`` name;
    ``autoFilter`` ``filterColumn`` ids at or after ``offset`` move with
    their columns.
    """
    cols_elm = table_elm.find(_TABLE_COLUMNS)
    if cols_elm is None:
        return
    existing = cols_elm.findall(_TABLE_COLUMN)
    next_id = max((int(c.get("id")) for c in existing if c.get("id", "").isdigit()), default=0)
    used = {(c.get("name") or "").casefold() for c in existing}
    anchor = existing[offset] if offset < len(existing) else None
    n = 0
    for _ in range(delta):
        n += 1
        while f"column{n}" in used:
            n += 1
        next_id += 1
        column = etree.Element(_TABLE_COLUMN)
        column.set("id", str(next_id))
        column.set("name", f"Column{n}")
        if anchor is None:
            cols_elm.append(column)
        else:
            anchor.addprevious(column)
    cols_elm.set("count", str(len(existing) + delta))
    af = table_elm.find(_AUTO_FILTER)
    if af is not None:
        for fc in af.iterchildren(_FILTER_COLUMN):
            col_id = fc.get("colId", "")
            if col_id.isdigit() and int(col_id) >= offset:
                fc.set("colId", str(int(col_id) + delta))


def shift_table_parts_cols(worksheet, at_col: int, delta: int) -> None:
    """Shift table ``ref`` / ``autoFilter`` columns after column insertion.

    Columns inserted strictly inside a table join it as new ``tableColumn``s
    (see ``_insert_table_columns``). Raises ``TableError``, before changing
    any table part, when such a table is bound to a query: its query table
    part would not know the new columns.
    """
    from xlsxedit.exceptions import TableError

    tp = worksheet._part.element.find(_TABLE_PARTS)
    if tp is None:
        return
    tables = []
    for rel_elm in tp.findall(_TABLE_PART):
        r_id = rel_elm.get(f"{{{OFFICE_REL_NS}}}id")
        if r_id is None:
            continue
        try:
            part = worksheet._part.rels[r_id].target_part
        except KeyError:
            continue
        table_elm = parse_xml(part.blob)
        tables.append((part, table_elm, _inner_column_offset(table_elm.get("ref"), at_col)))
    for _part, table_elm, offset in tables:
        if offset is not None and any(
            c.get("queryTableFieldId") is not None for c in table_elm.iter(_TABLE_COLUMN)
        ):
            raise TableError(
                f"cannot insert columns inside table {table_elm.get('displayName')!r} "
                f"({table_elm.get('ref')}): its columns are bound to a query "
                "(queryTableFieldId), and the query table part is not updated"
            )
    for part, table_elm, offset in tables:
        if offset is not None:
            _insert_table_columns(table_elm, offset, delta)
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


def _unquote_sheet_name(prefix: str) -> str:
    if prefix.startswith("'") and prefix.endswith("'") and len(prefix) >= 2:
        return prefix[1:-1].replace("''", "'")
    return prefix


def _sheet_prefix_matches(prefix: str, sheet_name: str) -> bool:
    if ":" in prefix and not (prefix.startswith("'") and prefix.endswith("'")):
        return False
    return _unquote_sheet_name(prefix) == sheet_name


def _parse_dollar_cell(token: str) -> tuple[bool, str, bool, int] | None:
    m = _CELL_TOKEN_RE.fullmatch(token)
    if not m:
        return None
    return bool(m.group(1)), m.group(2).upper(), bool(m.group(3)), int(m.group(4))


def _format_dollar_cell(abs_col: bool, col: str, abs_row: bool, row: int) -> str:
    return f"{'$' if abs_col else ''}{col}{'$' if abs_row else ''}{row}"


def shift_dollar_ref(ref: str, at_row: int, delta: int) -> str:
    """Shift rows in an A1 ref, preserving ``$``."""
    parts = ref.split(":")
    cells = [_parse_dollar_cell(p) for p in parts]
    if any(c is None for c in cells):
        return ref
    r2 = cells[-1][3]
    if r2 < at_row:
        return ref
    out = []
    for abs_c, col, abs_r, row in cells:
        out.append(
            _format_dollar_cell(abs_c, col, abs_r, shift_row_number(row, at_row, delta))
        )
    return ":".join(out)


def shift_dollar_ref_cols(ref: str, at_col: int, delta: int) -> str:
    """Shift columns in an A1 ref, preserving ``$`` (``at_col`` is 0-based)."""
    parts = ref.split(":")
    cells = [_parse_dollar_cell(p) for p in parts]
    if any(c is None for c in cells):
        return ref
    c2i = col_to_index(cells[-1][1])
    if c2i < at_col:
        return ref
    out = []
    for abs_c, col, abs_r, row in cells:
        new_idx = shift_col_index(col_to_index(col), at_col, delta)
        out.append(_format_dollar_cell(abs_c, index_to_col(new_idx), abs_r, row))
    return ":".join(out)


def _shift_defined_name_text(
    text: str,
    sheet_name: str,
    local: bool,
    at: int,
    delta: int,
    *,
    cols: bool,
) -> str:
    shifter = shift_dollar_ref_cols if cols else shift_dollar_ref

    def repl_qualified(match: re.Match[str]) -> str:
        prefix = match.group("sheet")
        if not _sheet_prefix_matches(prefix, sheet_name):
            return match.group(0)
        return f"{prefix}!{shifter(match.group('ref'), at, delta)}"

    text = _QUALIFIED_REF_RE.sub(repl_qualified, text)
    if not local:
        return text

    def repl_unqualified(match: re.Match[str]) -> str:
        return shifter(match.group(0), at, delta)

    out = []
    last = 0
    for match in _A1_RANGE_RE.finditer(text):
        start = match.start()
        if start > 0 and text[start - 1] == "!":
            out.append(text[last : match.end()])
            last = match.end()
            continue
        out.append(text[last:start])
        out.append(repl_unqualified(match))
        last = match.end()
    out.append(text[last:])
    return "".join(out)


def _sheet_index(workbook, sheet_name: str) -> int:
    names = list(workbook.sheetnames)
    return names.index(sheet_name)


def shift_defined_names(workbook, sheet_name: str, at_row: int, delta: int) -> None:
    """Shift A1 rows in definedNames that refer to ``sheet_name``."""
    _shift_defined_names(workbook, sheet_name, at_row, delta, cols=False)


def shift_defined_names_cols(workbook, sheet_name: str, at_col: int, delta: int) -> None:
    """Shift A1 columns in definedNames that refer to ``sheet_name``."""
    _shift_defined_names(workbook, sheet_name, at_col, delta, cols=True)


def _shift_defined_names(
    workbook, sheet_name: str, at: int, delta: int, *, cols: bool
) -> None:
    wb_elm = workbook._workbook_part.element
    block = wb_elm.find(_DEFINED_NAMES)
    if block is None:
        return
    try:
        index = _sheet_index(workbook, sheet_name)
    except ValueError:
        return
    for elm in block.findall(_DEFINED_NAME):
        local_id = elm.get("localSheetId")
        local = local_id is not None and int(local_id) == index
        text = elm.text
        if not text:
            continue
        elm.text = _shift_defined_name_text(
            text, sheet_name, local, at, delta, cols=cols
        )
