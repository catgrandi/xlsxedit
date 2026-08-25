"""Worksheet proxy."""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, Any, Iterator

from lxml import etree

from xlsxedit.bulk import write_cell_value
from xlsxedit.bulk_styles import (
    BulkStyleCache,
    merge_style_specs,
    normalize_style_spec,
    normalize_template_rows,
    validate_style_sources,
)
from xlsxedit.cell import Cell, _clear_cell_content
from xlsxedit.conditional_formatting import (
    ConditionalFormatting,
    add_cell_is_rule,
    clone_color_scale_rule,
    expand_conditional_formatting_to_row,
    iter_conditional_formatting,
)
from xlsxedit.dimensions import ColumnDimensions, RowDimensions
from xlsxedit.drawing import (
    Picture,
    Table,
    Chart,
    drawing_parts_for_worksheet,
    iter_charts,
    iter_pictures,
    table_parts_for_worksheet,
)
from xlsxedit.row_shift import (
    shift_col_dimensions,
    shift_defined_names,
    shift_defined_names_cols,
    shift_sheet_col_references,
    shift_sheet_data_cols,
    shift_sheet_data_rows,
    shift_sheet_row_references,
    shift_table_parts,
    shift_table_parts_cols,
)
from xlsxedit.merge import MergeMap, normalize_range, parse_range, ranges_overlap
from xlsxedit.opc.constants import CT, OFFICE_REL_NS, RT, SML_NS
from xlsxedit.oxml.address import col_to_index, index_to_col, join_address, split_address
from xlsxedit.shared_strings import SharedStringTable

if TYPE_CHECKING:
    from xlsxedit.parts import WorksheetPart
    from xlsxedit.workbook import Workbook

_ROW = f"{{{SML_NS}}}row"
_C = f"{{{SML_NS}}}c"
_F = f"{{{SML_NS}}}f"
_SHEET_DATA = f"{{{SML_NS}}}sheetData"
_MERGE_CELLS = f"{{{SML_NS}}}mergeCells"
_MERGE_CELL = f"{{{SML_NS}}}mergeCell"
_DIMENSION = f"{{{SML_NS}}}dimension"


class Worksheet:
    def __init__(self, name: str, part: WorksheetPart, workbook: Workbook):
        self._name = name
        self._part = part
        self._workbook = workbook
        self._merge_map: MergeMap | None = None
        self._column_dimensions = ColumnDimensions(self)
        self._row_dimensions = RowDimensions(self)
        self._bulk_indexes_built = False
        self._row_index: dict[int, etree._Element] = {}
        self._cell_index: dict[str, etree._Element] = {}
        self._bulk_max_row_num = 0

    def __repr__(self) -> str:
        return f"<Worksheet {self._name!r}>"

    @property
    def name(self) -> str:
        """Sheet name (read-only). Rename via :meth:`Workbook.rename_worksheet`."""
        return self._name

    @property
    def shared_strings(self) -> SharedStringTable | None:
        return self._workbook.shared_strings

    @property
    def column_dimensions(self) -> ColumnDimensions:
        return self._column_dimensions

    @property
    def row_dimensions(self) -> RowDimensions:
        return self._row_dimensions

    @property
    def _merge_index(self) -> MergeMap:
        if self._merge_map is None:
            refs = []
            block = self._part.element.find(_MERGE_CELLS)
            if block is not None:
                refs = [m.get("ref", "") for m in block.findall(_MERGE_CELL) if m.get("ref")]
            self._merge_map = MergeMap(refs)
        return self._merge_map

    @property
    def merged_ranges(self) -> list[str]:
        """List of merged-cell ranges on this sheet (e.g. ``["A1:C1"]``)."""
        return list(self._merge_index.refs)

    def _invalidate_merge_map(self) -> None:
        self._merge_map = None

    def _invalidate_bulk_indexes(self) -> None:
        self._bulk_indexes_built = False
        self._row_index.clear()
        self._cell_index.clear()
        self._bulk_max_row_num = 0

    def _build_bulk_indexes(self) -> None:
        if self._bulk_indexes_built:
            return
        self._row_index.clear()
        self._cell_index.clear()
        max_row = 0
        sheet_data = self._sheet_data()
        for row in sheet_data.findall(_ROW):
            row_num = int(row.get("r", "0"))
            if row_num:
                self._row_index[row_num] = row
                max_row = max(max_row, row_num)
            for c_elm in row.findall(_C):
                addr = c_elm.get("r")
                if addr:
                    self._cell_index[addr.upper()] = c_elm
        self._bulk_max_row_num = max_row
        self._bulk_indexes_built = True

    def _bulk_ensure_row(self, row_num: int) -> etree._Element:
        self._build_bulk_indexes()
        existing = self._row_index.get(row_num)
        if existing is not None:
            return existing
        sheet_data = self._sheet_data()
        row = etree.Element(_ROW, nsmap={None: SML_NS})
        row.set("r", str(row_num))
        # Sequential bulk export appends rows in order — avoid O(n²) scan per row.
        if row_num > self._bulk_max_row_num:
            sheet_data.append(row)
        else:
            inserted = False
            for i, child in enumerate(sheet_data):
                if child.tag != _ROW:
                    continue
                existing_num = int(child.get("r", "0"))
                if existing_num > row_num:
                    sheet_data.insert(i, row)
                    inserted = True
                    break
            if not inserted:
                sheet_data.append(row)
        self._row_index[row_num] = row
        if row_num > self._bulk_max_row_num:
            self._bulk_max_row_num = row_num
        return row

    def _bulk_get_or_create_cell(self, address: str, row_elm: etree._Element) -> etree._Element:
        address = address.upper()
        existing = self._cell_index.get(address)
        if existing is not None:
            return existing
        c_elm = etree.SubElement(row_elm, _C)
        c_elm.set("r", address)
        self._cell_index[address] = c_elm
        return c_elm

    def _style_map_for_row(self, row: int) -> dict[int, str]:
        styles: dict[int, str] = {}
        row_elm = self._find_row_element(row)
        if row_elm is None:
            return styles
        for c_elm in row_elm.findall(_C):
            addr = c_elm.get("r")
            s = c_elm.get("s")
            if addr is None or s is None:
                continue
            col, _ = split_address(addr)
            styles[col_to_index(col)] = s
        return styles

    def _resolve_write_address(self, address: str) -> str:
        anchor = self._merge_index.anchor_for(address)
        if anchor is not None and anchor.upper() != address.upper():
            return anchor
        return address

    def _sheet_data(self) -> etree._Element:
        sheet_data = self._part.element.find(_SHEET_DATA)
        if sheet_data is None:
            sheet_data = etree.SubElement(self._part.element, _SHEET_DATA)
        return sheet_data

    def _find_row_element(self, row_num: int) -> etree._Element | None:
        for row in self._sheet_data().findall(_ROW):
            if row.get("r") == str(row_num):
                return row
        return None

    def _find_cell_element(self, address: str) -> etree._Element | None:
        for row in self._sheet_data().findall(_ROW):
            for c_elm in row.findall(_C):
                if c_elm.get("r") == address:
                    return c_elm
        return None

    def _ensure_row(self, row_num: int) -> etree._Element:
        sheet_data = self._sheet_data()
        rows = list(sheet_data.findall(_ROW))
        for row in rows:
            if row.get("r") == str(row_num):
                return row
        row = etree.Element(_ROW, nsmap={None: SML_NS})
        row.set("r", str(row_num))
        inserted = False
        for i, existing in enumerate(rows):
            if int(existing.get("r", "0")) > row_num:
                sheet_data.insert(i, row)
                inserted = True
                break
        if not inserted:
            sheet_data.append(row)
        return row

    def _get_or_create_cell_element(self, address: str) -> etree._Element:
        address = self._resolve_write_address(address)
        existing = self._find_cell_element(address)
        if existing is not None:
            return existing
        col, row_num = split_address(address)
        row_elm = self._ensure_row(row_num)
        c_elm = etree.SubElement(row_elm, _C)
        c_elm.set("r", address)
        return c_elm

    def __getitem__(self, address: str) -> Cell:
        return Cell(self._get_or_create_cell_element(address), self)

    def iter_cells(self) -> Iterator[Cell]:
        sheet_data = self._part.element.find(_SHEET_DATA)
        if sheet_data is None:
            return
        for row in sheet_data.findall(_ROW):
            for c_elm in row.findall(_C):
                yield Cell(c_elm, self)

    def iter_rows(self) -> Iterator[tuple[int, list[Cell]]]:
        """Yield ``(row_number, cells)`` for each populated row."""
        sheet_data = self._part.element.find(_SHEET_DATA)
        if sheet_data is None:
            return
        for row_elm in sheet_data.findall(_ROW):
            row_num = int(row_elm.get("r", "0"))
            if not row_num:
                continue
            yield row_num, [Cell(c_elm, self) for c_elm in row_elm.findall(_C)]

    def values(
        self, *, max_rows: int | None = None, errors_as_nan: bool = False
    ) -> list[list]:
        """Dense row-major grid of cell values (empty cells → ``""``).

        Walks ``sheetData`` in Excel row order and inserts blank rows for gaps
        between row numbers. Formula cells use the cached calculated value when
        present; formulas without a cache return ``""`` (openpyxl ``data_only``
        style). Non-anchor cells inside a merge are always ``""``. When
        ``errors_as_nan`` is true, Excel error cells (``t=\"e\"``) become
        ``float(\"nan\")`` (pandas/openpyxl reader parity). Trailing empty rows
        are trimmed; each row is padded to the max width. Stop after
        ``max_rows`` output rows when set.
        """
        sheet_data = self._part.element.find(_SHEET_DATA)
        if sheet_data is None:
            return []

        merges = self._merge_index
        by_row: dict[int, dict[int, Any]] = {}
        for row_elm in sheet_data.findall(_ROW):
            row_num = int(row_elm.get("r", "0"))
            if not row_num:
                continue
            cols: dict[int, Any] = {}
            for i, c_elm in enumerate(row_elm.findall(_C)):
                addr = c_elm.get("r")
                if addr:
                    col_letters, _ = split_address(addr)
                    col_idx = col_to_index(col_letters)
                else:
                    col_idx = i
                    addr = None

                if addr and merges.anchor_for(addr) and not merges.is_anchor(addr):
                    cols[col_idx] = ""
                    continue

                cell = Cell(c_elm, self)
                if errors_as_nan and cell.data_type == "e":
                    cols[col_idx] = float("nan")
                    continue

                val = cell.value
                cols[col_idx] = "" if val is None else val
            by_row[row_num] = cols

        if not by_row:
            return []

        first = min(by_row)
        last = max(by_row)
        data: list[list] = []
        last_row_with_data = -1

        for excel_row in range(first, last + 1):
            cols = by_row.get(excel_row, {})
            if cols:
                max_col = max(cols)
                converted_row = [cols.get(c, "") for c in range(max_col + 1)]
                while converted_row and converted_row[-1] == "":
                    converted_row.pop()
            else:
                converted_row = []

            if converted_row:
                last_row_with_data = len(data)
            data.append(converted_row)
            if max_rows is not None and len(data) >= max_rows:
                break

        data = data[: last_row_with_data + 1]

        if data:
            max_width = max(len(row) for row in data)
            if max_width and min(len(row) for row in data) < max_width:
                empty: list = [""]
                data = [row + (max_width - len(row)) * empty for row in data]

        return data

    def _resolve_template_style_maps(
        self,
        *,
        template_rows: int | list[int] | None,
    ) -> list[dict[int, str]]:
        row_nums = normalize_template_rows(template_rows)
        return [self._style_map_for_row(n) for n in row_nums]

    def _write_bulk_rows(
        self,
        row_list: list,
        *,
        start_row: int,
        start_col_idx: int,
        template_rows: int | list[int] | None = None,
        row_styles: list[dict[str, Any]] | None = None,
        column_styles: list[dict[str, Any]] | None = None,
        string_columns: set[int] | None = None,
    ) -> int:
        validate_style_sources(
            template_rows=template_rows,
            row_styles=row_styles,
        )

        template_maps = self._resolve_template_style_maps(
            template_rows=template_rows,
        )
        styles = self._workbook.styles
        inline_cache: BulkStyleCache | None = None
        if row_styles is not None:
            inline_cache = BulkStyleCache(styles)

        sst = self.shared_strings
        sst_cache: dict[str, int] = {}
        string_columns = string_columns or set()
        count = 0
        max_width = 0

        for offset, row_values in enumerate(row_list):
            row_num = start_row + offset
            row_elm = self._bulk_ensure_row(row_num)
            if not isinstance(row_values, (list, tuple)):
                row_values = (row_values,)
            max_width = max(max_width, len(row_values))
            for col_offset, value in enumerate(row_values):
                col_idx = start_col_idx + col_offset
                address = join_address(index_to_col(col_idx), row_num)
                c_elm = self._bulk_get_or_create_cell(address, row_elm)

                col_spec = (
                    column_styles[col_offset]
                    if column_styles is not None and col_offset < len(column_styles)
                    else {}
                )
                row_spec = (
                    row_styles[offset % len(row_styles)]
                    if row_styles is not None and row_styles
                    else {}
                )

                style_index: str | None = None
                if row_styles is not None:
                    merged = merge_style_specs(col_spec, row_spec)
                    idx = inline_cache.index_for(merged) if inline_cache else None
                    if idx is not None:
                        style_index = str(idx)
                elif template_maps:
                    template_s = template_maps[offset % len(template_maps)].get(col_idx)
                    col_only = normalize_style_spec(col_spec)
                    if col_only:
                        base = int(template_s) if template_s is not None else 0
                        idx = styles.allocate_cell_style(base_xf=base, **col_only)
                        style_index = str(idx)
                    elif template_s is not None:
                        style_index = template_s
                elif col_spec:
                    if inline_cache is None:
                        inline_cache = BulkStyleCache(styles)
                    idx = inline_cache.index_for(col_spec)
                    if idx is not None:
                        style_index = str(idx)

                if style_index is not None:
                    c_elm.set("s", style_index)

                write_cell_value(
                    c_elm,
                    value,
                    sst,
                    sst_cache,
                    force_string=col_idx in string_columns,
                    worksheet=self,
                )
            count += 1

        if count:
            end_row = start_row + count - 1
            end_col = start_col_idx + max_width - 1
            self._update_dimension_to(
                join_address(index_to_col(start_col_idx), start_row),
                join_address(index_to_col(end_col), end_row),
            )
        return count

    def _write_rows_at(
        self,
        rows,
        *,
        start_row: int,
        start_col_idx: int = 0,
        template_rows: int | list[int] | None = None,
        row_styles: list[dict[str, Any]] | None = None,
        column_styles: list[dict[str, Any]] | None = None,
        string_columns: set[int] | None = None,
    ) -> int:
        """Write row sequences at ``start_row`` / ``start_col_idx`` (0-based column).

        Internal helper for the pandas ExcelWriter engine and other bulk callers
        that already have row tuples rather than a DataFrame.
        """
        return self._write_bulk_rows(
            list(rows),
            start_row=start_row,
            start_col_idx=start_col_idx,
            template_rows=template_rows,
            row_styles=row_styles,
            column_styles=column_styles,
            string_columns=string_columns,
        )

    def write_rows(
        self,
        rows,
        *,
        at_row: int | None = None,
        at_cell: str | None = None,
        template_rows: int | list[int] | None = None,
        row_styles: list[dict[str, Any]] | None = None,
        column_styles: list[dict[str, Any]] | None = None,
        string_columns: set[int] | None = None,
    ) -> int:
        """Write rows at fixed addresses (overwrites existing cells in range)."""
        if at_cell is not None:
            start_col, at_row = split_address(at_cell)
            start_col_idx = col_to_index(start_col)
        else:
            if at_row is None:
                raise ValueError("write_rows requires at_row or at_cell")
            start_col_idx = 0

        return self._write_rows_at(
            list(rows),
            start_row=at_row,
            start_col_idx=start_col_idx,
            template_rows=template_rows,
            row_styles=row_styles,
            column_styles=column_styles,
            string_columns=string_columns,
        )

    def insert_rows(
        self,
        rows,
        *,
        at_row: int | None = None,
        at_cell: str | None = None,
        template_rows: int | list[int] | None = None,
        row_styles: list[dict[str, Any]] | None = None,
        column_styles: list[dict[str, Any]] | None = None,
        string_columns: set[int] | None = None,
    ) -> int:
        """Insert rows at ``at_row``, shifting existing content down."""
        if at_cell is not None:
            start_col, at_row = split_address(at_cell)
            start_col_idx = col_to_index(start_col)
        else:
            if at_row is None:
                raise ValueError("insert_rows requires at_row or at_cell")
            start_col_idx = 0

        row_list = list(rows)
        if not row_list:
            return 0

        count = len(row_list)
        template_merge_bands = self._snapshot_template_row_merges(template_rows)
        sheet_data = self._sheet_data()
        shift_sheet_data_rows(sheet_data, at_row, count)
        shift_sheet_row_references(self._part.element, at_row, count)
        shift_table_parts(self, at_row, count)
        shift_defined_names(self._workbook, self.name, at_row, count)
        self._invalidate_bulk_indexes()

        written = self._write_bulk_rows(
            row_list,
            start_row=at_row,
            start_col_idx=start_col_idx,
            template_rows=template_rows,
            row_styles=row_styles,
            column_styles=column_styles,
            string_columns=string_columns,
        )
        self._apply_template_row_merges(at_row, written, template_merge_bands)
        self.update_dimension()
        return written

    def insert_columns(
        self,
        cols,
        *,
        at_col: str | int | None = None,
        at_cell: str | None = None,
        template_rows: int | list[int] | None = None,
        row_styles: list[dict[str, Any]] | None = None,
        column_styles: list[dict[str, Any]] | None = None,
        string_columns: set[int] | None = None,
    ) -> int:
        """Insert columns at ``at_col``, shifting existing content right.

        ``cols`` is a sequence of column vectors (each top→bottom). Does not
        rewrite formula text or move drawing/chart anchors.
        """
        if at_cell is not None:
            start_col, start_row = split_address(at_cell)
            at_col_idx = col_to_index(start_col)
        else:
            if at_col is None:
                raise ValueError("insert_columns requires at_col or at_cell")
            start_row = 1
            if isinstance(at_col, int):
                at_col_idx = at_col
            else:
                at_col_idx = col_to_index(str(at_col))

        col_list = list(cols)
        if not col_list:
            return 0

        count = len(col_list)
        sheet_data = self._sheet_data()
        shift_sheet_data_cols(sheet_data, at_col_idx, count)
        shift_sheet_col_references(self._part.element, at_col_idx, count)
        shift_table_parts_cols(self, at_col_idx, count)
        shift_col_dimensions(self._part.element, at_col_idx, count)
        shift_defined_names_cols(self._workbook, self.name, at_col_idx, count)
        self._invalidate_bulk_indexes()

        # Transpose column vectors → row tuples for _write_bulk_rows
        max_len = max((len(v) if isinstance(v, (list, tuple)) else 1) for v in col_list)
        row_list: list[tuple] = []
        for row_offset in range(max_len):
            row_vals = []
            for col_vec in col_list:
                if not isinstance(col_vec, (list, tuple)):
                    col_vec = (col_vec,)
                row_vals.append(col_vec[row_offset] if row_offset < len(col_vec) else None)
            row_list.append(tuple(row_vals))

        written = self._write_bulk_rows(
            row_list,
            start_row=start_row,
            start_col_idx=at_col_idx,
            template_rows=template_rows,
            row_styles=row_styles,
            column_styles=column_styles,
            string_columns=string_columns,
        )
        self.update_dimension()
        return count if written else 0

    def _detach_shared_formula_followers(
        self, si: str, *, exclude: etree._Element | None = None
    ) -> int:
        """Remove ``<f>`` from shared-formula followers with the given ``si``.

        Used when the shared-formula master is overwritten so Excel does not see
        orphan ``t=\"shared\"`` followers. Returns the number of followers detached.
        """
        removed = 0
        for c_elm in self._part.element.iter(_C):
            if exclude is not None and c_elm is exclude:
                continue
            f_elm = c_elm.find(_F)
            if f_elm is None:
                continue
            if f_elm.get("t") != "shared" or f_elm.get("si") != si:
                continue
            c_elm.remove(f_elm)
            removed += 1
        return removed

    def clear_range(self, cell_range: str) -> int:
        """Clear cell values in ``cell_range`` while preserving styles. Returns cells cleared."""
        c1, r1, c2, r2 = parse_range(cell_range)
        c1i, c2i = col_to_index(c1), col_to_index(c2)
        cleared = 0
        for row_num in range(r1, r2 + 1):
            for col_idx in range(c1i, c2i + 1):
                address = join_address(index_to_col(col_idx), row_num)
                c_elm = self._find_cell_element(address)
                if c_elm is None:
                    continue
                if c_elm.find(_F) is not None:
                    self._workbook._on_formula_removed(self, c_elm)
                _clear_cell_content(c_elm)
                c_elm.attrib.pop("t", None)
                cleared += 1
        if cleared:
            self._invalidate_bulk_indexes()
        return cleared

    def update_dimension(self) -> None:
        """Recompute ``<dimension ref>`` from populated cells."""
        sheet_data = self._part.element.find(_SHEET_DATA)
        if sheet_data is None:
            return
        min_row = min_col = None
        max_row = max_col = None
        for row in sheet_data.findall(_ROW):
            row_num = int(row.get("r", "0"))
            for c_elm in row.findall(_C):
                addr = c_elm.get("r")
                if not addr:
                    continue
                col, r = split_address(addr)
                ci = col_to_index(col)
                min_row = r if min_row is None else min(min_row, r)
                max_row = r if max_row is None else max(max_row, r)
                min_col = ci if min_col is None else min(min_col, ci)
                max_col = ci if max_col is None else max(max_col, ci)
        if min_row is None:
            return
        self._update_dimension_to(
            join_address(index_to_col(min_col), min_row),
            join_address(index_to_col(max_col), max_row),
        )

    def _dimension_insert_index(self) -> int:
        """Return index for ``<dimension>`` (after ``sheetPr``, before ``sheetViews``)."""
        root = self._part.element
        insert_at = 0
        for i, child in enumerate(root):
            local = etree.QName(child).localname
            if local == "sheetPr":
                insert_at = i + 1
                continue
            if local in ("sheetViews", "sheetFormatPr", "cols", "sheetData"):
                return i
        return insert_at

    def _update_dimension_to(self, top_left: str, bottom_right: str) -> None:
        ref = top_left if top_left == bottom_right else f"{top_left}:{bottom_right}"
        root = self._part.element
        dim = root.find(_DIMENSION)
        if dim is None:
            dim = etree.Element(_DIMENSION)
        else:
            root.remove(dim)
        dim.set("ref", ref)
        root.insert(self._dimension_insert_index(), dim)

    def expand_conditional_formatting(self, end_row: int) -> int:
        """Extend single-column CF ranges down to ``end_row``. Returns blocks updated."""
        return expand_conditional_formatting_to_row(self, end_row)

    @property
    def cells(self) -> list[Cell]:
        return list(self.iter_cells())

    @property
    def images(self) -> list[Picture]:
        pics: list[Picture] = []
        for dp in drawing_parts_for_worksheet(self):
            pics.extend(iter_pictures(dp, self))
        return pics

    @property
    def charts(self) -> list[Chart]:
        charts: list[Chart] = []
        for dp in drawing_parts_for_worksheet(self):
            charts.extend(iter_charts(dp))
        return charts

    @property
    def tables(self) -> list[Table]:
        return table_parts_for_worksheet(self)

    @property
    def conditional_formatting(self) -> list[ConditionalFormatting]:
        return iter_conditional_formatting(self)

    def _snapshot_template_row_merges(
        self, template_rows: int | list[int] | None
    ) -> list[list[tuple[str, str]]]:
        """Horizontal merge (c1, c2) pairs per template row, before insert shift."""
        row_nums = normalize_template_rows(template_rows)
        if not row_nums:
            return []
        by_row: dict[int, list[tuple[str, str]]] = {n: [] for n in row_nums}
        for ref in self.merged_ranges:
            c1, r1, c2, r2 = parse_range(ref)
            if r1 == r2 and r1 in by_row and c1 != c2:
                by_row[r1].append((c1, c2))
        return [by_row[n] for n in row_nums]

    def _apply_template_row_merges(
        self,
        at_row: int,
        written: int,
        bands: list[list[tuple[str, str]]],
    ) -> None:
        if not bands or written <= 0:
            return
        n_templates = len(bands)
        for offset in range(written):
            dest = at_row + offset
            for c1, c2 in bands[offset % n_templates]:
                self.merge_cells(f"{join_address(c1, dest)}:{join_address(c2, dest)}")

    def merge_cells(self, cell_range: str) -> None:
        """Merge cells in ``cell_range`` (e.g. ``A1:C1``)."""
        ref = normalize_range(cell_range)
        for existing in self._merge_index.refs:
            if ranges_overlap(existing, ref):
                raise ValueError(f"merge range {ref!r} overlaps existing {existing!r}")
        block = self._part.element.find(_MERGE_CELLS)
        if block is None:
            from xlsxedit.worksheet_order import insert_worksheet_child

            block = etree.Element(_MERGE_CELLS)
            insert_worksheet_child(self._part.element, block)
        merge_elm = etree.SubElement(block, _MERGE_CELL)
        merge_elm.set("ref", ref)
        block.set("count", str(len(block.findall(_MERGE_CELL))))
        self._invalidate_merge_map()
        self._invalidate_bulk_indexes()

    def unmerge_cells(self, cell_range: str) -> None:
        """Remove a merge range."""
        ref = normalize_range(cell_range)
        block = self._part.element.find(_MERGE_CELLS)
        if block is None:
            return
        removed = False
        for merge_elm in list(block.findall(_MERGE_CELL)):
            if merge_elm.get("ref") == ref:
                block.remove(merge_elm)
                removed = True
        if removed:
            count = len(block.findall(_MERGE_CELL))
            if count == 0:
                self._part.element.remove(block)
            else:
                block.set("count", str(count))
            self._invalidate_merge_map()
            self._invalidate_bulk_indexes()

    def add_conditional_formatting(
        self,
        cell_range: str,
        *,
        rule_type: str = "cellIs",
        operator: str = "greaterThan",
        formula: str = "0",
        font_color: str | None = None,
        bg_color: str | None = None,
        bold: bool = False,
    ) -> ConditionalFormatting:
        """Add a ``cellIs`` rule on ``cell_range``.

        Each cell in the range is tested against ``formula`` using ``operator``
        (e.g. ``greaterThan`` with formula ``"0"`` → value > 0). Matching cells
        receive the rule's differential format (fill/font).

        When ``bg_color`` / ``font_color`` are omitted, defaults to light-green
        fill (``C6EFCE``) and dark-green font (``006100``).
        """
        if rule_type != "cellIs":
            raise ValueError(f"unsupported rule_type: {rule_type!r}")
        return add_cell_is_rule(
            self,
            cell_range,
            operator=operator,
            formula=formula,
            font_color=font_color,
            bg_color=bg_color,
            bold=bold,
        )

    def add_color_scale_formatting(self, cell_range: str) -> ConditionalFormatting:
        """Add a ``colorScale`` conditional formatting rule to ``cell_range``."""
        return clone_color_scale_rule(self, cell_range)

    def add_image(
        self,
        image_path: str | Path,
        *,
        anchor: str,
        width: int | None = None,
        height: int | None = None,
        max_width: int = 200,
        max_height: int = 200,
        name: str | None = None,
        offset_x: int = 0,
        offset_y: int = 0,
    ) -> Picture:
        """Insert an image anchored at ``anchor`` (clones template drawing subgraph).

        ``offset_x`` / ``offset_y`` are pixel insets within the anchor cell (default 0 = flush).
        """
        return self._workbook._add_image_to_sheet(
            self,
            image_path,
            anchor,
            width,
            height,
            max_width,
            max_height,
            name,
            offset_x=offset_x,
            offset_y=offset_y,
        )

    def add_chart(
        self,
        chart_type: str = "bar",
        *,
        anchor: str = "E2",
        data_range: str = "A1:B5",
        title: str | None = None,
        name: str | None = None,
        to_anchor: str | None = None,
    ) -> Chart:
        """Add a chart from the bundled template subgraph.

        Placement is cell-based (``anchor``, optional ``to_anchor``). Offsets come from
        the template (flush). Tweak later via ``chart.offset_x`` / ``chart.offset_y`` if needed.
        """
        return self._workbook._add_chart_to_sheet(
            self,
            chart_type,
            anchor=anchor,
            data_range=data_range,
            title=title,
            name=name,
            to_anchor=to_anchor,
        )

    def add_table(
        self,
        cell_range: str,
        columns: list[str],
        *,
        name: str | None = None,
        display_name: str | None = None,
    ) -> Table:
        """Add an Excel table over ``cell_range``."""
        return self._workbook._add_table_to_sheet(
            self, cell_range, columns, name=name, display_name=display_name
        )

    def replace(self, old: str, new, *, value_type: str | None = None) -> int:
        """Replace placeholder text in string cells on this sheet.

        When at least one cell is changed, sets ``calcPr/@fullCalcOnLoad`` on
        the workbook so Excel recalculates formulas on open.
        """
        count = self._workbook._replace_on_sheet(self, old, new, value_type=value_type)
        if count > 0:
            self._workbook.set_full_calc_on_load()
        return count

    def find(self, value) -> Cell | None:
        """Return the first non-formula cell whose ``value`` equals ``value``, or ``None``."""
        from xlsxedit.cell import cell_matches

        for cell in self.iter_cells():
            if cell_matches(cell, value):
                return cell
        return None

    def findall(self, value) -> list[Cell]:
        """Return all non-formula cells whose ``value`` equals ``value``."""
        from xlsxedit.cell import cell_matches

        return [cell for cell in self.iter_cells() if cell_matches(cell, value)]
