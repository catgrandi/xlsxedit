"""Public Workbook API."""

from __future__ import annotations

import mimetypes
import re
import uuid
from copy import deepcopy
from datetime import date, datetime
from pathlib import Path
from typing import BinaryIO

from lxml import etree

from xlsxedit.api import _default_xlsx_path
from xlsxedit.exceptions import (
    DuplicateWorksheetError,
    InvalidRangeError,
    TableError,
    WorksheetNotFoundError,
)
from xlsxedit.drawing import (
    EMU_PER_PIXEL,
    Picture,
    Chart,
    Table,
    drawing_parts_for_worksheet,
    iter_charts,
    iter_pictures,
    _px_to_emu,
    _set_from_offset_emu,
    _validate_offset_px,
)
from xlsxedit.image_size import resolve_display_size
from xlsxedit.merge import merge_anchor, parse_range
from xlsxedit.opc.constants import CT, OFFICE_REL_NS, RT, SML_NS
from xlsxedit.opc.package import OpcPackage
from xlsxedit.opc.part import Part
from xlsxedit.opc.packuri import PackURI
from xlsxedit.oxml.address import col_to_index, index_to_col, join_address, split_address
from xlsxedit.oxml.parser import parse_template_xml, parse_xml, serialize_xml
from xlsxedit.parts import WorkbookPart, WorksheetPart, register_part_types
from xlsxedit.shared_strings import SharedStringTable
from xlsxedit.styles import Styles, datetime_to_serial
from xlsxedit.cell import Cell
from xlsxedit.worksheet import Worksheet

register_part_types()

_XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_XDR_TWO_CELL = f"{{{_XDR_NS}}}twoCellAnchor"
_XDR_ONE_CELL = f"{{{_XDR_NS}}}oneCellAnchor"
_XDR_GRAPHIC_FRAME = f"{{{_XDR_NS}}}graphicFrame"
_C_CHART = f"{{http://schemas.openxmlformats.org/drawingml/2006/chart}}chart"
_TABLE_PARTS = f"{{{SML_NS}}}tableParts"
_TABLE_PART = f"{{{SML_NS}}}tablePart"
_XDR_FROM = f"{{{_XDR_NS}}}from"
_XDR_TO = f"{{{_XDR_NS}}}to"
_XDR_COL = f"{{{_XDR_NS}}}col"
_XDR_ROW = f"{{{_XDR_NS}}}row"
_XDR_EXT = f"{{{_XDR_NS}}}ext"
_XDR_CNVPR = f"{{{_XDR_NS}}}cNvPr"
_A_EXT = f"{{{_A_NS}}}ext"
_WS_DRAWING = f"{{{SML_NS}}}drawing"


def _coerce_date(value) -> datetime | date:
    if isinstance(value, (datetime, date)):
        return value
    if isinstance(value, str):
        return datetime.fromisoformat(value)
    raise TypeError(f"date value must be datetime, date, or ISO string; got {type(value)!r}")


def _sheet_prefix(sheet_name: str) -> str:
    if not sheet_name.replace("_", "").isalnum():
        safe = sheet_name.replace("'", "''")
        return f"'{safe}'!"
    return f"{sheet_name}!"


def _sheet_abs_ref(sheet_name: str, col: str, row: int) -> str:
    return f"{_sheet_prefix(sheet_name)}${col.upper()}${row}"


def _sheet_abs_range(sheet_name: str, ref: str) -> str:
    c1, r1, c2, r2 = parse_range(ref)
    prefix = _sheet_prefix(sheet_name)
    if c1 == c2 and r1 == r2:
        return f"{prefix}${c1.upper()}${r1}"
    return f"{prefix}${c1.upper()}${r1}:${c2.upper()}${r2}"


class Workbook:
    """An open workbook.

    ``Workbook()`` or :meth:`create` — blank workbook from the bundled template.
    ``Workbook(path)`` or :meth:`open` — open an existing ``.xlsx``, OPC folder,
    or binary file-like (e.g. ``BytesIO``).
    """

    def __init__(
        self,
        path_or_package: str | Path | BinaryIO | OpcPackage | None = None,
        *,
        large: bool = False,
    ):
        self._large = large
        if path_or_package is None:
            package = OpcPackage.open(_default_xlsx_path())
        elif isinstance(path_or_package, OpcPackage):
            package = path_or_package
            self._large = getattr(package, "_large", False)
        else:
            package = OpcPackage.open(path_or_package, large=large)
        self._package = package
        self._workbook_part: WorkbookPart = package.main_document_part()
        self._refresh_sheets()
        sst_part = self._workbook_part.shared_strings_part
        self._shared_strings = (
            SharedStringTable(sst_part.element, defer_index=self._large)
            if sst_part is not None
            else None
        )
        styles_part = self._workbook_part.styles_part
        if styles_part is not None:
            # Live element: Styles mutations serialize automatically on save.
            self._styles = Styles(styles_part.element, workbook=self)
        else:
            self._styles = Styles(None, workbook=self)

    def __repr__(self) -> str:
        return f"<Workbook sheets={self.sheetnames!r}>"

    def _refresh_sheets(self) -> None:
        self._sheets = [
            Worksheet(name, part, self)
            for name, part in self._workbook_part.worksheet_parts()
        ]
        self._sheets_by_name = {ws.name: ws for ws in self._sheets}

    @classmethod
    def open(cls, path: str | Path | BinaryIO, *, large: bool = False) -> Workbook:
        """Open an ``.xlsx`` file, unpacked OPC directory, or binary file-like.

        ``path`` may be a filesystem path, an unpacked package folder, or a
        binary stream of a complete ``.xlsx`` zip (e.g. ``io.BytesIO``). Streams
        are not unpacked folders.

        Pass ``large=True`` to defer parsing of unaccessed worksheets and build the
        shared-string index on first use (lighter open for huge workbooks).
        """
        return cls(path, large=large)

    @classmethod
    def create(cls) -> Workbook:
        """Create a new blank workbook from the bundled default template."""
        return cls()

    def save(self, path: str | Path | BinaryIO, *, include_orphans: bool = False) -> None:
        """Save the workbook to a path or a binary file-like (e.g. ``BytesIO``).

        After ``save(buf)``, use ``buf.getvalue()`` or ``buf.seek(0)`` before reading.
        """
        self._package.save(path, include_orphans=include_orphans)

    def _invalidate_calc_chain(self) -> None:
        """Remove ``xl/calcChain.xml`` and its workbook relationship (idempotent).

        Excel rebuilds the chain on open. Call when a formula ``<f>`` is removed
        so the package does not list stale formula cells.
        """
        try:
            part = self._workbook_part.rels.part_with_reltype(RT.CALC_CHAIN)
        except KeyError:
            return
        for r_id, rel in list(self._workbook_part.rels._rels.items()):
            if rel.reltype == RT.CALC_CHAIN:
                self._workbook_part.rels._rels.pop(r_id, None)
        self._package._remove_part(part)

    _CALC_PR = f"{{{SML_NS}}}calcPr"

    def _workbook_child_insert_index(self, wb_elm: etree._Element, localname: str) -> int:
        from xlsxedit._ooxml_order import ORDER
        from xlsxedit.worksheet_order import ordered_insert_index

        return ordered_insert_index(wb_elm, localname, ORDER["CT_Workbook"])

    def set_full_calc_on_load(self) -> None:
        """Set ``calcPr/@fullCalcOnLoad`` so Excel recalculates on open.

        Idempotent. Does not evaluate formulas or strip cached ``<v>`` values;
        Excel (and hosts that honor this flag) refresh formula results after open.
        If ``calcMode="manual"``, some hosts may ignore ``fullCalcOnLoad``.
        """
        wb_elm = self._workbook_part.element
        calc_pr = wb_elm.find(self._CALC_PR)
        if calc_pr is None:
            calc_pr = etree.Element(self._CALC_PR)
            wb_elm.insert(self._workbook_child_insert_index(wb_elm, "calcPr"), calc_pr)
        calc_pr.set("fullCalcOnLoad", "1")

    def _on_formula_removed(self, worksheet: Worksheet, c_elm) -> None:
        """Bookkeeping before a cell loses its ``<f>`` element.

        Raises ``FormulaGroupError``, changing nothing, if the cell anchors an
        array or data-table formula over several cells: removing the anchor
        would leave the rest of the range as orphaned values. Otherwise drops
        stale ``calcChain.xml``. If the cell was a shared-formula master, also
        strips ``<f>`` from same-``si`` followers on that sheet.
        """
        from xlsxedit.cell import _F, _refuse_range_formula_edit

        f_elm = c_elm.find(_F)
        if f_elm is not None:
            _refuse_range_formula_edit(c_elm.get("r", ""), f_elm)
        if f_elm is not None and f_elm.get("t") == "shared" and (f_elm.text or "").strip():
            si = f_elm.get("si")
            if si is not None:
                worksheet._detach_shared_formula_followers(si, exclude=c_elm)
        self._invalidate_calc_chain()

    @property
    def orphan_partnames(self) -> tuple[str, ...]:
        return self._package.orphan_partnames

    @property
    def properties(self):
        """Core document properties (``/docProps/core.xml``) — title, author, etc."""
        from xlsxedit.parts import CorePropertiesPart

        try:
            part = self._package.rels.part_with_reltype(RT.CORE_PROPERTIES)
        except KeyError:
            part = CorePropertiesPart.default(self._package)
            self._package._add_part(part)
            r_id = self._package.rels.next_rId()
            self._package.rels.add_relationship(RT.CORE_PROPERTIES, part, r_id)
        return part.core_properties

    @property
    def sheetnames(self) -> list[str]:
        return [ws.name for ws in self._sheets]

    @property
    def worksheets(self) -> list[Worksheet]:
        return list(self._sheets)

    @property
    def shared_strings(self) -> SharedStringTable | None:
        return self._shared_strings

    @property
    def styles(self) -> Styles:
        return self._styles

    def __getitem__(self, name: str) -> Worksheet:
        try:
            return self._sheets_by_name[name]
        except KeyError:
            raise WorksheetNotFoundError(f"worksheet {name!r} not found") from None

    def __iter__(self):
        return iter(self._sheets)

    def write_dataframe(
        self,
        df,
        *,
        sheet: str | None = None,
        at_cell: str = "A1",
        header: bool = True,
        mode: str = "overwrite",
        template_rows: int | list[int] | None = None,
        row_styles: list[dict] | None = None,
        column_styles: list[dict] | None = None,
        resize_table: bool = True,
        table: str | Table | None = None,
        clear_range: str | None = None,
        expand_conditional_formatting: bool = False,
        string_columns: set[int] | None = None,
    ) -> int:
        """Write a pandas-like DataFrame to a worksheet.

        ``mode`` is ``"overwrite"`` (default) or ``"insert"`` (shifts rows below down).

        In overwrite mode, ``resize_table`` resizes the table that the written
        rows overlap so that it ends at the last written row; ``table`` (a
        table name or :class:`Table` on the sheet) picks the table instead,
        for example to grow it by writing below it. Before anything is
        written, :class:`~xlsxedit.exceptions.TableError` is raised when the
        rows overlap several tables, do not span exactly the table's columns,
        or would put anything but its column names in its header row, or when
        the resized table would overlap another table.
        """
        if mode not in ("overwrite", "insert"):
            raise ValueError(f"mode must be 'overwrite' or 'insert', got {mode!r}")
        if table is not None and (mode != "overwrite" or not resize_table):
            raise ValueError(
                "table= picks the table to resize, so it needs mode='overwrite' "
                "and resize_table=True"
            )

        if sheet is None:
            ws = self._sheets[0]
        else:
            ws = self[sheet]

        start_col, start_row = split_address(at_cell)
        start_col_idx = col_to_index(start_col)

        rows_to_write = []
        if header and hasattr(df, "columns"):
            rows_to_write.append(list(df.columns))

        if hasattr(df, "itertuples"):
            data_rows = [tuple(row) for row in df.itertuples(index=False)]
        elif hasattr(df, "values"):
            data_rows = [tuple(row) for row in df.values]
        else:
            data_rows = list(df)

        rows_to_write.extend(data_rows)
        data_start_row = start_row + (1 if header and hasattr(df, "columns") else 0)

        # Plan the table resize before anything is written, so a refusal
        # leaves the sheet as it was.
        target = _find_table(ws, table) if table is not None else None
        resize = None
        if mode == "overwrite" and resize_table and data_rows:
            resize = _plan_table_resize(
                ws,
                target,
                rows_to_write,
                start_row=start_row,
                start_col=start_col_idx,
                header_row=header and hasattr(df, "columns"),
            )
        elif mode == "overwrite" and resize_table and rows_to_write:
            _check_header_row_write(
                ws, rows_to_write[0], start_row=start_row, start_col=start_col_idx
            )

        if mode == "overwrite" and clear_range is not None:
            ws.clear_range(clear_range)

        effective_template = template_rows
        if effective_template is None and header and mode == "overwrite":
            effective_template = start_row

        style_kwargs = {
            "template_rows": effective_template,
            "row_styles": row_styles,
            "column_styles": column_styles,
            "string_columns": string_columns,
        }

        if mode == "insert":
            count = ws.insert_rows(
                rows_to_write,
                at_cell=join_address(start_col, start_row),
                **style_kwargs,
            )
        else:
            count = ws.write_rows(
                rows_to_write,
                at_cell=join_address(start_col, start_row),
                **style_kwargs,
            )

        data_rows_written = count - (1 if header and hasattr(df, "columns") else 0)
        if data_rows_written > 0 and mode == "overwrite":
            end_row = data_start_row + data_rows_written - 1
            if resize is not None:
                resized, new_ref = resize
                resized.resize(new_ref)
            if expand_conditional_formatting:
                ws.expand_conditional_formatting(end_row)
        elif data_rows_written > 0 and mode == "insert" and expand_conditional_formatting:
            end_row = start_row + count - 1
            ws.expand_conditional_formatting(end_row)

        return data_rows_written

    @classmethod
    def export_to_template(cls, template_path, df, save_path=None, **kwargs) -> Workbook:
        """Open ``template_path``, write ``df``, optionally save to ``save_path``."""
        wb = cls.open(template_path)
        wb.write_dataframe(df, **kwargs)
        if save_path is not None:
            wb.save(save_path)
        return wb

    def _sst_index_ref_count(self, index: int) -> int:
        """Count cells across the workbook referencing a shared-string index."""
        count = 0
        for ws in self._sheets:
            for cell in ws.iter_cells():
                if cell.data_type == "s" and cell._v_text() == str(index):
                    count += 1
        return count

    def replace(self, old: str, new, *, value_type: str | None = None) -> int:
        """Replace ``old`` across all sheets. See :meth:`Worksheet.replace`.

        When at least one cell is changed, sets ``calcPr/@fullCalcOnLoad`` so
        Excel recalculates formulas on open.
        """
        if not old:
            return 0
        total = 0
        for ws in self._sheets:
            total += self._replace_on_sheet(ws, old, new, value_type=value_type)
        if total > 0:
            self.set_full_calc_on_load()
        return total

    def find(self, value, *, sheet: str | None = None) -> Cell | None:
        """Return the first cell with ``value``, or ``None``.

        Search all sheets in workbook order, or only ``sheet`` when given.
        """
        for ws in self._sheets_for_find(sheet):
            found = ws.find(value)
            if found is not None:
                return found
        return None

    def findall(self, value, *, sheet: str | None = None) -> list[Cell]:
        """Return all cells with ``value`` (sheet order, then row-major)."""
        result: list[Cell] = []
        for ws in self._sheets_for_find(sheet):
            result.extend(ws.findall(value))
        return result

    def _sheets_for_find(self, sheet: str | None):
        if sheet is None:
            return self._sheets
        return [self[sheet]]

    def replace_image(self, name: str, image_path: str | Path) -> int:
        """Replace image(s) whose picture name equals ``name``."""
        total = 0
        for ws in self._sheets:
            for pic in ws.images:
                if pic.name == name:
                    pic.replace(image_path)
                    total += 1
        return total

    def insert_image_at_placeholder(
        self,
        placeholder: str,
        image_path: str | Path,
        *,
        width: int | None = None,
        height: int | None = None,
        max_width: int = 200,
        max_height: int = 200,
    ) -> int:
        """Replace whole-cell text placeholder with an image at that cell.

        Size kwargs are forwarded to :meth:`Worksheet.add_image` (pixels).
        With only ``width`` or only ``height``, aspect ratio is preserved.
        """
        total = 0
        for ws in self._sheets:
            for cell in ws.iter_cells():
                if cell.has_formula or not cell.is_string:
                    continue
                if cell.value != placeholder:
                    continue
                addr = cell.address
                cell.value = None
                ws.add_image(
                    image_path,
                    anchor=addr,
                    name=placeholder.strip("{}"),
                    width=width,
                    height=height,
                    max_width=max_width,
                    max_height=max_height,
                )
                total += 1
        return total

    def _replace_on_sheet(
        self, ws: Worksheet, old: str, new, *, value_type: str | None = None
    ) -> int:
        if value_type in (None, "text"):
            return self._substring_replace_on_sheet(ws, old, str(new))

        if value_type == "number":
            total = 0
            for cell in ws.iter_cells():
                if cell.has_formula:
                    continue
                if cell.value == old:
                    cell.value = new
                    total += 1
            return total

        if value_type == "date":
            dt = _coerce_date(new)
            serial = datetime_to_serial(dt)
            total = 0
            for cell in ws.iter_cells():
                if cell.has_formula:
                    continue
                if cell.value == old:
                    cell.value = serial
                    if not cell.style.is_date:
                        cell.apply_date_format()
                    total += 1
                elif isinstance(cell.value, (int, float)) and str(cell.value) == old:
                    cell.value = serial
                    if not cell.style.is_date:
                        cell.apply_date_format()
                    total += 1
            return total

        raise ValueError(f"unsupported value_type: {value_type!r}")

    def _substring_replace_on_sheet(self, ws: Worksheet, old: str, new: str) -> int:
        total = 0
        seen_sst_indexes: set[int] = set()
        for cell in ws.iter_cells():
            if cell.has_formula:
                continue
            if cell.data_type == "s":
                idx_text = cell._v_text()
                if idx_text is None:
                    continue
                idx = int(idx_text)
                if idx in seen_sst_indexes:
                    continue
                seen_sst_indexes.add(idx)
                if self._shared_strings is None:
                    continue
                total += self._shared_strings.replace_in_index(idx, old, new)
            elif cell.data_type == "inlineStr":
                total += cell.replace(old, new)
        return total

    def rename_worksheet(self, old_name: str, new_name: str) -> Worksheet:
        if new_name in self._sheets_by_name:
            raise DuplicateWorksheetError(f"worksheet {new_name!r} already exists")
        if old_name not in self._sheets_by_name:
            raise WorksheetNotFoundError(f"worksheet {old_name!r} not found")
        for sheet_elm in self._workbook_part.sheet_elements:
            if sheet_elm.get("name") == old_name:
                sheet_elm.set("name", new_name)
                break
        self._refresh_sheets()
        return self._sheets_by_name[new_name]

    def remove_worksheet(self, name: str) -> None:
        if len(self._sheets) <= 1:
            raise ValueError("cannot remove the only worksheet")
        if name not in self._sheets_by_name:
            raise WorksheetNotFoundError(f"worksheet {name!r} not found")
        removed = self._workbook_part.remove_sheet_element(name)
        if removed is None:
            raise WorksheetNotFoundError(f"worksheet {name!r} not found")
        r_id, _sheet_id = removed
        rel = self._workbook_part.rels.get(r_id)
        if rel is not None and not rel.is_external:
            part = rel.target_part
            self._workbook_part.rels._rels.pop(r_id, None)
            self._package._remove_part(part)
        self._refresh_sheets()

    def add_worksheet(self, name: str) -> Worksheet:
        """Add a new empty worksheet from the bundled worksheet template."""
        if name in self._sheets_by_name:
            raise DuplicateWorksheetError(f"worksheet {name!r} already exists")

        template_path = _default_xlsx_path().parent / "default-worksheet.xml"
        element = parse_template_xml(template_path.read_bytes())
        partname = self._package.next_partname("/xl/worksheets/sheet%d.xml")
        part = WorksheetPart(partname, CT.WORKSHEET, element, self._package)
        self._package._add_part(part)

        r_id = self._workbook_part.relate_to(part, RT.WORKSHEET)

        sheet_id = self._workbook_part.next_sheet_id()
        self._workbook_part.append_sheet_element(name, sheet_id, r_id)
        self._refresh_sheets()
        return self._sheets_by_name[name]

    def copy_worksheet(self, name: str, new_name: str) -> Worksheet:
        """Duplicate an existing worksheet (cells, styles, merges, sheet-owned parts)."""
        from xlsxedit.sheet_clone import clone_sheet_relationships, copy_local_defined_names

        if new_name in self._sheets_by_name:
            raise DuplicateWorksheetError(f"worksheet {new_name!r} already exists")
        if name not in self._sheets_by_name:
            raise WorksheetNotFoundError(f"worksheet {name!r} not found")

        source = self._sheets_by_name[name]
        source_index = list(self.sheetnames).index(name)
        element = deepcopy(source._part.element)
        partname = self._package.next_partname("/xl/worksheets/sheet%d.xml")
        part = WorksheetPart(partname, source._part.content_type, element, self._package)
        self._package._add_part(part)
        clone_sheet_relationships(source._part, part, self)

        r_id = self._workbook_part.relate_to(part, RT.WORKSHEET)
        sheet_id = self._workbook_part.next_sheet_id()
        self._workbook_part.append_sheet_element(new_name, sheet_id, r_id)
        copy_local_defined_names(self, source_index, len(self._sheets))
        self._refresh_sheets()
        return self._sheets_by_name[new_name]

    def _add_image_to_sheet(
        self,
        ws: Worksheet,
        image_path: str | Path,
        anchor: str,
        width: int | None,
        height: int | None,
        max_width: int,
        max_height: int,
        name: str | None,
        offset_x: int = 0,
        offset_y: int = 0,
    ) -> Picture:
        offset_x = _validate_offset_px(offset_x)
        offset_y = _validate_offset_px(offset_y)
        path = Path(image_path)
        display_w, display_h = resolve_display_size(
            path,
            width=width,
            height=height,
            max_width=max_width,
            max_height=max_height,
        )
        data = path.read_bytes()
        ext = path.suffix.lower().lstrip(".") or "jpeg"
        ct = mimetypes.types_map.get(f".{ext}", CT.JPEG)

        media_partname = self._package.next_partname("/xl/media/image%d." + ext)
        media_part = Part(media_partname, ct, data, self._package)
        self._package._add_part(media_part)

        template_path = _default_xlsx_path().parent / "default-picture-anchor.xml"
        anchor_template = parse_template_xml(template_path.read_bytes()).find(_XDR_ONE_CELL)

        drawing_parts = drawing_parts_for_worksheet(ws)
        if drawing_parts:
            drawing_part = drawing_parts[0]
            root = parse_xml(drawing_part.blob)
            anchor_elm = deepcopy(anchor_template)
        else:
            root = parse_xml(
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                b'<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"'
                b' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
                b' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
            )
            anchor_elm = deepcopy(anchor_template)
            drawing_partname = self._package.next_partname("/xl/drawings/drawing%d.xml")
            drawing_part = Part(drawing_partname, CT.DRAWING, b"", self._package)
            self._package._add_part(drawing_part)
            r_id = ws._part.relate_to(drawing_part, RT.DRAWING)
            from xlsxedit.worksheet_order import insert_worksheet_child

            drawing_elm = etree.Element(_WS_DRAWING)
            drawing_elm.set(f"{{{OFFICE_REL_NS}}}id", r_id)
            insert_worksheet_child(ws._part.element, drawing_elm)

        if anchor_elm is None:
            raise ValueError("picture anchor template missing")

        col, row = split_address(anchor)
        from_elm = anchor_elm.find(_XDR_FROM)
        if from_elm is not None:
            from_elm.find(_XDR_COL).text = str(col_to_index(col))
            from_elm.find(_XDR_ROW).text = str(row - 1)
        _set_from_offset_emu(anchor_elm, _px_to_emu(offset_x), _px_to_emu(offset_y))

        cx, cy = int(display_w * EMU_PER_PIXEL), int(display_h * EMU_PER_PIXEL)
        ext_elm = anchor_elm.find(_XDR_EXT)
        if ext_elm is not None:
            ext_elm.set("cx", str(cx))
            ext_elm.set("cy", str(cy))
        sp_pr = anchor_elm.find(f".//{{{_XDR_NS}}}spPr")
        if sp_pr is not None:
            a_ext = sp_pr.find(f".//{_A_EXT}")
            if a_ext is not None:
                a_ext.set("cx", str(cx))
                a_ext.set("cy", str(cy))

        pic_name = name or path.stem
        cnv = anchor_elm.find(f".//{_XDR_CNVPR}")
        if cnv is not None:
            cnv.set("name", pic_name)

        media_r_id = drawing_part.relate_to(media_part, RT.IMAGE)
        blip = anchor_elm.find(f".//{{{_A_NS}}}blip")
        if blip is not None:
            blip.set(f"{{{OFFICE_REL_NS}}}embed", media_r_id)

        root.append(anchor_elm)
        drawing_part._blob = serialize_xml(root)

        pics = list(iter_pictures(drawing_part, ws))
        return pics[-1]

    def _add_chart_to_sheet(
        self,
        ws: Worksheet,
        chart_type: str,
        *,
        anchor: str,
        data_range: str,
        title: str | None,
        name: str | None,
        to_anchor: str | None = None,
    ) -> Chart:
        if chart_type != "bar":
            raise ValueError(f"unsupported chart_type: {chart_type!r}")

        templates = _default_xlsx_path().parent
        drawing_template = parse_template_xml(
            (templates / "default-chart-anchor.xml").read_bytes()
        )
        chart_template = parse_template_xml(
            (templates / "default-bar-chart.xml").read_bytes()
        )

        chart_partname = self._package.next_partname("/xl/charts/chart%d.xml")
        chart_part = Part(chart_partname, CT.CHART, serialize_xml(chart_template), self._package)
        self._package._add_part(chart_part)

        drawing_parts = drawing_parts_for_worksheet(ws)
        anchor_template = deepcopy(drawing_template.find(_XDR_TWO_CELL))
        if anchor_template is None:
            raise ValueError("chart anchor template missing")

        if drawing_parts:
            drawing_part = drawing_parts[0]
            root = parse_xml(drawing_part.blob)
            anchor_elm = deepcopy(anchor_template)
        else:
            root = parse_xml(
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                b'<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"'
                b' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
                b' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
                b' xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>'
            )
            anchor_elm = deepcopy(anchor_template)
            drawing_partname = self._package.next_partname("/xl/drawings/drawing%d.xml")
            drawing_part = Part(drawing_partname, CT.DRAWING, b"", self._package)
            self._package._add_part(drawing_part)
            r_id = ws._part.relate_to(drawing_part, RT.DRAWING)
            from xlsxedit.worksheet_order import insert_worksheet_child

            drawing_elm = etree.Element(_WS_DRAWING)
            drawing_elm.set(f"{{{OFFICE_REL_NS}}}id", r_id)
            insert_worksheet_child(ws._part.element, drawing_elm)

        col, row = split_address(anchor)
        from_elm = anchor_elm.find(_XDR_FROM)
        to_elm = anchor_elm.find(_XDR_TO)
        if from_elm is not None:
            from_col_elm = from_elm.find(_XDR_COL)
            from_row_elm = from_elm.find(_XDR_ROW)
            to_col_elm = to_elm.find(_XDR_COL) if to_elm is not None else None
            to_row_elm = to_elm.find(_XDR_ROW) if to_elm is not None else None

            # Preserve template span so ``to`` stays below/right of ``from``.
            # (Only updating ``from`` left inverted boxes when anchor row >
            # template ``to`` row — Excel then repairs the drawing.)
            old_from_col = int(from_col_elm.text or "0")
            old_from_row = int(from_row_elm.text or "0")
            if to_col_elm is not None and to_row_elm is not None:
                d_col = max(1, int(to_col_elm.text or "0") - old_from_col)
                d_row = max(1, int(to_row_elm.text or "0") - old_from_row)
            else:
                d_col, d_row = 6, 13

            new_from_col = col_to_index(col)
            new_from_row = row - 1
            from_col_elm.text = str(new_from_col)
            from_row_elm.text = str(new_from_row)
            if to_col_elm is not None and to_row_elm is not None:
                if to_anchor is not None:
                    to_col, to_row = split_address(to_anchor)
                    to_col_elm.text = str(col_to_index(to_col))
                    to_row_elm.text = str(to_row - 1)
                else:
                    to_col_elm.text = str(new_from_col + d_col)
                    to_row_elm.text = str(new_from_row + d_row)

        chart_name = name or "Chart 1"
        frame = anchor_elm.find(_XDR_GRAPHIC_FRAME)
        if frame is not None:
            cnv = frame.find(f".//{_XDR_CNVPR}")
            if cnv is not None:
                cnv.set("name", chart_name)

        chart_r_id = drawing_part.relate_to(chart_part, RT.CHART)
        chart_ref = frame.find(f".//{_C_CHART}") if frame is not None else None
        if chart_ref is not None:
            chart_ref.set(f"{{{OFFICE_REL_NS}}}id", chart_r_id)

        root.append(anchor_elm)
        drawing_part._blob = serialize_xml(root)

        c1, r1, c2, r2 = parse_range(data_range)
        if col_to_index(c2) - col_to_index(c1) < 1:
            raise ValueError("data_range must span at least two columns")
        name_ref = _sheet_abs_ref(ws.name, c2, r1)
        cat_ref = _sheet_abs_range(
            ws.name,
            f"{c1}{r1 + 1}:{c1}{r2}",
        )
        val_ref = _sheet_abs_range(
            ws.name,
            f"{c2}{r1 + 1}:{c2}{r2}",
        )

        charts = list(iter_charts(drawing_part))
        chart = charts[-1]
        chart.set_series_formula(0, name_ref, ws)
        chart.set_series_formula(1, cat_ref, ws)
        chart.set_series_formula(2, val_ref, ws)
        if title is not None:
            chart.title = title
        return chart

    def _add_table_to_sheet(
        self,
        ws: Worksheet,
        ref: str,
        columns: list[str],
        *,
        name: str | None,
        display_name: str | None,
        write_header: bool = False,
    ) -> "Table":
        from xlsxedit.drawing import Table

        columns = list(columns)
        if not columns:
            raise TableError("columns must not be empty")
        seen: set[str] = set()
        for col_name in columns:
            if not isinstance(col_name, str) or not col_name:
                raise TableError(f"column names must be non-empty strings, got {col_name!r}")
            if col_name.casefold() in seen:
                raise TableError(f"duplicate column name {col_name!r} (names ignore case)")
            seen.add(col_name.casefold())

        box = _range_box(ref)
        ref = _box_ref(box)
        first_col, header_row, last_col, last_row = box
        if last_row == header_row:
            raise TableError(f"table range {ref} needs a header row and at least one data row")
        if last_col - first_col + 1 != len(columns):
            raise TableError(
                f"table range {ref} is {last_col - first_col + 1} columns wide "
                f"but {len(columns)} column names were given"
            )
        for other in ws.tables:
            other_box = _table_box(other)
            if other_box is not None and _boxes_overlap(box, other_box):
                raise TableError(f"table range {ref} overlaps table {other.name!r} ({other.ref})")

        used_names, used_ids = _table_names_and_ids(self)
        table_id = max(used_ids, default=0) + 1
        table_name = _next_table_name(used_names, table_id) if name is None else name
        disp = table_name if display_name is None else display_name
        for label, value in (("name", table_name), ("display_name", disp)):
            problem = _table_name_problem(value)
            if problem is None and value.casefold() in used_names:
                problem = "is already used by a table or defined name in this workbook"
            if problem is not None:
                raise TableError(f"table {label} {value!r} {problem}")

        cells = _row_cells(ws, header_row)
        if write_header:
            row_elm = ws._ensure_row(header_row)
            for offset, col_name in enumerate(columns):
                address = join_address(index_to_col(first_col + offset), header_row)
                c_elm = cells.get(address)
                if c_elm is None:
                    c_elm = _insert_cell(row_elm, address)
                Cell(c_elm, ws).value = col_name
            ws._invalidate_bulk_indexes()
        else:
            problems = _header_mismatches(ws, cells, header_row, first_col, columns)
            if problems:
                raise TableError(
                    f"the header cells of table range {ref} do not match the column names: "
                    f"{'; '.join(problems)}; write the header first or pass write_header=True"
                )

        templates = _default_xlsx_path().parent
        table_elm = parse_template_xml((templates / "default-table.xml").read_bytes())

        table_elm.set("ref", ref)
        table_elm.set("name", table_name)
        table_elm.set("displayName", disp)
        _refresh_revision_uids(table_elm)

        af = table_elm.find(f"{{{SML_NS}}}autoFilter")
        if af is not None:
            af.set("ref", ref)

        cols_elm = table_elm.find(f"{{{SML_NS}}}tableColumns")
        if cols_elm is None:
            cols_elm = etree.SubElement(table_elm, f"{{{SML_NS}}}tableColumns")
        for child in list(cols_elm):
            cols_elm.remove(child)
        cols_elm.set("count", str(len(columns)))
        for i, col_name in enumerate(columns, start=1):
            col_elm = etree.SubElement(cols_elm, f"{{{SML_NS}}}tableColumn")
            col_elm.set("id", str(i))
            col_elm.set("name", col_name)

        table_elm.set("id", str(table_id))

        table_partname = self._package.next_partname("/xl/tables/table%d.xml")
        table_part = Part(table_partname, CT.TABLE, serialize_xml(table_elm), self._package)
        self._package._add_part(table_part)

        from xlsxedit.worksheet_order import insert_worksheet_child, reposition_worksheet_child

        tp = ws._part.element.find(_TABLE_PARTS)
        if tp is None:
            tp = etree.Element(_TABLE_PARTS)
            insert_worksheet_child(ws._part.element, tp)
        else:
            reposition_worksheet_child(ws._part.element, tp)
        count = int(tp.get("count", "0")) + 1
        tp.set("count", str(count))
        rel_elm = etree.SubElement(tp, _TABLE_PART)
        r_id = ws._part.relate_to(table_part, RT.TABLE)
        rel_elm.set(f"{{{OFFICE_REL_NS}}}id", r_id)

        return Table(table_elm, table_part)


# --- Table part helpers ------------------------------------------------------

_MAX_ROW = 1048576
_MAX_COL = 16384  # columns A..XFD

_CELL = f"{{{SML_NS}}}c"
_EXT_LST = f"{{{SML_NS}}}extLst"
_DEFINED_NAMES = f"{{{SML_NS}}}definedNames"
_DEFINED_NAME = f"{{{SML_NS}}}definedName"
_TABLE_COLUMNS = f"{{{SML_NS}}}tableColumns"
_TABLE_COLUMN = f"{{{SML_NS}}}tableColumn"

# xr, xr2 and xr3: Excel's revision uids, which must not repeat across parts.
_REVISION_NAMESPACES = frozenset(
    {
        "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
        "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
        "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
    }
)

# Table names follow the defined-name rules: a letter, underscore or
# backslash, then letters, digits, underscores and periods, and never
# something Excel would read as an A1 or R1C1 reference inside the grid.
_TABLE_NAME_RE = re.compile(r"(?:[^\W\d]|\\)[\w.]*")
_A1_NAME_RE = re.compile(r"([A-Za-z]{1,3})(\d+)")
_R1C1_NAME_RE = re.compile(r"[Rr](\d*)(?:[Cc](\d*))?|[Cc](\d*)")


def _new_revision_uid() -> str:
    """A fresh revision uid in Excel's ``{GUID}`` form."""
    return "{" + str(uuid.uuid4()).upper() + "}"


def _refresh_revision_uids(root: etree._Element) -> None:
    """Give every revision ``uid`` attribute in ``root`` a fresh GUID.

    Elements that shared a uid still share one: Excel writes the same uid on
    a ``<table>`` and its ``<autoFilter>``.
    """
    fresh: dict[str, str] = {}
    for elm in root.iter(etree.Element):
        for attr, value in list(elm.attrib.items()):
            qname = etree.QName(attr)
            if qname.localname == "uid" and qname.namespace in _REVISION_NAMESPACES:
                elm.set(attr, fresh.setdefault(value.upper(), _new_revision_uid()))


def _table_name_problem(name) -> str | None:
    """Why ``name`` cannot name a table, or ``None`` when it can."""
    if not isinstance(name, str) or not name:
        return "must be a non-empty string"
    if len(name) > 255:
        return "is longer than 255 characters"
    if not _TABLE_NAME_RE.fullmatch(name):
        return (
            "must start with a letter, underscore or backslash and continue with "
            "letters, digits, underscores or periods"
        )
    if _is_cell_reference(name):
        return "looks like a cell reference"
    return None


def _is_cell_reference(name: str) -> bool:
    """Whether Excel reads ``name`` as an A1 or R1C1 reference inside the grid."""
    a1 = _A1_NAME_RE.fullmatch(name)
    if a1:
        return col_to_index(a1[1]) < _MAX_COL and 1 <= int(a1[2]) <= _MAX_ROW
    r1c1 = _R1C1_NAME_RE.fullmatch(name)
    if r1c1 is None:
        return False
    row, col, lone_col = r1c1.groups()
    if row is None:
        col = lone_col
    return (not row or 1 <= int(row) <= _MAX_ROW) and (not col or 1 <= int(col) <= _MAX_COL)


def _table_names_and_ids(
    workbook: Workbook, *, exclude: Part | None = None
) -> tuple[set[str], set[int]]:
    """Names in use (casefolded) and table ids across the workbook.

    Names are every table's ``name`` and ``displayName`` plus the defined
    names, which share their namespace. Every table part in the package is
    read, so the tables of a sheet still being copied count too.
    """
    names: set[str] = set()
    ids: set[int] = set()
    for part in workbook._package.get_parts_of_type(CT.TABLE):
        if part is exclude:
            continue
        elm = parse_xml(part.blob)
        names.update(v.casefold() for v in (elm.get("name"), elm.get("displayName")) if v)
        table_id = elm.get("id", "")
        if table_id.isdigit():
            ids.add(int(table_id))
    defined = workbook._workbook_part.element.find(_DEFINED_NAMES)
    if defined is not None:
        names.update(
            elm.get("name").casefold()
            for elm in defined.iterchildren(_DEFINED_NAME)
            if elm.get("name")
        )
    return names, ids


def _next_table_name(used: set[str], start: int) -> str:
    """The first ``TableN`` with ``N >= start`` not in ``used`` (casefolded names)."""
    n = max(start, 1)
    while f"table{n}" in used:
        n += 1
    return f"Table{n}"


def _range_box(ref: str) -> tuple[int, int, int, int]:
    """``(first column index, first row, last column index, last row)`` of an A1 range.

    Raises ``InvalidRangeError`` unless the range is ordered and inside the grid.
    """
    c1, r1, c2, r2 = parse_range(ref)
    box = (col_to_index(c1), r1, col_to_index(c2), r2)
    if not (0 <= box[0] <= box[2] < _MAX_COL and 1 <= box[1] <= box[3] <= _MAX_ROW):
        raise InvalidRangeError(f"range {ref!r} is reversed or outside A1:XFD1048576")
    return box


def _box_ref(box: tuple[int, int, int, int]) -> str:
    c1, r1, c2, r2 = box
    return f"{join_address(index_to_col(c1), r1)}:{join_address(index_to_col(c2), r2)}"


def _boxes_overlap(a: tuple[int, int, int, int], b: tuple[int, int, int, int]) -> bool:
    return a[0] <= b[2] and b[0] <= a[2] and a[1] <= b[3] and b[1] <= a[3]


def _table_box(table: Table) -> tuple[int, int, int, int] | None:
    """The box of ``table.ref``, or ``None`` when the ref is missing or malformed."""
    try:
        return _range_box(table.ref or "")
    except InvalidRangeError:
        return None


def _row_cells(ws: Worksheet, row: int) -> dict[str, etree._Element]:
    """``{address: <c>}`` of the cells stored in ``row``, without resolving merges."""
    row_elm = ws._find_row_element(row)
    if row_elm is None:
        return {}
    return {c.get("r", "").upper(): c for c in row_elm.iterchildren(_CELL)}


def _header_mismatches(
    ws: Worksheet,
    cells: dict[str, etree._Element],
    row: int,
    first_col: int,
    names: list[str],
) -> list[str]:
    """Describe each header cell that does not hold its column name as plain text."""
    problems = []
    for offset, name in enumerate(names):
        address = join_address(index_to_col(first_col + offset), row)
        c_elm = cells.get(address)
        cell = None if c_elm is None else Cell(c_elm, ws)
        if cell is not None and cell.has_formula:
            problems.append(f"{address} holds a formula, expected {name!r}")
            continue
        value = None if cell is None else cell.value
        if not (isinstance(value, str) and value == name):
            problems.append(f"{address} {_holds(value)}, expected {name!r}")
    return problems


def _holds(value) -> str:
    return "is empty" if value is None else f"holds {value!r}"


def _insert_cell(row_elm: etree._Element, address: str) -> etree._Element:
    """A new ``<c r=address>`` in ``row_elm``, before the first cell to its right."""
    col = col_to_index(split_address(address)[0])
    c_elm = etree.Element(_CELL)
    c_elm.set("r", address)
    for child in row_elm.iterchildren(_CELL, _EXT_LST):
        r = child.get("r")
        if child.tag == _EXT_LST or (r and col_to_index(split_address(r)[0]) > col):
            child.addprevious(c_elm)
            return c_elm
    row_elm.append(c_elm)
    return c_elm


def _find_table(ws: Worksheet, table) -> Table:
    """The table on ``ws`` that ``table``, a table name or ``Table``, refers to."""
    if isinstance(table, Table):
        for candidate in ws.tables:
            if table._part is not None and candidate._part is table._part:
                return candidate
        raise TableError(f"{table!r} is not a table on sheet {ws.name!r}")
    if not isinstance(table, str):
        raise TypeError(f"table must be a table name or a Table, got {type(table)!r}")

    def is_named(candidate: Table) -> bool:
        names = (candidate._element.get("name"), candidate._element.get("displayName"))
        return table.casefold() in {n.casefold() for n in names if n}

    for candidate in ws.tables:
        if is_named(candidate):
            return candidate
    for other in ws._workbook.worksheets:
        if other.name != ws.name and any(is_named(c) for c in other.tables):
            raise TableError(f"table {table!r} is on sheet {other.name!r}, not {ws.name!r}")
    raise TableError(f"the workbook has no table named {table!r}")


def _plan_table_resize(
    ws: Worksheet,
    target: Table | None,
    rows: list,
    *,
    start_row: int,
    start_col: int,
    header_row: bool,
) -> tuple[Table, str] | None:
    """The table that writing ``rows`` from ``start_row``/``start_col`` resizes, and its ref.

    Without a ``target`` this is the one table on ``ws`` that the written
    range overlaps, if any. ``header_row`` says whether the first row holds
    the column names. The table keeps its columns, and its header row unless
    the rows start at or above it. Raises ``TableError`` when the resize
    would desync the table part: the rows overlap several tables, do not
    span exactly the table's columns, or would put anything but the column
    names in its header row, or the resized table would overlap another.
    """
    width = max(len(row) if isinstance(row, (list, tuple)) else 1 for row in rows)
    if not width:
        return None
    written = (start_col, start_row, start_col + width - 1, start_row + len(rows) - 1)
    tables = ws.tables
    if target is None:
        hits = []
        for candidate in tables:
            box = _table_box(candidate)
            if box is not None and _boxes_overlap(box, written):
                hits.append(candidate)
        if not hits:
            return None
        if len(hits) > 1:
            names = ", ".join(repr(t.name) for t in hits)
            raise TableError(
                f"the written range {_box_ref(written)} overlaps tables {names}; "
                "write within one table, or pass resize_table=False"
            )
        target = hits[0]

    box = _table_box(target)
    if box is None:
        raise TableError(f"table {target.name!r} has a malformed ref {target.ref!r}")
    first_col, table_top, last_col, _ = box
    if (written[0], written[2]) != (first_col, last_col):
        raise TableError(
            f"the written range {_box_ref(written)} does not span exactly the columns of "
            f"table {target.name!r} ({target.ref}); write those columns, or pass "
            "resize_table=False"
        )
    cols_elm = target._element.find(_TABLE_COLUMNS)
    names = _column_names(target)
    count = None if cols_elm is None else cols_elm.get("count")
    if len(names) != width or (count is not None and count != str(len(names))):
        raise TableError(
            f"table {target.name!r} has {len(names)} tableColumn elements (count={count}) "
            f"but its ref {target.ref} is {width} columns wide"
        )

    # Body-only writes keep the table's header row: dropping it from the ref
    # would make Excel repair the table part.
    top = written[1] if header_row else min(table_top, written[1])
    bottom = written[3]
    new_box = (first_col, top, last_col, bottom)
    new_ref = _box_ref(new_box)
    for other in tables:
        other_box = _table_box(other)
        if other._part is not target._part and other_box and _boxes_overlap(other_box, new_box):
            raise TableError(
                f"resizing table {target.name!r} to {new_ref} would overlap table "
                f"{other.name!r} ({other.ref})"
            )
    if _has_header_row(target) and top == written[1]:
        if bottom == top:
            raise TableError(
                f"resizing table {target.name!r} to {new_ref} would leave it without a data row"
            )
        first = rows[0] if isinstance(rows[0], (list, tuple)) else (rows[0],)
        labels = [*first, *[None] * (width - len(first))]
        problems = _label_mismatches(labels, names, first_col, top)
        if problems:
            raise TableError(
                f"the first written row would become the header row of table "
                f"{target.name!r} but does not hold its column names: {'; '.join(problems)}; "
                "write the column names there, or start below the header row with header=False"
            )
    return target, new_ref


def _check_header_row_write(
    ws: Worksheet, labels: list, *, start_row: int, start_col: int
) -> None:
    """Refuse header ``labels`` that would overwrite a table's header with other names.

    For a header-only write, which resizes no table.
    """
    written = (start_col, start_row, start_col + len(labels) - 1, start_row)
    for table in ws.tables:
        box = _table_box(table)
        if box is None or box[1] != start_row or not _has_header_row(table):
            continue
        if not _boxes_overlap(box, written):
            continue
        first, last = max(box[0], written[0]), min(box[2], written[2])
        names = _column_names(table)[first - box[0] : last - box[0] + 1]
        heads = labels[first - start_col : last - start_col + 1]
        problems = _label_mismatches(heads, names, first, start_row)
        if problems:
            raise TableError(
                f"the header row would overwrite the header of table {table.name!r} with "
                f"other text: {'; '.join(problems)}; write its column names, or pass "
                "resize_table=False"
            )


def _column_names(table: Table) -> list[str | None]:
    cols_elm = table._element.find(_TABLE_COLUMNS)
    return [] if cols_elm is None else [c.get("name") for c in cols_elm.findall(_TABLE_COLUMN)]


def _has_header_row(table: Table) -> bool:
    from xlsxedit._ooxml_order import ATTRIBUTE_DEFAULTS

    default = ATTRIBUTE_DEFAULTS["CT_Table"]["headerRowCount"]
    return table._element.get("headerRowCount", default) != "0"


def _label_mismatches(labels: list, names: list, first_col: int, row: int) -> list[str]:
    """Describe each label, heading the columns from ``first_col`` on, that is not its name."""
    return [
        f"{join_address(index_to_col(first_col + i), row)} {_holds(label)}, expected {name!r}"
        for i, (label, name) in enumerate(zip(labels, names))
        if not (isinstance(label, str) and label == name)
    ]
