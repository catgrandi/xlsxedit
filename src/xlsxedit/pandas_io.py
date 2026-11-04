"""Optional pandas ExcelWriter / ExcelFile engine (``engine=\"xlsxedit\"``).

Pandas is an optional dependency. Call :func:`register` once before use::

    import xlsxedit.pandas_io as xpi
    xpi.register()

    with pd.ExcelWriter("template.xlsx", engine="xlsxedit", mode="a") as w:
        df.to_excel(w, sheet_name="Data", startrow=4, header=False, index=False)

    df = pd.read_excel("book.xlsx", engine="xlsxedit")

Pass xlsxedit options via ``engine_kwargs`` (writer: ``large``, ``mode``,
``template_rows``, ``clear_range``, ``row_styles``, ``column_styles``,
``string_columns``; reader: ``large``).

Pandas has no public ``register_reader``; :func:`register` also patches
``ExcelFile._engines`` so ``engine=\"xlsxedit\"`` works for ``read_excel``.
Or use :func:`read_excel`, which registers and forces the xlsxedit engine.
"""

from __future__ import annotations

from collections import defaultdict
from collections.abc import Callable
from typing import Any, NamedTuple

from xlsxedit.bulk_styles import validate_style_sources
from xlsxedit.merge import parse_range
from xlsxedit.oxml.address import index_to_col, join_address
from xlsxedit.range_set import CellRange
from xlsxedit.row_shift import (
    append_merges,
    check_block_in_grid,
    check_style_specs,
    check_values_writable,
)
from xlsxedit.styles import check_style_text
from xlsxedit.workbook import Workbook, _sheet_name_key

try:
    from pandas.io.excel import ExcelWriter, register_writer
    from pandas.io.excel._base import BaseExcelReader, ExcelFile
except ImportError:  # pragma: no cover
    ExcelWriter = object  # type: ignore[misc, assignment]
    register_writer = None  # type: ignore[assignment]
    BaseExcelReader = object  # type: ignore[misc, assignment]
    ExcelFile = None  # type: ignore[assignment]


def register() -> None:
    """Register ``engine=\"xlsxedit\"`` with pandas (writer + reader)."""
    if register_writer is None or ExcelFile is None:
        raise ImportError(
            "pandas is required for the xlsxedit Excel engine. "
            "Install with: pip install xlsxedit[pandas]"
        )
    register_writer(XlsxeditWriter)
    engines = dict(ExcelFile._engines)
    engines["xlsxedit"] = XlsxeditReader
    ExcelFile._engines = engines


def read_excel(io, *args, **kwargs):
    """``pd.read_excel`` with ``engine=\"xlsxedit\"`` after :func:`register`."""
    register()
    kwargs.setdefault("engine", "xlsxedit")
    import pandas as pd

    return pd.read_excel(io, *args, **kwargs)


class XlsxeditReader(BaseExcelReader):
    """pandas Excel reader backed by :class:`xlsxedit.Workbook`."""

    @property
    def _workbook_class(self):
        return Workbook

    def load_workbook(self, filepath_or_buffer, engine_kwargs):
        engine_kwargs = dict(engine_kwargs or {})
        large = bool(engine_kwargs.pop("large", False))
        return Workbook.open(filepath_or_buffer, large=large)

    @property
    def sheet_names(self) -> list[str]:
        return list(self.book.sheetnames)

    def get_sheet_by_name(self, name: str):
        self.raise_if_bad_sheet_by_name(name)
        return self.book[name]

    def get_sheet_by_index(self, index: int):
        self.raise_if_bad_sheet_by_index(index)
        return self.book.worksheets[index]

    def get_sheet_data(self, sheet, file_rows_needed: int | None = None):
        return sheet.values(max_rows=file_rows_needed, errors_as_nan=True)


class _Block(NamedTuple):
    """The cells of one ``to_excel`` call: rows of values from ``start_row`` and the
    0-based ``start_col_idx``, the number format of each date cell, and the merges."""

    rows: list[tuple]
    start_row: int
    start_col_idx: int
    formats: dict[tuple[int, int], str]
    merges: list[CellRange]


def _merges_to_add(existing: list[CellRange], new: list[CellRange]) -> list[CellRange]:
    """The merges of ``new`` that ``existing`` does not hold already, so that writing a
    frame over itself keeps its merges; ``ValueError``, as ``Worksheet.merge_cells``
    raises it, when one overlaps another merge."""
    held = {(a.min_col, a.min_row, a.max_col, a.max_row) for a in existing}
    added = []
    for area in new:
        bounds = (area.min_col, area.min_row, area.max_col, area.max_row)
        if bounds not in held:
            held.add(bounds)
            added.append(area)
    tagged = sorted(
        [(area, False) for area in existing] + [(area, True) for area in added],
        key=lambda item: (item[0].min_row, item[0].min_col),
    )
    reaching: list[tuple[CellRange, bool]] = []  # the merges that reach the current row
    for area, is_new in tagged:
        reaching = [(other, new_one) for other, new_one in reaching if other.max_row >= area.min_row]
        for other, other_new in reaching:
            if (is_new or other_new) and area.intersects(other):
                mine, theirs = (area, other) if is_new else (other, area)
                raise ValueError(f"merge range {str(mine)!r} overlaps existing {str(theirs)!r}")
        reaching.append((area, is_new))
    return added


class XlsxeditWriter(ExcelWriter):
    """pandas ``ExcelWriter`` implementation backed by :class:`xlsxedit.Workbook`."""

    _engine = "xlsxedit"
    _supported_extensions = (".xlsx", ".xlsm", ".xltx")

    def __init__(
        self,
        path,
        engine: str | None = None,
        date_format: str | None = None,
        datetime_format: str | None = None,
        mode: str = "w",
        storage_options=None,
        if_sheet_exists: str | None = None,
        engine_kwargs: dict[str, Any] | None = None,
        **kwargs,
    ) -> None:
        if ExcelWriter is object:
            raise ImportError(
                "pandas is required for the xlsxedit ExcelWriter engine. "
                "Install with: pip install xlsxedit[pandas]"
            )

        engine_kwargs = dict(engine_kwargs or {})
        engine_kwargs.update(kwargs)

        self._write_mode = engine_kwargs.pop("mode", "overwrite")
        if self._write_mode not in ("overwrite", "insert"):
            raise ValueError(
                f"engine_kwargs mode must be 'overwrite' or 'insert', "
                f"got {self._write_mode!r}"
            )
        self._large = bool(engine_kwargs.pop("large", False))
        self._template_rows = engine_kwargs.pop("template_rows", None)
        self._row_styles = engine_kwargs.pop("row_styles", None)
        self._column_styles = engine_kwargs.pop("column_styles", None)
        self._clear_range = engine_kwargs.pop("clear_range", None)
        self._string_columns = engine_kwargs.pop("string_columns", None)
        self._shift_drawings = bool(engine_kwargs.pop("shift_drawings", True))
        self._pending_default_replace = False

        super().__init__(
            path,
            mode=mode,
            storage_options=storage_options,
            if_sheet_exists=if_sheet_exists,
            engine_kwargs={},
            date_format=date_format,
            datetime_format=datetime_format,
        )

        if "r+" in self._mode:
            handle = self._handles.handle
            handle.seek(0)
            self._book = Workbook.open(handle, large=self._large)
            handle.seek(0)
        else:
            self._book = Workbook.create()
            self._pending_default_replace = True

    @property
    def book(self) -> Workbook:
        return self._book

    @property
    def sheets(self) -> dict[str, Any]:
        return {name: self._book[name] for name in self._book.sheetnames}

    def _unique_sheet_name(self, base: str) -> str:
        """``base``, or ``base`` with the first number appended that no sheet has,
        ignoring case as Excel does, cut so that the name keeps to 31 characters."""
        taken = {_sheet_name_key(name) for name in self._book.sheetnames}
        if _sheet_name_key(base) not in taken:
            return base
        n = 1
        while True:
            suffix = str(n)
            stem = base
            while len((stem + suffix).encode("utf-16-le")) // 2 > 31:
                stem = stem[:-1]
            if _sheet_name_key(stem + suffix) not in taken:
                return stem + suffix
            n += 1

    def _plan_sheet(self, sheet_name: str) -> tuple[Any, Callable[[], Any], bool]:
        """Where this write goes, changing nothing yet: ``(sheet, make, changes)``.

        ``sheet`` is the existing sheet the cells go into, ``None`` for a new one;
        ``make()`` makes the change that gives the sheet to write, and ``changes``
        says whether that changes the workbook. A name that matches a sheet
        ignoring case names that sheet, as in Excel. Raises, before anything
        changes, for ``if_sheet_exists="error"`` and for a sheet name Excel rejects.
        """
        book = self._book
        if self._pending_default_replace and book.sheetnames == ["Sheet1"]:
            # Mode "w": the first write takes over the blank Sheet1 of the new workbook.
            blank = book["Sheet1"]
            if sheet_name != "Sheet1":
                book._check_new_sheet_name(sheet_name, replacing="Sheet1")

            def take_over():
                self._pending_default_replace = False
                if sheet_name == "Sheet1":
                    return blank
                return book.rename_worksheet("Sheet1", sheet_name)

            return blank, take_over, sheet_name != "Sheet1"

        existing = book._existing_sheet_name(sheet_name)
        if existing is None:
            book._check_new_sheet_name(sheet_name)
            return None, lambda: book.add_worksheet(sheet_name), True
        ws = book[existing]
        if "r+" not in self._mode:  # mode "w": a later to_excel to a sheet writes over it
            return ws, lambda: ws, False
        how = self._if_sheet_exists
        if how == "error":
            raise ValueError(
                f"Sheet '{sheet_name}' already exists and if_sheet_exists is set to 'error'."
            )
        if how == "overlay":
            return ws, lambda: ws, False
        if how == "replace":
            book._check_new_sheet_name(sheet_name, replacing=existing)
            return None, lambda: self._replace_sheet(existing, sheet_name), True
        if how == "new":
            name = self._unique_sheet_name(sheet_name)
            book._check_new_sheet_name(name)
            return None, lambda: book.add_worksheet(name), True
        raise ValueError(
            f"'{how}' is not valid for if_sheet_exists. "
            "Valid options are 'error', 'new', 'replace' and 'overlay'."
        )

    def _replace_sheet(self, old: str, new: str):
        """Put a new empty sheet named ``new`` in place of sheet ``old``. The new sheet
        is added, last, before the old one is removed, so this works on a workbook's
        only sheet too."""
        book = self._book
        aside = self._unique_sheet_name("replaced")
        book.rename_worksheet(old, aside)
        ws = book.add_worksheet(new)
        book.remove_worksheet(aside)
        return ws

    def _style_kwargs(self) -> dict[str, Any]:
        return {
            "template_rows": self._template_rows,
            "row_styles": self._row_styles,
            "column_styles": self._column_styles,
            "string_columns": self._string_columns,
        }

    def _layout(self, cells, startrow: int, startcol: int) -> _Block:
        """The cells pandas formatted, as rows of values from their top-left cell."""
        by_row: dict[int, dict[int, Any]] = defaultdict(dict)
        formats: dict[tuple[int, int], str] = {}
        merges: list[CellRange] = []
        for cell in cells:
            val, fmt = self._value_with_fmt(cell.val)
            if isinstance(val, str) and not val:
                # pandas writes NaN, NaT and NA as na_rep, "" by default; the openpyxl
                # and XlsxWriter engines write "" as an empty cell too.
                val = None
            abs_row = startrow + cell.row + 1
            abs_col = startcol + cell.col  # 0-based
            by_row[abs_row][abs_col] = val
            if fmt is not None:
                formats[(abs_row, abs_col)] = fmt
            if cell.mergestart is not None and cell.mergeend is not None:
                # Match openpyxl ExcelWriter coordinates.
                area = CellRange.from_bounds(
                    abs_col + 1,
                    abs_row,
                    startcol + cell.mergeend + 1,
                    startrow + cell.mergestart + 1,
                )
                if not area.is_cell:
                    merges.append(area)
        if not by_row:
            return _Block([], 0, 0, formats, merges)
        min_row, max_row = min(by_row), max(by_row)
        min_col = min(min(cols) for cols in by_row.values())
        max_col = max(max(cols) for cols in by_row.values())
        rows = [
            tuple(by_row.get(r, {}).get(c) for c in range(min_col, max_col + 1))
            for r in range(min_row, max_row + 1)
        ]
        return _Block(rows, min_row, min_col, formats, merges)

    def _write_cells(
        self,
        cells,
        sheet_name: str | None = None,
        startrow: int = 0,
        startcol: int = 0,
        freeze_panes: tuple[int, int] | None = None,
        autofilter_range: str | None = None,
    ) -> None:
        """Write one ``to_excel`` call. Every value, merge, number format and sheet
        name is checked before the workbook changes, because pandas saves the
        workbook even when ``to_excel`` raises."""
        # freeze_panes / autofilter_range deferred (v1 no-op).
        _ = freeze_panes, autofilter_range

        sheet_name = self._get_sheet_name(sheet_name)
        block = self._layout(cells, startrow, startcol)
        target, make_sheet, changes = self._plan_sheet(sheet_name)
        inserting = self._write_mode == "insert"
        clear_range = None if inserting else self._clear_range
        style_kwargs = self._style_kwargs()

        validate_style_sources(
            template_rows=style_kwargs["template_rows"], row_styles=style_kwargs["row_styles"]
        )
        check_style_specs(style_kwargs["row_styles"], style_kwargs["column_styles"])
        for fmt in set(block.formats.values()):
            check_style_text({"num_format": fmt})
        if clear_range is not None:
            parse_range(clear_range)
        held = [] if inserting or target is None else [area for area, _ in target._merge_index.areas]
        merges = _merges_to_add(held, block.merges)  # an insert checks its own against the sheet's
        rows = block.rows
        if rows and changes:  # otherwise the write checks its rows before changing anything
            width = max(len(row) for row in rows)
            check_block_in_grid("to_excel", block.start_row, block.start_col_idx, len(rows), width)
            rows = check_values_writable(
                "to_excel",
                rows,
                start_col_idx=block.start_col_idx,
                string_columns=style_kwargs["string_columns"],
            )

        ws = make_sheet()
        if not rows:
            if clear_range is not None:
                ws.clear_range(clear_range)
            return
        if inserting:
            ws.insert_rows(
                rows,
                at_cell=join_address(index_to_col(block.start_col_idx), block.start_row),
                shift_drawings=self._shift_drawings,
                _merges=merges,
                **style_kwargs,
            )
        else:
            ws._write_rows_at(
                rows,
                start_row=block.start_row,
                start_col_idx=block.start_col_idx,
                clear_range=clear_range,
                method="to_excel",
                **style_kwargs,
            )
        for (row, col_idx), fmt in block.formats.items():
            ws[join_address(index_to_col(col_idx), row)].apply_number_format(fmt)
        if merges and not inserting:
            append_merges(ws._part.element, merges)
            ws._part.mark_dirty()
            ws._invalidate_merge_map()
            ws._invalidate_bulk_indexes()

    def _save(self) -> None:
        handle = self._handles.handle
        if "r+" in self._mode:
            handle.seek(0)
            self._book.save(handle)
            handle.truncate()
        else:
            self._book.save(handle)
