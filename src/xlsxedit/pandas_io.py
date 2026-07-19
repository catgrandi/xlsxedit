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
from typing import Any

from xlsxedit.oxml.address import index_to_col, join_address
from xlsxedit.workbook import Workbook

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
        if base not in self._book.sheetnames:
            return base
        n = 1
        while f"{base}{n}" in self._book.sheetnames:
            n += 1
        return f"{base}{n}"

    def _resolve_sheet(self, sheet_name: str):
        names = self._book.sheetnames
        exists = sheet_name in names

        if exists and self._if_sheet_exists != "new":
            if "r+" in self._mode:
                if self._if_sheet_exists == "error":
                    raise ValueError(
                        f"Sheet '{sheet_name}' already exists and "
                        f"if_sheet_exists is set to 'error'."
                    )
                if self._if_sheet_exists == "replace":
                    if len(names) == 1:
                        ws = self._book[sheet_name]
                        for cell in list(ws.iter_cells()):
                            cell.clear()
                        return ws
                    self._book.remove_worksheet(sheet_name)
                    return self._book.add_worksheet(sheet_name)
                if self._if_sheet_exists == "overlay":
                    return self._book[sheet_name]
                raise ValueError(
                    f"'{self._if_sheet_exists}' is not valid for if_sheet_exists. "
                    "Valid options are 'error', 'new', 'replace' and 'overlay'."
                )
            return self._book[sheet_name]

        if self._pending_default_replace and names == ["Sheet1"]:
            self._pending_default_replace = False
            if sheet_name != "Sheet1":
                return self._book.rename_worksheet("Sheet1", sheet_name)
            return self._book["Sheet1"]

        if self._if_sheet_exists == "new" and exists:
            sheet_name = self._unique_sheet_name(sheet_name)
        elif exists:
            return self._book[sheet_name]
        return self._book.add_worksheet(sheet_name)

    def _style_kwargs(self) -> dict[str, Any]:
        return {
            "template_rows": self._template_rows,
            "row_styles": self._row_styles,
            "column_styles": self._column_styles,
            "string_columns": self._string_columns,
        }

    def _write_cells(
        self,
        cells,
        sheet_name: str | None = None,
        startrow: int = 0,
        startcol: int = 0,
        freeze_panes: tuple[int, int] | None = None,
        autofilter_range: str | None = None,
    ) -> None:
        # freeze_panes / autofilter_range deferred (v1 no-op).
        _ = freeze_panes, autofilter_range

        sheet_name = self._get_sheet_name(sheet_name)
        ws = self._resolve_sheet(sheet_name)

        if self._write_mode == "overwrite" and self._clear_range is not None:
            ws.clear_range(self._clear_range)

        # Buffer pandas cells → dense row grid (values only; ignore formatter styles).
        by_row: dict[int, dict[int, Any]] = defaultdict(dict)
        formats: dict[tuple[int, int], str] = {}
        merges: list[tuple[int, int, int, int]] = []
        for cell in cells:
            val, fmt = self._value_with_fmt(cell.val)
            abs_row = startrow + cell.row + 1
            abs_col = startcol + cell.col  # 0-based
            by_row[abs_row][abs_col] = val
            if fmt is not None:
                formats[(abs_row, abs_col)] = fmt
            if cell.mergestart is not None and cell.mergeend is not None:
                # Match openpyxl ExcelWriter coordinates.
                merges.append(
                    (
                        abs_row,
                        abs_col,
                        startrow + cell.mergestart + 1,
                        startcol + cell.mergeend,
                    )
                )

        if not by_row:
            return

        min_row = min(by_row)
        max_row = max(by_row)
        min_col = min(min(cols) for cols in by_row.values())
        max_col = max(max(cols) for cols in by_row.values())

        row_list = []
        for r in range(min_row, max_row + 1):
            cols = by_row.get(r, {})
            row_list.append(tuple(cols.get(c) for c in range(min_col, max_col + 1)))

        style_kwargs = self._style_kwargs()
        if self._write_mode == "insert":
            ws.insert_rows(
                row_list,
                at_cell=join_address(index_to_col(min_col), min_row),
                **style_kwargs,
            )
        else:
            ws._write_rows_at(
                row_list,
                start_row=min_row,
                start_col_idx=min_col,
                **style_kwargs,
            )

        for (row, col_idx), fmt in formats.items():
            ws[join_address(index_to_col(col_idx), row)].apply_number_format(fmt)

        for r1, c1, r2, c2 in merges:
            ws.merge_cells(
                f"{join_address(index_to_col(c1), r1)}:{join_address(index_to_col(c2), r2)}"
            )

    def _save(self) -> None:
        handle = self._handles.handle
        if "r+" in self._mode:
            handle.seek(0)
            self._book.save(handle)
            handle.truncate()
        else:
            self._book.save(handle)
