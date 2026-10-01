# Features & API reference

Supported public API for **xlsxedit** — surgical `.xlsx` editing that preserves formatting.

Docs: [xlsxedit.jonasruilong.com](https://xlsxedit.jonasruilong.com)

**Tutorials:** `tutorial/pandas_tutorial.py` (pandas engine), `tutorial/run_tutorial.py` (full API walkthrough), `tutorial/export_pandas_example.py` (advanced bulk export), `tutorial/bench_large_export.py` (performance)

**Architecture:** [How xlsxedit works](how-xlsxedit-works.md)

---

## Workbook

| API | Summary | Notes |
|-----|---------|-------|
| `Workbook(path, *, large=False)` | Load an existing `.xlsx`, OPC folder, or binary file-like | Same as `Workbook.open(path)` |
| `Workbook.open(path, *, large=False)` | Load an existing `.xlsx` | Same as `Workbook(path)`; path, unpacked folder, or file-like; `large=True` for bulk |
| `Workbook()` | New workbook from bundled template | Same as `Workbook.create()` |
| `Workbook.create()` | New workbook from bundled template | Same as `Workbook()` |
| `wb.save(path, *, include_orphans=False)` | Write package to a path or binary file-like |  |
| `wb[sheet_name]` | Get worksheet by name | `KeyError` if missing |
| `wb.sheetnames` | List of sheet names | |
| `wb.worksheets` | List of `Worksheet` objects | |
| `wb.add_worksheet(name)` | Add a sheet | Blank template |
| `wb.copy_worksheet(name, new_name)` | Duplicate an existing sheet | Cells, styles, merges, drawings/tables; images shared |
| `wb.rename_worksheet(old, new)` | Rename a sheet | |
| `wb.remove_worksheet(name)` | Remove a sheet | Cannot remove last sheet |
| `wb.replace(old, new, *, value_type=None)` | SAR on all sheets | Substring or whole-cell typed; sets `fullCalcOnLoad` when any cell changes |
| `wb.set_full_calc_on_load()` | Request Excel full recalc on open | Sets `calcPr/@fullCalcOnLoad`; use after non-`replace` edits |
| `wb.find(value, *, sheet=None)` | First cell with exact ``value`` | All sheets if ``sheet`` omitted; skips formulas |
| `wb.findall(value, *, sheet=None)` | All matching cells as `list[Cell]` | Sheet order, then row-major; `[]` if none |
| `wb.replace_image(name, image_path)` | Replace picture by drawing name | |
| `wb.insert_image_at_placeholder(placeholder, image_path, *, width, height, max_width, max_height)` | Clear `{placeholder}` cell, insert image | Size kwargs (px) forwarded to `add_image`; width-only keeps aspect |
| `wb.write_dataframe(df, *, sheet, at_cell, header, mode, …)` | Bulk write rows | See [large-data-export.md](large-data-export.md) |
| `Workbook.export_to_template(template, df, save_path, **kwargs)` | Open template + `write_dataframe` + save | Class convenience |
| `wb.properties` | Core document metadata (`docProps/core.xml`) | `title`, `author`, `subject`, `keywords`, `comments`, `category`, `last_modified_by`, `created`, `modified` |
| `wb.orphan_partnames` | ZIP members not on relationship graph | Advanced |
| `wb.shared_strings` | Shared string table (advanced) | |
| `wb.styles` | Styles part (advanced) | |

`.xlsm` (macro-enabled) and `.xltx` / `.xltm` (template) workbooks open, edit, and save too — the workbook content type is preserved and VBA parts round-trip untouched (no macro inspection or signing).

---

## Worksheet

Access via `wb["Sheet1"]` or `ws = wb.worksheets[0]`. Cells via `ws["B2"]`.

| API | Summary | Notes |
|-----|---------|-------|
| `ws.name` | Sheet name | Read-only; rename via `wb.rename_worksheet` |
| `ws[address]` | Get `Cell` at address | e.g. `"A1"`, `"BC100"` |
| `ws.cells` | All populated cells | |
| `ws.iter_cells()` | Iterate cells | |
| `ws.iter_rows()` | Iterate `(row_num, [cells])` | |
| `ws.values(*, max_rows=None, errors_as_nan=False)` | Dense `list[list]` of cell values | Blanks as `""`; merges anchored top-left; `errors_as_nan` maps Excel errors to NaN |
| `ws.column_dimensions["C"]` | Column width proxy | `.width` get/set |
| `ws.row_dimensions[4]` | Row height proxy | `.height` get/set |
| `ws.merge_cells(cell_range)` | Merge range | e.g. `"A1:C1"` |
| `ws.unmerge_cells(cell_range)` | Unmerge range | |
| `ws.merged_ranges` | Merged ranges as `list[str]` | e.g. `["A1:C1"]` |
| `ws.clear_range(cell_range)` | Clear values in range | Returns count cleared |
| `ws.write_rows(rows, at_cell / at_row, …)` | Write rows at fixed positions | Bulk styling kwargs |
| `ws.insert_rows(rows, at_cell, …)` | Insert rows, shift below down | Shifts merges/CF/tables; does **not** rewrite formulas or drawing/chart anchors |
| `ws.insert_columns(cols, at_col / at_cell, …)` | Insert columns, shift right | ``cols`` = column vectors (top→bottom); same caveats as ``insert_rows`` |
| `ws.update_dimension()` | Refresh `<dimension>` | Usually automatic |
| `ws.replace(old, new, *, value_type=None)` | SAR on this sheet only | Sets `fullCalcOnLoad` when any cell changes |
| `ws.find(value)` | First cell with exact ``value`` | Skips formulas; ``None`` if missing |
| `ws.findall(value)` | All matching cells as `list[Cell]` | `[]` if none |
| `ws.images` | List of `Picture` | Read existing drawings |
| `ws.add_image(path, *, anchor, width, height, name, offset_x, offset_y)` | Insert image | Pixel offsets within anchor cell (default 0) |
| `ws.charts` | List of `Chart` | |
| `ws.add_chart(type, *, anchor, data_range, title, name, to_anchor)` | Add bar chart | `chart_type="bar"` today; flush from template; ``to_anchor`` optional end cell |
| `ws.tables` | List of `Table` | |
| `ws.add_table(cell_range, columns, *, name, display_name)` | Add Excel table | `name` defaults to `TableN`, N being the new table's id or the next free number |
| `ws.conditional_formatting` | Read CF blocks | colorScale, dataBar, cellIs, … |
| `ws.add_conditional_formatting(cell_range, *, operator, formula, …)` | Add `cellIs` rule | With optional dxf colors |
| `ws.add_color_scale_formatting(cell_range)` | Add color-scale rule | |
| `ws.expand_conditional_formatting(end_row)` | Extend CF ranges after bulk insert | Optional with `write_dataframe` |

---

## Cell

| API | Summary | Notes |
|-----|---------|-------|
| `cell.value` | Get/set typed value | str, int, float, bool, `date`, `datetime`, `None`; overwriting a formula drops `calcChain.xml` (Excel rebuilds it) |
| `cell.data_type` | Raw OOXML cell type | `"s"`, `"str"`, `"inlineStr"`, `"b"`, `"e"`, or `None` (numeric) |
| `cell.formula` | Get/set formula string | e.g. `"=A2*2"`; setting to `None` removes the formula |
| `cell.clear()` | Clear cell content | Keeps style index; drops calcChain if cell had a formula |
| `cell.address` | Cell address | |
| `cell.worksheet` | Owning worksheet | |
| `cell.offset(cols=0, rows=0)` | Neighbor cell | e.g. `offset(cols=1, rows=1)` is one right and one down |
| `cell.style` | `CellStyle` read proxy | bold, colors, fonts, alignment, num format |
| `cell.apply_style(**kwargs)` | Write style | `bold`, `font_color`, `bg_color`, `font_size`, … |
| `cell.apply_date_format()` | Apply standard date format | |
| `cell.apply_number_format(code)` | Apply custom number format | e.g. `"$#,##0.00"` |
| `cell.hyperlink` | `Hyperlink` proxy | `.url`, `.location`, `.display` |
| `cell.replace(old, new)` | Substring replace in cell text | |
| `cell.has_formula` | Whether cell has formula | |
| `cell.is_string` | Shared-string cell | |

### CellStyle (read via `cell.style`)

`bold`, `italic`, `underline`, `font_size`, `font_name`, `font_color`, `bg_color`, `horizontal_align`, `vertical_align`, `num_format`, `is_date`, `is_percent`

---

## Picture

From `ws.images` or `ws.add_image(...)`.

| API | Summary |
|-----|---------|
| `pic.name` | Drawing name get/set |
| `pic.anchor` | Top-left anchor cell get/set |
| `pic.offset_x`, `pic.offset_y` | Inset within anchor cell (pixels; default 0 = flush) |
| `pic.width`, `pic.height` | Size in pixels get/set |
| `pic.media_path` | Path inside package, e.g. `xl/media/image1.jpeg` |
| `pic.replace(image_path)` | Swap image bytes |

---

## Chart

From `ws.charts` or `ws.add_chart(...)`.

| API | Summary |
|-----|---------|
| `chart.name` | Drawing name get/set |
| `chart.anchor` | Top-left anchor cell get/set (preserves size span) |
| `chart.to_anchor` | Bottom-right corner cell get/set |
| `chart.offset_x`, `chart.offset_y` | Pixel inset within anchor cell (post-create tweak; add is flush) |
| `chart.title` | Chart title get/set |
| `chart.partname` | Chart part path |
| `chart.set_series_formula(...)` | Update series range |

---

## Table

From `ws.tables` or `ws.add_table(...)`.

| API | Summary |
|-----|---------|
| `table.name` | Table name |
| `table.ref` | Range ref, e.g. `A1:B10` |
| `table.columns` | Column header names |
| `table.resize(ref)` | Change table range |

---

## Search-and-replace

| `value_type` | Match | Write |
|--------------|-------|-------|
| `None` (default) | Substring in text cells | Replace substring in SST / inline string |
| `"text"` | Same as default — substring in text cells | Replace substring (not whole-cell) |
| `"number"` | Whole cell equals `old` | Numeric `<v>` |
| `"date"` | Whole cell equals `old` | Date serial; preserves date style |

Formula cells are always skipped.

Image SAR: `wb.replace_image(name, path)`, `wb.insert_image_at_placeholder("{logo}", path)`.

---

## Bulk export

| API | Summary |
|-----|---------|
| `wb.write_dataframe(df, …)` | Main entry — pandas-like `columns` + `itertuples` / `values` |
| `mode="overwrite"` | Write at fixed rows (default) |
| `mode="insert"` | Insert rows, shift content below |
| `template_rows` | Copy styles from existing rows (`int` or `list[int]`) |
| `row_styles` | List of style dicts cycled per row |
| `column_styles` | List of style dicts per column |
| `clear_range` | Clear before overwrite when re-exporting fewer rows |
| `resize_table` | Resize the table the written rows overlap (overwrite mode) |
| `table` | Table to resize instead (name or `Table`), e.g. to grow a table from below |
| `expand_conditional_formatting` | Extend conditional-formatting ranges |

Full guide: [large-data-export.md](large-data-export.md)

Pandas ExcelWriter / reader engine: [pandas.md](pandas.md) (`pip install xlsxedit[pandas]` then `xlsxedit.pandas_io.register()`).

---

## Exceptions

All library errors derive from `XlsxeditError`, so you can catch everything with a single `except xlsxedit.XlsxeditError`. Each subclass also inherits the builtin it replaces, so existing `except ValueError` / `except KeyError` handlers keep working.

| Exception | Also a | Raised when |
|-----------|--------|-------------|
| `XlsxeditError` | `Exception` | Base class for all xlsxedit errors |
| `WorksheetNotFoundError` | `KeyError` | A worksheet name is not present |
| `DuplicateWorksheetError` | `ValueError` | Adding/renaming to an existing sheet name |
| `InvalidRangeError` | `ValueError` | Malformed cell address or range (e.g. `"A1:C3"`) |
| `InvalidColorError` | `ValueError` | Invalid color value |
| `InvalidImageError` | `ValueError` | Unsupported or corrupt image data |
| `MissingPartError` | `RuntimeError` | A required package part is absent |

---

## Fidelity-tested fixtures

Round-trip tests in `tests/test_fidelity.py` on:

`TestBook`, `MergedCells`, `DifferentCellTypes`, `CustomCellSize`, `ConditionalFormatting`, `SimpleFormula`, `images`, `ChartsAndTables`, `TextSytle`, `hyperlink`

---

## Not yet supported

| Area | Status |
|------|--------|
| Pivot tables | Preserved as opaque parts; not editable |
| Comments / threaded comments | Not exposed |
| VBA / macros | `.xlsm` opens/saves; VBA is opaque round-trip only |
| All chart types | Add: bar template; read: existing charts |
| All CF rule types | Read most; write `cellIs` and color scale |
| Byte-identical ZIP | Semantic preservation, not diff tool |
| Style authoring from scratch | Clone/apply xf; not full theme designer |

See roadmap in [how-xlsxedit-works.md](how-xlsxedit-works.md) §9.
