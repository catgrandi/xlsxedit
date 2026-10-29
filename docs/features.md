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
| `wb.save(path, *, include_orphans=False)` | Write package to a path or binary file-like | Raises `ValueError`, writing nothing, if two parts would share a name (ignoring case) |
| `wb[sheet_name]` | Get worksheet by name | `KeyError` if missing |
| `wb.sheetnames` | List of sheet names | |
| `wb.worksheets` | List of `Worksheet` objects | |
| `wb.add_worksheet(name)` | Add a sheet | Blank template |
| `wb.copy_worksheet(name, new_name)` | Duplicate an existing sheet | Cells, styles, merges, drawings/tables; images shared; drops `calcChain.xml`. A copy that fails changes nothing |
| `wb.rename_worksheet(old, new)` | Rename a sheet | |
| `wb.remove_worksheet(name, *, keep_unreachable=False)` | Remove a sheet | Cannot remove last sheet; also drops `calcChain.xml`, the sheet's own defined names and the parts only that sheet reached (drawings, charts, tables) unless `keep_unreachable=True`; renumbers later `localSheetId`s and the workbook view's `activeTab`/`firstSheet` |
| `wb.replace(old, new, *, value_type=None)` | SAR on all sheets | Substring or whole-cell typed; sets `fullCalcOnLoad` when any cell changes |
| `wb.set_full_calc_on_load()` | Request Excel full recalc on open | Sets `calcPr/@fullCalcOnLoad`; `replace`, `cell.formula`, `insert_rows` and `insert_columns` already set it |
| `wb.find(value, *, sheet=None)` | First cell with exact ``value`` | All sheets if ``sheet`` omitted; skips formulas |
| `wb.findall(value, *, sheet=None)` | All matching cells as `list[Cell]` | Sheet order, then row-major; `[]` if none |
| `wb.replace_image(name, image_path)` | Replace picture by drawing name | |
| `wb.insert_image_at_placeholder(placeholder, image_path, *, width, height, max_width, max_height)` | Clear `{placeholder}` cell, insert image | Size kwargs (px) forwarded to `add_image`; width-only keeps aspect |
| `wb.write_dataframe(df, *, sheet, at_cell, header, mode, …)` | Bulk write rows | See [large-data-export.md](large-data-export.md) |
| `Workbook.export_to_template(template, df, save_path, **kwargs)` | Open template + `write_dataframe` + save | Class convenience |
| `wb.properties` | Core document metadata (`docProps/core.xml`) | `title`, `author`, `subject`, `keywords`, `comments`, `category`, `last_modified_by`, `created`, `modified` |
| `wb.orphan_partnames` | ZIP members not on relationship graph | Advanced |
| `wb.shared_strings` | Shared string table (advanced) | |
| `wb.styles` | Styles part (advanced) | `effective_xf_index(ws, address)` returns the `cellXfs` index that formats a cell |

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
| `ws.clear_range(cell_range)` | Clear values in range | Returns count cleared; refuses before clearing anything (see `FormulaGroupError`) |
| `ws.write_rows(rows, at_cell / at_row, …)` | Write rows at fixed positions | Bulk styling kwargs; a write it cannot complete raises before any cell changes ([details](large-data-export.md#two-write-modes)) |
| `ws.insert_rows(rows, at_cell, …)` | Insert rows, shift below down | See [Inserting rows and columns](#inserting-rows-and-columns) |
| `ws.insert_columns(cols, at_col / at_cell, …)` | Insert columns, shift right | ``cols`` = column vectors (top→bottom); see [Inserting rows and columns](#inserting-rows-and-columns) |
| `ws.update_dimension()` | Refresh `<dimension>` | Usually automatic |
| `ws.replace(old, new, *, value_type=None)` | SAR on this sheet only | Sets `fullCalcOnLoad` when any cell changes |
| `ws.find(value)` | First cell with exact ``value`` | Skips formulas; ``None`` if missing |
| `ws.findall(value)` | All matching cells as `list[Cell]` | `[]` if none |
| `ws.drawing_objects` | Every drawing object as `DrawingObject` | Pictures, charts, shapes, groups, slicers, …; see [Drawing objects](#drawing-objects) |
| `ws.images` | List of `Picture` | Read existing drawings |
| `ws.add_image(path, *, anchor, width, height, name, offset_x, offset_y)` | Insert image | Pixel offsets within anchor cell (default 0) |
| `ws.charts` | List of `Chart` | |
| `ws.add_chart(type, *, anchor, data_range, title, name, to_anchor)` | Add bar chart | `chart_type="bar"` today; flush from template; ``to_anchor`` optional end cell |
| `ws.tables` | List of [`Table`](#table) | The same object for a table on every call |
| `ws.add_table(cell_range, columns, *, name, display_name, write_header)` | Add Excel table | Header row must hold `columns`, or pass `write_header=True`; `name` defaults to `TableN`, N being the new table's id or the next free number; raises `TableError`, changing nothing, for invalid or duplicate names or columns and for overlapping tables, and `FormulaGroupError` when `write_header=True` would overwrite the anchor of a multi-cell array or data-table formula |
| `ws.conditional_formatting` | Read CF blocks | colorScale, dataBar, cellIs, … |
| `ws.add_conditional_formatting(cell_range, *, operator, formula, …)` | Add `cellIs` rule | With optional dxf colors |
| `ws.add_color_scale_formatting(cell_range)` | Add color-scale rule | |
| `ws.expand_conditional_formatting(end_row)` | Extend CF ranges after bulk insert | Optional with `write_dataframe` |
| `ws.data_validations` | List of `DataValidation` | `dataValidation` rules, then read-only x14 rules; see [Data validation](#data-validation) |
| `ws.add_data_validation(sqref, *, type, formula1, formula2, operator, …)` | Add a validation | Dropdown lists and number, date, time, text-length and custom rules; checks every argument before changing anything |
| `ws.remove_data_validation(dv)` | Remove a validation | Not an x14 rule |

### Inserting rows and columns

`insert_rows` and `insert_columns` move everything in the workbook that names the shifted cells:

- on the sheet: cells, merges, conditional formats and data validations with their formulas (and their x14 extensions and sparklines), hyperlinks, tables with their `autoFilter` and sort state, the sheet's `autoFilter` and its sort state, page breaks, sheet views, protected ranges, ignored errors, cell watches, scenarios, shared, array and data-table formula ranges and, for columns, the `<cols>` widths and row `spans`;
- elsewhere: defined names, reference by reference, the series, titles and labels of every chart, on a worksheet or a chartsheet (with the copies Excel keeps in chart extensions; a chartEx chart reads its data through hidden defined names, which move with the other names), and the references to the sheet in other sheets' conditional formats, data validations (with their x14 extensions and sparklines) and internal hyperlink locations.

A range that straddles the insert point grows; one that spans every row (column) ignores a row (column) insert. Excel evaluates a conditional format's or data validation's formula relative to the top-left cell of its ranges, and adjusts it for each cell, splitting the format where cells come to need different formulas; xlsxedit keeps one format, with the formula Excel gives its top-left cell. A format over every row (column) that a row (column) insert leaves in place keeps its relative references, and only the absolute ones move; Excel does the same for most such formulas, but moves some ranges that mix `$` and relative bounds as a whole. String literals, structured references and names in these formulas stay as they are. A chart keeps its cached values; Excel refreshes them from the cells. The inserts do **not** rewrite formulas in cells, drawings, comments or pivot tables. Columns inserted strictly inside a table join it: each becomes a `tableColumn` with the next free id, and the table's `filterColumn` ids and sort conditions move with their columns. `insert_columns` names a joining column after the value it writes into the column's header cell, which must be a non-empty string that no other column of the table uses (ignoring case) and must not lie inside a merge, which would carry the value to its top-left cell; where it writes nothing there, the column gets the first unused name `Column1`, `Column2`, …, and that name is written into the header cell, as Excel does. Inserting at a table's first column moves the whole table; inserting just past its last column leaves it alone. Rows inserted at a table's totals row join its data rows, so its `autoFilter` and sort range still end at the last data row.

Merges, formats, validations, hyperlinks and other ranges pushed past the last row or column are dropped; a single reference pushed off becomes `#REF!`. `template_rows` name rows as they are before the insert, and their one-row merges are copied onto the new rows. Both methods drop `calcChain.xml`, set `fullCalcOnLoad`, and return the number of rows (columns) inserted.

They refuse an insert, changing nothing, by raising:

- `GridOverflowError` when the new cells, cell content, a table, a scenario input cell or a data-table input cell would pass row 1,048,576 or column XFD;
- `InvalidRangeError` for a reference in the sheet they cannot parse, or an insert point off the grid (`TypeError` if it is not an integer);
- `TableError` when columns would be inserted strictly inside a table whose columns are bound to a query (`queryTableFieldId`), or when `insert_columns` would write into a joining column's header cell a value that cannot name the column, or one that a merge covers;
- `FormulaGroupError` when a value aimed at a cell that a merge covers would overwrite a merge anchor holding a multi-cell array or data-table formula;
- `ValueError` when a `template_rows` merge would overlap a merge, or when both `template_rows` and `row_styles` are given;
- `TypeError` or `ValueError` for a value or inline style they cannot write.

Values aimed at cells that a merge covers follow the rule in [large-data-export.md](large-data-export.md).

Both methods move whole rows (columns), as Excel's *Insert Sheet Rows* does. Excel's *Insert Cells* and *Delete Cells*, which shift only the cells in some columns (rows) and leave the rest of the sheet in place, exist as an internal step, `xlsxedit.row_shift.shift_cells`, for table row inserts to build on. No public method runs it yet. Its docstring describes what it moves and the cases it refuses, which follow Excel's own refusals.

---

## Cell

| API | Summary | Notes |
|-----|---------|-------|
| `cell.value` | Get/set typed value | str, int, float, bool, `date`, `datetime`, `None`; overwriting a formula drops `calcChain.xml` (Excel rebuilds it) |
| `cell.data_type` | Raw OOXML cell type | `"s"`, `"str"`, `"inlineStr"`, `"b"`, `"e"`, or `None` (numeric) |
| `cell.formula` | Get/set formula text | Stored and returned without the leading `=` (`"A2*2"`); the setter also accepts `"=A2*2"`. Setting drops the cached result and its type and sets `fullCalcOnLoad`. A shared-formula follower reads `None` and becomes a standalone formula when set; a single-cell array formula stays an array formula. Refused edits raise `FormulaGroupError`. Setting to `None` removes the formula and keeps its cached result as a value; on a shared-formula master the followers lose their formulas too, as with `cell.value` |
| `cell.formula_type` | Kind of formula (read-only) | `"normal"`, `"shared"`, `"array"`, `"dataTable"`, or `None` without a formula, including the non-anchor cells of an array or data-table range |
| `cell.clear()` | Clear cell content | Keeps style index; drops calcChain if cell had a formula |
| `cell.address` | Cell address | |
| `cell.worksheet` | Owning worksheet | |
| `cell.offset(cols=0, rows=0)` | Neighbor cell | e.g. `offset(cols=1, rows=1)` is one right and one down |
| `cell.style` | `CellStyle` read proxy | Effective style: the cell's own; an empty cell without one takes its row's (row formatted as a whole), else its column's; otherwise the default, as Excel shows it |
| `cell.apply_style(**kwargs)` | Change only the passed style properties | `bold`, `font_color`, `bg_color`, `font_size`, …; a colour is six or eight hex digits (`#` optional) or a `Color` with an `rgb` value; builds on the effective style and keeps the current font's other properties; `False` turns bold/italic/underline off; reuses an identical existing format |
| `cell.apply_date_format()` | Apply standard date format | Builds on the effective style |
| `cell.apply_number_format(code)` | Apply custom number format | e.g. `"$#,##0.00"`; builds on the effective style |
| `cell.hyperlink` | `Hyperlink` proxy | `.url`, `.location`, `.display` |
| `cell.replace(old, new)` | Substring replace in cell text | |
| `cell.has_formula` | Whether cell has formula | |
| `cell.is_string` | Shared-string cell | |

### CellStyle (read via `cell.style`)

`bold`, `italic`, `underline`, `font_size`, `font_name`, `font_color`, `bg_color`, `horizontal_align`, `vertical_align`, `num_format`, `is_date`, `is_percent`

`font_size` is an `int`, or a `float` for a fractional size such as `10.5`. `font_color` is a `xlsxedit.Color` with `rgb`, `theme`, `tint`, `indexed`, and `auto` fields, or `None` when the font sets no colour; it compares equal to its `rgb` string, so `font_color == "FFFF0000"` works. `bg_color` is the fill's `rgb` string or `None`.

---

## Picture

From `ws.images` or `ws.add_image(...)`.

| API | Summary |
|-----|---------|
| `pic.name` | Drawing name get/set |
| `pic.anchor` | Top-left anchor cell get/set; a two-cell anchor's bottom-right corner moves by the same number of cells |
| `pic.offset_x`, `pic.offset_y` | Inset within anchor cell (pixels; default 0 = flush) |
| `pic.width`, `pic.height` | Size in pixels get/set (positive) |
| `pic.media_path` | Path inside package, e.g. `xl/media/image1.jpeg` |
| `pic.replace(image_path)` | Swap image bytes |

`ws.images` lists the pictures that sit directly in top-level two-cell anchors, then those in one-cell anchors, then the rest in document order: pictures inside groups, inside the `mc:Choice` branch of `mc:AlternateContent`, or in absolute anchors. Pictures in an `mc:Fallback` branch are left out; `ws.drawing_objects` has them.

Setting `width` or `height` writes `xdr:spPr/a:xfrm/a:ext` and resizes the anchor. A one-cell or absolute anchor takes the size in `xdr:ext`. A two-cell anchor gets a new bottom-right corner (`xdr:to`), computed from the sheet's column widths and row heights for Excel's default 11-point body font; hidden columns and rows count as zero. Reading the size uses `a:xfrm/a:ext`, then `xdr:ext`, then the two-cell anchor box. Resizing also removes a bare `a:ext` that earlier releases wrote directly under `xdr:spPr`.

A picture inside a group cannot be moved, offset, or resized, and one with an absolute anchor cannot be moved or offset: those setters raise `ValueError`, as does a move that would put a corner outside the sheet.

Each picture `add_image` creates gets a `cNvPr/@id` unique within the drawing part and a fresh `a16:creationId`.

---

## Chart

From `ws.charts` or `ws.add_chart(...)`.

| API | Summary |
|-----|---------|
| `chart.name` | Drawing name get/set |
| `chart.anchor` | Top-left anchor cell get/set (preserves size span) |
| `chart.to_anchor` | Bottom-right corner cell get/set |
| `chart.offset_x`, `chart.offset_y` | Pixel inset within anchor cell (post-create tweak; add is flush) |
| `chart.title` | Chart title get/set; `None` removes the title text |
| `chart.partname` | Chart part path |
| `chart.set_series_formula(index, formula, worksheet=None)` | Update series range; with `worksheet`, rebuild the series cache from its cells |

`ws.charts` lists the charts that sit directly in top-level two-cell anchors, then the rest in document order: charts in absolute or one-cell anchors, inside groups, or in an `mc:Choice` branch. Other graphic frames, such as chartEx charts, slicers, and timelines, are not in `ws.charts`; `ws.drawing_objects` has them.

`chart.anchor` and the offsets follow the same rules as a picture's. Setting `chart.to_anchor` raises `ValueError` for a chart inside a group or one whose anchor has no `xdr:to`.

A rebuilt series cache has `c:ptCount` equal to the size of the range. A numeric cache (`c:numCache`) holds only numeric cells: blank, text, boolean, and error cells get no `c:pt`, so point indexes can skip. Its `c:formatCode` is the number format of the first numeric cell, and a point whose cell has another format carries that format. A string cache (`c:strCache`) leaves out blank cells.

Each chart `add_chart` creates gets a `cNvPr/@id` unique within the drawing part and a fresh `a16:creationId`.

---

## Drawing objects

From `ws.drawing_objects`: every object of the sheet's drawing in document order, read-only. Group members and both branches of `mc:AlternateContent` are included, around an anchor or inside it.

| API | Summary |
|-----|---------|
| `obj.kind` | `"picture"`, `"chart"`, `"chartEx"`, `"slicer"`, `"timeline"`, `"shape"`, `"connector"`, `"group"`, `"contentPart"`, or `"graphicFrame"` for any other graphic frame |
| `obj.name`, `obj.id` | `cNvPr` name and id |
| `obj.anchor` | Top-left cell of the enclosing anchor; `""` for an absolute anchor |
| `obj.anchor_type` | `"twoCellAnchor"`, `"oneCellAnchor"`, or `"absoluteAnchor"` |
| `obj.alternate` | `"choice"` or `"fallback"` inside `mc:AlternateContent`, else `None` |
| `obj.group` | The group object it belongs to, else `None` |
| `obj.element` | The object's XML element, such as `xdr:pic` |

---

## Table

From `ws.tables` or `ws.add_table(...)`. A `Table` reads and edits the table part in place, so a table you hold follows later edits such as `insert_rows`. Attributes the file leaves out read as their schema defaults: one header row, no totals row, `totals_row_shown` true. `insertRow`, `insertRowShift` and everything the API does not name are kept as they are.

Setters and `resize` check everything before they change anything. When they raise, the table part and the cells are as they were; when they would change nothing, the table part keeps its exact bytes.

| API | Summary |
|-----|---------|
| `table.name` | Table name (read-only) |
| `table.display_name` | The name formulas use; settable. Raises `TableError` for a name `add_table` would refuse, or while a formula, a defined name, a pivot cache or a data connection refers to the table by its current name: none of them is rewritten |
| `table.id` | Workbook-wide id (read-only) |
| `table.ref` | Range, e.g. `A1:B10` |
| `table.columns` | Column names, left to right |
| `table.column_objects` | A read-only `TableColumn` per column: `name`, `id`, `unique_name`, `calculated_column_formula`, `totals_row_function` (`"none"` when absent), `totals_row_formula`, `totals_row_label`, `data_dxf_id`, `query_table_field_id` |
| `table.header_row_count` | 1 with a header row, else 0; settable (see [Header and totals rows](#header-and-totals-rows)) |
| `table.totals_row_count` | 1 with a totals row, else 0; settable |
| `table.totals_row_shown` | Whether the totals row has been shown; settable, but not to `False` while the totals row is shown |
| `table.style` | `TableStyle(name, show_first_column, show_last_column, show_row_stripes, show_column_stripes)`, or `None` without `tableStyleInfo`; settable. The name must be a built-in table style, such as `"TableStyleMedium2"`, a custom table style in `styles.xml`, or `None` |
| `table.header_row_dxf_id`, `data_dxf_id`, `totals_row_dxf_id`, `header_row_border_dxf_id`, `table_border_dxf_id`, `totals_row_border_dxf_id` | Index of a `dxfs` entry in `styles.xml`, or `None`; settable to an existing entry, such as one `wb.styles.ensure_dxf(...)` returns |
| `table.header_row_cell_style`, `data_cell_style`, `totals_row_cell_style` | Cell style names (read-only) |
| `table.header_range`, `data_range`, `totals_range` | The header, data and totals rows as ranges, `None` when the table has none |
| `table.first_data_row`, `last_data_row` | First and last data row numbers |
| `table.resize(ref, *, columns=None)` | Change the table's range (see [Resizing a table](#resizing-a-table)) |

### Resizing a table

`resize` works like Excel's Resize Table: the range takes in whatever its cells hold, and no cell moves. With a totals row, the last row of `ref` is the totals row. The `autoFilter` covers `ref` without the totals row, and is removed from a table without a header row. The sort range follows the data rows and is clamped into them; a sort condition whose column leaves the table is dropped. A `ref` equal to the current range changes nothing.

`columns` names every column of `ref`, left to right, and is required when `ref` adds or drops columns. A column the table keeps keeps its `tableColumn` (id, formulas, totals function) and must keep its name. A new column's header cell must hold its name or be empty; an empty one gets the name written into it, the only cell `resize` writes.

`resize` raises `TableError`, changing nothing, when:

- the header row would move;
- `ref` does not overlap the current range, would leave no data row, or overlaps another table on the sheet;
- the columns change without `columns`, or `columns` does not fit `ref`, repeats a name (ignoring case) or renames a column the table keeps;
- the columns change and the table is bound to a query (`queryTableFieldId`), whose query table part would not know them;
- a new column's header cell holds anything but its name;
- a dropped column is named in a structured reference, in a formula that names the table or in one of the table's column formulas, or a table slicer filters on it.

It raises `InvalidRangeError` for a `ref` that is reversed or off the grid.

### Header and totals rows

Setting `header_row_count` or `totals_row_count` keeps `ref` and writes no cell:

- `header_row_count = 0` turns the header row into the first data row and removes the `autoFilter`, which Excel writes only with a header row. `header_row_count = 1` makes the first row the header row; its cells must already hold the column names, and no `autoFilter` is added. The sort range follows the data rows.
- `totals_row_count = 1` turns the last data row into the totals row and sets `totals_row_shown`; `totals_row_count = 0` turns the totals row into a data row. The `autoFilter` and the sort range follow. Write the totals row's labels and formulas yourself.

Both raise, changing nothing, `TypeError` for a value that is not an `int` (including `True`) and `TableError` for any other value than 0 or 1 or when the table would be left without a data row.

---

## Data validation

From `ws.data_validations` or `ws.add_data_validation(...)`. Each `DataValidation` is one rule and the cells it applies to.

| API | Summary |
|-----|---------|
| `dv.sqref` | The cells, as a `SheetRangeSet`; settable |
| `dv.type` | `none`, `whole`, `decimal`, `list`, `date`, `time`, `textLength` or `custom` |
| `dv.operator` | `between`, `notBetween`, `equal`, `notEqual`, `lessThan`, `lessThanOrEqual`, `greaterThan` or `greaterThanOrEqual` for the types that compare; `None` for `list`, `custom` and `none` |
| `dv.formula1`, `dv.formula2` | Formula text without a leading `=`, or `None`; a formula Excel wraps in `mc:AlternateContent` reads from its `mc:Fallback` |
| `dv.allow_blank`, `dv.show_input_message`, `dv.show_error_message` | `bool` |
| `dv.prompt_title`, `dv.prompt` | Input message, or `None` |
| `dv.error_style` | `stop`, `warning` or `information` |
| `dv.error_title`, `dv.error` | Error alert, or `None` |
| `dv.is_x14` | Whether the rule is an `x14:dataValidation` from the sheet's extension list |

An attribute the file leaves out reads as its schema default: `type` `none`, `operator` `between`, `error_style` `stop`, and `False` for the flags.

`str(dv.sqref)` is the range text, such as `"B2:B10 D2"`. Iterating it gives one `CellRange` per area, with 1-based `min_col`, `min_row`, `max_col` and `max_row`, and `dv.sqref.intersects("C5")` tells whether the rule covers a cell or range. Range sets compare by identity, so compare their text.

Excel 2010 and later store a validation whose formulas refer to another sheet in the worksheet's extension list. `ws.data_validations` lists those after the others, with `is_x14` true. They are read-only: setting their `sqref` or removing them raises `DataValidationError`.

### Adding a validation

```python
ws.add_data_validation("B2:B100", type="list", formula1='"Open,Closed"', prompt="Pick a status")
ws.add_data_validation("C2:C100", type="whole", operator="between", formula1=1, formula2=10)
```

- `sqref` is a range string, a `CellRange` or a `SheetRangeSet`. It is written as Excel writes it: in capitals, without `$`, `B:B` as `B1:B1048576`, `3:3` as `A3:XFD3`, and each area once.
- Every type but `none` needs `formula1`. `whole`, `decimal`, `date`, `time` and `textLength` compare the cell with `operator`, which defaults to `between`; `between` and `notBetween` also need `formula2`. Other types take no `operator` and no `formula2`.
- A formula is a `str`, written without surrounding whitespace or a leading `=`, or an `int` or `float`, written as a number. A `list` formula is a quoted list such as `'"Yes,No"'`, a range on the same sheet, or a defined name. A reference qualified with the sheet's own name, such as `Sheet1!$H$1:$H$5` on `Sheet1`, is written without the qualifier.
- `allow_blank`, `show_input_message` and `show_error_message` default to `True`, as in Excel's dialog. Attributes equal to their schema default are left out.
- Excel allows titles of 32 characters, messages of 255, a quoted list of 255 between its quotes, and formulas of 8,192, counted in UTF-16 code units.

A formula that refers to another sheet or workbook is refused: Excel writes such a validation in its x14 form, which xlsxedit does not write yet. Refer to a defined name that the workbook already holds instead.

`add_data_validation` raises, changing nothing:

- `InvalidRangeError` for a malformed or empty `sqref`, or when another validation on the sheet has a malformed range;
- `TypeError` when `sqref` is not a range string, `CellRange` or `SheetRangeSet`, `type`, `operator`, `error_style` or a message is not a `str`, a flag is not a `bool`, or a formula is not a `str`, `int` or `float`;
- `DataValidationError` for a `type`, `operator` or `error_style` the schema does not list, a missing or extra formula or `operator`, text that XML cannot store or that is longer than Excel allows, a number that is not finite, a formula that refers to another sheet or workbook, a `sqref` that overlaps another validation on the sheet, because Excel keeps one validation per cell, and a sheet that was removed from its workbook.

### Changing and removing a validation

`dv.sqref = "B2:B200"` rewrites the range by the same rules and raises the same errors. It may overlap the rule's own cells. It also raises `DataValidationError` for an x14 rule, one that is no longer on its sheet, or one whose sheet was removed.

`ws.remove_data_validation(dv)` removes the rule, and the sheet's `dataValidations` element with its last rule. It raises `DataValidationError`, changing nothing, for an x14 rule, one that is not on the sheet, or a removed sheet, and `TypeError` for anything but a `DataValidation`.

`insert_rows` and `insert_columns` move validations with their cells, as [Inserting rows and columns](#inserting-rows-and-columns) describes.

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
| `GridOverflowError` | `InvalidRangeError` | An insert or bulk write would put cells past row 1,048,576 or column XFD ([details](#inserting-rows-and-columns)) |
| `InvalidColorError` | `ValueError` | Invalid color value |
| `InvalidImageError` | `ValueError` | Unsupported or corrupt image data |
| `MissingPartError` | `RuntimeError` | A required package part is absent |
| `DataValidationError` | `ValueError` | A data-validation edit Excel would reject or cannot apply, or a change to a read-only x14 validation ([details](#data-validation)) |
| `TableError` | `ValueError` | A table edit would leave an invalid or inconsistent table: `add_table` names, columns, range, overlap or header cells; `Table.resize` and the `Table` setters; a `write_dataframe` resize that does not fit the table; `insert_columns` inside a table bound to a query, or with a header value that cannot name a column joining a table, or whose header cell a merge covers |
| `FormulaGroupError` | `ValueError` | An edit would split a formula group: setting `cell.formula` on a shared-formula master that other cells still derive from, or on any cell of a multi-cell array or data-table formula; or overwriting, clearing, or removing the formula of such a multi-cell formula's anchor (`cell.value`, `cell.clear()`, `cell.formula = None`, `clear_range`, bulk writes, `add_table(write_header=True)`). Raised before anything changes, so a refused bulk write or `clear_range` changes no cell |
| `DTDForbiddenError` | `ValueError` | An XML part xlsxedit reads declares a DTD (`<!DOCTYPE>`), which Excel never writes. Raised by `Workbook.open`, so never partway through an edit: the parts parsed only on first use (tables, drawings, charts, worksheets opened with `large=True`, and the pivot caches, data connections and slicer caches that table renames and resizes read) are checked without being parsed. Parts xlsxedit never parses, such as custom XML, may keep a DTD. Parsing never expands entities or fetches external resources |

---

## Fidelity-tested fixtures

Round-trip tests in `tests/test_fidelity.py` run on every workbook at the top of `fixtures/`, including Excel-authored ones with rich tables, data validation, conditional formatting, all three `editAs` modes of a drawing anchor, array and dynamic-array formulas, and a macro-enabled book with comments, a pivot table and a slicer. [`fixtures/MANIFEST.md`](../fixtures/MANIFEST.md) lists what each one contains.

---

## Not yet supported

| Area | Status |
|------|--------|
| Pivot tables | Preserved as opaque parts; not editable |
| Comments / threaded comments | Not exposed |
| VBA / macros | `.xlsm` opens/saves; VBA is opaque round-trip only |
| All chart types | Add: bar template; read: existing charts |
| All CF rule types | Read most; write `cellIs` and color scale |
| Data validation in x14 form | Read only; formulas that refer to another sheet cannot be written yet |
| Byte-identical ZIP | The ZIP container is rewritten; parts an edit does not change keep their exact bytes ([how](how-xlsxedit-works.md)) |
| Style authoring from scratch | Clone/apply xf; not full theme designer |

See roadmap in [how-xlsxedit-works.md](how-xlsxedit-works.md) §9.
