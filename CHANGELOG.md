# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## Unreleased

Several changes below are breaking, so the next release is a major one.

### Breaking

- **Formulas** — `Cell.formula` returns formulas without the leading `=`: a cell set to `"=A1*2"` reads back `"A1*2"`, the form files written by Excel already returned. The setter still accepts a leading `=`. Setting an empty formula raises `ValueError`; a value that is neither `str` nor `None` raises `TypeError`.
- **Formulas** — setting `cell.formula` on a shared-formula master that other cells derive from, or on any cell of a multi-cell array or data-table formula, raises `FormulaGroupError`. So does overwriting, clearing or removing the formula of such a formula's anchor through `cell.value`, `cell.clear()`, `cell.formula = None`, `clear_range` or a bulk write.
- **Security** — opening a workbook whose XML parts declare a DTD (`<!DOCTYPE>`) raises `DTDForbiddenError`. Tables, drawings, charts, worksheets deferred by `large=True`, and the pivot caches, data connections and slicer caches that table renames and resizes read are checked on open too, without being parsed, so the error never comes partway through an edit.
- **Styles** — `cell.style.font_color` returns a `Color` (`rgb`, `theme`, `tint`, `indexed`, `auto`) instead of a `str`. It compares equal to its ARGB string; use `.rgb` where a string is needed. A theme colour no longer reads as the bare theme index.
- **Tables** — `add_table` requires the header row to hold the column names already, or `write_header=True` to write them. It raises `TableError` for a table name Excel rejects or that is already in use, empty or duplicate column names, a column list that does not fit the range, a range without a data row, and a range that overlaps another table.
- **Tables** — without `name=`, `add_table` names the table `TableN` after its workbook-wide id, or the next free number, instead of always `Table1`.
- **Bulk export** — `write_dataframe` resizes only the table the written rows overlap; rows written directly below a table no longer grow it unless `table=` names it. It raises `TableError`, before writing, when the rows overlap several tables, do not span exactly the table's columns, would put anything but the column names in the header row, when the table has a totals row, or when the resized table would overlap another, or would no longer overlap a table named by `table=` because the rows start above it or with a header row.
- **Bulk export** — `write_rows`, `write_dataframe` in overwrite mode and the pandas engine check the whole write before changing anything, and `clear_range` checks its range before clearing it. A refused write, with its `clear_range`, changes nothing; it used to keep the cells written or cleared before the refused one. A write past row 1,048,576 or column XFD raises `GridOverflowError` instead of writing cells Excel cannot open.
- **Inserts** — `insert_rows` and `insert_columns` move drawing anchors as Excel does, by each object's placement (`editAs`): an object that moves and sizes with cells moves, and grows when the new rows or columns go through it; one that only moves keeps its size; one that does neither stays where it is on the page. `write_dataframe(mode="insert")` and the pandas engine's insert mode do the same. An insert that would push an object off the grid raises `GridOverflowError`. Pass `shift_drawings=False` (in `engine_kwargs` for the pandas engine) to leave every anchor as it was. See [Drawings on inserts](docs/features.md#drawings-on-inserts).
- **Inserts** — `insert_rows` and `insert_columns` refuse an insert they cannot complete, changing nothing, including `insert_columns` inside a table when the value for a new column's header cell cannot name the column: see [Inserting rows and columns](docs/features.md#inserting-rows-and-columns) for each case.
- **Tables** — `Table.resize` raises `TableError`, changing nothing, when the header row would move, the new range does not overlap the current one, leaves no data row or overlaps another table, or adds or drops columns without the new `columns=` option or of a table bound to a query. See [Resizing a table](docs/features.md#resizing-a-table).
- **Tables** — `Table` lives in `xlsxedit.table` (still importable from `xlsxedit` and `xlsxedit.drawing`), and its constructor takes the table part and its worksheet.

### Added

- **Workbook** — `copy_worksheet(name, new_name)`, which changes nothing when the copy fails. `remove_worksheet(name, keep_unreachable=True)` keeps the parts only that sheet reached.
- **Exceptions** — `GridOverflowError`, `FormulaGroupError`, `DTDForbiddenError`, `TableError` and `DataValidationError`.
- **Formulas** — `Cell.formula_type`: `"normal"`, `"shared"`, `"array"`, `"dataTable"` or `None`.
- **Drawings** — `ws.drawing_objects` lists every object of a sheet's drawing as a `DrawingObject`: pictures, charts, shapes, connectors, groups and their members, slicers, timelines and other graphic frames.
- **Drawings** — `Picture.placement` and `Chart.placement` read and set how an object follows its cells: `"twoCell"` (move and size), `"oneCell"` (move only) or `"absolute"` (neither), stored as `editAs` on a two-cell anchor.
- **Tables** — `write_dataframe(..., table=...)` names the table to resize; `add_table(..., write_header=True)` writes the header cells.
- **Styles** — `Color`, and `Styles.effective_xf_index(worksheet, address)`.
- **Data validation** — `ws.data_validations` lists a sheet's validations as `DataValidation` objects, x14 ones read-only. `ws.add_data_validation(...)` and `ws.remove_data_validation(dv)` add and remove them, and `dv.sqref` is settable; a refused edit raises the new `DataValidationError`, `InvalidRangeError` or `TypeError` and changes nothing. `SheetRangeSet` and `CellRange`, the types of `dv.sqref` and its areas, are exported.
- **Tables** — `Table` reads and writes `display_name`, `header_row_count`, `totals_row_count`, `totals_row_shown`, `style` (a `TableStyle`) and the six table dxf ids, and reads `id`, the cell style names, `header_range`, `data_range`, `totals_range`, `first_data_row` and `last_data_row`. `table.column_objects` lists each column as a `TableColumn`. `Table.resize(ref, columns=[...])` adds and drops columns.
- **Inserts** — an internal step for Excel's *Insert Cells* and *Delete Cells*, `xlsxedit.row_shift.shift_cells`, which moves only the cells in a span of columns (rows). It refuses, changing nothing, where Excel refuses (a table it would move only part of, an array or data-table formula it would change only part of, content or a drawing object it would push off the sheet) and where Excel would unmerge cells or split a shared formula. It moves a drawing object only where Excel's *Insert Cells* does: when every column (row) the object covers is in the span and the object starts at or after the first new line. No public method runs it yet; table row inserts will.
- **Ranges** — `CellRange.shift_pieces` gives an area's cells after a shift restricted to a span as areas, split where the area crosses the span's edge.

### Changed

- **Inserts** — `insert_rows` and `insert_columns` plan the whole move before changing anything, and now also move shared, array and data-table formula ranges, x14 conditional formats, validations and sparklines, internal hyperlinks, `autoFilter` and sort state, page breaks, sheet views, protected ranges, ignored errors, cell watches, scenarios, `<cols>` and row `spans`, and rewrite defined names reference by reference. Both drop `calcChain.xml` and set `fullCalcOnLoad`.
- **Inserts** — `template_rows` name rows as they are before the insert. `insert_columns` inside a table adds table columns, named after the text it writes into their header cells, or `ColumnN`, which it then writes there.
- **Tables** — `ws.tables` returns the same `Table` for a table on every call, over the live table part, so a table you hold follows `insert_rows` and `insert_columns`. A table part no edit changes keeps its exact bytes, including after `resize` to the range it already has.
- **Performance** — new rows are placed through the row index instead of a scan of the sheet, so the rows `insert_rows` writes into the gap it opens, and single-cell writes to new rows in any order, cost the same on a large sheet as on a small one: inserting 5,000 rows into a 50,000-row sheet takes about a second instead of two minutes. `update_dimension`, which every insert runs, reads the cell addresses in one regular-expression pass and takes half as long.
- **Bulk writes** — a value aimed at a cell that a merge covers goes to the merge's top-left cell, unless it is `None` or the same write fills that cell.
- **Styles** — `apply_style` changes only the properties passed and builds on the cell's current style. Identical styles reuse existing `cellXfs`, font, fill and dxf entries. An empty cell without its own style takes its row's or column's style.
- **Drawings** — `ws.charts` lists only chart frames. Geometry setters that cannot apply, such as moving a grouped object, raise `ValueError`.
- **Drawings** — `add_image` writes the anchor Excel writes for an inserted picture, a `twoCellAnchor` placed `oneCell` with a computed bottom-right corner, instead of a `oneCellAnchor`, and raises `ValueError`, adding nothing, for a picture that would pass the edge of the sheet or an anchor off it. `insert_image_at_placeholder` places every image before it clears any cell, so such a picture replaces no placeholder. In `ws.images` an added picture now comes with the two-cell ones.
- **Drawings** — moving or offsetting a picture or chart placed `oneCell` or `absolute` keeps its size, computing its bottom-right corner again over the sheet's column widths and row heights, instead of keeping its span of cells.
- **Saving** — parts, `.rels` items and `[Content_Types].xml` that an edit does not change keep their exact bytes, so saving an unedited workbook changes no member. Rewritten parts use Excel's XML declaration. The source's `Default` content types are kept.
- **Saving** — cell, bulk-write, merge, `clear_range`, row-height, column-width, shared-string, table and data-validation edits mark the parts they change, so saving them no longer parses their source a second time, and a changed part releases its source bytes. A part an edit changed is rewritten even if a later edit restores it.
- **Bulk writes** — the cells of rows a bulk write appends join the cell index only when a single-cell access reaches their row, which cuts peak memory for writing 100,000 × 8 cells by about a fifth.
- **Search and replace** — `replace` sets `fullCalcOnLoad` when it changes a cell.
- **Performance** — repeated `ws[...]` access no longer rescans the sheet.
- **Ranges** — `SheetRangeSet.shift` with a `span` splits an area that crosses the span's edge, as Excel splits the range of a conditional format or validation, instead of leaving the area where it was.

### Fixed

- **Inserts** — `insert_rows` and `insert_columns` rewrite the formulas of conditional formats and data validations, x14 formulas that hold several references, the series, titles and labels of charts, and the references to the sheet in other sheets' formats and validations. A conditional format comparing `$B10` keeps testing its own row after rows are inserted above it.
- **Inserts** — `insert_rows` and `insert_columns` no longer raise `AttributeError` in a workbook that has a chartsheet.
- **Cells** — writing one row through both the cell API and a bulk API no longer saves duplicate `<row>` elements, and cells stay in column order. `ws["a6"]` reaches the same cell as `ws["A6"]`. A value `cell.value` or a bulk write cannot hold raises before the cell's formula is touched, so it no longer drops `calcChain.xml` or detaches shared-formula followers first, and text XML cannot hold no longer leaves an empty shared string behind. A string written over a cell whose shared-string index names no string replaces it, where it used to raise or, for a negative index, overwrite another string; `add_table(write_header=True)` writes such header cells too. A NumPy `float64` is written as a number instead of `np.float64(...)`, and a `count` on `sharedStrings.xml` that is not a number is rewritten instead of raising.
- **Large mode** — a string written to a workbook opened with `large=True` no longer points one entry past its text in `sharedStrings.xml`.
- **Formulas** — `Cell.formula` drops the stale cached value and cell type, sets `fullCalcOnLoad`, and turns a shared-formula follower into a standalone formula.
- **Tables** — `write_dataframe` no longer moves or resizes a table it did not write into. `add_table` and `copy_worksheet` give tables unique ids and names.
- **Workbook** — the charts of a sheet `copy_worksheet` copies, and the defined names scoped to it, refer to the copy, as in Excel, instead of to the source sheet, and its drawing objects get fresh `a16:creationId`s.
- **Tables** — `Table.resize`, `insert_rows` and `insert_columns` keep a table's totals row out of its `autoFilter`, and move the table's sort state; `resize` removes the `autoFilter` of a table without a header row.
- **Styles** — exports with `template_rows` and `column_styles` allocate one cell format per template row and styled column instead of one per cell. `font_size` reads fractional sizes, and `<b val="0"/>` no longer reads as bold. A `num_format`, `font_name` or alignment the stylesheet cannot hold raises `TypeError` or `ValueError` before anything changes, instead of leaving an incomplete `numFmt` or an invalid alignment in `styles.xml`. `apply_style` accepts the `Color` that `font_color` returns, and a colour that is not six or eight hex digits raises `InvalidColorError`.
- **Workbook** — `remove_worksheet` drops `calcChain.xml`, the sheet's own defined names and the parts only it reached, and renumbers later `localSheetId`s. New parts never take an orphan's name, and saving raises `ValueError` rather than write two members under one name.
- **Drawings** — editing a chart and then a picture on one sheet keeps both edits. `Picture.anchor` moves the whole anchor, and `Picture.width` / `height` resize the picture with its anchor. Added images and charts get ids unique within their drawing. Rebuilt chart caches hold only numbers in `c:numCache`. A refused chart title or series formula leaves the chart unchanged, and clearing a title no longer writes an empty `c:tx`.
- **Bulk export** — `write_rows` with rows that hold no values no longer raises or writes an inverted `<dimension>`.
- **Worksheet XML order** — new worksheet and stylesheet children are placed in schema order around `legacyDrawing` and `mc:AlternateContent`.

### Security

- **XML parsing** — entities are never expanded and nothing is fetched from the network, which closes an XXE hole on lxml < 5.

## [1.0.1] - 2026-07-27

### Added

- **Charts** — `to_anchor` on `add_chart` / `Chart` (optional end cell for chart box size).
- **Charts & images** — `offset_x` / `offset_y` on `Picture` and `Chart` (post-create placement tweaks).
- **Templates** — `parse_template_xml` strips indent from bundled templates so serialized chart/table parts stay compact even when template sources are pretty-printed.

### Fixed

- **Worksheet XML order** — table, conditional formatting, drawing, merge, and hyperlink inserts now follow ECMA-376 child sequence (`worksheet_order.py`); fixes Excel repair when those features share one sheet.
- **Charts** — `add_chart` preserves `from`/`to` span when anchoring below the template default (no inverted box); default anchor offsets are flush (`colOff`/`rowOff` zero).
- **Bulk export** — `write_dataframe(..., header=False)` keeps the table header row when resizing the table range.
- **Conditional formatting** — `expand_conditional_formatting` only grows multi-row ranges; single-cell rules (e.g. a KPI cell) are no longer stretched down the column.

### Changed

- README hero image and [xlsx-sar-test](https://github.com/jonas-kupferschmid/xlsx-sar-test) companion-repo description.

## [1.0.0] - 2026-07-19

First public release.

### Added

- **Workbook** — `open`, `create`, `save`, `sheetnames`, `worksheets`,
  `add_worksheet`, `rename_worksheet`, `remove_worksheet`, `properties`
  (core document metadata). Opens `.xlsx`, `.xlsm`, `.xltx` / `.xltm`; VBA and
  unknown parts round-trip untouched.
- **Cells** — typed `value` (str, int, float, bool, `date`, `datetime`),
  `formula`, `data_type`, `clear`, `offset`, `find` / `findall`, `iter_cells`,
  `iter_rows`.
- **Styles** — `cell.style` read proxy, `apply_style`, `apply_date_format`,
  `apply_number_format`.
- **Layout** — column/row dimensions, `merge_cells`, `unmerge_cells`,
  `merged_ranges`, `clear_range`, `write_rows`, `insert_rows`, `insert_columns`.
- **Images** — `add_image`, `images`, `Picture.replace`, `replace_image`,
  `insert_image_at_placeholder`.
- **Charts** — `add_chart`, `charts`, read/set `title`, `set_series_formula`.
- **Tables** — `add_table`, `tables`, `Table.resize`.
- **Conditional formatting** — read blocks and rules; `add_conditional_formatting`
  (`cellIs`), `add_color_scale_formatting`, `expand_conditional_formatting`.
- **Hyperlinks** — `cell.hyperlink` with `url`, `location`, `display`.
- **Search and replace** — workbook/sheet `replace` (substring or whole-cell
  typed: text, number, date).
- **Bulk export** — `write_dataframe`, `Workbook.export_to_template`, `write_rows`,
  row/column styles, and `Workbook.open(..., large=True)` for big files.
- **Pandas** — optional `engine="xlsxedit"` for `ExcelWriter` and `read_excel`
  via `xlsxedit.pandas_io.register()`.
- **Typing** — ships `py.typed` (PEP 561).
- **Exceptions** — `XlsxeditError` base with `WorksheetNotFoundError`,
  `DuplicateWorksheetError`, `InvalidRangeError`, `InvalidColorError`,
  `InvalidImageError`, and `MissingPartError` (each also subclasses the builtin
  it replaces).

[1.0.1]: https://github.com/jonas-kupferschmid/xlsxedit/compare/v1.0.0...v1.0.1
[1.0.0]: https://xlsxedit.jonasruilong.com/changelog
