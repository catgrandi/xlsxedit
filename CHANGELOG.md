# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## Unreleased

Several changes below are breaking, so the next release is a major one.

### Breaking

- **Formulas** — `Cell.formula` returns formulas without the leading `=`: a cell set to `"=A1*2"` reads back `"A1*2"`, the form files written by Excel already returned. The setter still accepts a leading `=`. Setting an empty formula raises `ValueError`; a value that is neither `str` nor `None` raises `TypeError`.
- **Formulas** — setting `cell.formula` on a shared-formula master that other cells derive from, or on any cell of a multi-cell array or data-table formula, raises `FormulaGroupError`. So does overwriting, clearing or removing the formula of such a formula's anchor through `cell.value`, `cell.clear()`, `cell.formula = None`, `clear_range` or a bulk write.
- **Security** — a workbook whose XML parts declare a DTD (`<!DOCTYPE>`) raises `DTDForbiddenError` when the part is first parsed.
- **Styles** — `cell.style.font_color` returns a `Color` (`rgb`, `theme`, `tint`, `indexed`, `auto`) instead of a `str`. It compares equal to its ARGB string; use `.rgb` where a string is needed. A theme colour no longer reads as the bare theme index.
- **Tables** — `add_table` requires the header row to hold the column names already, or `write_header=True` to write them. It raises `TableError` for a table name Excel rejects or that is already in use, empty or duplicate column names, a column list that does not fit the range, a range without a data row, and a range that overlaps another table.
- **Tables** — without `name=`, `add_table` names the table `TableN` after its workbook-wide id, or the next free number, instead of always `Table1`.
- **Bulk export** — `write_dataframe` resizes only the table the written rows overlap; rows written directly below a table no longer grow it unless `table=` names it. It raises `TableError`, before writing, when the rows overlap several tables, do not span exactly the table's columns, would put anything but the column names in the header row, or the resized table would overlap another.
- **Inserts** — `insert_rows` and `insert_columns` refuse an insert they cannot complete, changing nothing: see [Inserting rows and columns](docs/features.md#inserting-rows-and-columns) for each case.

### Added

- **Workbook** — `copy_worksheet(name, new_name)`. `remove_worksheet(name, keep_unreachable=True)` keeps the parts only that sheet reached.
- **Exceptions** — `GridOverflowError`, `FormulaGroupError`, `DTDForbiddenError` and `TableError`.
- **Formulas** — `Cell.formula_type`: `"normal"`, `"shared"`, `"array"`, `"dataTable"` or `None`.
- **Drawings** — `ws.drawing_objects` lists every object of a sheet's drawing as a `DrawingObject`: pictures, charts, shapes, connectors, groups and their members, slicers, timelines and other graphic frames.
- **Tables** — `write_dataframe(..., table=...)` names the table to resize; `add_table(..., write_header=True)` writes the header cells.
- **Styles** — `Color`, and `Styles.effective_xf_index(worksheet, address)`.

### Changed

- **Inserts** — `insert_rows` and `insert_columns` plan the whole move before changing anything, and now also move shared, array and data-table formula ranges, x14 conditional formats, validations and sparklines, internal hyperlinks, `autoFilter` and sort state, page breaks, sheet views, protected ranges, ignored errors, cell watches, scenarios, `<cols>` and row `spans`, and rewrite defined names reference by reference. Both drop `calcChain.xml` and set `fullCalcOnLoad`.
- **Inserts** — `template_rows` name rows as they are before the insert. `insert_columns` inside a table adds table columns.
- **Bulk writes** — a value aimed at a cell that a merge covers goes to the merge's top-left cell, unless it is `None` or the same write fills that cell.
- **Styles** — `apply_style` changes only the properties passed and builds on the cell's current style. Identical styles reuse existing `cellXfs`, font, fill and dxf entries. An empty cell without its own style takes its row's or column's style.
- **Drawings** — `ws.charts` lists only chart frames. Geometry setters that cannot apply, such as moving a grouped object, raise `ValueError`.
- **Saving** — parts, `.rels` items and `[Content_Types].xml` that an edit does not change keep their exact bytes, so saving an unedited workbook changes no member. Rewritten parts use Excel's XML declaration. The source's `Default` content types are kept.
- **Search and replace** — `replace` sets `fullCalcOnLoad` when it changes a cell.
- **Performance** — repeated `ws[...]` access no longer rescans the sheet.

### Fixed

- **Cells** — writing one row through both the cell API and a bulk API no longer saves duplicate `<row>` elements, and cells stay in column order. `ws["a6"]` reaches the same cell as `ws["A6"]`.
- **Formulas** — `Cell.formula` drops the stale cached value and cell type, sets `fullCalcOnLoad`, and turns a shared-formula follower into a standalone formula.
- **Tables** — `write_dataframe` no longer moves or resizes a table it did not write into. `add_table` and `copy_worksheet` give tables unique ids and names.
- **Styles** — exports with `template_rows` and `column_styles` allocate one cell format per template row and styled column instead of one per cell. `font_size` reads fractional sizes, and `<b val="0"/>` no longer reads as bold. A `num_format`, `font_name` or alignment the stylesheet cannot hold raises `TypeError` or `ValueError` before anything changes, instead of leaving an incomplete `numFmt` or an invalid alignment in `styles.xml`. `apply_style` accepts the `Color` that `font_color` returns, and a colour that is not six or eight hex digits raises `InvalidColorError`.
- **Workbook** — `remove_worksheet` drops `calcChain.xml`, the sheet's own defined names and the parts only it reached, and renumbers later `localSheetId`s. New parts never take an orphan's name, and saving raises `ValueError` rather than write two members under one name.
- **Drawings** — editing a chart and then a picture on one sheet keeps both edits. `Picture.anchor` moves the whole anchor, and `Picture.width` / `height` resize the picture with its anchor. Added images and charts get ids unique within their drawing. Rebuilt chart caches hold only numbers in `c:numCache`. A refused chart title or series formula leaves the chart unchanged, and clearing a title no longer writes an empty `c:tx`.
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
