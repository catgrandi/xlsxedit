# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
