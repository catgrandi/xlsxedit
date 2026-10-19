# insert-rows-above-table

Two whole rows inserted above a table, at the first row of a conditional format and two validations.

## Steps in Excel

1. Open `before.xlsx`.
2. Select rows 3:4 on `Stock` by their row headers.
3. Home > Insert > Insert Sheet Rows.
4. Save as `after.xlsx`.

## How this pair was made

`python scripts/excel_fixtures/build_reference.py insert-rows-above-table` built `before.xlsx` in Excel, opened it again and ran the steps through Excel's object model:

```vb
Worksheets("Stock").Rows("3:4").Insert()
```

- Excel: 16.0.20430.20146 (64-bit)
- Platform: Windows 11 10.0.26200 (AMD64)
- Date: 2026-10-19

## What Excel did

- Everything from row 3 down moves two rows: the table (A4:C9 to A6:C11) and its autoFilter, the merge A12:B12, the defined name `Reorder`, and the formula in E5 (now E7, `SUM(B5:B9)` rewritten to `SUM(B7:B11)`).
- A conditional format or validation whose first row is the insert row moves; it does not grow. F3:F10 becomes F5:F12 for the conditional format and the x14 list validation, and G3:G10 becomes G5:G12 for the base list validation.
- The expression rule's relative formula follows its range: `F3="Yes"` becomes `F5="Yes"`.
- Rows 1 and 2, the merged title and the note, stay where they are, and so does the defined name `Title`.
- The inserted rows take the format of row 2 above them, which is the default.
