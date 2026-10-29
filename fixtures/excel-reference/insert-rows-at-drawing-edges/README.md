# insert-rows-at-drawing-edges

Two whole rows inserted where drawings of every placement start or end exactly on a gridline.

## Steps in Excel

1. Open `before.xlsx`.
2. Select rows 7:8 on `Edges` by their row headers.
3. Home > Insert > Insert Sheet Rows.
4. Save as `after.xlsx`.

## How this pair was made

`python scripts/excel_fixtures/build_reference.py insert-rows-at-drawing-edges` built `before.xlsx` in Excel, opened it again and ran the steps through Excel's object model:

```vb
Worksheets("Edges").Rows("7:8").Insert()
```

- Excel: 16.0.20527.20088 (64-bit)
- Platform: Windows 11 10.0.26200 (AMD64)
- Date: 2026-10-29

## What Excel did

- A drawing whose bottom edge lies on the gridline at the top of row 7, where the rows go in, does not change, whatever its placement: its `to` stays at row 6 (zero-based) with offset 0.
- Two-cell and one-cell drawings whose top edge lies on that gridline move down two rows, both corners.
- The two-cell drawing across the insert grows by two rows. The one-cell one keeps its size, and Excel writes its recomputed `to`, which falls on a gridline, as that row with offset 0 (row 10), not as the row above with a full offset.
- Absolute drawings stay where they were on the page; every row has the default height, so their anchor cells do not change.
