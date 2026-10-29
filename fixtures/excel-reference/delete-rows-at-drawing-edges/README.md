# delete-rows-at-drawing-edges

Three whole rows deleted where drawings of every placement start or end exactly on a gridline.

## Steps in Excel

1. Open `before.xlsx`.
2. Select rows 5:7 on `Edges` by their row headers.
3. Home > Delete > Delete Sheet Rows.
4. Save as `after.xlsx`.

## How this pair was made

`python scripts/excel_fixtures/build_reference.py delete-rows-at-drawing-edges` built `before.xlsx` in Excel, opened it again and ran the steps through Excel's object model:

```vb
Worksheets("Edges").Rows("5:7").Delete()
```

- Excel: 16.0.20527.20088 (64-bit)
- Platform: Windows 11 10.0.26200 (AMD64)
- Date: 2026-10-29

## What Excel did

- A drawing whose bottom edge lies on the gridline at the top of row 5, the first deleted row, does not change, whatever its placement.
- The two-cell drawing that covers exactly the deleted rows, from the top of row 5 to the top of row 8, is deleted. The one-cell drawing there survives at the top of row 5 with its size kept, and the absolute one stays where it was.
- A two-cell drawing whose bottom edge lies on a gridline inside the deleted rows (the top of row 7) now ends at the top of row 5. One whose top edge lies inside them (the top of row 6) now starts at the top of row 5, and its bottom moves up three rows.
- One-cell drawings that start inside the deleted rows start at the top of row 5, those that start below them move up three rows, and both keep their size; the one that ends inside them keeps its size and place.
- Absolute drawings stay where they were on the page; every row has the default height, so their anchor cells do not change.
