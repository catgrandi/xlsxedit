# delete-rows-under-cf

Three whole rows deleted under conditional formats, validations, merges, defined names and drawings.

## Steps in Excel

1. Open `before.xlsx`.
2. Select rows 5:7 on `Sheet1` by their row headers.
3. Home > Delete > Delete Sheet Rows.
4. Save as `after.xlsx`.

## How this pair was made

`python scripts/excel_fixtures/build_reference.py delete-rows-under-cf` built `before.xlsx` in Excel, opened it again and ran the steps through Excel's object model:

```vb
Worksheets("Sheet1").Rows("5:7").Delete()
```

- Excel: 16.0.20430.20146 (64-bit)
- Platform: Windows 11 10.0.26200 (AMD64)
- Date: 2026-10-19

## What Excel did

- Conditional formats: a range that contains the deleted rows shrinks (A2:A20 to A2:A17); one that starts at the first deleted row keeps its start (B5:B12 to B5:B9); one that ends inside them is cut (C3:C6 to C3:C4); one entirely inside them is removed (D5:D7); one below moves up (E10:E15 to E7:E12); the x14 data bar shrinks (G2:G10 to G2:G7).
- A multi-area range whose areas become adjacent is merged: `F2:F4 F6:F9` becomes `F2:F6`.
- Validations follow the same rules, base and x14: H5:H7 is removed, H9:H12 becomes H6:H9, and the x14 list I2:I10 becomes I2:I7. Excel had already stored H5:H7 and H9:H12, added as two identical validations, as one with sqref `H5:H7 H9:H12`.
- Merges: J5:K6, inside the deleted rows, is removed; J9:K10 becomes J6:K7.
- Defined names: `Block` shrinks from A3:B9 to A3:B6; `Watch` named A6, a deleted cell, and becomes `Sheet1!#REF!`.
- Drawings inside the deleted rows: the two-cell (move and size) shape is deleted; the `editAs="oneCell"` shape survives at the top of row 5 with its size kept; the `editAs="absolute"` shape survives where it was.
- Drawings below the deleted rows: two-cell and one-cell shapes move up three rows; the absolute one stays where it was on the page, which here leaves its anchor cells unchanged because the rows are all the same height. The two-cell shape across the deleted rows shrinks by three rows.
