# table-append-row

A row added to the bottom of a table by pressing Tab in its last cell.

## Steps in Excel

1. Open `before.xlsx`.
2. Select E6, the last cell of the `Orders` table.
3. Press Tab.
4. Save as `after.xlsx`.

## How this pair was made

`python scripts/excel_fixtures/build_reference.py table-append-row` built `before.xlsx` in Excel, opened it again and ran the steps through Excel's object model:

```vb
Worksheets("Orders").ListObjects("Orders").ListRows.Add()
```

- Excel: 16.0.20430.20146 (64-bit)
- Platform: Windows 11 10.0.26200 (AMD64)
- Date: 2026-10-19

## What Excel did

- `ListRows.Add()` stood in for Tab. With `AlwaysInsert` omitted it shifted the cells below the table down one row within the table's columns, although row 7 was empty: the note in A8 moved to A9, while G2 beside the table stayed. Whether Tab does the same when the row below is empty was not checked in the UI.
- The table and its autoFilter grow from A1:E6 to A1:E7. The new row is empty except D7, which gets the calculated-column formula.
- C7, D7 and E7 take their columns' number formats (the `dataDxfId` formats).
- The conditional format on the Amount column grows from D2:D6 to D2:D7, and the validation on the Qty column from B2:B6 to B2:B7.
