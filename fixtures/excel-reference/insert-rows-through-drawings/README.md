# insert-rows-through-drawings

Three whole rows inserted through and above drawings of every placement.

## Steps in Excel

1. Open `before.xlsx`.
2. Select rows 7:9 on `Drawings` by their row headers.
3. Home > Insert > Insert Sheet Rows.
4. Save as `after.xlsx`.

## How this pair was made

`python scripts/excel_fixtures/build_reference.py insert-rows-through-drawings` built `before.xlsx` in Excel, opened it again and ran the steps through Excel's object model:

```vb
Worksheets("Drawings").Rows("7:9").Insert()
```

- Excel: 16.0.20430.20146 (64-bit)
- Platform: Windows 11 10.0.26200 (AMD64)
- Date: 2026-10-19

## What Excel did

- The inserted rows 7 to 9 copy the 30 pt height of row 6; the custom heights of rows 8 and 15 move to rows 11 and 18.
- A two-cell anchor without `editAs` that straddles the insert grows: `from` stays and `to` moves down three rows. The chart does the same.
- `editAs="oneCell"` and `editAs="absolute"` anchors that straddle the insert keep their size: `from` stays and Excel recomputes `to`, cell and offset, over the new row heights.
- Two-cell and one-cell anchors below the insert move down three rows and keep their offsets.
- `editAs="absolute"` below the insert stays where it was on the page: Excel rewrites `from` and `to`, cells and offsets, to describe the same position over the new rows (`from` row 13 + 25400 EMU becomes row 10 + 254000 EMU).
- The anchor above the insert does not change.
