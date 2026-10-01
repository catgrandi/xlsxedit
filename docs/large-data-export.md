# Large data export

Bulk APIs for exporting many rows with **xlsxedit**. Normal `ws["A1"]` surgical edits are unchanged.

Docs: [xlsxedit.jonasruilong.com/docs/large-data-export](https://xlsxedit.jonasruilong.com/docs/large-data-export)

**Tutorials:**

- `python tutorial/pandas_tutorial.py` → beginner pandas engine (write + read)
- `python tutorial/export_pandas_example.py` → `tutorial/output/export_example.xlsx` (three sheets: overwrite, insert, style rows)
- `python tutorial/bench_large_export.py` → zebra export + timing (default 50k rows)

## Template layout (header + data slot + footer)

The worked example puts the footer **directly under the header** so overwrite vs insert is easy to see in Excel:

```
Row  1-3   Title, client, date          ← untouched
Row  4     Column headers               ← untouched
Row  5     Footer                       ← overwritten on first data row (overwrite)
                                      ← shifts down on insert (e.g. row 8 for 3 rows)
Rows 5-7   Data from DataFrame          ← written at at_cell="A5" in the tutorial
```

## Two write modes

| Mode | API | Behavior |
|------|-----|----------|
| **overwrite** (default) | `write_rows` / `write_dataframe(mode="overwrite")` | Writes at fixed rows; overwrites cells in that range |
| **insert** | `insert_rows` / `write_dataframe(mode="insert")` | Inserts rows at `at_cell`; shifts footer and content below down |

Insert is **not** faster than overwrite — choose insert when rows below `at_cell` must move down. Insert shifts merges, conditional-formatting ranges, and tables; it does **not** rewrite formula expressions (e.g. a formula that moves to `B6` may still say `=A5*2`) or move chart/drawing anchors.

```python
# Overwrite — replaces cells at A5+ (including footer if it sits on A5)
wb.write_dataframe(df, at_cell="A5", header=False, mode="overwrite")

# Re-export into the same slot with fewer rows — clear leftovers first
wb.write_dataframe(
    df, at_cell="A5", header=False, mode="overwrite", clear_range="A5:D100",
)

# Insert — footer on row 5 moves down by len(df)
wb.write_dataframe(df, at_cell="A5", header=False, mode="insert")
```

`write_dataframe` only overwrites cells it writes. It does **not** clear extra cells outside the new data rectangle. Use `clear_range` when re-filling a slot with **fewer** rows than a previous export.

## Tables in the slot

In overwrite mode, `write_dataframe` resizes the Excel table that the written rows overlap so that it ends at the last written row (`resize_table=True`, the default). The table keeps its columns and, when you write only the body, its header row. Rows that overlap no table leave every table as it is, so to grow a table by writing directly below it, name it with `table`:

```python
# Table "Sales" spans A4:D5: the header row and one sample row.
wb.write_dataframe(df, at_cell="A5", header=False)  # 3 rows: Sales now spans A4:D7
wb.write_dataframe(more, at_cell="A8", header=False, table="Sales")  # appends below it
```

`table` takes a table name or a `Table` from `ws.tables`. Before writing anything, `write_dataframe` raises `xlsxedit.exceptions.TableError` instead of resizing when:

- The rows overlap more than one table, or the resized table would overlap another one.
- The rows do not span exactly the table's columns.
- The table's header row would not hold its column names: a written header (`header=True`) differs from them, or body rows start at or above the header row without repeating them.

Pass `resize_table=False` to write such rows and leave every table as it is.

Pandas is optional — only `lxml` is required. Use any object with `.columns` and `.itertuples()` / `.values`.

To use xlsxedit as a pandas `ExcelWriter` / `read_excel` engine (`engine="xlsxedit"`), start with `tutorial/pandas_tutorial.py` or see [pandas.md](pandas.md).

## Three ways to style data

### A. Template sample rows (copy from Excel)

Put formatted sample cells on template rows (e.g. 7–8 for zebra). Bulk write copies the `s` style index per column:

```python
wb.write_dataframe(df, at_cell="A7", header=False, template_rows=[7, 8])
# single template row: template_rows=7
```

### B. Inline row styles (zebra in Python)

`row_styles` is a **list of dicts** — one per stripe, cycling row by row. Each dict can be `{}` (no extra style).

```python
row_styles=[
    {"bg_color": "FFFFFFFF"},
    {"bg_color": "FF9DC3E6"},
]
```

### C. Inline column styles (formats in Python)

`column_styles` is a **list of dicts** — one per column (A, B, C, … from `at_cell`).

```python
column_styles=[
    {},
    {"num_format": "0"},
    {"num_format": "$#,##0.00"},
    {"num_format": "$#,##0.00"},
]
```

**B + C together** (most common for code-only styling):

```python
wb.write_dataframe(
    df,
    at_cell="A5",
    header=False,
    row_styles=[{"bg_color": "FFFFFFFF"}, {"bg_color": "FF9DC3E6"}],
    column_styles=[{}, {"num_format": "0"}, {"num_format": "$#,##0.00"}, {"num_format": "$#,##0.00"}],
)
```

Keys match `apply_style` (`bold`, `bg_color`, `font_color`, …) plus `num_format` for number formats.

Do not mix `template_rows` with `row_styles` (raises `ValueError`). `column_styles` may combine with either.

## Parameters

| Parameter | Purpose |
|-----------|---------|
| `at_cell="A5"` | First row (and column) written |
| `header=False` | Skip writing DataFrame column names (use template row 4) |
| `clear_range="A5:D100"` | Overwrite mode only — clear old values before write |
| `template_rows=[7, 8]` | Zebra from template rows (`int` or `list[int]`) |
| `row_styles` / `column_styles` | Inline styles (lists of dicts; allocated once, cached) |
| `mode="overwrite"` \| `"insert"` | Overwrite slot vs push rows down |
| `resize_table=False` | Overwrite mode only — leave Excel tables as they are (see [Tables in the slot](#tables-in-the-slot)) |
| `table="Sales"` | Resize this table (name or `Table`) instead of the one the rows overlap |

## Anti-patterns

- Do not loop `ws[f"A{i}"]` for thousands of rows.
- Do not call `apply_style()` per cell on large exports.
- Do not use `ws.cells` on large sheets.

## Performance

Rough expectations for numeric data (post bulk-row-append fix):

| Scenario | Rough expectation |
|----------|-------------------|
| 50k–200k rows, 4 cols, overwrite | ~20k rows/s write; scales roughly linearly |
| Zebra + column formats | ~1.3× slower than plain numeric |
| Insert mode | Similar write speed; shifts merges/CF/tables; does **not** rewrite formula text or chart/drawing anchors |

Measure on your machine: `python tutorial/bench_large_export.py` (set `ROWS=200000` to stress-test).

Reference run: 200k rows × 4 cols, zebra `row_styles`, overwrite — ~20k rows/s for `write_dataframe` on Apple M2 (macOS). Total wall time ~12 s including pandas build and save; your hardware and column count will vary.

## `large=True` on open

Only when opening a huge existing file for light edits:

```python
wb = Workbook.open("huge.xlsx", large=True)
```

Export from a small template does not need this.
