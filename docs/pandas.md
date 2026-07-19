# Pandas integration

xlsxedit works with pandas in three ways:

1. **ExcelWriter engine** — `pd.ExcelWriter(..., engine="xlsxedit")`
2. **Excel reader engine** — `pd.read_excel(..., engine="xlsxedit")`
3. **Direct API** — `wb.write_dataframe(df, …)` (no engine registration)

Docs: [xlsxedit.jonasruilong.com/docs/pandas](https://xlsxedit.jonasruilong.com/docs/pandas)

## Install

```bash
pip install xlsxedit[pandas]
# or: pip install pandas
```

Core xlsxedit only requires `lxml`. Pandas is optional.

## Tutorial

```bash
python tutorial/pandas_tutorial.py
```

Creates `tutorial/output/pandas_tutorial.xlsx` — register, write, read, and a simple template fill.

## Quick start (copy-paste)

```python
import pandas as pd
import xlsxedit.pandas_io as xpi

xpi.register()  # once per process

df = pd.DataFrame({"Item": ["Widget", "Gadget"], "Qty": [2, 5]})

# Write
with pd.ExcelWriter("out.xlsx", engine="xlsxedit") as writer:
    df.to_excel(writer, sheet_name="Data", index=False)

# Read
got = pd.read_excel("out.xlsx", engine="xlsxedit")
```

## Register once

```python
import xlsxedit.pandas_io as xpi

xpi.register()
```

This registers the **writer** via pandas' public `register_writer`, and installs the
**reader** by patching `ExcelFile._engines` (pandas has no public `register_reader`
today). After that, `engine="xlsxedit"` works for both `to_excel` / `ExcelWriter`
and `read_excel`.

Today you call `xlsxedit.pandas_io.register()` once per process. A future pandas PR
may add a public `register_reader`; until then the reader is installed by patching
`ExcelFile._engines` (documented, supported plugin hook).

## ExcelWriter engine

```python
import pandas as pd
import xlsxedit.pandas_io as xpi

xpi.register()

df = pd.DataFrame({"Item": ["Widget"], "Qty": [2]})

# Create a new workbook
with pd.ExcelWriter("out.xlsx", engine="xlsxedit") as writer:
    df.to_excel(writer, sheet_name="Data", index=False)

# Fill an existing template (preserves other sheets / charts / styles)
with pd.ExcelWriter(
    "template.xlsx",
    engine="xlsxedit",
    mode="a",
    if_sheet_exists="overlay",
) as writer:
    df.to_excel(writer, sheet_name="Report", startrow=4, header=False, index=False)
```

### `if_sheet_exists` (append mode)

| Value | Behavior |
|-------|----------|
| `overlay` (recommended for templates) | Write at `startrow`/`startcol` only |
| `error` | Raise if the sheet already exists |
| `replace` | Clear / recreate the sheet, then write |
| `new` | Always create a uniquely named sheet |

### Writer `engine_kwargs`

Passed through to xlsxedit (not to pandas):

```python
engine_kwargs={
    "large": True,                 # lighter open for huge templates
    "mode": "overwrite",           # or "insert" (shifts rows down)
    "template_rows": [7, 8],       # copy template sample-row styles
    "clear_range": "A5:D1000",     # clear before overwrite
    "row_styles": [...],
    "column_styles": [...],
}
```

**Style policy (v1):** pandas per-cell fonts/fills from `to_excel` are ignored. Use template styles (`template_rows`) or `row_styles` / `column_styles` in `engine_kwargs`.

**File types:** the writer accepts `.xlsx`, `.xlsm`, and `.xltx` paths, so append mode can fill macro-enabled and template workbooks (VBA round-trips untouched).

**Dates / datetimes:** formats follow pandas `date_format` / `datetime_format` (defaults `YYYY-MM-DD` and `YYYY-MM-DD HH:MM:SS`), so values round-trip as dates rather than Excel serial numbers.

**MultiIndex:** column and index merges are written when `merge_cells=True` (pandas default). Pandas still requires `index=True` when writing MultiIndex **columns**.

`freeze_panes` and autofilter from pandas are not applied yet.

## Excel reader engine

```python
import pandas as pd
import xlsxedit.pandas_io as xpi

xpi.register()

df = pd.read_excel("book.xlsx", engine="xlsxedit")
df = pd.read_excel("book.xlsx", engine="xlsxedit", sheet_name="Report")
```

Or skip remembering the patch and use the helper (registers + forces the engine):

```python
from xlsxedit.pandas_io import read_excel

df = read_excel("book.xlsx")
```

Semantics match openpyxl's pandas reader defaults where it matters:

- Formula cells return the **cached calculated value** (`data_only` style), not `=A1+1`
- Excel errors in that cache (`#REF!`, `#N/A`, …) become **NaN** (openpyxl parity); `Cell.value` still returns the error string
- Merged cells: value only at the **top-left** anchor; other cells in the merge are blank
- Empty cells become `""` in the raw grid (pandas often shows them as `NaN` in the DataFrame)
- Sparse sheets keep row/column alignment (blank rows for gaps; pad to max width)

Reader `engine_kwargs`:

```python
pd.read_excel("huge.xlsx", engine="xlsxedit", engine_kwargs={"large": True})
```

Prefer **calamine** or **openpyxl** for general-purpose reads. Use the xlsxedit
reader when you already use xlsxedit / want one stack for write + read.

## Direct `write_dataframe`

Prefer this when you want full xlsxedit control without the ExcelWriter adapter:

```python
from xlsxedit import Workbook

wb = Workbook.open("template.xlsx")
wb.write_dataframe(
    df,
    sheet="Report",
    at_cell="A5",
    header=False,
    mode="overwrite",
    template_rows=[7, 8],
)
wb.save("out.xlsx")
```

See [large-data-export.md](large-data-export.md).

## Learn more

| Topic | Where |
|-------|-------|
| Beginner engine tutorial | `python tutorial/pandas_tutorial.py` |
| Bulk export / overwrite / insert / styles | `python tutorial/export_pandas_example.py` |
| Large-data details | [large-data-export.md](large-data-export.md) |

## When to use which

| Use case | API |
|----------|-----|
| Template fill from pandas | `ExcelWriter(..., engine="xlsxedit", mode="a", if_sheet_exists="overlay")` |
| Styles / insert / clear_range | `write_dataframe` or `engine_kwargs` |
| Create blank workbook from scratch | xlsxwriter or openpyxl may be simpler |
| Read with the same stack | `pd.read_excel(..., engine="xlsxedit")` or `xlsxedit.pandas_io.read_excel` |
| Fast / general Excel → DataFrame | calamine or openpyxl |
