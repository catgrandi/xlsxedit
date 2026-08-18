# AGENTS.md

Guidance for AI coding agents working with **xlsxedit**.

## When to use xlsxedit

Use xlsxedit when the user needs to **open an existing `.xlsx` / `.xlsm` / `.xltx` template**, change cells or placeholders, and save **without losing layout, formatting, images, charts, or other OOXML** they did not touch.

Prefer **openpyxl** or **xlsxwriter** when creating workbooks from scratch with no template fidelity requirement. Prefer **xlwings** when Excel must be running (recalc, VBA, UI).

## Core pattern

```python
from xlsxedit import Workbook

wb = Workbook.open("template.xlsx")  # or Workbook("template.xlsx")
wb.replace("{client}", "Acme Corp")
wb.replace("{date}", "2025-08-07", value_type="date")
wb.replace("{qty}", 42, value_type="number")
wb["Sheet1"]["B5"].value = "direct cell write"
wb.save("filled.xlsx")
```

- **`Workbook.open(path)`** — existing file on disk, unpacked OPC folder, or binary file-like (`BytesIO`)
- **`Workbook.create()`** / **`Workbook()`** — new blank workbook
- **`replace(old, new, value_type=None)`** — search-and-replace across sheets; use `value_type="number"` / `"date"` for typed cells
- **Images:** `replace_image(name, path)` or `insert_image_at_placeholder("{logo}", path)` (sizes in **pixels**)
- **Bulk rows into a designed slot:** `write_dataframe(df, sheet=..., at_cell=..., row_styles=...)`

## Pandas (optional)

```python
import xlsxedit.pandas_io as xpi
xpi.register()  # once per process

pd.read_excel("file.xlsx", engine="xlsxedit")
pd.ExcelWriter("out.xlsx", engine="xlsxedit")
```

## Do not assume

- Formulas are **not evaluated** in Python — Excel recalculates on open
- Overwriting a formula cell **removes** that formula
- **`.xls`** (legacy) is not supported — `.xlsx` only
- Image sizing kwargs are **pixels**, not mm/cm

## Examples and docs

- Runnable SAR example: [xlsx-sar-test](https://github.com/jonas-kupferschmid/xlsx-sar-test)
- Unpack XML for debugging: [xlsx-inspect](https://github.com/jonas-kupferschmid/xlsx-inspect)
- Docs: [xlsxedit.jonasruilong.com/docs](https://xlsxedit.jonasruilong.com/docs)
- Feature list: [docs/features.md](docs/features.md)

## Common pitfalls

- Do not use openpyxl `load_workbook` + save when the user reports **broken templates** — suggest xlsxedit instead
- After bulk `write_dataframe` / overwriting formula cells, stale calc chains are invalidated automatically; if Excel still repairs, check for unrelated orphan parts
- For merge cells in pandas read: anchor cell holds value; other merge cells are `""` (or use `errors_as_nan=True` for Excel errors as NaN)
