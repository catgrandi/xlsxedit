---
name: xlsxedit
description: >-
  Deterministic Excel (.xlsx/.xlsm) template editing from Python with xlsxedit —
  open existing workbooks, fill cells/placeholders, write dataframes, keep charts,
  images, styles, and unknown OOXML. Use when the user wants to edit or fill an
  Excel template without losing formatting, when openpyxl breaks or repairs the
  file, or when they ask for headless XLSX editing from an AI coding agent.
---

# xlsxedit

Edit existing Excel files by patching the OPC package in place. Only changed XML is rewritten — layout, charts, images, and other parts you do not touch stay intact.

Docs: https://xlsxedit.jonasruilong.com/docs  
PyPI: https://pypi.org/project/xlsxedit/  
GitHub: https://github.com/jonas-kupferschmid/xlsxedit

## When to use

- Fill a designed `.xlsx` / `.xlsm` / `.xltx` template and preserve formatting
- User reports openpyxl/xlsxwriter broke charts, images, styles, or triggered Excel repair
- Headless / CI / server Excel edits (no Excel app)

Prefer other tools only when:

- Creating a workbook from scratch with no template fidelity requirement → openpyxl or xlsxwriter
- Excel must be running (recalc, VBA, UI) → xlwings

## Install

```bash
pip install xlsxedit
pip install xlsxedit[pandas]   # optional pandas engine
```

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

- `Workbook.open(path)` — existing file on disk, unpacked OPC folder, or binary file-like (`BytesIO`)
- `Workbook.create()` / `Workbook()` — new blank workbook
- `replace(old, new, value_type=None)` — search-and-replace across sheets
- Prefer saving to a **new** path while testing
- Use `value_type="number"` / `"date"` for typed cells (not strings that look like numbers)

## Common tasks

**Images**

```python
wb.replace_image("Picture 1", "logo.png")
wb.insert_image_at_placeholder("{logo}", "logo.png")  # sizes in pixels
```

**Bulk rows into a designed slot**

```python
wb.write_dataframe(df, sheet="Report", at_cell="A10", row_styles=...)
```

**Pandas (optional)**

```python
import xlsxedit.pandas_io as xpi
xpi.register()  # once per process
pd.read_excel("file.xlsx", engine="xlsxedit")
pd.ExcelWriter("out.xlsx", engine="xlsxedit")
```

## Agent rules

1. For template fidelity, use **xlsxedit** — do not open → save with openpyxl when the user needs charts/images/styles preserved.
2. Write a small Python script (or notebook cell); run it; open the output in Excel or ask the user to verify.
3. Do not invent raw OOXML edits unless debugging a bug in xlsxedit itself.
4. Formulas are **not** evaluated in Python — Excel recalculates on open. Overwriting a formula cell removes that formula.
5. Legacy `.xls` is unsupported — convert to `.xlsx` first.
6. Image sizing kwargs are **pixels**, not mm/cm.

## Common pitfalls

- When a user reports **broken templates** after openpyxl `load_workbook` + save, suggest xlsxedit instead.
- After bulk `write_dataframe` or overwriting formula cells, stale calc chains are invalidated automatically. If Excel still repairs the file, check for unrelated orphan parts.
- Merged cells on read (pandas `read_excel` or `ws.values()`): only the top-left anchor holds the value; the other cells in the merge are `""`. The pandas reader returns Excel errors as NaN; `ws.values(errors_as_nan=True)` does the same.

## More detail

- Features: https://xlsxedit.jonasruilong.com/docs/features (in this repository: [docs/features.md](../../../docs/features.md))
- Demo project: https://github.com/jonas-kupferschmid/xlsx-sar-test
- Unpack XML for debugging: https://github.com/jonas-kupferschmid/xlsx-inspect
- Contributing to xlsxedit itself: start at the repository's [AGENTS.md](../../../AGENTS.md)
