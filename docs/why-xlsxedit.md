# Why xlsxedit?

You designed a report in Excel — logos, charts, conditional formatting, merged headers. Your Python script opens it, fills in the numbers, saves. Then Excel warns the file is corrupted, or the chart vanished, or the theme reset.

That usually happens because most libraries **rebuild** the workbook from a Python object model. Anything the model does not know about gets dropped on save.

**xlsxedit** takes a different path: load the `.xlsx` as an OPC package, patch only the XML nodes you change, and write the file back. Styles, drawings, charts, and unknown parts stay intact.

Docs: [xlsxedit.jonasruilong.com](https://xlsxedit.jonasruilong.com)

---

## Use cases

- Fill invoice or report **templates** designed in Excel
- **Search-and-replace** placeholders (`{client}`, `{date}`, `{logo}`)
- Refresh **data tables** without rebuilding the sheet layout
- **Bulk export** thousands of rows into a styled slot (`write_dataframe`)
- Preserve **images and charts** when updating numbers underneath

---

## How fidelity works

`xlsxedit` keeps the on-disk package as the source of truth. Recognized parts are edited surgically with `lxml`; everything else round-trips as opaque bytes on the relationship graph.

For architecture detail, see [How xlsxedit works](how-xlsxedit-works.md).

---

## Compared to other libraries

| | **xlsxedit** | **openpyxl** | **xlsxwriter** | **xlwings** |
|--|--------------|--------------|----------------|-------------|
| Format | `.xlsx` | `.xlsx` | `.xlsx` | `.xlsx` (via Excel) |
| Open existing file | Yes (direct file) | Yes | No | Yes (launches Excel) |
| Template fidelity | Designed for preservation | Often loses styles/charts/unknown XML | N/A (create only) | Depends on Excel |
| Headless / server / CI | Yes | Yes | Yes | **No** (desktop Excel) |
| Typical speed | Fast direct I/O; ~20k rows/s bulk export | Moderate | Fast (create) | **Slow** (Excel startup + COM) |
| Best for | Headless template fill, SAR, bulk export | General read/write | Fast new exports | Desktop Excel automation, macros |
| Excel app required | No | No | No | **Yes — must run Excel** |

### xlsxedit

Headless editing of `.xlsx` files on disk. No Excel process. Built for template fidelity and batch pipelines.

### openpyxl

Loads into an object model and rewrites the workbook. Fine for many tasks, but round-trips on complex templates often alter or drop formatting, charts, and OOXML the model does not implement.

### xlsxwriter

Excellent for creating new workbooks quickly. Cannot open an existing template.

### xlwings

Drives the **Excel application** (COM on Windows, AppleScript on Mac). Excel must be installed, launched, and kept in the loop — workable for desktop automation and macros, but **slow and impractical** for servers, CI, cron jobs, or bulk report generation.

---

## What works today

- Cells: typed values, formulas, `cell.clear()`
- Styles: read cell style; `apply_style`, date/number formats
- Layout: column/row dimensions, `merge_cells`
- Hyperlinks, images, charts, tables (add/replace)
- Conditional formatting (read; add `cellIs` rules)
- Search-and-replace: `replace`, `replace_image`, typed replace
- Bulk export: `write_dataframe`, `write_rows`, `insert_rows`, row/column styles
- Sheets: add, rename, remove

Fidelity is tested on real fixtures: merges, images, charts, tables, conditional formatting, hyperlinks, and more (`tests/test_fidelity.py`).

See [features.md](features.md) for the full API list.

---

## When to use xlsxedit

- Report pipelines, cron jobs, Docker / CI
- Filling styled templates **without** installing Excel
- You need charts, themes, or layout to survive after Python edits
- Search-and-replace or bulk row export into a designed workbook

## When not to use xlsxedit

- You need **live Excel recalc**, VBA, or controlling the Excel UI → use **xlwings**
- You are creating a workbook from scratch with no template → **xlsxwriter** or openpyxl may be simpler
- You need full authoring coverage (every chart type, pivot editing, etc.) → openpyxl’s broader API may fit better today; xlsxedit’s API is still growing

---

## Limitations

- Not byte-identical ZIP output — semantic preservation of unedited parts, not a diff tool
- Some conditional-formatting types are read-only
- Pivot tables and comments are not edited yet
- Unknown features are preserved as opaque parts, not exposed in the API

---

## Acknowledgments

The OPC package layout and surgical `lxml` editing pattern follow work by [Steve Canny (scanny)](https://github.com/scanny) on [python-docx](https://github.com/python-docx/python-docx) and [python-pptx](https://github.com/python-pptx/python-pptx). xlsxedit is independent and not affiliated with those projects.
