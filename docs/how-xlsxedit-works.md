# How `xlsxedit` works

Choosing a library? See [Why xlsxedit?](why-xlsxedit.md) first. API catalog: [features.md](features.md).

`xlsxedit` is a surgical `.xlsx` editor. It does **not** rebuild the workbook from a high-level model (the failure mode that makes tools like openpyxl drop formatting). It loads the OPC ZIP into parts, mutates only the XML it understands, and writes the package back so unrelated parts survive.

If you need Excel’s on-disk layout (sheets, shared strings, styles), read [excel-xlsx-structure.md](excel-xlsx-structure.md) first.

---

## 1. Design goal

| Goal | How |
|------|-----|
| Change cell text | Edit shared-string / inline-string XML |
| Keep cell look | Never strip `s` (style index) on `<c>` |
| Keep unknown Excel features | Leave unrecognized parts as opaque byte blobs |
| Simple workflow | `open` → mutate → `save` |

Public surface today:

```python
from xlsxedit import Workbook

wb = Workbook.create()
wb = Workbook.open("report.xlsx")
ws = wb["Sheet1"]
ws["A1"].value = "hello"
ws["A2"].value = 42
ws["A3"].apply_date_format()
ws.column_dimensions["C"].width = 20.0
ws.merge_cells("A1:C1")
ws["B1"].hyperlink.url = "https://xlsxedit.jonasruilong.com"
ws.add_image("logo.jpg", anchor="E2")
ws.add_chart("bar", anchor="G2", data_range="A1:B5", title="Sales")
ws.add_table("A1:C10", ["Item", "Price"])
ws.add_conditional_formatting("A2:A10", operator="greaterThan", formula="0")
wb.add_worksheet("Report")
wb.replace("old", "new")          # SAR convenience — calls base cell APIs
wb.replace("{n}", 888, value_type="number")
wb.save("report_out.xlsx")
```

---

## 1b. Creating documents (template model)

`Workbook.create()` does not generate OPC XML in Python. It opens a bundled minimal package and mutates it — the same load/save path as editing an existing file.

| Asset | Role |
|-------|------|
| `templates/default.xlsx` | Runtime: `Workbook.create()` opens this |
| `templates/default-xlsx-template/` | Maintainer source (unpack/edit/repack) |
| `scripts/repack_default.py` | Pack folder → `default.xlsx` |

The template is a valid minimal workbook: one empty sheet, styles, theme, empty shared-string table. See section 9 for feature status. Base APIs are implemented; SAR wraps them.

---

## 2. Architecture

```mermaid
flowchart LR
  xlsx[".xlsx ZIP"] --> reader[PackageReader]
  reader --> relwalk[Relationship walk]
  relwalk --> factory[PartFactory]
  factory --> opaque[Part opaque blob]
  factory --> xml[XmlPart live lxml]
  xml --> wb[WorkbookPart]
  xml --> ws[WorksheetPart]
  xml --> sst[SharedStringsPart]
  api[Workbook API] --> xml
  api --> writer[PackageWriter]
  opaque --> writer
  writer --> out[".xlsx ZIP"]
```

| Layer | Role |
|-------|------|
| `Workbook` / `Worksheet` / `Cell` | Public API |
| `OpcPackage` | In-memory package + relationships |
| `Part` / `XmlPart` / `PartFactory` | Blob vs parsed XML |
| `PackageReader` / `PackageWriter` | ZIP ↔ parts |

### Load path

1. Open the ZIP (or an unpacked directory).
2. Read `[Content_Types].xml` and `/_rels/.rels`. A part neither lists falls back to `application/xml` (`.xml`) or `application/octet-stream`.
3. **Walk the relationship graph** from the package root (same idea as Word’s OPC model): follow each internal relationship, load that part, then follow *its* relationships, and so on.
4. For each part, `PartFactory` picks a class by content type:
   - workbook / worksheet / sharedStrings → `XmlPart` subclass (parsed to lxml)
   - **everything else** → base `Part` (keeps original bytes)
5. Wire relationship objects so `rId`s point at real part instances.
6. `Workbook` wraps the workbook part, builds sheet proxies, and attaches the shared string table if present.

### Save path

1. Collect all loaded parts. Two parts under one name (part names compare case-insensitively) raise `ValueError` before anything is written.
2. Write `[Content_Types].xml` and each part’s `.rels`: their source bytes while those still describe the in-memory graph exactly, otherwise regenerated from it. A regenerated `[Content_Types].xml` keeps the source `Default` entries and adds an `Override` only for a part whose content type differs from its extension’s `Default`.
3. Write every part’s `blob`:
   - opaque `Part` → original bytes unchanged
   - `XmlPart` → original bytes while the live lxml tree is unchanged (never handed out, or serializing exactly as a fresh parse of the original bytes does); otherwise the tree, serialized with Excel’s XML declaration
4. Produce a new ZIP.

---

## 3. What the library understands today

Only these content types are registered as live XML:

| Content type | Class | Edited for `replace`? |
|--------------|-------|------------------------|
| workbook (`.xlsx` / `.xlsm` / `.xltx` / `.xltm` variants) | `WorkbookPart` | No (read for sheet list) |
| worksheet | `WorksheetPart` | Read cells; inlineStr may be edited |
| sharedStrings | `SharedStringsPart` | Yes — `<t>` text mutated in place |
| styles | `StylesPart` | Live element; serves original bytes until a style is mutated |
| core properties (`docProps/core.xml`) | `CorePropertiesPart` | Via `wb.properties` |

Everything else (theme, drawings, charts, pivot caches, VBA, slicers, customXml, printer settings, …) falls through to opaque `Part`.

So for a text-only replace:

- Every part the replace does not change round-trips **byte-identical**: theme, styles, drawings, untouched sheets, `.rels` items, and `[Content_Types].xml`.
- Shared strings XML is re-serialized with your text changes, and so is a sheet whose inline strings changed (attribute order / insignificant whitespace can change; structure and `s` on cells stay unless you edit them). The workbook part is re-serialized too, because a replace that changes a cell sets `calcPr/@fullCalcOnLoad`.

---

## 4. How `replace` works

1. Walk each sheet’s `<c>` cells in `sheetData`.
2. Skip formula cells (`<f>` present) and non-string types (numbers, bools, errors).
3. For `t="s"`: read SST index from `<v>`, edit that `<si>` **once** per index (workbook-wide; Sheet1 and Sheet2 share the table).
4. For `t="inlineStr"`: edit `<t>` nodes under `<is>`.
5. Rich text: replace only when `old` sits entirely inside one `<t>` (run-local). Cross-run matches are skipped so `<rPr>` next door stays intact.
6. Do not rebuild or dedupe the SST; indices on cells stay valid.

That is why formatting survives: style indexes and property XML are left alone.

---

## 5. Unknown / new Excel features — are they kept?

**Short answer:** Yes for normal Office parts that are linked into the package through relationships — even if this library has never heard of them. They are loaded as opaque blobs and written back unchanged.

### What is preserved

| Situation | Behavior |
|-----------|----------|
| Chart, drawing, image, pivot cache, VBA, slicer, timeline, custom XML part, future Microsoft part | Loaded via the relationship graph → opaque `Part` → **same bytes on save** |
| New content type you have never registered | Same — default is opaque `Part` |
| Relationships pointing at those parts | Re-emitted on save from the in-memory graph |
| Cell style indexes (`s`), column widths, merged-region XML inside a sheet | Left alone by `replace` (the sheet is re-serialized only if a cell in it changed) |

**Politeness:** do not interpret what you do not understand; do not drop it from the package.

Example flow for a chart:

```text
workbook.xml.rels → worksheets/sheet1.xml
sheet1.xml.rels     → drawings/drawing1.xml
drawing1.xml.rels   → charts/chart1.xml  (+ maybe media/image1.png)

None of those content types are registered → each is Part(blob=original)
save → all those blobs written again
```

You never parse the chart XML, so you cannot corrupt it by misunderstanding it.

### Important caveats (honest limits)

1. **Relationship-unreachable ZIP members are “orphans.”**  
   OPC is a graph, not “every ZIP member.” Related unknown features (charts, VBA, …) are always kept. Members with **no** relationship chain are still **discovered** on open as `Workbook.orphan_partnames`, but they are only written on save when you pass `include_orphans=True`. Default `save()` stays OPC-strict and omits them. An orphan whose name is not a legal OPC part name, such as `[trash]/0000.dat`, is listed and written percent-encoded (`/%5Btrash%5D/0000.dat`). An orphan is skipped when its name cannot be made legal (an empty segment, or a segment ending in `.`) or would equal, case-insensitively, the name of a related part or of an orphan already kept. Numbered parts the library adds (sheets, drawings, charts, tables, media) never take an orphan’s name.

```python
wb = Workbook.open("report.xlsx")
if wb.orphan_partnames:
    print("unrelated ZIP cargo:", wb.orphan_partnames)
wb.save("out.xlsx", include_orphans=True)  # keep that cargo
```

2. **Removing a worksheet removes what only it reached.**  
   `remove_worksheet` drops the parts no other relationship chain reaches once the sheet is gone (its drawings, charts, chart styles, tables, …) instead of leaving them as unreachable cargo; parts the workbook or another sheet still reaches, such as a shared image, stay. Pass `keep_unreachable=True` to keep them in the package. It also drops `calcChain.xml`, as `copy_worksheet` does; Excel rebuilds it on open. `docProps/app.xml` is left as it was.

3. **XML parts we *do* understand are re-serialized when they change.**  
   An edited workbook, worksheet, or sharedStrings part is written from its lxml tree, so whitespace, attribute order, or namespace-declaration placement may differ from the source even outside the nodes you edited. Semantic content for unedited nodes remains. A part whose tree did not change keeps its exact bytes.

4. **Package bookkeeping is regenerated when the graph changes.**  
   `[Content_Types].xml` and each `.rels` item keep their source bytes while they still describe the package; once a part or relationship is added or removed they are rebuilt from the loaded graph. Targets and types are preserved; exact original XML formatting is not. A rebuilt `[Content_Types].xml` keeps the source `Default` entries, adds an `Override` only for a part whose type differs from its extension’s `Default`, and drops entries for parts no longer in the package. Orphans included via `include_orphans=True` get content-type entries when written.

5. **We do not implement every Excel behavior.**  
   Unknown features are preserved as cargo, not edited. `replace` will not search text inside charts, text boxes, headers as drawing text, or pivot caches — only worksheet string cells / SST / inlineStr as documented.

6. **Macros / binary parts**  
   If present and related (e.g. `vbaProject.bin`), they round-trip as blobs. The library does not inspect or resign them.

### Practical checklist

After open → replace → save on a “fancy” workbook:

| Expect kept | Expect maybe re-serialized | Expect ignored by replace |
|-------------|----------------------------|---------------------------|
| Charts, images, theme, styles, untouched sheets (byte-identical) | `sharedStrings.xml`, sheets whose cells changed, workbook | Chart titles, comments UI drawing text*, pivots |
| VBA / customXml / slicers (as blob), `[Content_Types].xml`, `.rels` | `[Content_Types].xml`, `.rels` only when parts or relationships are added or removed | Non-string cell values |
| Orphan ZIP members | only if `include_orphans=True` | never edited |

\*Comments and some UI text live in other parts; v1 does not target them.

---

## 6. Mental model vs openpyxl

| | `xlsxedit` | openpyxl-style |
|--|---------------|----------------|
| Source of truth | Live package XML / blobs | Python object model |
| Unknown feature | Opaque part round-trips | Often dropped when writing a model that does not know it |
| Change one string | Rewrites the SST (and a sheet only if its cells changed) | May rewrite styles / sheets heavily |
| Goal | Preserve on-disk fidelity while editing text | Rich authoring API |

Neither approach produces a byte-identical ZIP: the container itself (member order, compression, timestamps) is rewritten. The promise here is that **parts an edit does not change keep their exact bytes**, and that unknown parts on the relationship graph always survive.

---

## 7. Code map

| Path | Responsibility |
|------|----------------|
| `workbook.py` | `Workbook.open` / `create` / `save` / SAR / sheet lifecycle |
| `worksheet.py` | `ws[address]`, dimensions, merges, images/charts/tables |
| `cell.py` | Typed `value`, `formula`, `hyperlink`, style writes |
| `dimensions.py` | Column/row width and height |
| `styles.py` | Read/write `cellXfs` (clone xf, number formats) |
| `merge.py` | Merge range parsing and anchor map |
| `hyperlinks.py` | Cell hyperlink get/set |
| `conditional_formatting.py` | CF list and `cellIs` add |
| `drawing.py` | `Picture`, `Chart`, `Table` proxies |
| `shared_strings.py` | SST display text + in-place `<t>` replace |
| `parts.py` | Registers workbook / worksheet / SST part types |
| `opc/package.py` | Load/save orchestration |
| `opc/pkgreader.py` | Relationship-graph walk |
| `opc/pkgwriter.py` | ZIP write + content types |
| `opc/part.py` | Opaque `Part` vs `XmlPart` |

---

## 8. Extending safely later

To support a new feature **without** losing round-trip safety:

1. Prefer leaving it opaque until you must edit it.
2. When you need to edit it, register a new `XmlPart` subclass for that content type only.
3. Mutate the smallest nodes possible; keep sibling XML intact.
4. Add tests that open a fixture containing that feature, `replace` something unrelated, save, and assert the feature part’s bytes (or critical XML) still match.

Until then, unknown = blob = kept (if it was in the relationship graph).

---

## 9. Roadmap status

| Phase | Status | Features |
|-------|--------|----------|
| 1 — Create | done | `Workbook.create()`, `default.xlsx` template |
| 2 — Cells | done | `ws["B2"]`, typed `value`, `formula`, `cell.clear()` |
| 3 — Sheets | done | `add_worksheet`, `rename_worksheet`, `remove_worksheet` |
| 4 — Styles & layout | done | `cell.style` read, `apply_date_format`, `apply_number_format`, column/row dimensions, `merge_cells` |
| 5 — Images | done | `add_image`, `Picture` geometry, `replace_image`, `insert_image_at_placeholder` |
| 6 — Features | partial | Hyperlinks, CF read/add (`cellIs`), chart add/title, table add/resize; pivots/comments future |
| SAR | done | `replace`, typed replace (delegates to cell APIs), image SAR |

**Architecture rule:** new SAR helpers must call base methods first (e.g. `replace(..., value_type="date")` → `cell.value` + `apply_date_format()`).
