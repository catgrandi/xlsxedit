# Vendored OOXML schemas

This directory holds W3C XML Schema (XSD) files for the formats xlsxedit
reads and writes:

- `transitional/` holds the schemas that define SpreadsheetML in ECMA-376
  Part 4, *Transitional Migration Features*, 5th edition (December 2016):
  `sml.xsd` and the nine schemas it imports directly or indirectly.
- `microsoft/` holds Microsoft's schemas for the extension namespaces that
  Excel writes, such as x14 conditional formatting, data validation and
  sparklines, from [MS-XLSX] Appendix A, with the files that let them
  compile against `transitional/`.

Three tools use them:

- [`scripts/gen_ooxml_order.py`](../scripts/gen_ooxml_order.py) generates
  the child-element order tables in `src/xlsxedit/_ooxml_order.py`: `ORDER`
  from `transitional/`, `EXTENSION_ORDER` from `microsoft/ms-xlsx/`. Run it
  again after changing a schema; `tests/test_ooxml_order.py` fails until the
  generated module matches.
- [`tests/schema_validate.py`](../tests/schema_validate.py) validates the
  parts of saved workbooks during the test suite.
- [`scripts/fetch_ms_xlsx_schemas.py`](../scripts/fetch_ms_xlsx_schemas.py)
  writes `microsoft/`, except `sml-shim.xsd`. Nothing is fetched when the
  tests run.

The directory is outside `src/`, and `MANIFEST.in` does not include it, so
the schemas never ship in the wheel or the source distribution.

## ECMA-376 transitional schemas

### Why the transitional set

Workbooks that Excel writes use the transitional namespaces, such as
`http://schemas.openxmlformats.org/spreadsheetml/2006/main`. The strict
schemas target `http://purl.oclc.org/ooxml/` namespaces and cannot validate
those workbooks.

### Source

| Item | Value |
| --- | --- |
| Download | <https://ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip> |
| Download SHA-256 | `bd25da1109f73762356596918bf5ff8b74a1331642dba5f1c1d1dfc6bed34ecd` |
| Inner archive | `OfficeOpenXML-XMLSchema-Transitional.zip` |
| Inner archive SHA-256 | `d34187520749998af306faf1b730e568b0ca6d88ad24638a407c0a9bb4ca04fc` |
| Retrieved | 2026-10-01 |

The files are copied from the inner archive without edits, except that line
endings are LF: the repository stores every text file with LF, and five of
the files use CRLF in the archive. The hashes are of the committed files.

| File | SHA-256 |
| --- | --- |
| `dml-chart.xsd` | `41b93bd8857cc68b1e43be2806a872d736a9bdd6566900062d8fdb57d7bbb354` |
| `dml-chartDrawing.xsd` | `3fd0586f2637b98bb9886f0e0b67d89e1cc987c2d158cc7deb5f5b9890ced412` |
| `dml-diagram.xsd` | `29b254ee0d10414a8504b5a08149c7baec35a60d5ff607d6b3f492aa36815f40` |
| `dml-lockedCanvas.xsd` | `5cb76dabd8b97d1e9308a1700b90c20139be4d50792d21a7f09789f5cccd6026` |
| `dml-main.xsd` | `5375417f0f5394b8dd1a7035b9679151f19a6b65df309dec10cfb4a420cb00e9` |
| `dml-picture.xsd` | `5d389d42befbebd91945d620242347caecd3367f9a3a7cf8d97949507ae1f53c` |
| `dml-spreadsheetDrawing.xsd` | `b4532b6d258832953fbb3ee4c711f4fe25d3faf46a10644b2505f17010d01e88` |
| `shared-commonSimpleTypes.xsd` | `8df6a1d927bbefb091a5233bd4e872f7992f62b4065fe1f129caf4fa5da68ad7` |
| `shared-relationshipReference.xsd` | `12264f3c03d738311cd9237d212f1c07479e70f0cbe1ae725d29b36539aef637` |
| `sml.xsd` | `beffeed56945c22a77440122c8bdc426f3fcbe7f3b12ea0976c770d1f8d54578` |

### Verify the copy

From anywhere inside the repository, download the archive into a temporary
directory, check its hash, and compare each file with the committed copy
while ignoring line endings:

```bash
repo=$(git rev-parse --show-toplevel)
cd "$(mktemp -d)"
curl -LO https://ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip
sha256sum ECMA-376-4_5th_edition_december_2016.zip
unzip -o ECMA-376-4_5th_edition_december_2016.zip OfficeOpenXML-XMLSchema-Transitional.zip
unzip -o -d ecma OfficeOpenXML-XMLSchema-Transitional.zip
for f in "$repo"/schemas/transitional/*.xsd; do diff --strip-trailing-cr "$f" "ecma/${f##*/}"; done
```

The loop prints nothing when every file matches.

### License

ECMA-376 is copyright Ecma International, which publishes it free of
charge. The schema files carry no copyright or license notice of their own.
The [Ecma text copyright policy](https://ecma-international.org/policies/by-ipr/ecma-text-copyright-policy/)
permits copying and distributing Ecma standards with the copyright notice,
and permits using a standard to implement its functionality in conforming
products. This repository uses the files only to generate code and to run
its tests, and does not include them in its distributions.

## Microsoft extension schemas

[MS-XLSX] Appendix A publishes one schema per extension namespace. This
directory vendors the schemas for x14
(`http://schemas.microsoft.com/office/spreadsheetml/2009/9/main`), x15
(`.../spreadsheetml/2010/11/main`), xm (`.../excel/2006/main`) and x14ac
(`.../spreadsheetml/2009/9/ac`), and for every Microsoft namespace they
import.

### Files

- `microsoft/ms-xlsx/` holds each schema as Microsoft publishes it, named
  after its namespace without the `http://schemas.microsoft.com/office/`
  prefix: x14 is `spreadsheetml-2009-9-main.xsd`.
- Microsoft's schemas import each other, and the ECMA-376 namespaces, by the
  file names Microsoft uses, such as `xl15.xsd` or `xlbasictypes.xsd`.
  `microsoft/ms-xlsx/` has an alias file for each of those names. An alias
  includes the file that defines its namespace: a vendored schema,
  `../sml-shim.xsd` for the SpreadsheetML names, or
  `transitional/shared-relationshipReference.xsd` for `orel.xsd`.
- `microsoft/sml-shim.xsd` includes `transitional/sml.xsd` and declares
  `ST_Xstring` and `ST_Guid`. Microsoft's schemas refer to both types in the
  SpreadsheetML namespace, where the ECMA-376 5th edition does not declare
  them: it declares them in the shared-types namespace. The shim derives each
  from its shared-types counterpart. No other type is missing; every
  vendored schema compiles with these two. It is the only file in
  `microsoft/` written by hand.
- `microsoft/extensions.xsd` imports SpreadsheetML through the shim and then
  every vendored namespace. Compile it to validate extension content.

libxml2 loads only the first schema document it meets for a namespace and
skips every later import of that namespace. So every SpreadsheetML alias
resolves to the one shim, and the shim, not `transitional/sml.xsd`, must
come first.

### Source

Each vendored schema is the only `<pre>` block on its Appendix A page,
where the schema is escaped text indented with no-break spaces. The script
decodes the character references, turns each no-break space into a space
(XML treats a no-break space as text, not as whitespace), removes the common indentation and trailing whitespace, and ends every line
with LF. It makes no other change.

The pages were retrieved on 2026-10-19. *Updated* is the `updated_at` date in
each page's metadata on that day. The hashes are of the committed files.

| File | Page | Updated | SHA-256 |
| --- | --- | --- | --- |
| `excel-2006-main.xsd` | [5.1](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/b0ddba06-ac73-4d3d-a913-8fbe4a56cf28) | 2025-02-18 | `81fb7525afd8ac05047edb0dfa000101b1caa848ee48a1c315ef30acf1c4570e` |
| `spreadsheetml-2010-11-main.xsd` | [5.3](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/e42bbfd7-2a3d-4308-a4f3-30313fc506b9) | 2026-10-16 | `2e09286512a4b4c72355ec4fcfb2702e0a0b16d1fc6ca3b09e6d45a5de61e32f` |
| `spreadsheetml-2009-9-main.xsd` | [5.4](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/0a377581-c743-4ace-bfcd-22f359e70165) | 2026-10-16 | `b1e7eb12a7f52aca6cda096a9c12f2758437c8491821396615f8cfc5fe4d2c17` |
| `spreadsheetml-2009-9-ac.xsd` | [5.5](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/7148c81d-4d6e-4204-be07-a7c09f4277c7) | 2024-04-16 | `03f59bdb37efbb6ead5b9e3f9768116eed77ef4691b854fe673e19379590f4dd` |
| `spreadsheetml-2014-11-main.xsd` | [5.10](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/4fe24a21-6f69-4680-882f-f86b16a69119) | 2026-05-19 | `7c507c61cb55dea65c5b2edb74ec846b393ee9b320488420a53e7ef816371100` |
| `spreadsheetml-2016-revision10.xsd` | [5.12](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/8a818a7f-92d0-4390-9855-1a484d8f9347) | 2024-04-16 | `76ef11d1ca0cf614cdd2a7c1f73ce776c69636b316091ef0783b3903e4480422` |
| `spreadsheetml-2016-revision6.xsd` | [5.14](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/e47e052c-40d0-43cd-add4-73a220966b69) | 2024-04-16 | `099b9e20ea51a46370735d2d9147edb19a08c40fabd1511bf0c0e95899816876` |
| `spreadsheetml-2014-revision.xsd` | [5.15](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/5ba37a83-255b-4e2d-ae7b-b4d14d44ad3a) | 2025-09-16 | `4cfcb84939deee3a195a3a91a6a18322b3d8618ffae304cbf1efc2c33cead9eb` |
| `spreadsheetml-2015-revision2.xsd` | [5.16](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/7b729443-d0cd-434e-8d8c-b228f701a55b) | 2024-04-16 | `59dc66129aec14de4e9758836facd6432efbc6f19897f3bb9bcd0264b6fb7dd0` |

### Update or verify the copy

The script needs network access and lxml. To rewrite `microsoft/` from the
pages, then print the rows of the table above:

```bash
python scripts/fetch_ms_xlsx_schemas.py
```

After a rewrite, replace the table rows with the printed ones and run
`scripts/gen_ooxml_order.py`. To report the committed files that differ from
the pages, with exit status 1, without writing anything:

```bash
python scripts/fetch_ms_xlsx_schemas.py --check
```

To compare each vendored schema with its copy in the
[OpenXLIFF](https://github.com/maxprograms-com/OpenXLIFF) project's
`catalog/openxml/MS-XLSX-AppendixA/` directory at commit
`f3f1c1e5e7758b7a0e9287ac707c6e6f7aa6fcd5`, ignoring whitespace, and print
every difference:

```bash
python scripts/fetch_ms_xlsx_schemas.py --verify
```

On 2026-10-19 the comparison found no difference in any of the nine files.
The mirror is a cross-check only: no file here is copied from it, so its
EPL-1.0 license does not apply to them.

### License

The schemas are copyright Microsoft. The
[Intellectual Property Rights Notice for Open Specifications Documentation](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/2c5dee00-eff2-4b22-92b6-0738acd4475e)
at the front of [MS-XLSX] permits copying the documentation to develop
implementations of the technologies it describes, and distributing the
schemas it includes in an implementation, modified or not. This repository
uses the files only to generate code and to run its tests, and does not
include them in its distributions.
