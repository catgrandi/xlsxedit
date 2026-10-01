# Vendored OOXML schemas

`transitional/` holds the W3C XML Schema (XSD) files that define
SpreadsheetML in ECMA-376 Part 4, *Transitional Migration Features*, 5th
edition (December 2016): `sml.xsd` and the nine schemas it imports directly
or indirectly.

[`tests/schema_validate.py`](../tests/schema_validate.py) validates the parts
of saved workbooks against them during the test suite.

The directory is outside `src/`, and `MANIFEST.in` does not include it, so
the schemas never ship in the wheel or the source distribution.

## Why the transitional set

Workbooks that Excel writes use the transitional namespaces, such as
`http://schemas.openxmlformats.org/spreadsheetml/2006/main`. The strict
schemas target `http://purl.oclc.org/ooxml/` namespaces and cannot validate
those workbooks.

## Source

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

## Verify the copy

From the repository root, download the archive, check its hash, and compare
each file while ignoring line endings:

```bash
curl -LO https://ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip
sha256sum ECMA-376-4_5th_edition_december_2016.zip
unzip -o ECMA-376-4_5th_edition_december_2016.zip OfficeOpenXML-XMLSchema-Transitional.zip
unzip -o -d ecma OfficeOpenXML-XMLSchema-Transitional.zip
for f in schemas/transitional/*.xsd; do diff --strip-trailing-cr "$f" "ecma/${f##*/}"; done
```

The loop prints nothing when every file matches.

## License

ECMA-376 is copyright Ecma International, which publishes it free of
charge. The schema files carry no copyright or license notice of their own.
The [Ecma text copyright policy](https://ecma-international.org/policies/by-ipr/ecma-text-copyright-policy/)
permits copying and distributing Ecma standards with the copyright notice,
and permits using a standard to implement its functionality in conforming
products. This repository uses the files only to run its tests and does not
include them in its distributions.
