#!/usr/bin/env python3
"""Print a summary of an xlsxedit workbook, or compare two workbooks.

``--compare OTHER`` lists the package members whose bytes differ.

``--semantic`` reads the package with the test harness instead
(``tests/excel_facts.py`` and the canonicaliser in ``tests/preservation.py``)
and prints it aspect by aspect:

- sheets, in tab order, with their visibility;
- defined names;
- tables: every attribute and child, each column, dxf indexes resolved to the
  format they name;
- conditional formats and data validations, including their x14 forms;
- drawing anchors, with anchor type, ``editAs``, position and object;
- merges;
- row heights and column widths, with their hidden and outline state;
- formulas, by kind;
- cell values;
- parts: every package member's canonical XML (C14N), and for
  ``[Content_Types].xml`` the content type each member resolves to.

Revision ids such as ``xr:uid`` are left out, as in ``tests/excel_facts.py``.
With ``--compare OTHER`` it reports what differs in each aspect, ignoring how
the XML is serialised, lists the members that were only re-serialised, and
exits with status 1 when anything differs.
"""

from __future__ import annotations

import argparse
import hashlib
import sys
import zipfile
from collections import Counter
from collections.abc import Iterable, Mapping
from pathlib import Path
from typing import NamedTuple

from lxml import etree

ROOT = Path(__file__).resolve().parents[1]

if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from tests import excel_facts as facts  # noqa: E402
from tests.preservation import CONTENT_TYPES, canon, content_types, read_pkg  # noqa: E402
from xlsxedit import Workbook  # noqa: E402

Aspects = dict[str, dict[str, str]]

_WHOLE_SHEET = "A1:XFD1048576"
_SHOWN = 30

# Attributes that point into styles.xml's dxfs; their index is replaced by the format it names.
_DXF_ATTRIBUTES = frozenset(
    {
        "headerRowDxfId",
        "dataDxfId",
        "totalsRowDxfId",
        "headerRowBorderDxfId",
        "tableBorderDxfId",
        "totalsRowBorderDxfId",
    }
)


def summarize(path: Path) -> None:
    wb = Workbook.open(path)
    print(f"Workbook: {path}")
    print(f"  sheets: {wb.sheetnames}")
    if wb.orphan_partnames:
        print(f"  orphans: {wb.orphan_partnames}")
    for ws in wb.worksheets:
        print(f"\n  [{ws.name}]")
        print(f"    cells: {len(ws.cells)}")
        print(f"    merges: {len(ws.merged_ranges)}")
        if ws.images:
            print(f"    images: {len(ws.images)}")
            for pic in ws.images:
                print(f"      - {pic.name!r} @ {pic.anchor} ({pic.width}x{pic.height}px)")
        if ws.charts:
            print(f"    charts: {len(ws.charts)}")
            for ch in ws.charts:
                print(f"      - {ch.name!r} @ {ch.anchor}")
        if ws.tables:
            print(f"    tables: {len(ws.tables)}")
            for tbl in ws.tables:
                print(f"      - {tbl.name!r} {tbl.ref} cols={tbl.columns}")


def compare(a: Path, b: Path) -> None:
    with zipfile.ZipFile(a) as za, zipfile.ZipFile(b) as zb:
        names = sorted(set(za.namelist()) | set(zb.namelist()))
        diffs = []
        for n in names:
            try:
                ba, bb = za.read(n), zb.read(n)
            except KeyError:
                diffs.append(n)
                continue
            if ba != bb:
                diffs.append(n)
        print(f"Compare {a.name} vs {b.name}: {len(diffs)} differing parts")
        for n in diffs[:30]:
            print(f"  {n}")
        if len(diffs) > 30:
            print(f"  ... and {len(diffs) - 30} more")


# semantic aspects ------------------------------------------------------------


def _unique(items: Iterable[tuple[str, str]]) -> dict[str, str]:
    """Key the items, appending ``#n`` from the second use of a key on."""
    found: dict[str, str] = {}
    seen: Counter[str] = Counter()
    for key, value in items:
        seen[key] += 1
        found[key if seen[key] == 1 else f"{key} #{seen[key]}"] = value
    return found


def _with_formats(book: facts.Book, elm) -> str:
    """``facts.describe`` of ``elm``, with each dxf index replaced by the format it names."""
    text = facts.describe(elm, drop=_DXF_ATTRIBUTES)
    formats = [
        f"{name}={book.dxf(elm.get(name))}" for name in sorted(_DXF_ATTRIBUTES) if elm.get(name)
    ]
    return " ".join([text, *formats])


def _sheets(book: facts.Book) -> dict[str, str]:
    return {sheet.name: f"position {i}, {sheet.state}" for i, sheet in enumerate(book.sheets, 1)}


def _defined_names(book: facts.Book) -> dict[str, str]:
    return {
        f"{scope}!{name}" if scope else name: formula
        for (name, scope), formula in sorted(facts.defined_names(book).items())
    }


def _tables(book: facts.Book) -> dict[str, str]:
    items = []
    for sheet in book.sheets:
        for member in book.related(sheet.member, "table"):
            root = book.tree(member)
            key = f"{sheet.name} table {root.get('name', '')}"
            attributes = {k: v for k, v in root.attrib.items() if not k.startswith(f"{{{facts.MC}}}")}
            shallow = root.makeelement(root.tag, attributes)
            items.append((key, _with_formats(book, shallow)))
            for child in root.iterchildren(etree.Element):
                if child.tag == facts.sml("tableColumns"):
                    for column in child.iterchildren(facts.sml("tableColumn")):
                        name = f"{key} column {column.get('name', '')}"
                        items.append((name, _with_formats(book, column)))
                else:
                    items.append((f"{key} {etree.QName(child).localname}", facts.describe(child)))
    return _unique(items)


def _conditional_formats(book: facts.Book) -> dict[str, str]:
    items = []
    for sheet, rules in facts.conditional_formats(book).items():
        for rank, rule in enumerate(rules, 1):
            source = f" ({rule.source})" if rule.source else ""
            items.append((f"{sheet} {rule.sqref} {rule.type}{source}", f"#{rank} {rule.detail}"))
    return _unique(items)


def _data_validations(book: facts.Book) -> dict[str, str]:
    items = []
    for sheet, validations in facts.data_validations(book).items():
        for v in validations:
            source = f" ({v.source})" if v.source else ""
            items.append((f"{sheet} {v.sqref}{source}", f"{v.type} {v.detail}".rstrip()))
    return _unique(items)


def _marker(marker: tuple[int, int, int, int] | None) -> str:
    if marker is None:
        return "?"
    col, col_off, row, row_off = marker
    return f"{facts.col_letters(col + 1)}{row + 1} (+{col_off}, +{row_off})"


def _anchor(anchor: facts.Anchor) -> str:
    kind = anchor.kind + (f' editAs="{anchor.edit_as}"' if anchor.edit_as else "")
    if anchor.kind == "absoluteAnchor":
        where = f"at {anchor.position} size {anchor.extent}"
    elif anchor.kind == "oneCellAnchor":
        where = f"{_marker(anchor.start)} size {anchor.extent}"
    else:
        where = f"{_marker(anchor.start)} to {_marker(anchor.end)}"
    objects = ", ".join(f"{kind} {name!r}" for kind, name in anchor.objects) or "nothing"
    return f"{kind} {where}: {objects}"


def _drawings(book: facts.Book) -> dict[str, str]:
    return {
        f"{sheet}/{name}": _anchor(anchor)
        for sheet, anchors in facts.drawings_by_name(book).items()
        for name, anchor in anchors.items()
    }


def _merges(book: facts.Book) -> dict[str, str]:
    return {f"{sheet}!{ref}": "" for sheet, refs in facts.merges(book).items() for ref in refs}


def _rows_and_columns(book: facts.Book) -> dict[str, str]:
    found = {}
    for sheet in book.sheets:
        for number, text in facts.rows(book, f"'{sheet.name}'!{_WHOLE_SHEET}").items():
            found[f"{sheet.name} row {number}"] = text
        for col in book.tree(sheet.member).iter(facts.sml("col")):
            parts = []
            if col.get("customWidth") in ("1", "true"):
                parts.append(f"width={col.get('width')}")
            for flag in ("hidden", "collapsed"):
                if col.get(flag) in ("1", "true"):
                    parts.append(flag)
            if col.get("outlineLevel") not in (None, "0"):
                parts.append(f"outlineLevel={col.get('outlineLevel')}")
            if col.get("style") is not None:
                parts.append(f"format {book.cell_style(col.get('style'))}".rstrip())
            if parts:
                first = facts.col_letters(int(col.get("min", "1")))
                last = facts.col_letters(int(col.get("max", "1")))
                found[f"{sheet.name} columns {first}:{last}"] = " ".join(parts)
    return found


def _formulas_and_cells(book: facts.Book) -> tuple[dict[str, str], dict[str, str]]:
    formulas, cells = {}, {}
    for sheet in book.sheets:
        for address, fact in facts.cells(book, f"'{sheet.name}'!{_WHOLE_SHEET}").items():
            key = f"{sheet.name}!{address}"
            formula = fact.formula
            if formula is None:
                cells[key] = f"{fact.type} {fact.value!r}"
                continue
            parts = [formula.kind]
            if formula.ref:
                parts.append(f"ref={formula.ref}")
            if formula.options:
                parts.append(formula.options)
            parts.append(f"={formula.text}" if formula.text is not None else "(shared follower)")
            formulas[key] = " ".join(parts)
    return formulas, cells


def _parts(pkg: Mapping[str, bytes]) -> dict[str, str]:
    """Each member's canonical form as a digest; content types as their effective map."""
    found = {}
    for name, blob in pkg.items():
        if name == CONTENT_TYPES:
            blob = repr(sorted(content_types(pkg).items())).encode("utf-8")
        else:
            blob = canon(name, blob)
        found[name] = hashlib.sha256(blob).hexdigest()
    return found


def semantic_aspects(source) -> Aspects:
    """What a workbook holds, by aspect: ``{aspect: {item: description}}``.

    ``source`` is a ``read_pkg`` dictionary or anything ``read_pkg`` accepts.
    """
    pkg = source if isinstance(source, Mapping) else read_pkg(source)
    book = facts.Book(pkg)
    formulas, cells = _formulas_and_cells(book)
    return {
        "sheets": _sheets(book),
        "defined names": _defined_names(book),
        "tables": _tables(book),
        "conditional formats": _conditional_formats(book),
        "data validations": _data_validations(book),
        "drawings": _drawings(book),
        "merges": _merges(book),
        "rows and columns": _rows_and_columns(book),
        "formulas": formulas,
        "cells": cells,
        "parts": _parts(pkg),
    }


class Difference(NamedTuple):
    key: str
    before: str | None  # None when the item was added
    after: str | None  # None when the item was removed


def semantic_differences(before: Aspects, after: Aspects) -> dict[str, list[Difference]]:
    """The items of each aspect that differ, in document order; ``[]`` for an equal aspect."""
    found = {}
    for aspect, old in before.items():
        new = after[aspect]
        keys = [*old, *(key for key in new if key not in old)]
        found[aspect] = [
            Difference(key, old.get(key), new.get(key))
            for key in keys
            if old.get(key) != new.get(key)
        ]
    return found


def _item(sign: str, key: str, value: str | None, aspect: str) -> str:
    shown = value and aspect != "parts"
    return f"    {sign} {key}" + (f": {value}" if shown else "")


def print_aspects(path: Path, aspects: Aspects) -> None:
    print(f"Semantic view of {path}")
    for aspect, items in aspects.items():
        print(f"  {aspect}: {len(items)}")
        for key in list(items)[:_SHOWN]:
            print(_item(" ", key, items[key], aspect))
        if len(items) > _SHOWN:
            print(f"    ... and {len(items) - _SHOWN} more")


def print_differences(differences: dict[str, list[Difference]]) -> None:
    for aspect, items in differences.items():
        if not items:
            print(f"  {aspect}: same")
            continue
        print(f"  {aspect}: {len(items)} item(s) differ")
        for item in items[:_SHOWN]:
            if item.after is None:
                print(_item("-", item.key, item.before, aspect))
            elif item.before is None:
                print(_item("+", item.key, item.after, aspect))
            else:
                print(_item("~", item.key, None, aspect))
                if aspect != "parts":
                    print(f"        - {item.before}")
                    print(f"        + {item.after}")
        if len(items) > _SHOWN:
            print(f"    ... and {len(items) - _SHOWN} more")


def semantic_compare(a: Path, b: Path) -> int:
    """Print what differs in meaning between two workbooks; 1 when anything does."""
    pkg_a, pkg_b = read_pkg(a), read_pkg(b)
    before, after = semantic_aspects(pkg_a), semantic_aspects(pkg_b)
    differences = semantic_differences(before, after)
    differing = sum(1 for items in differences.values() if items)
    print(f"Semantic compare {a.name} vs {b.name}: {differing} of {len(differences)} aspects differ")
    print_differences(differences)
    reserialised = sorted(
        name
        for name in pkg_a.keys() & pkg_b.keys()
        if pkg_a[name] != pkg_b[name] and before["parts"][name] == after["parts"][name]
    )
    if reserialised:
        print(f"  re-serialised, canonically equal: {', '.join(reserialised)}")
    return 1 if differing else 0


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("workbook", type=Path, help=".xlsx file to inspect")
    p.add_argument("--compare", type=Path, metavar="OTHER", help="compare with another workbook")
    p.add_argument(
        "--semantic",
        action="store_true",
        help="read the package aspect by aspect; with --compare, exit 1 when anything differs",
    )
    args = p.parse_args(argv)
    for path in (args.workbook, args.compare):
        if path is not None and not path.is_file():
            raise SystemExit(f"not found: {path}")
    if args.semantic:
        if args.compare:
            return semantic_compare(args.workbook, args.compare)
        print_aspects(args.workbook, semantic_aspects(args.workbook))
        return 0
    summarize(args.workbook)
    if args.compare:
        print()
        compare(args.workbook, args.compare)
    return 0


if __name__ == "__main__":
    sys.exit(main())
