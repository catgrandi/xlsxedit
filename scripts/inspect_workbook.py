#!/usr/bin/env python3
"""Print a semantic summary of an xlsxedit workbook."""

from __future__ import annotations

import argparse
import sys
import zipfile
from pathlib import Path

from xlsxedit import Workbook


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


def main(argv: list[str] | None = None) -> None:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("workbook", type=Path, help=".xlsx file to inspect")
    p.add_argument("--compare", type=Path, metavar="OTHER", help="compare opaque parts")
    args = p.parse_args(argv)
    if not args.workbook.is_file():
        raise SystemExit(f"not found: {args.workbook}")
    summarize(args.workbook)
    if args.compare:
        if not args.compare.is_file():
            raise SystemExit(f"not found: {args.compare}")
        print()
        compare(args.workbook, args.compare)


if __name__ == "__main__":
    main()
