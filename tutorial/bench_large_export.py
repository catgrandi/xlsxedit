#!/usr/bin/env python3
"""
Benchmark large zebra-styled write_dataframe exports.

Run from the project folder:
    python tutorial/bench_large_export.py

Change row count:
    ROWS=100000 python tutorial/bench_large_export.py

Optional:
    MODE=insert     — insert instead of overwrite
    NO_PANDAS=1     — force lazy fallback

Output: tutorial/output/bench_large_export.xlsx
"""

from __future__ import annotations

import os
import sys
import time
from pathlib import Path

# Allow ``from xlsxedit import …`` when running from a GitHub clone without
# ``pip install -e .`` (src layout: repo/src/xlsxedit).
_SRC = Path(__file__).resolve().parent.parent / "src"
if _SRC.is_dir() and str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from xlsxedit import Workbook
from xlsxedit.oxml.address import index_to_col, join_address

ROWS = int(os.environ.get("ROWS", "50000"))
MODE = os.environ.get("MODE", "overwrite")
COLS = int(os.environ.get("COLS", "4"))
START = os.environ.get("START", "A2")
NO_PANDAS = os.environ.get("NO_PANDAS", "").lower() in ("1", "true", "yes")

HERE = Path(__file__).resolve().parent
OUTPUT = HERE / "output" / "bench_large_export.xlsx"


class _LazyNumericFrame:
    def __init__(self, n: int, ncol: int):
        self.columns = [f"Col{i + 1}" for i in range(ncol)]
        self._n = n
        self._ncol = ncol

    def itertuples(self, index=False):
        for i in range(self._n):
            yield tuple(float(i + c * 0.01) for c in range(self._ncol))


def make_dataframe(rows: int, cols: int):
    if NO_PANDAS:
        return _LazyNumericFrame(rows, cols), "lazy"

    try:
        import numpy as np
        import pandas as pd
    except ImportError:
        return _LazyNumericFrame(rows, cols), "lazy"

    base = np.arange(rows, dtype=np.float64)
    data = {f"Col{i + 1}": base + i * 0.01 for i in range(cols)}
    return pd.DataFrame(data), "pandas"


def style_header_row(ws, columns) -> None:
    """Row 1 = bold column titles (not part of the bulk write)."""
    for col_idx, name in enumerate(columns):
        cell = ws[join_address(index_to_col(col_idx), 1)]
        cell.value = name
        cell.apply_style(bold=True, bg_color="FF4472C4", font_color="FFFFFFFF")


def main() -> None:
    if MODE not in ("overwrite", "insert"):
        raise SystemExit(f"MODE must be overwrite or insert, got {MODE!r}")

    print(f"Rows={ROWS:,}  cols={COLS}  start={START}  mode={MODE}  zebra", flush=True)
    print(f"Output → {OUTPUT}", flush=True)

    marks: list[tuple[str, float]] = []
    clock = time.perf_counter()

    df, backend = make_dataframe(ROWS, COLS)
    now = time.perf_counter()
    marks.append((f"build DataFrame ({backend})", now - clock))
    clock = now

    wb = Workbook.create()
    now = time.perf_counter()
    marks.append(("Workbook.create()", now - clock))
    clock = now

    ws = wb["Sheet1"]
    style_header_row(ws, df.columns)
    now = time.perf_counter()
    marks.append(("header row (bold)", now - clock))
    clock = now

    print("writing zebra rows…", flush=True)
    written = wb.write_dataframe(
        df,
        at_cell=START,
        header=False,
        mode=MODE,
        resize_table=False,
        row_styles=[
            {"bg_color": "FFFFFFFF"},
            {"bg_color": "FF9DC3E6"},
        ],
        column_styles=[
            {},
            {"num_format": "0"},
            {"num_format": "$#,##0.00"},
            {"num_format": "$#,##0.00"},
        ][:COLS],
    )
    now = time.perf_counter()
    marks.append(("write_dataframe (zebra)", now - clock))
    clock = now

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    print("saving…", flush=True)
    wb.save(OUTPUT)
    now = time.perf_counter()
    marks.append(("save .xlsx", now - clock))

    size_mb = OUTPUT.stat().st_size / (1024 * 1024)
    total = sum(sec for _, sec in marks)

    print()
    print(f"rows written: {written:,}")
    print(f"file size:    {size_mb:.1f} MB")
    print()
    print(f"{'step':<28} {'seconds':>8}  {'%':>6}")
    print("-" * 46)
    for label, sec in marks:
        pct = 100.0 * sec / total if total else 0
        extra = f"  ({written / sec:,.0f} rows/s)" if "write_dataframe" in label and sec > 0 else ""
        print(f"{label:<28} {sec:8.2f}  {pct:5.1f}%{extra}")
    print("-" * 46)
    print(f"{'TOTAL':<28} {total:8.2f}  100.0%")


if __name__ == "__main__":
    main()
