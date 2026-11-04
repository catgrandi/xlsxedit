"""Benchmark inserts and the save after them (#56).

Deselected by default (see ``addopts`` in pyproject.toml). Run it with
``pytest -m bench -s``; set ``BENCH_ROWS=50000`` for the sheet the issue
measures.
"""

from __future__ import annotations

import gc
import io
import os
import time
from collections.abc import Callable

import pytest

from xlsxedit import Workbook

BENCH_ROWS = int(os.environ.get("BENCH_ROWS", "2000"))


def _timed(fn: Callable[[], object]) -> float:
    gc.collect()
    start = time.perf_counter()
    fn()
    return time.perf_counter() - start


def _saved(wb: Workbook) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


@pytest.mark.bench
def test_bench_one_row_insert_with_a_chart_and_formats():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws.write_rows(
        [[n, n * 1.5, n % 2 == 0, f"s{n % 97}", n * 3, n / 7, "t", n % 13] for n in range(1, BENCH_ROWS + 1)],
        at_row=1,
    )
    ws.add_chart("bar", anchor="J2", data_range=f"A1:B{BENCH_ROWS}", title="Data")
    ws.add_conditional_formatting(f"A2:A{BENCH_ROWS}", operator="greaterThan", formula="100")
    ws.add_conditional_formatting(f"E2:E{BENCH_ROWS}", operator="lessThan", formula="$A$1")
    ws.add_color_scale_formatting(f"B2:B{BENCH_ROWS}")
    wb = Workbook.open(io.BytesIO(_saved(wb)))
    ws = wb["Sheet1"]

    row = [1, 2.5, True, "new", 5, 6, "t", 7]
    insert = _timed(lambda: ws.insert_rows([row], at_row=BENCH_ROWS // 2))
    save = _timed(lambda: _saved(wb))

    print(f"\nBENCH_ROWS={BENCH_ROWS} one-row insert={insert:.3f}s save={save:.3f}s")
    assert ws[f"D{BENCH_ROWS // 2}"].value == "new"


@pytest.mark.bench
def test_bench_twenty_inserts_with_ten_charts():
    wb = Workbook.create()
    for i in range(10):
        ws = wb["Sheet1"] if i == 0 else wb.add_worksheet(f"Data{i + 1}")
        ws.write_rows([[n, n * 2, n * 3, f"r{n}"] for n in range(1, 301)], at_row=1)
        ws.add_chart("bar", anchor="F2", data_range="A1:B300", title=f"Chart {i + 1}")
    wb = Workbook.open(io.BytesIO(_saved(wb)))
    ws = wb["Sheet1"]

    insert = _timed(lambda: [ws.insert_rows([[n]], at_row=10) for n in range(20)])
    save = _timed(lambda: _saved(wb))

    print(f"\n20 inserts with 10 charts={insert:.3f}s save={save:.3f}s")
    assert ws["A10"].value == 19
