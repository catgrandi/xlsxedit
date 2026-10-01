"""Benchmark bulk export.

Deselected by default (see ``addopts`` in pyproject.toml). Run it with
``pytest -m bench -s``; set ``BENCH_ROWS=100000`` for full scale.
"""

from __future__ import annotations

import os
import time
from pathlib import Path

import pytest

from xlsxedit import Workbook

BENCH_ROWS = int(os.environ.get("BENCH_ROWS", "2000"))


class _NumericFrame:
    def __init__(self, n: int):
        self.columns = ["A", "B", "C"]
        self._n = n

    def itertuples(self, index=False):
        for i in range(self._n):
            yield (i, i * 2, i * 3)


@pytest.mark.bench
def test_bench_write_rows(tmp_path: Path):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "Col1"
    ws["B1"].value = "Col2"
    ws["C1"].value = "Col3"
    df = _NumericFrame(BENCH_ROWS)

    t0 = time.perf_counter()
    wb.write_dataframe(
        df,
        at_cell="A2",
        header=False,
        template_rows=1,
    )
    write_s = time.perf_counter() - t0

    out = tmp_path / "bench.xlsx"
    t1 = time.perf_counter()
    wb.save(out)
    save_s = time.perf_counter() - t1

    print(f"\nBENCH_ROWS={BENCH_ROWS} write={write_s:.2f}s save={save_s:.2f}s")
    assert out.stat().st_size > 0
