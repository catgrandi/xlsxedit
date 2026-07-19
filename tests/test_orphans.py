from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

from xlsxedit import Workbook

ORPHAN_NAME = "custom/orphan.bin"
ORPHAN_BYTES = b"orphan-cargo-bytes"


def _book_with_orphan(book1_path: Path, dest: Path) -> Path:
    shutil.copy(book1_path, dest)
    with zipfile.ZipFile(dest, "a") as zf:
        zf.writestr(ORPHAN_NAME, ORPHAN_BYTES)
    return dest


def test_normal_book_has_no_orphans(book1_path: Path):
    wb = Workbook.open(book1_path)
    assert wb.orphan_partnames == ()


def test_orphan_is_discovered(book1_path: Path, tmp_path: Path):
    path = _book_with_orphan(book1_path, tmp_path / "with_orphan.xlsx")
    wb = Workbook.open(path)
    assert wb.orphan_partnames == ("/custom/orphan.bin",)


def test_save_drops_orphans_by_default(book1_path: Path, tmp_path: Path):
    path = _book_with_orphan(book1_path, tmp_path / "with_orphan.xlsx")
    out = tmp_path / "out_default.xlsx"
    wb = Workbook.open(path)
    wb.save(out)
    with zipfile.ZipFile(out) as zf:
        assert ORPHAN_NAME not in zf.namelist()


def test_save_include_orphans_keeps_bytes(book1_path: Path, tmp_path: Path):
    path = _book_with_orphan(book1_path, tmp_path / "with_orphan.xlsx")
    out = tmp_path / "out_keep.xlsx"
    wb = Workbook.open(path)
    n = wb.replace("lol", "hi")
    assert n == 1
    wb.save(out, include_orphans=True)

    with zipfile.ZipFile(book1_path) as orig, zipfile.ZipFile(out) as saved:
        assert ORPHAN_NAME in saved.namelist()
        assert saved.read(ORPHAN_NAME) == ORPHAN_BYTES
        assert orig.read("xl/styles.xml") == saved.read("xl/styles.xml")
        assert orig.read("xl/theme/theme1.xml") == saved.read("xl/theme/theme1.xml")

    wb2 = Workbook.open(out)
    assert any(c.value == "hi" for c in wb2["test"].cells)
    assert wb2.orphan_partnames == ("/custom/orphan.bin",)
