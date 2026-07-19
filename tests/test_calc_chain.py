"""Stale formula metadata: calcChain + shared-formula followers."""

from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

from xlsxedit import Workbook
from xlsxedit.opc.constants import SML_NS
from xlsxedit.oxml.parser import parse_xml
from tests.conftest import INSPECT_FIXTURES

_F = f"{{{SML_NS}}}f"
_C = f"{{{SML_NS}}}c"
_FIXTURE = INSPECT_FIXTURES["SimpleFormula"]


def _has_calc_chain(path: Path) -> bool:
    with ZipFile(path) as z:
        return "xl/calcChain.xml" in z.namelist()


def _formulas(path: Path) -> dict[str, tuple[str | None, str | None, str]]:
    """Return {address: (f@t, f@si, f.text)} for formula cells in sheet1."""
    with ZipFile(path) as z:
        root = parse_xml(z.read("xl/worksheets/sheet1.xml"))
    out: dict[str, tuple[str | None, str | None, str]] = {}
    for c_elm in root.iter(_C):
        f_elm = c_elm.find(_F)
        if f_elm is None:
            continue
        out[c_elm.get("r", "")] = (f_elm.get("t"), f_elm.get("si"), f_elm.text or "")
    return out


def test_fixture_has_calc_chain():
    assert _has_calc_chain(_FIXTURE)


def test_overwrite_normal_formula_drops_calc_chain(tmp_path: Path):
    out = tmp_path / "out.xlsx"
    wb = Workbook.open(_FIXTURE)
    ws = wb.worksheets[0]
    assert ws["C3"].has_formula
    ws["C3"].value = 99
    assert not ws["C3"].has_formula
    wb.save(out)

    assert not _has_calc_chain(out)
    formulas = _formulas(out)
    assert "C3" not in formulas
    assert "C8" in formulas  # SUM remains
    assert formulas["C8"][2].startswith("SUM")


def test_overwrite_shared_master_detaches_followers(tmp_path: Path):
    out = tmp_path / "master.xlsx"
    wb = Workbook.open(_FIXTURE)
    ws = wb.worksheets[0]
    # C4 is shared master with formula text; C5:C7 are followers.
    assert ws["C4"].has_formula
    ws["C4"].value = 42
    wb.save(out)

    assert not _has_calc_chain(out)
    formulas = _formulas(out)
    assert "C4" not in formulas
    assert "C5" not in formulas
    assert "C6" not in formulas
    assert "C7" not in formulas
    assert "C3" in formulas
    assert "C8" in formulas


def test_overwrite_shared_follower_keeps_master(tmp_path: Path):
    out = tmp_path / "follower.xlsx"
    wb = Workbook.open(_FIXTURE)
    ws = wb.worksheets[0]
    ws["C5"].value = 7
    wb.save(out)

    assert not _has_calc_chain(out)
    formulas = _formulas(out)
    assert "C5" not in formulas
    assert formulas["C4"][0] == "shared"
    assert formulas["C4"][2]  # master still has text
    assert "C6" in formulas
    assert "C7" in formulas


def test_clear_formula_drops_calc_chain(tmp_path: Path):
    out = tmp_path / "clear.xlsx"
    wb = Workbook.open(_FIXTURE)
    wb.worksheets[0]["C3"].clear()
    wb.save(out)
    assert not _has_calc_chain(out)
    assert "C3" not in _formulas(out)


def test_clear_range_drops_calc_chain(tmp_path: Path):
    out = tmp_path / "range.xlsx"
    wb = Workbook.open(_FIXTURE)
    wb.worksheets[0].clear_range("C3:C3")
    wb.save(out)
    assert not _has_calc_chain(out)


def test_write_rows_at_over_formula_drops_calc_chain(tmp_path: Path):
    out = tmp_path / "bulk.xlsx"
    wb = Workbook.open(_FIXTURE)
    ws = wb.worksheets[0]
    # C3 is Excel row 3, col index 2 (0-based C).
    ws._write_rows_at([(123,)], start_row=3, start_col_idx=2)
    wb.save(out)
    assert not _has_calc_chain(out)
    assert "C3" not in _formulas(out)


def test_non_formula_edit_preserves_calc_chain(tmp_path: Path):
    out = tmp_path / "keep.xlsx"
    wb = Workbook.open(_FIXTURE)
    ws = wb.worksheets[0]
    ws["A3"].value = 100  # operand, not a formula cell
    wb.save(out)
    assert _has_calc_chain(out)
    assert "C3" in _formulas(out)


def test_formula_none_drops_calc_chain(tmp_path: Path):
    out = tmp_path / "fnone.xlsx"
    wb = Workbook.open(_FIXTURE)
    wb.worksheets[0]["C3"].formula = None
    wb.save(out)
    assert not _has_calc_chain(out)
    assert "C3" not in _formulas(out)
