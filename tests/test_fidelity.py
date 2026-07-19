"""Fidelity round-trip tests for inspect fixture workbooks."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from xlsxedit import Workbook
from xlsxedit.oxml.parser import parse_xml
from xlsxedit.opc.constants import SML_NS
from tests.conftest import INSPECT_FIXTURES

_C = f"{{{SML_NS}}}c"
_MERGE = f"{{{SML_NS}}}mergeCell"
_CF = f"{{{SML_NS}}}conditionalFormatting"
_F = f"{{{SML_NS}}}f"


def _tier_a_unchanged(orig: zipfile.ZipFile, saved: zipfile.ZipFile, paths: list[str]) -> None:
    for p in paths:
        if p not in orig.namelist():
            continue
        assert orig.read(p) == saved.read(p), p


def _save_roundtrip(src: Path, tmp_path: Path, *, edit: bool) -> Path:
    out = tmp_path / f"{src.stem}_out.xlsx"
    wb = Workbook.open(src)
    if edit:
        wb[list(wb.sheetnames)[0]]["Z99"].value = "probe"
    wb.save(out)
    return out


@pytest.mark.parametrize("name", list(INSPECT_FIXTURES))
def test_fidelity_no_edit(name: str, fixtures_dir: Path, tmp_path: Path):
    src = INSPECT_FIXTURES[name]
    out = _save_roundtrip(src, tmp_path, edit=False)
    with zipfile.ZipFile(src) as orig, zipfile.ZipFile(out) as saved:
        assert set(orig.namelist()) == set(saved.namelist())
        if name == "images":
            _tier_a_unchanged(orig, saved, [
                "xl/media/image1.jpeg", "xl/drawings/drawing1.xml", "xl/drawings/drawing2.xml",
                "xl/styles.xml", "xl/theme/theme1.xml",
            ])
        elif name == "ChartsAndTables":
            _tier_a_unchanged(orig, saved, [
                "xl/charts/chart1.xml", "xl/charts/chart2.xml", "xl/tables/table1.xml",
                "xl/drawings/drawing1.xml", "xl/drawings/drawing2.xml",
                "xl/styles.xml", "xl/theme/theme1.xml",
            ])
        else:
            _tier_a_unchanged(orig, saved, ["xl/styles.xml", "xl/theme/theme1.xml"])


@pytest.mark.parametrize("name", list(INSPECT_FIXTURES))
def test_fidelity_harmless_edit(name: str, tmp_path: Path):
    src = INSPECT_FIXTURES[name]
    out = _save_roundtrip(src, tmp_path, edit=True)
    wb = Workbook.open(out)
    ws = wb[wb.sheetnames[0]]
    assert ws["Z99"].value == "probe"


def test_merged_cells_semantic(tmp_path: Path, fixtures_dir: Path):
    src = INSPECT_FIXTURES["MergedCells"]
    out = _save_roundtrip(src, tmp_path, edit=False)
    with zipfile.ZipFile(out) as z:
        root = parse_xml(z.read("xl/worksheets/sheet1.xml"))
    merges = root.findall(f".//{_MERGE}")
    assert len(merges) == 5


def test_conditional_formatting_semantic(tmp_path: Path):
    src = INSPECT_FIXTURES["ConditionalFormatting"]
    out = _save_roundtrip(src, tmp_path, edit=False)
    with zipfile.ZipFile(out) as z:
        root = parse_xml(z.read("xl/worksheets/sheet1.xml"))
    assert len(root.findall(f".//{_CF}")) == 5
    assert root.find(f"{{{SML_NS}}}extLst") is not None


def test_custom_cell_size_semantic(tmp_path: Path):
    src = INSPECT_FIXTURES["CustomCellSize"]
    out = _save_roundtrip(src, tmp_path, edit=False)
    with zipfile.ZipFile(out) as z:
        root = parse_xml(z.read("xl/worksheets/sheet1.xml"))
    cols = root.find(f"{{{SML_NS}}}cols")
    assert cols is not None
    assert cols[0].get("width") == "32"
    row4 = root.find(f".//{{{SML_NS}}}row[@r='4']")
    assert row4 is not None
    assert row4.get("ht") == "47"


def test_different_cell_types_values(tmp_path: Path):
    src = INSPECT_FIXTURES["DifferentCellTypes"]
    wb = Workbook.open(src)
    assert wb["Sheet1"]["B5"].value.day == 7
    assert wb["Sheet1"]["B6"].value.month == 8


def test_simple_formula_shared(tmp_path: Path):
    src = INSPECT_FIXTURES["SimpleFormula"]
    out = _save_roundtrip(src, tmp_path, edit=False)
    with zipfile.ZipFile(out) as z:
        root = parse_xml(z.read("xl/worksheets/sheet1.xml"))
    shared = [c for c in root.iter(_C) if (f := c.find(_F)) is not None and f.get("t") == "shared"]
    assert len(shared) >= 4
