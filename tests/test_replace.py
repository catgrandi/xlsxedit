from __future__ import annotations

import zipfile
from pathlib import Path

from xlsxedit import Workbook
from xlsxedit.opc.constants import SML_NS
from xlsxedit.oxml.parser import parse_xml


def test_workbook_replace_plain_sst(book1_path: Path, tmp_path: Path):
    out = tmp_path / "out.xlsx"
    wb = Workbook.open(book1_path)
    n = wb.replace("lol", "hi")
    assert n == 1
    wb.save(out)

    wb2 = Workbook.open(out)
    values = [c.value for c in wb2["test"].cells if c.is_string]
    assert "hi" in values
    assert "lol" not in values


def test_sheet_replace(book1_path: Path):
    wb = Workbook.open(book1_path)
    n = wb["Sheet1"].replace("this is a1", "ALPHA")
    assert n == 1
    assert wb["Sheet1"].cells[0].value == "ALPHA"


def test_replace_skips_numbers(book1_path: Path):
    wb = Workbook.open(book1_path)
    n = wb.replace("100", "999")
    assert n == 0
    assert {c.address: c.value for c in wb["Sheet1"].cells}["A5"] == 100


def test_rich_text_within_run_preserves_rpr(book1_path: Path, tmp_path: Path):
    out = tmp_path / "out.xlsx"
    wb = Workbook.open(book1_path)
    n = wb.replace("edited", "CHANGED")
    assert n == 1
    wb.save(out)

    with zipfile.ZipFile(out) as z:
        sst = parse_xml(z.read("xl/sharedStrings.xml"))

    sis = sst.findall(f"{{{SML_NS}}}si")
    rich = sis[2]
    runs = rich.findall(f"{{{SML_NS}}}r")
    bold_runs = [
        r
        for r in runs
        if (rPr := r.find(f"{{{SML_NS}}}rPr")) is not None
        and rPr.find(f"{{{SML_NS}}}b") is not None
    ]
    assert len(bold_runs) == 1
    assert bold_runs[0].find(f"{{{SML_NS}}}t").text == "CHANGED"

    # Adjacent runs unchanged
    assert runs[0].find(f"{{{SML_NS}}}t").text == "this is b2 and i "
    assert runs[2].find(f"{{{SML_NS}}}t").text == " it"


def test_rich_text_cross_run_not_replaced(book1_path: Path):
    wb = Workbook.open(book1_path)
    # Spans end of first run and bold run — must not match within a single <t>
    n = wb.replace("i edited", "X")
    assert n == 0
    b2 = next(c for c in wb["Sheet1"].cells if c.address == "B2")
    assert b2.value == "this is b2 and i edited it"


def test_empty_old_is_noop(book1_path: Path):
    wb = Workbook.open(book1_path)
    assert wb.replace("", "x") == 0
