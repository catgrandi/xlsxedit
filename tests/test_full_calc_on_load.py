"""fullCalcOnLoad after replace and via set_full_calc_on_load."""

from __future__ import annotations

from pathlib import Path

from lxml import etree

from xlsxedit import Workbook
from xlsxedit.cell import _ensure_v
from xlsxedit.opc.constants import SML_NS

_CALC_PR = f"{{{SML_NS}}}calcPr"


def _calc_pr(wb) -> etree._Element | None:
    return wb._workbook_part.element.find(_CALC_PR)


def _full_calc_on_load(wb) -> str | None:
    calc_pr = _calc_pr(wb)
    return None if calc_pr is None else calc_pr.get("fullCalcOnLoad")


def test_workbook_replace_sets_full_calc_on_load(tmp_path: Path):
    wb = Workbook.create()
    wb["Sheet1"]["A1"].value = "{qty}"
    assert _calc_pr(wb) is None

    count = wb.replace("{qty}", 21, value_type="number")
    assert count == 1
    assert _full_calc_on_load(wb) == "1"

    out = tmp_path / "out.xlsx"
    wb.save(out)
    wb2 = Workbook.open(out)
    assert _full_calc_on_load(wb2) == "1"


def test_worksheet_replace_sets_full_calc_on_load():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "hello"

    count = ws.replace("hello", "world")
    assert count == 1
    assert _full_calc_on_load(wb) == "1"


def test_replace_noop_does_not_add_calc_pr():
    wb = Workbook.create()
    wb["Sheet1"]["A1"].value = "hello"

    count = wb.replace("missing", "world")
    assert count == 0
    assert _calc_pr(wb) is None


def test_set_full_calc_on_load_idempotent_and_preserves_attrs():
    wb = Workbook.create()
    wb_elm = wb._workbook_part.element
    calc_pr = etree.Element(_CALC_PR)
    calc_pr.set("calcId", "191029")
    calc_pr.set("calcMode", "auto")
    wb_elm.append(calc_pr)

    wb.set_full_calc_on_load()
    wb.set_full_calc_on_load()

    found = _calc_pr(wb)
    assert found is not None
    assert found.get("fullCalcOnLoad") == "1"
    assert found.get("calcId") == "191029"
    assert found.get("calcMode") == "auto"
    assert len(wb_elm.findall(_CALC_PR)) == 1


def test_replace_leaves_stale_formula_cache():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "{qty}"
    ws["C1"].formula = "=A1*2"
    _ensure_v(ws["C1"]._element).text = "0"

    wb.replace("{qty}", 21, value_type="number")
    assert ws["A1"].value == 21
    assert ws["C1"].formula == "A1*2"
    assert ws["C1"].value == 0
    assert _full_calc_on_load(wb) == "1"
