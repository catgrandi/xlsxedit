"""``scripts/inspect_workbook.py --semantic`` compares workbooks by meaning, not bytes."""

from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest

from tests.conftest import FIXTURES
from tests.preservation import read_pkg
from xlsxedit import Workbook
from xlsxedit.opc.part import XmlPart

SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "inspect_workbook.py"
REFERENCE = FIXTURES / "excel-reference"


@pytest.fixture(scope="module")
def inspect():
    spec = importlib.util.spec_from_file_location("inspect_workbook", SCRIPT)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _reserialised(source: Path, target: Path) -> Path:
    """Open and save ``source`` with every XML part written again from its tree."""
    wb = Workbook.open(source)
    for part in wb._package.iter_parts():
        if isinstance(part, XmlPart):
            part.mark_dirty()
    wb.save(target)
    return target


@pytest.mark.parametrize(
    "fixture",
    [
        "Preservation.xlsm",
        "TableRich.xlsx",
        "ConditionalFormattingRich.xlsx",
        "DataValidation.xlsx",
        "DrawingAnchors.xlsx",
        "Formulas.xlsx",
    ],
)
def test_a_reserialised_copy_compares_equal(inspect, tmp_path, capsys, fixture):
    source = FIXTURES / fixture
    copy = _reserialised(source, tmp_path / fixture)
    before, after = read_pkg(source), read_pkg(copy)
    assert any(before[name] != after[name] for name in before), "the copy must differ in bytes"

    assert inspect.main([str(source), "--semantic", "--compare", str(copy)]) == 0
    out = capsys.readouterr().out
    assert f"0 of {len(inspect.semantic_aspects(source))} aspects differ" in out
    assert "re-serialised, canonically equal:" in out


@pytest.mark.parametrize(
    ("operation", "aspects"),
    [
        (
            "insert-rows-above-table",
            {
                "defined names",
                "tables",
                "conditional formats",
                "data validations",
                "merges",
                "formulas",
                "cells",
                "parts",
            },
        ),
        (
            "delete-rows-under-cf",
            {
                "defined names",
                "conditional formats",
                "data validations",
                "drawings",
                "merges",
                "cells",
                "parts",
            },
        ),
    ],
)
def test_a_change_in_content_is_reported_by_aspect(inspect, capsys, operation, aspects):
    before, after = REFERENCE / operation / "before.xlsx", REFERENCE / operation / "after.xlsx"
    differences = inspect.semantic_differences(
        inspect.semantic_aspects(before), inspect.semantic_aspects(after)
    )
    assert {aspect for aspect, items in differences.items() if items} == aspects

    assert inspect.main([str(before), "--semantic", "--compare", str(after)]) == 1
    out = capsys.readouterr().out
    assert f"{len(aspects)} of {len(differences)} aspects differ" in out
    assert "  sheets: same" in out


def test_an_edit_shows_where_it_landed(inspect, tmp_path, capsys):
    source = FIXTURES / "TableRich.xlsx"
    wb = Workbook.open(source)
    wb["Orders"]["B2"].value = 251
    edited = tmp_path / "edited.xlsx"
    wb.save(edited)

    assert inspect.main([str(source), "--semantic", "--compare", str(edited)]) == 1
    out = capsys.readouterr().out
    assert "    ~ Orders!B2\n        - n 250.0\n        + n 251.0\n" in out
    differences = inspect.semantic_differences(
        inspect.semantic_aspects(source), inspect.semantic_aspects(edited)
    )
    assert {aspect for aspect, items in differences.items() if items} == {"cells", "parts"}


def test_semantic_view_includes_extension_rules(inspect, capsys):
    assert inspect.main([str(FIXTURES / "ConditionalFormattingRich.xlsx"), "--semantic"]) == 0
    out = capsys.readouterr().out
    assert "  conditional formats: 5\n" in out
    assert " x14@" in out or "(x14)" in out


def test_byte_compare_still_lists_reserialised_parts(inspect, tmp_path, capsys):
    source = FIXTURES / "TableRich.xlsx"
    copy = _reserialised(source, tmp_path / "copy.xlsx")
    assert inspect.main([str(source), "--compare", str(copy)]) == 0
    assert "differing parts" in capsys.readouterr().out
