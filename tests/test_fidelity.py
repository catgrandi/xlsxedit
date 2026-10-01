"""Fidelity of every registered fixture: consistency, a no-edit save, a single-cell edit.

Built on ``tests/preservation.py``; a fixture registered in ``INSPECT_FIXTURES`` is
covered automatically.
"""

from __future__ import annotations

import io

import pytest

from xlsxedit import Workbook
from xlsxedit.opc.constants import CT, SML_NS
from xlsxedit.oxml.parser import parse_xml
from tests.conftest import BOOK1, INSPECT_FIXTURES
from tests.preservation import (
    CONTENT_TYPES,
    CellState,
    PreservationReport,
    assert_preserved,
    check_consistency,
    content_types,
    read_pkg,
    worksheet_diff,
    worksheet_members,
)

FIXTURES = {"Book1": BOOK1, **INSPECT_FIXTURES}
PROBE = "Z99"

# Parts xlsxedit rebuilds on every save: its live XmlParts and all relationships.
# Every other member must keep its exact bytes, as StylesPart promises styles.xml.
REBUILT_ON_SAVE = frozenset(
    {
        CT.WORKBOOK,
        CT.WORKBOOK_MACRO_ENABLED,
        CT.WORKBOOK_TEMPLATE,
        CT.WORKBOOK_MACRO_ENABLED_TEMPLATE,
        CT.WORKSHEET,
        CT.SHARED_STRINGS,
        CT.OPC_CORE_PROPERTIES,
        CT.OPC_RELATIONSHIPS,
    }
)

_C = f"{{{SML_NS}}}c"
_MERGE = f"{{{SML_NS}}}mergeCell"
_CF = f"{{{SML_NS}}}conditionalFormatting"
_F = f"{{{SML_NS}}}f"


def _probe_first_sheet(name: str, value) -> tuple[dict[str, bytes], dict[str, bytes], str]:
    """Write ``value`` to ``PROBE`` on the first sheet; return before, after, sheet member."""
    before = read_pkg(FIXTURES[name])
    wb = Workbook.open(FIXTURES[name])
    sheet = wb.sheetnames[0]
    wb[sheet][PROBE].value = value
    out = io.BytesIO()
    wb.save(out)
    assert Workbook.open(io.BytesIO(out.getvalue()))[sheet][PROBE].value == value
    return before, read_pkg(out.getvalue()), worksheet_members(before)[sheet]


def _assert_bytes_kept(before: dict[str, bytes], report: PreservationReport) -> None:
    types = content_types(before)
    rewritten = {
        m for m in report.reserialised - {CONTENT_TYPES} if types[m] not in REBUILT_ON_SAVE
    }
    assert not rewritten, f"re-serialised parts that should keep their bytes: {sorted(rewritten)}"


def _saved_sheet(name: str, member: str = "xl/worksheets/sheet1.xml"):
    return parse_xml(read_pkg(Workbook.open(FIXTURES[name]))[member])


@pytest.mark.parametrize("name", FIXTURES)
def test_fixture_is_consistent(name: str):
    check_consistency(FIXTURES[name])


@pytest.mark.parametrize("name", FIXTURES)
def test_no_edit_save_changes_nothing(name: str):
    """Not even ``[Content_Types].xml``: rewriting its ``Default`` entries as
    ``Override``s keeps every part's effective content type."""
    before = read_pkg(FIXTURES[name])
    _assert_bytes_kept(before, assert_preserved(before, Workbook.open(FIXTURES[name])))


@pytest.mark.parametrize("name", FIXTURES)
def test_number_edit_changes_only_the_probed_sheet(name: str):
    before, after, member = _probe_first_sheet(name, 42)
    _assert_bytes_kept(before, assert_preserved(before, after, expected_changed={member}))
    diff = worksheet_diff(before[member], after[member])
    assert diff.added == {PROBE: CellState(t=None, s=None, formula=None, value="42")}
    assert diff.rows in ({}, {99: (None, {"r": "99"})})
    assert not (diff.removed or diff.changed or diff.children or diff.root or diff.order)


@pytest.mark.parametrize("name", FIXTURES)
def test_string_edit_changes_only_the_probed_sheet_and_shared_strings(name: str):
    sst = [
        m for m, ct in content_types(read_pkg(FIXTURES[name])).items() if ct == CT.SHARED_STRINGS
    ]
    if not sst:
        pytest.skip("no shared strings part to append to")
    before, after, member = _probe_first_sheet(name, "probe")
    _assert_bytes_kept(before, assert_preserved(before, after, expected_changed={member, *sst}))
    diff = worksheet_diff(before[member], after[member])
    assert list(diff.added) == [PROBE]
    assert diff.added[PROBE].t == "s"
    assert not (diff.removed or diff.changed or diff.children or diff.root or diff.order)


def test_merged_cells_semantic():
    assert len(_saved_sheet("MergedCells").findall(f".//{_MERGE}")) == 5


def test_conditional_formatting_semantic():
    root = _saved_sheet("ConditionalFormatting")
    assert len(root.findall(f".//{_CF}")) == 5
    assert root.find(f"{{{SML_NS}}}extLst") is not None


def test_custom_cell_size_semantic():
    root = _saved_sheet("CustomCellSize")
    cols = root.find(f"{{{SML_NS}}}cols")
    assert cols is not None
    assert cols[0].get("width") == "32"
    row4 = root.find(f".//{{{SML_NS}}}row[@r='4']")
    assert row4 is not None
    assert row4.get("ht") == "47"


def test_different_cell_types_values():
    wb = Workbook.open(FIXTURES["DifferentCellTypes"])
    assert wb["Sheet1"]["B5"].value.day == 7
    assert wb["Sheet1"]["B6"].value.month == 8


def test_simple_formula_shared():
    root = _saved_sheet("SimpleFormula")
    shared = [c for c in root.iter(_C) if (f := c.find(_F)) is not None and f.get("t") == "shared"]
    assert len(shared) >= 4
