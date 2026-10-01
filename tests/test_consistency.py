"""``check_consistency``: what each invariant catches, and the known bugs that break one.

The first tests corrupt one fixture member at a time and pin the violation it
produces. The ``_known_bug`` tests at the end apply an operation that an open
issue documents as breaking an invariant: each is a strict xfail that also
waives exactly the invariants it breaks in the suite guard. The issue named
there turns those invariants on; when it lands, the test XPASSes, then its
waiver goes stale, so delete the ``_known_bug`` line.
"""

from __future__ import annotations

import io
import zipfile

import pytest
from lxml import etree

from xlsxedit import Workbook
from xlsxedit.opc.constants import SML_NS
from tests.conftest import BOOK1, INSPECT_FIXTURES
from tests.preservation import (
    check_consistency,
    consistency_violations,
    read_pkg,
    waives,
    worksheet_members,
)

SHEET = "xl/worksheets/sheet1.xml"
TABLE = "xl/tables/table1.xml"
TABLE_SHEET = "xl/worksheets/sheet3.xml"
_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _pkg(name: str) -> dict[str, bytes]:
    return dict(read_pkg(INSPECT_FIXTURES[name]))


def _edit(
    pkg: dict[str, bytes], member: str, old: bytes, new: bytes, count: int = 1
) -> dict[str, bytes]:
    """Replace ``old`` in a member; LookupError (never an AssertionError a
    ``_known_bug`` xfail would absorb) when the fixture text has moved."""
    found = pkg[member].count(old)
    if found != count:
        raise LookupError(f"{old!r} occurs {found} times in {member}, expected {count}")
    pkg[member] = pkg[member].replace(old, new)
    return pkg


def _add_rel(pkg: dict[str, bytes], rels: str, r_id: str, kind: str, target: str) -> None:
    rel = f'<Relationship Id="{r_id}" Type="{_RT}/{kind}" Target="{target}"/>'
    _edit(pkg, rels, b"</Relationships>", rel.encode() + b"</Relationships>")


def _codes(pkg) -> list[str]:
    return sorted(v.code for v in consistency_violations(pkg))


def _only(pkg, code: str) -> str:
    """Assert ``pkg`` breaks exactly one invariant, ``code``, once; return the detail."""
    violations = consistency_violations(pkg)
    assert [v.code for v in violations] == [code], violations
    return violations[0].detail


def _zip(pkg: dict[str, bytes]) -> io.BytesIO:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        for member, blob in pkg.items():
            z.writestr(member, blob)
    buf.seek(0)
    return buf


def _known_bug(issue: str, *codes: str, reason: str):
    """Strict xfail on the ``codes`` violations ``issue`` documents, waived in the suite guard."""
    xfail = pytest.mark.xfail(strict=True, raises=AssertionError, reason=f"{issue}: {reason}")
    return lambda test: xfail(waives(issue, *codes)(test))


def _chain(*entries: str) -> bytes:
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        b'<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        + "".join(entries).encode()
        + b"</calcChain>"
    )


def test_shared_formula_member_outside_the_master_ref():
    pkg = _edit(_pkg("SimpleFormula"), SHEET, b'ref="C4:C7"', b'ref="C4:C6"')
    assert "C7 outside master C4's ref C4:C6" in _only(pkg, "shared-formula")


def test_shared_formula_follower_with_text_is_a_second_master():
    pkg = _edit(
        _pkg("SimpleFormula"),
        SHEET,
        b'<c r="C5"><f t="shared" si="0"/>',
        b'<c r="C5"><f t="shared" si="0">A5*B5</f>',
    )
    assert "2 definitions (C4, C5)" in _only(pkg, "shared-formula")


def test_shared_formula_master_without_ref():
    pkg = _edit(
        _pkg("SimpleFormula"), SHEET, b'<f t="shared" ref="C4:C7" si="0">', b'<f t="shared" si="0">'
    )
    assert "master C4 has no ref" in _only(pkg, "shared-formula")


def test_shared_formula_without_si():
    pkg = _edit(
        _pkg("SimpleFormula"),
        SHEET,
        b'<c r="C5"><f t="shared" si="0"/>',
        b'<c r="C5"><f t="shared"/>',
    )
    assert "C5: shared formula without si" in _only(pkg, "shared-formula")


@pytest.mark.parametrize(
    ("f", "ok"),
    [
        (b'<f t="array" ref="C3:C4">A3*B3</f>', True),
        (b'<f t="array" ref="D3:D4">A3*B3</f>', False),
        (b'<f t="array">A3*B3</f>', False),
        (b'<f t="dataTable" ref="C3:D9" dt2D="0" dtr="0" r1="A1"/>', True),
        (b'<f t="dataTable" ref="D3:E9" dt2D="0" dtr="0" r1="A1"/>', False),
    ],
)
def test_array_and_data_table_masters_lie_inside_their_ref(f: bytes, ok: bool):
    pkg = _edit(_pkg("SimpleFormula"), SHEET, b"<f>A3*B3</f>", f)
    if ok:
        assert _codes(pkg) == []
    else:
        assert "C3: " in _only(pkg, "formula-ref")


def test_calc_chain_entries_inherit_the_previous_sheet_id():
    pkg = _pkg("SimpleFormula")
    pkg["xl/calcChain.xml"] = _chain('<c r="C3" i="1"/>', '<c r="C4"/>', '<c r="C8"/>')
    assert _codes(pkg) == []


def test_calc_chain_entry_without_a_formula():
    pkg = _pkg("SimpleFormula")
    pkg["xl/calcChain.xml"] = _chain('<c r="C3" i="1"/>', '<c r="A3"/>', '<c r="$B$3"/>')
    assert _only(pkg, "calc-chain") == "entries name cells without <f> on 'Sheet1': A3, B3"


def test_calc_chain_entry_on_a_missing_sheet():
    pkg = _pkg("SimpleFormula")
    pkg["xl/calcChain.xml"] = _chain('<c r="C3" i="1"/>', '<c r="C3" i="7"/>')
    assert "name sheetId 7, which no worksheet has" in _only(pkg, "calc-chain")


@pytest.mark.parametrize(
    ("old", "new", "what"),
    [
        (b'<c r="A3">', b'<c r="XFE3">', "cell XFE3"),
        (b'<row r="8" ', b'<row r="1048577" ', "row 1048577"),
        (
            b'<dimension ref="A2:C8"/>',
            b'<dimension ref="A2:C1048577"/>',
            "dimension/@ref 'A2:C1048577'",
        ),
        (b'sqref="D7"', b'sqref="D7 #REF!"', "selection/@sqref '#REF!'"),
        (
            b"<sheetData>",
            b'<cols><col min="1" max="16385"/></cols><sheetData>',
            "col min='1' max='16385'",
        ),
    ],
)
def test_grid_bounds_in_worksheets(old: bytes, new: bytes, what: str):
    pkg = _edit(_pkg("SimpleFormula"), SHEET, old, new)
    assert what in _only(pkg, "grid-bounds")


def test_grid_bounds_in_tables():
    pkg = _edit(_pkg("ChartsAndTables"), TABLE, b'ref="A1:C11"', b'ref="A1:C1048577"', count=2)
    assert "table/@ref 'A1:C1048577'" in _only(pkg, "grid-bounds")


def test_table_columns_count_attribute():
    pkg = _edit(
        _pkg("ChartsAndTables"), TABLE, b'<tableColumns count="3">', b'<tableColumns count="4">'
    )
    assert _only(pkg, "table-columns") == "tableColumns/@count=4 but 3 tableColumn children"


def test_table_columns_match_the_ref_width():
    pkg = _edit(_pkg("ChartsAndTables"), TABLE, b'ref="A1:C11"', b'ref="A1:D11"', count=2)
    assert _only(pkg, "table-columns") == "3 tableColumn children but ref A1:D11 is 4 wide"


def _second_table(display_name: bytes) -> dict[str, bytes]:
    pkg = _pkg("ChartsAndTables")
    pkg["xl/tables/table2.xml"] = (
        pkg[TABLE]
        .replace(
            b'name="Table2" displayName="Table2"',
            b'name="Other" displayName="' + display_name + b'"',
        )
        .replace(b"A1:C11", b"E1:G11")
    )
    _edit(
        pkg,
        TABLE_SHEET,
        b'<tableParts count="1"><tablePart r:id="rId1"/></tableParts>',
        b'<tableParts count="2"><tablePart r:id="rId1"/><tablePart r:id="rId2"/></tableParts>',
    )
    _add_rel(pkg, "xl/worksheets/_rels/sheet3.xml.rels", "rId2", "table", "../tables/table2.xml")
    return pkg


def test_table_display_names_are_unique_case_insensitively():
    assert _codes(_second_table(b"Table3")) == []
    detail = _only(_second_table(b"TABLE2"), "table-name")
    assert "Table2 (xl/tables/table1.xml), TABLE2 (xl/tables/table2.xml)" in detail


@pytest.mark.parametrize(
    ("totals", "auto_filter", "ok"),
    [
        (b'totalsRowShown="0"', b"A1:C11", True),
        (b'totalsRowShown="0"', b"A1:C10", False),
        (b'totalsRowCount="1"', b"A1:C10", True),
        (b'totalsRowCount="1"', b"A1:C11", False),
    ],
)
def test_table_auto_filter_excludes_the_totals_rows(totals: bytes, auto_filter: bytes, ok: bool):
    pkg = _edit(_pkg("ChartsAndTables"), TABLE, b'totalsRowShown="0"', totals)
    _edit(pkg, TABLE, b'<autoFilter ref="A1:C11"', b'<autoFilter ref="' + auto_filter + b'"')
    if ok:
        assert _codes(pkg) == []
    else:
        assert "but table ref A1:C11 minus" in _only(pkg, "table-autofilter")


_METADATA = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    b'<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    b'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
    b'<metadataTypes count="1">'
    b'<metadataType name="XLDAPR" minSupportedVersion="120000"/></metadataTypes>'
    b'<futureMetadata name="XLDAPR" count="1"><bk><extLst>'
    b'<ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
    b'<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk>'
    b"</futureMetadata>"
    b'<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>'
)


def _with_metadata(cell_attrs: bytes) -> dict[str, bytes]:
    pkg = _edit(_pkg("SimpleFormula"), SHEET, b'<c r="C3">', b'<c r="C3" ' + cell_attrs + b">")
    pkg["xl/metadata.xml"] = _METADATA
    _add_rel(pkg, "xl/_rels/workbook.xml.rels", "rId9", "sheetMetadata", "metadata.xml")
    return pkg


def test_cell_metadata_indexes_resolve():
    assert _codes(_with_metadata(b'cm="1"')) == []
    assert (
        _only(_with_metadata(b'cm="2"'), "cell-metadata")
        == "unresolved indexes: C3 cm=2 (metadata has 1)"
    )
    assert (
        _only(_with_metadata(b'vm="1"'), "cell-metadata")
        == "unresolved indexes: C3 vm=1 (metadata has 0)"
    )


def test_cell_metadata_without_a_metadata_part():
    pkg = _edit(_pkg("SimpleFormula"), SHEET, b'<c r="C3">', b'<c r="C3" cm="1">')
    assert _only(pkg, "cell-metadata") == "unresolved indexes: C3 cm=1 (no metadata part)"


_X14_ID = b"{BAAC1153-E264-9646-9BE7-9455373CB1E6}"


def test_x14_cf_twin_must_share_the_base_sqref():
    pkg = _edit(
        _pkg("ConditionalFormatting"),
        SHEET,
        b"<xm:sqref>C2:C7</xm:sqref>",
        b"<xm:sqref>C2:C8</xm:sqref>",
    )
    assert "xm:sqref differs from base sqref 'C2:C7'" in _only(pkg, "x14-cf")


def test_x14_cf_rule_id_must_have_a_twin():
    pkg = _edit(
        _pkg("ConditionalFormatting"),
        SHEET,
        b"<x14:id>" + _X14_ID,
        b"<x14:id>{00000000-0000-0000-0000-000000000000}",
    )
    assert "has no x14:cfRule twin" in _only(pkg, "x14-cf")


def test_x14_cf_matching_ignores_guid_case_and_area_order():
    pkg = _edit(_pkg("ConditionalFormatting"), SHEET, b'id="' + _X14_ID, b'id="' + _X14_ID.lower())
    _edit(
        pkg,
        SHEET,
        b'<conditionalFormatting sqref="C2:C7">',
        b'<conditionalFormatting sqref="C2:C7 $E$9">',
    )
    _edit(pkg, SHEET, b"<xm:sqref>C2:C7</xm:sqref>", b"<xm:sqref>E9 C2:C7</xm:sqref>")
    assert _codes(pkg) == []


def test_relationship_to_a_missing_part():
    pkg = dict(read_pkg(BOOK1))
    del pkg["xl/worksheets/sheet2.xml"]
    detail = _only(pkg, "relationship")
    assert detail == "rId2 (worksheet) targets missing part 'xl/worksheets/sheet2.xml'"


def test_every_relationships_part_is_checked():
    pkg = _pkg("ChartsAndTables")
    del pkg["xl/drawings/drawing1.xml"]  # its own rels part is then ignored
    detail = _only(pkg, "relationship")
    assert detail == "rId1 (drawing) targets missing part 'xl/drawings/drawing1.xml'"


@pytest.mark.parametrize(
    "target",
    [
        b"../../xl/worksheets/sheet2.xml",
        b"/xl/../xl/worksheets/sheet2.xml",
        b"worksheets/sheet%32.xml",
    ],
    ids=["above-root", "absolute", "percent-encoded"],
)
def test_relationship_targets_resolve_like_the_library(target: bytes):
    pkg = _edit(
        dict(read_pkg(BOOK1)),
        "xl/_rels/workbook.xml.rels",
        b'Target="worksheets/sheet2.xml"',
        b'Target="' + target + b'"',
    )
    assert _codes(pkg) == []


def test_table_part_reference_without_a_relationship():
    pkg = _edit(
        _pkg("ChartsAndTables"),
        TABLE_SHEET,
        b'<tablePart r:id="rId1"/>',
        b'<tablePart r:id="rId7"/>',
    )
    assert "tablePart r:id 'rId7' is not a table relationship" in _only(pkg, "relationship")


def test_check_consistency_raises_listing_every_violation():
    pkg = _edit(_pkg("SimpleFormula"), SHEET, b'ref="C4:C7"', b'ref="C4:C6"')
    pkg["xl/calcChain.xml"] = _chain('<c r="A3" i="1"/>')
    with pytest.raises(AssertionError) as excinfo:
        check_consistency(pkg)
    message = str(excinfo.value)
    assert "violates 2 consistency invariant(s)" in message
    assert "[shared-formula] xl/worksheets/sheet1.xml: si=0:" in message
    assert "[calc-chain] xl/calcChain.xml:" in message


def test_check_consistency_returns_waived_violations():
    pkg = _edit(_pkg("SimpleFormula"), SHEET, b'ref="C4:C7"', b'ref="C4:C6"')
    assert [v.code for v in check_consistency(pkg, waive={"shared-formula"})] == ["shared-formula"]
    with pytest.raises(ValueError, match="unknown invariant codes: shared-formulas"):
        check_consistency(pkg, waive={"shared-formulas"})


def test_worksheet_members_follow_workbook_order():
    assert worksheet_members(INSPECT_FIXTURES["ChartsAndTables"]) == {
        "bar chart": "xl/worksheets/sheet1.xml",
        "line chart": "xl/worksheets/sheet2.xml",
        "Table": TABLE_SHEET,
    }


@_known_bug(
    "#5",
    "shared-formula",
    "calc-chain",
    reason="insert_rows leaves shared refs and calcChain stale",
)
def test_insert_rows_keeps_formulas_and_calc_chain_consistent():
    wb = Workbook.open(INSPECT_FIXTURES["SimpleFormula"])
    wb["Sheet1"].insert_rows([[None]], at_row=5)
    check_consistency(wb)


@_known_bug("#5", "x14-cf", reason="insert_rows never shifts x14 xm:sqref")
def test_insert_rows_moves_x14_conditional_formatting():
    wb = Workbook.open(INSPECT_FIXTURES["ConditionalFormatting"])
    wb["Sheet1"].insert_rows([[None]], at_row=3)
    check_consistency(wb)


@_known_bug("#5", "grid-bounds", reason="insert_rows pushes cells past row 1048576")
def test_insert_rows_stays_inside_the_grid():
    wb = Workbook.create()
    ws = wb.worksheets[0]
    ws["A1048576"].value = 1
    ws.insert_rows([[None]], at_row=1)
    check_consistency(wb)


@_known_bug("#5", "formula-ref", reason="insert_rows never shifts array <f ref>")
def test_insert_rows_moves_array_formula_refs():
    wb = Workbook.create()
    ws = wb.worksheets[0]
    f = etree.SubElement(ws["D2"]._element, f"{{{SML_NS}}}f", t="array", ref="D2:D3")
    f.text = "A2:A3*2"
    ws.insert_rows([[None], [None]], at_row=1)
    check_consistency(wb)


def test_write_dataframe_leaves_other_tables_alone():
    wb = Workbook.open(INSPECT_FIXTURES["ChartsAndTables"])
    wb.write_dataframe([("a", 1), ("b", 2)], sheet="Table", at_cell="F1", header=False)
    check_consistency(wb)


@_known_bug("#7", "table-columns", reason="insert_columns widens a table without a tableColumn")
def test_insert_columns_inside_a_table_adds_a_table_column():
    wb = Workbook.open(INSPECT_FIXTURES["ChartsAndTables"])
    wb["Table"].insert_columns([[None]], at_col="B")
    check_consistency(wb)


@_known_bug("#7", "table-name", reason="add_table reuses displayName Table1")
def test_add_table_names_each_table_uniquely():
    wb = Workbook.create()
    ws = wb.worksheets[0]
    for cell, header in zip(("A1", "B1", "D1", "E1"), ("x", "y", "p", "q")):
        ws[cell].value = header
    ws.add_table("A1:B3", ["x", "y"])
    ws.add_table("D1:E3", ["p", "q"])
    check_consistency(wb)


@_known_bug("#8", "calc-chain", reason="remove_worksheet keeps calcChain entries of the sheet")
def test_remove_worksheet_drops_its_calc_chain_entries():
    wb = Workbook.open(INSPECT_FIXTURES["SimpleFormula"])
    wb.add_worksheet("Other")
    wb.remove_worksheet("Sheet1")
    check_consistency(wb)


@_known_bug("#13", "table-autofilter", reason="Table.resize puts the totals row inside autoFilter")
def test_table_resize_keeps_the_totals_row_out_of_autofilter():
    pkg = _edit(_pkg("ChartsAndTables"), TABLE, b'totalsRowShown="0"', b'totalsRowCount="1"')
    _edit(pkg, TABLE, b'<autoFilter ref="A1:C11"', b'<autoFilter ref="A1:C10"')
    wb = Workbook.open(_zip(pkg))
    wb["Table"].tables[0].resize("A1:C12")
    check_consistency(wb)
