"""Saved packages validate against the vendored ECMA-376 transitional schemas."""

from __future__ import annotations

import io
import warnings
from pathlib import Path

import pytest
from lxml import etree

from xlsxedit import Workbook
from xlsxedit.drawing import Chart
from xlsxedit.opc.constants import OFFICE_REL_NS, SML_NS
from xlsxedit.oxml.parser import parse_template_xml
from tests.conftest import ASSETS, FIXTURES
from tests.preservation import check_consistency
from tests.schema_validate import (
    CHART_ALLOWLIST,
    SchemaWarning,
    apply_markup_compatibility,
    assert_valid_package,
    schema_errors,
)
from tests.smoke_workbook import build_smoke_workbook

FIXTURE_PATHS = sorted(FIXTURES.glob("*.xlsx"))
TEMPLATES = Path(__file__).resolve().parents[1] / "src" / "xlsxedit" / "templates"

_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_X14AC = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"


def _ids(paths: list[Path]) -> list[str]:
    return [p.stem for p in paths]


# insert_rows and insert_columns still break these consistency invariants on
# some fixtures (issues #5 and #7). The two tests below assert schema validity,
# so they opt out of the suite guard and tolerate exactly these codes; every
# other invariant must still hold. The strict xfails in test_consistency.py
# report when the bugs are fixed. Remove a code here when its issue lands.
_KNOWN_STRUCTURAL_BUGS = frozenset({"shared-formula", "calc-chain", "x14-cf", "table-columns"})


def _assert_consistent_apart_from_known_bugs(wb: Workbook) -> None:
    buffer = io.BytesIO()
    wb.save(buffer)
    check_consistency(buffer.getvalue(), waive=_KNOWN_STRUCTURAL_BUGS)


def test_fixtures_are_found():
    assert len(FIXTURE_PATHS) >= 11


@pytest.mark.parametrize("path", FIXTURE_PATHS, ids=_ids(FIXTURE_PATHS))
def test_fixture_validates(path: Path):
    with warnings.catch_warnings():
        warnings.simplefilter("error", SchemaWarning)
        checked = assert_valid_package(path)
    assert "xl/workbook.xml" in checked
    assert "xl/styles.xml" in checked


@pytest.mark.parametrize("path", FIXTURE_PATHS, ids=_ids(FIXTURE_PATHS))
def test_fixture_validates_after_insert_rows(path: Path, unchecked_workbooks):
    wb = Workbook.open(path)
    for ws in wb.worksheets:
        ws.insert_rows([["inserted", 1], ["inserted", 2]], at_row=2)
    checked = assert_valid_package(wb)
    _assert_consistent_apart_from_known_bugs(wb)
    assert sum(kind == "spreadsheetml" for kind in checked.values()) >= 3


@pytest.mark.parametrize("path", FIXTURE_PATHS, ids=_ids(FIXTURE_PATHS))
def test_fixture_validates_after_insert_columns(path: Path, unchecked_workbooks):
    wb = Workbook.open(path)
    for ws in wb.worksheets:
        ws.insert_columns([["inserted", 1]], at_col="B")
    assert_valid_package(wb)
    _assert_consistent_apart_from_known_bugs(wb)


@pytest.mark.parametrize("path", FIXTURE_PATHS, ids=_ids(FIXTURE_PATHS))
def test_fixture_validates_after_copy_worksheet(path: Path):
    wb = Workbook.open(path)
    source = wb.sheetnames[0]
    wb.copy_worksheet(source, f"{source[:20]} copy")
    assert_valid_package(wb)


def test_smoke_workbook_validates():
    checked = assert_valid_package(build_smoke_workbook())
    assert set(checked.values()) == {"spreadsheetml", "drawing", "theme", "chart"}


def test_sheet_with_legacy_drawing_validates_after_edits():
    wb = Workbook.create()
    ws = wb["Sheet1"]
    legacy = etree.SubElement(ws._part.element, f"{{{SML_NS}}}legacyDrawing")
    legacy.set(f"{{{OFFICE_REL_NS}}}id", "rId90")
    ws["A1"].value = "Day"
    ws["B1"].value = "N"
    ws["A2"].value = "Mon"
    ws["B2"].value = 1
    ws.merge_cells("C1:D1")
    ws.add_conditional_formatting("B2", operator="greaterThan", formula="0")
    ws["A5"].value = "link"
    ws["A5"].hyperlink.url = "https://example.com"
    ws.add_image(ASSETS / "coco-happy-swiss-nature.jpg", anchor="F2")
    ws.add_table("A1:B2", ["Day", "N"], name="Days")
    assert_valid_package(wb)


def test_out_of_order_worksheet_fails():
    wb = Workbook.create()
    root = wb["Sheet1"]._part.element
    root.append(root.find(f"{{{SML_NS}}}sheetViews"))
    with pytest.raises(AssertionError, match=r"xl/worksheets/sheet1\.xml line \d+: .*sheetViews"):
        assert_valid_package(wb)


def test_alternate_content_resolves_to_fallback_in_place():
    root = etree.fromstring(
        f'<r xmlns:mc="{_MC}"><a/>'
        '<mc:AlternateContent><mc:Choice Requires="x"><c1/></mc:Choice>'
        "<mc:Fallback><f1/><f2/></mc:Fallback></mc:AlternateContent>"
        '<mc:AlternateContent><mc:Choice Requires="x"><c2/></mc:Choice></mc:AlternateContent>'
        "<z/></r>"
    )
    apply_markup_compatibility(root)
    assert [child.tag for child in root] == ["a", "f1", "f2", "z"]


def test_ignorable_content_is_dropped_outside_extensions_only():
    root = etree.fromstring(
        f'<r xmlns:mc="{_MC}" xmlns:x14ac="{_X14AC}" xmlns:x14="{_X14}"'
        ' mc:Ignorable="x14ac x14" xml:space="preserve">'
        '<row x14ac:dyDescent="0.25"/>'
        "<x14ac:note/>"
        '<extLst><ext uri="{X}"><x14:kept x14ac:attr="1"/></ext></extLst>'
        "</r>"
    )
    apply_markup_compatibility(root)
    assert dict(root.attrib) == {}
    assert dict(root.find("row").attrib) == {}
    assert [etree.QName(child).localname for child in root] == ["row", "extLst"]
    kept = root.find("extLst/ext")[0]
    assert etree.QName(kept).localname == "kept"
    assert dict(kept.attrib) == {}


def test_ignorable_extension_content_keeps_ext_valid():
    """Stripping x14 inside ``ext`` would leave it without its required child."""
    styles = (
        f'<styleSheet xmlns="{SML_NS}" xmlns:mc="{_MC}" xmlns:x14="{_X14}"'
        ' mc:Ignorable="x14">'
        '<extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}">'
        '<x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>'
        "</ext></extLst></styleSheet>"
    ).encode()
    assert schema_errors(styles) == []


def _workbook_with_chart() -> tuple[Workbook, Chart]:
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "Day"
    ws["B1"].value = "N"
    ws["A2"].value = "Mon"
    ws["B2"].value = 1
    return wb, ws.add_chart("bar", anchor="D2", data_range="A1:B2", title="T")


def test_known_chart_violation_is_rewritten_not_reported():
    template = parse_template_xml((TEMPLATES / "default-bar-chart.xml").read_bytes())
    errors = schema_errors(template)
    assert len(errors) == 1
    assert "showDLblsOverMax" in errors[0]
    for _reason, rewrite in CHART_ALLOWLIST:
        rewrite(template)
    assert schema_errors(template) == []

    wb, _chart = _workbook_with_chart()
    with warnings.catch_warnings():
        warnings.simplefilter("error", SchemaWarning)
        checked = assert_valid_package(wb)
    assert checked["xl/charts/chart1.xml"] == "chart"


def test_unexpected_chart_violation_warns():
    wb, chart = _workbook_with_chart()
    root = chart._chart_root()
    chart_elm = root.find(f"{{{_C}}}chart")
    chart_elm.remove(chart_elm.find(f"{{{_C}}}plotArea"))
    chart._save_chart_root(root)
    with pytest.warns(SchemaWarning, match=r"xl/charts/chart1\.xml line \d+: .*plotArea"):
        assert_valid_package(wb)


def test_chart_error_after_the_known_violation_still_warns():
    """libxml2 reports one content-model error per element, so the known one
    must not stand in for a later ``c:chart`` error."""
    wb, chart = _workbook_with_chart()
    root = chart._chart_root()
    etree.SubElement(root.find(f"{{{_C}}}chart"), f"{{{_C}}}legend")
    chart._save_chart_root(root)
    with pytest.warns(SchemaWarning, match=r"xl/charts/chart1\.xml line \d+: .*legend"):
        assert_valid_package(wb)
