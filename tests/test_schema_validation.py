"""Saved packages validate against the vendored ECMA-376 transitional schemas."""

from __future__ import annotations

import io
import warnings
import zipfile
from collections.abc import Callable
from pathlib import Path

import pytest
from lxml import etree

from xlsxedit import Workbook
from xlsxedit.drawing import Chart
from xlsxedit.opc.constants import OFFICE_REL_NS, SML_NS
from xlsxedit.oxml.parser import parse_template_xml
from tests.conftest import ASSETS, FIXTURES, INSPECT_FIXTURES
from tests.preservation import check_consistency
from tests.schema_validate import (
    CHART_ALLOWLIST,
    SchemaWarning,
    apply_markup_compatibility,
    assert_valid_package,
    extension_namespaces,
    extension_payloads,
    schema_errors,
)
from tests.smoke_workbook import build_smoke_workbook

FIXTURE_PATHS = sorted(FIXTURES.glob("*.xlsx"))
TEMPLATES = Path(__file__).resolve().parents[1] / "src" / "xlsxedit" / "templates"
CF_FIXTURE = INSPECT_FIXTURES["ConditionalFormatting"]
CF_SHEET = "xl/worksheets/sheet1.xml"

_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_X14AC = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
_XM = "http://schemas.microsoft.com/office/excel/2006/main"
_XR = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"


def _ids(paths: list[Path]) -> list[str]:
    return [p.stem for p in paths]


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
def test_fixture_validates_after_insert_rows(path: Path):
    wb = Workbook.open(path)
    for ws in wb.worksheets:
        ws.insert_rows([["inserted", 1], ["inserted", 2]], at_row=2)
    checked = assert_valid_package(wb)
    check_consistency(wb)
    assert sum(kind == "spreadsheetml" for kind in checked.values()) >= 3


@pytest.mark.parametrize("path", FIXTURE_PATHS, ids=_ids(FIXTURE_PATHS))
def test_fixture_validates_after_insert_columns(path: Path):
    wb = Workbook.open(path)
    for ws in wb.worksheets:
        ws.insert_columns([["inserted", 1]], at_col="B")
    assert_valid_package(wb)
    check_consistency(wb)


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


def test_alternate_content_takes_the_first_understood_choice():
    root = etree.fromstring(
        f'<r xmlns:mc="{_MC}" xmlns:x14="{_X14}" xmlns:u="urn:unknown">'
        '<mc:AlternateContent><mc:Choice Requires="u"><c1/></mc:Choice>'
        '<mc:Choice Requires="x14"><c2/></mc:Choice><mc:Choice Requires="x14"><c3/></mc:Choice>'
        "<mc:Fallback><f1/></mc:Fallback></mc:AlternateContent></r>"
    )
    apply_markup_compatibility(root, frozenset({_X14}))
    assert [child.tag for child in root] == ["c2"]


def test_extension_namespaces_are_the_vendored_microsoft_ones():
    namespaces = extension_namespaces()
    assert {_X14, _XM, _X14AC, _XR} <= namespaces
    assert all(ns.startswith("http://schemas.microsoft.com/office/") for ns in namespaces)


def _with_member(path: Path, member: str, edit: Callable[[bytes], bytes]) -> bytes:
    """Return a copy of the package at ``path`` with ``member`` rewritten by ``edit``."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(buffer, "w") as target:
        for info in source.infolist():
            data = source.read(info)
            target.writestr(info, edit(data) if info.filename == member else data)
    return buffer.getvalue()


def test_x14_conditional_formatting_of_the_fixture_is_validated():
    with zipfile.ZipFile(CF_FIXTURE) as package:
        sheet = package.read(CF_SHEET)
    assert f"{{{_X14}}}conditionalFormattings" in extension_payloads(sheet)
    assert schema_errors(sheet) == []
    assert assert_valid_package(CF_FIXTURE)[CF_SHEET] == "spreadsheetml"


def test_corrupted_x14_conditional_formatting_fails():
    def corrupt(xml: bytes) -> bytes:
        assert xml.count(b"<x14:cfRule ") == 1
        return xml.replace(b"<x14:cfRule ", b"<x14:unknown/><x14:cfRule ")

    corrupted = _with_member(CF_FIXTURE, CF_SHEET, corrupt)
    with pytest.raises(
        AssertionError, match=r"xl/worksheets/sheet1\.xml line \d+: .*unknown': This element is not expected"
    ):
        assert_valid_package(corrupted)


def _styles_with_payload(payload: str, namespaces: str = "") -> bytes:
    return (
        f'<styleSheet xmlns="{SML_NS}" xmlns:x14="{_X14}" {namespaces}>'
        f'<extLst><ext uri="{{X}}">{payload}</ext></extLst></styleSheet>'
    ).encode()


def test_an_undeclared_element_in_a_vendored_namespace_fails():
    errors = schema_errors(_styles_with_payload("<x14:slicerStylez/>"))
    assert len(errors) == 1
    assert "slicerStylez': No matching global declaration" in errors[0]


def test_a_payload_inside_alternate_content_is_validated():
    payload = (
        '<mc:AlternateContent><mc:Choice Requires="x14"><x14:slicerStylez/></mc:Choice>'
        "<mc:Fallback><u:other/></mc:Fallback></mc:AlternateContent>"
    )
    styles = _styles_with_payload(payload, f'xmlns:mc="{_MC}" xmlns:u="urn:unknown"')
    assert extension_payloads(styles) == [f"{{{_X14}}}slicerStylez"]
    errors = schema_errors(styles)
    assert len(errors) == 1
    assert "slicerStylez': No matching global declaration" in errors[0]


def test_ignorable_content_inside_an_unknown_payload_is_kept():
    """libxml2 still validates declared elements inside a payload it skips."""
    styles = _styles_with_payload(
        "<u:a><u:b><styleSheet><bogus/></styleSheet></u:b></u:a>",
        f'xmlns:mc="{_MC}" xmlns:u="urn:unknown" mc:Ignorable="u"',
    )
    errors = schema_errors(styles)
    assert len(errors) == 1
    assert "bogus" in errors[0]


def test_payloads_in_other_namespaces_are_skipped():
    styles = _styles_with_payload("<u:any><u:thing/></u:any>", 'xmlns:u="urn:unknown"')
    assert extension_payloads(styles) == []
    assert schema_errors(styles) == []


def test_a_nested_payload_is_validated_with_its_outer_payload():
    rule = (
        '<x14:cfRule type="expression" id="{00000000-0000-0000-0000-000000000001}">'
        "<xm:f>A1&gt;0</xm:f>"
        '<x14:extLst><ext uri="{Y}"><x14:id>{00000000-0000-0000-0000-000000000002}</x14:id>'
        "</ext></x14:extLst></x14:cfRule>"
    )
    payload = (
        "<x14:conditionalFormattings><x14:conditionalFormatting>"
        f"{rule}<xm:sqref>A1</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings>"
    )
    styles = _styles_with_payload(payload, f'xmlns:xm="{_XM}"')
    assert extension_payloads(styles) == [f"{{{_X14}}}conditionalFormattings"]
    assert schema_errors(styles) == []
    errors = schema_errors(styles.replace(b"-000000000002}", b"}"))
    assert len(errors) == 1
    assert f"{{{_X14}}}id'" in errors[0]


def _sheet_with_x14_validation(uid: str) -> bytes:
    return (
        f'<worksheet xmlns="{SML_NS}" xmlns:mc="{_MC}" xmlns:x14ac="{_X14AC}" xmlns:xr="{_XR}"'
        ' mc:Ignorable="x14ac xr" xr:uid="{00000000-0000-0000-0000-000000000003}">'
        '<sheetData><row r="1" x14ac:dyDescent="0.25"/></sheetData>'
        f'<extLst><ext uri="{{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}}" xmlns:x14="{_X14}">'
        f'<x14:dataValidations count="1" xmlns:xm="{_XM}">'
        f'<x14:dataValidation type="list" xr:uid="{uid}">'
        "<x14:formula1><xm:f>Lists!$A$1:$A$5</xm:f></x14:formula1><xm:sqref>B1</xm:sqref>"
        "</x14:dataValidation></x14:dataValidations></ext></extLst></worksheet>"
    ).encode()


def test_ignorable_attributes_are_validated_on_extension_elements_only():
    """``xr:uid`` is part of the x14 type; on the base ``worksheet`` and
    ``row`` elements, extension attributes are dropped as before."""
    assert schema_errors(_sheet_with_x14_validation("{00000000-0000-0000-0000-000000000004}")) == []
    errors = schema_errors(_sheet_with_x14_validation("not-a-guid"))
    assert len(errors) == 1
    assert "revision}uid" in errors[0]


def test_sparkline_groups_are_validated():
    group = (
        '<x14:sparklineGroups><x14:sparklineGroup displayEmptyCellsAs="gap">'
        '<x14:colorSeries rgb="FF376092"/>{}</x14:sparklineGroup></x14:sparklineGroups>'
    )
    sparklines = (
        "<x14:sparklines><x14:sparkline><xm:f>Sheet1!A1:E1</xm:f><xm:sqref>F1</xm:sqref>"
        "</x14:sparkline></x14:sparklines>"
    )
    styles = _styles_with_payload(group.format(sparklines), f'xmlns:xm="{_XM}"')
    assert extension_payloads(styles) == [f"{{{_X14}}}sparklineGroups"]
    assert schema_errors(styles) == []
    errors = schema_errors(_styles_with_payload(group.format(""), f'xmlns:xm="{_XM}"'))
    assert len(errors) == 1
    assert "Missing child element(s)" in errors[0]


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
