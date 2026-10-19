"""Chart references follow the cells they plot on insert (issue #12)."""

from __future__ import annotations

import io
from copy import deepcopy

import pytest
from lxml import etree

from xlsxedit import Workbook
from xlsxedit.drawing import A_NS, C15_NS, CHART_NS, XDR_NS, chart_parts
from xlsxedit.exceptions import DTDForbiddenError, GridOverflowError
from xlsxedit.opc.constants import CT, OFFICE_REL_NS, RT, SML_NS
from xlsxedit.opc.packuri import PackURI
from xlsxedit.opc.part import Part
from tests.conftest import INSPECT_FIXTURES
from tests.preservation import assert_preserved, read_pkg
from tests.schema_validate import assert_valid_package
from tests.test_insert_rows import _defined_name_text

NS = {"c": CHART_NS, "c15": C15_NS}


def _chart(wb, partname: str) -> etree._Element:
    return next(part.element for part in chart_parts(wb._package) if part.partname == partname)


def _refs(wb, partname: str) -> list[str]:
    root = _chart(wb, partname)
    return [elm.text for elm in root.iter(f"{{{CHART_NS}}}f", f"{{{C15_NS}}}f", f"{{{C15_NS}}}sqref")]


def test_a_chart_follows_the_cells_it_plots():
    wb = Workbook.open(INSPECT_FIXTURES["ChartsAndTables"])
    before = read_pkg(wb)
    wb["bar chart"].insert_rows([[None], [None]], at_row=4)
    assert _refs(wb, "/xl/charts/chart1.xml") == [
        "'bar chart'!$B$2",
        "'bar chart'!$A$3:$A$8",
        "'bar chart'!$B$3:$B$8",
    ]
    assert_preserved(
        before,
        read_pkg(wb),
        expected_changed={"xl/worksheets/sheet1.xml", "xl/charts/chart1.xml", "xl/workbook.xml"},
    )
    assert_valid_package(wb)


def test_a_chart_and_its_hidden_names_stay_in_step():
    wb = Workbook.open(INSPECT_FIXTURES["ChartsAndTables"])
    wb["line chart"].insert_columns([[None]], at_col="A")
    assert _refs(wb, "/xl/charts/chart2.xml") == [
        "'line chart'!$C$2",
        "'line chart'!$B$2:$B$12",
        "'line chart'!$C$3:$C$12",
    ]
    assert _defined_name_text(wb, "_xlchart.v1.0") == "'line chart'!$B$2"
    assert _defined_name_text(wb, "_xlchart.v1.3") == "'line chart'!$C$3:$C$12"
    assert _refs(wb, "/xl/charts/chart1.xml")[0] == "'bar chart'!$B$2"


_C15_COPIES = [
    "<c15:fullRef><c15:sqref>'bar chart'!$B$3:$B$9</c15:sqref></c15:fullRef>",
    "<c15:datalabelsRange><c15:f>'bar chart'!$C$3:$C$6</c15:f></c15:datalabelsRange>",
    "<c15:filteredSeriesTitle><c15:tx><c:strRef><c:f>'bar chart'!$D$2</c:f></c:strRef></c15:tx>"
    "</c15:filteredSeriesTitle>",
]


def test_the_c15_copies_of_chart_references_move():
    wb = Workbook.open(INSPECT_FIXTURES["ChartsAndTables"])
    ext_lst = _chart(wb, "/xl/charts/chart1.xml").find(".//c:ser/c:extLst", NS)
    for copy in _C15_COPIES:
        ext_lst.append(
            etree.fromstring(
                f'<c:ext xmlns:c="{CHART_NS}" xmlns:c15="{C15_NS}" '
                f'uri="{{02D57815-91ED-43cb-92C2-25804820EDAC}}">{copy}</c:ext>'
            )
        )
    wb["bar chart"].insert_rows([[None]], at_row=3)
    assert _refs(wb, "/xl/charts/chart1.xml")[3:] == [
        "'bar chart'!$B$4:$B$10",
        "'bar chart'!$C$4:$C$7",
        "'bar chart'!$D$2",
    ]
    assert_valid_package(wb)


def test_a_chart_follows_another_sheet_it_plots():
    wb = Workbook.open(INSPECT_FIXTURES["ChartsAndTables"])
    values = _chart(wb, "/xl/charts/chart1.xml").findall(".//c:f", NS)[2]
    values.text = "('line chart'!$B$3:$B$6,'bar chart'!$B$3:$B$6)"
    wb["line chart"].insert_rows([[None]], at_row=1)
    assert values.text == "('line chart'!$B$4:$B$7,'bar chart'!$B$3:$B$6)"
    wb["bar chart"].insert_rows([[None]], at_row=6)
    assert values.text == "('line chart'!$B$4:$B$7,'bar chart'!$B$3:$B$7)"


def test_a_series_pushed_off_the_grid_becomes_a_ref_error():
    wb = Workbook.open(INSPECT_FIXTURES["ChartsAndTables"])
    name = _chart(wb, "/xl/charts/chart1.xml").find(".//c:tx//c:f", NS)
    name.text = "'bar chart'!$B$1048576"
    wb["bar chart"].insert_rows([[None]], at_row=1)
    assert name.text == "'bar chart'!#REF!"


def test_a_chart_that_declares_a_dtd_refuses_the_insert(unchecked_workbooks):
    wb = Workbook.open(INSPECT_FIXTURES["ChartsAndTables"])
    part = next(p for p in chart_parts(wb._package) if p.partname == "/xl/charts/chart1.xml")
    part._orig_blob = part._orig_blob.replace(b"?>", b"?><!DOCTYPE c:chartSpace>", 1)
    before = read_pkg(wb)
    with pytest.raises(DTDForbiddenError):
        wb["bar chart"].insert_rows([[None]], at_row=2)
    assert_preserved(before, read_pkg(wb))


def test_a_refused_insert_leaves_charts_alone():
    wb = Workbook.open(INSPECT_FIXTURES["ChartsAndTables"])
    ws = wb["bar chart"]
    ws["A1048576"].value = "last"
    before = read_pkg(wb)
    with pytest.raises(GridOverflowError):
        ws.insert_rows([[None]], at_row=2)
    assert_preserved(before, read_pkg(wb))


_CHARTSHEET_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"
_CHARTSHEET_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"


def _with_chartsheet() -> Workbook:
    """ChartsAndTables plus a chartsheet whose own chart plots 'bar chart'!A3:B6."""
    wb = Workbook.open(INSPECT_FIXTURES["ChartsAndTables"])
    package = wb._package
    chart = Part(
        PackURI("/xl/charts/chart9.xml"),
        CT.CHART,
        etree.tostring(deepcopy(_chart(wb, "/xl/charts/chart1.xml"))),
        package,
    )
    drawing = Part(PackURI("/xl/drawings/drawing9.xml"), CT.DRAWING, b"", package)
    drawing._blob = (
        f'<xdr:wsDr xmlns:xdr="{XDR_NS}" xmlns:a="{A_NS}"><xdr:absoluteAnchor>'
        '<xdr:pos x="0" y="0"/><xdr:ext cx="8666585" cy="6271846"/><xdr:graphicFrame macro="">'
        '<xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 9"/><xdr:cNvGraphicFramePr/>'
        '</xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
        f'<a:graphic><a:graphicData uri="{CHART_NS}"><c:chart xmlns:c="{CHART_NS}" '
        f'xmlns:r="{OFFICE_REL_NS}" r:id="{drawing.relate_to(chart, RT.CHART)}"/></a:graphicData>'
        "</a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>"
    ).encode()
    sheet = Part(PackURI("/xl/chartsheets/sheet9.xml"), _CHARTSHEET_CT, b"", package)
    sheet._blob = (
        f'<chartsheet xmlns="{SML_NS}" xmlns:r="{OFFICE_REL_NS}"><sheetViews>'
        '<sheetView workbookViewId="0"/></sheetViews><pageMargins left="0.7" right="0.7" '
        'top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
        f'<drawing r:id="{sheet.relate_to(drawing, RT.DRAWING)}"/></chartsheet>'
    ).encode()
    for part in (chart, drawing, sheet):
        package._add_part(part)
    workbook_part = wb._workbook_part
    r_id = workbook_part.relate_to(sheet, _CHARTSHEET_RT)
    workbook_part.append_sheet_element("Chart9", workbook_part.next_sheet_id(), r_id)
    buf = io.BytesIO()
    wb.save(buf)
    return Workbook.open(io.BytesIO(buf.getvalue()))


def test_inserts_work_in_a_workbook_with_a_chartsheet_and_move_its_chart():
    wb = _with_chartsheet()
    assert wb.sheetnames[-1] == "Chart9"
    wb["bar chart"].insert_rows([[None]], at_row=4)
    assert _refs(wb, "/xl/charts/chart9.xml") == [
        "'bar chart'!$B$2",
        "'bar chart'!$A$3:$A$7",
        "'bar chart'!$B$3:$B$7",
    ]
    wb["line chart"].insert_columns([[None]], at_col="A")
    assert_valid_package(wb)
