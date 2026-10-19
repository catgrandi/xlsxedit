"""Chart references follow the cells they plot on insert (issue #12)."""

from __future__ import annotations

import pytest
from lxml import etree

from xlsxedit import Workbook
from xlsxedit.drawing import C15_NS, CHART_NS, chart_parts
from xlsxedit.exceptions import GridOverflowError
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


def test_a_refused_insert_leaves_charts_alone():
    wb = Workbook.open(INSPECT_FIXTURES["ChartsAndTables"])
    ws = wb["bar chart"]
    ws["A1048576"].value = "last"
    before = read_pkg(wb)
    with pytest.raises(GridOverflowError):
        ws.insert_rows([[None]], at_row=2)
    assert_preserved(before, read_pkg(wb))
