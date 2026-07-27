"""Bundled templates may be pretty-printed; output parts must stay compact."""

from __future__ import annotations

import subprocess
import zipfile
from pathlib import Path

from xlsxedit import Workbook
from xlsxedit.api import _default_xlsx_path
from xlsxedit.oxml.parser import parse_template_xml, parse_xml, serialize_xml

TEMPLATES = _default_xlsx_path().parent

# Serialized size of default-bar-chart after parse_template_xml (compact baseline).
_COMPACT_BAR_CHART_BYTES = 6031


def test_parse_template_xml_strips_indentation():
    raw = (TEMPLATES / "default-bar-chart.xml").read_bytes()
    assert b"\n  " in raw
    loose = serialize_xml(parse_xml(raw))
    tight = serialize_xml(parse_template_xml(raw))
    assert len(tight) < len(loose)
    assert len(tight) == _COMPACT_BAR_CHART_BYTES


def test_fragment_templates_serialize_compact():
    """Pretty-printed fragment sources must not bloat serialized parts."""
    cases = (
        "default-bar-chart.xml",
        "default-chart-anchor.xml",
        "default-table.xml",
        "default-cf-colorscale.xml",
        "default-picture-anchor.xml",
        "default-worksheet.xml",
    )
    for name in cases:
        raw = (TEMPLATES / name).read_bytes()
        loose = serialize_xml(parse_xml(raw))
        tight = serialize_xml(parse_template_xml(raw))
        assert len(tight) <= len(loose), name


def test_add_chart_writes_compact_chart_part(tmp_path: Path):
    wb = Workbook.create()
    ws = wb["Sheet1"]
    ws["A1"].value = "Day"
    ws["B1"].value = "N"
    ws["A2"].value = "Mon"
    ws["B2"].value = 1
    ws.add_chart("bar", anchor="D2", data_range="A1:B2", title="T")
    out = tmp_path / "chart.xlsx"
    wb.save(out)

    template_bloated = len(
        serialize_xml(parse_xml((TEMPLATES / "default-bar-chart.xml").read_bytes()))
    )
    with zipfile.ZipFile(out) as z:
        chart_blob = z.read("xl/charts/chart1.xml")
    # Title/series edits grow the part; must stay far below pretty-print bloat (~9k+).
    assert len(chart_blob) < template_bloated - 2500
    assert len(chart_blob) < 7000


def test_matches_historical_minified_bar_chart_if_available():
    """Guard against template edits: compact output should match last minified roundtrip."""
    result = subprocess.run(
        ["git", "show", "HEAD:src/xlsxedit/templates/default-bar-chart.xml"],
        capture_output=True,
        cwd=Path(__file__).resolve().parents[1],
    )
    if result.returncode != 0:
        return
    historical = serialize_xml(parse_xml(result.stdout))
    current = serialize_xml(
        parse_template_xml((TEMPLATES / "default-bar-chart.xml").read_bytes())
    )
    assert current == historical
