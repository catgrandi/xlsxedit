"""``xlsxedit._ooxml_order`` is generated from the vendored schemas and must not drift."""

from __future__ import annotations

import ast
import importlib.util
from pathlib import Path

import pytest
from lxml import etree

from xlsxedit import _ooxml_order
from xlsxedit._ooxml_order import ATTRIBUTE_DEFAULTS, ENUMS, EXTENSION_ORDER, ORDER
from xlsxedit.worksheet_order import _WS_CHILD_ORDER

GENERATOR = Path(__file__).resolve().parents[1] / "scripts" / "gen_ooxml_order.py"

_X14 = "{http://schemas.microsoft.com/office/spreadsheetml/2009/9/main}"
_XM = "{http://schemas.microsoft.com/office/excel/2006/main}"


@pytest.fixture(scope="module")
def generator():
    spec = importlib.util.spec_from_file_location("gen_ooxml_order", GENERATOR)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_generated_module_matches_schemas(generator):
    committed = generator.OUTPUT.read_text(encoding="utf-8")
    assert committed == generator.render(), "run: python scripts/gen_ooxml_order.py"
    assert generator.main(["--check"]) == 0


def test_tables_cover_the_requested_types():
    assert set(ORDER) == {
        "CT_Worksheet",
        "CT_Stylesheet",
        "CT_Xf",
        "CT_Workbook",
        "CT_Table",
        "CT_TableColumn",
        "CT_Row",
        "CT_Cell",
        "CT_CfRule",
        "CT_ConditionalFormatting",
        "CT_DataValidation",
        "CT_TwoCellAnchor",
        "CT_Chart",
        "CT_ChartSpace",
    }
    assert set(EXTENSION_ORDER) == {
        f"{_X14}{name}"
        for name in (
            "CT_ConditionalFormattings",
            "CT_ConditionalFormatting",
            "CT_CfRule",
            "CT_ColorScale",
            "CT_DataBar",
            "CT_IconSet",
            "CT_Cfvo",
            "CT_DataValidations",
            "CT_DataValidation",
            "CT_DataValidationFormula",
            "CT_SparklineGroups",
            "CT_SparklineGroup",
            "CT_Sparklines",
            "CT_Sparkline",
        )
    }
    assert set(ATTRIBUTE_DEFAULTS) == {"CT_Table", "CT_TableColumn"}
    assert set(ENUMS) == {"ST_TableType", "ST_TotalsRowFunction"}


def test_module_reads_no_file():
    """The package ships without ``schemas/``, so the tables must be literals."""
    tree = ast.parse(Path(_ooxml_order.__file__).read_text(encoding="utf-8"))
    imports = [node for node in tree.body if isinstance(node, (ast.Import, ast.ImportFrom))]
    assert [ast.unparse(node) for node in imports] == ["from __future__ import annotations"]


def test_x14_children_are_named_with_their_namespace():
    assert EXTENSION_ORDER[f"{_X14}CT_ConditionalFormatting"] == (
        f"{_X14}cfRule",
        f"{_XM}sqref",
        f"{_X14}extLst",
    )
    assert EXTENSION_ORDER[f"{_X14}CT_CfRule"] == (
        f"{_XM}f",
        f"{_X14}colorScale",
        f"{_X14}dataBar",
        f"{_X14}iconSet",
        f"{_X14}dxf",
        f"{_X14}extLst",
    )
    assert EXTENSION_ORDER[f"{_X14}CT_DataValidation"] == (
        f"{_X14}formula1",
        f"{_X14}formula2",
        f"{_XM}sqref",
    )
    group = EXTENSION_ORDER[f"{_X14}CT_SparklineGroup"]
    assert group[0] == f"{_X14}colorSeries"
    assert group[-2:] == (f"{_XM}f", f"{_X14}sparklines")
    assert len(group) == 10


def test_worksheet_order_matches_schema():
    order = ORDER["CT_Worksheet"]
    assert _WS_CHILD_ORDER == order
    assert order.index("legacyDrawing") == order.index("drawing") + 1
    assert order.index("drawingHF") == order.index("legacyDrawingHF") + 1


def test_choice_alternatives_are_listed_in_place():
    assert ORDER["CT_TwoCellAnchor"] == (
        "from",
        "to",
        "sp",
        "grpSp",
        "graphicFrame",
        "cxnSp",
        "pic",
        "contentPart",
        "clientData",
    )


def test_table_attribute_defaults_and_enums():
    assert ATTRIBUTE_DEFAULTS["CT_Table"]["headerRowCount"] == "1"
    assert ATTRIBUTE_DEFAULTS["CT_Table"]["totalsRowShown"] == "true"
    assert ATTRIBUTE_DEFAULTS["CT_TableColumn"] == {"totalsRowFunction": "none"}
    assert ENUMS["ST_TableType"] == ("worksheet", "xml", "queryTable")
    assert ENUMS["ST_TotalsRowFunction"][0] == "none"
    assert "countNums" in ENUMS["ST_TotalsRowFunction"]


def test_generator_refuses_a_repeating_choice(generator):
    complex_type = etree.fromstring(
        '<xsd:complexType xmlns:xsd="http://www.w3.org/2001/XMLSchema">'
        '<xsd:choice maxOccurs="unbounded"><xsd:element name="a"/><xsd:element name="b"/>'
        "</xsd:choice></xsd:complexType>"
    )
    with pytest.raises(ValueError, match="repeating xsd:choice"):
        generator._element_names("transitional/sml.xsd", complex_type)


def test_ecma_tables_refuse_a_child_from_another_namespace(generator):
    with pytest.raises(ValueError, match="list it as an extension"):
        generator.child_order(generator._X14_SCHEMA, "CT_ConditionalFormatting")
