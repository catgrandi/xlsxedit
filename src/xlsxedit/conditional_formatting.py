"""Conditional formatting helpers."""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import TYPE_CHECKING

from lxml import etree

from xlsxedit.api import _default_xlsx_path
from xlsxedit.opc.constants import SML_NS
from xlsxedit.oxml.parser import parse_template_xml

if TYPE_CHECKING:
    from xlsxedit.worksheet import Worksheet

_CF = f"{{{SML_NS}}}conditionalFormatting"
_CF_RULE = f"{{{SML_NS}}}cfRule"
_FORMULA = f"{{{SML_NS}}}formula"
_COLOR_SCALE = f"{{{SML_NS}}}colorScale"

_COLOR_SCALE_TEMPLATE: etree._Element | None = None


def _color_scale_template() -> etree._Element:
    global _COLOR_SCALE_TEMPLATE
    if _COLOR_SCALE_TEMPLATE is None:
        path = _default_xlsx_path().parent / "default-cf-colorscale.xml"
        _COLOR_SCALE_TEMPLATE = parse_template_xml(path.read_bytes())
    return deepcopy(_COLOR_SCALE_TEMPLATE)


class ConditionalFormatRule:
    """Read-only view of a single ``cfRule``."""

    def __init__(self, element: etree._Element):
        self._element = element

    @property
    def type(self) -> str | None:
        return self._element.get("type")

    @property
    def priority(self) -> int | None:
        val = self._element.get("priority")
        return int(val) if val is not None else None

    @property
    def operator(self) -> str | None:
        return self._element.get("operator")

    @property
    def formulas(self) -> list[str]:
        return [f.text or "" for f in self._element.findall(_FORMULA)]


class ConditionalFormatting:
    """A ``conditionalFormatting`` block on a worksheet."""

    def __init__(self, element: etree._Element):
        self._element = element

    def __repr__(self) -> str:
        return f"<ConditionalFormatting {self.cell_range!r}>"

    @property
    def cell_range(self) -> str:
        return self._element.get("sqref", "")

    @property
    def rules(self) -> list[ConditionalFormatRule]:
        return [ConditionalFormatRule(r) for r in self._element.findall(_CF_RULE)]


def iter_conditional_formatting(worksheet: Worksheet) -> list[ConditionalFormatting]:
    root = worksheet._part.element
    return [ConditionalFormatting(elm) for elm in root.findall(_CF)]


def _next_priority(worksheet: Worksheet) -> int:
    max_pri = 0
    for block in worksheet._part.element.findall(_CF):
        for rule in block.findall(_CF_RULE):
            p = rule.get("priority")
            if p is not None:
                max_pri = max(max_pri, int(p))
    return max_pri + 1


def add_cell_is_rule(
    worksheet: Worksheet,
    cell_range: str,
    *,
    operator: str = "greaterThan",
    formula: str = "0",
    priority: int | None = None,
    font_color: str | None = None,
    bg_color: str | None = None,
    bold: bool = False,
) -> ConditionalFormatting:
    """Add a ``cellIs`` conditional-format rule to ``cell_range``."""
    root = worksheet._part.element
    if priority is None:
        priority = _next_priority(worksheet)

    cf_elm = etree.Element(_CF)
    cf_elm.set("sqref", cell_range)
    rule_elm = etree.SubElement(cf_elm, _CF_RULE)
    rule_elm.set("type", "cellIs")
    rule_elm.set("priority", str(priority))
    rule_elm.set("operator", operator)
    formula_elm = etree.SubElement(rule_elm, _FORMULA)
    formula_elm.text = formula

    if font_color is None and bg_color is None and not bold:
        bg_color = "C6EFCE"
        font_color = "006100"
    styles = worksheet._workbook.styles
    dxf_id = styles.ensure_dxf(font_color=font_color, bg_color=bg_color, bold=bold)
    rule_elm.set("dxfId", str(dxf_id))

    from xlsxedit.worksheet_order import insert_worksheet_child

    insert_worksheet_child(root, cf_elm)
    return ConditionalFormatting(cf_elm)


def clone_color_scale_rule(
    worksheet: Worksheet, cell_range: str, template: etree._Element | None = None
) -> ConditionalFormatting:
    """Clone a ``colorScale`` rule from a template ``conditionalFormatting`` element."""
    from xlsxedit.worksheet_order import insert_worksheet_child

    root = worksheet._part.element
    cf_elm = deepcopy(template if template is not None else _color_scale_template())
    cf_elm.set("sqref", cell_range)
    for rule in cf_elm.findall(_CF_RULE):
        rule.set("priority", str(_next_priority(worksheet)))
    insert_worksheet_child(root, cf_elm)
    return ConditionalFormatting(cf_elm)


def _expand_sqref_range(sqref: str, end_row: int) -> str:
    from xlsxedit.merge import parse_range
    from xlsxedit.oxml.address import join_address

    parts = []
    for ref in sqref.split():
        try:
            c1, r1, c2, r2 = parse_range(ref)
        except ValueError:
            parts.append(ref)
            continue
        # Only grow ranges that already span multiple rows. Single-cell CF
        # (e.g. a KPI on D7) must not expand down the column.
        if c1 == c2 and r2 > r1 and r2 < end_row:
            ref = f"{join_address(c1, r1)}:{join_address(c2, end_row)}"
        parts.append(ref)
    return " ".join(parts)


def expand_conditional_formatting_to_row(worksheet: Worksheet, end_row: int) -> int:
    """Extend multi-row single-column CF ``sqref`` ranges down to ``end_row``."""
    updated = 0
    for block in iter_conditional_formatting(worksheet):
        sqref = block.cell_range
        new_sqref = _expand_sqref_range(sqref, end_row)
        if new_sqref != sqref:
            block._element.set("sqref", new_sqref)
            updated += 1
    return updated
