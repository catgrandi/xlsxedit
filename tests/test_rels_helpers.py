"""Tests for Relationships.get_or_add / Part.relate_to."""

from __future__ import annotations

from xlsxedit.opc.constants import CT, RT
from xlsxedit.opc.packuri import PackURI
from xlsxedit.opc.part import Part


def _part(name: str) -> Part:
    return Part(PackURI(name), CT.XML)


def test_get_or_add_reuses_rid():
    source = _part("/xl/worksheets/sheet1.xml")
    target = _part("/xl/drawings/drawing1.xml")
    r_id1 = source.rels.get_or_add(RT.DRAWING, target)
    r_id2 = source.rels.get_or_add(RT.DRAWING, target)
    assert r_id1 == r_id2
    assert len(source.rels) == 1


def test_get_or_add_distinct_parts_get_new_rid():
    source = _part("/xl/worksheets/sheet1.xml")
    a = _part("/xl/drawings/drawing1.xml")
    b = _part("/xl/drawings/drawing2.xml")
    r_id_a = source.rels.get_or_add(RT.DRAWING, a)
    r_id_b = source.rels.get_or_add(RT.DRAWING, b)
    assert r_id_a != r_id_b
    assert len(source.rels) == 2


def test_get_or_add_distinct_reltypes_get_new_rid():
    source = _part("/xl/drawings/drawing1.xml")
    target = _part("/xl/media/image1.png")
    r_id_img = source.rels.get_or_add(RT.IMAGE, target)
    r_id_other = source.rels.get_or_add(RT.CHART, target)
    assert r_id_img != r_id_other


def test_relate_to_wraps_get_or_add():
    source = _part("/xl/worksheets/sheet1.xml")
    target = _part("/xl/tables/table1.xml")
    r_id1 = source.relate_to(target, RT.TABLE)
    r_id2 = source.relate_to(target, RT.TABLE)
    assert r_id1 == r_id2


def test_get_or_add_ext_rel_reuses_rid():
    source = _part("/xl/worksheets/sheet1.xml")
    url = "https://example.com/"
    r_id1 = source.rels.get_or_add_ext_rel(RT.HYPERLINK, url)
    r_id2 = source.rels.get_or_add_ext_rel(RT.HYPERLINK, url)
    assert r_id1 == r_id2
    assert len(source.rels) == 1
    r_id3 = source.rels.get_or_add_ext_rel(RT.HYPERLINK, "https://other.example/")
    assert r_id3 != r_id1
