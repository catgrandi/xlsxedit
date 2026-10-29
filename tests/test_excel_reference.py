"""Excel-reference comparisons: an xlsxedit operation against what Excel itself wrote.

Each folder in ``fixtures/excel-reference/`` holds ``before.xlsx``, Excel's
``after.xlsx`` for one operation, and a README with the Excel steps and build.
A test applies the matching xlsxedit operation to ``before.xlsx`` and compares
the result with ``after.xlsx`` one aspect at a time through
``assert_matches_excel``. The comparison reads targeted facts
(``tests/excel_facts.py``), never bytes or whole packages: Excel rewrites
revision ids, ``calcId``, document properties, the dimension and the ZIP
layout on every save.

An aspect xlsxedit does not match yet is a strict xfail naming the roadmap
issue that fixes it, so the mark has to go when the fix lands.
"""

from __future__ import annotations

import difflib
from collections.abc import Callable, Iterable, Iterator, Mapping
from pathlib import Path
from typing import NamedTuple

import pytest

from xlsxedit import Workbook
from tests import excel_facts as facts
from tests.conftest import FIXTURES

REFERENCE = FIXTURES / "excel-reference"

ASPECTS: dict[str, Callable[[facts.Book, str | None], object]] = {
    "tables": lambda book, band: facts.tables(book),
    "conditional_formatting": lambda book, band: facts.conditional_formats(book),
    "data_validation": lambda book, band: facts.data_validations(book),
    "drawings": lambda book, band: facts.drawings_by_name(book),
    "defined_names": lambda book, band: facts.defined_names(book),
    "merges": lambda book, band: facts.merges(book),
    "values": lambda book, band: facts.cells(book, band),
    "styles": lambda book, band: facts.styles(book, band),
    "rows": lambda book, band: facts.rows(book, band),
}
BAND_ASPECTS = frozenset({"values", "styles", "rows"})


def assert_matches_excel(result, after, aspects: Iterable[str], *, band: str | None = None) -> None:
    """Assert that ``result`` shows what Excel's ``after`` shows in each of ``aspects``.

    ``result`` and ``after`` are anything ``tests.preservation.read_pkg``
    accepts, a ``Workbook`` included. ``aspects`` are keys of ``ASPECTS``;
    ``band`` (``"Sheet!A1:F20"``) is the range the ``values``, ``styles`` and
    ``rows`` aspects read. The ``AssertionError`` lists, per aspect, the facts
    only one side has.
    """
    aspects = list(aspects)
    unknown = sorted(set(aspects) - ASPECTS.keys())
    if not aspects or unknown:
        raise ValueError(f"aspects must come from {sorted(ASPECTS)}, got {aspects!r}")
    if band is None and BAND_ASPECTS.intersection(aspects):
        raise ValueError(f"{sorted(BAND_ASPECTS.intersection(aspects))} need a band")
    ours, theirs = facts.Book(result), facts.Book(after)
    problems = []
    for aspect in aspects:
        got, want = ASPECTS[aspect](ours, band), ASPECTS[aspect](theirs, band)
        if got != want:
            diff = difflib.unified_diff(
                list(_lines(want)), list(_lines(got)), "excel", "xlsxedit", lineterm="", n=0
            )
            problems.append(f"{aspect}:\n" + "\n".join(diff))
    if problems:
        raise AssertionError("does not match Excel's result in " + "\n".join(problems))


def _lines(value, prefix: str = "") -> Iterator[str]:
    """One line per fact, keyed by its path, so a diff shows only what differs."""
    if isinstance(value, Mapping):
        for key in sorted(value, key=_order):
            yield from _lines(value[key], f"{prefix}{key}: ")
    elif isinstance(value, (tuple, list)) and not hasattr(value, "_fields"):
        for item in value:
            yield from _lines(item, prefix)
        if not value:
            yield f"{prefix}()"
    else:
        yield f"{prefix}{value!r}"


def _order(key) -> tuple:
    try:
        return (0, *facts.parse_cell(key), "")
    except (ValueError, AttributeError):
        return (1, 0, 0, str(key))


class Operation(NamedTuple):
    """An Excel-reference pair and the xlsxedit call that should reproduce it."""

    apply: Callable[[Workbook], object]
    band: str
    changed: tuple[str, ...]  # aspects Excel changed, so before and after differ there
    known: Mapping[str, str]  # aspect -> "#issue: reason" while xlsxedit differs
    raises: type[BaseException] = AssertionError


_NO_APPEND = "#17: Table.append_rows, which Excel's ListRows.Add maps to, does not exist yet"
_NO_DELETE = "#22: Worksheet.delete_rows does not exist yet"

OPERATIONS: dict[str, Operation] = {
    "table-append-row": Operation(
        apply=lambda wb: wb["Orders"].tables[0].append_rows([[]]),
        band="Orders!A1:G10",
        changed=("tables", "conditional_formatting", "data_validation", "values", "styles"),
        known=dict.fromkeys(ASPECTS, _NO_APPEND),
        raises=AttributeError,
    ),
    "insert-rows-above-table": Operation(
        apply=lambda wb: wb["Stock"].insert_rows([[], []], at_row=3),
        band="Stock!A1:H14",
        changed=(
            "tables",
            "conditional_formatting",
            "data_validation",
            "defined_names",
            "merges",
            "values",
        ),
        known={
            "values": "#23: cell formula text is not rewritten on insert "
            "(SUM(B5:B9) stays, Excel writes SUM(B7:B11))",
        },
    ),
    "insert-rows-through-drawings": Operation(
        apply=lambda wb: wb["Drawings"].insert_rows([[]] * 3, at_row=7),
        band="Drawings!A1:B35",
        changed=("drawings", "values", "rows"),
        known={
            "drawings": "#21: drawing anchors do not move, grow or keep their size on insert",
            "rows": "#19: inserted rows do not take the height of the row above",
        },
    ),
    "delete-rows-under-cf": Operation(
        apply=lambda wb: wb["Sheet1"].delete_rows(5, 3),
        band="Sheet1!A1:K22",
        changed=(
            "conditional_formatting",
            "data_validation",
            "drawings",
            "defined_names",
            "merges",
            "values",
        ),
        known=dict.fromkeys(ASPECTS, _NO_DELETE),
        raises=AttributeError,
    ),
    "insert-rows-at-drawing-edges": Operation(
        apply=lambda wb: wb["Edges"].insert_rows([[]] * 2, at_row=7),
        band="Edges!A1:B20",
        changed=("drawings",),
        known={"drawings": "#21: drawing anchors do not move on insert"},
    ),
    "delete-rows-at-drawing-edges": Operation(
        apply=lambda wb: wb["Edges"].delete_rows(5, 3),
        band="Edges!A1:B20",
        changed=("drawings",),
        known=dict.fromkeys(ASPECTS, _NO_DELETE),
        raises=AttributeError,
    ),
}


def _pair(name: str) -> tuple[Path, Path]:
    return REFERENCE / name / "before.xlsx", REFERENCE / name / "after.xlsx"


def _cases() -> Iterator:
    for name, op in OPERATIONS.items():
        for aspect in ASPECTS:
            marks = ()
            if aspect in op.known:
                marks = pytest.mark.xfail(strict=True, raises=op.raises, reason=op.known[aspect])
            yield pytest.param(name, aspect, id=f"{name}-{aspect}", marks=marks)


@pytest.mark.parametrize(("name", "aspect"), list(_cases()))
def test_matches_excel(name: str, aspect: str):
    before, after = _pair(name)
    op = OPERATIONS[name]
    wb = Workbook.open(before)
    op.apply(wb)
    assert_matches_excel(wb, after, [aspect], band=op.band)


@pytest.mark.parametrize("name", OPERATIONS)
def test_harness_sees_what_excel_changed(name: str):
    """Before and after differ in exactly the aspects the operation lists as changed."""
    before, after = _pair(name)
    op = OPERATIONS[name]
    assert_matches_excel(after, after, ASPECTS, band=op.band)
    differ = []
    for aspect in ASPECTS:
        try:
            assert_matches_excel(before, after, [aspect], band=op.band)
        except AssertionError:
            differ.append(aspect)
    assert differ == list(op.changed)


def test_every_pair_is_tested_and_documented():
    folders = {p.name for p in REFERENCE.iterdir() if p.is_dir()}
    assert folders == set(OPERATIONS)
    for name in folders:
        assert {p.name for p in (REFERENCE / name).iterdir()} == {
            "README.md",
            "before.xlsx",
            "after.xlsx",
        }
        readme = (REFERENCE / name / "README.md").read_text(encoding="utf-8")
        assert "- Excel: 16." in readme, f"{name}/README.md does not record the Excel build"


def test_mismatch_lists_only_the_differing_facts():
    before, after = _pair("insert-rows-above-table")
    with pytest.raises(AssertionError) as caught:
        assert_matches_excel(before, after, ["merges", "tables"])
    message = str(caught.value)
    assert "-Stock: 'A14:B14'" in message
    assert "+Stock: 'A12:B12'" in message
    assert "'A1:D1'" not in message


def test_bad_arguments_are_refused():
    before, after = _pair("insert-rows-above-table")
    with pytest.raises(ValueError, match="aspects must come from"):
        assert_matches_excel(before, after, ["bytes"])
    with pytest.raises(ValueError, match="need a band"):
        assert_matches_excel(before, after, ["values"])
