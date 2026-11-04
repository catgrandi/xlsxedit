"""Sheet names Excel rejects are refused before anything changes.

Recorded with Excel 16 on 2026-11-04 by renaming a sheet through COM: Excel
refuses an empty name, one over 31 characters counted in UTF-16 code units,
one holding ``: \\ / ? * [ ]`` or beginning or ending with an apostrophe
("You typed an invalid name for a sheet or chart."), "History" in any case
("History is a reserved name."), and another sheet's name in another case
("That name is already taken."). It does not open, without repairing it, a file
whose sheet names break any of these rules but "History".
"""

from __future__ import annotations

import io

import pytest

from xlsxedit import DuplicateWorksheetError, Workbook
from tests.conftest import INSPECT_FIXTURES
from tests.preservation import assert_preserved, read_pkg

EMOJI = "\N{GRINNING FACE}"

INVALID = {
    "empty": "",
    "32 characters": "a" * 32,
    "31 characters and an emoji's second code unit": "a" * 30 + EMOJI,
    "History": "History",
    "history": "history",
    "HISTORY": "HISTORY",
    "slash": "a/b",
    "backslash": "a\\b",
    "question mark": "a?b",
    "asterisk": "a*b",
    "opening bracket": "a[b",
    "closing bracket": "a]b",
    "colon": "a:b",
    "leading apostrophe": "'abc",
    "trailing apostrophe": "abc'",
    "a character XML cannot store": "a\x01b",
}

VALID = {
    "31 characters": "a" * 31,
    "30 code units with an emoji": "a" * 28 + EMOJI,
    "History1": "History1",
    "inner apostrophe": "a'b",
    "a space": " ",
    "leading space": " abc",
    "a tab": "a\tb",
}

# Pairs Excel takes for one name, and pairs it does not, after naming the first sheet.
CLASHES = {
    "ASCII case": ("Ax", "ax"),
    "accented case": ("\N{LATIN SMALL LETTER E WITH ACUTE}x", "\N{LATIN CAPITAL LETTER E WITH ACUTE}x"),
    "sigma": ("\N{GREEK SMALL LETTER SIGMA}x", "\N{GREEK CAPITAL LETTER SIGMA}x"),
    "final sigma": ("\N{GREEK SMALL LETTER FINAL SIGMA}x", "\N{GREEK SMALL LETTER SIGMA}x"),
    "Kelvin sign": ("\N{KELVIN SIGN}x", "kx"),
    "Ohm sign": ("\N{OHM SIGN}x", "\N{GREEK SMALL LETTER OMEGA}x"),
    "capital sharp s": ("\N{LATIN CAPITAL LETTER SHARP S}x", "\N{LATIN SMALL LETTER SHARP S}x"),
    "titlecase digraph": ("\N{LATIN CAPITAL LETTER D WITH SMALL LETTER Z WITH CARON}x", "\N{LATIN SMALL LETTER DZ WITH CARON}x"),
}
DISTINCT = {
    "sharp s and SS": ("Stra\N{LATIN SMALL LETTER SHARP S}e", "STRASSE"),
    "long s": ("\N{LATIN SMALL LETTER LONG S}x", "Sx"),
    "dotless i": ("\N{LATIN SMALL LETTER DOTLESS I}x", "Ix"),
    "dotted capital I": ("ix", "\N{LATIN CAPITAL LETTER I WITH DOT ABOVE}x"),
    "full width": ("\N{FULLWIDTH LATIN CAPITAL LETTER A}x", "Ax"),
    "kana": ("\N{HIRAGANA LETTER A}x", "\N{KATAKANA LETTER A}x"),
    "combining accent": ("e\N{COMBINING ACUTE ACCENT}x", "\N{LATIN SMALL LETTER E WITH ACUTE}x"),
    "accent": ("ax", "\N{LATIN SMALL LETTER A WITH ACUTE}x"),
}

NAMINGS = {
    "add_worksheet": lambda wb, name: wb.add_worksheet(name),
    "copy_worksheet": lambda wb, name: wb.copy_worksheet("bar chart", name),
    "rename_worksheet": lambda wb, name: wb.rename_worksheet("bar chart", name),
}


def _book() -> Workbook:
    """ChartsAndTables: sheets "bar chart", "line chart" and "Table"."""
    return Workbook.open(INSPECT_FIXTURES["ChartsAndTables"])


def _reopened(wb: Workbook) -> Workbook:
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return Workbook.open(buf)


@pytest.mark.parametrize("naming", NAMINGS)
@pytest.mark.parametrize("name", INVALID.values(), ids=INVALID)
def test_a_name_excel_rejects_is_refused_before_anything_changes(name: str, naming: str):
    wb = _book()
    before = read_pkg(wb)
    with pytest.raises(ValueError, match="invalid sheet name"):
        NAMINGS[naming](wb, name)
    assert_preserved(before, read_pkg(wb))


@pytest.mark.parametrize("naming", NAMINGS)
def test_a_name_that_is_not_a_str_is_refused_before_anything_changes(naming: str):
    wb = _book()
    before = read_pkg(wb)
    with pytest.raises(TypeError, match="sheet name must be a str"):
        NAMINGS[naming](wb, 2024)
    assert_preserved(before, read_pkg(wb))


@pytest.mark.parametrize("naming", NAMINGS)
@pytest.mark.parametrize("name", VALID.values(), ids=VALID)
def test_names_excel_accepts_are_kept(name: str, naming: str):
    wb = _book()
    ws = NAMINGS[naming](wb, name)
    assert ws.name == name
    assert _reopened(wb)[name].name == name


@pytest.mark.parametrize("naming", NAMINGS)
@pytest.mark.parametrize("pair", CLASHES.values(), ids=CLASHES)
def test_another_sheet_s_name_in_another_case_is_refused(pair: tuple[str, str], naming: str):
    wb = _book()
    wb.rename_worksheet("line chart", pair[0])
    before = read_pkg(wb)
    with pytest.raises(DuplicateWorksheetError, match=f"already exists as {pair[0]!r}"):
        NAMINGS[naming](wb, pair[1])
    assert_preserved(before, read_pkg(wb))


@pytest.mark.parametrize("naming", NAMINGS)
@pytest.mark.parametrize("pair", DISTINCT.values(), ids=DISTINCT)
def test_names_excel_tells_apart_are_two_sheets(pair: tuple[str, str], naming: str):
    wb = _book()
    wb.rename_worksheet("line chart", pair[0])
    NAMINGS[naming](wb, pair[1])
    assert {pair[0], pair[1]} <= set(_reopened(wb).sheetnames)


def test_a_sheet_keeps_or_changes_the_case_of_its_own_name():
    wb = _book()
    first = wb.sheetnames[0]
    assert wb.rename_worksheet(first, first).name == first
    assert wb.rename_worksheet(first, first.upper()).name == first.upper()
    with pytest.raises(DuplicateWorksheetError):
        wb.add_worksheet(first.lower())

