"""``fixtures/MANIFEST.md`` is generated from the fixtures and must not drift."""

from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest

from tests.conftest import BOOK1, FIXTURES, INSPECT_FIXTURES

GENERATOR = Path(__file__).resolve().parents[1] / "scripts" / "describe_fixtures.py"


@pytest.fixture(scope="module")
def generator():
    spec = importlib.util.spec_from_file_location("describe_fixtures", GENERATOR)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_manifest_matches_fixtures(generator):
    committed = generator.OUTPUT.read_text(encoding="utf-8")
    assert committed == generator.render(), "run: python scripts/describe_fixtures.py"
    assert generator.main(["--check"]) == 0


def test_every_fixture_is_registered():
    """A workbook added to ``fixtures/`` is covered by ``test_fidelity`` once registered."""
    on_disk = {p for p in FIXTURES.iterdir() if p.suffix.lower() in (".xlsx", ".xlsm")}
    assert on_disk == {BOOK1, *INSPECT_FIXTURES.values()}
