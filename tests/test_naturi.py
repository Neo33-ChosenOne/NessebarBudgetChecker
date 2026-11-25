"""Tests for the natural-indicators parser, importer and staff query."""

from __future__ import annotations

from pathlib import Path

import pytest
from sqlalchemy import create_engine, func, select
from sqlalchemy.orm import Session

from nessebar_budget.db.models import Base
from nessebar_budget.db.staff_models import NaturalIndicator
from nessebar_budget.db.staff_repo import import_natural_indicators, staff_by_activity
from nessebar_budget.parsers.naturi import clean_indicator_name, map_activity_codes, parse_naturi

ROOT = Path(__file__).resolve().parents[1]
SAMPLE_2021 = ROOT / "data/samples/naturi_2021_5206.xls"  # new form, report
CACHE_2019 = ROOT / "data/cache/nesebar_site/2024/02/NaturiPokazateli_2019_5206.xlsx"  # old form


def _find(rows, level, indicator, title_part=None):
    return [
        r for r in rows
        if r["level"] == level and r["indicator_code"] == indicator and (not title_part or title_part in r["block_title"])
    ]


def test_clean_indicator_name_collapses_letter_spacing() -> None:
    assert clean_indicator_name("Щ а т н и   б р о й к и") == "Щатни бройки"
    assert clean_indicator_name("ДEЖУPHИ  ") == "ДЕЖУРНИ"  # Latin look-alikes folded to Cyrillic


def test_parse_new_form_report() -> None:
    d = parse_naturi(SAMPLE_2021, "https://www.nesebar.bg/otcheti/Naturi_2021_5206.xls")
    assert (d["year"], d["basis"], d["form"]) == (2021, "отчет", "new")
    rows = d["rows"]
    assert all(any(r[k] for k in ("plan", "total", "state", "local", "cofinance")) for r in rows)
    assert _find(rows, "municipality", "100")[0]["total"] == 1551.5
    admin = _find(rows, "activity", "100", "ОБЩИНСKА АДМИНИСТРАЦИЯ")[0]
    assert (admin["activity_code"], admin["total"]) == ("1122", 220.0)
    assert _find(rows, "activity", "100", "ДЕТСKИ ГРАДИНИ")[0]["activity_code"] == "3311"
    assert _find(rows, "activity", "100", "ЧИСТОТА")[0]["total"] == 431.0


def test_basis_hint_makes_plan() -> None:
    assert parse_naturi(SAMPLE_2021, "https://www.nesebar.bg/budget/Naturi_2021_5206.xls")["basis"] == "план"


@pytest.mark.skipif(not CACHE_2019.exists(), reason="2019 workbook is only in the local cache")
def test_parse_old_form_and_code_normalisation() -> None:
    d = parse_naturi(CACHE_2019)
    assert (d["year"], d["basis"], d["form"]) == (2019, "отчет", "old")
    muni = _find(d["rows"], "municipality", "100")[0]  # "01-00" -> "100"
    assert muni["total"] == 1930.5 and muni["indicator_name"] == "Щатни бройки"
    mapping, unmapped = map_activity_codes(d["rows"], {"Общинска администрация": "1122"})
    assert mapping["ОБЩИНСKА АДМИНИСТРАЦИЯ"][0] == "1122" and "ЧИСТОТА" in unmapped


def test_map_activity_codes_never_guesses() -> None:
    rows = [{"level": "activity", "block_title": t, "activity_code": None, "function_code": "6"}
            for t in ("Чистота", "Нещо непознато")]
    mapping, unmapped = map_activity_codes(rows, {"6623 Чистота": "6623", "Озеленяване": "6622"})
    assert mapping == {"Чистота": ("6623", "exact")} and unmapped == ["Нещо непознато"]


def test_import_twice_does_not_duplicate() -> None:
    engine = create_engine("sqlite://")
    Base.metadata.create_all(engine)
    with Session(engine) as s:
        n1 = import_natural_indicators(s, SAMPLE_2021, "u")
        n2 = import_natural_indicators(s, SAMPLE_2021, "u")
        assert n1 == n2 > 0
        assert s.scalar(select(func.count()).select_from(NaturalIndicator)) == n1
        staff = staff_by_activity(s)
        assert staff[2021]["1122"] == 220.0 and staff[2021]["6623"] == 431.0
