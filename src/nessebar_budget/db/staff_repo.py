"""Import and query natural indicators (staff head-counts per activity)."""

from __future__ import annotations

from pathlib import Path

from sqlalchemy import delete, select, text
from sqlalchemy.exc import OperationalError
from sqlalchemy.orm import Session

from nessebar_budget.db.staff_models import NaturalIndicator
from nessebar_budget.parsers.naturi import STAFF_CODE, map_activity_codes, parse_naturi


def _activity_names(session: Session, year: int) -> dict[str, str]:
    """Activity name -> code, closest year first so a same-year spelling wins.

    Sources: names on the activity total rows of `cash_activity_lines`, and
    the numbered titles of already-imported new-form workbooks (so import the
    newer files before the 2019 one).
    """
    names: dict[str, str] = {}
    try:
        rows = session.execute(
            text(
                "SELECT DISTINCT activity_name, activity_code, CAST(substr(period,1,4) AS INTEGER) "
                "FROM cash_activity_lines WHERE activity_name IS NOT NULL"
            )
        ).all()
    except OperationalError:  # database pre-dates cash_activity_lines
        rows = []
    stored = session.execute(
        select(NaturalIndicator.block_title, NaturalIndicator.activity_code, NaturalIndicator.year)
        .where(NaturalIndicator.level == "activity", NaturalIndicator.activity_code.is_not(None))
        .distinct()
    ).all()
    for n, c, y in sorted([*rows, *stored], key=lambda t: abs((t[2] or 0) - year)):
        names.setdefault(n, c)
    return names


def import_natural_indicators(session: Session, path: str | Path, source_url: str | None = None) -> int:
    """Parse one workbook and replace the rows of its (year, basis). Returns
    the number stored; does not commit. Old-form activity blocks are mapped to
    activity codes by title; unmatched ones keep `activity_code=NULL`."""
    data = parse_naturi(path, source_url)
    year, basis, rows = data["year"], data["basis"], data["rows"]
    mapping, _unmapped = map_activity_codes(rows, _activity_names(session, year))
    session.execute(delete(NaturalIndicator).where(NaturalIndicator.year == year, NaturalIndicator.basis == basis))
    for r in rows:
        code = r["activity_code"] or (mapping.get(r["block_title"], (None,))[0] if r["level"] == "activity" else None)
        session.add(
            NaturalIndicator(
                year=year, basis=basis, level=r["level"], activity_code=code, block_title=r["block_title"],
                indicator_code=r["indicator_code"], indicator_name=r["indicator_name"], plan=r["plan"],
                total=r["total"], state=r["state"], local=r["local"], cofinance=r["cofinance"],
                source_url=source_url,
            )
        )
    session.flush()
    return len(rows)


def staff_by_activity(session: Session) -> dict[int, dict[str, float]]:
    """`{year: {activity_code: staff head-count}}` from the activity blocks'
    total staff indicator ("Щатна численост", code 100). For a year with both
    a report and a plan the report wins; see `staff_basis_by_year`."""
    basis = staff_basis_by_year(session)
    out: dict[int, dict[str, float]] = {}
    q = select(NaturalIndicator).where(
        NaturalIndicator.level == "activity",
        NaturalIndicator.indicator_code == STAFF_CODE,
        NaturalIndicator.activity_code.is_not(None),
    )
    for r in session.scalars(q):
        if r.basis == basis[r.year] and r.total:
            out.setdefault(r.year, {})[r.activity_code] = r.total
    return out


def staff_basis_by_year(session: Session) -> dict[int, str]:
    """`{year: "отчет" | "план"}` for the basis `staff_by_activity` uses."""
    out: dict[int, str] = {}
    for y, b in session.execute(select(NaturalIndicator.year, NaturalIndicator.basis).distinct()):
        if out.get(y) != "отчет":
            out[y] = b
    return out
