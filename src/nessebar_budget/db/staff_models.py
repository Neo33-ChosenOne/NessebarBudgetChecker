"""SQLAlchemy model for the municipality's natural indicators (head-counts).

Extends the shared `Base`; importing this module registers the table so
`init_db()` creates it (see `db/__init__.py`). One row is one indicator of one
block (whole municipality, a function, a group or an activity) of one year's
"натурални показатели" workbook, as parsed by `parsers.naturi`.
"""

from __future__ import annotations

from sqlalchemy import Float, Index, Integer, String, Text
from sqlalchemy.orm import Mapped, mapped_column

from nessebar_budget.db.models import Base


class NaturalIndicator(Base):
    """A non-zero indicator value from a natural-indicators workbook.

    `basis` is "отчет" (actual, finished year) or "план" (budgeted). New-form
    files have no separate plan column, so for a "план" file the planned
    figure sits in `total` and `plan` is NULL. `activity_code` is the 4-digit
    code used in `cash_activity_lines`; NULL for municipality/function/group
    blocks and for activity blocks whose title could not be matched.
    """

    __tablename__ = "natural_indicators"
    __table_args__ = (Index("ix_natural_indicators_year_basis", "year", "basis"),)

    id: Mapped[int] = mapped_column(primary_key=True)
    year: Mapped[int] = mapped_column(Integer)
    basis: Mapped[str] = mapped_column(String(8))
    level: Mapped[str] = mapped_column(String(16))
    activity_code: Mapped[str | None] = mapped_column(String(4), default=None, index=True)
    block_title: Mapped[str] = mapped_column(Text)
    indicator_code: Mapped[str] = mapped_column(String(8))
    indicator_name: Mapped[str] = mapped_column(Text)
    plan: Mapped[float | None] = mapped_column(Float, default=None)
    total: Mapped[float | None] = mapped_column(Float, default=None)
    state: Mapped[float | None] = mapped_column(Float, default=None)
    local: Mapped[float | None] = mapped_column(Float, default=None)
    cofinance: Mapped[float | None] = mapped_column(Float, default=None)
    source_url: Mapped[str | None] = mapped_column(Text, default=None)
