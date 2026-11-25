"""SQLAlchemy models for parsed budget-execution line items.

These extend the shared `Base` declared in `db.models` (which also defines
`BudgetReport`, the row recording a single fetched report file). Importing
this module registers the tables below on `Base.metadata`, so they are
created by `db.session.init_db()` as long as this module has been imported
at least once before `init_db()` runs -- see `db/__init__.py`.
"""

from __future__ import annotations

from sqlalchemy import JSON, ForeignKey, Numeric, String, Text
from sqlalchemy.orm import Mapped, mapped_column

from nessebar_budget.db.models import Base


class BudgetLineItem(Base):
    """A single row from a capital-expenditure ledger ("Разчет за финансиране
    на капиталовите разходи"), as parsed by `parsers.budget_capital`.

    Rows are denormalized: every object-level row carries its own
    `function_code`/`paragraph`/`subparagraph` context, and aggregate rows
    (paragraph subtotals, function subtotals, the sheet's grand "ОБЩО" row)
    are also stored, distinguished via `extra_json["row_type"]`.
    """

    __tablename__ = "budget_line_items"

    id: Mapped[int] = mapped_column(primary_key=True)

    report_id: Mapped[int | None] = mapped_column(ForeignKey("budget_reports.id"), default=None)

    period: Mapped[str | None] = mapped_column(String(32), default=None)
    unit: Mapped[str | None] = mapped_column(String(128), default=None)

    function_code: Mapped[str | None] = mapped_column(String(16), default=None)
    paragraph: Mapped[str | None] = mapped_column(String(16), default=None)
    subparagraph: Mapped[str | None] = mapped_column(String(16), default=None)

    object_code: Mapped[str | None] = mapped_column(String(32), default=None)
    object_name: Mapped[str | None] = mapped_column(Text, default=None)
    years: Mapped[str | None] = mapped_column(String(32), default=None)

    estimated_total: Mapped[float | None] = mapped_column(Numeric(18, 2), default=None)
    spent_prior: Mapped[float | None] = mapped_column(Numeric(18, 2), default=None)
    plan_current: Mapped[float | None] = mapped_column(Numeric(18, 2), default=None)
    spent_period: Mapped[float | None] = mapped_column(Numeric(18, 2), default=None)

    currency: Mapped[str | None] = mapped_column(String(8), default=None)
    extra_json: Mapped[dict | None] = mapped_column(JSON, default=None)


class CashExecutionLine(Base):
    """A single row from a "B1" cash-execution-by-paragraph report, as parsed
    by `parsers.budget_b1`.
    """

    __tablename__ = "cash_execution_lines"

    id: Mapped[int] = mapped_column(primary_key=True)

    report_id: Mapped[int | None] = mapped_column(ForeignKey("budget_reports.id"), default=None)

    period: Mapped[str | None] = mapped_column(String(32), default=None)
    section: Mapped[str | None] = mapped_column(String(32), default=None)

    paragraph: Mapped[str | None] = mapped_column(String(16), default=None)
    name: Mapped[str | None] = mapped_column(Text, default=None)

    plan_annual: Mapped[float | None] = mapped_column(Numeric(18, 2), default=None)
    plan_adjusted: Mapped[float | None] = mapped_column(Numeric(18, 2), default=None)
    actual_ytd: Mapped[float | None] = mapped_column(Numeric(18, 2), default=None)

    extra_json: Mapped[dict | None] = mapped_column(JSON, default=None)


class CashActivityLine(Base):
    """One money-carrying row of a B1/B3 report's per-activity expenditure
    pages ("II.1. РАЗХОДИ ПО ДЕЙНОСТИ"), as parsed by `parsers.budget_b1`:
    how much of a §§/под-§§ was planned and paid in one activity (дейност).

    Deliberately lean -- no `extra_json`, no names repeated per row, NULL
    instead of zero: there are ~700 of these per monthly report, the database is committed to git,
    and GitHub rejects files over 100 MB. `activity_name` is set only on the
    activity's total row (`paragraph` "99-99"); §§ names come from
    `CashExecutionLine`. `parent_paragraph` is NULL on a paragraph row and
    the paragraph's code on a sub-paragraph row.
    """

    __tablename__ = "cash_activity_lines"

    id: Mapped[int] = mapped_column(primary_key=True)

    report_id: Mapped[int | None] = mapped_column(
        ForeignKey("budget_reports.id"), default=None, index=True
    )

    period: Mapped[str | None] = mapped_column(String(32), default=None)
    activity_code: Mapped[str | None] = mapped_column(String(8), default=None)
    activity_name: Mapped[str | None] = mapped_column(Text, default=None)

    paragraph: Mapped[str | None] = mapped_column(String(16), default=None)
    parent_paragraph: Mapped[str | None] = mapped_column(String(16), default=None)

    plan: Mapped[float | None] = mapped_column(Numeric(18, 2), default=None)
    actual: Mapped[float | None] = mapped_column(Numeric(18, 2), default=None)

    # Who funds it: the report splits every amount into state-delegated
    # activities, local activities and local co-financing of state ones.
    # NULL where zero, which is most cells.
    plan_state: Mapped[float | None] = mapped_column(Numeric(18, 2), default=None)
    plan_local: Mapped[float | None] = mapped_column(Numeric(18, 2), default=None)
    plan_cofinance: Mapped[float | None] = mapped_column(Numeric(18, 2), default=None)
    actual_state: Mapped[float | None] = mapped_column(Numeric(18, 2), default=None)
    actual_local: Mapped[float | None] = mapped_column(Numeric(18, 2), default=None)
    actual_cofinance: Mapped[float | None] = mapped_column(Numeric(18, 2), default=None)

    source_row: Mapped[int | None] = mapped_column(default=None)


class FundLine(Base):
    """One money-carrying row of an EU-funds or third-party-funds ledger
    (`IB1_*`/`IB3_*` workbooks: DES, DMP, KSF, RA = EU-funds accounts, K33 =
    third-party funds), parsed by the same `parsers.budget_b1` as the cash
    report: the ledgers share its `OTCHET` layout.

    Table rows (`section` "приходи", "разходи", "трансфери", "временни
    заеми", "финансиране", "наличности") and per-activity rows (`section`
    "дейност") share one lean table: no JSON, only rows with a plan or paid
    amount plus each table's total. `name` is the row's own name for a
    table row, and the activity's name on an activity total row only.
    `programme_code`/`programme_name` are set on KSF activity rows (the
    same activity recurs under several programmes); `programme_name` on
    the activity total row only. Amounts are EUR.
    """

    __tablename__ = "fund_lines"

    id: Mapped[int] = mapped_column(primary_key=True)
    report_id: Mapped[int | None] = mapped_column(
        ForeignKey("budget_reports.id"), default=None, index=True
    )
    period: Mapped[str | None] = mapped_column(String(32), default=None)
    ledger: Mapped[str | None] = mapped_column(String(8), default=None)
    section: Mapped[str | None] = mapped_column(String(32), default=None)
    paragraph: Mapped[str | None] = mapped_column(String(16), default=None)
    parent_paragraph: Mapped[str | None] = mapped_column(String(16), default=None)
    name: Mapped[str | None] = mapped_column(Text, default=None)
    activity_code: Mapped[str | None] = mapped_column(String(8), default=None)
    programme_code: Mapped[str | None] = mapped_column(String(16), default=None)
    programme_name: Mapped[str | None] = mapped_column(Text, default=None)
    plan: Mapped[float | None] = mapped_column(Numeric(18, 2), default=None)
    actual: Mapped[float | None] = mapped_column(Numeric(18, 2), default=None)
    row_type: Mapped[str | None] = mapped_column(String(16), default=None)
