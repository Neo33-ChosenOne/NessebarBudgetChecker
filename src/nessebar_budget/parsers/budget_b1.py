"""Parser for the standard MinFin "B1" (monthly) and "B3" (quarterly)
municipal cash-execution report xls.

Layout, learned from `data/samples/nesebar_cash_execution_B1_2026_8.xls` and
cross-checked against the official MinFin макети (`data/minfin/maketi/`,
`.../касови отчети/B3_2026_3_Mun.xls` and its `IB3_*` siblings) and against
real `B1`/`B3` files cached from 2019 on: see `docs/sources/BUDGET_FORMS.md`.
`B3` quarterly workbooks share `B1`'s exact `OTCHET` schema (same column
layout, same §§/под-§§ header shape) -- the only difference is the reporting
cadence (quarter-end cumulative instead of month-end cumulative) and, for
quarters before 2026, the period's "от"/"до" header dates spanning the whole
year-to-date range rather than a single month -- so this parser accepts
both without any kind-specific branching.

This workbook concatenates *several* logical report pages onto one `OTCHET`
sheet, each re-printing the organisation header block before its own table:
"I. ПРИХОДИ..." (revenue by §§/под-§§), "II. РАЗХОДИ - РЕКАПИТУЛАЦИЯ..."
(expenditure by §§/под-§§), "III. ТРАНСФЕРИ" (money from and to other
budgets, chiefly the state's subsidies), "VI. ОПЕРАЦИИ С ФИНАНСОВИ АКТИВИ И
ПАСИВИ" (loans, deposits and the change in cash balances that finance the
difference), and then one "II.1. РАЗХОДИ ПО ДЕЙНОСТИ" page per activity
(дейност: kindergartens, street cleaning, the municipal administration,
...) repeating the expenditure table for that activity alone. This parser
extracts the first two tables in full, the money-carrying rows of the
transfers and financing tables (all four in `rows`, told apart by
`section`), and the money-carrying rows of every activity page
(`activities`) -- see BUDGET_FORMS.md.

Within each table, every row has (after the fixed `(a)` sort-key column in
position 0):
- col 1 (§§): filled only on a paragraph-aggregate row (its name then sits
  in col 2), OR the literal string "ВСИЧКО" marking the table's grand-total
  row (name in col 3, pseudo-code in col 2, e.g. "99-99").
- col 2 (под-§§): filled only on a subparagraph/detail row (its name then
  sits in col 3).
- cols 4-11: Уточнен план Общо, план by (държавни/местни/дофинансиране)
  дейности, ОТЧЕТ by the same three, ОТЧЕТ Общо.

Currency: the workbook carries an explicit "(в лева)" / "(в евро)" marker
(row 18, 1-indexed, in every file inspected so far, 2019-2026) just above
the header rows; this is checked first. When no marker is found, the period
decides (< 2026-01 -> BGN, >= 2026-01 -> EUR). Every monetary value is
converted to EUR at parse time (2 dp, at the fixed 1 EUR = 1.95583 BGN
rate); `extra_json` additionally records `original_currency`,
`conversion_rate`, and (only when a conversion actually happened) the
the conversion rate under `extra_json["conversion_rate"]`. `currency` is also
set on every row for forward-compatibility, even though the current
`CashExecutionLine` table/`upsert_cash_execution_lines` helper has no
`currency` column to persist it to (out of scope here -- see
`db/budget_models.py`); it is still recorded in `extra_json`, which *is*
persisted.

Provenance: every row's `extra_json` carries `source_sheet` ("OTCHET") and
`source_row` (1-based Excel row number) of the line it was read from.

Period: `detect_period(path)` reads the "до" date of the header's "за
периода от...до" pair; `parse-budget` stores that as the report's period
instead of the month the file was published under on nesebar.bg.
"""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Any

import pandas as pd

logger = logging.getLogger(__name__)

SHEET_NAME = "OTCHET"

COL_SORT = 0
COL_MAJOR = 1
COL_SUB = 2
COL_NAME_IF_MAJOR = 2
COL_NAME_IF_SUB = 3
COL_PLAN_TOTAL = 4
COL_PLAN_STATE = 5
COL_PLAN_LOCAL = 6
COL_PLAN_COFINANCE = 7
COL_ACTUAL_STATE = 8
COL_ACTUAL_LOCAL = 9
COL_ACTUAL_COFINANCE = 10
COL_ACTUAL_TOTAL = 11

_MONEY_FIELDS = (
    "plan_total",
    "plan_state",
    "plan_local",
    "plan_cofinance",
    "actual_state",
    "actual_local",
    "actual_cofinance",
    "actual_total",
)

_SECTIONS_IN_ORDER = ("приходи", "разходи")

#: The further tables, recognised by their total row's name (their position
#: among the report's pages is not relied on). "IV. ВРЕМЕННИ БЕЗЛИХВЕНИ
#: ЗАЕМИ" has no header row of its own: it follows the transfers total on
#: the same page. With these, the report's own identity holds:
#: приходи - разходи + трансфери + временни заеми = -финансиране.
_LATER_SECTIONS = (
    ("ВРЕМЕННИ", "временни заеми"),
    ("ТРАНСФЕРИ", "трансфери"),
    ("ФИНАНСОВИ", "финансиране"),
)

#: Sheet (present since 2021) that breaks the closing cash down by where it
#: came from; and the `section` its rows are stored under.
BALANCES_SHEET = "NALICHNOST"
BALANCES_SECTION = "наличности"

#: The row that opens an activity page: "<------  ДЕЙНОСТ  -  код  по  ЕБК"
#: in the name column, the activity's 4-digit code beside it. (Col 1 says
#: "оп/дейност" only since 2026, so the name-column text is what is matched.)
_ACTIVITY_MARKER_RE = re.compile(r"ДЕЙНОСТ\s*-\s*код", re.IGNORECASE)
#: Pseudo-code of an activity page's total row.
ACTIVITY_TOTAL = "99-99"
#: KSF (EU-funds) pages open differently: programme code + name, a "ПЕРИОД
#: 2014-2020" line, then the activity's code and "<code> <name>" row.
_PROGRAMME_PERIOD_RE = re.compile(r"^ПЕРИОД\b", re.IGNORECASE)

#: 1 EUR = 1.95583 BGN -- Bulgaria's official euro-adoption rate (2026-01-01).
BGN_PER_EUR = 1.95583

# "евро" must end the word: the EU-funds ledgers' title says "ЕВРОПЕЙСКИЯ СЪЮЗ".
_CURRENCY_EUR_RE = re.compile(r"\beur\b|евро(?![а-я])", re.IGNORECASE)
_CURRENCY_BGN_RE = re.compile(r"\bлв\b|лев[а-я]*", re.IGNORECASE)


def _num(value: Any) -> float | None:
    if value is None:
        return None
    try:
        if isinstance(value, float) and pd.isna(value):
            return None
    except TypeError:
        pass
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _text(value: Any) -> str | None:
    if value is None:
        return None
    try:
        if isinstance(value, float) and pd.isna(value):
            return None
    except TypeError:
        pass
    text = str(value).strip()
    return text or None


def _code_str(value: Any) -> str | None:
    """Render a §§/под-§§ code cell as a plain string ("100", not "100.0")."""
    if value is None:
        return None
    try:
        if isinstance(value, float) and pd.isna(value):
            return None
        if isinstance(value, (int, float)):
            return str(int(value))
    except (TypeError, ValueError):
        pass
    return str(value).strip() or None


def _convert(value: float | None, original_currency: str) -> float | None:
    """Convert a monetary value from `original_currency` to EUR (2 dp)."""
    if value is None:
        return None
    if original_currency == "EUR":
        return round(value, 2)
    return round(value / BGN_PER_EUR, 2)


def _detect_original_currency(df: pd.DataFrame, period: str | None) -> str:
    """Detect the workbook's original currency from its own "(в лева)" /
    "(в евро)" marker (seen at row 18, 1-indexed, in every file inspected,
    2019-2026), falling back to the period (< 2026-01 -> BGN, >= 2026-01 ->
    EUR) when no marker is found.
    """
    for i in range(min(len(df), 25)):
        for cell in df.iloc[i]:
            if not isinstance(cell, str):
                continue
            if _CURRENCY_EUR_RE.search(cell):
                return "EUR"
            if _CURRENCY_BGN_RE.search(cell):
                return "BGN"

    if period and period < "2026-01":
        return "BGN"
    if period:
        return "EUR"
    logger.warning("budget_b1: no currency marker and no period given; defaulting to BGN")
    return "BGN"


def _find_header_rows(df: pd.DataFrame) -> list[int]:
    """Return the row indices of every "§§ / под-§§ / НАИМЕНОВАНИЕ" header."""
    header_rows = []
    for i in range(len(df)):
        row = df.iloc[i]
        if _text(row[COL_MAJOR]) == "§§" and _text(row[COL_SUB]) == "под-§§":
            header_rows.append(i)
    return header_rows


def _row_values(row: pd.Series, original_currency: str) -> dict[str, float | None]:
    raw = {
        "plan_total": _num(row[COL_PLAN_TOTAL]),
        "plan_state": _num(row[COL_PLAN_STATE]),
        "plan_local": _num(row[COL_PLAN_LOCAL]),
        "plan_cofinance": _num(row[COL_PLAN_COFINANCE]),
        "actual_state": _num(row[COL_ACTUAL_STATE]),
        "actual_local": _num(row[COL_ACTUAL_LOCAL]),
        "actual_cofinance": _num(row[COL_ACTUAL_COFINANCE]),
        "actual_total": _num(row[COL_ACTUAL_TOTAL]),
    }
    converted = {key: _convert(value, original_currency) for key, value in raw.items()}
    return {"converted": converted, "raw": raw}


def _currency_extra(original_currency: str, raw: dict[str, float | None]) -> dict[str, Any]:
    extra: dict[str, Any] = {"original_currency": original_currency}
    if original_currency != "EUR":
        extra["conversion_rate"] = BGN_PER_EUR  # raw BGN = EUR value × rate
    return extra


def _parse_block(
    df: pd.DataFrame,
    header_row: int,
    section: str,
    period: str,
    original_currency: str,
    end_row: int | None = None,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    current_major: str | None = None

    r = header_row + 2  # skip the "(1) (2) ..." column-position row
    while r < (len(df) if end_row is None else end_row):
        row = df.iloc[r]
        major_raw = row[COL_MAJOR]
        major = _text(major_raw) if isinstance(major_raw, str) else _code_str(major_raw)
        sub = _code_str(row[COL_SUB])

        if major is None and sub is None:
            r += 1
            continue  # blank spacer or a footnote/"в т.ч." memo row with no code

        values = _row_values(row, original_currency)
        extra_base = _currency_extra(original_currency, values["raw"])
        # Provenance: `df` is read with header=None and no skipped rows, so
        # df index r is Excel row r + 1 of the OTCHET sheet.
        extra_base["source_sheet"] = SHEET_NAME
        extra_base["source_row"] = r + 1

        if isinstance(major_raw, str) and major_raw.strip().upper() == "ВСИЧКО":
            rows.append(
                {
                    "period": period,
                    "section": section,
                    "paragraph": _code_str(row[COL_SUB]) or "ВСИЧКО",
                    "name": _text(row[COL_NAME_IF_SUB]),
                    "plan_adjusted": values["converted"]["plan_total"],
                    "actual_ytd": values["converted"]["actual_total"],
                    "plan_annual": None,
                    "currency": "EUR",
                    "extra_json": {"row_type": "section_total", **values["converted"], **extra_base},
                }
            )
            break  # end of this block

        if major is not None:
            current_major = major
            rows.append(
                {
                    "period": period,
                    "section": section,
                    "paragraph": major,
                    "name": _text(row[COL_NAME_IF_MAJOR]),
                    "plan_adjusted": values["converted"]["plan_total"],
                    "actual_ytd": values["converted"]["actual_total"],
                    "plan_annual": None,
                    "currency": "EUR",
                    "extra_json": {"row_type": "paragraph", **values["converted"], **extra_base},
                }
            )
        elif sub is not None:
            rows.append(
                {
                    "period": period,
                    "section": section,
                    "paragraph": sub,
                    "name": _text(row[COL_NAME_IF_SUB]),
                    "plan_adjusted": values["converted"]["plan_total"],
                    "actual_ytd": values["converted"]["actual_total"],
                    "plan_annual": None,
                    "currency": "EUR",
                    "extra_json": {
                        "row_type": "subparagraph",
                        "parent_paragraph": current_major,
                        **values["converted"],
                        **extra_base,
                    },
                }
            )
        r += 1

    return rows


def _parse_activity_block(
    df: pd.DataFrame, header_row: int, end_row: int, period: str, original_currency: str
) -> list[dict[str, Any]] | None:
    """One "II.1. РАЗХОДИ ПО ДЕЙНОСТИ" page -> its rows that carry money.

    Returns None when the page is not an activity page (transfers,
    financing). Rows with neither a plan nor a paid amount are dropped --
    each page repeats all ~120 §§/под-§§ of the classification and most are
    empty -- and so is an activity with nothing at all. The activity's total
    row (paragraph `ACTIVITY_TOTAL`) carries its name; the other rows do not
    repeat it. On a KSF programme page the rows also carry `programme_code`
    (and, on the total row, `programme_name`); both are None on a plain page. Amounts are EUR, like everything this module returns.
    """
    marker = next(
        (
            r
            for r in range(header_row + 1, min(header_row + 8, end_row))
            if _ACTIVITY_MARKER_RE.search(_text(df.iat[r, COL_NAME_IF_SUB]) or "")
        ),
        None,
    )
    programme_code = programme_name = None
    if marker is not None:
        code = _code_str(df.iat[marker, COL_SUB])
        name_row, first = marker + 1, marker + 2
    else:
        period_row = next(
            (
                r
                for r in range(header_row + 2, min(header_row + 8, end_row - 2))
                if _PROGRAMME_PERIOD_RE.match(_text(df.iat[r, COL_NAME_IF_SUB]) or "")
            ),
            None,
        )
        if period_row is None:
            return None
        programme_code = _code_str(df.iat[period_row - 1, COL_SUB])
        programme_name = _text(df.iat[period_row - 1, COL_NAME_IF_SUB])
        code = _code_str(df.iat[period_row + 1, COL_SUB])
        name_row, first = period_row + 1, period_row + 2
        if code is None:
            return None
    # "122 Общинска администрация" -> the name; the number is the code's tail.
    name = re.sub(r"^\d+\s+", "", _text(df.iat[name_row, COL_NAME_IF_SUB]) or "") or None

    rows: list[dict[str, Any]] = []
    current_major: str | None = None
    for r in range(first, end_row):
        row = df.iloc[r]
        major_raw = row[COL_MAJOR]
        major = _text(major_raw) if isinstance(major_raw, str) else _code_str(major_raw)
        sub = _code_str(row[COL_SUB])
        is_total = sub == ACTIVITY_TOTAL
        if major is not None and not is_total:
            current_major = major
            paragraph, parent = major, None
        elif is_total or (sub is not None and sub.isdigit()):
            paragraph, parent = sub, (None if is_total else current_major)
        else:
            continue  # spacer, "(наименование на дейността)", or a memo row with no code

        values = _row_values(row, original_currency)["converted"]
        plan, actual = values["plan_total"], values["actual_total"]
        if plan or actual:
            rows.append(
                {
                    "period": period,
                    "activity_code": code,
                    "activity_name": name if is_total else None,
                    "programme_code": programme_code,
                    "programme_name": programme_name if is_total else None,
                    "paragraph": paragraph,
                    "parent_paragraph": parent,
                    "plan": plan,
                    "actual": actual,
                    # Who funds it: state-delegated, local, co-financed.
                    **{key: values[key] or None for key in _MONEY_FIELDS if "_total" not in key},
                    "source_row": r + 1,  # 1-based row of the OTCHET sheet
                }
            )
        if is_total:
            break
    return rows


def parse_budget_b1(path: str | Path, period: str | None = None) -> dict[str, Any]:
    """Parse the revenue and expenditure-by-paragraph tables of a B1/B3 xls,
    and the per-activity expenditure pages.

    Returns `{"rows": [...], "activities": [...], "skipped": [...]}`:
    `rows` are the two main tables, `activities` the money-carrying rows of
    every "РАЗХОДИ ПО ДЕЙНОСТИ" page (see `_parse_activity_block`), and
    `skipped` documents the report pages that were not parsed.
    """
    path = Path(path)
    xl = pd.ExcelFile(path, engine="xlrd")
    if SHEET_NAME not in xl.sheet_names:
        logger.warning("budget_b1: %s has no %r sheet; sheets=%s", path, SHEET_NAME, xl.sheet_names)
        return {"rows": [], "activities": [], "skipped": [f"no {SHEET_NAME!r} sheet"]}

    df = xl.parse(SHEET_NAME, header=None)

    header_rows = _find_header_rows(df)
    if not header_rows:
        logger.warning("budget_b1: no §§/под-§§ header row found in %s", path)
        return {"rows": [], "activities": [], "skipped": ["no header row found"]}

    inferred_period = period or _infer_period(df) or ""
    original_currency = _detect_original_currency(df, inferred_period)

    rows: list[dict[str, Any]] = []
    for section, header_row in zip(_SECTIONS_IN_ORDER, header_rows):
        rows.extend(_parse_block(df, header_row, section, inferred_period, original_currency))

    activities: list[dict[str, Any]] = []
    other_pages = 0
    for i in range(len(_SECTIONS_IN_ORDER), len(header_rows)):
        end_row = header_rows[i + 1] if i + 1 < len(header_rows) else len(df)
        block = _parse_activity_block(
            df, header_rows[i], end_row, inferred_period, original_currency
        )
        if block is not None:
            activities.extend(block)
            continue
        # Transfers (followed on the same page by temporary loans) or
        # financing: the same table shape as the main two. Only the rows
        # that carry money are kept (these tables list ~60 and ~140
        # §§/под-§§, a dozen of them in use), plus each total.
        start, recognised = header_rows[i], False
        while start + 2 < end_row:
            table = _parse_block(df, start, "", inferred_period, original_currency, end_row)
            total = (
                table[-1]
                if table and table[-1]["extra_json"]["row_type"] == "section_total"
                else None
            )
            section = next(
                (name for marker, name in _LATER_SECTIONS if total and marker in (total["name"] or "")),
                None,
            )
            if section is None:
                break
            recognised = True
            for row in table:
                if row is total or any(row["extra_json"].get(key) for key in _MONEY_FIELDS):
                    rows.append({**row, "section": section})
            start = total["extra_json"]["source_row"] - 2  # the block after this total
        if not recognised:
            other_pages += 1

    skipped = []
    if other_pages:
        skipped.append(
            f"{other_pages} further report page(s) on this sheet were not recognised and "
            "not parsed"
        )
    if len(header_rows) < len(_SECTIONS_IN_ORDER):
        skipped.append(
            f"expected {len(_SECTIONS_IN_ORDER)} report pages (приходи, разходи) but only "
            f"found {len(header_rows)}"
        )

    if BALANCES_SHEET in xl.sheet_names:
        rows.extend(
            _parse_balances(xl.parse(BALANCES_SHEET, header=None), inferred_period, original_currency)
        )

    return {"rows": rows, "activities": activities, "skipped": skipped}


def _parse_balances(df: pd.DataFrame, period: str, original_currency: str) -> list[dict[str, Any]]:
    """The `NALICHNOST` sheet: the cash held at the period's end, by where
    it came from (own revenue, state transfers, carried-over balance, ...).

    One table: a MinFin row code in col 0, the label in col 1, then pairs of
    (amount, of which on term deposits) for state activities, local
    activities, co-financing and the total (cols 2-9). Row 999 is the grand
    total; the 9999/99999 rows repeat it and check it. Rows are stored with
    the row code as `paragraph`, the total amount as `actual_ytd`, and the
    split in `extra_json`; rows without money are dropped.
    """
    rows: list[dict[str, Any]] = []
    for r in range(len(df)):
        code = _code_str(df.iat[r, 0])
        label = _text(df.iat[r, 1])
        if not code or not code.isdigit() or not label or int(code) >= 9999:
            continue
        # K33 (third-party funds) has a 5-column sheet: just one total in col 2.
        cols = range(2, 10) if df.shape[1] >= 10 else ()
        amounts = [_convert(_num(df.iat[r, c]), original_currency) for c in cols]
        if not cols:
            amounts = [None] * 6 + [_convert(_num(df.iat[r, 2]), original_currency), None]
        state, _, local, _, cofin, _, total, total_dep = amounts
        if not total:
            continue
        rows.append(
            {
                "period": period,
                "section": BALANCES_SECTION,
                "paragraph": code,
                "name": label.lstrip("- ").strip(),
                "plan_adjusted": None,
                "actual_ytd": total,
                "plan_annual": None,
                "currency": "EUR",
                "extra_json": {
                    "row_type": "section_total" if code == "999" else "balance",
                    "actual_state": state, "actual_local": local, "actual_cofinance": cofin,
                    "on_deposit": total_dep,
                    **_currency_extra(original_currency, {}),
                    "source_sheet": BALANCES_SHEET,
                    "source_row": r + 1,
                },
            }
        )
    return rows


def _is_date(value: Any) -> bool:
    return hasattr(value, "year") and hasattr(value, "month") and hasattr(value, "day")


def _infer_period(df: pd.DataFrame) -> str | None:
    """The period a B1/B3 report covers, from its own OTCHET header.

    The header reads "за периода от | до" with the two dates on the row
    below (e.g. 2019-01-01 | 2019-03-31). Both B1 and B3 are year-to-date
    cumulative, so "от" is always 1 January; the period is the "до" date's
    month. The cell under the "до" label is used when found; otherwise the
    *latest* date on the first header row that has any date (never the
    first one found -- that would report every file as January).
    """
    head = min(len(df), 25)
    for i in range(head):
        for j, cell in enumerate(df.iloc[i]):
            if not (isinstance(cell, str) and cell.strip().casefold() == "до"):
                continue
            for below in range(i + 1, min(i + 3, len(df))):
                value = df.iloc[below, j]
                if _is_date(value):
                    return f"{value.year:04d}-{value.month:02d}"

    best: Any = None
    for i in range(head):
        for cell in df.iloc[i]:
            if _is_date(cell) and (best is None or cell > best):
                best = cell
        if best is not None:
            break
    if best is None:
        return None
    return f"{best.year:04d}-{best.month:02d}"


def detect_period(path: str | Path) -> str | None:
    """Return the period ("YYYY-MM") a B1/B3 workbook *covers* -- the month
    of its "до" (to) date -- read from the OTCHET sheet's own header, not the
    month it was published under on nesebar.bg. A quarterly `B3_2019_1`
    gives "2019-03" (Q1), a `B3_2018_4` posted in January 2019 gives
    "2018-12". Returns `None` if the file can't be read or states no date.
    """
    try:
        df = pd.read_excel(Path(path), sheet_name=SHEET_NAME, header=None, nrows=25, engine="xlrd")
    except Exception as exc:  # noqa: BLE001 -- any unreadable file just has no period
        logger.warning("budget_b1: could not read %s to detect its period: %s", path, exc)
        return None
    return _infer_period(df)
