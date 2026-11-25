"""Parser for the municipality's annual natural-indicator workbooks.

The MinFin form "РАЗШИФРОВКА НА НАТУРАЛНИТЕ ПОКАЗАТЕЛИ ПО БЮДЖЕТА НА ОБЩИНИТЕ"
lists head-counts (staff, mayors, councillors, ...) and service volumes
(children, pupils, places, km, ...) for the whole municipality, for each
function, and for each budget activity. Combined with the salary spending in
`cash_activity_lines` it lets the site show how many people work in an
activity and the average pay there.

Two layouts exist (verified by reading the cells of the 2019, 2021, 2023,
2024 and 2026 files):

* **old** (2019 `.xlsx`): name in column 0, code like "01-00" in column 1,
  then plan / total / state / local / co-financing in columns 2-6. A block is
  a free-text title row followed by a "НАТУРАЛНИ ПОКАЗАТЕЛИ:" row; the title of
  an activity block carries the NAME only, no activity code.
* **new** (2021+ `.xls`): every data row is self-describing: a tag
  ("recapitulation" / "by function" / "by activity"), the function digit, the
  4-digit activity code, the code ("100", "111", ...) and total / state /
  local / co-financing (no plan column). The activity code is therefore read
  from the cell, not guessed from the title. 2021 starts at column 0; 2023+
  has four leading identification columns, so everything is found by header
  position, not hard-coded.

Old codes are normalised to the new scheme ("01-00" -> "100", "87-00" ->
"8700") so one indicator has one code across years.

Whether a file is a finished-year REPORT or a budgeted PLAN is not stated in
the new form (its column is headed "ОТЧЕТ/ГОДИШЕН ПЛАН"), so it is decided
from the 2019 header when unambiguous, otherwise from the hint passed in
(the site publishes plans under `/budget/` and reports under `/otcheti/`),
otherwise from whether the period has already ended. `basis_source` records
which rule fired.

Workbooks are treated as data only.
"""

from __future__ import annotations

import datetime as dt
import difflib
import re
from pathlib import Path
from typing import Any

# Latin letters that the municipality's files mix into Cyrillic words
# (e.g. "ОБЩИНСKА" with a Latin K, "ДEЖУPHИ"); folded to Cyrillic for names
# and matching so the same word compares equal.
_LATIN_TO_CYR = str.maketrans("ABCEHKMOPTXYabcehkmoptxy", "АВСЕНКМОРТХУавсенкмортху")

_FUNCTION_DIGITS = {
    "ОБЩИ ДЪРЖАВНИ СЛУЖБИ": "1",
    "ОТБРАНА И СИГУРНОСТ": "2",
    "ОБРАЗОВАНИЕ": "3",
    "ЗДРАВЕОПАЗВАНЕ": "4",
    "СОЦИАЛНО ОСИГУРЯВАНЕ": "5",
    "ЖИЛИЩНО СТРОИТЕЛСТВО": "6",
    "КУЛТУРА": "7",
    "ИКОНОМИЧЕСКИ ДЕЙНОСТИ": "8",
}

# Block titles that cannot be matched by name; extend when `unmapped` lists one.
# Key: normalised title (see `_norm`), value: 4-digit activity code.
ACTIVITY_TITLE_OVERRIDES: dict[str, str] = {}

STAFF_CODE = "100"  # "ЩАТНА ЧИСЛЕНОСТ" (old form: "01-00")


def _s(v: Any) -> str:
    return "" if v is None else str(v).strip()


def _num(v: Any) -> float | None:
    if isinstance(v, bool) or v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip().replace(",", "."))
    except ValueError:
        return None


def _read_rows(path: Path) -> tuple[list[list[Any]], int]:
    """First sheet as a list of rows (1-based row n is index n-1), plus the
    xls date mode (0 for xlsx, where dates arrive as datetimes)."""
    if path.suffix.lower() == ".xls":
        import xlrd

        wb = xlrd.open_workbook(str(path))
        sh = wb.sheet_by_index(0)
        return [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(sh.nrows)], wb.datemode
    import openpyxl

    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return rows, 0


def clean_indicator_name(raw: str) -> str:
    """Collapse the form's letter-spacing ("Щ а т н и   б р о й к и" ->
    "Щатни бройки"), fold Latin look-alikes, normalise whitespace."""
    text = re.sub(r"\s+", lambda m: "  " if len(m.group()) > 1 else " ", raw.strip())
    words = []
    for chunk in text.split("  "):
        toks = chunk.split(" ")
        words.append("".join(toks) if len(toks) > 2 and all(len(t) == 1 for t in toks) else chunk)
    return " ".join(" ".join(words).split()).translate(_LATIN_TO_CYR)


def _norm(title: str) -> str:
    """Matching key: no leading number, Cyrillic-folded, upper case, letters
    and digits only."""
    t = re.sub(r"^\s*\d+\s*[.)]?\s*", "", title.translate(_LATIN_TO_CYR)).upper()
    return " ".join(re.sub(r"[^0-9А-ЯA-Z]+", " ", t).split())


def _function_digit(heading: str) -> str | None:
    h = _norm(heading)
    for name, d in _FUNCTION_DIGITS.items():
        if name in h:
            return d
    return None


def _to_date(v: Any, datemode: int) -> dt.date | None:
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, (int, float)) and v > 30000:
        import xlrd

        return xlrd.xldate_as_datetime(v, datemode).date()
    m = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", _s(v))
    return dt.date(int(m[3]), int(m[2]), int(m[1])) if m else None


def _period(rows: list[list[Any]], datemode: int) -> tuple[dt.date | None, dt.date | None]:
    """Find the "от"/"до" labels in the top rows; the date is the next
    non-empty cell to the right."""
    found: dict[str, dt.date] = {}
    for row in rows[8:13]:
        for i, cell in enumerate(row):
            label = _s(cell).lower().replace("за периода:", "").strip()
            if label in ("от", "до") and label not in found:
                for nxt in row[i + 1 :]:
                    if _s(nxt):
                        d = _to_date(nxt, datemode)
                        if d:
                            found[label] = d
                        break
    return found.get("от"), found.get("до")


def _decide_basis(rows: list[list[Any]], hint: str | None, end: dt.date | None) -> tuple[str, str]:
    cells = {_s(c).lower() for r in rows[13:17] for c in r}
    if "отчет" in cells:  # 2019 form: "Годишен" / "Отчет" column labels
        return "отчет", "header says Отчет"
    if hint:
        h = hint.lower()
        if "/budget/" in h or Path(h).name.startswith("budget_"):
            return "план", "path/URL is in the budget folder"
        if "/otcheti/" in h or Path(h).name.startswith("otcheti_"):
            return "отчет", "path/URL is in the otcheti folder"
    if end and end > dt.datetime.now(dt.UTC).date():
        return "план", "period has not ended"
    return "отчет", "default: finished period, no other evidence"


def parse_naturi(path: str | Path, source_url: str | None = None) -> dict:
    """Parse one natural-indicators workbook.

    Returns `{"year", "basis" ("отчет"|"план"), "basis_source", "period_text",
    "form" ("old"|"new"), "rows": [...]}`. Each row has `level`
    (municipality / function / group / activity), `block_title` (as
    written, trimmed), `function_code` (digit or None), `activity_code`
    (4 digits for new-form activity blocks, else None), `indicator_code`,
    `indicator_name`, and `plan`, `total`, `state`, `local`, `cofinance`
    (None when the cell is empty). Only indicators with a non-zero value are
    kept. Old-form activity blocks have `activity_code=None`; resolve them
    with `map_activity_codes`.

    `source_url` (or the file name) is only used to tell plan from report
    when the workbook itself does not say.
    """
    path = Path(path)
    rows, datemode = _read_rows(path)
    hdr = next(i for i, r in enumerate(rows[:40]) if any(_s(c) == "§§" for c in r))
    code_col = next(i for i, c in enumerate(rows[hdr]) if _s(c) == "§§")
    old = _s(rows[hdr][code_col + 1]).lower() == "план"
    if old:
        name_c, tag_c, plan_c, tot_c = 0, None, code_col + 1, code_col + 2
    else:
        name_c = next(i for i, c in enumerate(rows[hdr]) if _s(c).startswith("ПОКАЗАТЕЛИ"))
        tag_c, plan_c, tot_c = name_c + 1, None, code_col + 1
    func_c, act_c = (tag_c + 1, tag_c + 2) if tag_c is not None else (None, None)

    start, end = _period(rows, datemode)
    basis, basis_source = _decide_basis(rows, source_url or str(path), end)
    period_text = f"{start:%d.%m.%Y} - {end:%d.%m.%Y}" if start and end else ""

    def cell(row: list[Any], i: int) -> Any:
        return row[i] if i < len(row) else None

    out: list[dict] = []
    pending: list[tuple[int, str]] = []  # title-like lines since the last data row
    cur_key: tuple | None = None
    cur: dict = {}
    func_digit: str | None = None
    for r in range(hdr + 1, len(rows)):
        row = rows[r]
        name = _s(cell(row, name_c))
        code_raw = _s(cell(row, code_col))
        tag = _s(cell(row, tag_c)) if tag_c is not None else ""
        is_data = bool(name) and (
            bool(tag) if not old else bool(re.fullmatch(r"\d\d-\d\d", code_raw)) and code_raw != "99-99"
        )
        if not is_data:
            if name and not any(_s(c) for i, c in enumerate(row) if i != name_c):
                if name.startswith(("НАТУРАЛНИ ПОКАЗАТЕЛИ", "СТАТУС")):
                    if old and "РЕКАПИТУЛАЦИЯ" in name and name.startswith("НАТУРАЛНИ"):
                        pending.append((r, name))
                    continue
                pending.append((r, name))
                if re.match(r"\s*ФУНКЦИЯ\b", name):
                    func_digit = _function_digit(name)
            continue
        # a data row: decide which block it belongs to
        if old:
            title = ""
            if pending:
                title = pending[-1][1]
                if len(pending) > 1 and pending[-2][0] == pending[-1][0] - 1 and pending[-2][1].lstrip().startswith(
                    "РЕКАПИТУЛАЦИЯ"
                ):
                    title = f"{pending[-2][1].strip()} {title.strip()}"
            up = " ".join(title.upper().split())
            if "ПОКАЗАТЕЛИ" in up and "РЕКАПИТУЛАЦИЯ" in up and "ФУНКЦИЯ" not in up and "ГРУПА" not in up:
                level = "municipality"
            elif up.startswith("РЕКАПИТУЛАЦИЯ ЗА ФУНКЦИЯ"):
                level = "function"
            elif up.startswith(("РЕКАПИТУЛАЦИЯ ЗА ГРУПА", "ГРУПА")):
                level = "group"
            else:
                level = "activity"
            key: tuple = (level, title.strip()) if pending else cur_key  # type: ignore[assignment]
            if pending:
                cur = {"level": level, "block_title": title.strip(), "function_code": func_digit,
                       "activity_code": None}
                if level in ("function", "group"):
                    cur["function_code"] = _function_digit(title) or func_digit
        else:
            tag_level = {"recapitulation": "municipality", "by function": "function", "by activity": "activity"}
            level = tag_level.get(tag, tag)
            fn = _num(cell(row, func_c))
            ac = _num(cell(row, act_c))
            key = (tag, fn, ac)
            if key != cur_key:
                if level == "municipality":
                    title = "НАТУРАЛНИ ПОКАЗАТЕЛИ - РЕКАПИТУЛАЦИЯ"
                else:
                    title = pending[-1][1] if pending else ""
                    if len(pending) > 1 and pending[-2][0] == pending[-1][0] - 1 and pending[-2][1].lstrip().startswith(
                        "РЕКАПИТУЛАЦИЯ"
                    ):
                        title = f"{pending[-2][1].strip()} {title.strip()}"
                cur = {
                    "level": level,
                    "block_title": title.strip(),
                    "function_code": str(int(fn)) if fn else None,
                    "activity_code": str(int(ac)) if ac else None,
                }
        cur_key = key
        pending = []

        vals = {
            "plan": _num(cell(row, plan_c)) if plan_c is not None else None,
            "total": _num(cell(row, tot_c)),
            "state": _num(cell(row, tot_c + 1)),
            "local": _num(cell(row, tot_c + 2)),
            "cofinance": _num(cell(row, tot_c + 3)),
        }
        if not any(v for v in vals.values()):
            continue
        code = str(int(code_raw.replace("-", ""))) if old else str(int(_num(cell(row, code_col)) or 0))
        out.append({**cur, "indicator_code": code, "indicator_name": clean_indicator_name(name), **vals})

    return {
        "year": end.year if end else None,
        "basis": basis,
        "basis_source": basis_source,
        "period_text": period_text,
        "form": "old" if old else "new",
        "rows": out,
    }


def map_activity_codes(
    rows: list[dict], names: dict[str, str]
) -> tuple[dict[str, tuple[str, str]], list[str]]:
    """Resolve the activity code of old-form activity blocks by title.

    `names` maps an activity NAME (any spelling, any year) to its 4-digit
    code; build it from `cash_activity_lines` plus the numbered titles of the
    newer workbooks. Matching is exact on `_norm`, then a manual override,
    then a close fuzzy match (>= 0.9) restricted to the block's function
    digit when known. Returns (`{title: (code, method)}`, unmapped titles);
    an unmapped title is never guessed.
    """
    index: dict[str, str] = {}
    for n, c in names.items():
        index.setdefault(_norm(n), c)
    mapping: dict[str, tuple[str, str]] = {}
    unmapped: list[str] = []
    seen: dict[str, str | None] = {}
    for r in rows:
        if r["level"] != "activity" or r.get("activity_code") or r["block_title"] in seen:
            continue
        title, fd = r["block_title"], r.get("function_code")
        seen[title] = fd
        key = _norm(title)
        code = index.get(key)
        how = "exact"
        if code is None and key in ACTIVITY_TITLE_OVERRIDES:
            code, how = ACTIVITY_TITLE_OVERRIDES[key], "override"
        if code is None:
            cands = {k: c for k, c in index.items() if not fd or c.startswith(fd)}
            best = difflib.get_close_matches(key, list(cands), n=1, cutoff=0.9)
            if best:
                code, how = cands[best[0]], "fuzzy"
        if code is None:
            unmapped.append(title)
        else:
            mapping[title] = (code, how)
    return mapping, unmapped
