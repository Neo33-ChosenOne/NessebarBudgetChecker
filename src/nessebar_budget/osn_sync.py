"""Copy newly published council documents from os-nessebar.eu to the OSN site.

Two kinds, handled the same way: the invitation (покана) to each session and
the decisions (решения) it took. The council publishes both as web pages, not
files. OSN stores each as a Chrome "print to PDF" copy of the page and reads
the session date (and, for decisions, the protocol number) from the file name,
so this does the same for whatever is listed above the newest one OSN holds.

    python -m nessebar_budget.osn_sync            # render and upload
    python -m nessebar_budget.osn_sync --dry-run  # render only, upload nothing

Nothing is stored between runs: OSN's own lists are the record of what exists.
"""

from __future__ import annotations

import argparse
import re
import sys
import traceback
from collections.abc import Callable, Hashable
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import httpx
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

from nessebar_budget.config import get_settings

SITE = "https://os-nessebar.eu"
OUT_DIR = Path("data/cache/osn_sync")  # gitignored; the workflow keeps it as an artifact
MAX_PAGES = 30  # the listings have 11 pages of 10; this only bounds a broken pager
# The council holds about one session a month. More new items than this means the
# comparison with OSN broke, and uploading them would flood the site with duplicates.
MAX_UNATTENDED = 12
_DATE = r"(\d{2})\.(\d{2})\.(\d{4})"
_PROTOCOL = re.compile(rf"№\s*(\d+)\s*[/_]\s*{_DATE}")


def protocol_key(text: str | None) -> tuple[int, str] | None:
    """(protocol number, ISO session date) from a title, or None if it has no date.

    The number alone is not unique: it restarts with every council term.
    """
    m = _PROTOCOL.search(text or "")
    return (int(m[1]), f"{m[4]}-{m[3]}-{m[2]}") if m else None


def session_date(text: str | None) -> str | None:
    """ISO session date from an invitation title, or None if it carries none."""
    m = re.search(_DATE, text or "")
    return f"{m[3]}-{m[2]}-{m[1]}" if m else None


def pdf_name(key: tuple[int, str]) -> str:
    """The decisions file name OSN parses protocol and date from (Chrome's name for the page)."""
    number, (year, month, day) = key[0], key[1].split("-")
    return f"РЕШЕНИЯ ОТ ПРОТОКОЛ № {number}_{day}.{month}.{year} г. _ ОС Несебър.pdf"


def invite_pdf_name(date: str, title: str) -> str:
    """The invitation file name, in the one spelling OSN is known to accept.

    The council's own titles vary ("П О К А Н А за ...", "Покана - за ..."), so only
    the date and whether the session is extraordinary are taken from them.
    """
    year, month, day = date.split("-")
    sort = "ИЗВЪНРЕДНО" if "извънредно" in title.lower() else "РЕДОВНО"
    return (
        f"ПОКАНА ЗА {sort} ЗАСЕДАНИЕ НА ОБЩИНСКИ СЪВЕТ-НЕСЕБЪР {day}.{month}.{year} Г. _ ОС Несебър.pdf"
    )


@dataclass(frozen=True)
class Kind:
    name: str  # also the path segment of OSN's API for this kind
    section: str  # path of the council's listing
    key: Callable[[str | None], Hashable | None]  # identity of an item, from its title
    osn_key: Callable[[dict[str, Any]], Hashable | None]  # the same identity, from OSN's record
    pdf_name: Callable[[Any, str], str]  # (key, title) -> file name


KINDS = (
    # Invitations first: they only matter before the session.
    Kind(
        "invite",
        "/zasedaniya-na-obshtinski-savet-i-materiali",
        session_date,
        lambda record: record.get("date"),
        invite_pdf_name,
    ),
    Kind(
        "decision",
        "/resheniya",
        protocol_key,
        lambda record: protocol_key(record.get("protocol")),
        lambda key, _title: pdf_name(key),
    ),
)


def new_items(
    fetch_page: Callable[[int], str], have: set[Hashable], kind: Kind
) -> tuple[list[tuple[Hashable, str, str]], list[str]]:
    """Items listed above the newest one in `have`, oldest first, as (key, title, url).

    The listing is newest first, so reading stops at the first item OSN already
    has. Also returns the titles above that point that carry no date and
    therefore cannot be matched or named.
    """
    new: list[tuple[Hashable, str, str]] = []
    unreadable: list[str] = []
    for page_no in range(1, MAX_PAGES + 1):
        links = BeautifulSoup(fetch_page(page_no), "lxml").select(f'h3 a[href*="{kind.section}/"]')
        for a in links:
            title = " ".join(a.get_text().split())
            key = kind.key(title)
            if key is None:
                unreadable.append(title)
            elif key in have:
                return new[::-1], unreadable
            else:
                new.append((key, title, a["href"]))
        if not links:
            break
    return new[::-1], unreadable


def osn_keys(client: httpx.Client, token: str | None, kind: Kind) -> set[Hashable]:
    """What OSN already holds of one kind: the admin list when logged in, else the public one."""
    if token:
        r = client.get(f"/api/admin/os-documents/{kind.name}/all", headers=_auth(token))
        r.raise_for_status()
        records = r.json()
    else:
        r = client.get("/api/public/os-documents/decision/all")
        r.raise_for_status()
        records = r.json()
        if kind.name == "invite":
            # shortcut: the public API shows only invitations already tied to a decision,
            # so a login-less dry run can re-list an upcoming session's. Real runs log in.
            records = [d["invite"] for d in records if d.get("invite")]
    keys = {kind.osn_key(record) for record in records}
    if None in keys:
        # Without a full picture of what exists, "new" could mean a duplicate.
        raise RuntimeError(f"OSN lists a {kind.name} that cannot be identified; refusing to upload.")
    return keys


def _auth(token: str | None) -> dict[str, str]:
    return {"Authorization": f"Bearer {token}"} if token else {}


def _fetch_listing(kind: Kind, page_no: int) -> str:
    r = httpx.get(SITE + kind.section, params={"page": page_no}, timeout=60, follow_redirects=True)
    r.raise_for_status()
    return r.text


def _sync(kind: Kind, osn: httpx.Client, token: str | None, page: Any, limit: int, dry_run: bool):
    new, unreadable = new_items(lambda n: _fetch_listing(kind, n), osn_keys(osn, token, kind), kind)
    if len(new) > MAX_UNATTENDED and not limit:
        raise RuntimeError(f"{len(new)} new items is implausible; rerun with --limit to override.")
    new = new[: limit or None]
    print(f"{kind.name}: {len(new)} new {[key for key, _, _ in new]}")

    for key, title, url in new:  # oldest first, so a failure never leaves a gap
        response = page.goto(url, wait_until="networkidle")
        if not (response and response.ok) or kind.key(page.title()) != key:
            raise RuntimeError(f"{url} did not load as {key}; nothing uploaded for it.")
        # 0.4in is Chrome's print-dialog default, as in the copies already on OSN.
        edges = ("top", "right", "bottom", "left")
        pdf = page.pdf(format="A4", margin=dict.fromkeys(edges, "0.4in"))
        name = kind.pdf_name(key, title)
        (OUT_DIR / name).write_bytes(pdf)
        if dry_run:
            print(f"would upload {name} ({len(pdf) // 1024} KB)")
            continue
        r = osn.post(
            f"/api/admin/os-documents/{kind.name}/upload",
            headers=_auth(token),
            files={"file": (name, pdf, "application/pdf")},
        )
        r.raise_for_status()
        # OSN derives the identity from the file on its side; confirm it read it as
        # expected before sending the next one.
        if key not in osn_keys(osn, token, kind):
            raise RuntimeError(f"Uploaded {name} but OSN does not list {key}; stopping.")
        print(f"uploaded {name}")

    if unreadable:
        raise RuntimeError(f"Titles without a date, upload these by hand: {unreadable}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--dry-run", action="store_true", help="render PDFs, upload nothing")
    parser.add_argument("--limit", type=int, default=0, help="send at most N per kind, oldest first")
    args = parser.parse_args()

    settings = get_settings()
    if not settings.osn_api_url:
        sys.exit("OSN_API_URL is not set.")
    have_login = bool(settings.osn_username and settings.osn_password)
    if not args.dry_run and not have_login:
        sys.exit("OSN_USERNAME / OSN_PASSWORD are not set (use --dry-run to only render).")

    failed = False
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    # TLS verification stays on: this client sends the admin password.
    with httpx.Client(base_url=settings.osn_api_url, timeout=300) as osn, sync_playwright() as pw:
        token = None
        if have_login:
            # shortcut: one login per run. If the token expires mid-backlog the run
            # fails and the next run resumes; add a re-login if that ever happens.
            r = osn.post(
                "/api/authenticate",
                json={"username": settings.osn_username, "password": settings.osn_password},
            )
            r.raise_for_status()
            token = r.json()["accessToken"]

        browser = pw.chromium.launch()
        page = browser.new_page()
        for kind in KINDS:
            try:
                _sync(kind, osn, token, page, args.limit, args.dry_run)
            except Exception:  # noqa: BLE001 -- one kind failing must not hold back the other
                traceback.print_exc()
                failed = True
        browser.close()
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
