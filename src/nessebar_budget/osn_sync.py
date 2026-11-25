"""Copy newly published council decisions from os-nessebar.eu to the OSN site.

The council publishes each session's decisions as a web page, not a file.
OSN stores every session as a Chrome "print to PDF" copy of that page and
reads the protocol number and date from the file name, so this does the same
for the sessions listed above the newest one OSN already holds.

    python -m nessebar_budget.osn_sync            # render and upload
    python -m nessebar_budget.osn_sync --dry-run  # render only, upload nothing

Nothing is stored between runs: OSN's own list is the record of what exists.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections.abc import Callable
from pathlib import Path

import httpx
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

from nessebar_budget.config import get_settings

SOURCE_URL = "https://os-nessebar.eu/resheniya"
OUT_DIR = Path("data/cache/osn_decisions")  # gitignored; the workflow keeps it as an artifact
MAX_PAGES = 30  # the listing has 11 pages of 10; this only bounds a broken pager
_PROTOCOL = re.compile(r"№\s*(\d+)\s*[/_]\s*(\d{2})\.(\d{2})\.(\d{4})")

Key = tuple[int, str]


def protocol_key(text: str | None) -> Key | None:
    """(protocol number, ISO session date) from a title, or None if it has no date.

    The number alone is not unique: it restarts with every council term.
    """
    m = _PROTOCOL.search(text or "")
    return (int(m[1]), f"{m[4]}-{m[3]}-{m[2]}") if m else None


def pdf_name(key: Key) -> str:
    """The file name OSN parses the protocol and date from (Chrome's name for the page)."""
    number, (year, month, day) = key[0], key[1].split("-")
    return f"РЕШЕНИЯ ОТ ПРОТОКОЛ № {number}_{day}.{month}.{year} г. _ ОС Несебър.pdf"


def new_sessions(
    fetch_page: Callable[[int], str], have: set[Key]
) -> tuple[list[tuple[Key, str]], list[str]]:
    """Sessions listed above the newest one in `have`, oldest first, as (key, url).

    The listing is newest first, so reading stops at the first session OSN
    already has. Also returns the titles above that point that carry no
    protocol date and therefore cannot be matched or named.
    """
    new: list[tuple[Key, str]] = []
    unreadable: list[str] = []
    for page_no in range(1, MAX_PAGES + 1):
        links = BeautifulSoup(fetch_page(page_no), "lxml").select('h3 a[href*="/resheniya/"]')
        for a in links:
            title = " ".join(a.get_text().split())
            key = protocol_key(title)
            if key is None:
                unreadable.append(title)
            elif key in have:
                return new[::-1], unreadable
            else:
                new.append((key, a["href"]))
        if not links:
            break
    return new[::-1], unreadable


def osn_keys(client: httpx.Client, token: str | None) -> set[Key]:
    """What OSN already holds: the admin list when logged in, else the public one."""
    scope = "admin" if token else "public"
    r = client.get(f"/api/{scope}/os-documents/decision/all", headers=_auth(token))
    r.raise_for_status()
    keys = {protocol_key(d.get("protocol")) for d in r.json()}
    if None in keys:
        # Without a full picture of what exists, "new" could mean a duplicate.
        sys.exit("OSN lists a decision whose protocol cannot be read; refusing to upload.")
    return keys


def _auth(token: str | None) -> dict[str, str]:
    return {"Authorization": f"Bearer {token}"} if token else {}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--dry-run", action="store_true", help="render PDFs, upload nothing")
    parser.add_argument("--limit", type=int, default=0, help="send at most N, oldest first")
    args = parser.parse_args()
    dry_run = args.dry_run

    settings = get_settings()
    if not settings.osn_api_url:
        sys.exit("OSN_API_URL is not set.")
    if not dry_run and not (settings.osn_username and settings.osn_password):
        sys.exit("OSN_USERNAME / OSN_PASSWORD are not set (use --dry-run to only render).")

    # TLS verification stays on: this client sends the admin password.
    with httpx.Client(base_url=settings.osn_api_url, timeout=300) as osn:
        token = None
        if not dry_run:
            # shortcut: one login per run. If the token expires mid-backlog the run
            # fails and the next run resumes; add a re-login if that ever happens.
            r = osn.post(
                "/api/authenticate",
                json={"username": settings.osn_username, "password": settings.osn_password},
            )
            r.raise_for_status()
            token = r.json()["accessToken"]

        def fetch_page(page_no: int) -> str:
            r = httpx.get(SOURCE_URL, params={"page": page_no}, timeout=60, follow_redirects=True)
            r.raise_for_status()
            return r.text

        new, unreadable = new_sessions(fetch_page, osn_keys(osn, token))
        new = new[: args.limit or None]
        print(f"{len(new)} new session(s): {[f'{n}/{d}' for (n, d), _ in new]}")

        OUT_DIR.mkdir(parents=True, exist_ok=True)
        with sync_playwright() as pw:
            browser = pw.chromium.launch()
            page = browser.new_page()
            for key, url in new:  # oldest first, so a failure never leaves a gap
                response = page.goto(url, wait_until="networkidle")
                if not (response and response.ok) or protocol_key(page.title()) != key:
                    sys.exit(f"{url} did not load as protocol {key}; nothing uploaded for it.")
                # 0.4in is Chrome's print-dialog default, as in the copies already on OSN.
                edges = ("top", "right", "bottom", "left")
                pdf = page.pdf(format="A4", margin=dict.fromkeys(edges, "0.4in"))
                name = pdf_name(key)
                (OUT_DIR / name).write_bytes(pdf)
                if dry_run:
                    print(f"would upload {name} ({len(pdf) // 1024} KB)")
                    continue
                r = osn.post(
                    "/api/admin/os-documents/decision/upload",
                    headers=_auth(token),
                    files={"file": (name, pdf, "application/pdf")},
                )
                r.raise_for_status()
                # OSN derives protocol/date from the file on its side; confirm it read
                # them as expected before sending the next one.
                if key not in osn_keys(osn, token):
                    sys.exit(f"Uploaded {name} but OSN does not list protocol {key}; stopping.")
                print(f"uploaded {name}")
            browser.close()

    if unreadable:
        print(f"Titles without a protocol date, upload these by hand: {unreadable}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
