from nessebar_budget.osn_sync import (
    KINDS,
    invite_pdf_name,
    new_items,
    pdf_name,
    protocol_key,
    session_date,
)

INVITE, DECISION = KINDS


def _page(section: str, *titles: str) -> str:
    return "".join(
        f'<article><h3><a href="https://os-nessebar.eu{section}/s{i}">\n  {t}\n</a></h3></article>'
        for i, t in enumerate(titles)
    )


def test_protocol_key_reads_site_titles_and_osn_values():
    assert protocol_key("РЕШЕНИЯ ОТ ПРОТОКОЛ № 31/19.11.2026 г.") == (31, "2026-11-19")
    assert protocol_key("№ 26/05.05.2026 Г.") == (26, "2026-05-05")
    assert protocol_key("РЕШЕНИЯ ОТ ПРОТОКОЛ  № 12 /12.05.2017г.") == (12, "2017-05-12")
    assert protocol_key("РЕШЕНИЯ ОТ ПРОТОКОЛ № 27") is None


def test_pdf_name_round_trips_to_the_same_key():
    name = pdf_name((31, "2026-11-19"))
    assert name == "РЕШЕНИЯ ОТ ПРОТОКОЛ № 31_19.11.2026 г. _ ОС Несебър.pdf"
    assert protocol_key(name) == (31, "2026-11-19")


def test_invite_names_use_the_spelling_osn_accepts_whatever_the_title_looks_like():
    regular = "ПОКАНА ЗА РЕДОВНО ЗАСЕДАНИЕ НА ОБЩИНСКИ СЪВЕТ-НЕСЕБЪР 19.11.2026 Г."
    extraordinary = "П О К А Н А за извънредно заседание на Общински съвет-Несебър на 24.09.2026 г."
    assert session_date(regular) == "2026-11-19"
    assert session_date(extraordinary) == "2026-09-24"
    assert session_date("П О К А Н А за редовно заседание на Общински съвет-Несебър") is None
    # the first is, letter for letter, the name of the one invitation OSN already held
    assert invite_pdf_name("2026-05-05", regular) == (
        "ПОКАНА ЗА РЕДОВНО ЗАСЕДАНИЕ НА ОБЩИНСКИ СЪВЕТ-НЕСЕБЪР 05.05.2026 Г. _ ОС Несебър.pdf"
    )
    assert invite_pdf_name("2026-09-24", extraordinary) == (
        "ПОКАНА ЗА ИЗВЪНРЕДНО ЗАСЕДАНИЕ НА ОБЩИНСКИ СЪВЕТ-НЕСЕБЪР 24.09.2026 Г. _ ОС Несебър.pdf"
    )
    assert session_date(invite_pdf_name("2026-09-24", extraordinary)) == "2026-09-24"


def test_new_items_stops_at_the_newest_one_osn_has():
    section = DECISION.section
    pages = {
        1: _page(
            section,
            "РЕШЕНИЯ ОТ ПРОТОКОЛ № 28/04.09.2026 г.",
            "РЕШЕНИЯ ОТ ПРОТОКОЛ № 27",  # no date: cannot be matched
            "РЕШЕНИЯ ОТ ПРОТОКОЛ № 26/26.06.2026 г.",
        ),
        2: _page(
            section,
            "РЕШЕНИЯ ОТ ПРОТОКОЛ № 25/05.05.2026 г.",  # already on OSN
            "РЕШЕНИЯ ОТ ПРОТОКОЛ № 1/10.11.2023 г.",  # older and absent: not our job
        ),
    }
    fetched = []

    def fetch(n):
        fetched.append(n)
        return pages.get(n, "")

    new, unreadable = new_items(fetch, {(25, "2026-05-05")}, DECISION)

    assert [key for key, _, _ in new] == [(26, "2026-06-26"), (28, "2026-09-04")]  # oldest first
    assert new[0][2] == "https://os-nessebar.eu/resheniya/s2"
    assert unreadable == ["РЕШЕНИЯ ОТ ПРОТОКОЛ № 27"]
    assert fetched == [1, 2]


def test_new_items_ends_on_an_empty_page_when_osn_has_nothing():
    only_page = _page(DECISION.section, "РЕШЕНИЯ ОТ ПРОТОКОЛ № 2/17.11.2015 г.")
    new, _ = new_items(lambda n: only_page if n == 1 else "", set(), DECISION)
    assert [key for key, _, _ in new] == [(2, "2015-11-17")]


def test_new_invites_are_matched_by_session_date_and_ignore_other_sections():
    page = _page(
        INVITE.section,
        "ПОКАНА ЗА РЕДОВНО ЗАСЕДАНИЕ НА ОБЩИНСКИ СЪВЕТ-НЕСЕБЪР 19.11.2026 Г.",
        "П О К А Н А за извънредно заседание на Общински съвет-Несебър на 24.09.2026 г.",
        "ПОКАНА ЗА РЕДОВНО ЗАСЕДАНИЕ НА ОБЩИНСКИ СЪВЕТ-НЕСЕБЪР 05.05.2026 Г.",
    ) + _page(DECISION.section, "РЕШЕНИЯ ОТ ПРОТОКОЛ № 31/19.11.2026 г.")

    new, unreadable = new_items(lambda n: page if n == 1 else "", {"2026-05-05"}, INVITE)

    assert [(key, title[:9]) for key, title, _ in new] == [
        ("2026-09-24", "П О К А Н"),
        ("2026-11-19", "ПОКАНА ЗА"),
    ]
    assert unreadable == []
    assert INVITE.osn_key({"date": "2026-05-05"}) == "2026-05-05"
    assert DECISION.osn_key({"protocol": "№ 26/05.05.2026 Г."}) == (26, "2026-05-05")
