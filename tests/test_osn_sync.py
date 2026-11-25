from nessebar_budget.osn_sync import new_sessions, pdf_name, protocol_key


def _page(*titles: str) -> str:
    return "".join(
        f'<article><h3><a href="https://os-nessebar.eu/resheniya/s{i}">\n  {t}\n</a></h3></article>'
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


def test_new_sessions_stops_at_the_newest_one_osn_has():
    pages = {
        1: _page(
            "РЕШЕНИЯ ОТ ПРОТОКОЛ № 28/04.09.2026 г.",
            "РЕШЕНИЯ ОТ ПРОТОКОЛ № 27",  # no date: cannot be matched
            "РЕШЕНИЯ ОТ ПРОТОКОЛ № 26/26.06.2026 г.",
        ),
        2: _page(
            "РЕШЕНИЯ ОТ ПРОТОКОЛ № 25/05.05.2026 г.",  # already on OSN
            "РЕШЕНИЯ ОТ ПРОТОКОЛ № 1/10.11.2023 г.",  # older and absent: not our job
        ),
    }
    fetched = []

    def fetch(n):
        fetched.append(n)
        return pages.get(n, "")

    new, unreadable = new_sessions(fetch, {(25, "2026-05-05")})

    assert [key for key, _ in new] == [(26, "2026-06-26"), (28, "2026-09-04")]  # oldest first
    assert new[0][1] == "https://os-nessebar.eu/resheniya/s2"
    assert unreadable == ["РЕШЕНИЯ ОТ ПРОТОКОЛ № 27"]
    assert fetched == [1, 2]


def test_new_sessions_ends_on_an_empty_page_when_osn_has_nothing():
    only_page = _page("РЕШЕНИЯ ОТ ПРОТОКОЛ № 2/17.11.2015 г.")
    new, _ = new_sessions(lambda n: only_page if n == 1 else "", set())
    assert [key for key, _ in new] == [(2, "2015-11-17")]
