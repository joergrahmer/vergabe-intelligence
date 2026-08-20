from unittest.mock import patch

from fetch_evergabe import extract_results, main

BASE_URL = "https://www.evergabe-online.de/search/awardedProcedure.html"


def make_result_link(index):
    return (
        f'<a class="text-wrap" href="../contractAward.html?id={index}">'
        f"Treffer {index}</a>"
    )


def test_extract_results_returns_titles_and_urls():
    html = "<html><body>" + "".join(make_result_link(i) for i in range(1, 8)) + "</body></html>"

    results = extract_results(html, BASE_URL)

    assert results == [
        {
            "title": f"Treffer {i}",
            "url": f"https://www.evergabe-online.de/contractAward.html?id={i}",
        }
        for i in range(1, 6)
    ]


def test_extract_results_without_results_returns_empty_list():
    html = "<html><body><p>Kein Treffer</p></body></html>"

    results = extract_results(html, BASE_URL)

    assert results == []


def test_extract_results_skips_links_without_title_or_href():
    html = (
        '<html><body>'
        '<a class="text-wrap" href="../contractAward.html?id=0"></a>'
        '<a class="text-wrap">Ohne Href</a>'
        f'{make_result_link(1)}'
        '</body></html>'
    )

    results = extract_results(html, BASE_URL)

    assert results == [
        {"title": "Treffer 1", "url": "https://www.evergabe-online.de/contractAward.html?id=1"}
    ]


def test_extract_results_ignores_navigation_links():
    html = (
        '<html><body>'
        '<a href="/search.html">Ausschreibungen suchen</a>'
        '<a class="text-wrap" href="../tenderdetails.html?id=99">Kein Zuschlag</a>'
        f'{make_result_link(1)}'
        '</body></html>'
    )

    results = extract_results(html, BASE_URL)

    assert results == [
        {"title": "Treffer 1", "url": "https://www.evergabe-online.de/contractAward.html?id=1"}
    ]


@patch("fetch_evergabe.fetch_search_results_html")
def test_main_handles_fetch_error_gracefully(mock_fetch, capsys):
    mock_fetch.side_effect = RuntimeError("Timeout")

    main()

    captured = capsys.readouterr()
    assert "Fehler" in captured.out


@patch("fetch_evergabe.fetch_search_results_html")
def test_main_prints_titles_and_urls(mock_fetch, capsys):
    mock_fetch.return_value = (f"<html><body>{make_result_link(1)}</body></html>", BASE_URL)

    main()

    captured = capsys.readouterr()
    assert "Titel: Treffer 1" in captured.out
    assert "URL: https://www.evergabe-online.de/contractAward.html?id=1" in captured.out
    assert "---" in captured.out


@patch("fetch_evergabe.fetch_search_results_html")
def test_main_prints_message_when_no_results(mock_fetch, capsys):
    mock_fetch.return_value = ("<html><body></body></html>", BASE_URL)

    main()

    captured = capsys.readouterr()
    assert "Keine Treffer gefunden." in captured.out
