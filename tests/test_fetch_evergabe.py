from unittest.mock import Mock, patch

import requests

from fetch_evergabe import get_search_results, main


def make_response(status_code, text=""):
    response = Mock()
    response.status_code = status_code
    response.text = text
    return response


def test_get_search_results_with_status_200_returns_titles():
    html = "<html><body>" + "".join(f"<h2>Treffer {i}</h2>" for i in range(1, 8)) + "</body></html>"
    response = make_response(200, html)

    titles = get_search_results(response)

    assert titles == [f"Treffer {i}" for i in range(1, 6)]


def test_get_search_results_with_non_200_status_returns_empty_list():
    response = make_response(503, "<html></html>")

    titles = get_search_results(response)

    assert titles == []


def test_get_search_results_without_headlines_returns_empty_list():
    response = make_response(200, "<html><body><p>Kein Treffer</p></body></html>")

    titles = get_search_results(response)

    assert titles == []


def test_get_search_results_skips_empty_headlines():
    html = "<html><body><h2></h2><h3>Echter Treffer</h3></body></html>"
    response = make_response(200, html)

    titles = get_search_results(response)

    assert titles == ["Echter Treffer"]


@patch("fetch_evergabe.requests.get")
def test_main_handles_connection_error_gracefully(mock_get, capsys):
    mock_get.side_effect = requests.exceptions.ConnectionError("Verbindung fehlgeschlagen")

    main()

    captured = capsys.readouterr()
    assert "Fehler" in captured.out


@patch("fetch_evergabe.requests.get")
def test_main_prints_numbered_results(mock_get, capsys):
    mock_get.return_value = make_response(200, "<html><body><h2>BAAINBw Ausschreibung</h2></body></html>")

    main()

    captured = capsys.readouterr()
    assert "1. BAAINBw Ausschreibung" in captured.out
