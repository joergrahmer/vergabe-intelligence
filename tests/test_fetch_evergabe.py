from unittest.mock import Mock, patch

import requests

from fetch_evergabe import get_page_title, main


def make_response(status_code, text=""):
    response = Mock()
    response.status_code = status_code
    response.text = text
    return response


def test_get_page_title_with_status_200_returns_title():
    response = make_response(200, "<html><head><title>Test Titel</title></head></html>")

    title = get_page_title(response)

    assert title == "Test Titel"


def test_get_page_title_with_non_200_status_returns_none():
    response = make_response(404, "<html></html>")

    title = get_page_title(response)

    assert title is None


def test_get_page_title_without_title_tag_returns_none():
    response = make_response(200, "<html><head></head></html>")

    title = get_page_title(response)

    assert title is None


def test_get_page_title_with_empty_body_returns_none():
    response = make_response(200, "")

    title = get_page_title(response)

    assert title is None


@patch("fetch_evergabe.requests.get")
def test_main_handles_connection_error_gracefully(mock_get, capsys):
    mock_get.side_effect = requests.exceptions.ConnectionError("Verbindung fehlgeschlagen")

    main()

    captured = capsys.readouterr()
    assert "Fehler" in captured.out
