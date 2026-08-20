import requests
from bs4 import BeautifulSoup

SEARCH_URL = "https://www.evergabe-online.de/search.html"
SEARCH_STRING = "BAAINBw"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )
}


def get_search_results(response, limit=5):
    if response.status_code != 200:
        print(f"Fehler: Statuscode {response.status_code}")
        return []

    soup = BeautifulSoup(response.text, "html.parser")
    headlines = soup.select("h2, h3")
    titles = [h.get_text(strip=True) for h in headlines if h.get_text(strip=True)]
    return titles[:limit]


def main():
    try:
        response = requests.get(
            SEARCH_URL, params={"searchString": SEARCH_STRING}, headers=HEADERS
        )
    except requests.exceptions.RequestException as e:
        print(f"Fehler: Verbindung fehlgeschlagen ({e})")
        return

    results = get_search_results(response)
    if not results:
        print("Keine Treffer gefunden.")
        return

    for i, title in enumerate(results, start=1):
        print(f"{i}. {title}")


if __name__ == "__main__":
    main()
