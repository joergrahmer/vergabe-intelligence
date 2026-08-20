from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup

SEARCH_URL = "https://www.evergabe-online.de/search/awardedProcedure.html?8"
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
    links = soup.select('a[data-evid="search_list_result"]')

    results = []
    for a in links:
        title = a.get_text(strip=True)
        href = a.get("href")
        if not title or not href:
            continue
        results.append({"title": title, "url": urljoin(response.url, href)})

    return results[:limit]


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

    for result in results:
        print(f"Titel: {result['title']}")
        print(f"URL: {result['url']}")
        print("---")


if __name__ == "__main__":
    main()
