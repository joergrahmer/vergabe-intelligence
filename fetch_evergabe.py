import requests
from bs4 import BeautifulSoup

URL = "https://www.evergabe-online.de"


def get_page_title(response):
    if response.status_code != 200:
        print(f"Fehler: Statuscode {response.status_code}")
        return None

    soup = BeautifulSoup(response.text, "html.parser")
    title = soup.title.string if soup.title else None
    print(f"Titel der Seite: {title}")
    return title


def main():
    try:
        response = requests.get(URL)
    except requests.exceptions.RequestException as e:
        print(f"Fehler: Verbindung fehlgeschlagen ({e})")
        return

    get_page_title(response)


if __name__ == "__main__":
    main()
