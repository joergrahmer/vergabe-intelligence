import requests
from bs4 import BeautifulSoup

URL = "https://www.evergabe-online.de"


def main():
    response = requests.get(URL)

    if response.status_code != 200:
        print(f"Fehler: Statuscode {response.status_code}")
        return

    soup = BeautifulSoup(response.text, "html.parser")
    title = soup.title.string if soup.title else None
    print(f"Titel der Seite: {title}")


if __name__ == "__main__":
    main()
