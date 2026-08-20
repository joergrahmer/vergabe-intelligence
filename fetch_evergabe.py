from urllib.parse import urljoin

from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

AWARDED_URL = "https://www.evergabe-online.de/search/awardedProcedure.html"
SEARCH_STRING = "Software"
SEARCH_INPUT_SELECTOR = "#keywordString"
SUBMIT_BUTTON_SELECTOR = '[name="submitButton"]'
RESULT_LINK_SELECTOR = 'a.text-wrap[href*="contractAward.html"]'


def extract_results(html, base_url, limit=5):
    soup = BeautifulSoup(html, "html.parser")
    links = soup.select(RESULT_LINK_SELECTOR)

    results = []
    for a in links:
        title = a.get_text(strip=True)
        href = a.get("href")
        if not title or not href:
            continue
        results.append({"title": title, "url": urljoin(base_url, href)})

    return results[:limit]


def fetch_search_results_html(search_string=SEARCH_STRING):
    with sync_playwright() as p:
        browser = p.chromium.launch()
        try:
            page = browser.new_page()
            page.goto(AWARDED_URL)
            page.fill(SEARCH_INPUT_SELECTOR, search_string)
            with page.expect_response(
                lambda r: "submitButton" in r.url and r.request.method == "POST"
            ):
                page.click(SUBMIT_BUTTON_SELECTOR)
            page.wait_for_load_state("networkidle")
            return page.content(), page.url
        finally:
            browser.close()


def main():
    try:
        html, url = fetch_search_results_html()
    except Exception as e:
        print(f"Fehler: Suche fehlgeschlagen ({e})")
        return

    results = extract_results(html, url)
    if not results:
        print("Keine Treffer gefunden.")
        return

    for result in results:
        print(f"Titel: {result['title']}")
        print(f"URL: {result['url']}")
        print("---")


if __name__ == "__main__":
    main()
