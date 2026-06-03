"""
Vergabe Intelligence Agent — v2
Quellen:  TED Europa (Prototyp 1), bund.de + evergabe-online (Prototyp 2)
Datenbank: SQLite (data/vergabe.db) — akkumuliert sich über alle Läufe
Output:   Excel (data/vergabe_YYYYMMDD.xlsx) → GitHub Artifact

Lokaler Test:
    pip install requests openpyxl anthropic
    export ANTHROPIC_API_KEY="sk-..."
    python src/agent.py
"""

import os
import json
import re
import sqlite3
import time
from datetime import datetime
from pathlib import Path

import requests
import anthropic
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────────────────────
# KONFIGURATION
# ─────────────────────────────────────────────────────────

BASE_DIR   = Path(__file__).parent.parent
DATA_DIR   = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)

DB_PATH    = DATA_DIR / "vergabe.db"
EXCEL_PATH = DATA_DIR / f"vergabe_{datetime.now().strftime('%Y%m%d')}.xlsx"

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

# CPV-Codes: Verteidigung + Cyber/IT + Kommunikation
CPV_CODES = [
    "35000000",   # Sicherheits- & Verteidigungsausrüstung
    "35700000",   # Militärische Ausrüstung
    "72200000",   # Softwareprogrammierung & -beratung
    "72300000",   # Datendienstleistungen
    "72500000",   # Computerbezogene Dienstleistungen
    "72700000",   # Computernetzdienstleistungen
    "72800000",   # Computer-Audit & -Testdienste
    "64200000",   # Fernmeldedienstleistungen
]

COUNTRIES   = ["DE"]
YEAR_FROM   = 2022
MAX_RECORDS = 200     # Pro Lauf — erhöhen für Bulk-Historik

# Taxonomie (einmalig definiert, für AI-Prompt verwendet)
TAXONOMY = """
C1 – Cyber & IT-Sicherheit: SOC, SIEM, Penetration Testing, GRC, Kryptographie, Cyber-Ausbildung
C2 – Electronic Warfare: Jamming, SIGINT, Signalverarbeitung, EW-Software, EW-Integration
C3 – C2 / C4ISR & Führungssysteme: Lagebilder, Mission Planning, taktische Netze, Sensorfusion
C4 – Nachrichtenwesen & OSINT: OSINT-Plattformen, GEOINT, Satellitenbild-Analyse, AI-Analyse
C5 – Simulation & Ausbildung: Gefechtssimulatoren, Schießtrainer, E-Learning, Übungsunterstützung
C6 – IT-Infrastruktur & Betrieb: Rechenzentrum, Netzwerk, Workplace, Managed Services, Lizenzen
C7 – Plattformen & Hardware: Fahrzeuge, UAV, Marine, Munition, Waffensysteme
C8 – Sonstiges: Bau, Logistik, Beratung, Instandhaltung, nicht eindeutig zuordenbar
"""

KATEGORIE_COLORS = {
    "C1": "DBEAFE", "C2": "EDE9FE", "C3": "D1FAE5", "C4": "FEF3C7",
    "C5": "FCE7F3", "C6": "F3F4F6", "C7": "DCFCE7", "C8": "FEE2E2",
}

EXCEL_COLUMNS = [
    "id", "datum_veroeffentlichung", "quelle", "url", "vergabestelle",
    "titel_original", "cpv_code", "wert_eur", "land", "laufzeit_monate",
    "auftragnehmer_name", "auftragnehmer_name_norm", "auftragnehmer_land",
    "kategorie_code", "beschreibung_ai", "pe_relevant", "notizen",
]


# ─────────────────────────────────────────────────────────
# DATENBANK
# ─────────────────────────────────────────────────────────

def init_db(conn: sqlite3.Connection):
    conn.execute("""
        CREATE TABLE IF NOT EXISTS vergaben (
            id                      TEXT PRIMARY KEY,
            datum_veroeffentlichung TEXT,
            quelle                  TEXT,
            url                     TEXT,
            vergabestelle           TEXT,
            titel_original          TEXT,
            cpv_code                TEXT,
            wert_eur                REAL,
            land                    TEXT,
            laufzeit_monate         INTEGER,
            auftragnehmer_name      TEXT,
            auftragnehmer_name_norm TEXT,
            auftragnehmer_land      TEXT,
            kategorie_code          TEXT,
            beschreibung_ai         TEXT,
            pe_relevant             TEXT DEFAULT '',
            notizen                 TEXT DEFAULT '',
            created_at              TEXT DEFAULT (datetime('now'))
        )
    """)
    conn.commit()


def upsert_vergabe(conn: sqlite3.Connection, rec: dict) -> bool:
    """Fügt Vergabe ein. Gibt True zurück wenn neu, False wenn bereits vorhanden."""
    existing = conn.execute(
        "SELECT id FROM vergaben WHERE id = ?", (rec["id"],)
    ).fetchone()

    if existing:
        return False  # Duplikat — überspringen

    conn.execute("""
        INSERT INTO vergaben (
            id, datum_veroeffentlichung, quelle, url, vergabestelle,
            titel_original, cpv_code, wert_eur, land, laufzeit_monate,
            auftragnehmer_name, auftragnehmer_name_norm, auftragnehmer_land,
            kategorie_code, beschreibung_ai, pe_relevant, notizen
        ) VALUES (
            :id, :datum_veroeffentlichung, :quelle, :url, :vergabestelle,
            :titel_original, :cpv_code, :wert_eur, :land, :laufzeit_monate,
            :auftragnehmer_name, :auftragnehmer_name_norm, :auftragnehmer_land,
            :kategorie_code, :beschreibung_ai, :pe_relevant, :notizen
        )
    """, rec)
    conn.commit()
    return True


def load_all_vergaben(conn: sqlite3.Connection) -> list[dict]:
    cur = conn.execute(
        "SELECT * FROM vergaben ORDER BY datum_veroeffentlichung DESC"
    )
    cols = [d[0] for d in cur.description]
    return [dict(zip(cols, row)) for row in cur.fetchall()]


# ─────────────────────────────────────────────────────────
# QUELLE 1: TED EUROPA
# ─────────────────────────────────────────────────────────

def fetch_ted(max_records: int = MAX_RECORDS) -> list[dict]:
    """Holt Zuschlagsbekanntmachungen von TED Search API (anonym)."""

    cpv_part     = " OR ".join(f"cpv={c}" for c in CPV_CODES)
    country_part = " OR ".join(f"buyer-country={c}" for c in COUNTRIES)
    query        = (
        f"({cpv_part}) AND notice-type=can"
        f" AND ({country_part})"
        f" AND publication-date>={YEAR_FROM}0101"
    )

    fields = [
        "publication-number", "notice-title", "award-value-total",
        "organisation-name-winner", "winner-country", "buyer-name",
        "buyer-country", "cpv", "notice-date", "contract-duration",
        "notice-url", "description",
    ]

    results, page = [], 1
    while len(results) < max_records:
        try:
            resp = requests.post(
                "https://api.ted.europa.eu/v3/notices/search",
                json={"q": query, "page": page, "limit": 10, "fields": fields},
                timeout=30,
            )
            resp.raise_for_status()
        except requests.RequestException as e:
            print(f"  [TED] Fehler Seite {page}: {e}")
            break

        data    = resp.json()
        notices = data.get("notices", [])
        if not notices:
            break

        results.extend(notices)
        total = data.get("total", 0)
        print(f"  [TED] Seite {page} — {len(results)}/{min(total, max_records)}")

        if len(results) >= min(total, max_records):
            break
        page += 1
        time.sleep(0.5)

    return results[:max_records]


def normalize_ted(raw: dict) -> dict:
    def first(v):
        return (v[0] if isinstance(v, list) and v else v) or ""

    pub   = first(raw.get("publication-number", ""))
    title = first(raw.get("notice-title", ""))
    cpv   = first(raw.get("cpv", ""))

    # Wert
    try:
        wert = float(str(raw.get("award-value-total", "")).replace(",", ""))
    except (ValueError, TypeError):
        wert = None

    # Datum
    d = str(first(raw.get("notice-date", "")))
    try:
        datum = datetime.strptime(d[:8], "%Y%m%d").strftime("%Y-%m-%d")
    except ValueError:
        datum = d[:10]

    # Laufzeit
    dur      = raw.get("contract-duration", {})
    laufzeit = None
    if isinstance(dur, dict):
        laufzeit = dur.get("durationInMonths") or (
            round(dur["durationInDays"] / 30) if dur.get("durationInDays") else None
        )

    url = first(raw.get("notice-url", ""))
    if not url and pub:
        url = f"https://ted.europa.eu/udl?uri=TED:NOTICE:{pub}:TEXT:DE:HTML"

    return {
        "id":                       f"TED-{pub}",
        "datum_veroeffentlichung":  datum,
        "quelle":                   "TED",
        "url":                      url,
        "vergabestelle":            first(raw.get("buyer-name", "")),
        "titel_original":           title,
        "cpv_code":                 cpv,
        "wert_eur":                 wert,
        "land":                     first(raw.get("buyer-country", "DE")),
        "laufzeit_monate":          laufzeit,
        "auftragnehmer_name":       first(raw.get("organisation-name-winner", "")),
        "auftragnehmer_name_norm":  first(raw.get("organisation-name-winner", "")),
        "auftragnehmer_land":       first(raw.get("winner-country", "")),
        "kategorie_code":           "",
        "beschreibung_ai":          "",
        "pe_relevant":              "",
        "notizen":                  "",
        "_beschaffungstext":        first(raw.get("description", "")),
    }


# ─────────────────────────────────────────────────────────
# QUELLE 2: SERVICE.BUND.DE (RSS)
# ─────────────────────────────────────────────────────────

def fetch_bund() -> list[dict]:
    """
    Holt Vergaben von service.bund.de via RSS-Feed.
    Filter: Vergabestelle Bundeswehr / BMVg
    """
    # RSS-Feed für Bundeswehrverwaltung
    feeds = [
        "https://www.service.bund.de/Content/DE/Ausschreibungen/Suche/Formular.html"
        "?view=processForm&resultCount=50&searchType=1"
        "&tenderOrganisationHierarchySearch=Bundeswehrverwaltung"
        "&output=rss",
    ]

    results = []
    for feed_url in feeds:
        try:
            resp = requests.get(feed_url, timeout=20,
                                headers={"User-Agent": "VergabeIntelligence/1.0"})
            resp.raise_for_status()
            # Einfaches XML-Parsing ohne externe Bibliothek
            import xml.etree.ElementTree as ET
            root = ET.fromstring(resp.content)
            items = root.findall(".//item")
            print(f"  [bund.de] {len(items)} Einträge im Feed")
            for item in items:
                results.append({
                    "title":       (item.findtext("title") or "").strip(),
                    "link":        (item.findtext("link") or "").strip(),
                    "description": (item.findtext("description") or "").strip(),
                    "pubDate":     (item.findtext("pubDate") or "").strip(),
                    "category":    (item.findtext("category") or "").strip(),
                })
        except Exception as e:
            print(f"  [bund.de] Fehler: {e}")

    return results


def normalize_bund(raw: dict) -> dict:
    # Datum parsen (RSS-Format: "Mon, 15 Apr 2024 00:00:00 +0000")
    datum = ""
    try:
        from email.utils import parsedate_to_datetime
        datum = parsedate_to_datetime(raw["pubDate"]).strftime("%Y-%m-%d")
    except Exception:
        datum = raw.get("pubDate", "")[:10]

    # Eindeutige ID aus URL
    url  = raw.get("link", "")
    slug = re.sub(r"[^a-zA-Z0-9]", "-", url[-40:]) if url else raw["title"][:30]

    return {
        "id":                       f"BUND-{slug}",
        "datum_veroeffentlichung":  datum,
        "quelle":                   "bund.de",
        "url":                      url,
        "vergabestelle":            raw.get("category", "Bundeswehrverwaltung"),
        "titel_original":           raw.get("title", ""),
        "cpv_code":                 "",
        "wert_eur":                 None,
        "land":                     "DE",
        "laufzeit_monate":          None,
        "auftragnehmer_name":       "",
        "auftragnehmer_name_norm":  "",
        "auftragnehmer_land":       "",
        "kategorie_code":           "",
        "beschreibung_ai":          "",
        "pe_relevant":              "",
        "notizen":                  "",
        "_beschaffungstext":        raw.get("description", ""),
    }


# ─────────────────────────────────────────────────────────
# QUELLE 3: EVERGABE-ONLINE.DE (RSS)
# ─────────────────────────────────────────────────────────

def fetch_evergabe() -> list[dict]:
    """
    Holt Vergaben von evergabe-online.de.
    Hinweis: evergabe hat keine öffentliche API — wir nutzen den
    Suchergebnis-RSS für Bundeswehr-Vergaben.
    """
    # evergabe bietet RSS für gespeicherte Suchen
    feed_url = (
        "https://www.evergabe-online.de/tenderdetails.html"
        "?rss&filter.organisation=Bundeswehr&filter.type=award"
    )

    results = []
    try:
        resp = requests.get(feed_url, timeout=20,
                            headers={"User-Agent": "VergabeIntelligence/1.0"})
        resp.raise_for_status()
        import xml.etree.ElementTree as ET
        root  = ET.fromstring(resp.content)
        items = root.findall(".//item")
        print(f"  [evergabe] {len(items)} Einträge im Feed")
        for item in items:
            results.append({
                "title":       (item.findtext("title") or "").strip(),
                "link":        (item.findtext("link") or "").strip(),
                "description": (item.findtext("description") or "").strip(),
                "pubDate":     (item.findtext("pubDate") or "").strip(),
            })
    except Exception as e:
        print(f"  [evergabe] Fehler: {e}")
        print(f"  [evergabe] Hinweis: evergabe-online erfordert ggf. Account für RSS-Zugang")

    return results


def normalize_evergabe(raw: dict) -> dict:
    datum = ""
    try:
        from email.utils import parsedate_to_datetime
        datum = parsedate_to_datetime(raw["pubDate"]).strftime("%Y-%m-%d")
    except Exception:
        datum = raw.get("pubDate", "")[:10]

    url  = raw.get("link", "")
    slug = re.sub(r"[^a-zA-Z0-9]", "-", url[-40:]) if url else raw["title"][:30]

    return {
        "id":                       f"EVG-{slug}",
        "datum_veroeffentlichung":  datum,
        "quelle":                   "evergabe",
        "url":                      url,
        "vergabestelle":            "",
        "titel_original":           raw.get("title", ""),
        "cpv_code":                 "",
        "wert_eur":                 None,
        "land":                     "DE",
        "laufzeit_monate":          None,
        "auftragnehmer_name":       "",
        "auftragnehmer_name_norm":  "",
        "auftragnehmer_land":       "",
        "kategorie_code":           "",
        "beschreibung_ai":          "",
        "pe_relevant":              "",
        "notizen":                  "",
        "_beschaffungstext":        raw.get("description", ""),
    }


# ─────────────────────────────────────────────────────────
# AI ENRICHMENT
# ─────────────────────────────────────────────────────────

def enrich(records: list[dict]) -> list[dict]:
    """Klassifiziert und beschreibt Vergaben via Claude Haiku (schnell + günstig)."""

    if not ANTHROPIC_API_KEY:
        print("  [AI] Kein API-Key — Enrichment übersprungen (alle → C8)")
        for r in records:
            r["kategorie_code"]  = "C8"
            r["beschreibung_ai"] = "(kein API-Key)"
        return records

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    total  = len(records)

    for i, rec in enumerate(records, 1):
        text = f"{rec.get('titel_original','')} {rec.get('_beschaffungstext','')}"[:800]

        prompt = f"""Klassifiziere diese Vergabe. Antworte NUR mit JSON, kein Text davor/danach.

Vergabe: {text}
CPV: {rec.get('cpv_code','')}
Vergabestelle: {rec.get('vergabestelle','')}

Taxonomie:
{TAXONOMY}

{{"kategorie": "C1", "beschreibung": "1-2 Sätze deutsch: was konkret beschafft?"}}

Regeln: Genau einen Code C1-C8. Bei Unklarheit C8."""

        try:
            msg = client.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=150,
                messages=[{"role": "user", "content": prompt}],
            )
            raw = msg.content[0].text.strip()
            m   = re.search(r'\{.*\}', raw, re.DOTALL)
            if m:
                obj = json.loads(m.group())
                rec["kategorie_code"]  = obj.get("kategorie", "C8")
                rec["beschreibung_ai"] = obj.get("beschreibung", "")
            else:
                rec["kategorie_code"]  = "C8"
                rec["beschreibung_ai"] = "(parse error)"
        except Exception as e:
            rec["kategorie_code"]  = "C8"
            rec["beschreibung_ai"] = f"(fehler: {str(e)[:40]})"

        if i % 5 == 0:
            print(f"  [AI] {i}/{total} enriched ...")
        time.sleep(0.2)

    return records


# ─────────────────────────────────────────────────────────
# EXCEL EXPORT
# ─────────────────────────────────────────────────────────

def write_excel(all_records: list[dict], new_ids: set[str], filepath: Path):
    wb = Workbook()

    # ── Sheet 1: Alle Vergaben ──────────────────────────────
    ws = wb.active
    ws.title = "Vergaben"

    header_fill  = PatternFill("solid", start_color="1E3A5F")
    header_font  = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    thin_border  = Border(
        bottom=Side(style="thin", color="DDDDDD"),
        right=Side(style="thin",  color="DDDDDD"),
    )

    DISPLAY_COLS = [
        ("ID",                      "id"),
        ("Datum",                   "datum_veroeffentlichung"),
        ("Quelle",                  "quelle"),
        ("URL",                     "url"),
        ("Vergabestelle",           "vergabestelle"),
        ("Titel",                   "titel_original"),
        ("CPV",                     "cpv_code"),
        ("Wert EUR",                "wert_eur"),
        ("Land",                    "land"),
        ("Laufzeit Mon.",           "laufzeit_monate"),
        ("Auftragnehmer (Original)","auftragnehmer_name"),
        ("Auftragnehmer (Norm.)",   "auftragnehmer_name_norm"),
        ("AN-Land",                 "auftragnehmer_land"),
        ("Kategorie",               "kategorie_code"),
        ("Beschreibung (AI)",       "beschreibung_ai"),
        ("PE Relevant",             "pe_relevant"),
        ("Notizen",                 "notizen"),
    ]

    col_widths = [22,12,8,12,32,52,10,14,6,10,32,32,8,10,55,10,30]

    for ci, (label, _) in enumerate(DISPLAY_COLS, 1):
        cell = ws.cell(1, ci, label)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = thin_border
        ws.column_dimensions[get_column_letter(ci)].width = col_widths[ci - 1]

    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"

    new_fill  = PatternFill("solid", start_color="FFFBEB")  # Goldgelb = neu diese Woche

    for ri, rec in enumerate(all_records, 2):
        kat      = rec.get("kategorie_code", "C8")
        bg_hex   = KATEGORIE_COLORS.get(kat, "FFFFFF")
        is_new   = rec.get("id", "") in new_ids
        row_fill = PatternFill("solid", start_color="FFFBEB") if is_new else \
                   PatternFill("solid", start_color=bg_hex)

        for ci, (_, field) in enumerate(DISPLAY_COLS, 1):
            val  = rec.get(field, "")
            cell = ws.cell(ri, ci, val)
            cell.font      = Font(name="Arial", size=9)
            cell.fill      = row_fill
            cell.alignment = Alignment(vertical="top", wrap_text=False)
            cell.border    = thin_border

            if field == "url" and val:
                cell.hyperlink = val
                cell.font = Font(name="Arial", size=9, color="0563C1", underline="single")
            if field == "wert_eur" and val:
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal="right", vertical="top")

    ws.auto_filter.ref = f"A1:{get_column_letter(len(DISPLAY_COLS))}1"

    # ── Sheet 2: Diese Woche (nur neue) ─────────────────────
    ws2 = wb.create_sheet("Neu diese Woche")
    for ci, (label, _) in enumerate(DISPLAY_COLS, 1):
        cell = ws2.cell(1, ci, label)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = thin_border
        ws2.column_dimensions[get_column_letter(ci)].width = col_widths[ci - 1]

    new_records = [r for r in all_records if r.get("id", "") in new_ids]
    for ri, rec in enumerate(new_records, 2):
        kat     = rec.get("kategorie_code", "C8")
        bg_hex  = KATEGORIE_COLORS.get(kat, "FFFFFF")
        for ci, (_, field) in enumerate(DISPLAY_COLS, 1):
            val  = rec.get(field, "")
            cell = ws2.cell(ri, ci, val)
            cell.font      = Font(name="Arial", size=9)
            cell.fill      = PatternFill("solid", start_color=bg_hex)
            cell.alignment = Alignment(vertical="top")
            cell.border    = thin_border
            if field == "url" and val:
                cell.hyperlink = val
                cell.font = Font(name="Arial", size=9, color="0563C1", underline="single")
            if field == "wert_eur" and val:
                cell.number_format = '#,##0'

    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = f"A1:{get_column_letter(len(DISPLAY_COLS))}1"

    # ── Sheet 3: Statistik ───────────────────────────────────
    ws3 = wb.create_sheet("Statistik")
    ws3.column_dimensions["A"].width = 28
    ws3.column_dimensions["B"].width = 16
    ws3.column_dimensions["C"].width = 20
    ws3.column_dimensions["D"].width = 14

    stat_header = Font(bold=True, name="Arial", size=10)
    ws3.cell(1, 1, "Kategorie").font = stat_header
    ws3.cell(1, 2, "Anzahl Vergaben").font = stat_header
    ws3.cell(1, 3, "Volumen EUR").font = stat_header
    ws3.cell(1, 4, "Davon neu").font = stat_header

    from collections import Counter
    kat_count  = Counter(r.get("kategorie_code","C8") for r in all_records)
    kat_volume = {}
    kat_new    = Counter(r.get("kategorie_code","C8") for r in all_records if r.get("id") in new_ids)
    for r in all_records:
        k = r.get("kategorie_code", "C8")
        kat_volume[k] = kat_volume.get(k, 0) + (r.get("wert_eur") or 0)

    TAX_NAMES = {
        "C1":"Cyber & IT-Sicherheit", "C2":"Electronic Warfare",
        "C3":"C2 / C4ISR",           "C4":"Nachrichtenwesen & OSINT",
        "C5":"Simulation & Ausbildung","C6":"IT-Infrastruktur & Betrieb",
        "C7":"Plattformen & Hardware","C8":"Sonstiges",
    }
    for ri, code in enumerate(["C1","C2","C3","C4","C5","C6","C7","C8"], 2):
        bg = KATEGORIE_COLORS.get(code, "FFFFFF")
        fill = PatternFill("solid", start_color=bg)
        cells = [
            ws3.cell(ri, 1, f"{code} – {TAX_NAMES[code]}"),
            ws3.cell(ri, 2, kat_count.get(code, 0)),
            ws3.cell(ri, 3, kat_volume.get(code, 0)),
            ws3.cell(ri, 4, kat_new.get(code, 0)),
        ]
        for c in cells:
            c.fill = fill
            c.font = Font(name="Arial", size=9)
        cells[2].number_format = '#,##0'

    # Summenzeile
    ri = 10
    ws3.cell(ri, 1, "GESAMT").font = Font(bold=True, name="Arial", size=9)
    ws3.cell(ri, 2, f"=SUM(B2:B9)").font = Font(bold=True, name="Arial", size=9)
    ws3.cell(ri, 3, f"=SUM(C2:C9)").font = Font(bold=True, name="Arial", size=9)
    ws3.cell(ri, 3).number_format = '#,##0'
    ws3.cell(ri, 4, f"=SUM(D2:D9)").font = Font(bold=True, name="Arial", size=9)

    # ── Sheet 4: Taxonomie ───────────────────────────────────
    ws4 = wb.create_sheet("Taxonomie")
    ws4.column_dimensions["A"].width = 8
    ws4.column_dimensions["B"].width = 30
    ws4.column_dimensions["C"].width = 65
    ws4.column_dimensions["D"].width = 10

    for ci, h in enumerate(["Code","Kategorie","Typische Inhalte","PE-Fokus"], 1):
        ws4.cell(1, ci, h).font = Font(bold=True, name="Arial", size=10)

    TAX_ROWS = [
        ("C1","Cyber & IT-Sicherheit",       "SOC, SIEM, Pen-Test, GRC, Kryptographie, Cyber-Ausbildung",           "hoch"),
        ("C2","Electronic Warfare",           "Jamming, SIGINT, Signalverarbeitung, EW-Software, EW-Integration",    "hoch"),
        ("C3","C2 / C4ISR",                  "Lagebilder, Mission Planning, taktische Netze, Sensorfusion",          "hoch"),
        ("C4","Nachrichtenwesen & OSINT",     "OSINT-Plattformen, GEOINT, Satellitenbild-Analyse, AI-Analyse",        "hoch"),
        ("C5","Simulation & Ausbildung",      "Gefechtssimulatoren, Schießtrainer, E-Learning",                       "mittel"),
        ("C6","IT-Infrastruktur & Betrieb",  "Rechenzentrum, Netzwerk, Workplace, Managed Services",                 "mittel"),
        ("C7","Plattformen & Hardware",       "Fahrzeuge, UAV, Marine, Munition, Waffensysteme",                      "gering"),
        ("C8","Sonstiges",                    "Bau, Logistik, Beratung — nicht eindeutig zuordenbar",                 "gering"),
    ]
    for ri, (code, name, desc, pe) in enumerate(TAX_ROWS, 2):
        fill = PatternFill("solid", start_color=KATEGORIE_COLORS.get(code, "FFFFFF"))
        for ci, v in enumerate([code, name, desc, pe], 1):
            c = ws4.cell(ri, ci, v)
            c.fill = fill
            c.font = Font(name="Arial", size=9)

    wb.save(filepath)
    print(f"\n[Excel] Gespeichert: {filepath}")
    print(f"[Excel] {len(all_records)} Vergaben gesamt · {len(new_ids)} neu · 4 Sheets")


# ─────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  Vergabe Intelligence Agent v2")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M UTC')}")
    print("=" * 60)

    # DB initialisieren
    conn = sqlite3.connect(DB_PATH)
    init_db(conn)

    existing_count = conn.execute("SELECT COUNT(*) FROM vergaben").fetchone()[0]
    print(f"\n[DB] Bestehende Einträge: {existing_count}")

    new_ids = set()

    # ── Quelle 1: TED ───────────────────────────────────────
    print("\n[1/3] TED Europa ...")
    try:
        ted_raw     = fetch_ted(max_records=MAX_RECORDS)
        ted_records = [normalize_ted(r) for r in ted_raw]
        ted_records = enrich(ted_records)
        for rec in ted_records:
            if upsert_vergabe(conn, {k: v for k, v in rec.items() if not k.startswith("_")}):
                new_ids.add(rec["id"])
        print(f"  [TED] {len(ted_raw)} geholt · {len([r for r in ted_records if r['id'] in new_ids])} neu")
    except Exception as e:
        print(f"  [TED] Fehler: {e}")

    # ── Quelle 2: bund.de ───────────────────────────────────
    print("\n[2/3] service.bund.de ...")
    try:
        bund_raw     = fetch_bund()
        bund_records = [normalize_bund(r) for r in bund_raw]
        bund_records = enrich(bund_records)
        for rec in bund_records:
            if upsert_vergabe(conn, {k: v for k, v in rec.items() if not k.startswith("_")}):
                new_ids.add(rec["id"])
        print(f"  [bund.de] {len(bund_raw)} geholt · {len([r for r in bund_records if r['id'] in new_ids])} neu")
    except Exception as e:
        print(f"  [bund.de] Fehler: {e}")

    # ── Quelle 3: evergabe ──────────────────────────────────
    print("\n[3/3] evergabe-online.de ...")
    try:
        evg_raw     = fetch_evergabe()
        evg_records = [normalize_evergabe(r) for r in evg_raw]
        evg_records = enrich(evg_records)
        for rec in evg_records:
            if upsert_vergabe(conn, {k: v for k, v in rec.items() if not k.startswith("_")}):
                new_ids.add(rec["id"])
        print(f"  [evergabe] {len(evg_raw)} geholt · {len([r for r in evg_records if r['id'] in new_ids])} neu")
    except Exception as e:
        print(f"  [evergabe] Fehler: {e}")

    # ── Excel Export ────────────────────────────────────────
    print("\n[Export] Excel schreiben ...")
    all_records = load_all_vergaben(conn)
    write_excel(all_records, new_ids, EXCEL_PATH)
    conn.close()

    # ── Zusammenfassung ─────────────────────────────────────
    from collections import Counter
    kat_count = Counter(r.get("kategorie_code","C8") for r in all_records)
    total_eur = sum(r.get("wert_eur") or 0 for r in all_records)

    print("\n" + "=" * 60)
    print("  ZUSAMMENFASSUNG")
    print("=" * 60)
    print(f"  Vergaben gesamt:  {len(all_records)}")
    print(f"  Neu diese Woche:  {len(new_ids)}")
    print(f"  Gesamtvolumen:    EUR {total_eur:,.0f}")
    print(f"  Nach Kategorie:")
    for code in ["C1","C2","C3","C4","C5","C6","C7","C8"]:
        n = kat_count.get(code, 0)
        if n:
            print(f"    {code}: {n}")
    print("=" * 60)


if __name__ == "__main__":
    main()
