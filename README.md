# Vergabe Intelligence

Automatisiertes Monitoring von Verteidigungsvergaben (Bundeswehr / BMVg).  
Läuft wöchentlich via GitHub Actions. Output: Excel als Artifact zum Download.

## Quellen
| Quelle | Inhalt | Zugang |
|---|---|---|
| TED Europa | EU-weite Zuschlagsbekanntmachungen | Kostenlos, anonym |
| service.bund.de | Nationale Vergaben Bundeswehrverwaltung | Kostenlos, RSS |
| evergabe-online.de | BAAINBw / BMVg Vergaben | RSS (ggf. Account nötig) |

## Setup (einmalig, ~15 Minuten)

### 1. Repository klonen
```bash
git clone https://github.com/DEIN-USERNAME/vergabe-intelligence.git
cd vergabe-intelligence
```

### 2. Anthropic API Key als GitHub Secret hinterlegen
1. GitHub → Repository → Settings → Secrets and variables → Actions
2. "New repository secret"
3. Name: `ANTHROPIC_API_KEY`
4. Value: dein Anthropic API Key (`sk-ant-...`)

### 3. Ersten Lauf manuell starten
1. GitHub → Actions → "Vergabe Intelligence — Weekly Run"
2. "Run workflow" → Branch: main → "Run workflow"
3. Nach ~5 Minuten: unter "Artifacts" das Excel herunterladen

### 4. Automatischer Wochenlauf
Läuft jeden Montag 06:00 Uhr (DE Sommerzeit) automatisch.

## Lokaler Test
```bash
pip install requests openpyxl anthropic
export ANTHROPIC_API_KEY="sk-ant-..."
python src/agent.py
# → data/vergabe_YYYYMMDD.xlsx
```

## Standalone-Skript: evergabe-online.de Suche
`fetch_evergabe.py` durchsucht evergabe-online.de nach einem festen Suchbegriff
(`SEARCH_STRING` im Skript, aktuell `"BAAINBw"`) und gibt die ersten 5 Trefferüberschriften
auf der Konsole aus.

```bash
pip install -r requirements.txt
python3 fetch_evergabe.py
```

Beispielausgabe:
```
1. Messumgebung für Immissionsmessungen - X/U2CE/VA021/TA212
2. Beschaffung eines Kaltwasser-Erzeugers - X/U2CE/VA133/VC171
3. 5,56mm x 45 2DK-1LS-1HK, 200er Gurt
4. 6003058358-BAAINBw E2.1
5. Herstellung und Lieferung von LS10 Leuchtköper, Fallschirm, Handabfeuerung,( MOD9163100) 38mm, PT, Einzelstern, Rot
```

Tests (offline, ohne echten HTTP-Call):
```bash
pip install -r requirements-dev.txt
pytest tests/ -v
```

## Taxonomie
| Code | Kategorie | PE-Fokus |
|---|---|---|
| C1 | Cyber & IT-Sicherheit | hoch |
| C2 | Electronic Warfare | hoch |
| C3 | C2 / C4ISR & Führungssysteme | hoch |
| C4 | Nachrichtenwesen & OSINT | hoch |
| C5 | Simulation & Ausbildung | mittel |
| C6 | IT-Infrastruktur & Betrieb | mittel |
| C7 | Plattformen & Hardware | gering |
| C8 | Sonstiges | gering |

## Kosten (Schätzung pro Monat)
- GitHub Actions: **kostenlos** (Free Tier: 2000 Min/Monat)
- Anthropic API: **~$0.50–2.00** (abhängig von Datenmenge, Haiku-Modell)
- Gesamt: **< $2/Monat**

## Datenbank
`data/vergabe.db` — SQLite, liegt im Repo, wird bei jedem Lauf erweitert.  
Duplikate werden automatisch übersprungen (PRIMARY KEY = Bekanntmachungs-ID).

## Excel-Output (4 Sheets)
1. **Vergaben** — alle Einträge, farbcodiert nach Kategorie, neue Einträge goldgelb markiert
2. **Neu diese Woche** — nur neue Einträge dieses Laufs
3. **Statistik** — Anzahl + Volumen je Kategorie, automatisch berechnet
4. **Taxonomie** — Referenz-Sheet

## Erweiterung auf Phase 2 (Marktanalyse)
Folgende Felder können später manuell ergänzt werden:
- `pe_relevant` → Ja/Nein
- `auftragnehmer_name_norm` → normalisierter Firmenname für Pivot-Tabellen
- `notizen` → freies Kommentarfeld
