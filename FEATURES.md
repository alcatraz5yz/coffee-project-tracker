# PCS Kaffee Dashboard — Feature-Übersicht

## Was ist das?

Ein internes Web-Dashboard zur Verwaltung von PCS-Projekten (Approbation, Zulassung, Builds, Nachweise) für Kaffeemaschinen-Projekte. Läuft lokal oder im Firmennetzwerk, alle 6 Standorte können gleichzeitig darauf zugreifen.

---

## Technischer Stack

| Komponente | Technologie |
|---|---|
| Frontend | HTML / CSS / JavaScript (kein Framework) |
| Backend | Node.js + Express |
| Datenbank | SQLite (via better-sqlite3) |
| Deployment | Lokal oder interner Server |

---

## Seitenleiste (Sidebar)

- **Projektliste** mit Suchfeld (sucht nach ID, Name, Phase, Markt)
- **Gesundheitsstatus** pro Projekt: Gut / Beobachten / Blockiert
- **Build-Stufen-Pills** pro Projekt (PT1, OOT, TS1, TS2, ...) direkt unter dem Projektnamen
  - Werden automatisch aus den Ziffer-Testprogrammen abgeleitet
  - Die aktive Build-Stufe bleibt blau hervorgehoben
  - Klick auf eine Stufe öffnet sofort den Approbation-Tab gefiltert auf diese Stufe
- **P:\PCS Scanner** Button unten in der Sidebar mit letztem Scan-Zeitstempel
- **Scan-Button** startet manuellen Ordner-Scan, aktualisiert alle Daten live

---

## Tab 1 — PCS Status (Übersicht)

### Header
- Projekttitel und Familie oben links
- **Produktbild** in der Mitte des Headers — wird automatisch geladen aus `evidence-{n}/product.jpg`
- Tab-Navigation oben rechts

### Kennzahlen-Strip
- Abschlussstatus (Bei VDE / In Arbeit / Abgeschlossen / ...)
- Aktueller Build + Musterzahl
- Anzahl offene Build-Gates
- Blockierte PCS Punkte
- Anzahl indexierter Word-Berichte

### Abschluss-Gates
Zeigt den technischen Abschlussfortschritt automatisch:
- VDE Unterlagen eingereicht
- Prüfungen bestanden
- Finaler VDE/CB Bericht vorhanden
- EMC / ErP Reports vorhanden
- Keine offenen VDE Rückfragen
- Finale Dokumente abgelegt

### Panels
- **Zulassungsstatus**: Checkliste aller Zertifizierungsschritte
- **Builds & Muster**: Timeline aller Builds mit Status, Musterzahl, Notiz
- **PCS Blocker**: Risikoliste mit Hoch/Mittel Einstufung
- **Nächste PCS Massnahmen**: Die 5 nächsten offenen Tasks

---

## Tab 2 — Approbation (Ziffer-Checkliste)

- Vollständige **IEC-Ziffer-Checkliste** pro Projekt (18 Ziffern für EF1157)
- Kompaktes **3-Spalten-Layout**: Ziffer-Nr. / Thema / Status
- Notiz erscheint als kleine graue Unterzeile unter dem Titel
- Direkte Ziffer-Nachweise erscheinen als kompakte Links unter dem Titel
- Status per Klick änderbar: **Offen → Erledigt → Nicht nötig → Blockiert**
- Status wird in der **Datenbank gespeichert** (sofort für alle Standorte sichtbar)
- Filterbar nach Status (Alle / Erledigt / Offen / Nicht nötig / Blockiert)
- **Build-Filterung über Sidebar**: Klick auf PT1/OOT/TS1/TS2 in der Seitenleiste filtert die Ziffer-Liste auf die Prüfungen dieser Testserie
  - Aktiver Build erscheint im Tabellen-Titel (z.B. "Ziffer Checkliste — OOT")
  - Ziffern mit `Alle`-Zuweisung erscheinen immer
- Jede Ziffer hat eine **Build-Zuweisung** (welche Testserien diese Ziffer umfasst)

### Build-Zuweisung EF1157 (IEC 60335-2-14)
| Ziffer | Bereich | Teststufen |
|---|---|---|
| 4 | General conditions | Alle |
| 7 | Marking | PT1, TS1 |
| 8 | Live parts | PT1, TS1 |
| 10 | Input / current | PT1, TS1, TS2 |
| 11 | Heating | PT1, TS1 |
| 13 | Leakage current | PT1, TS1 |
| 15 | Moisture resistance | TS1 |
| 16 | Leakage after moisture | TS1 |
| 17 | Overload protection | PT1, TS1 |
| 19 | Abnormal operation | PT1, OOT, TS1 |
| 20 | Stability / mech. | PT1 |
| 22 | Construction | PT1, TS1 |
| 24 | Components | Alle |
| 25 | Supply connection | PT1, TS1 |
| 27 | Earthing | PT1, TS1 |
| 29 | Clearances / creepage | PT1, TS1 |
| 30 | Resistance to heat/fire | PT1, TS1 |
| 32 | Radiation / toxicity | PT1, TS1 |

---

## Tab 3 — Massnahmen (Task-Board)

- Vollständige Task-Liste pro Projekt: Bereich, Aufgabe, Verantwortlicher, Deadline, Status
- Filtert nach Bereich: Zulassung / Elektronik / Build / PCS / Approbation
- **Status per Klick änderbar**: Offen → In Arbeit → Erledigt → Blockiert
- Erledigte Tasks werden durchgestrichen und ausgegraut
- Status wird sofort in der Datenbank gespeichert
- Sortiert nach Deadline

---

## Tab 4 — Nachweise (Evidence Index)

### PCS Nachweisindex (Karten)
- 16 IEC-Ordner als Karten, **4-Spalten-Grid**
- Jede Karte zeigt:
  - Bereichsname (fett)
  - Status-Badge (Vorhanden / Blockiert / Offen)
  - **Grosse Dateianzahl** als visuelle Hauptinfo
  - "Ordner öffnen →" Link
  - **Farbiger linker Rand** je nach Status: grün = Vorhanden, rot = Blockiert, amber = Offen
  - Hover: leichter Schatten und mini Anheben

### PCS Approbationsberichte (Tabelle)
- Alle Word-Prüfberichte aus "12 Untersuchungen", automatisch vom Scanner indexiert
- Spalten: Projekt / Build / Version / Geändert + Grösse / Status / Dateiname
- Status: Aktuell / Archiviert / Referenz
- **Word-Dateien** (`.doc`, `.docx`, `.docm`) öffnen direkt in Microsoft Word via `ms-word:` Protokoll
- **PDF-Dateien** öffnen im Browser

---

## Tab 5 — Fachliche Freigabe

Manueller Bestätigungsbereich — nur durch Rücksprache mit VDE, Approbation oder Projektleitung auszufüllen.

### Bestätigungs-Gates
**Status pro Gate:** Offen → Ja → Nein → Teilweise (per Klick wechseln)

**Gesamtstatus** wird automatisch berechnet:
- Alle Ja → Abgeschlossen
- Mindestens ein Nein → Blockiert
- Alle Offen → Nicht gestartet
- Gemischt → In Prüfung

### Rücksprache & Dokumentation
- Bestätigt von (Name / Funktion)
- Datum
- Notiz / Rücksprache
- Automatisch gespeichert nach 600ms

---

## P:\PCS Ordner-Scanner

Scannt den `P:\PCS` Ordner (oder auf Mac `~/Desktop` für Tests) und indexiert automatisch alle Projekte.

### Was der Scanner macht
- Erkennt Projektordner (beginnend mit Zahl)
- Findet den IEC-Unterordner (`Zulassungen/IEC/` oder ähnlich)
- Zählt Dateien pro IEC-Unterordner (alle Dateitypen)
- Indexiert Word-Dokumente aus "12 Untersuchungen" als Berichtsversionen
- Leitet Build-Stufe aus Ordnernamen ab (Prototype→PT1, TS1, OOT, 127V→Brazil, etc.)
- Erkennt Archiv-Dokumente (`/Archiv/`, `kopie`, `Vorlage`)

### Was der Scanner NICHT ändert
- Ziffer-Status (manuell gesetzt)
- Fachliche Freigabe (manuell gesetzt)
- Tasks, Risiken, Zertifizierung (manuell gepflegt)

### Konfiguration
```bash
# Standard: ~/Desktop (Mac Entwicklung)
node scanner.js

# Windows Produktion
PCS_ROOT="P:\PCS" node scanner.js
```

### Dokument-Verlinkung
- Lokal auf Mac: `evidence-{n}/` Symlink → wird automatisch über Express serviert
- Windows Produktion: `/files/{n}/` → Express serviert `P:\PCS` direkt
- Symlinks ausserhalb des Projektordners werden korrekt aufgelöst (realpathSync)

---

## Datenbank (SQLite)

Alle Daten in `pcs.db`. Keine Cloud, keine externen Abhängigkeiten.

| Tabelle | Inhalt |
|---|---|
| `projects` | Stammdaten aller Projekte |
| `builds` | Build-Timeline pro Projekt |
| `tasks` | Task-Board Einträge |
| `risks` | PCS Blocker / Risiken |
| `certification` | Zertifizierungs-Checkliste |
| `ziffern` | IEC-Ziffer-Checkliste inkl. `builds`-Spalte |
| `evidence_links` | Direkte Dateilinks pro Ziffer |
| `subtopic_summaries` | Beschreibungen je Checkliste |
| `closeout_gates` | Automatische Abschluss-Gates |
| `fachfreigabe_gates` | Manuelle Freigabe-Gates |
| `fachfreigabe_meta` | Bestätigt-von, Datum, Notiz |
| `document_groups` | PCS Nachweisindex (Scanner) |
| `report_versions` | Word-Berichtsversionen (Scanner) |
| `project_stats` | Kennzahlen-Karten |
| `scan_log` | Scan-Historie |

---

## Dokumentenzugriff

- **Lokal (Mac)**: `evidence-{n}/` Symlink → `ms-word:ofe|u|http://localhost:8090/...`
- **Windows Intranet**: `/files/{n}/` → `ms-word:ofe|u|http://<server>:8090/...`
- **SharePoint**: Konfigurierbar in `config.js`
- `.doc`, `.docx`, `.docm` → Microsoft Word
- `.pdf` → Browser (inline)
- Konfiguration in `config.js` (`documentMode: "local"` | `"sharepoint"`)

---

## Produktbild

- Automatisch geladen aus `evidence-{n}/product.jpg` (oder `.png`, `thumbnail.jpg`, `thumbnail.png`)
- Erscheint mittig in der Topbar jedes Projekts
- Wird ausgeblendet wenn kein Bild vorhanden
- Für EF1157: `/Users/allan/Desktop/1157/product.jpg` (aus Ordner 13 Fotos)

---

## Mehrsprachigkeit

- Oberfläche vollständig auf **Deutsch**
- Alle Statuswerte auf Deutsch
- Originale Dateinamen bleiben unverändert

---

## Responsive Design

- Optimiert für Desktop (1200px+)
- Tablet (1100px)
- Grundfunktionen Mobile (760px)

---

## Starten

```bash
# Einmalig: Datenbank befüllen
node seed.js

# Server starten
npm start
# → http://localhost:8090

# Mit explizitem PCS-Pfad (Windows)
PCS_ROOT="P:\PCS" npm start
```

---

## Geplante Features

- **Mehrbenutzer-Konflikte**: Optimistic Locking bei gleichzeitigen Änderungen
- **Firmen-Login**: Microsoft Entra ID / Office 365
- **Auto-Scan**: Zeitgesteuerter Scan (z.B. stündlich) ohne manuellen Button-Klick
- **Export**: Projektbericht als PDF oder Excel
- **Bild-Upload**: Produktbild direkt im Dashboard hochladen
