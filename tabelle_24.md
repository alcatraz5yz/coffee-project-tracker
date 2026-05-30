# Tabelle 24 — Component Certificate Workflow

Knowledge dump about the Tabelle 24 compliance task at Eugster/Frismag and the
PCS-Dashboard scanner/editor we've built for it.

## What "Tabelle 24" Is

Tabelle 24 is the **component-certificate table** inside the *internal Bauteilliste*
Word file of every EF project (EF1157, EF1175, …). It lists every component used in
the product alongside the certificate(s) that prove its compliance. The chapter is
typically titled "24.1 TABLE: components EF<NNNN>".

The task: per EF project, go through every row, verify each certificate is still
current, download the newest PDF from the relevant authority website, place it in
the right component folder, and update the Word file's color coding.

This is recurring compliance work — one pass per product, per regulatory cycle.

## Folder Structure (per EF project)

```
<EF-project-root>/
└── Zulassungen/
    └── IEC/
        ├── 01 Administration
        ├── 02 Änderungen
        ├── 03 User Manual
        ├── 04 Typenschild
        ├── 05 Pflichtenheft Offerte Auftrag
        ├── 06 Materialliste Explo
        ├── 07 Elektronik
        ├── 08 Schema Gerät
        ├── 09 Bautelliste/                ← (note: typo "Bautelliste", not "Bauteilliste")
        │   ├── EF<NNNN> Bautl _Intern_ <date>.docx   ← latest internal version (target file)
        │   ├── EF<NNNN> Bautl _Intern_ <older>.doc   ← older internal versions
        │   ├── EF<NNNN> Bautl_VDE_ <date>.doc        ← customer/VDE version — NEVER touch
        │   ├── <variant subfolders>/                  ← e.g. "EF1175 200V JP/"
        │   └── Archiv/                                ← older versions kept here
        ├── 10 Komponenten/                ← component PDFs live here
        │   ├── <Komponente A>/
        │   │   ├── <newest cert PDF>
        │   │   └── Archiv/                ← replaced PDFs moved here
        │   └── <Komponente B>/
        ├── 11 Materialprüfungen
        ├── 12 Untersuchungen
        ├── 13 Fotos
        ├── 14 PAK Bewertung
        ├── 15 CB Prüfberichte Safety EMC ErP
        └── 16 Länderumsetzungen
```

**File-naming convention:**
- `*_Intern_*` = internal working version (has dates, gets edited)
- `*_VDE_*` = customer/external version (no dates, never modify)
- Newest by date wins; older intern files stay in `09 Bautelliste/` or `Archiv/`
- Variants (regional voltage versions, e.g. Japan 200V) get their own subfolder
  inside `09 Bautelliste/`, named like `EF1175 200V JP/`

## Two Possible File Formats — Same Logical Content

### 1. Paragraph layout (modern `.docx`)

Some newer Intern files (e.g. `EF1157 Bautl _Intern_ 20.04.2026.docx`) have **no
Word tables at all**. The table content is laid out as a sequence of paragraphs:
one paragraph per logical cell, with line breaks inside paragraphs for multi-line
cells. Column structure is implied by paragraph order.

Markers we use:
- Chapter heading paragraph: contains `TABLE: components`
- Column header paragraphs follow (until "mark(s) of conformity")
- Then data: every 6 paragraphs ≈ one row, but multi-line cells split into more
- Rows always END with the colored mark-of-conformity paragraph

### 2. Word-table layout (`.doc` → LibreOffice → `.docx`)

Older Intern files in `.doc` format use a real Word table. Converting via
LibreOffice headless (`soffice --headless --convert-to docx`) yields a `.docx` with
a real `<w:tbl>` element — typically 10 columns wide because merged cells get
expanded to duplicated columns (e.g. cols 0=1, 5=6, 7=8).

**Important:** LibreOffice conversion **loses run-level color information**. So
parsing a converted `.doc` returns paragraphs/cells with default colors, and
status (green/red/yellow) becomes `unknown`. The user has to manually mark via UI
buttons in that case.

## Column Layout (6 logical columns)

| # | Header                | Content                                                       |
|---|-----------------------|---------------------------------------------------------------|
| 1 | object/part No.       | Component name + 7-digit M3 part number(s)                    |
| 2 | manufacturer/trademark| Maker name                                                    |
| 3 | type/model            | Model designation                                             |
| 4 | technical data        | Voltage / current / dimensions                                |
| 5 | standard              | IEC/EN/SN/JIS reference                                       |
| 6 | mark(s) of conformity | Cert authority + number + date(s) — **this cell is colored**  |

The left-side numbers (column 1) are **M3 part numbers** — the user looks each up
in the Infor M3 ERP to get a validity code. The right-side cell (column 6) is the
**mark of conformity** with cert authority (VDE/ESTI/ENEC/UL/JET/…) and number;
the user verifies these via each authority's website and downloads PDFs.

## Status Colors (the heart of the system)

Used in the mark-of-conformity column. The COLORED part is always the date.

| Color hex | Color    | Meaning                                                   |
|-----------|----------|-----------------------------------------------------------|
| `002060`  | dark blue| Default text (cert label, body content)                   |
| `00B050`  | green    | Confirmed current — verified, no action needed            |
| `FF0000`  | red      | Expired / withdrawn — cert no longer valid                |
| `FFFF00`  | yellow hl| Outdated, awaiting update (highlight + strikethrough)     |

**Row-level color** (the whole row): the user's process colors entire rows
- No color → still to be processed
- Green row → already confirmed, skip
- Blue row → newly added last cycle, re-verify
- Red row → invalid, skip unless new valid cert found

**Date encoding rules:**
- **Confirm** (date still valid): date GREEN
- **Expire** (no replacement): date RED
- **Update** (newer PDF exists): old date YELLOW-highlighted + strikethrough,
  newline, new date GREEN below

`word_updater.py` references both — `00B050` for green, `FF0000` for red, yellow
as `<w:highlight w:val="yellow">` not as font color.

## Validity Rule (critical)

> A certificate is only outdated if a **newer PDF exists with a different date**.
> A past date in the table alone does NOT mean invalid.
> No PDF found, or PDF date matches table → GREEN (still valid).

This avoids accidentally marking valid certs as expired just because some time
passed. The authoritative date is always what the PDF itself says — the table is
a cache.

## The 3-Step Workflow (per row)

### Step 1 — M3 lookup (left-side numbers)
If column 1 has a 7-digit number (e.g. `0156693`, `0112824`):
1. Open Infor M3, enter the number
2. Response code `20` = valid; `90` = invalid
3. Mark the number green or red in the Bauteilliste
4. From M3, open the linked browser page → download the right PDFs → place in
   the component folder

### Step 2 — Folder structure (`10 Komponenten/`)
1. Find / create the correct subfolder for each component
2. If the folder structure exists from a sibling/previous project, copy it
3. Each component subfolder may need an `Archiv/` for older PDFs

### Step 3 — Certificate placement
1. For each cert authority on the right (VDE/ESTI/ENEC/UL/JET):
   - Look up the cert number on the authority's website
   - Compare its date with the date in the Word file
   - If newer: download PDF, move old PDF into `Archiv/`, place new PDF in main folder
   - Update the Word file's date coloring accordingly

The Word file edits happen LAST, after all PDFs are placed.

## Certificate-Source Websites

| Authority | URL                                                                                | Notes                                                  |
|-----------|------------------------------------------------------------------------------------|--------------------------------------------------------|
| VDE       | https://www.vde.com/tic-en/marks-and-certificates/vde-approved-products/search     | Cookie banner in shadow DOM — see VDE section below    |
| ESTI      | https://www.electrosuisse.ch                                                       | Swiss certs                                            |
| ENEC      | URL printed on the certificate itself                                              | European cooperation scheme                            |
| UL        | https://productiq.ulprospector.com                                                 | US / international                                     |
| JET       | https://www.jet.or.jp                                                              | Japan — used in 200V JP variants                       |
| Infor M3  | Internal ERP                                                                       | Only reachable from office network (CH16186 work laptop)|
| IDM       | intranet.emea.efgroup.int                                                          | Office-network only; skip when at home                 |

## VDE Specifics (worked out by experimentation)

**Direct cert detail URL:**
```
https://www.vde.com/tic-en/marks-and-certificates/vde-approved-products/certificate?id=<CERT_NO>&type=zertreg%7Ccertificate
```

**Cookie banner — Usercentrics in shadow DOM.** Plain text selectors like
`button:has-text("Accept all Cookies")` are unreliable. Must pierce the shadow root:

```javascript
document.querySelector('#usercentrics-root').shadowRoot
  .querySelector('[data-testid="uc-accept-all-button"]')
  .click()
```

**PDF links** live on `zertreg-proxy.vde.com/documents/download?uuid=…` once cookies
are accepted. The cert page renders a "Documents" section with one or two cards:

- *Appendix 100 — Type Code* (general doc)
- *Appendix 200 — Technical Characteristics* (preferred for date verification)

**Date extraction from PDF:** the PDF contains a line like
`(letzte Änderung / updated 2025-10-31 )` — usually on page 1, sometimes in an
appendix-style table cell on page 2-3. The `extract_updated_date()` helper tries
four patterns:
1. `updated YYYY-MM-DD` directly after the label
2. `updated DD.MM.YYYY` directly after the label
3. Same label, but value in a separate cell up to 300 chars later
4. Single ISO date in the whole document — take it

**Reality of cert availability:** old short (6-digit) certs from before ~2020
(e.g. `120557` from 2014, `109835` from 2017) usually have **no public PDFs** on
vde.com. Newer 8-digit `40xxxxxx` certs almost always have 1-2 PDFs. Withdrawn
certs (e.g. EF1157's `4000125`) don't return any results at all.

## Logical-Row Grouping (continuation rows)

Some logical entries span multiple **table rows** in Word — for example, a Cord
set JP entry has the cord set primary row + a continuation row "With power cord"
that's still part of the same assembly. In Word, the row separator is invisible
(no top border on the continuation row), so it looks like one tall row.

**Detection:** check whether the continuation row's first cell has a top border:

```python
def row_is_continuation(row):
    cell = row.cells[0]
    tcPr = cell._tc.find(qn("w:tcPr"))
    if tcPr is None: return False
    tcBorders = tcPr.find(qn("w:tcBorders"))
    if tcBorders is None: return False
    top = tcBorders.find(qn("w:top"))
    if top is None: return True   # tcBorders exists but no top → no separator
    val = top.get(qn("w:val"))
    return val in ("nil", "none")
```

The parser emits `groupId` + `groupPosition: "primary"|"continuation"` per row.
The dashboard UI renders continuations with a `↳ ` marker and no top border so
the visual grouping matches Word.

This signal only exists in **table-layout files**. For paragraph-layout `.docx`,
there's no equivalent — those rows can't be automatically grouped.

## Example: EF1175 Japan Variant Cord Set

A real logical row from `EF1175 200V JP/EF1175 Bauteilliste 200V JP Intern PCS8-E
Rev. 14.08.2025.doc` showing the grouping complexity:

```
table_row[3]:
  col0: Cord set JP / 0112826 / 0080424 plug / 0043372 appl. outlet
  last: JET0670-43001-1013 / 11.11.2031 / JET0670-43004-1001 / 19.01.2026
  top border: single sz=6 → visible separator → primary row

table_row[4]:
  col0: With power cord / 0112824
  last: JET0670-12009-1004 / 07.11.2029
  top border: none → no separator → continuation of row 3

table_row[5]:
  col0: Appliance inlet / 0108881
  last: VDE 40030341
  top border: single sz=4 → visible → primary row (new logical entry)
```

So logically:
- **One entry** = Cord set JP with 4 part numbers (0112826, 0080424, 0043372, 0112824)
  and 3 JET certificates
- The user looks up each of the 4 part numbers separately in M3
- The user checks each of the 3 JET certs separately

## What We Built — PCS Dashboard Integration

### Files

| Path                                                           | Purpose                                                       |
|----------------------------------------------------------------|---------------------------------------------------------------|
| `~/Desktop/scripts-tabelle24/parse_bauteilliste.py`            | Read .docx / .doc → JSON of Tabelle-24 rows                   |
| `~/Desktop/scripts-tabelle24/update_bauteilliste.py`           | Apply staged status changes → write new .docx                 |
| `~/Desktop/scripts-tabelle24/vde_lookup.py`                    | Look up a single VDE cert, download PDF, extract date         |
| `~/Desktop/scripts-tabelle24/vde_checker_final.py`             | Pre-existing batch checker (EF1157-hardcoded, reference)      |
| `~/Desktop/scripts-tabelle24/enec_esti_check.py`               | Pre-existing ENEC/ESTI batch checker (not yet integrated)     |
| `~/Desktop/scripts-tabelle24/word_updater.py`                  | Pre-existing Word colorizer (EF1157-hardcoded, reference)     |
| `~/Desktop/coffee-project-tracker/server.js`                   | Express server with Tabelle-24 endpoints                      |
| `~/Desktop/coffee-project-tracker/tabelle24.html`              | Standalone UI page                                            |

Convention: scripts live in `~/Desktop/scripts-tabelle24/` (macOS dev) or
`C:\Users\CH16186\Desktop\scripts\` (Windows work laptop). **Never inside the
project folder.**

### REST endpoints

| Method | Path                              | Purpose                                                       |
|--------|-----------------------------------|---------------------------------------------------------------|
| GET    | `/api/tabelle24/files?root=…`     | List *Intern* Bauteilliste files under a project              |
| POST   | `/api/tabelle24/parse`            | Parse one file → rows JSON                                    |
| POST   | `/api/tabelle24/save`             | Apply staged changes → write new .docx                        |
| POST   | `/api/tabelle24/vde-lookup`       | Look up one VDE cert → PDF list + date + verdict              |
| GET    | `/api/tabelle24/vde-pdf?path=…`   | Serve a downloaded PDF from the lookup's temp dir             |

### Parser output schema

```jsonc
{
  "sourceFile": "/path/to/source.docx",
  "convertedFromDoc": false,        // true if .doc → LibreOffice → .docx happened
  "format": "paragraph" | "table",
  "rowCount": 18,
  "rows": [
    {
      "rowIdx": 1,                  // 1-based, sequential
      "groupId": 1,                 // table format only
      "groupPosition": "primary" | "continuation",
      // Addressing (one of two shapes):
      "startParaIdx": 10,           // paragraph format only
      "endParaIdx": 15,             // paragraph format only — also the mark cell paragraph
      "tableIdx": 0,                // table format only
      "tableRowIdx": 2,             // table format only
      "markCellIdx": 8,             // table format only — index in row.cells
      // Content:
      "cells": ["Cord set CH …", "Sun Fai …", "SF-285", "10A 250 V~", "IEC 60884-1 …"],
      "markOfConformity": "ESTI 24.0540\n14.08.2027",
      "markOfConformityRuns": [
        { "text": "ESTI 24.0540", "color": null,     "highlight": null,     "strike": false },
        { "text": "\n",           "color": null,     "highlight": null,     "strike": false },
        { "text": "14.08.2027",   "color": "00B050", "highlight": null,     "strike": false }
      ],
      "status": "green" | "red" | "yellow" | "unknown",
      "statusHex": "00B050"         // or null
    }
  ]
}
```

The `markOfConformityRuns` array lets the UI render the cell exactly like Word —
colored dates, yellow highlights, strikethroughs — by emitting styled `<span>`s.

### Edit actions (UI → server)

Each edit is one entry in the `changes[]` array passed to `/save`. Addressing
must match the format used for the source:

**Paragraph-format change:**
```json
{ "paraIdx": 43, "action": "expire" }
```

**Table-format change:**
```json
{ "tableIdx": 0, "tableRowIdx": 6, "markCellIdx": 8, "action": "update", "newDate": "2025-10-31" }
```

Available actions:
- `confirm` — color the last date in the cell green
- `expire` — color the last date in the cell red
- `update` — yellow-highlight + strikethrough the last date, then add `<w:br>` +
  new date in green

The updater always targets the **last date** in the cell. Multi-date cells (e.g.
existing update state with old + new date) lose the old-date strikethrough when
re-edited — this is a known limitation.

### VDE lookup flow

1. `POST /api/tabelle24/vde-lookup` with `{ certNumber, currentDate? }`
2. Node spawns `vde_lookup.py`, pipes JSON via stdin
3. Python:
   - Playwright headless Chromium
   - Goto search URL, click `uc-accept-all-button` via shadow DOM
   - Goto `certificate?id=<cert>` URL
   - Check `"Certificate number:"` in body → `found`
   - Collect all `zertreg-proxy.vde.com` links → `pdfList`
   - Prefer Appendix 200, fall back to 100, then last
   - Download chosen PDF via `requests` with playwright cookies
   - Run `extract_updated_date()` to get `onlineDate`
   - Compare with `currentDate` → `current` | `newer` | `older` | `unknown`
4. Return JSON:

```jsonc
{
  "certNumber": "40042427",
  "found": true,
  "pdfList": [{ "label": "Appendix 100 : Type Code", "url": "…" }, …],
  "chosen": { "label": "Appendix 200 …", "url": "…" },
  "downloadedPdf": "/private/tmp/vde-lookups/vde_40042427_appendix_200.pdf",
  "onlineDate": "2025-10-31",
  "comparison": "newer"
}
```

5. UI renders a drawer below the row with the verdict, PDF list, a link to open
   the downloaded PDF, and a "→ apply as <action>" suggestion button that
   pre-stages the appropriate edit (confirm if current; update with the new date
   if newer; expire if not found).

Performance: ~9 seconds per cert (Playwright + cookie + nav + download + PDF parse).
Acceptable for per-row clicks; could batch or persist browser session if needed.

## Known Limitations

1. **`.doc` color information is lost on LibreOffice conversion.** Status comes
   back as `unknown` for all rows; user has to mark manually. Output of the save
   endpoint is always `.docx`, not `.doc`.

2. **Continuation grouping only works on table-format files.** Paragraph-layout
   `.docx` files don't have an analogous structural signal we can detect.

3. **Multi-date cells lose old strikethrough on re-edit.** The updater rebuilds
   the paragraph from scratch when modifying the last date, dropping run-level
   formatting on prefix text. Workaround: edit-once.

4. **VDE lookup is per-cert, not batch.** ~9 seconds each. For a 30-row file, a
   "Check all VDE" action would take ~5 minutes.

5. **Old VDE certs often have no public PDFs.** 6-digit certs from before ~2020
   typically return `found: true, pdfList: []` because vde.com doesn't expose
   their appendix downloads anymore. Not a script bug — actual website state.

6. **Withdrawn certs return `found: false`.** That's the signal to suggest
   `expire` — the cert no longer exists in the VDE database.

7. **No automation for ESTI / ENEC / UL / JET yet.** Phase 2b would integrate
   the existing `enec_esti_check.py` and add ENEC + ESTI buttons; UL and JET
   would need new scrapers.

8. **No M3 automation.** Infor M3 is intranet-only and the [work-laptop is on
   the security watchlist](#) — automation needs more thought before being
   built.

9. **No folder-structure setup yet** (Phase 3). The "create / copy `10 Komponenten/`
   subfolders" workflow step isn't automated yet.

10. **No PDF placement / Archiv handling yet** (Phase 4). Downloads land in
    `/tmp/vde-lookups/`, not in the project's component folders. The user still
    has to copy them by hand.

## Open Workflow Questions

- For VDE "Keine PDFs" (cert exists but no public PDF): is the right default
  to leave it as-is (assume valid), or to suggest expire (can't verify)?
- How to surface logical-row grouping in paragraph-layout `.docx` (no top-border
  signal available)?
- Should M3 automation use the work-laptop's logged-in browser session, or skip
  entirely and ask the user to enter results manually?

## Reference: Pre-existing scripts (EF1157-specific, useful as templates)

- `vde_checker_final.py` — batch flow that worked April 2026; has hardcoded
  `CERTS[]` and `CERT_FOLDERS{}` maps for EF1157. Good model for "navigate
  → download → date extract → status compare → save to correct folder".
- `enec_esti_check.py` — analogous batch flow for ENEC / ESTI / UL.
- `word_updater.py` — pre-existing colorizer with hardcoded `CHANGES{}` mapping
  date strings to (action, new_date, comment) tuples. Same color/highlight/strike
  logic the new `update_bauteilliste.py` uses, but indexed by date text rather
  than paragraph/cell coordinates.

The new scripts replace the hardcoded lists with parameters and project-folder
addressing, so they generalize across EF projects.
