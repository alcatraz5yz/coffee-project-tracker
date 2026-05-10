# EF1157 PCS Folder Audit

Audit date: 2026-05-09

Source folder: `/Users/allan/Desktop/1157`

## Scope

The scan covered the full EF1157 folder under `Zulassungen/IEC`.

Total files found: 503

Main file types:

- 276 PDF
- 129 JPG
- 30 DOCX
- 18 XLSX
- 9 DOC
- 8 XLS
- 4 XLSM
- 3 DOCM

Ignored as real evidence:

- `.DS_Store`
- `Thumbs.db`
- `~$...` Microsoft Office lock files
- `*.tmp`

## Important Document Groups

The PCS dashboard now includes a Critical Document Index for:

- Administration: project plan, dispatch lists, minutes, delivery notes.
- Standards / Changes: IEC/EN 60335 documents and change documents.
- Manual / Labels: user manual draft and EU/CH/DE/Brazil type labels.
- Order / Scope: VDE order, offers, approbation order form, forms.
- BOM / Exploded Views: approval BOM and exploded assemblies.
- Electronics: mainboard/MMI schematics, PCB/artwork, BOMs, supplier/component evidence.
- Device Schemas: wiring and fluid diagrams.
- Bautelliste: internal and VDE part-list versions.
- Components: component approval evidence.
- Material / GWT: Table 30, material test drafts, UL yellow cards, GWT evidence.
- Investigations: Approbation Word reports, Ziffer 11/19 measurements, EF1234 127V PT1 report.
- Photos: TS1 and measurement photo evidence, Ziffer 19 and Ziffer 30 photos.
- PAK / LFGB: PAK BOM/list and PAK exploded assembly.
- CB / Safety / EMC / ErP: final report area.
- Country Implementation: country rollout area.

## Key Findings

The folder is strong in raw evidence but weak in final-package clarity.

Strong areas:

- Electronics evidence is extensive.
- Component certificates are extensive.
- Measurement evidence for Ziffer 11 and Ziffer 19 exists.
- Ziffer 30 photo and material evidence exists.
- TS1/TS2 dispatch lists exist.
- EF1234 Brazil 127V PT1 report exists inside `12 Untersuchungen`.

Weak or blocked areas:

- `15 CB Prüfberichte Safety EMC ErP` does not contain a clear final CB/Safety/EMC/ErP report package yet.
- `16 Länderumsetzungen` is empty, which matters for Brazil/EF1234.
- `14 PAK Bewertung` has high-level PAK files, but LFGB and individual test subfolders appeared empty in this scan.
- Some folders contain old Office lock files and system files that should not be treated as real versions.

## Dashboard Changes From This Audit

- Added a PCS Evidence Index to the Evidence tab.
- Added missing project tasks for final CB/EMC/ErP clarification, TS1/TS2 dispatch overview, and country implementation scope.
- Kept Approbation Word reports as their own table.
- Kept long file paths hidden in the UI for readability.

## Recommended Next Company Step

For a real company pilot, do not import all 503 files as flat documents.

Use the current structure:

- Critical groups for management overview.
- Approbation Word report table for versions.
- Ziffer checklist rows for detailed compliance evidence.
- SharePoint links later for secure document opening and audit history.
