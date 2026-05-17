import openpyxl, json, sys, os

EXCEL_PATH = os.environ.get("ARCHIVE_EXCEL", os.path.join(os.path.expanduser("~"), "Desktop", "PCS_Archiv_Muster.xlsx"))

try:
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active
except FileNotFoundError:
    print(json.dumps({"error": f"Datei nicht gefunden: {EXCEL_PATH}"}))
    sys.exit(1)

# Group locations by EF number (unique, preserve order)
by_ef = {}
for row in ws.iter_rows(min_row=2, values_only=True):
    if not row or not row[0]:
        continue
    ef_nr   = str(row[0]).strip()
    location = str(row[4]).strip() if row[4] else ""
    if not ef_nr or not location or location == "—":
        continue
    if ef_nr not in by_ef:
        by_ef[ef_nr] = []
    if location not in by_ef[ef_nr]:
        by_ef[ef_nr].append(location)

result = [{"project_id": ef, "location": " · ".join(locs)} for ef, locs in by_ef.items()]
print(json.dumps(result))
