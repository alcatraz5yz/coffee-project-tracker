/**
 * tabelle30-node.js — Pure-Node replacements for the Python helpers used by the
 * Tabelle 30 comparison, so the Windows build needs neither Python nor LibreOffice.
 *
 *   parseTabelle30(wordPath)        ← parse_tabelle30.py   (uses docx-reader.js)
 *   parseExcelErgaenzung(xlsxPath)  ← parse_excel_ergaenzung.py (uses SheetJS)
 *
 * Both return the same JSON shape the Python scripts produced, so the existing
 * compare logic and frontend renderer keep working unchanged.
 */

"use strict";

const path = require("path");
const XLSX = require("xlsx");
const { readTables } = require("./docx-reader");
const { assertExcelReadable } = require("./excel-safety");
const { convertDocToDocx } = require("./doc-convert");

const TABLE_MARKER = "Resistance to heat and fire";

// ── parse_tabelle30.py port ──────────────────────────────────────────────────
function parseTabelle30(wordPath) {
  let convertedFromDoc = false;
  if (/\.doc$/i.test(wordPath)) {
    // Legacy binary .doc — convert to .docx via MS Word (Windows only).
    wordPath = convertDocToDocx(wordPath);
    convertedFromDoc = true;
  } else if (!/\.(docx|docm)$/i.test(wordPath)) {
    throw new Error(`Nicht unterstütztes Word-Format für Tabelle 30: ${path.basename(wordPath)}`);
  }
  const tables = readTables(wordPath);
  let tableIdx = -1;
  let target = null;
  for (let i = 0; i < tables.length; i++) {
    const flat = tables[i].map((r) => r.map((c) => c.text).join(" ")).join(" ");
    if (flat.includes(TABLE_MARKER)) { tableIdx = i; target = tables[i]; break; }
  }
  if (!target) return { tableIdx: null, rowCount: 0, rows: [] };

  // Find the column-label header row ("Object/part No." …); data rows follow it and
  // have a non-empty first cell (the numeric sub-header rows have an empty first cell).
  let labelRow = -1;
  for (let i = 0; i < target.length && i < 6; i++) {
    const c0 = (target[i][0]?.text || "").toLowerCase();
    if (c0.includes("object") && c0.includes("part")) { labelRow = i; break; }
  }
  const start = labelRow >= 0 ? labelRow + 1 : 2;
  const headerRows = target.slice(0, start).map((r) => r.map((c) => c.text));

  const rows = [];
  for (let ri = start; ri < target.length; ri++) {
    const cells = target[ri].map((c) => c.text);
    const cellRuns = target[ri].map((c) => c.runs);
    if (!(cells[0] || "").trim()) continue;          // skip sub-header / blank rows
    if (!cells.some((t) => (t || "").trim())) continue;
    rows.push({ rowIdx: rows.length + 1, tableRowIdx: ri, cells, cellRuns });
  }
  return { sourceFile: wordPath, convertedFromDoc, tableIdx, rowCount: rows.length, headerRows, rows };
}

// ── parse_excel_ergaenzung.py port ───────────────────────────────────────────
const HEADER_KEYWORDS = [
  "drw no", "drawing no", "part name",
  "component name", "ef part number", "specification",
];
const RELEVANT_CATEGORIES = new Set(["plastic", "rubber", "lacquer", "sub assembly", "others"]);

function clean(v) {
  if (v === null || v === undefined) return "";
  return String(v).replace(/ /g, " ").trim();
}

function findHeaderRow(rowsArr) {
  for (let ri = 0; ri < Math.min(12, rowsArr.length); ri++) {
    const joined = (rowsArr[ri] || []).map((c) => (c ? String(c).toLowerCase() : "")).join(" | ");
    if (HEADER_KEYWORDS.some((kw) => joined.includes(kw))) return ri;
  }
  return null;
}

function getColIndices(header) {
  const idx = {};
  for (let ci = 0; ci < header.length; ci++) {
    const cell = header[ci];
    if (cell === null || cell === undefined || cell === "") continue;
    const label = String(cell).toLowerCase().trim().replace(/\n/g, " ");
    if (["drw no", "drawing no", "art no"].includes(label) || label.startsWith("drawing index")) idx.drwNo = ci;
    else if (label === "comments" || label.startsWith("comments")) idx.comments = ci;
    else if (label.includes("new part") && label.includes("existing part")) idx.newExisting = ci;
    else if (label === "level") idx.level = ci;
    else if (label.includes("ef part number") || label.includes("ef art no") || label === "ef partnumber") idx.efArtNo = ci;
    else if (label.includes("explo position") || label === "position" || label.startsWith("pos")) idx.position = ci;
    else if (label.includes("part name") || label.includes("component name")) idx.partName = ci;
    else if (label.includes("raw material supplier")) idx.materialSupplier = ci;
    else if (label.includes("supplier") && !label.includes("material")) idx.supplier = ci;
    else if (label === "category" || label.startsWith("category")) idx.category = ci;
    else if (label === "material type" || label.startsWith("material type")) idx.materialType = ci;
    else if (label === "trade name" || label.startsWith("trade name")) idx.tradeName = ci;
    else if (label === "color" || label.startsWith("color ")) idx.color = ci;
    else if (label === "material no." || label.startsWith("material no")) idx.materialNo = ci;
    else if (label.includes("ul file")) idx.ulFile = ci;
    else if (label.includes("iec") && label.includes("certificate")) idx.iecCert = ci;
    else if (label.includes("specification")) idx.specification = ci;
    else if (label.includes("today used") || label === "today used") idx.todayUsed = ci;
    else if (label.includes("change to") || label.includes("plan to be used")) idx.changeTo = ci;
    else if (label.includes("active") && (label.includes("alt") || label.includes("alternative"))) idx.activeAlt = ci;
    else if (label === "type") idx.type = ci;
  }
  return idx;
}

function parseExcelErgaenzung(xlsxPath) {
  assertExcelReadable(xlsxPath);
  const wb = XLSX.readFile(xlsxPath, { cellStyles: false });
  let sheetName = null; let rowsArr = null; let headerRi = null;
  for (const name of wb.SheetNames) {
    const arr = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: null, raw: true, blankrows: true });
    const ri = findHeaderRow(arr);
    if (ri !== null) { sheetName = name; rowsArr = arr; headerRi = ri; break; }
  }
  if (headerRi === null) {
    const err = new Error("no sheet with recognised headers (Drw no / Part name / Component name)");
    err.code = "NO_HEADER";
    throw err;
  }

  const header = rowsArr[headerRi];
  const cols = getColIndices(header);
  const isChangeList = "changeTo" in cols;
  const fmt = isChangeList ? "change-list" : "part-list";

  const get = (row, key) => (key in cols && cols[key] < (row ? row.length : 0)) ? clean(row[cols[key]]) : "";

  let hasTab30Markers = false;
  if (!isChangeList) {
    for (let ri = headerRi + 1; ri < rowsArr.length; ri++) {
      const r = rowsArr[ri];
      if (r && r.length > 0 && clean(r[0]).toLowerCase().replace(/ /g, "") === "tab30") { hasTab30Markers = true; break; }
    }
  }
  let hasChangeMarkers = false;
  if (!isChangeList && "comments" in cols) {
    for (let ri = headerRi + 1; ri < rowsArr.length; ri++) {
      const cm = get(rowsArr[ri], "comments").toLowerCase();
      if (cm.includes("version") && cm.includes("add the part number")) { hasChangeMarkers = true; break; }
    }
  }

  const outRows = [];
  const seen = new Set();
  for (let ri = headerRi + 1; ri < rowsArr.length; ri++) {
    const row = rowsArr[ri];
    const partName = get(row, "partName");
    if (!partName) continue;

    if (isChangeList) {
      outRows.push({
        rowIdx: outRows.length + 1, excelRow: ri + 1,
        drwNo: get(row, "drwNo"), position: get(row, "position"), type: get(row, "type"),
        partName, supplier: get(row, "supplier"), todayUsed: get(row, "todayUsed"),
        changeTo: get(row, "changeTo"), reference: "", category: "",
      });
    } else {
      const marker = get(row, "comments") || (row && row.length > 0 ? clean(row[0]) : "");
      const newExisting = get(row, "newExisting");
      const level = get(row, "level");

      if (hasTab30Markers && marker.toLowerCase().replace(/ /g, "") !== "tab30") continue;
      if (!hasTab30Markers && hasChangeMarkers) {
        const ml = marker.toLowerCase();
        if (!ml.includes("add the part number") || newExisting.toLowerCase() !== "new") continue;
      }

      const category = get(row, "category").toLowerCase();
      if (category && !RELEVANT_CATEGORIES.has(category)) continue;

      const matType = get(row, "materialType");
      const trade = get(row, "tradeName");
      const color = get(row, "color");
      let material = [matType, trade, color].filter(Boolean).join(" ");
      if (!material) material = get(row, "specification");
      if (!material) continue;

      const dedupeKey = [
        get(row, "position"),
        partName.toLowerCase().replace(/\n/g, " ").trim(),
        material.toLowerCase().replace(/\n/g, " ").trim(),
      ].join("\u0000");
      if (hasChangeMarkers && seen.has(dedupeKey)) continue;
      seen.add(dedupeKey);

      const supplier = get(row, "materialSupplier") || get(row, "supplier");
      const reference = [get(row, "ulFile"), get(row, "iecCert")].filter(Boolean).join(" / ");

      outRows.push({
        rowIdx: outRows.length + 1, excelRow: ri + 1,
        drwNo: get(row, "efArtNo") || get(row, "drwNo"), position: get(row, "position"),
        type: get(row, "activeAlt"), partName, supplier, todayUsed: "",
        changeTo: material, reference, category, marker, newExisting, level,
      });
    }
  }

  return {
    sourceFile: xlsxPath, sheetName, headerRow: headerRi + 1, format: fmt,
    hasTab30Markers, hasChangeMarkers, rowCount: outRows.length, rows: outRows,
  };
}

module.exports = { parseTabelle30, parseExcelErgaenzung };
