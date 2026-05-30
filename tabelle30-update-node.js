"use strict";

const fs = require("fs");
const path = require("path");
const { readZipEntries, writeZipEntries } = require("./docx-reader");

const TABLE_MARKER = "TABLE: Resistance to heat and fire";
const FOOTNOTE_RE = /^\s*\d+\)\s|supplementary information/i;

function xmlEscape(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function textParagraph(text) {
  const lines = String(text || "").split(/\r?\n/);
  const body = lines.map((line, i) => `${i ? "<w:br/>" : ""}<w:t xml:space="preserve">${xmlEscape(line)}</w:t>`).join("");
  return `<w:p><w:r>${body}</w:r></w:p>`;
}

function decodeXml(s) {
  return String(s || "")
    .replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, d) => String.fromCharCode(Number(d)))
    .replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCharCode(parseInt(h, 16)));
}

function splitBlocks(xml, openRe, closeTag) {
  const out = [];
  let i = 0;
  while (true) {
    openRe.lastIndex = i;
    const mm = openRe.exec(xml);
    if (!mm) break;
    const start = mm.index;
    const end = xml.indexOf(closeTag, start);
    if (end < 0) break;
    out.push({ start, end: end + closeTag.length, xml: xml.slice(start, end + closeTag.length) });
    i = end + closeTag.length;
  }
  return out;
}

function cellText(tcXml) {
  const parts = [];
  for (const m of tcXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)) parts.push(decodeXml(m[1]));
  return parts.join(" ").replace(/\s+/g, " ").trim();
}

function rowCells(trXml) {
  return splitBlocks(trXml, /<w:tc(?:\s[^>]*)?>/g, "</w:tc>");
}

function replaceCellText(tcXml, text) {
  const tcPr = tcXml.match(/<w:tcPr[\s\S]*?<\/w:tcPr>/)?.[0] || "";
  return tcXml.replace(/(<w:tc(?:\s[^>]*)?>)[\s\S]*?(<\/w:tc>)/, `$1${tcPr}${textParagraph(text)}$2`);
}

function replaceCellAt(rowXml, cellBlock, text) {
  const nextCell = replaceCellText(cellBlock.xml, text);
  return rowXml.slice(0, cellBlock.start) + nextCell + rowXml.slice(cellBlock.end);
}

function populateRow(templateRowXml, entry) {
  let row = templateRowXml;
  let cells = rowCells(row);
  if (cells.length < 4) throw new Error(`template row has only ${cells.length} cells; need at least 4`);

  const replacements = new Map();
  replacements.set(0, entry.partName || "");
  replacements.set(1, entry.supplier || "");
  replacements.set(2, entry.material || "");
  for (let i = 3; i < cells.length - 1; i++) replacements.set(i, "--");
  replacements.set(cells.length - 1, entry.reference || "");

  for (const [i, value] of [...replacements.entries()].sort((a, b) => b[0] - a[0])) {
    cells = rowCells(row);
    row = replaceCellAt(row, cells[i], value);
  }
  return row;
}

function findTabelle30Table(documentXml) {
  const tables = splitBlocks(documentXml, /<w:tbl(?:\s[^>]*)?>/g, "</w:tbl>");
  return tables.find((t) => t.xml.includes(TABLE_MARKER));
}

function findLastDataRow(rows) {
  for (let i = rows.length - 1; i >= 0; i--) {
    const cells = rowCells(rows[i].xml);
    const text = cells.map((c) => cellText(c.xml)).join(" ");
    if (!text.trim()) continue;
    if (FOOTNOTE_RE.test(text)) continue;
    if (cells.length < 4) continue;
    return i;
  }
  return -1;
}

function defaultOutputPath(sourceFile) {
  const parsed = path.parse(sourceFile);
  const ts = new Date().toISOString().replace(/[-:T]/g, "").slice(0, 14);
  return path.join(parsed.dir, `${parsed.name}_T30_updated_${ts}.docx`);
}

function updateTabelle30({ sourceFile, outputFile, entries }) {
  if (!sourceFile || typeof sourceFile !== "string") throw new Error("missing sourceFile");
  if (!Array.isArray(entries) || !entries.length) throw new Error("no entries provided");
  if (!/\.(docx|docm)$/i.test(sourceFile)) {
    throw new Error("Only .docx/.docm can be updated without Python/LibreOffice. Convert legacy .doc to .docx first.");
  }

  const zipEntries = readZipEntries(fs.readFileSync(sourceFile));
  const docEntry = zipEntries.find((e) => e.name === "word/document.xml");
  if (!docEntry) throw new Error("word/document.xml not found");

  const documentXml = docEntry.data.toString("utf8");
  const table = findTabelle30Table(documentXml);
  if (!table) throw new Error("Tabelle 30 (Resistance to heat and fire) not found");

  const rows = splitBlocks(table.xml, /<w:tr(?:\s[^>]*)?>/g, "</w:tr>");
  const lastDataIdx = findLastDataRow(rows);
  if (lastDataIdx < 0) throw new Error("no data rows in Tabelle 30");

  const template = rows[lastDataIdx].xml;
  const newRows = entries.map((entry) => populateRow(template, entry)).join("");
  const insertAt = rows[lastDataIdx].end;
  const updatedTable = table.xml.slice(0, insertAt) + newRows + table.xml.slice(insertAt);
  const updatedDocument = documentXml.slice(0, table.start) + updatedTable + documentXml.slice(table.end);

  docEntry.data = Buffer.from(updatedDocument, "utf8");
  const out = path.resolve(outputFile || defaultOutputPath(sourceFile));
  fs.mkdirSync(path.dirname(out), { recursive: true });
  fs.writeFileSync(out, writeZipEntries(zipEntries));
  return { outputFile: out, applied: entries.length, warnings: [] };
}

module.exports = { updateTabelle30 };
