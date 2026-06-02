"use strict";

const fs = require("fs");
const path = require("path");
const { readDocumentXml, replaceDocumentXml } = require("./docx-reader");
const { convertDocToDocx } = require("./doc-convert");

const GREEN_HEX = "00B050";
const RED_HEX = "FF0000";
const BLUE_HEX = "002060";
const YELLOW_HEX = "FFFF00";
const STATUS_COLORS = { [GREEN_HEX]: "green", [RED_HEX]: "red", [YELLOW_HEX]: "yellow" };
const HEADER_MARKERS = ["TABLE: components", "Table: components"];
const DATE_RE = /\b\d{1,2}\.\d{1,2}\.\d{4}\b|\b\d{4}[-.]\d{1,2}[-.]\d{1,2}\b/g;

function decodeXml(s) {
  return String(s || "")
    .replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, d) => String.fromCharCode(Number(d)))
    .replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCharCode(parseInt(h, 16)));
}

function xmlEscape(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
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

function runColor(rXml) {
  const m = rXml.match(/<w:color\s+w:val="([0-9A-Fa-f]{6})"/);
  return m ? m[1].toUpperCase() : null;
}

function runHighlight(rXml) {
  const m = rXml.match(/<w:highlight\s+w:val="([^"]+)"/);
  return m ? m[1] : null;
}

function runStrike(rXml) {
  return /<w:strike(?:\s[^>]*)?\/>/.test(rXml) || /<w:dstrike(?:\s[^>]*)?\/>/.test(rXml);
}

function paragraphRuns(pXml) {
  const out = [];
  for (const rm of pXml.matchAll(/<w:r(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/g)) {
    const rXml = rm[1];
    const parts = [];
    for (const child of rXml.matchAll(/<w:(t|br|tab)(?:\s[^>]*)?(?:>([\s\S]*?)<\/w:\1>|\/>)/g)) {
      if (child[1] === "t") parts.push(decodeXml(child[2] || ""));
      else if (child[1] === "br") parts.push("\n");
      else if (child[1] === "tab") parts.push("\t");
    }
    const text = parts.join("");
    if (!text) continue;
    out.push({ text, color: runColor(rXml), highlight: runHighlight(rXml), strike: runStrike(rXml) });
  }
  return out;
}

function runsText(runs) {
  return runs.map((r) => r.text).join("");
}

function paragraphText(pXml) {
  return runsText(paragraphRuns(pXml));
}

function paragraphStatus(runs) {
  for (const run of runs) {
    if (run.color && STATUS_COLORS[run.color]) return [STATUS_COLORS[run.color], run.color];
  }
  return null;
}

function cellRuns(tcXml) {
  const out = [];
  const paras = splitBlocks(tcXml, /<w:p(?:\s[^>]*)?>/g, "</w:p>");
  paras.forEach((p, i) => {
    if (i > 0) out.push({ text: "\n", color: null, highlight: null, strike: false });
    out.push(...paragraphRuns(p.xml));
  });
  return out;
}

function cellText(tcXml) {
  return runsText(cellRuns(tcXml)).trim();
}

function cellStatus(tcXml) {
  return paragraphStatus(cellRuns(tcXml));
}

function tableCells(trXml) {
  const raw = splitBlocks(trXml, /<w:tc(?:\s[^>]*)?>/g, "</w:tc>");
  const deduped = [];
  let prev = null;
  for (const cell of raw) {
    const text = cellText(cell.xml);
    if (text !== prev) deduped.push({ ...cell, text });
    prev = text;
  }
  return deduped;
}

function rowIsContinuation(trXml) {
  const first = tableCells(trXml)[0]?.xml || "";
  const borders = first.match(/<w:tcBorders[\s\S]*?<\/w:tcBorders>/)?.[0];
  if (!borders) return false;
  const top = borders.match(/<w:top(?:\s[^>]*)?\/>/)?.[0] || borders.match(/<w:top[\s\S]*?<\/w:top>/)?.[0];
  if (!top) return true;
  const val = top.match(/w:val="([^"]+)"/)?.[1];
  return val === "nil" || val === "none";
}

function parseTableFormat(xml) {
  const rows = [];
  let currentGroupId = 0;
  const tables = splitBlocks(xml, /<w:tbl(?:\s[^>]*)?>/g, "</w:tbl>");
  for (let tblIdx = 0; tblIdx < tables.length; tblIdx++) {
    const tbl = tables[tblIdx];
    if (!HEADER_MARKERS.some((m) => tbl.xml.includes(m))) continue;
    let inData = false;
    const trs = splitBlocks(tbl.xml, /<w:tr(?:\s[^>]*)?>/g, "</w:tr>");
    for (let tableRowIdx = 0; tableRowIdx < trs.length; tableRowIdx++) {
      const tr = trs[tableRowIdx];
      const cells = tableCells(tr.xml);
      const rowText = cells.map((c) => c.text).join(" ").toLowerCase();
      if (!inData) {
        if (rowText.includes("mark(s) of conformity")) inData = true;
        continue;
      }
      if (!cells.length) continue;
      const markCell = cells[cells.length - 1];
      const bodyCells = cells.slice(0, -1);
      if (!markCell.text && bodyCells.every((c) => !c.text)) continue;
      const status = cellStatus(markCell.xml);
      const [colorName, colorHex] = status || ["unknown", null];
      const isCont = rows.length && rowIsContinuation(tr.xml);
      if (!isCont) currentGroupId += 1;
      rows.push({
        rowIdx: rows.length + 1,
        groupId: currentGroupId,
        groupPosition: isCont ? "continuation" : "primary",
        tableIdx: tblIdx,
        tableRowIdx,
        markCellIdx: cells.length - 1,
        cells: bodyCells.map((c) => c.text),
        cellRuns: bodyCells.map((c) => cellRuns(c.xml)),
        markOfConformity: markCell.text,
        markOfConformityRuns: cellRuns(markCell.xml),
        status: colorName,
        statusHex: colorHex,
      });
    }
    break;
  }
  return rows;
}

function parseParagraphFormat(xml) {
  const paras = splitBlocks(xml, /<w:p(?:\s[^>]*)?>/g, "</w:p>");
  const rows = [];
  let inTable = false;
  let inData = false;
  let buffer = [];
  let rowStartIdx = null;

  for (let idx = 0; idx < paras.length; idx++) {
    const runs = paragraphRuns(paras[idx].xml);
    const text = runsText(runs).trim();
    if (!inTable) {
      if (HEADER_MARKERS.some((m) => text.includes(m))) inTable = true;
      continue;
    }
    if (!inData) {
      if (text.toLowerCase().includes("mark(s) of conformity")) inData = true;
      continue;
    }
    if (!text) continue;
    const status = paragraphStatus(runs);
    if (rowStartIdx === null) rowStartIdx = idx;
    buffer.push({ text: runsText(runs), paraIdx: idx, runs });
    if (status) {
      const [colorName, colorHex] = status;
      const mark = buffer[buffer.length - 1];
      const body = buffer.slice(0, -1);
      rows.push({
        rowIdx: rows.length + 1,
        startParaIdx: rowStartIdx,
        endParaIdx: idx,
        cells: body.map((c) => c.text),
        cellRuns: body.map((c) => c.runs),
        markOfConformity: mark.text,
        markOfConformityRuns: mark.runs,
        status: colorName,
        statusHex: colorHex,
      });
      buffer = [];
      rowStartIdx = null;
    }
  }
  return rows;
}

function fileContainsTabelle24(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  try {
    if (ext === ".docx" || ext === ".docm") {
      const xml = readDocumentXml(filePath);
      return HEADER_MARKERS.some((m) => xml.includes(m));
    }
    if (ext === ".doc") {
      const raw = fs.readFileSync(filePath);
      return HEADER_MARKERS.some((m) => raw.includes(Buffer.from(m, "utf16le")) || raw.includes(Buffer.from(m, "utf8")));
    }
  } catch {
    return false;
  }
  return false;
}

function parseTabelle24(filePath) {
  let docxPath = filePath;
  let convertedFromDoc = false;
  if (/\.doc$/i.test(filePath)) {
    // Legacy binary .doc — convert to .docx via MS Word (Windows only).
    docxPath = convertDocToDocx(filePath);
    convertedFromDoc = true;
  } else if (!/\.(docx|docm)$/i.test(filePath)) {
    throw new Error(`Nicht unterstütztes Word-Format für Tabelle 24: ${path.basename(filePath)}`);
  }
  const xml = readDocumentXml(docxPath);
  let rows = parseTableFormat(xml);
  let format = "table";
  if (!rows.length) {
    rows = parseParagraphFormat(xml);
    format = "paragraph";
  }
  return { sourceFile: filePath, convertedFromDoc, format, rowCount: rows.length, rows };
}

function makeRun(text, { color, highlight, strike } = {}) {
  const props = [];
  if (color) props.push(`<w:color w:val="${color}"/>`);
  if (highlight) props.push(`<w:highlight w:val="${highlight}"/>`);
  if (strike) props.push("<w:strike/>");
  const rPr = props.length ? `<w:rPr>${props.join("")}</w:rPr>` : "";
  return `<w:r>${rPr}<w:t xml:space="preserve">${xmlEscape(text)}</w:t></w:r>`;
}

function makeBreak() {
  return "<w:r><w:br/></w:r>";
}

function runsForAction(fullText, action, newDate) {
  const matches = [...String(fullText || "").matchAll(DATE_RE)];
  if (!matches.length) return null;
  const last = matches[matches.length - 1];
  const date = last[0];
  const before = fullText.slice(0, last.index);
  const after = fullText.slice(last.index + date.length);
  const out = [];
  const addText = (text, opts) => {
    if (!text) return;
    const parts = text.split("\n");
    parts.forEach((part, i) => {
      if (i > 0) out.push(makeBreak());
      if (part) out.push(makeRun(part, opts));
    });
  };
  addText(before);
  if (action === "confirm") addText(date, { color: GREEN_HEX });
  else if (action === "expire") addText(date, { color: RED_HEX });
  else if (action === "yellow") addText(date, { color: YELLOW_HEX });
  else if (action === "blue") addText(date, { color: BLUE_HEX });
  else if (action === "update") {
    addText(date, { highlight: "yellow", strike: true });
    if (newDate) {
      out.push(makeBreak());
      out.push(makeRun(newDate, { color: GREEN_HEX }));
    }
  } else {
    return null;
  }
  addText(after);
  return out.join("");
}

function replaceParagraphRuns(pXml, newRunsXml) {
  return pXml.replace(/(<w:p(?:\s[^>]*)?>)[\s\S]*?(<\/w:p>)/, `$1${newRunsXml}$2`);
}

function updateParagraphAt(xml, paraIdx, action, newDate) {
  const paras = splitBlocks(xml, /<w:p(?:\s[^>]*)?>/g, "</w:p>");
  const target = paras[paraIdx];
  if (!target) return { xml, ok: false, reason: "paraIdx out of range" };
  const text = paragraphText(target.xml);
  const runs = runsForAction(text, action, newDate);
  if (!runs) return { xml, ok: false, reason: "no date found in paragraph" };
  const next = replaceParagraphRuns(target.xml, runs);
  return { xml: xml.slice(0, target.start) + next + xml.slice(target.end), ok: true };
}

function defaultOutputPath(sourceFile) {
  const parsed = path.parse(sourceFile);
  const ts = new Date().toISOString().replace(/[-:T]/g, "").slice(0, 14);
  return path.join(parsed.dir, `${parsed.name}_updated_${ts}.docx`);
}

function updateTabelle24({ sourceFile, outputFile, changes }) {
  if (!/\.(docx|docm)$/i.test(sourceFile)) {
    throw new Error("Only .docx/.docm can be updated without Python/LibreOffice. Convert legacy .doc to .docx first.");
  }
  let xml = readDocumentXml(sourceFile);
  let applied = 0;
  const skipped = [];
  for (const change of changes || []) {
    if (change.tableIdx !== undefined && change.tableIdx !== null) {
      skipped.push({ change, reason: "table-format Tabelle 24 updates are not implemented in the Node updater yet" });
      continue;
    }
    const result = updateParagraphAt(xml, change.paraIdx, change.action, change.newDate);
    xml = result.xml;
    if (result.ok) applied += 1;
    else skipped.push({ change, reason: result.reason });
  }
  const out = path.resolve(outputFile || defaultOutputPath(sourceFile));
  fs.mkdirSync(path.dirname(out), { recursive: true });
  replaceDocumentXml(sourceFile, out, xml);
  return { outputFile: out, applied, skipped };
}

module.exports = { fileContainsTabelle24, parseTabelle24, updateTabelle24 };
