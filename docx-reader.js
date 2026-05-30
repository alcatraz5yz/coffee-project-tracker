/**
 * docx-reader.js — Pure-Node reader for Word .docx / .docm tables.
 *
 * No external dependencies and NO LibreOffice: a .docx/.docm is a ZIP whose
 * `word/document.xml` holds the body. Node's built-in zlib inflates it, and we
 * walk the table XML directly. This replaces the Python (python-docx + soffice)
 * path so the Windows build needs nothing beyond Node.
 *
 * .docm (macro-enabled) has the exact same document.xml structure as .docx, so
 * it is read directly — the macros (vbaProject.bin) are simply ignored.
 */

"use strict";

const fs = require("fs");
const zlib = require("zlib");

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

// ── Minimal ZIP: read one entry via the End-Of-Central-Directory record ──────
function readZipEntry(buf, target) {
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0; i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error("not a zip (no EOCD)");
  const count = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  for (let k = 0; k < count; k++) {
    if (buf.readUInt32LE(p) !== 0x02014b50) throw new Error("bad central directory");
    const method = buf.readUInt16LE(p + 10);
    const compSize = buf.readUInt32LE(p + 20);
    const nameLen = buf.readUInt16LE(p + 28);
    const extraLen = buf.readUInt16LE(p + 30);
    const commLen = buf.readUInt16LE(p + 32);
    const localOff = buf.readUInt32LE(p + 42);
    const name = buf.toString("utf8", p + 46, p + 46 + nameLen);
    if (name === target) {
      if (buf.readUInt32LE(localOff) !== 0x04034b50) throw new Error("bad local header");
      const lNameLen = buf.readUInt16LE(localOff + 26);
      const lExtraLen = buf.readUInt16LE(localOff + 28);
      const dataStart = localOff + 30 + lNameLen + lExtraLen;
      const comp = buf.subarray(dataStart, dataStart + compSize);
      return method === 0 ? comp : zlib.inflateRawSync(comp);
    }
    p = p + 46 + nameLen + extraLen + commLen;
  }
  throw new Error(`zip entry not found: ${target}`);
}

function readZipEntries(buf) {
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0; i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error("not a zip (no EOCD)");
  const count = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  const entries = [];
  for (let k = 0; k < count; k++) {
    if (buf.readUInt32LE(p) !== 0x02014b50) throw new Error("bad central directory");
    const method = buf.readUInt16LE(p + 10);
    const compSize = buf.readUInt32LE(p + 20);
    const nameLen = buf.readUInt16LE(p + 28);
    const extraLen = buf.readUInt16LE(p + 30);
    const commLen = buf.readUInt16LE(p + 32);
    const localOff = buf.readUInt32LE(p + 42);
    const name = buf.toString("utf8", p + 46, p + 46 + nameLen);
    if (buf.readUInt32LE(localOff) !== 0x04034b50) throw new Error("bad local header");
    const lNameLen = buf.readUInt16LE(localOff + 26);
    const lExtraLen = buf.readUInt16LE(localOff + 28);
    const dataStart = localOff + 30 + lNameLen + lExtraLen;
    const comp = buf.subarray(dataStart, dataStart + compSize);
    const data = method === 0 ? Buffer.from(comp) : zlib.inflateRawSync(comp);
    entries.push({ name, data });
    p = p + 46 + nameLen + extraLen + commLen;
  }
  return entries;
}

function writeZipEntries(entries) {
  const locals = [];
  const centrals = [];
  let offset = 0;
  for (const entry of entries) {
    const nameBuf = Buffer.from(entry.name, "utf8");
    const data = Buffer.isBuffer(entry.data) ? entry.data : Buffer.from(entry.data);
    const comp = zlib.deflateRawSync(data);
    const crc = crc32(data);

    const local = Buffer.alloc(30 + nameBuf.length);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0, 6);
    local.writeUInt16LE(8, 8);
    local.writeUInt16LE(0, 10);
    local.writeUInt16LE(0, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(comp.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    local.writeUInt16LE(0, 28);
    nameBuf.copy(local, 30);
    locals.push(local, comp);

    const central = Buffer.alloc(46 + nameBuf.length);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0, 8);
    central.writeUInt16LE(8, 10);
    central.writeUInt16LE(0, 12);
    central.writeUInt16LE(0, 14);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(comp.length, 20);
    central.writeUInt32LE(data.length, 24);
    central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt16LE(0, 30);
    central.writeUInt16LE(0, 32);
    central.writeUInt16LE(0, 34);
    central.writeUInt16LE(0, 36);
    central.writeUInt32LE(0, 38);
    central.writeUInt32LE(offset, 42);
    nameBuf.copy(central, 46);
    centrals.push(central);
    offset += local.length + comp.length;
  }

  const centralSize = centrals.reduce((sum, b) => sum + b.length, 0);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(centralSize, 12);
  eocd.writeUInt32LE(offset, 16);
  eocd.writeUInt16LE(0, 20);
  return Buffer.concat([...locals, ...centrals, eocd]);
}

function decodeXml(s) {
  return s
    .replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, d) => String.fromCharCode(Number(d)))
    .replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCharCode(parseInt(h, 16)));
}

// ── Run / cell / row / table extraction ──────────────────────────────────────
// A cell's runs (<w:r><w:t>) are joined with a space, so visually-separate runs
// like "ABS" + "Ecoex GAR-011" + "black" become "ABS Ecoex GAR-011 black"
// instead of "ABSEcoex GAR-011black". Paragraphs (<w:p>) become newlines.
function paragraphText(pXml) {
  let x = pXml.replace(/<w:tab\/?>/g, " ").replace(/<w:br\s*\/?>/g, "\n");
  const parts = [];
  for (const m of x.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)) parts.push(decodeXml(m[1]));
  return parts.join(" ");
}

function cellText(tcXml) {
  const paras = [];
  const re = /<w:p(?:\s[^>]*)?>([\s\S]*?)<\/w:p>/g;
  let m;
  while ((m = re.exec(tcXml))) paras.push(paragraphText(m[1]));
  // also handle self-closing empty paragraphs <w:p/>
  return paras.join("\n").replace(/[ \t]+/g, " ").replace(/ *\n */g, "\n").trim();
}

function runColor(rXml) {
  const m = rXml.match(/<w:color\s+w:val="([0-9A-Fa-f]{6})"/);
  return m ? m[1].toUpperCase() : null;
}

// Runs with colour info, for optional coloured rendering parity with the Mac build.
function cellRuns(tcXml) {
  const out = [];
  const paraRe = /<w:p(?:\s[^>]*)?>([\s\S]*?)<\/w:p>/g;
  let pm; let first = true;
  while ((pm = paraRe.exec(tcXml))) {
    if (!first) out.push({ text: "\n", color: null });
    first = false;
    const pXml = pm[1].replace(/<w:tab\/?>/g, " ").replace(/<w:br\s*\/?>/g, "\n");
    const runRe = /<w:r(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/g;
    let rm;
    while ((rm = runRe.exec(pXml))) {
      const rXml = rm[1];
      let text = "";
      for (const t of rXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)) text += decodeXml(t[1]);
      if (!text) continue;
      out.push({ text, color: runColor(rXml) });
    }
  }
  return out;
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
    out.push(xml.slice(start, end + closeTag.length));
    i = end + closeTag.length;
  }
  return out;
}

function tableRows(tblXml) {
  return splitBlocks(tblXml, /<w:tr(?:\s[^>]*)?>/g, "</w:tr>").map((tr) => {
    const tcs = splitBlocks(tr, /<w:tc(?:\s[^>]*)?>/g, "</w:tc>");
    return tcs.map((tc) => ({ text: cellText(tc), runs: cellRuns(tc) }));
  });
}

/** Return every table in the document as an array of rows (each row = array of {text,runs}). */
function readTables(filePath) {
  const buf = fs.readFileSync(filePath);
  const xml = readZipEntry(buf, "word/document.xml").toString("utf8");
  const tbls = splitBlocks(xml, /<w:tbl(?:\s[^>]*)?>/g, "</w:tbl>");
  return tbls.map(tableRows);
}

/** Raw document.xml text (used for marker/heading scans). */
function readDocumentXml(filePath) {
  return readZipEntry(fs.readFileSync(filePath), "word/document.xml").toString("utf8");
}

function replaceDocumentXml(filePath, outputPath, documentXml) {
  const entries = readZipEntries(fs.readFileSync(filePath));
  const doc = entries.find((e) => e.name === "word/document.xml");
  if (!doc) throw new Error("word/document.xml not found");
  doc.data = Buffer.from(documentXml, "utf8");
  fs.writeFileSync(outputPath, writeZipEntries(entries));
}

module.exports = { readTables, readDocumentXml, readZipEntries, writeZipEntries, replaceDocumentXml };
