"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");

const VDE_CERT_URL = "https://www.vde.com/tic-en/marks-and-certificates/vde-approved-products/certificate?id={cert}&type=zertreg%7Ccertificate";
const DOWNLOAD_DIR = path.join(os.tmpdir(), "vde-lookups");

function normalizeIso(s) {
  s = String(s || "").trim();
  if (/^20\d{2}-\d{2}-\d{2}$/.test(s)) return s;
  const m = s.match(/^(\d{2})\.(\d{2})\.(20\d{2})$/);
  return m ? `${m[3]}-${m[2]}-${m[1]}` : null;
}

function compareDates(onlineIso, currentIso) {
  const online = normalizeIso(onlineIso);
  const current = normalizeIso(currentIso);
  if (!online || !current) return "unknown";
  if (online > current) return "newer";
  if (online === current) return "current";
  return "older";
}

function decodeHtml(s) {
  return String(s || "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&#(\d+);/g, (_, d) => String.fromCharCode(Number(d)));
}

function stripTags(s) {
  return decodeHtml(String(s || "").replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim());
}

function absolutize(url) {
  try {
    return new URL(decodeHtml(url), "https://www.vde.com").toString();
  } catch {
    return decodeHtml(url);
  }
}

function extractPdfEntries(html) {
  const entries = [];
  for (const m of html.matchAll(/<a\b([^>]*)>([\s\S]*?)<\/a>/gi)) {
    const attrs = m[1] || "";
    const href = attrs.match(/\bhref\s*=\s*["']([^"']+)["']/i)?.[1];
    if (!href) continue;
    const label = stripTags(m[2]) || "PDF";
    const hay = `${href} ${label}`.toLowerCase();
    if (!hay.includes(".pdf") && !hay.includes("zertreg-proxy.vde.com") && !hay.includes("download")) continue;
    entries.push({ label: label.slice(0, 80), url: absolutize(href) });
  }
  return entries;
}

function chooseAppendix(entries) {
  if (!entries.length) return null;
  return entries.find((e) => /\b200\b/.test(`${e.label} ${e.url}`))
    || entries.find((e) => /\b100\b/.test(`${e.label} ${e.url}`))
    || entries[entries.length - 1];
}

function pdfTextBytes(buf) {
  const latin = buf.toString("latin1");
  const utf16 = [];
  for (let i = 0; i + 1 < buf.length; i += 2) {
    const c = buf.readUInt16BE(i);
    if ((c >= 32 && c <= 126) || c === 10 || c === 13 || c === 9) utf16.push(String.fromCharCode(c));
  }
  return `${latin}\n${utf16.join("")}`;
}

function extractUpdatedDate(pdfPath) {
  try {
    const text = pdfTextBytes(fs.readFileSync(pdfPath));
    let m = text.match(/(?:letzte\s*[AÄ]nderung|updated)\s+(20\d{2}-\d{2}-\d{2})/i);
    if (m) return m[1];
    m = text.match(/(?:letzte\s*[AÄ]nderung|updated)\s+(\d{2}\.\d{2}\.\d{4})/i);
    if (m) return normalizeIso(m[1]);
    m = text.match(/(?:letzte\s*[AÄ]nderung|updated)[\s\S]{0,300}?(20\d{2}-(?:0[1-9]|1[0-2])-(?:0[1-9]|[12]\d|3[01]))/i);
    if (m) return m[1];
    const dates = [...text.matchAll(/20\d{2}-(?:0[1-9]|1[0-2])-(?:0[1-9]|[12]\d|3[01])/g)].map((x) => x[0]);
    return dates.length === 1 ? dates[0] : null;
  } catch {
    return null;
  }
}

async function fetchWithTimeout(url, options = {}, timeoutMs = 30000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(url, { ...options, signal: controller.signal });
  } finally {
    clearTimeout(timer);
  }
}

async function lookupVde({ certNumber, currentDate }) {
  const cert = String(certNumber || "").trim();
  if (!cert) throw new Error("missing certNumber");

  const result = {
    certNumber: cert,
    found: false,
    pdfList: [],
    chosen: null,
    downloadedPdf: null,
    onlineDate: null,
    comparison: "unknown",
  };

  const certUrl = VDE_CERT_URL.replace("{cert}", encodeURIComponent(cert));
  const page = await fetchWithTimeout(certUrl, {
    headers: { "User-Agent": "Mozilla/5.0", "Accept": "text/html,application/xhtml+xml" },
  }, 30000);
  if (!page.ok) {
    result.error = `certificate page failed: HTTP ${page.status}`;
    return result;
  }
  const html = await page.text();
  result.found = html.includes(`Certificate number: ${cert}`) || html.includes(`Certificate number:${cert}`);
  result.pdfList = extractPdfEntries(html);
  if (!result.found || !result.pdfList.length) {
    if (result.found) result.error = "certificate found, but no PDF link was exposed in the static VDE page";
    return result;
  }

  const chosen = chooseAppendix(result.pdfList);
  result.chosen = chosen;
  const pdf = await fetchWithTimeout(chosen.url, { headers: { "User-Agent": "Mozilla/5.0", "Referer": "https://www.vde.com/" } }, 30000);
  const bytes = Buffer.from(await pdf.arrayBuffer());
  if (!pdf.ok || bytes.subarray(0, 5).toString() !== "%PDF-") {
    result.error = `download failed: HTTP ${pdf.status}, ${bytes.length} bytes`;
    return result;
  }

  fs.mkdirSync(DOWNLOAD_DIR, { recursive: true });
  const appNum = (chosen.label.match(/\d{2,3}/) || chosen.url.match(/\d{2,3}/) || ["x"])[0];
  const savePath = path.join(DOWNLOAD_DIR, `vde_${cert}_appendix_${appNum}.pdf`);
  fs.writeFileSync(savePath, bytes);
  result.downloadedPdf = savePath;
  result.onlineDate = extractUpdatedDate(savePath);
  result.comparison = compareDates(result.onlineDate, currentDate);
  return result;
}

module.exports = { lookupVde, extractUpdatedDate, compareDates };
