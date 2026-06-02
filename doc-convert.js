/**
 * doc-convert.js — Convert a legacy binary .doc to .docx using MS Word.
 *
 * The pure-Node docx-reader only understands the ZIP-based .docx/.docm format.
 * Old .doc files (OLE compound binary) cannot be parsed reliably in pure JS, and
 * this build intentionally has NO Python and NO LibreOffice. So on Windows we ask
 * the installed MS Word to do the conversion via a short PowerShell + COM call.
 *
 * - Windows only. On any other platform (incl. the macOS dev machine) it throws a
 *   clear error instead of pretending — the caller surfaces that to the user.
 * - The converted .docx is cached in the OS temp dir, keyed by the source path +
 *   size + mtime, so each .doc is only sent through Word once (until it changes).
 * - On-demand only: callers convert a single selected file when it is parsed,
 *   never for every .doc in a folder listing.
 */

const fs = require("fs");
const os = require("os");
const path = require("path");
const crypto = require("crypto");
const { execFileSync } = require("child_process");

const CACHE_DIR = path.join(os.tmpdir(), "pcs-doc-cache");

function psQuote(value) {
  // Single-quoted PowerShell string; escape embedded single quotes by doubling.
  return "'" + String(value).replace(/'/g, "''") + "'";
}

/**
 * Convert srcDoc (.doc) to a .docx and return the .docx path.
 * Throws a clear, user-facing error if Word/Windows is unavailable or it fails.
 */
function convertDocToDocx(srcDoc) {
  if (process.platform !== "win32") {
    throw new Error("Alte .doc-Dateien brauchen MS Word (nur Windows) zur Umwandlung in .docx.");
  }

  let stat;
  try {
    stat = fs.statSync(srcDoc);
  } catch {
    throw new Error(`Quelldatei nicht gefunden: ${srcDoc}`);
  }

  const key = crypto
    .createHash("sha1")
    .update(`${path.resolve(srcDoc)}\0${stat.size}\0${stat.mtimeMs}`)
    .digest("hex");
  fs.mkdirSync(CACHE_DIR, { recursive: true });
  const out = path.join(CACHE_DIR, `${key}.docx`);

  // Reuse a previous conversion of the exact same file (size + mtime unchanged).
  if (fs.existsSync(out)) return out;

  // wdFormatXMLDocument = 16 (.docx). Open read-only, no dialogs, quit cleanly.
  const ps =
    "$ErrorActionPreference='Stop';" +
    "$w=New-Object -ComObject Word.Application;" +
    "$w.Visible=$false;$w.DisplayAlerts=0;" +
    "try{" +
    `$d=$w.Documents.Open(${psQuote(srcDoc)},$false,$true);` +
    `$d.SaveAs(${psQuote(out)},16);` +
    "$d.Close($false)" +
    "}finally{$w.Quit()}";

  try {
    execFileSync(
      "powershell.exe",
      ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", ps],
      { stdio: "ignore", timeout: 120000, windowsHide: true }
    );
  } catch (err) {
    throw new Error(
      `Word-Umwandlung von .doc nach .docx fehlgeschlagen (ist MS Word installiert?): ${err.message}`
    );
  }

  if (!fs.existsSync(out)) {
    throw new Error(`Word lieferte keine .docx-Ausgabe für: ${srcDoc}`);
  }
  return out;
}

module.exports = { convertDocToDocx };
