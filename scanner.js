/**
 * PCS Folder Scanner — async, incremental, parallel
 *
 * Scans a root directory (P:\PCS on Windows, ~/Desktop for Mac testing)
 * and writes document groups + report versions into the database.
 *
 * Manual data is NEVER touched:
 *   ziffern status, fachfreigabe, tasks, risks, certification, closeout
 *
 * Usage:
 *   node scanner.js                        — uses PCS_ROOT env var or ~/Desktop
 *   PCS_ROOT=/Volumes/PCS node scanner.js  — explicit path
 */

const fs = require("fs");
const fsp = fs.promises;
const path = require("path");
const os = require("os");
const { db } = require("./db");

// ── Config ──────────────────────────────────────────────────
const PCS_ROOT = process.env.PCS_ROOT || path.join(os.homedir(), "Desktop");
const PROJECT_PREFIX = process.env.PROJECT_PREFIX || "EF";
const CONCURRENCY = 8;

// IEC subfolder number → display area
const IEC_FOLDER_MAP = {
  "01": "Administration",
  "02": "Standards / Changes",
  "03": "Manual / Labels",
  "04": "Typenschild",
  "05": "Order / Scope",
  "06": "BOM / Exploded Views",
  "07": "Electronics",
  "08": "Device Schemas",
  "09": "Bautelliste",
  "10": "Components",
  "11": "Material / GWT",
  "12": "Investigations",
  "13": "Photos",
  "14": "PAK / LFGB",
  "15": "CB / Safety / EMC / ErP",
  "16": "Country Implementation"
};

const CRITICAL_FOLDERS = new Set(["07", "12", "15"]);
const WORD_EXTS = new Set([".doc", ".docx", ".docm"]);

// ── Helpers ──────────────────────────────────────────────────
function fmtSize(bytes) {
  if (bytes >= 1e6) return `${(bytes / 1e6).toFixed(1)} MB`;
  if (bytes >= 1e3) return `${Math.round(bytes / 1e3)} KB`;
  return `${bytes} B`;
}

function fmtDate(mtime) {
  return mtime.toISOString().slice(0, 16).replace("T", " ");
}

function isLockFile(name) {
  return name.startsWith("~$") || name.startsWith(".");
}

async function countFilesAsync(dir) {
  try {
    const entries = await fsp.readdir(dir, { withFileTypes: true });
    const counts = await Promise.all(entries.map(async (entry) => {
      if (isLockFile(entry.name)) return 0;
      if (entry.isDirectory()) return countFilesAsync(path.join(dir, entry.name));
      if (entry.isFile()) return 1;
      return 0;
    }));
    return counts.reduce((a, b) => a + b, 0);
  } catch {
    return 0;
  }
}

async function findWordDocsAsync(dir, iecRoot) {
  try {
    const entries = await fsp.readdir(dir, { withFileTypes: true });
    const nested = await Promise.all(entries.map(async (entry) => {
      if (isLockFile(entry.name)) return [];
      const full = path.join(dir, entry.name);
      if (entry.isDirectory()) return findWordDocsAsync(full, iecRoot);
      if (entry.isFile() && WORD_EXTS.has(path.extname(entry.name).toLowerCase())) {
        try {
          const stats = await fsp.stat(full);
          return [{ file: entry.name, fullPath: full, relFromIec: path.relative(iecRoot, full), stats }];
        } catch { return []; }
      }
      return [];
    }));
    return nested.flat();
  } catch {
    return [];
  }
}

function inferBuild(relFrom12) {
  const parts = relFrom12.split(path.sep);
  if (!parts[0]) return "Misc";
  const sub = parts[0].toLowerCase();
  if (sub.includes("prototype") || sub.includes("proto")) return "PT / Pre-Approval";
  if (sub.match(/ts2/)) return "TS2 / Typenprüfung";
  if (sub.match(/ts1/)) return "TS1 / Typenprüfung";
  if (sub.match(/127v/)) return "127V PT1 Brazil";
  if (sub.match(/pvt/)) return "PVT";
  if (sub.match(/oot/)) return "OOT";
  return parts[0];
}

function inferReportState(relFromIec) {
  const lower = relFromIec.toLowerCase();
  if (lower.includes(`${path.sep}archiv${path.sep}`) || lower.endsWith(`${path.sep}archiv`)) return "Archived";
  if (lower.includes("kopie") || lower.includes(" - kopie")) return "Archived";
  if (lower.includes("vorlage") || lower.includes("template")) return "Reference";
  return "Current";
}

function inferDocProject(filename) {
  const match = filename.match(/ef[\s_-]?(\d{4})/i);
  if (match) return `${PROJECT_PREFIX}${match[1]}`;
  return null;
}

async function findIecFolderAsync(projectDir) {
  const candidates = [
    path.join(projectDir, "Zulassungen", "IEC"),
    path.join(projectDir, "Zulassung", "IEC"),
    path.join(projectDir, "Zulassungen"),
    path.join(projectDir, "IEC")
  ];
  for (const c of candidates) {
    try {
      const st = await fsp.stat(c);
      if (st.isDirectory()) return c;
    } catch { /* try next */ }
  }
  return null;
}

function encodePathSegment(value) {
  return String(value).split(path.sep).map(encodeURIComponent).join("/");
}

function isProjectFolderName(name) {
  return /^\d/.test(name) || new RegExp(`^${PROJECT_PREFIX}[\\s_-]*\\d`, "i").test(name);
}

// Extract all EF numbers from a folder name.
// "EF254+698 Nivona"       → ["254", "698"]
// "EF230+231+232 NNSA"     → ["230", "231", "232"]
// "1157 EF1157 CoffeeB"    → ["1157"]
// "EF 1175"                → ["1175"]
function extractProjectNumbers(folderName) {
  const efMatch = folderName.match(/^EF[\s_-]*(\d+(?:[+]\d+)*)/i);
  if (efMatch) return efMatch[1].split("+").map((n) => n.trim()).filter(Boolean);
  const digitMatch = folderName.match(/^(\d+)/);
  if (digitMatch) return [digitMatch[1]];
  return [];
}

function getWebPrefix(projectNumber, projectFolderName) {
  const symlinkPath = path.join(__dirname, `evidence-${projectNumber}`);
  try {
    fs.statSync(symlinkPath);
    return `evidence-${projectNumber}/`;
  } catch {
    return `files/${encodePathSegment(projectFolderName)}/`;
  }
}

// ── Per-project scan ─────────────────────────────────────────
async function scanProject(projectNumber, projectDir, storedMtimes, allNumbers = []) {
  const projectId = `${PROJECT_PREFIX}${projectNumber}`;
  const allProjectIds = allNumbers.length > 0
    ? allNumbers.map((n) => `${PROJECT_PREFIX}${n}`)
    : [projectId];
  const iecFolder = await findIecFolderAsync(projectDir);
  const projectFolderName = path.basename(projectDir);
  const webPrefix = getWebPrefix(projectNumber, projectFolderName);
  const now = new Date().toISOString();
  const errors = [];

  if (!iecFolder) {
    errors.push(`Kein IEC-Ordner gefunden in ${projectDir}`);
    return { projectId, allProjectIds, documentGroups: [], reportVersions: [], errors, skipped: 0, scanned: 0 };
  }

  let iecEntries = [];
  try {
    iecEntries = (await fsp.readdir(iecFolder, { withFileTypes: true }))
      .filter((e) => e.isDirectory() && !isLockFile(e.name));
  } catch (err) {
    errors.push(`Fehler beim Lesen von ${iecFolder}: ${err.message}`);
    return { projectId, documentGroups: [], reportVersions: [], errors, skipped: 0, scanned: 0 };
  }

  let skipped = 0;
  let scanned = 0;
  const documentGroups = [];
  const skippedAreas = [];

  // Process all IEC areas in parallel, with mtime-based skip
  await Promise.all(iecEntries.map(async (entry) => {
    const numMatch = entry.name.match(/^(\d+)/);
    if (!numMatch) return;

    const num = numMatch[1].padStart(2, "0");
    const area = IEC_FOLDER_MAP[num] || entry.name;
    const folderPath = path.join(iecFolder, entry.name);

    let currentMtime = null;
    try {
      const st = await fsp.stat(folderPath);
      currentMtime = st.mtime.toISOString();
    } catch { return; }

    if (storedMtimes.has(area) && storedMtimes.get(area) === currentMtime) {
      skippedAreas.push(area);
      skipped++;
      return;
    }

    scanned++;
    const count = await countFilesAsync(folderPath);
    const status = count === 0
      ? (CRITICAL_FOLDERS.has(num) ? "Blocked" : "Open")
      : (num === "12" ? "Current" : "Available");

    const relPathFromProject = path.relative(projectDir, folderPath).split(path.sep).join("/");
    const href = `${webPrefix}${relPathFromProject}/`;

    documentGroups.push({
      area,
      status,
      count: `${count} ${count === 1 ? "Datei" : "Dateien"}`,
      summary: `${entry.name} — ${count} Datei${count !== 1 ? "en" : ""} gefunden`,
      primary_doc: entry.name,
      href,
      last_scanned: now,
      folder_mtime: currentMtime
    });
  }));

  // Report versions from folder 12 Untersuchungen
  const reportVersions = [];
  let reportsScanned = false;
  const unterEntry = iecEntries.find((e) => e.name.match(/^12/));
  if (unterEntry) {
    reportsScanned = true;
    const unterFolder = path.join(iecFolder, unterEntry.name);
    const wordDocs = await findWordDocsAsync(unterFolder, iecFolder);
    wordDocs.sort((a, b) => b.stats.mtime - a.stats.mtime);
    const seenBuild = new Map();

    for (const doc of wordDocs) {
      const relFrom12 = path.relative(unterFolder, doc.fullPath);
      const build = inferBuild(relFrom12);
      let state = inferReportState(doc.relFromIec);

      if (state === "Current") {
        if (seenBuild.has(build)) state = "Archived";
        else seenBuild.set(build, true);
      }

      const docProject = inferDocProject(doc.file) || projectId;
      const relFromProjectDir = path.relative(projectDir, doc.fullPath).split(path.sep).join("/");
      const href = `${webPrefix}${relFromProjectDir}`;

      reportVersions.push({
        project: docProject,
        build,
        version: fmtDate(doc.stats.mtime),
        modified: fmtDate(doc.stats.mtime),
        size: fmtSize(doc.stats.size),
        state,
        file: doc.file,
        href,
        last_scanned: now
      });
    }
  }

  return { projectId, allProjectIds, documentGroups, reportVersions, reportsScanned, errors, skipped, scanned, skippedAreas };
}

// ── Prepared statements ──────────────────────────────────────
const stmts = {
  upsertProject: db.prepare(
    `INSERT OR IGNORE INTO projects (id, name, phase, health, progress, updated_at)
     VALUES (?, ?, 'Neu gefunden', 'Watch', 0, ?)`
  ),
  touchLastScanned: db.prepare(
    "UPDATE document_groups SET last_scanned = ? WHERE project_id = ? AND area = ?"
  ),
  deleteDocArea: db.prepare(
    "DELETE FROM document_groups WHERE project_id = ? AND area = ?"
  ),
  insertDocGroup: db.prepare(
    `INSERT INTO document_groups (project_id, area, status, count, summary, primary_doc, href, last_scanned, folder_mtime)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
  ),
  deleteReports: db.prepare("DELETE FROM report_versions WHERE project_id = ?"),
  insertReport: db.prepare(
    `INSERT INTO report_versions (project_id, build, version, modified, size, state, file, href, last_scanned)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
  ),
  getStoredMtimes: db.prepare(
    "SELECT area, folder_mtime FROM document_groups WHERE project_id = ? AND folder_mtime IS NOT NULL"
  ),
  insertScanLog: db.prepare(
    `INSERT INTO scan_log (scanned_at, projects_found, files_found, notes) VALUES (?, ?, ?, ?)`
  )
};

// ── DB write for one project (inside a transaction) ──────────
function writeProjectToDB(res, now) {
  const { projectId, allProjectIds, documentGroups, reportVersions, reportsScanned, skippedAreas } = res;
  const ids = allProjectIds?.length ? allProjectIds : [projectId];

  for (const id of ids) {
    stmts.upsertProject.run(id, `${id} (gescannt)`, now);
  }

  // Touch last_scanned for unchanged areas (all IDs share the same folder)
  for (const area of (skippedAreas || [])) {
    for (const id of ids) {
      stmts.touchLastScanned.run(now, id, area);
    }
  }

  // Replace only re-scanned areas — write same document groups for all IDs
  for (const g of documentGroups) {
    for (const id of ids) {
      stmts.deleteDocArea.run(id, g.area);
      stmts.insertDocGroup.run(
        id, g.area, g.status, g.count, g.summary,
        g.primary_doc, g.href, g.last_scanned, g.folder_mtime || null
      );
    }
  }

  if (reportsScanned) {
    const reportProjectIds = new Set([...ids, ...reportVersions.map((r) => r.project || projectId)]);
    for (const rId of reportProjectIds) {
      stmts.upsertProject.run(rId, `${rId} (gescannt)`, now);
      stmts.deleteReports.run(rId);
    }
    for (const r of reportVersions) {
      stmts.insertReport.run(
        r.project || projectId, r.build, r.version, r.modified,
        r.size, r.state, r.file, r.href, r.last_scanned
      );
    }
  }
}

// ── Main scan ────────────────────────────────────────────────
async function scan(customRoot, onProgress) {
  const root = customRoot || PCS_ROOT;
  const now = new Date().toISOString();
  const result = { scannedAt: now, root, projects: [], totalFiles: 0, errors: [] };

  let rootEntries = [];
  try {
    rootEntries = (await fsp.readdir(root, { withFileTypes: true }))
      .filter((e) => e.isDirectory() && !isLockFile(e.name) && isProjectFolderName(e.name));
  } catch (err) {
    result.errors.push(`Kann PCS-Root nicht lesen (${root}): ${err.message}`);
    return result;
  }

  const total = rootEntries.length;
  let done = 0;

  for (let i = 0; i < rootEntries.length; i += CONCURRENCY) {
    const batch = rootEntries.slice(i, i + CONCURRENCY);

    const batchResults = await Promise.all(batch.map(async (entry) => {
      const numbers = extractProjectNumbers(entry.name);
      if (!numbers.length) return null;

      const primaryNumber = numbers[0];
      const primaryId = `${PROJECT_PREFIX}${primaryNumber}`;
      const projectDir = path.join(root, entry.name);

      // Load stored mtimes using primary project ID
      const storedMtimes = new Map(
        stmts.getStoredMtimes.all(primaryId).map((r) => [r.area, r.folder_mtime])
      );

      const res = await scanProject(primaryNumber, projectDir, storedMtimes, numbers);

      done++;
      if (onProgress) onProgress({ done, total, current: primaryId });
      return res;
    }));

    // Write batch to DB in one transaction
    db.transaction(() => {
      for (const res of batchResults) {
        if (res) writeProjectToDB(res, now);
      }
    })();

    for (const res of batchResults) {
      if (!res) continue;
      const fileCount = res.documentGroups.reduce((sum, g) => sum + (parseInt(g.count) || 0), 0);
      result.totalFiles += fileCount;
      result.projects.push({
        id: res.projectId,
        documentGroups: res.documentGroups.length,
        reportVersions: res.reportVersions.length,
        files: fileCount,
        skipped: res.skipped,
        scanned: res.scanned,
        errors: res.errors
      });
      result.errors.push(...res.errors);
    }
  }

  stmts.insertScanLog.run(
    result.scannedAt,
    result.projects.length,
    result.totalFiles,
    result.errors.length ? result.errors.join("; ") : null
  );

  return result;
}

module.exports = { scan, PCS_ROOT };

// ── Run standalone ───────────────────────────────────────────
if (require.main === module) {
  console.log(`Scanne: ${PCS_ROOT}\n`);
  scan(null, ({ done, total, current }) => {
    process.stdout.write(`\r  ${done}/${total}  ${current}          `);
  }).then((result) => {
    console.log(`\n\n${result.projects.length} Projekt(e) gefunden, ${result.totalFiles} Dateien\n`);
    result.projects.forEach((p) => {
      const skipInfo = p.skipped > 0 ? `  übersprungen:${p.skipped}` : "";
      console.log(`  ${p.id}  Gruppen:${p.documentGroups}  Berichte:${p.reportVersions}  Dateien:${p.files}${skipInfo}`);
      if (p.errors.length) p.errors.forEach((e) => console.log(`    ! ${e}`));
    });
    if (result.errors.length) {
      console.log("\nFehler:");
      result.errors.forEach((e) => console.log(`  - ${e}`));
    }
  }).catch((err) => {
    console.error("Scan fehlgeschlagen:", err);
    process.exit(1);
  });
}
