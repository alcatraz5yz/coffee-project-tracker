/**
 * PCS Folder Scanner
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
const path = require("path");
const os = require("os");
const { db } = require("./db");

// ── Config ──────────────────────────────────────────────────
const PCS_ROOT = process.env.PCS_ROOT || path.join(os.homedir(), "Desktop");
const PROJECT_PREFIX = process.env.PROJECT_PREFIX || "EF";

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

// Folders that are important — 0 files = Blocked instead of Open
const CRITICAL_FOLDERS = new Set(["07", "12", "15"]);

// Word/Excel/PDF extensions we care about
const DOC_EXTS = new Set([".doc", ".docx", ".docm", ".pdf", ".xls", ".xlsx", ".xlsm"]);
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

// Recursively count all real files (skip lock files and hidden files)
function countFiles(dir) {
  let count = 0;
  try {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      if (isLockFile(entry.name)) continue;
      if (entry.isDirectory()) count += countFiles(path.join(dir, entry.name));
      else if (entry.isFile()) count++;
    }
  } catch { /* unreadable folder */ }
  return count;
}

// Recursively find Word docs, return [{file, fullPath, stats, relFromIec}]
function findWordDocs(dir, iecRoot, results = []) {
  try {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      if (isLockFile(entry.name)) continue;
      const full = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        findWordDocs(full, iecRoot, results);
      } else if (entry.isFile() && WORD_EXTS.has(path.extname(entry.name).toLowerCase())) {
        try {
          const stats = fs.statSync(full);
          results.push({
            file: entry.name,
            fullPath: full,
            relFromIec: path.relative(iecRoot, full),
            stats
          });
        } catch { /* skip */ }
      }
    }
  } catch { /* unreadable */ }
  return results;
}

// Derive build label from path relative to "12 Untersuchungen"
function inferBuild(relFrom12) {
  const parts = relFrom12.split(path.sep);
  if (!parts[0]) return "Misc";
  const sub = parts[0].toLowerCase();
  if (sub.includes("prototype") || sub.includes("proto")) return "PT / Pre-Approval";
  if (sub.match(/ts1/)) return "TS1 / Typenprüfung";
  if (sub.match(/ts2/)) return "TS2 / Typenprüfung";
  if (sub.match(/127v/)) return "127V PT1 Brazil";
  if (sub.match(/pvt/)) return "PVT";
  if (sub.match(/oot/)) return "OOT";
  return parts[0]; // use folder name as-is
}

// Infer report state from full relative path
function inferReportState(relFromIec) {
  const lower = relFromIec.toLowerCase();
  if (lower.includes(`${path.sep}archiv${path.sep}`) || lower.endsWith(`${path.sep}archiv`)) return "Archived";
  if (lower.includes("kopie") || lower.includes(" - kopie")) return "Archived";
  if (lower.includes("vorlage") || lower.includes("template")) return "Reference";
  return "Current";
}

// Extract project ID that a Word doc belongs to (from filename)
function inferDocProject(filename) {
  const match = filename.match(/ef[\s-]?(\d{4})/i) || filename.match(/(\d{4})/);
  if (match) return `${PROJECT_PREFIX}${match[1]}`;
  return null;
}

// Find the IEC folder inside a project directory
function findIecFolder(projectDir) {
  const candidates = [
    path.join(projectDir, "Zulassungen", "IEC"),
    path.join(projectDir, "Zulassung", "IEC"),
    path.join(projectDir, "Zulassungen"),
    path.join(projectDir, "IEC")
  ];
  return candidates.find((c) => {
    try { return fs.statSync(c).isDirectory(); } catch { return false; }
  }) || null;
}

// Web-accessible href prefix for a project number
// Prefers an existing evidence-<n> symlink in the project dir, falls back to /files/<n>/
function getWebPrefix(projectNumber) {
  const symlinkPath = path.join(__dirname, `evidence-${projectNumber}`);
  try {
    fs.statSync(symlinkPath);
    return `evidence-${projectNumber}/`;
  } catch {
    return `files/${projectNumber}/`;
  }
}

// ── Per-project scan ─────────────────────────────────────────
function scanProject(projectNumber, projectDir) {
  const projectId = `${PROJECT_PREFIX}${projectNumber}`;
  const iecFolder = findIecFolder(projectDir);
  const webPrefix = getWebPrefix(projectNumber);
  const now = new Date().toISOString();
  const errors = [];

  if (!iecFolder) {
    errors.push(`Kein IEC-Ordner gefunden in ${projectDir}`);
    return { projectId, documentGroups: [], reportVersions: [], errors };
  }

  // ── Document groups from numbered subfolders ────────────────
  const documentGroups = [];
  let iecEntries = [];
  try {
    iecEntries = fs.readdirSync(iecFolder, { withFileTypes: true })
      .filter((e) => e.isDirectory() && !isLockFile(e.name));
  } catch (err) {
    errors.push(`Fehler beim Lesen von ${iecFolder}: ${err.message}`);
  }

  for (const entry of iecEntries) {
    const numMatch = entry.name.match(/^(\d+)/);
    if (!numMatch) continue;

    const num = numMatch[1].padStart(2, "0");
    const area = IEC_FOLDER_MAP[num] || entry.name;
    const folderPath = path.join(iecFolder, entry.name);
    const count = countFiles(folderPath);
    const status = count === 0
      ? (CRITICAL_FOLDERS.has(num) ? "Blocked" : "Open")
      : (num === "12" ? "Current" : "Available");

    const relPath = path.relative(projectDir, folderPath).split(path.sep).join("/");
    const href = `${webPrefix}${relPath}/`;

    documentGroups.push({
      area,
      status,
      count: `${count} ${count === 1 ? "Datei" : "Dateien"}`,
      summary: `${entry.name} — ${count} Datei${count !== 1 ? "en" : ""} gefunden`,
      primary_doc: entry.name,
      href,
      last_scanned: now
    });
  }

  // ── Report versions from 12 Untersuchungen ──────────────────
  const reportVersions = [];
  const unterFolder = iecEntries.find((e) => e.name.match(/^12/))
    ? path.join(iecFolder, iecEntries.find((e) => e.name.match(/^12/)).name)
    : null;

  if (unterFolder) {
    const wordDocs = findWordDocs(unterFolder, iecFolder);

    // Sort by mtime descending so first one per build gets "Current"
    wordDocs.sort((a, b) => b.stats.mtime - a.stats.mtime);
    const seenBuild = new Map();

    for (const doc of wordDocs) {
      const relFrom12 = path.relative(unterFolder, doc.fullPath);
      const build = inferBuild(relFrom12);
      let state = inferReportState(doc.relFromIec);

      // If forced state is Current, only the first (newest) per build is Current
      if (state === "Current") {
        if (seenBuild.has(build)) {
          state = "Archived";
        } else {
          seenBuild.set(build, true);
        }
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

  return { projectId, documentGroups, reportVersions, errors };
}

// ── DB upsert helpers ────────────────────────────────────────
const upsertDocGroup = db.prepare(
  `INSERT INTO document_groups (project_id, area, status, count, summary, primary_doc, href, last_scanned)
   VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
);
const upsertReport = db.prepare(
  `INSERT INTO report_versions (project_id, build, version, modified, size, state, file, href, last_scanned)
   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
);
const upsertProject = db.prepare(
  `INSERT OR IGNORE INTO projects (id, name, phase, health, progress, updated_at)
   VALUES (?, ?, 'Neu gefunden', 'Watch', 0, ?)`
);

// ── Main scan ────────────────────────────────────────────────
function scan(customRoot) {
  const root = customRoot || PCS_ROOT;
  const result = {
    scannedAt: new Date().toISOString(),
    root,
    projects: [],
    totalFiles: 0,
    errors: []
  };

  let rootEntries = [];
  try {
    rootEntries = fs.readdirSync(root, { withFileTypes: true })
      .filter((e) => e.isDirectory() && !isLockFile(e.name) && e.name.match(/^\d/));
  } catch (err) {
    result.errors.push(`Kann PCS-Root nicht lesen (${root}): ${err.message}`);
    return result;
  }

  const doScan = db.transaction(() => {
    for (const entry of rootEntries) {
      const projectNumber = entry.name.replace(/\D/g, "");
      if (!projectNumber) continue;

      const projectId = `${PROJECT_PREFIX}${projectNumber}`;
      const projectDir = path.join(root, entry.name);

      // Ensure project exists in DB (won't overwrite existing)
      upsertProject.run(projectId, `${projectId} (gescannt)`, new Date().toISOString());

      const { documentGroups, reportVersions, errors } = scanProject(projectNumber, projectDir);
      const reportProjectIds = new Set([projectId, ...reportVersions.map((r) => r.project || projectId)]);

      // Replace scanned data — keep manual data untouched
      db.prepare("DELETE FROM document_groups WHERE project_id = ?").run(projectId);
      for (const reportProjectId of reportProjectIds) {
        db.prepare("DELETE FROM report_versions WHERE project_id = ?").run(reportProjectId);
        upsertProject.run(reportProjectId, `${reportProjectId} (gescannt)`, new Date().toISOString());
      }

      for (const g of documentGroups) {
        upsertDocGroup.run(projectId, g.area, g.status, g.count, g.summary, g.primary_doc, g.href, g.last_scanned);
      }
      for (const r of reportVersions) {
        upsertReport.run(r.project || projectId, r.build, r.version, r.modified, r.size, r.state, r.file, r.href, r.last_scanned);
      }

      const fileCount = documentGroups.reduce((sum, g) => sum + parseInt(g.count) || 0, 0);
      result.totalFiles += fileCount;
      result.projects.push({
        id: projectId,
        folder: entry.name,
        documentGroups: documentGroups.length,
        reportVersions: reportVersions.length,
        files: fileCount,
        errors
      });
      result.errors.push(...errors);
    }

    // Write scan log
    db.prepare(
      `INSERT INTO scan_log (scanned_at, projects_found, files_found, notes)
       VALUES (?, ?, ?, ?)`
    ).run(
      result.scannedAt,
      result.projects.length,
      result.totalFiles,
      result.errors.length ? result.errors.join("; ") : null
    );
  });

  doScan();
  return result;
}

module.exports = { scan, PCS_ROOT };

// ── Run standalone ───────────────────────────────────────────
if (require.main === module) {
  console.log(`Scanne: ${PCS_ROOT}\n`);
  const result = scan();
  console.log(`${result.projects.length} Projekt(e) gefunden, ${result.totalFiles} Dateien\n`);
  result.projects.forEach((p) => {
    console.log(`  ${p.id}  (${p.folder})`);
    console.log(`    Dokumentgruppen: ${p.documentGroups}  |  Berichte: ${p.reportVersions}  |  Dateien: ${p.files}`);
    if (p.errors.length) p.errors.forEach((e) => console.log(`    ! ${e}`));
  });
  if (result.errors.length) {
    console.log("\nFehler:");
    result.errors.forEach((e) => console.log(`  - ${e}`));
  }
}
