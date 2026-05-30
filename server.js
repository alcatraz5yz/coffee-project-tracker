const express = require("express");
const path = require("path");
const os = require("os");
const fs = require("fs");
const { spawn, spawnSync } = require("child_process");
const XLSX = require("xlsx");
const {
  getProjects,
  getProject,
  updateZifferStatus,
  updateZifferReason,
  updateFachfreigabeGate,
  updateFachfreigabeMeta,
  updateTaskStatus,
  updateArchiveLocation,
  updateProjectNo,
  updateSwVersion,
  updateHwVersion,
  updateMachineType,
  updateMachineUse,
  updateColors,
  getShipments,
  getAllOpenShipments,
  upsertShipment,
  deleteShipment,
  getLabs,
  upsertLab,
  getArchiveItems,
  upsertArchiveItem,
  deleteArchiveItem,
  db
} = require("./db");
const { scan, PCS_ROOT } = require("./scanner");

const app = express();
const PORT = process.env.PORT || 8090;
const evidenceRoots = new Map();

app.use(express.json());

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function hrefJoin(baseUrl, name, isDirectory) {
  const cleanBase = baseUrl.endsWith("/") ? baseUrl : `${baseUrl}/`;
  return `${cleanBase}${encodeURIComponent(name)}${isDirectory ? "/" : ""}`;
}

function sendDirectoryListing(req, res, root, mountPath) {
  const mount = mountPath.endsWith("/") ? mountPath.slice(0, -1) : mountPath;
  const reqPath = decodeURIComponent(req.originalUrl.split("?")[0]);
  const relUrl = reqPath.startsWith(mount) ? reqPath.slice(mount.length) : reqPath;
  const relPath = relUrl.replace(/^\/+/, "");
  const safeRoot = fs.realpathSync(root);
  const target = path.resolve(safeRoot, relPath);

  if (!target.startsWith(safeRoot)) return res.status(403).send("Forbidden");
  if (!fs.existsSync(target)) return false;

  const stat = fs.statSync(target);
  if (!stat.isDirectory()) return false;

  const entries = fs.readdirSync(target, { withFileTypes: true })
    .filter((entry) => !entry.name.startsWith(".") && !entry.name.startsWith("~$") && !/\.db$/i.test(entry.name) && !/\.tmp$/i.test(entry.name))
    .sort((a, b) => Number(b.isDirectory()) - Number(a.isDirectory()) || a.name.localeCompare(b.name, "de"));

  const parentHref = relPath
    ? path.posix.dirname(reqPath.endsWith("/") ? reqPath.slice(0, -1) : reqPath) + "/"
    : null;

  res.type("html").send(`<!doctype html>
<html lang="de">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>PCS Ordner</title>
    <style>
      :root { color-scheme: light; font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
      body { margin: 0; background: #eef2f4; color: #172027; }
      main { max-width: 1100px; margin: 0 auto; padding: 26px; }
      header { display: flex; justify-content: space-between; gap: 16px; align-items: end; margin-bottom: 16px; }
      h1 { margin: 0; font-size: 24px; }
      p { margin: 5px 0 0; color: #63707a; }
      a { color: #1f6f8b; font-weight: 750; text-decoration: none; }
      a:hover { text-decoration: underline; }
      button { border: 0; border-radius: 7px; padding: 9px 12px; background: #1f6f8b; color: white; font: inherit; font-weight: 800; cursor: pointer; }
      button:hover { background: #185b73; }
      .actions { display: flex; gap: 10px; align-items: center; }
      .panel { overflow: hidden; border: 1px solid #d6dde2; border-radius: 8px; background: white; box-shadow: 0 16px 34px rgba(20, 32, 42, 0.08); }
      .row { display: grid; grid-template-columns: 1fr 120px 180px; gap: 14px; align-items: center; padding: 11px 14px; border-bottom: 1px solid #d6dde2; }
      .row:last-child { border-bottom: 0; }
      .head { background: #f3f6f8; color: #63707a; font-size: 12px; font-weight: 800; text-transform: uppercase; }
      .type { justify-self: start; border-radius: 999px; padding: 3px 8px; background: #edf5f8; color: #1f6f8b; font-size: 12px; font-weight: 800; }
      .file { background: #e9eef2; color: #52606b; }
      time, .size { color: #63707a; font-size: 13px; }
      @media (max-width: 700px) { main { padding: 14px; } .row { grid-template-columns: 1fr; gap: 4px; } }
    </style>
  </head>
  <body>
    <main>
      <header>
        <div>
          <h1>PCS Ordner</h1>
          <p>${escapeHtml(reqPath)}</p>
        </div>
        <div class="actions">
          <button id="finder-btn" type="button">Im Finder öffnen</button>
          <a href="/">Dashboard</a>
        </div>
      </header>
      <section class="panel">
        <div class="row head"><span>Name</span><span>Typ</span><span>Geändert</span></div>
        ${parentHref ? `<div class="row"><a href="${escapeHtml(parentHref)}">..</a><span class="type">Ordner</span><span></span></div>` : ""}
        ${entries.map((entry) => {
          const full = path.join(target, entry.name);
          const entryStat = fs.statSync(full);
          const isDirectory = entry.isDirectory();
          const href = hrefJoin(reqPath, entry.name, isDirectory);
          return `<div class="row">
            <a href="${escapeHtml(href)}">${escapeHtml(entry.name)}</a>
            <span class="type ${isDirectory ? "" : "file"}">${isDirectory ? "Ordner" : "Datei"}</span>
            <time>${entryStat.mtime.toLocaleString("de-CH", { dateStyle: "short", timeStyle: "short" })}</time>
          </div>`;
        }).join("")}
      </section>
      <script>
        document.querySelector("#finder-btn").addEventListener("click", async () => {
          await fetch("/api/open-path", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ href: location.pathname })
          });
        });
      </script>
    </main>
  </body>
</html>`);
  return true;
}

function decodeHrefPath(href) {
  let value = String(href || "").split("?")[0];
  for (let i = 0; i < 5; i++) {
    try {
      const decoded = decodeURIComponent(value);
      if (decoded === value) break;
      value = decoded;
    } catch {
      break;
    }
  }
  return value;
}

function resolveAllowedHref(href) {
  if (!href || typeof href !== "string") return null;
  const rawPath = decodeHrefPath(href);
  const urlPath = rawPath.startsWith("/") ? rawPath : `/${rawPath}`;

  for (const [name, root] of evidenceRoots.entries()) {
    const prefix = `/${name}/`;
    if (!urlPath.startsWith(prefix)) continue;
    const rel = urlPath.slice(prefix.length);
    const safeRoot = fs.realpathSync(root);
    const target = path.resolve(safeRoot, rel);
    if (target.startsWith(safeRoot) && fs.existsSync(target)) return target;
  }

  if (urlPath.startsWith("/files/")) {
    const safeRoot = fs.realpathSync(PCS_ROOT);
    const rel = urlPath.slice("/files/".length);
    const target = path.resolve(safeRoot, rel);
    if (target.startsWith(safeRoot) && fs.existsSync(target)) return target;

    const parts = rel.split("/");
    const projectFolder = parts[0] || "";
    const projectMatch = projectFolder.match(/^EF[\s_-]*(\d+)$/i);
    if (projectMatch) {
      const fallback = path.resolve(safeRoot, projectMatch[1], ...parts.slice(1));
      if (fallback.startsWith(safeRoot) && fs.existsSync(fallback)) return fallback;
    }
  }

  return null;
}

function resolveAllowedParentHref(href) {
  const rawPath = decodeHrefPath(href);
  const urlPath = rawPath.startsWith("/") ? rawPath : `/${rawPath}`;
  const parentUrl = path.posix.dirname(urlPath.replace(/\/$/, ""));
  return resolveAllowedHref(parentUrl);
}

function invalidateListPathCache() {
  listPathCache.clear();
}

function copyRecursive(src, dst) {
  const stat = fs.statSync(src);
  if (stat.isDirectory()) {
    fs.mkdirSync(dst, { recursive: true });
    for (const entry of fs.readdirSync(src)) {
      copyRecursive(path.join(src, entry), path.join(dst, entry));
    }
  } else {
    fs.copyFileSync(src, dst);
  }
}

function uniquePath(target) {
  if (!fs.existsSync(target)) return target;
  const parsed = path.parse(target);
  let i = 2;
  while (true) {
    const candidate = path.join(parsed.dir, `${parsed.name} (${i})${parsed.ext}`);
    if (!fs.existsSync(candidate)) return candidate;
    i += 1;
  }
}

function openInSystemFileManager(target, appName) {
  let opener = process.platform === "darwin"
    ? "open"
    : process.platform === "win32"
      ? "explorer"
      : "xdg-open";
  let args = [target];

  if (appName === "word") {
    if (process.platform === "darwin") {
      args = ["-a", "Microsoft Word", target];
    } else if (process.platform === "win32") {
      opener = "cmd";
      args = ["/c", "start", "", "winword", target];
    }
  } else if (appName === "excel") {
    if (process.platform === "darwin") {
      args = ["-a", "Microsoft Excel", target];
    } else if (process.platform === "win32") {
      opener = "cmd";
      args = ["/c", "start", "", "excel", target];
    }
  }

  const child = spawn(opener, args, { detached: true, stdio: "ignore" });
  child.unref();
}

function formatFileSize(bytes) {
  if (!Number.isFinite(bytes) || bytes <= 0) return "";
  const units = ["B", "KB", "MB", "GB"];
  let size = bytes;
  let unit = 0;
  while (size >= 1024 && unit < units.length - 1) {
    size /= 1024;
    unit += 1;
  }
  return `${size >= 10 || unit === 0 ? Math.round(size) : size.toFixed(1)} ${units[unit]}`;
}

function directoryListingMiddleware(root, mountPath) {
  return (req, res, next) => {
    try {
      if (sendDirectoryListing(req, res, root, mountPath)) return;
    } catch (err) {
      return next(err);
    }
    next();
  };
}

// Resolve evidence-* symlinks to their real paths so Express can serve them
// (Express blocks symlinks that point outside the root directory by default)
try {
  for (const name of fs.readdirSync(__dirname)) {
    if (!name.startsWith("evidence-")) continue;
    const real = fs.realpathSync(path.join(__dirname, name));
    evidenceRoots.set(name, real);
    app.use(`/${name}`, directoryListingMiddleware(real, `/${name}`));
    app.use(`/${name}`, express.static(real, { dotfiles: "ignore", index: false }));
  }
} catch {}

app.use(express.static(path.join(__dirname)));

// Serve P:\PCS (or local equivalent) at /files/ so scanner hrefs work
app.use("/files", (req, _res, next) => {
  const safeRoot = fs.realpathSync(PCS_ROOT);
  const rel = decodeHrefPath(req.path).replace(/^\/+/, "");
  const target = path.resolve(safeRoot, rel);
  if (target.startsWith(safeRoot) && fs.existsSync(target)) return next();

  const parts = rel.split("/");
  const projectMatch = (parts[0] || "").match(/^EF[\s_-]*(\d+)$/i);
  if (projectMatch) {
    const fallbackParts = [projectMatch[1], ...parts.slice(1)];
    const fallback = path.resolve(safeRoot, ...fallbackParts);
    if (fallback.startsWith(safeRoot) && fs.existsSync(fallback)) {
      const query = req.url.includes("?") ? req.url.slice(req.url.indexOf("?")) : "";
      req.url = `/${fallbackParts.map(encodeURIComponent).join("/")}${query}`;
    }
  }
  next();
});
app.use("/files", directoryListingMiddleware(PCS_ROOT, "/files"));
app.use("/files", express.static(PCS_ROOT, { dotfiles: "ignore", index: false }));

// ── Projects ───────────────────────────────────────────────
app.get("/api/projects", (_req, res) => {
  res.json(getProjects());
});

app.get("/api/projects/:id", (req, res) => {
  const project = getProject(req.params.id);
  if (!project) return res.status(404).json({ error: "Nicht gefunden" });
  res.json(project);
});

// ── Ziffer status ──────────────────────────────────────────
app.put("/api/projects/:id/ziffern/:subtopic/:nr", (req, res) => {
  const { id, subtopic, nr } = req.params;
  const { status } = req.body;
  if (!status) return res.status(400).json({ error: "status fehlt" });
  updateZifferStatus(id, subtopic, nr, status);
  res.json({ ok: true });
});

app.put("/api/projects/:id/ziffern/:subtopic/:nr/reason", (req, res) => {
  const { id, subtopic, nr } = req.params;
  updateZifferReason(id, subtopic, nr, req.body?.reason);
  res.json({ ok: true });
});

// ── Task status ────────────────────────────────────────────
app.put("/api/projects/:id/tasks/:taskId", (req, res) => {
  const { id, taskId } = req.params;
  const { status, block_reason } = req.body;
  if (!status) return res.status(400).json({ error: "status fehlt" });
  updateTaskStatus(id, taskId, status, block_reason);
  res.json({ ok: true });
});

// ── Fachfreigabe ───────────────────────────────────────────
app.put("/api/projects/:id/fachfreigabe/gates", (req, res) => {
  const { id } = req.params;
  const { label, status } = req.body;
  if (!label || !status) return res.status(400).json({ error: "label und status benötigt" });
  updateFachfreigabeGate(id, label, status);
  res.json({ ok: true });
});

app.put("/api/projects/:id/fachfreigabe/meta", (req, res) => {
  const { id } = req.params;
  const { confirmed_by, datum, notiz } = req.body;
  updateFachfreigabeMeta(id, confirmed_by, datum, notiz);
  res.json({ ok: true });
});

// ── Packaging / Shipments ──────────────────────────────────
app.get("/api/projects/:id/shipments", (req, res) => {
  res.json(getShipments(req.params.id));
});
app.post("/api/projects/:id/shipments", (req, res) => {
  const newId = upsertShipment(req.params.id, req.body);
  res.json({ id: newId });
});
app.put("/api/projects/:id/shipments/:sid", (req, res) => {
  upsertShipment(req.params.id, { ...req.body, id: Number(req.params.sid) });
  res.json({ ok: true });
});
app.delete("/api/projects/:id/shipments/:sid", (req, res) => {
  deleteShipment(req.params.id, Number(req.params.sid));
  res.json({ ok: true });
});
app.get("/api/shipments/open", (_req, res) => {
  res.json(getAllOpenShipments());
});

// ── Labs ──────────────────────────────────────────────────
app.get("/api/labs", (_req, res) => res.json(getLabs()));
app.post("/api/labs", (req, res) => res.json({ id: upsertLab(req.body) }));
app.put("/api/labs/:id", (req, res) => {
  upsertLab({ ...req.body, id: Number(req.params.id) });
  res.json({ ok: true });
});

// ── Archiv ────────────────────────────────────────────────
app.put("/api/projects/:id/archive-location", (req, res) => {
  updateArchiveLocation(req.params.id, req.body?.location);
  res.json({ ok: true });
});

app.put("/api/projects/:id/project-no", (req, res) => {
  updateProjectNo(req.params.id, req.body?.project_no);
  res.json({ ok: true });
});

app.put("/api/projects/:id/sw-version", (req, res) => {
  updateSwVersion(req.params.id, req.body?.sw_version);
  res.json({ ok: true });
});

app.put("/api/projects/:id/hw-version", (req, res) => {
  updateHwVersion(req.params.id, req.body?.hw_version);
  res.json({ ok: true });
});

app.put("/api/projects/:id/machine-type", (req, res) => {
  updateMachineType(req.params.id, req.body?.machine_type);
  res.json({ ok: true });
});

app.put("/api/projects/:id/machine-use", (req, res) => {
  updateMachineUse(req.params.id, req.body?.machine_use);
  res.json({ ok: true });
});

app.put("/api/projects/:id/colors", (req, res) => {
  updateColors(req.params.id, req.body?.colors);
  res.json({ ok: true });
});

app.post("/api/open-folder", (req, res) => {
  const folderPath = req.body?.path;
  if (!folderPath) return res.status(400).json({ error: "no path" });
  const { exec } = require("child_process");
  exec(`open "${folderPath.replace(/"/g, '\\"')}"`, (err) => {
    if (err) return res.status(500).json({ error: err.message });
    res.json({ ok: true });
  });
});

const ARCHIVE_EXCEL_PATH = process.env.ARCHIVE_EXCEL || path.join(os.homedir(), "Desktop", "PCS_Archiv_Muster.xlsx");

function _xlsxProjectNo(xlsxPath) {
  try {
    const wb = XLSX.readFile(xlsxPath, { sheetRows: 15 });
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: null });
    for (const row of rows) {
      for (let i = 0; i < row.length; i++) {
        if (row[i] == null) continue;
        const label = String(row[i]).trim().toLowerCase();
        if (label.includes("project no") || label.includes("projektnr")) {
          const val = row[i + 4];
          if (val != null) return String(val).trim();
        }
      }
    }
  } catch (_) {}
  return null;
}

function _findApprobationXlsx(projectDir) {
  const results = [];
  function walk(dir) {
    let entries;
    try { entries = fs.readdirSync(dir, { withFileTypes: true }); } catch { return; }
    if (path.basename(dir).toLowerCase() === "approbationsauftrag") {
      for (const e of entries.sort((a, b) => a.name.localeCompare(b.name))) {
        if (e.isFile() && /\.xlsx$/i.test(e.name) && !e.name.startsWith("~$") && !e.name.startsWith("."))
          results.push(path.join(dir, e.name));
      }
      return;
    }
    for (const e of entries) {
      if (e.name.startsWith(".") || e.name.startsWith("~$")) continue;
      if (e.isDirectory()) walk(path.join(dir, e.name));
    }
  }
  walk(projectDir);
  return results;
}

function _projectIdFromFolder(name) {
  const m = name.match(/^EF[\s_-]*(\d+)/i) || name.match(/^(\d+)/);
  return m ? `EF${m[1]}` : null;
}

function runProjectNoScan(root) {
  return new Promise((resolve, reject) => {
    let entries;
    try {
      entries = fs.readdirSync(root, { withFileTypes: true })
        .filter(e => e.isDirectory() && !e.name.startsWith("."))
        .sort((a, b) => a.name.localeCompare(b.name));
    } catch (e) { return reject(new Error(e.message)); }
    let count = 0;
    for (const entry of entries) {
      const projectId = _projectIdFromFolder(entry.name);
      if (!projectId) continue;
      for (const xlsxPath of _findApprobationXlsx(path.join(root, entry.name))) {
        const no = _xlsxProjectNo(xlsxPath);
        if (no) { updateProjectNo(projectId, no); count++; break; }
      }
    }
    resolve(count);
  });
}

function runArchiveScan(excelPath) {
  return new Promise((resolve, reject) => {
    try {
      const wb = XLSX.readFile(excelPath);
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: null });
      const byEf = new Map();
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row || !row[0]) continue;
        const efNr = String(row[0]).trim();
        const loc = row[4] ? String(row[4]).trim() : "";
        if (!efNr || !loc || loc === "—") continue;
        if (!byEf.has(efNr)) byEf.set(efNr, []);
        const locs = byEf.get(efNr);
        if (!locs.includes(loc)) locs.push(loc);
      }
      let count = 0;
      for (const [ef, locs] of byEf) {
        updateArchiveLocation(ef, locs.join(" · "));
        count++;
      }
      resolve(count);
    } catch (e) { reject(new Error(`Archiv-Excel Fehler: ${e.message}`)); }
  });
}

app.post("/api/scan-archive", async (req, res) => {
  const excelPath = req.body?.path || ARCHIVE_EXCEL_PATH;
  try {
    const updated = await runArchiveScan(excelPath);
    res.json({ ok: true, updated });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post("/api/scan-project-no", async (_req, res) => {
  try {
    const updated = await runProjectNoScan(PCS_ROOT);
    res.json({ ok: true, updated });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get("/api/open-archive", (req, res) => {
  const projectId = (req.query.project_id || "").replace(/[^A-Za-z0-9]/g, "");
  const excelPath = ARCHIVE_EXCEL_PATH;

  // Find the row number for this project in the archive Excel
  let rowNum = null;
  try {
    const wb2 = XLSX.readFile(excelPath);
    const rows2 = XLSX.utils.sheet_to_json(wb2.Sheets[wb2.SheetNames[0]], { header: 1, defval: null });
    for (let i = 1; i < rows2.length; i++) {
      if (rows2[i] && rows2[i][0] != null && String(rows2[i][0]).trim() === projectId) {
        rowNum = i + 1;
        break;
      }
    }
  } catch (_) {}

  // Open the file with the default app (Numbers)
  spawn("open", [excelPath], { detached: true }).unref();

  // If found, navigate Numbers to that row after a short delay
  if (rowNum) {
    const cellRef = `A${rowNum}`;
    const appleScript = [
      'tell application "Numbers"',
      "    activate",
      "    delay 2",
      "    tell front document",
      "        tell active sheet",
      "            tell table 1",
      `                set selection range to cell "${cellRef}"`,
      "            end tell",
      "        end tell",
      "    end tell",
      "end tell",
    ].join("\n");
    setTimeout(() => {
      spawn("osascript", ["-e", appleScript], { detached: true }).unref();
    }, 500);
  }

  res.json({ ok: true });
});
app.get("/api/archive", (_req, res) => res.json(getArchiveItems()));
app.post("/api/archive", (req, res) => res.json({ id: upsertArchiveItem(req.body) }));
app.put("/api/archive/:id", (req, res) => {
  upsertArchiveItem({ ...req.body, id: Number(req.params.id) });
  res.json({ ok: true });
});
app.delete("/api/archive/:id", (req, res) => {
  deleteArchiveItem(Number(req.params.id));
  res.json({ ok: true });
});

// ── Scanner ────────────────────────────────────────────────
let scanInProgress = false;
let scanProgress = { done: 0, total: 0, current: "" };

// Start scan in background, return immediately
app.post("/api/scan/start", (_req, res) => {
  if (scanInProgress) return res.json({ started: false });
  scanInProgress = true;
  scanProgress = { done: 0, total: 0, current: "" };
  scan(null, (p) => { scanProgress = p; })
    .then(async (r) => {
      console.log(`Scan fertig: ${r.projects.length} Projekt(e)`);
      try {
        const n = await runProjectNoScan(PCS_ROOT);
        console.log(`Projektnummern: ${n} gescannt`);
      } catch (e) {
        console.warn(`Projektnr-Scan fehlgeschlagen: ${e.message}`);
      }
    })
    .catch((err) => console.error("Scan-Fehler:", err.message))
    .finally(() => { scanInProgress = false; });
  res.json({ started: true });
});

app.get("/api/scan/status", (_req, res) => {
  const last = db.prepare("SELECT * FROM scan_log ORDER BY id DESC LIMIT 1").get();
  res.json({
    ...(last || { scanned_at: null, projects_found: 0, files_found: 0 }),
    in_progress: scanInProgress,
    progress: scanInProgress ? scanProgress : null
  });
});

// ── Excel file scan across all document groups ──────────────
app.get("/api/projects/:id/excel-files", (req, res) => {
  const project = getProject(req.params.id);
  if (!project) return res.status(404).json({ error: "Projekt nicht gefunden" });

  const EXCEL_RE = /\.(xls|xlsx|xlsm|xlsb|xlam)$/i;
  const SKIP_RE = /^[.~]|\.tmp$/i;
  const results = [];

  function scan(dir, urlBase) {
    let entries;
    try { entries = fs.readdirSync(dir, { withFileTypes: true }); } catch { return; }
    for (const e of entries) {
      if (SKIP_RE.test(e.name)) continue;
      const fullPath = path.join(dir, e.name);
      const href = `${urlBase}${encodeURIComponent(e.name)}`;
      if (e.isDirectory()) {
        scan(fullPath, `${href}/`);
      } else if (EXCEL_RE.test(e.name)) {
        results.push({ name: e.name, href });
      }
    }
  }

  for (const group of project.documentGroups || []) {
    const resolved = resolveAllowedHref(group.href);
    if (!resolved) continue;
    const urlBase = group.href.endsWith("/") ? group.href : `${group.href}/`;
    scan(resolved, urlBase);
  }

  res.json(results);
});

// ── File preview (Word / Excel → HTML) ─────────────────────
// ── Local file manager opener ───────────────────────────────
app.post("/api/open-path", (req, res) => {
  const target = resolveAllowedHref(req.body?.href);
  if (!target) return res.status(400).json({ error: "Pfad nicht erlaubt oder nicht gefunden" });
  openInSystemFileManager(target, req.body?.app);
  res.json({ ok: true });
});

// Best-effort: ensure the user's write bit is set on `dir` before writing into it.
// Some EF project folders ship as read-only (dr-xr-xr-x); silently restore +w so file ops
// (mkdir / copy / move) work the same as in Finder. Failure is swallowed — the next
// fs op will throw with a real EACCES which we surface verbatim.
function ensureWritable(dir) {
  try {
    const st = fs.statSync(dir);
    if (!(st.mode & 0o200)) fs.chmodSync(dir, st.mode | 0o200);
  } catch {}
}

// Wrap fs errors so the client sees a useful message instead of generic 500.
function handleFsError(res, err, label) {
  const msg = err?.code === "EACCES"
    ? `${label}: Schreibrecht fehlt (${err.path || ""}). Ordner ist read-only — chmod +w wurde versucht und ist trotzdem fehlgeschlagen.`
    : `${label}: ${err?.message || String(err)}`;
  res.status(500).json({ error: msg });
}

app.post("/api/files/mkdir", (req, res) => {
  const parent = resolveAllowedHref(req.body?.parentHref);
  const name = String(req.body?.name || "").trim();
  if (!parent || !fs.existsSync(parent) || !fs.statSync(parent).isDirectory()) return res.status(400).json({ error: "Zielordner nicht erlaubt" });
  if (!name || /[\\/:*?"<>|]/.test(name)) return res.status(400).json({ error: "Ungültiger Ordnername" });
  ensureWritable(parent);
  const target = uniquePath(path.join(parent, name));
  try {
    fs.mkdirSync(target, { recursive: true });
    invalidateListPathCache();
    res.json({ ok: true, path: target });
  } catch (err) {
    handleFsError(res, err, "mkdir fehlgeschlagen");
  }
});

app.post("/api/files/copy", (req, res) => {
  const source = resolveAllowedHref(req.body?.sourceHref);
  const destDir = resolveAllowedHref(req.body?.destHref);
  if (!source || !destDir || !fs.statSync(destDir).isDirectory()) return res.status(400).json({ error: "Quelle oder Ziel nicht erlaubt" });
  ensureWritable(destDir);
  const target = uniquePath(path.join(destDir, path.basename(source)));
  try {
    copyRecursive(source, target);
    invalidateListPathCache();
    res.json({ ok: true, path: target });
  } catch (err) {
    handleFsError(res, err, "copy fehlgeschlagen");
  }
});

app.post("/api/files/move", (req, res) => {
  const source = resolveAllowedHref(req.body?.sourceHref);
  const destDir = resolveAllowedHref(req.body?.destHref);
  if (!source || !destDir || !fs.statSync(destDir).isDirectory()) return res.status(400).json({ error: "Quelle oder Ziel nicht erlaubt" });
  ensureWritable(destDir);
  ensureWritable(path.dirname(source));
  const target = uniquePath(path.join(destDir, path.basename(source)));
  try {
    try {
      fs.renameSync(source, target);
    } catch (err) {
      if (err.code !== "EXDEV") throw err;
      copyRecursive(source, target);
      fs.rmSync(source, { recursive: true, force: true });
    }
    invalidateListPathCache();
    res.json({ ok: true, path: target });
  } catch (err) {
    handleFsError(res, err, "move fehlgeschlagen");
  }
});

app.post("/api/files/delete", (req, res) => {
  const source = resolveAllowedHref(req.body?.href);
  if (!source) return res.status(400).json({ error: "Pfad nicht erlaubt" });
  ensureWritable(path.dirname(source));
  try {
    sendToTrash(source);
    invalidateListPathCache();
    res.json({ ok: true });
  } catch (err) {
    handleFsError(res, err, "delete fehlgeschlagen");
  }
});

const listPathCache = new Map();
const LIST_PATH_CACHE_MS = 60_000;

function clearListPathCache() {
  listPathCache.clear();
}

function uniquePathInDir(parent, name) {
  const parsed = path.parse(name);
  let candidate = path.join(parent, name);
  let counter = 2;
  while (fs.existsSync(candidate)) {
    candidate = path.join(parent, `${parsed.name} (${counter})${parsed.ext}`);
    counter += 1;
  }
  return candidate;
}

function sendToTrash(target) {
  if (process.platform === "darwin") {
    const script = `tell application "Finder" to delete POSIX file "${target.replace(/"/g, '\\"')}"`;
    const result = spawnSync("osascript", ["-e", script], { encoding: "utf8" });
    if (result.status === 0) return;
    throw new Error(result.stderr || "Finder trash failed");
  }

  if (process.platform === "win32") {
    const ps = [
      "Add-Type -AssemblyName Microsoft.VisualBasic",
      `$p = ${JSON.stringify(target)}`,
      "if (Test-Path -LiteralPath $p -PathType Container) {",
      "  [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($p, 'OnlyErrorDialogs', 'SendToRecycleBin')",
      "} else {",
      "  [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($p, 'OnlyErrorDialogs', 'SendToRecycleBin')",
      "}",
    ].join("; ");
    const result = spawnSync("powershell.exe", ["-NoProfile", "-Command", ps], { encoding: "utf8" });
    if (result.status === 0) return;
    throw new Error(result.stderr || "Recycle Bin failed");
  }

  const trashDir = path.join(path.dirname(target), "_Papierkorb");
  fs.mkdirSync(trashDir, { recursive: true });
  fs.renameSync(target, uniquePathInDir(trashDir, path.basename(target)));
}

app.get("/api/list-path", (req, res) => {
  const href = req.query?.href;
  const target = resolveAllowedHref(href);
  if (!target) return res.status(400).json({ error: "Pfad nicht erlaubt oder nicht gefunden" });

  const stat = fs.statSync(target);
  if (!stat.isDirectory()) return res.status(400).json({ error: "Pfad ist kein Ordner" });

  const urlPath = decodeHrefPath(href);
  const cacheKey = `${target}:${stat.mtimeMs}`;
  const cached = listPathCache.get(cacheKey);
  if (cached && Date.now() - cached.createdAt < LIST_PATH_CACHE_MS) return res.json(cached.payload);

  const entries = fs.readdirSync(target, { withFileTypes: true })
    .filter((entry) => !entry.name.startsWith(".") && !entry.name.startsWith("~$") && !/\.db$/i.test(entry.name) && !/\.tmp$/i.test(entry.name))
    .map((entry) => {
      const full = path.join(target, entry.name);
      const entryStat = fs.statSync(full);
      const isDirectory = entry.isDirectory();
      const HIDE = /^[.~]|\.tmp$|\.db$/i;
      const childCount = isDirectory
        ? (() => { try { return fs.readdirSync(full).filter(n => !HIDE.test(n)).length; } catch { return 0; } })()
        : 0;
      return {
        name: entry.name,
        type: isDirectory ? "Ordner" : "Datei",
        href: hrefJoin(urlPath, entry.name, isDirectory),
        modified: entryStat.mtime.toLocaleString("de-CH", { dateStyle: "short", timeStyle: "short" }),
        mtime: entryStat.mtime.getTime(),
        size: isDirectory ? "" : formatFileSize(entryStat.size),
        empty: childCount === 0,
        childCount: isDirectory ? childCount : null
      };
    })
    .sort((a, b) => Number(b.type === "Ordner") - Number(a.type === "Ordner") || b.mtime - a.mtime)
    .map(({ mtime, ...rest }) => rest);

  const payload = { href: urlPath, entries };
  listPathCache.set(cacheKey, { createdAt: Date.now(), payload });
  if (listPathCache.size > 200) {
    const firstKey = listPathCache.keys().next().value;
    listPathCache.delete(firstKey);
  }
  res.json(payload);
});

app.post("/api/file-action", (req, res) => {
  const { action, sourceHref, targetDirHref, name } = req.body || {};
  const targetDir = targetDirHref ? resolveAllowedHref(targetDirHref) : null;

  try {
    if (action === "mkdir") {
      if (!targetDir || !fs.statSync(targetDir).isDirectory()) return res.status(400).json({ error: "target folder invalid" });
      const safeName = String(name || "").trim().replace(/[\\/:*?"<>|]/g, " ");
      if (!safeName) return res.status(400).json({ error: "folder name missing" });
      const created = uniquePathInDir(targetDir, safeName);
      fs.mkdirSync(created, { recursive: false });
      clearListPathCache();
      return res.json({ ok: true, path: created });
    }

    const source = sourceHref ? resolveAllowedHref(sourceHref) : null;
    if (!source || !fs.existsSync(source)) return res.status(400).json({ error: "source invalid" });

    if (action === "delete") {
      sendToTrash(source);
      clearListPathCache();
      return res.json({ ok: true });
    }

    if (!targetDir || !fs.statSync(targetDir).isDirectory()) return res.status(400).json({ error: "target folder invalid" });
    const destination = uniquePathInDir(targetDir, path.basename(source));
    if (path.resolve(destination).startsWith(path.resolve(source) + path.sep)) {
      return res.status(400).json({ error: "cannot move/copy folder into itself" });
    }

    if (action === "copy") {
      fs.cpSync(source, destination, { recursive: true, errorOnExist: true });
      clearListPathCache();
      return res.json({ ok: true, path: destination });
    }

    if (action === "move") {
      try {
        fs.renameSync(source, destination);
      } catch (err) {
        if (err.code !== "EXDEV") throw err;
        fs.cpSync(source, destination, { recursive: true, errorOnExist: true });
        fs.rmSync(source, { recursive: true, force: true });
      }
      clearListPathCache();
      return res.json({ ok: true, path: destination });
    }

    res.status(400).json({ error: "unknown action" });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Tabelle 24 (component-certificate table scanner) ──────────────
const TABELLE24_SCRIPT = path.join(os.homedir(), "Desktop", "scripts-tabelle24", "parse_bauteilliste.py");

function runJsonProcess(command, args, stdinJson) {
  return new Promise((resolve, reject) => {
    const proc = spawn(command, args, { stdio: [stdinJson ? "pipe" : "ignore", "pipe", "pipe"] });
    let stdout = "";
    let stderr = "";
    proc.stdout.on("data", (d) => { stdout += d; });
    proc.stderr.on("data", (d) => { stderr += d; });
    proc.on("close", (code) => {
      if (code !== 0) return reject(new Error(stderr.trim() || `${command} exited ${code}`));
      try {
        resolve(JSON.parse(stdout));
      } catch (err) {
        reject(new Error(`non-JSON output: ${err.message}`));
      }
    });
    if (stdinJson) proc.stdin.end(JSON.stringify(stdinJson));
  });
}

function resolveTabelle24Root(query) {
  if (query.root && typeof query.root === "string") {
    const resolved = path.resolve(query.root);
    return fs.existsSync(resolved) ? resolved : null;
  }

  if (query.href && typeof query.href === "string") {
    return resolveAllowedHref(query.href);
  }

  const projectId = query.projectId;
  if (!projectId || typeof projectId !== "string") return null;

  const project = getProject(projectId);
  const bautelliste = project?.documentGroups?.find((group) =>
    group.area === "Bautelliste" || /^09\b/.test(group.primary || "")
  );
  if (bautelliste?.href) {
    const resolved = resolveAllowedHref(bautelliste.href);
    if (resolved) return resolved;
  }

  const numeric = projectId.replace(/^EF/i, "");
  const candidates = [
    path.join(PCS_ROOT, projectId),
    path.join(PCS_ROOT, numeric),
    path.join(PCS_ROOT, `EF ${numeric}`),
    path.join(os.homedir(), "Desktop", numeric),
    path.join(os.homedir(), "Desktop", projectId),
    path.join(os.homedir(), "Desktop", `EF ${numeric}`),
  ];
  return candidates.find((candidate) => fs.existsSync(candidate)) || null;
}

function normalizeTabelle24SearchRoot(root) {
  const base = path.basename(root).toLowerCase();
  if (/^09\b/.test(base) && base.includes("baut")) return root;

  const candidates = [
    path.join(root, "Zulassungen", "IEC"),
    path.join(root, "IEC"),
    root,
  ];
  for (const iecRoot of candidates) {
    let entries;
    try { entries = fs.readdirSync(iecRoot, { withFileTypes: true }); } catch { continue; }
    const folder = entries.find((entry) =>
      entry.isDirectory() && /^09\b/i.test(entry.name) && /baut.*liste/i.test(entry.name)
    );
    if (folder) return path.join(iecRoot, folder.name);
  }
  return root;
}

function deriveTabelle24TargetRoot(sourceFile) {
  const parts = path.resolve(sourceFile).split(path.sep);
  const iecIdx = parts.findIndex((part) => part === "IEC");
  if (iecIdx >= 0) {
    const iecRoot = parts.slice(0, iecIdx + 1).join(path.sep) || path.sep;
    const filePath = path.resolve(sourceFile);
    const is200vJp = /200v\s*jp/i.test(filePath);
    const target = path.join(iecRoot, "10 Komponenten", is200vJp ? "200V JP" : "");
    return fs.existsSync(target) ? target : path.join(iecRoot, "10 Komponenten");
  }
  return path.join(path.dirname(sourceFile), "..", "10 Komponenten");
}

function normalizePartNumber(raw) {
  if (!raw) return null;
  const clean = String(raw).replace(/\D/g, "");
  if (clean.length === 6) return `0${clean}`;
  return clean.length === 7 ? clean : null;
}

function extractPartNumbers(text) {
  const found = new Set();
  for (const match of String(text || "").matchAll(/(?:^|\D)(\d{6,7})(?!\d)/g)) {
    const normalized = normalizePartNumber(match[1]);
    if (normalized) found.add(normalized);
  }
  return [...found];
}

function extractCertificates(markText) {
  const text = String(markText || "");
  const certs = [];
  const add = (authority, number, raw) => {
    const key = `${authority}:${number || raw}`;
    if (!certs.some((cert) => cert.key === key)) certs.push({ key, authority, number, raw: raw.trim() });
  };

  for (const match of text.matchAll(/\bVDE\s+([0-9]{5,8})\b/gi)) add("VDE", match[1], match[0]);
  for (const match of text.matchAll(/\bESTI\s+([0-9]{2}\.[0-9]{4})\b/gi)) add("ESTI", match[1], match[0]);
  for (const match of text.matchAll(/\b(JET[0-9A-Z-]{8,})\b/gi)) add("JET", match[1], match[1]);
  for (const match of text.matchAll(/\bENEC\s*([0-9]{1,2})\b(?:[\s\S]{0,60}?([A-Z]{2,6}[- ][A-Z0-9-]{3,}|[0-9]{2,3}-[0-9]{3,}|V[0-9]{4}))?/gi)) {
    add("ENEC", match[2] || match[1], match[0]);
  }
  for (const match of text.matchAll(/\b(?:KEMA|DEKRA)\b[\s\S]{0,40}?([0-9]{2}-[0-9]{5,}|[0-9]{7}\.[0-9]{2})/gi)) add("DEKRA/KEMA", match[1], match[0]);
  for (const match of text.matchAll(/\bT[UÜ]V\b[\s\S]{0,35}?([A-Z0-9-]{6,})/gi)) add("TÜV", match[1], match[0]);
  if (/tested\s+with\s+appliance/i.test(text)) add("Appliance", null, "Tested with appliance");

  return certs.map(({ key, ...cert }) => cert);
}

function extractLatestIsoDateFromText(text) {
  const dates = [];
  for (const match of String(text || "").matchAll(/(20\d{2})[-.](\d{2})[-.](\d{2})|(\d{2})\.(\d{2})\.(20\d{2})/g)) {
    if (match[1]) dates.push(`${match[1]}-${match[2]}-${match[3]}`);
    else dates.push(`${match[6]}-${match[5]}-${match[4]}`);
  }
  dates.sort();
  return dates.at(-1) || null;
}

function listFilesDeep(dir, maxDepth = 2) {
  const files = [];
  function walk(current, depth) {
    if (depth > maxDepth) return;
    let entries;
    try { entries = fs.readdirSync(current, { withFileTypes: true }); } catch { return; }
    for (const entry of entries) {
      if (entry.name.startsWith(".") || entry.name === "Thumbs.db") continue;
      const full = path.join(current, entry.name);
      if (entry.isDirectory()) walk(full, depth + 1);
      else files.push(full);
    }
  }
  walk(dir, 0);
  return files;
}

function findChildFolderByPart(sectionRoot, partNumbers) {
  if (!sectionRoot || !fs.existsSync(sectionRoot)) return null;
  const wanted = new Set(partNumbers);
  const queue = [sectionRoot];
  while (queue.length) {
    const current = queue.shift();
    let entries;
    try { entries = fs.readdirSync(current, { withFileTypes: true }); } catch { continue; }
    for (const entry of entries) {
      if (!entry.isDirectory()) continue;
      const full = path.join(current, entry.name);
      const nums = extractPartNumbers(entry.name);
      if (nums.some((num) => wanted.has(num))) return full;
      queue.push(full);
    }
  }
  return null;
}

function normalizeNameForMatch(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/\b\d{1,7}\b/g, " ")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\b(co|ltd|gmbh|ag|inc|srl|spa|type|series)\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function findChildFolderByLabel(sectionRoot, label) {
  if (!sectionRoot || !fs.existsSync(sectionRoot)) return null;
  const labelNorm = normalizeNameForMatch(label);
  if (!labelNorm) return null;
  const labelTokens = new Set(labelNorm.split(" ").filter((token) => token.length >= 3));
  let best = null;
  let bestScore = 0;

  let entries;
  try { entries = fs.readdirSync(sectionRoot, { withFileTypes: true }); } catch { return null; }
  for (const entry of entries) {
    if (!entry.isDirectory()) continue;
    const nameNorm = normalizeNameForMatch(entry.name);
    if (!nameNorm) continue;
    const guardedTokens = ["x2", "y2", "f1", "f2", "hmi", "emc"];
    const guardMismatch = guardedTokens.some((token) =>
      labelNorm.split(" ").includes(token) && !nameNorm.split(" ").includes(token)
    );
    if (guardMismatch) continue;
    const nameTokens = nameNorm.split(" ").filter((token) => token.length >= 3);
    const overlap = nameTokens.filter((token) => labelTokens.has(token)).length;
    const score = overlap / Math.max(1, Math.min(nameTokens.length, labelTokens.size));
    if (score > bestScore) {
      bestScore = score;
      best = path.join(sectionRoot, entry.name);
    }
  }
  return bestScore >= 0.55 ? best : null;
}

function slugFromRow(row) {
  const cellText = (row.cells || []).join(" ").replace(/\s+/g, " ").trim();
  const parts = extractPartNumbers(cellText);
  const words = cellText
    .replace(/\b\d{6,7}\b/g, "")
    .replace(/[\\/:*?"<>|]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .split(" ")
    .slice(0, 8)
    .join(" ");
  return `${parts[0] ? `${parts[0]} ` : ""}${words || `Row ${row.rowIdx}`}`.trim();
}

function analyzeTabelle24Rows(parsed, targetRoot) {
  const analysis = [];
  let section = "Chassis";
  const sectionMap = {
    chassis: "Chassis",
    electronic: "Electronic",
    "coffee modul": "Coffee Modul",
    "coffee module": "Coffee Modul",
  };

  for (const row of parsed.rows || []) {
    const cellText = (row.cells || []).join("\n").trim();
    const mark = String(row.markOfConformity || "").trim();
    const combined = `${cellText}\n${mark}`;
    const title = (!cellText && mark) ? mark.toLowerCase() : "";

    if (title.includes("chassis")) { section = "Chassis"; continue; }
    if (title.includes("direct water")) { section = "Chassis"; continue; }
    if (title.includes("coffee modul")) { section = "Coffee Modul"; continue; }
    if (title.includes("mhbu")) { section = "MHBU"; continue; }
    if (/^electronic\b/i.test(cellText) && section === "Chassis") section = "Electronic";
    if (section === "MHBU") continue;

    const normalizedSection = sectionMap[section.toLowerCase()] || section;
    const partNumbers = extractPartNumbers(cellText);
    const certificates = extractCertificates(mark);
    const sectionRoot = path.join(targetRoot, normalizedSection);
    const label = slugFromRow(row);
    const matchedFolder = (partNumbers.length ? findChildFolderByPart(sectionRoot, partNumbers) : null)
      || findChildFolderByLabel(sectionRoot, label);
    const targetFolder = matchedFolder || path.join(sectionRoot, label);
    const exists = fs.existsSync(targetFolder);
    const files = exists ? listFilesDeep(targetFolder, 2) : [];
    const pdfs = files.filter((file) => /\.pdf$/i.test(file));
    const hasArchiv = exists && fs.existsSync(path.join(targetFolder, "Archiv"));

    analysis.push({
      rowIdx: row.rowIdx,
      groupId: row.groupId || null,
      section: normalizedSection,
      label,
      targetFolder,
      targetExists: exists,
      hasArchiv,
      partNumbers,
      certificates,
      markOfConformity: mark,
      currentDate: extractLatestIsoDateFromText(mark),
      pdfCount: pdfs.length,
      fileCount: files.length,
      oldPdfs: pdfs.map((file) => ({ name: path.basename(file), path: file })),
      actions: [
        ...(!exists ? ["create-folder"] : []),
        ...(!hasArchiv && pdfs.length ? ["create-archiv"] : []),
        ...(partNumbers.length ? ["m3-check"] : []),
        ...certificates.map((cert) => `${cert.authority.toLowerCase()}-check`),
      ],
    });
  }

  return analysis;
}

// List Intern Bauteilliste files in a project (under 09 Bautelliste / 09 Bauteilliste, incl. variant subfolders).
app.get("/api/tabelle24/files", (req, res) => {
  const resolved = resolveTabelle24Root(req.query);
  if (!resolved) return res.status(400).json({ error: "missing or unresolved ?root=, ?href=, or ?projectId=" });
  const searchRoot = normalizeTabelle24SearchRoot(resolved);

  const candidates = [];
  function walk(dir, depth) {
    if (depth > 4) return;
    let entries;
    try { entries = fs.readdirSync(dir, { withFileTypes: true }); } catch { return; }
    for (const e of entries) {
      const full = path.join(dir, e.name);
      if (e.isDirectory()) {
        if (/^archiv$/i.test(e.name)) continue;
        walk(full, depth + 1);
      } else if (/\.(docx?|docm)$/i.test(e.name) && /intern/i.test(e.name) && !/^~\$/.test(e.name)) {
        const stat = fs.statSync(full);
        candidates.push({ path: full, name: e.name, size: stat.size, mtime: stat.mtime.toISOString() });
      }
    }
  }
  walk(searchRoot, 0);
  candidates.sort((a, b) => b.mtime.localeCompare(a.mtime));
  res.json({ root: searchRoot, files: candidates });
});

// Look up a VDE certificate, return PDF list + extracted date + comparison verdict.
const VDE_LOOKUP_SCRIPT = path.join(os.homedir(), "Desktop", "scripts-tabelle24", "vde_lookup.py");
app.post("/api/tabelle24/vde-lookup", (req, res) => {
  const { certNumber, currentDate } = req.body || {};
  if (!certNumber || typeof certNumber !== "string") return res.status(400).json({ error: "missing certNumber" });
  if (!fs.existsSync(VDE_LOOKUP_SCRIPT)) return res.status(500).json({ error: `vde_lookup script missing: ${VDE_LOOKUP_SCRIPT}` });

  const proc = spawn("python3", [VDE_LOOKUP_SCRIPT], { stdio: ["pipe", "pipe", "pipe"] });
  let stdout = "", stderr = "";
  proc.stdout.on("data", (d) => { stdout += d; });
  proc.stderr.on("data", (d) => { stderr += d; });
  proc.on("close", (code) => {
    if (code !== 0) return res.status(500).json({ error: stderr.trim() || `lookup exited ${code}` });
    try { res.json(JSON.parse(stdout)); }
    catch (e) { res.status(500).json({ error: `bad JSON: ${e.message}`, raw: stdout.slice(0, 500) }); }
  });
  proc.stdin.end(JSON.stringify({ certNumber, currentDate }));
});

// Serve a file from anywhere under ~/Desktop/ (used by Tabelle 24 PDF thumbnails).
// Path is given as an absolute filesystem path; we resolve + verify it stays inside Desktop.
app.get("/api/tabelle24/file", (req, res) => {
  const p = req.query.path;
  if (!p || typeof p !== "string") return res.status(400).send("missing ?path");
  const safe = path.resolve(p);
  const desktop = path.resolve(os.homedir(), "Desktop");
  if (!safe.startsWith(desktop) || !fs.existsSync(safe)) return res.status(404).send("not found");
  const ext = path.extname(safe).toLowerCase();
  const types = { ".pdf": "application/pdf", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png", ".gif": "image/gif", ".webp": "image/webp" };
  res.setHeader("Content-Type", types[ext] || "application/octet-stream");
  res.setHeader("Content-Disposition", `inline; filename="${path.basename(safe)}"`);
  fs.createReadStream(safe).pipe(res);
});

// Serve a downloaded VDE PDF (from the lookup's temp dir) — so the UI can offer "PDF öffnen".
app.get("/api/tabelle24/vde-pdf", (req, res) => {
  const p = req.query.path;
  if (!p || typeof p !== "string") return res.status(400).send("missing ?path");
  const safe = path.resolve(p);
  // Only allow paths inside the OS temp dir's vde-lookups subdir (where vde_lookup.py writes).
  const allowed = path.resolve(os.tmpdir(), "vde-lookups");
  if (!safe.startsWith(allowed) || !fs.existsSync(safe)) return res.status(404).send("not found");
  res.setHeader("Content-Type", "application/pdf");
  fs.createReadStream(safe).pipe(res);
});

function uniqueDestination(dir, filename) {
  const parsed = path.parse(filename);
  let candidate = path.join(dir, filename);
  let counter = 2;
  while (fs.existsSync(candidate)) {
    candidate = path.join(dir, `${parsed.name}_${counter}${parsed.ext}`);
    counter += 1;
  }
  return candidate;
}

app.post("/api/tabelle24/vde-place", async (req, res) => {
  const { certNumber, currentDate, targetFolder } = req.body || {};
  if (!certNumber || typeof certNumber !== "string") return res.status(400).json({ error: "missing certNumber" });
  if (!targetFolder || typeof targetFolder !== "string") return res.status(400).json({ error: "missing targetFolder" });
  const target = path.resolve(targetFolder);
  if (!fs.existsSync(target)) return res.status(404).json({ error: `target folder not found: ${target}` });
  if (!fs.existsSync(VDE_LOOKUP_SCRIPT)) return res.status(500).json({ error: `vde_lookup script missing: ${VDE_LOOKUP_SCRIPT}` });

  try {
    const lookup = await runJsonProcess("python3", [VDE_LOOKUP_SCRIPT], { certNumber, currentDate });
    if (!lookup.downloadedPdf || !fs.existsSync(lookup.downloadedPdf)) {
      return res.json({ lookup, placed: false, archived: [], saved: null, message: "no downloaded PDF available" });
    }

    try { fs.chmodSync(target, fs.statSync(target).mode | 0o200); } catch {}
    const archiveDir = path.join(target, "Archiv");
    fs.mkdirSync(archiveDir, { recursive: true });
    try { fs.chmodSync(archiveDir, fs.statSync(archiveDir).mode | 0o200); } catch {}

    const archived = [];
    for (const entry of fs.readdirSync(target, { withFileTypes: true })) {
      if (!entry.isFile() || !/\.pdf$/i.test(entry.name)) continue;
      if (!entry.name.toLowerCase().includes("vde") && !entry.name.includes(certNumber)) continue;
      if (!entry.name.includes(certNumber)) continue;
      const src = path.join(target, entry.name);
      const dst = uniqueDestination(archiveDir, entry.name);
      fs.renameSync(src, dst);
      archived.push(dst);
    }

    const appendix = lookup.chosen?.label?.match(/\b(\d{2,3})\b/)?.[1] || "PDF";
    const datePart = lookup.onlineDate ? ` ${lookup.onlineDate}` : "";
    const filename = `VDE ${certNumber} Appendix_${appendix}${datePart}.pdf`;
    const savePath = uniqueDestination(target, filename);
    fs.copyFileSync(lookup.downloadedPdf, savePath);

    res.json({ lookup, placed: true, archived, saved: savePath });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Apply status changes to a Bauteilliste, write a new .docx, return its path.
const TABELLE24_UPDATE_SCRIPT = path.join(os.homedir(), "Desktop", "scripts-tabelle24", "update_bauteilliste.py");
app.post("/api/tabelle24/save", (req, res) => {
  const { sourceFile, outputFile, changes } = req.body || {};
  if (!sourceFile || typeof sourceFile !== "string") return res.status(400).json({ error: "missing sourceFile" });
  if (!Array.isArray(changes) || !changes.length) return res.status(400).json({ error: "missing changes[]" });
  if (!fs.existsSync(sourceFile)) return res.status(404).json({ error: `source not found: ${sourceFile}` });
  if (!fs.existsSync(TABELLE24_UPDATE_SCRIPT)) return res.status(500).json({ error: `update script missing: ${TABELLE24_UPDATE_SCRIPT}` });

  const proc = spawn("python3", [TABELLE24_UPDATE_SCRIPT], { stdio: ["pipe", "pipe", "pipe"] });
  let stdout = "", stderr = "";
  proc.stdout.on("data", (d) => { stdout += d; });
  proc.stderr.on("data", (d) => { stderr += d; });
  proc.on("close", (code) => {
    if (code !== 0) return res.status(500).json({ error: stderr.trim() || `updater exited ${code}` });
    try { res.json(JSON.parse(stdout)); }
    catch (e) { res.status(500).json({ error: `updater returned non-JSON: ${e.message}`, raw: stdout.slice(0, 500) }); }
  });
  proc.stdin.end(JSON.stringify({ sourceFile, outputFile, changes }));
});

// ── Tabelle 30 (VDE Typenprüfung — "Resistance to heat and fire") ─────────────
const TABELLE30_SCRIPT = path.join(os.homedir(), "Desktop", "scripts-tabelle24", "parse_tabelle30.py");

// List candidate Typenprüfung reports under a project (12 Untersuchungen folder, .doc/.docx
// with "typenpruf"/"vde" in the name).
app.get("/api/tabelle30/files", (req, res) => {
  const root = req.query.root;
  if (!root || typeof root !== "string") return res.status(400).json({ error: "missing ?root=<absolute path to EF project>" });
  const resolved = path.resolve(root);
  if (!fs.existsSync(resolved)) return res.status(404).json({ error: `not found: ${resolved}` });

  const candidates = [];
  function walk(dir, depth) {
    if (depth > 5) return;
    let entries;
    try { entries = fs.readdirSync(dir, { withFileTypes: true }); } catch { return; }
    for (const e of entries) {
      const full = path.join(dir, e.name);
      if (e.isDirectory()) {
        walk(full, depth + 1);
      } else if (/\.(docx?|docm)$/i.test(e.name) && !/^~\$/.test(e.name) && /(typenpr[uü]f|vde[_ ])/i.test(e.name)) {
        const stat = fs.statSync(full);
        candidates.push({ path: full, name: e.name, size: stat.size, mtime: stat.mtime.toISOString() });
      }
    }
  }
  walk(resolved, 0);
  candidates.sort((a, b) => b.mtime.localeCompare(a.mtime));
  res.json({ root: resolved, files: candidates });
});

// Parse Tabelle 30 from a Typenprüfung file.
app.post("/api/tabelle30/parse", (req, res) => {
  const file = req.body?.file;
  if (!file || typeof file !== "string") return res.status(400).json({ error: "missing { file }" });
  const resolved = path.resolve(file);
  if (!fs.existsSync(resolved)) return res.status(404).json({ error: `not found: ${resolved}` });
  if (!fs.existsSync(TABELLE30_SCRIPT)) return res.status(500).json({ error: `parser script missing: ${TABELLE30_SCRIPT}` });

  const proc = spawn("python3", [TABELLE30_SCRIPT, resolved], { stdio: ["ignore", "pipe", "pipe"] });
  let stdout = "", stderr = "";
  proc.stdout.on("data", (d) => { stdout += d; });
  proc.stderr.on("data", (d) => { stderr += d; });
  proc.on("close", (code) => {
    if (code !== 0) return res.status(500).json({ error: stderr.trim() || `parser exited ${code}` });
    try { res.json(JSON.parse(stdout)); }
    catch (e) { res.status(500).json({ error: `parser returned non-JSON: ${e.message}`, raw: stdout.slice(0, 500) }); }
  });
});

// Parse one Bauteilliste file → Tabelle 24 rows as JSON.
app.post("/api/tabelle24/parse", (req, res) => {
  const file = req.body?.file;
  if (!file || typeof file !== "string") return res.status(400).json({ error: "missing { file }" });
  const resolved = path.resolve(file);
  if (!fs.existsSync(resolved)) return res.status(404).json({ error: `not found: ${resolved}` });
  if (!fs.existsSync(TABELLE24_SCRIPT)) return res.status(500).json({ error: `parser script missing: ${TABELLE24_SCRIPT}` });

  const proc = spawn("python3", [TABELLE24_SCRIPT, resolved], { stdio: ["ignore", "pipe", "pipe"] });
  let stdout = "", stderr = "";
  proc.stdout.on("data", (d) => { stdout += d; });
  proc.stderr.on("data", (d) => { stderr += d; });
  proc.on("close", (code) => {
    if (code !== 0) return res.status(500).json({ error: stderr.trim() || `parser exited ${code}` });
    try { res.json(JSON.parse(stdout)); }
    catch (e) { res.status(500).json({ error: `parser returned non-JSON: ${e.message}`, raw: stdout.slice(0, 500) }); }
  });
});

// Build the automation worklist: row -> target folder -> M3/certificate tasks.
app.post("/api/tabelle24/analyze", async (req, res) => {
  const file = req.body?.file;
  if (!file || typeof file !== "string") return res.status(400).json({ error: "missing { file }" });
  const resolved = path.resolve(file);
  if (!fs.existsSync(resolved)) return res.status(404).json({ error: `not found: ${resolved}` });
  if (!fs.existsSync(TABELLE24_SCRIPT)) return res.status(500).json({ error: `parser script missing: ${TABELLE24_SCRIPT}` });

  try {
    const parsed = await runJsonProcess("python3", [TABELLE24_SCRIPT, resolved]);
    const targetRoot = req.body?.targetRoot
      ? path.resolve(req.body.targetRoot)
      : deriveTabelle24TargetRoot(resolved);
    const rows = analyzeTabelle24Rows(parsed, targetRoot);
    res.json({
      sourceFile: resolved,
      targetRoot,
      rowCount: rows.length,
      totals: {
        foldersMissing: rows.filter((row) => !row.targetExists).length,
        rowsWithM3: rows.filter((row) => row.partNumbers.length).length,
        rowsWithCertificates: rows.filter((row) => row.certificates.length).length,
        vde: rows.reduce((sum, row) => sum + row.certificates.filter((cert) => cert.authority === "VDE").length, 0),
        jet: rows.reduce((sum, row) => sum + row.certificates.filter((cert) => cert.authority === "JET").length, 0),
        enec: rows.reduce((sum, row) => sum + row.certificates.filter((cert) => cert.authority === "ENEC").length, 0),
      },
      rows,
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Start ──────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`PCS Dashboard → http://localhost:${PORT}`);
  console.log(`Datenbank: pcs.db  |  Dokumente: evidence-1157/`);
  runArchiveScan(ARCHIVE_EXCEL_PATH)
    .then((n) => console.log(`Archiv-Excel: ${n} Einträge synchronisiert`))
    .catch((e) => console.warn(`Archiv-Excel nicht geladen: ${e.message}`));
});
