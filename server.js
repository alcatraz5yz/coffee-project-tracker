const express = require("express");
const path = require("path");
const os = require("os");
const fs = require("fs");
const { spawn } = require("child_process");
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

function resolveAllowedHref(href) {
  if (!href || typeof href !== "string") return null;
  const rawPath = decodeURIComponent(href.split("?")[0]);
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
    const target = path.resolve(safeRoot, urlPath.slice("/files/".length));
    if (target.startsWith(safeRoot) && fs.existsSync(target)) return target;
  }

  return null;
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
app.get("/api/preview-file", async (req, res) => {
  const target = resolveAllowedHref(req.query?.href);
  if (!target) return res.status(400).send("Pfad nicht erlaubt");
  const ext = path.extname(target).toLowerCase();

  if (/\.(xlsx|xls|xlsm|xlsb)$/i.test(target)) {
    try {
      const wb = XLSX.readFile(target);
      const sheets = wb.SheetNames.map((name) => ({
        name,
        html: XLSX.utils.sheet_to_html(wb.Sheets[name])
      }));
      res.type("html").send(`<!doctype html><html><head><meta charset="utf-8"><style>
        *{box-sizing:border-box}body{margin:0;font-family:system-ui,sans-serif;font-size:13px;background:#fff}
        .tabs{display:flex;gap:2px;padding:6px 8px;background:#f3f6f8;border-bottom:1px solid #d6dde2;overflow-x:auto;flex-shrink:0}
        .tab{padding:3px 10px;border:1px solid #c8d0d8;background:#fff;border-radius:4px;cursor:pointer;font-size:12px;font-weight:700;white-space:nowrap}
        .tab.active{background:#1f6f8b;color:#fff;border-color:#1f6f8b}
        .sheet{display:none;overflow:auto;padding:10px}
        .sheet.active{display:block}
        table{border-collapse:collapse;font-size:12px}
        td,th{border:1px solid #d6dde2;padding:3px 8px;white-space:nowrap;vertical-align:top}
        tr:first-child td,tr:first-child th{background:#f3f6f8;font-weight:700}
      </style></head><body>
      <div class="tabs">${sheets.map((s, i) => `<button class="tab${i === 0 ? " active" : ""}" onclick="show(${i})">${s.name}</button>`).join("")}</div>
      ${sheets.map((s, i) => `<div class="sheet${i === 0 ? " active" : ""}" id="s${i}">${s.html}</div>`).join("")}
      <script>function show(i){document.querySelectorAll(".tab").forEach((b,j)=>b.classList.toggle("active",i===j));document.querySelectorAll(".sheet").forEach((c,j)=>c.classList.toggle("active",i===j));}</script>
      </body></html>`);
    } catch (e) { res.status(500).send(`<p style="padding:16px;color:red">Fehler: ${e.message}</p>`); }

  } else if (/\.docx$/i.test(target)) {
    try {
      const mammoth = require("mammoth");
      const result = await mammoth.convertToHtml({ path: target });
      res.type("html").send(`<!doctype html><html><head><meta charset="utf-8"><style>
        body{margin:24px 28px;font-family:system-ui,sans-serif;font-size:14px;line-height:1.65;color:#172027;max-width:860px}
        img{max-width:100%}table{border-collapse:collapse;width:100%;margin:8px 0}
        td,th{border:1px solid #d6dde2;padding:5px 8px}th{background:#f3f6f8;font-weight:700}
        h1,h2,h3{margin:1em 0 0.4em}p{margin:0.4em 0}
      </style></head><body>${result.value}</body></html>`);
    } catch (e) { res.status(500).send(`<p style="padding:16px;color:red">Fehler: ${e.message}</p>`); }

  } else {
    res.status(400).send("Dateityp nicht unterstützt");
  }
});

// ── Local file manager opener ───────────────────────────────
app.post("/api/open-path", (req, res) => {
  const target = resolveAllowedHref(req.body?.href);
  if (!target) return res.status(400).json({ error: "Pfad nicht erlaubt oder nicht gefunden" });
  openInSystemFileManager(target, req.body?.app);
  res.json({ ok: true });
});

app.get("/api/list-path", (req, res) => {
  const href = req.query?.href;
  const target = resolveAllowedHref(href);
  if (!target) return res.status(400).json({ error: "Pfad nicht erlaubt oder nicht gefunden" });

  const stat = fs.statSync(target);
  if (!stat.isDirectory()) return res.status(400).json({ error: "Pfad ist kein Ordner" });

  const urlPath = decodeURIComponent(String(href).split("?")[0]);
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

  res.json({ href: urlPath, entries });
});

// ── Start ──────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`PCS Dashboard → http://localhost:${PORT}`);
  console.log(`Datenbank: pcs.db  |  Dokumente: evidence-1157/`);
  runArchiveScan(ARCHIVE_EXCEL_PATH)
    .then((n) => console.log(`Archiv-Excel: ${n} Einträge synchronisiert`))
    .catch((e) => console.warn(`Archiv-Excel nicht geladen: ${e.message}`));
});
