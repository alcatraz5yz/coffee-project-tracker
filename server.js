const express = require("express");
const path = require("path");
const os = require("os");
const fs = require("fs");
const { spawn } = require("child_process");
const {
  getProjects,
  getProject,
  updateZifferStatus,
  updateFachfreigabeGate,
  updateFachfreigabeMeta,
  updateTaskStatus,
  updateArchiveLocation,
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
    .filter((entry) => !entry.name.startsWith(".") && !entry.name.startsWith("~$"))
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
      opener = "open";
      args = ["-a", "Microsoft Word", target];
    } else if (process.platform === "win32") {
      opener = "cmd";
      args = ["/c", "start", "", "winword", target];
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

const ARCHIVE_EXCEL_PATH = process.env.ARCHIVE_EXCEL || path.join(os.homedir(), "Desktop", "PCS_Archiv_Muster.xlsx");

function runArchiveScan(excelPath) {
  return new Promise((resolve, reject) => {
    const scriptPath = path.join(__dirname, "scan_archive.py");
    const proc = spawn("python3", [scriptPath], { env: { ...process.env, ARCHIVE_EXCEL: excelPath } });
    let out = "";
    let err = "";
    proc.stdout.on("data", (d) => { out += d; });
    proc.stderr.on("data", (d) => { err += d; });
    proc.on("close", (code) => {
      if (code !== 0) return reject(new Error(err || "Scanner-Fehler"));
      try {
        const entries = JSON.parse(out);
        if (entries.error) return reject(new Error(entries.error));
        entries.forEach(({ project_id, location }) => updateArchiveLocation(project_id, location));
        resolve(entries.length);
      } catch (e) {
        reject(e);
      }
    });
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

app.get("/api/open-archive", (req, res) => {
  const projectId = (req.query.project_id || "").replace(/[^A-Za-z0-9]/g, "");
  const excelPath = ARCHIVE_EXCEL_PATH;

  // Find the row number for this project via the scan_archive.py data
  let rowNum = null;
  try {
    const { execSync } = require("child_process");
    const pyCode = [
      "import openpyxl, os",
      `wb = openpyxl.load_workbook(r'${excelPath}')`,
      "ws = wb.active",
      "row_num = 0",
      "for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):",
      `    if str(row[0]).strip() == '${projectId}':`,
      "        row_num = i",
      "        break",
      "print(row_num)",
    ].join("\n");
    rowNum = parseInt(execSync(`python3 -c "${pyCode.replace(/"/g, '\\"')}"`).toString().trim(), 10) || null;
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
    .then((r) => console.log(`Scan fertig: ${r.projects.length} Projekt(e)`))
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
    .filter((entry) => !entry.name.startsWith(".") && !entry.name.startsWith("~$"))
    .sort((a, b) => Number(b.isDirectory()) - Number(a.isDirectory()) || a.name.localeCompare(b.name, "de"))
    .map((entry) => {
      const full = path.join(target, entry.name);
      const entryStat = fs.statSync(full);
      const isDirectory = entry.isDirectory();
      return {
        name: entry.name,
        type: isDirectory ? "Ordner" : "Datei",
        href: hrefJoin(urlPath, entry.name, isDirectory),
        modified: entryStat.mtime.toLocaleString("de-CH", { dateStyle: "short", timeStyle: "short" }),
        size: isDirectory ? "" : formatFileSize(entryStat.size)
      };
    });

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
