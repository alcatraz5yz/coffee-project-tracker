const express = require("express");
const path = require("path");
const os = require("os");
const fs = require("fs");
const fsp = fs.promises;
const { spawn, spawnSync } = require("child_process");
const XLSX = require("xlsx");
const { trackerConfig } = require("./config");
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
const { classifyNewEntries } = require("./tabelle30-match");
const { parseTabelle30, parseExcelErgaenzung } = require("./tabelle30-node");
const { updateTabelle30 } = require("./tabelle30-update-node");
const { readDocumentXml } = require("./docx-reader");
const { fileContainsTabelle24, parseTabelle24, updateTabelle24 } = require("./tabelle24-node");
const { lookupVde } = require("./vde-lookup-node");
const { assertExcelReadable } = require("./excel-safety");

const app = express();
const PORT = process.env.PORT || 8090;
const HOST = process.env.HOST || "127.0.0.1";
const FS_TIMEOUT_MS = Number(process.env.FS_TIMEOUT_MS || 6000);
const evidenceRoots = new Map();

app.disable("x-powered-by");
app.use((_req, res, next) => {
  res.setHeader("X-Content-Type-Options", "nosniff");
  res.setHeader("Referrer-Policy", "same-origin");
  res.setHeader("X-Frame-Options", "SAMEORIGIN");
  next();
});
app.use(express.json({ limit: "2mb" }));

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

// True only when `target` is `root` itself or a path genuinely inside it.
// A plain startsWith(root) check would also accept sibling folders that merely
// share a name prefix (e.g. "evidence-1157-backup" for root "evidence-1157"),
// which would let crafted ../ hrefs escape the allowed root.
function isWithin(root, target) {
  return target === root || target.startsWith(root + path.sep);
}

class FsTimeoutError extends Error {
  constructor(label) {
    super(`${label}: Netzlaufwerk antwortet nicht innerhalb von ${FS_TIMEOUT_MS} ms`);
    this.name = "FsTimeoutError";
    this.status = 504;
  }
}

function fsTimeout(promise, label) {
  let timer;
  const timeout = new Promise((_, reject) => {
    timer = setTimeout(() => reject(new FsTimeoutError(label)), FS_TIMEOUT_MS);
  });
  return Promise.race([promise, timeout]).finally(() => clearTimeout(timer));
}

async function fsRealpath(p, label = `realpath ${p}`) {
  return fsTimeout(fsp.realpath(p), label);
}

async function fsStat(p, label = `stat ${p}`) {
  return fsTimeout(fsp.stat(p), label);
}

async function fsAccess(p, label = `access ${p}`) {
  try {
    await fsTimeout(fsp.access(p), label);
    return true;
  } catch (err) {
    if (err?.code === "ENOENT" || err?.code === "ENOTDIR") return false;
    throw err;
  }
}

async function fsReaddir(p, options, label = `readdir ${p}`) {
  return fsTimeout(fsp.readdir(p, options), label);
}

async function fsIsDirectory(p) {
  try {
    return (await fsStat(p)).isDirectory();
  } catch {
    return false;
  }
}

async function fsExists(p) {
  return fsAccess(p);
}

const realpathCache = new Map();

async function cachedRealpath(root) {
  if (realpathCache.has(root)) return realpathCache.get(root);
  const real = await fsRealpath(root, `realpath ${root}`);
  realpathCache.set(root, real);
  return real;
}

function boundedCacheGet(cache, key, producer, maxEntries = 250) {
  if (cache.has(key)) return cache.get(key);
  const value = producer();
  cache.set(key, value);
  if (cache.size > maxEntries) cache.delete(cache.keys().next().value);
  return value;
}

function fileCacheKey(filePath) {
  const stat = fs.statSync(filePath);
  return `${path.resolve(filePath)}\0${stat.size}\0${stat.mtimeMs}`;
}

const markerCache = new Map();
const parsedTabelle24Cache = new Map();
const parsedTabelle30Cache = new Map();
const parsedExcelCache = new Map();

function cachedMarker(filePath, markerName, producer) {
  return boundedCacheGet(markerCache, `${markerName}\0${fileCacheKey(filePath)}`, producer, 500);
}

function cachedParse(cache, filePath, producer) {
  return boundedCacheGet(cache, fileCacheKey(filePath), producer, 120);
}

async function sendDirectoryListing(req, res, root, mountPath) {
  const mount = mountPath.endsWith("/") ? mountPath.slice(0, -1) : mountPath;
  const reqPath = decodeURIComponent(req.originalUrl.split("?")[0]);
  const relUrl = reqPath.startsWith(mount) ? reqPath.slice(mount.length) : reqPath;
  const relPath = relUrl.replace(/^\/+/, "");
  const safeRoot = await cachedRealpath(root);
  const target = path.resolve(safeRoot, relPath);

  if (!isWithin(safeRoot, target)) return res.status(403).send("Forbidden");
  if (!await fsExists(target)) return false;

  const stat = await fsStat(target);
  if (!stat.isDirectory()) return false;

  const entries = (await fsReaddir(target, { withFileTypes: true }))
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
        ${(await Promise.all(entries.map(async (entry) => {
          const full = path.join(target, entry.name);
          const entryStat = await fsStat(full);
          const isDirectory = entry.isDirectory();
          const href = hrefJoin(reqPath, entry.name, isDirectory);
          return `<div class="row">
            <a href="${escapeHtml(href)}">${escapeHtml(entry.name)}</a>
            <span class="type ${isDirectory ? "" : "file"}">${isDirectory ? "Ordner" : "Datei"}</span>
            <time>${entryStat.mtime.toLocaleString("de-CH", { dateStyle: "short", timeStyle: "short" })}</time>
          </div>`;
        }))).join("")}
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

async function resolveAllowedHref(href) {
  if (!href || typeof href !== "string") return null;
  const rawPath = decodeHrefPath(href);
  const urlPath = rawPath.startsWith("/") ? rawPath : `/${rawPath}`;

  for (const [name, root] of evidenceRoots.entries()) {
    const prefix = `/${name}/`;
    if (!urlPath.startsWith(prefix)) continue;
    const rel = urlPath.slice(prefix.length);
    const safeRoot = await cachedRealpath(root);
    const target = path.resolve(safeRoot, rel);
    if (isWithin(safeRoot, target) && await fsExists(target)) return target;
  }

  if (urlPath.startsWith("/files/")) {
    const safeRoot = await cachedRealpath(PCS_ROOT);
    const rel = urlPath.slice("/files/".length);
    const target = path.resolve(safeRoot, rel);
    if (isWithin(safeRoot, target) && await fsExists(target)) return target;

    const parts = rel.split("/");
    const projectFolder = parts[0] || "";
    const projectMatch = projectFolder.match(/^EF[\s_-]*(\d+)$/i);
    if (projectMatch) {
      const fallback = path.resolve(safeRoot, projectMatch[1], ...parts.slice(1));
      if (isWithin(safeRoot, fallback) && await fsExists(fallback)) return fallback;
    }
  }

  return null;
}

function resolveAllowedHrefSync(href) {
  if (!href || typeof href !== "string") return null;
  const rawPath = decodeHrefPath(href);
  const urlPath = rawPath.startsWith("/") ? rawPath : `/${rawPath}`;

  for (const [name, root] of evidenceRoots.entries()) {
    const prefix = `/${name}/`;
    if (!urlPath.startsWith(prefix)) continue;
    const rel = urlPath.slice(prefix.length);
    const safeRoot = fs.realpathSync(root);
    const target = path.resolve(safeRoot, rel);
    if (isWithin(safeRoot, target) && fs.existsSync(target)) return target;
  }

  if (urlPath.startsWith("/files/")) {
    const safeRoot = fs.realpathSync(PCS_ROOT);
    const rel = urlPath.slice("/files/".length);
    const target = path.resolve(safeRoot, rel);
    if (isWithin(safeRoot, target) && fs.existsSync(target)) return target;

    const parts = rel.split("/");
    const projectFolder = parts[0] || "";
    const projectMatch = projectFolder.match(/^EF[\s_-]*(\d+)$/i);
    if (projectMatch) {
      const fallback = path.resolve(safeRoot, projectMatch[1], ...parts.slice(1));
      if (isWithin(safeRoot, fallback) && fs.existsSync(fallback)) return fallback;
    }
  }

  return null;
}

async function resolveAllowedParentHref(href) {
  const rawPath = decodeHrefPath(href);
  const urlPath = rawPath.startsWith("/") ? rawPath : `/${rawPath}`;
  const parentUrl = path.posix.dirname(urlPath.replace(/\/$/, ""));
  return resolveAllowedHref(parentUrl);
}

function invalidateListPathCache() {
  listPathCache.clear();
}

async function copyRecursiveAsync(src, dst) {
  await fsTimeout(fsp.cp(src, dst, { recursive: true, errorOnExist: true }), `copy ${src}`);
}

async function uniquePathAsync(target) {
  if (!await fsExists(target)) return target;
  const parsed = path.parse(target);
  let i = 2;
  while (true) {
    const candidate = path.join(parsed.dir, `${parsed.name} (${i})${parsed.ext}`);
    if (!await fsExists(candidate)) return candidate;
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
  return async (req, res, next) => {
    try {
      if (await sendDirectoryListing(req, res, root, mountPath)) return;
    } catch (err) {
      return next(err);
    }
    next();
  };
}

function denySensitiveFiles(req, res, next) {
  const name = path.basename(decodeHrefPath(req.path));
  if (name.startsWith(".") || name.startsWith("~$") || /\.(db|db-shm|db-wal|sqlite|tmp)$/i.test(name)) {
    return res.status(404).send("not found");
  }
  next();
}

const STATIC_FILES = new Set([
  "index.html",
  "app.js",
  "styles.css",
  "config.js",
  "tabelle24.html",
  "tabelle30.html",
]);

app.get("/", (_req, res) => {
  res.sendFile(path.join(__dirname, "index.html"));
});

for (const file of STATIC_FILES) {
  app.get(`/${file}`, (_req, res) => {
    res.sendFile(path.join(__dirname, file));
  });
}

for (const dir of ["assets", "public"]) {
  const full = path.join(__dirname, dir);
  if (fs.existsSync(full)) {
    app.use(`/${dir}`, express.static(full, { dotfiles: "ignore", index: false, fallthrough: false }));
  }
}

// Resolve evidence-* symlinks to their real paths so Express can serve them
// (Express blocks symlinks that point outside the root directory by default)
try {
  for (const name of fs.readdirSync(__dirname)) {
    if (!name.startsWith("evidence-")) continue;
    const real = fs.realpathSync(path.join(__dirname, name));
    evidenceRoots.set(name, real);
    app.use(`/${name}`, directoryListingMiddleware(real, `/${name}`));
    app.use(`/${name}`, denySensitiveFiles);
    app.use(`/${name}`, express.static(real, { dotfiles: "ignore", index: false }));
  }
} catch {}

// Serve P:\PCS (or local equivalent) at /files/ so scanner hrefs work
app.use("/files", async (req, _res, next) => {
  try {
    const safeRoot = await cachedRealpath(PCS_ROOT);
    const rel = decodeHrefPath(req.path).replace(/^\/+/, "");
    const target = path.resolve(safeRoot, rel);
    if (isWithin(safeRoot, target) && await fsExists(target)) return next();

    const parts = rel.split("/");
    const projectMatch = (parts[0] || "").match(/^EF[\s_-]*(\d+)$/i);
    if (projectMatch) {
      const fallbackParts = [projectMatch[1], ...parts.slice(1)];
      const fallback = path.resolve(safeRoot, ...fallbackParts);
      if (isWithin(safeRoot, fallback) && await fsExists(fallback)) {
        const query = req.url.includes("?") ? req.url.slice(req.url.indexOf("?")) : "";
        req.url = `/${fallbackParts.map(encodeURIComponent).join("/")}${query}`;
      }
    }
    next();
  } catch (err) {
    next(err);
  }
});
app.use("/files", directoryListingMiddleware(PCS_ROOT, "/files"));
app.use("/files", denySensitiveFiles);
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
  try {
    openInSystemFileManager(folderPath);   // plattformübergreifend (Finder/Explorer)
    res.json({ ok: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

const CONFIG_ARCHIVE_EXCEL = process.platform === "win32" ? trackerConfig.archiveExcel : null;
const ARCHIVE_EXCEL_PATH = process.env.ARCHIVE_EXCEL || CONFIG_ARCHIVE_EXCEL || path.join(os.homedir(), "Desktop", "PCS_Archiv_Muster.xlsx");

function _xlsxProjectNo(xlsxPath) {
  try {
    assertExcelReadable(xlsxPath);
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

async function _findApprobationXlsx(projectDir) {
  const results = [];
  async function walk(dir) {
    let entries;
    try { entries = await fsp.readdir(dir, { withFileTypes: true }); } catch { return; }
    if (path.basename(dir).toLowerCase() === "approbationsauftrag") {
      for (const e of entries.sort((a, b) => a.name.localeCompare(b.name))) {
        if (e.isFile() && /\.xlsx$/i.test(e.name) && !e.name.startsWith("~$") && !e.name.startsWith("."))
          results.push(path.join(dir, e.name));
      }
      return;
    }
    for (const e of entries) {
      if (e.name.startsWith(".") || e.name.startsWith("~$")) continue;
      if (e.isDirectory()) await walk(path.join(dir, e.name));
    }
  }
  await walk(projectDir);
  return results;
}

function _projectIdFromFolder(name) {
  const m = name.match(/^EF[\s_-]*(\d+)/i) || name.match(/^(\d+)/);
  return m ? `EF${m[1]}` : null;
}

async function runProjectNoScan(root) {
  let entries;
  try {
    entries = (await fsp.readdir(root, { withFileTypes: true }))
      .filter(e => e.isDirectory() && !e.name.startsWith("."))
      .sort((a, b) => a.name.localeCompare(b.name));
  } catch (e) { throw new Error(e.message); }
  let count = 0;
  for (const entry of entries) {
    const projectId = _projectIdFromFolder(entry.name);
    if (!projectId) continue;
    for (const xlsxPath of await _findApprobationXlsx(path.join(root, entry.name))) {
      const no = _xlsxProjectNo(xlsxPath);
      if (no) { updateProjectNo(projectId, no); count++; break; }
    }
  }
  return count;
}

function runArchiveScan(excelPath) {
  return new Promise((resolve, reject) => {
    try {
      assertExcelReadable(excelPath);
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
    assertExcelReadable(excelPath);
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
app.post("/api/scan/start", (req, res) => {
  if (scanInProgress) return res.json({ started: false });
  // projectId gesetzt → nur dieses eine Projekt scannen (z.B. das geöffnete),
  // sonst alle. Spart auf dem Netzlaufwerk viel Zeit.
  const onlyProjectId = String(req.body?.projectId || "").trim() || null;
  scanInProgress = true;
  scanProgress = { done: 0, total: 0, current: "" };
  scan(null, (p) => { scanProgress = p; }, { onlyProjectId })
    .then(async (r) => {
      console.log(`Scan fertig: ${r.projects.length} Projekt(e)${onlyProjectId ? ` (nur ${onlyProjectId})` : ""}`);
      // Projektnummern-Abgleich (liest eine Excel) nur beim Voll-Scan.
      if (!onlyProjectId) {
        try {
          const n = await runProjectNoScan(PCS_ROOT);
          console.log(`Projektnummern: ${n} gescannt`);
        } catch (e) {
          console.warn(`Projektnr-Scan fehlgeschlagen: ${e.message}`);
        }
      }
    })
    .catch((err) => console.error("Scan-Fehler:", err.message))
    .finally(() => { scanInProgress = false; });
  res.json({ started: true, scope: onlyProjectId || "all" });
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
app.get("/api/projects/:id/excel-files", async (req, res, next) => {
  const project = getProject(req.params.id);
  if (!project) return res.status(404).json({ error: "Projekt nicht gefunden" });

  const EXCEL_RE = /\.(xls|xlsx|xlsm|xlsb|xlam)$/i;
  const SKIP_RE = /^[.~]|\.tmp$/i;
  const results = [];

  async function scan(dir, urlBase) {
    let entries;
    try { entries = await fsReaddir(dir, { withFileTypes: true }); } catch { return; }
    for (const e of entries) {
      if (SKIP_RE.test(e.name)) continue;
      const fullPath = path.join(dir, e.name);
      const href = `${urlBase}${encodeURIComponent(e.name)}`;
      if (e.isDirectory()) {
        await scan(fullPath, `${href}/`);
      } else if (EXCEL_RE.test(e.name)) {
        results.push({ name: e.name, href });
      }
    }
  }

  try {
    for (const group of project.documentGroups || []) {
      const resolved = await resolveAllowedHref(group.href);
      if (!resolved) continue;
      const urlBase = group.href.endsWith("/") ? group.href : `${group.href}/`;
      await scan(resolved, urlBase);
    }
    res.json(results);
  } catch (err) {
    next(err);
  }
});

// ── File preview (Word / Excel → HTML) ─────────────────────
// ── Local file manager opener ───────────────────────────────
app.post("/api/open-path", async (req, res, next) => {
  try {
    const target = await resolveAllowedHref(req.body?.href);
    if (!target) return res.status(400).json({ error: "Pfad nicht erlaubt oder nicht gefunden" });
    openInSystemFileManager(target, req.body?.app);
    res.json({ ok: true });
  } catch (err) {
    next(err);
  }
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

function validateEntryName(name) {
  const value = String(name || "").trim();
  if (!value) return null;
  if (value === "." || value === "..") return null;
  if (/[\\/:*?"<>|\x00-\x1f]/.test(value)) return null;
  if (value !== path.basename(value)) return null;
  return value;
}

app.post("/api/files/mkdir", async (req, res, next) => {
  try {
    const parent = await resolveAllowedHref(req.body?.parentHref);
    const name = validateEntryName(req.body?.name);
    if (!parent || !await fsIsDirectory(parent)) return res.status(400).json({ error: "Zielordner nicht erlaubt" });
    if (!name) return res.status(400).json({ error: "Ungültiger Ordnername" });
    ensureWritable(parent);
    const target = await uniquePathAsync(path.join(parent, name));
    fs.mkdirSync(target, { recursive: true });
    invalidateListPathCache();
    res.json({ ok: true, path: target });
  } catch (err) {
    if (err instanceof FsTimeoutError) return next(err);
    handleFsError(res, err, "mkdir fehlgeschlagen");
  }
});

app.post("/api/files/rename", async (req, res, next) => {
  try {
    const source = await resolveAllowedHref(req.body?.href);
    const newName = validateEntryName(req.body?.newName);
    if (!source || !await fsExists(source)) return res.status(400).json({ error: "Pfad nicht erlaubt" });
    if (!newName) return res.status(400).json({ error: "Ungültiger Name" });

    const parent = path.dirname(source);
    const target = path.join(parent, newName);
    if (path.resolve(source) === path.resolve(target)) return res.json({ ok: true, path: source });
    if (await fsExists(target)) return res.status(409).json({ error: "Ein Element mit diesem Namen existiert bereits" });

    ensureWritable(parent);
    await fsTimeout(fsp.rename(source, target), `rename ${source}`);
    invalidateListPathCache();
    res.json({ ok: true, path: target });
  } catch (err) {
    if (err instanceof FsTimeoutError) return next(err);
    handleFsError(res, err, "rename fehlgeschlagen");
  }
});

app.post("/api/files/copy", async (req, res, next) => {
  try {
    const source = await resolveAllowedHref(req.body?.sourceHref);
    const destDir = await resolveAllowedHref(req.body?.destHref);
    if (!source || !destDir || !await fsIsDirectory(destDir)) return res.status(400).json({ error: "Quelle oder Ziel nicht erlaubt" });
    ensureWritable(destDir);
    const target = await uniquePathAsync(path.join(destDir, path.basename(source)));
    await copyRecursiveAsync(source, target);
    invalidateListPathCache();
    res.json({ ok: true, path: target });
  } catch (err) {
    if (err instanceof FsTimeoutError) return next(err);
    handleFsError(res, err, "copy fehlgeschlagen");
  }
});

app.post("/api/files/move", async (req, res, next) => {
  try {
    const source = await resolveAllowedHref(req.body?.sourceHref);
    const destDir = await resolveAllowedHref(req.body?.destHref);
    if (!source || !destDir || !await fsIsDirectory(destDir)) return res.status(400).json({ error: "Quelle oder Ziel nicht erlaubt" });
    if (path.resolve(destDir) === path.resolve(source) || path.resolve(destDir).startsWith(path.resolve(source) + path.sep)) {
      return res.status(400).json({ error: "Ordner kann nicht in sich selbst verschoben werden" });
    }
    ensureWritable(destDir);
    ensureWritable(path.dirname(source));
    const target = await uniquePathAsync(path.join(destDir, path.basename(source)));
    try {
      await fsTimeout(fsp.rename(source, target), `move ${source}`);
    } catch (err) {
      if (err.code !== "EXDEV") throw err;
      await copyRecursiveAsync(source, target);
      await fsTimeout(fsp.rm(source, { recursive: true, force: true }), `remove ${source}`);
    }
    invalidateListPathCache();
    res.json({ ok: true, path: target });
  } catch (err) {
    if (err instanceof FsTimeoutError) return next(err);
    handleFsError(res, err, "move fehlgeschlagen");
  }
});

app.post("/api/files/delete", async (req, res, next) => {
  try {
    const source = await resolveAllowedHref(req.body?.href);
    if (!source) return res.status(400).json({ error: "Pfad nicht erlaubt" });
    ensureWritable(path.dirname(source));
    sendToTrash(source);
    invalidateListPathCache();
    res.json({ ok: true });
  } catch (err) {
    if (err instanceof FsTimeoutError) return next(err);
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

async function uniquePathInDirAsync(parent, name) {
  const parsed = path.parse(name);
  let candidate = path.join(parent, name);
  let counter = 2;
  while (await fsExists(candidate)) {
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
    // WICHTIG: Pfad als EINFACH-gequoteter PowerShell-Literal übergeben. In doppelten
    // Anführungszeichen sind Backslashes literal, JSON.stringify hätte sie aber als
    // "\\" verdoppelt → kaputter Pfad → Fehler. In einfachen Quotes wird nur ' → ''.
    const psPath = `'${String(target).replace(/'/g, "''")}'`;
    const ps = [
      "Add-Type -AssemblyName Microsoft.VisualBasic",
      `$p = ${psPath}`,
      "if (Test-Path -LiteralPath $p -PathType Container) {",
      "  [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($p, 'OnlyErrorDialogs', 'SendToRecycleBin')",
      "} else {",
      "  [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($p, 'OnlyErrorDialogs', 'SendToRecycleBin')",
      "}",
    ].join("; ");
    const result = spawnSync("powershell.exe", ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", ps], { encoding: "utf8" });
    if (result.status === 0) return;
    // Netzlaufwerke (z.B. P:\) haben keinen Papierkorb → SendToRecycleBin scheitert.
    // Dann recoverbar in einen Papierkorb-Ordner auf demselben Laufwerk verschieben,
    // statt endgültig zu löschen oder mit 500 abzubrechen.
    const reason = (result.stderr || result.error?.message || "").trim();
    try {
      moveToRecycleFolder(target);
      return;
    } catch (fallbackErr) {
      throw new Error(`Papierkorb nicht verfügbar (${reason || "unbekannt"}); Verschieben in Papierkorb-Ordner ebenfalls fehlgeschlagen: ${fallbackErr.message}`);
    }
  }

  moveToRecycleFolder(target);
}

// Recoverbarer Fallback: Element in einen ".Papierkorb"-Ordner im selben Verzeichnis
// verschieben (mit Zeitstempel-Unterordner, damit nichts überschrieben wird). Der
// Ordner beginnt mit "." → wird in der Dateiliste und vom Scanner ausgeblendet.
function moveToRecycleFolder(target) {
  const trashDir = path.join(path.dirname(target), ".Papierkorb");
  fs.mkdirSync(trashDir, { recursive: true });
  const dest = uniquePathInDir(trashDir, path.basename(target));
  fs.renameSync(target, dest);
}

app.get("/api/list-path", async (req, res, next) => {
  try {
    const href = req.query?.href;
    const target = await resolveAllowedHref(href);
    if (!target) return res.status(400).json({ error: "Pfad nicht erlaubt oder nicht gefunden" });

    const stat = await fsStat(target);
    if (!stat.isDirectory()) return res.status(400).json({ error: "Pfad ist kein Ordner" });

    const urlPath = decodeHrefPath(href);
    const cacheKey = `${target}:${stat.mtimeMs}`;
    const cached = listPathCache.get(cacheKey);
    if (cached && Date.now() - cached.createdAt < LIST_PATH_CACHE_MS) return res.json(cached.payload);

    const rawEntries = await fsReaddir(target, { withFileTypes: true });
    const HIDE = /^[.~]|\.tmp$|\.db$/i;
    const entries = (await Promise.all(rawEntries
      .filter((entry) => !entry.name.startsWith(".") && !entry.name.startsWith("~$") && !/\.db$/i.test(entry.name) && !/\.tmp$/i.test(entry.name))
      .map(async (entry) => {
        const full = path.join(target, entry.name);
        const entryStat = await fsStat(full);
        const isDirectory = entry.isDirectory();
        let childCount = 0;
        if (isDirectory) {
          try {
            childCount = (await fsReaddir(full)).filter((n) => !HIDE.test(n)).length;
          } catch {
            childCount = 0;
          }
        }
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
      })))
      .sort((a, b) => Number(b.type === "Ordner") - Number(a.type === "Ordner") || b.mtime - a.mtime);
    // mtime bleibt im Payload, damit der Client nach Datum sortieren kann.

    const payload = { href: urlPath, entries };
    listPathCache.set(cacheKey, { createdAt: Date.now(), payload });
    if (listPathCache.size > 200) {
      const firstKey = listPathCache.keys().next().value;
      listPathCache.delete(firstKey);
    }
    res.json(payload);
  } catch (err) {
    next(err);
  }
});

app.post("/api/file-action", async (req, res, next) => {
  const { action, sourceHref, targetDirHref, name } = req.body || {};

  try {
    const targetDir = targetDirHref ? await resolveAllowedHref(targetDirHref) : null;
    if (action === "mkdir") {
      if (!targetDir || !await fsIsDirectory(targetDir)) return res.status(400).json({ error: "target folder invalid" });
      const safeName = validateEntryName(name);
      if (!safeName) return res.status(400).json({ error: "folder name missing" });
      const created = await uniquePathInDirAsync(targetDir, safeName);
      await fsTimeout(fsp.mkdir(created, { recursive: false }), `mkdir ${created}`);
      clearListPathCache();
      return res.json({ ok: true, path: created });
    }

    const source = sourceHref ? await resolveAllowedHref(sourceHref) : null;
    if (!source || !await fsExists(source)) return res.status(400).json({ error: "source invalid" });

    if (action === "delete") {
      sendToTrash(source);
      clearListPathCache();
      return res.json({ ok: true });
    }

    if (action === "rename") {
      const safeName = validateEntryName(name);
      if (!safeName) return res.status(400).json({ error: "name missing" });
      const parent = path.dirname(source);
      const destination = path.join(parent, safeName);
      if (path.resolve(source) === path.resolve(destination)) return res.json({ ok: true, path: source });
      if (await fsExists(destination)) return res.status(409).json({ error: "target exists" });
      ensureWritable(parent);
      await fsTimeout(fsp.rename(source, destination), `rename ${source}`);
      clearListPathCache();
      return res.json({ ok: true, path: destination });
    }

    if (!targetDir || !await fsIsDirectory(targetDir)) return res.status(400).json({ error: "target folder invalid" });
    const destination = await uniquePathInDirAsync(targetDir, path.basename(source));
    if (path.resolve(destination).startsWith(path.resolve(source) + path.sep)) {
      return res.status(400).json({ error: "cannot move/copy folder into itself" });
    }

    if (action === "copy") {
      await copyRecursiveAsync(source, destination);
      clearListPathCache();
      return res.json({ ok: true, path: destination });
    }

    if (action === "move") {
      try {
        await fsTimeout(fsp.rename(source, destination), `move ${source}`);
      } catch (err) {
        if (err.code !== "EXDEV") throw err;
        await copyRecursiveAsync(source, destination);
        await fsTimeout(fsp.rm(source, { recursive: true, force: true }), `remove ${source}`);
      }
      clearListPathCache();
      return res.json({ ok: true, path: destination });
    }

    res.status(400).json({ error: "unknown action" });
  } catch (err) {
    if (err instanceof FsTimeoutError) return next(err);
    res.status(500).json({ error: err.message });
  }
});

// ── Tabelle 24 (component-certificate table scanner) ──────────────

async function resolveTabelle24Root(query) {
  if (query.root && typeof query.root === "string") {
    const resolved = path.resolve(query.root);
    return (await fsExists(resolved)) ? resolved : null;
  }

  if (query.href && typeof query.href === "string") {
    return resolveAllowedHrefSync(query.href);
  }

  const projectId = query.projectId;
  if (!projectId || typeof projectId !== "string") return null;

  const project = getProject(projectId);
  const bautelliste = project?.documentGroups?.find((group) =>
    group.area === "Bautelliste" || /^09\b/.test(group.primary || "")
  );
  if (bautelliste?.href) {
    const resolved = resolveAllowedHrefSync(bautelliste.href);
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
  for (const candidate of candidates) { if (await fsExists(candidate)) return candidate; }
  return null;
}

async function normalizeTabelle24SearchRoot(root) {
  const base = path.basename(root).toLowerCase();
  if (/^09\b/.test(base) && base.includes("baut")) return root;

  const candidates = [
    path.join(root, "Zulassungen", "IEC"),
    path.join(root, "IEC"),
    root,
  ];
  for (const iecRoot of candidates) {
    let entries;
    try { entries = await fsReaddir(iecRoot, { withFileTypes: true }); } catch { continue; }
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

// List Intern Bauteilliste files in a project that actually contain Tabelle 24.
// Walk (skip Archiv subdirs, require "intern" in name per workflow rule — never
// the customer VDE copies), then content-filter via pure Node marker scan.
// Nativer Datei-Auswahldialog (Finder/Explorer) → liefert den echten Pfad zurück,
// damit der Benutzer die Bauteilliste manuell wählen kann statt per Pfad-Suche.
function pickFileDialog() {
  return new Promise((resolve) => {
    let cmd, args;
    if (process.platform === "darwin") {
      const script = [
        'try',
        '  set f to choose file with prompt "Bauteilliste / Datei wählen" of type {"docx","doc","docm","xlsx","xlsm","xls"}',
        '  POSIX path of f',
        'on error number -128',
        '  return ""',
        'end try',
      ].join("\n");
      cmd = "osascript"; args = ["-e", script];
    } else if (process.platform === "win32") {
      // Topmost-Owner-Form, damit der Dialog VOR dem Browser erscheint (sonst öffnet
      // er sich unsichtbar dahinter). Filter optional via env, Standard: Office-Dateien.
      const ps = WIN_TOPMOST_OWNER
        + "$d = New-Object System.Windows.Forms.OpenFileDialog; "
        + "$d.Filter = 'Word/Excel|*.docx;*.doc;*.docm;*.xlsx;*.xlsm;*.xls|Alle Dateien (*.*)|*.*'; "
        + "$d.Title = 'Bauteilliste / Datei waehlen'; "
        + "$r = $d.ShowDialog($o); $o.Close(); "
        + "if ($r -eq [System.Windows.Forms.DialogResult]::OK) { [Console]::Out.Write($d.FileName) }";
      cmd = "powershell.exe"; args = WIN_PS_ARGS(ps);
    } else {
      resolve({ error: "Plattform nicht unterstützt" }); return;
    }
    let out = "", errOut = "";
    try {
      const p = spawn(cmd, args);
      p.stdout.on("data", (d) => { out += d.toString(); });
      p.stderr.on("data", (d) => { errOut += d.toString(); });
      p.on("error", (e) => resolve({ error: e.message }));
      p.on("close", () => {
        const pth = out.trim();
        if (pth) return resolve({ path: pth });
        resolve(errOut.trim() ? { error: errOut.trim() } : { canceled: true });
      });
    } catch (e) {
      resolve({ error: e.message });
    }
  });
}

// Gemeinsamer PowerShell-Vorspann: unsichtbare, immer-im-Vordergrund Owner-Form ($o),
// damit Datei-/Ordner-Dialoge zuverlässig VOR dem Browser erscheinen.
const WIN_TOPMOST_OWNER =
  "Add-Type -AssemblyName System.Windows.Forms; "
  + "$o = New-Object System.Windows.Forms.Form; $o.TopMost=$true; $o.ShowInTaskbar=$false; "
  + "$o.Opacity=0; $o.Width=1; $o.Height=1; $o.StartPosition='CenterScreen'; "
  + "$o.Add_Shown({ $o.Activate() }); $o.Show(); ";
const WIN_PS_ARGS = (ps) => ["-NoProfile", "-ExecutionPolicy", "Bypass", "-STA", "-WindowStyle", "Hidden", "-Command", ps];

app.post("/api/pick-file", async (_req, res) => {
  try { res.json(await pickFileDialog()); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

// Nativer Ordner-Auswahldialog (Finder/Explorer) → echter Ordnerpfad.
function pickFolderDialog() {
  return new Promise((resolve) => {
    let cmd, args;
    if (process.platform === "darwin") {
      const script = [
        'try',
        '  set f to choose folder with prompt "Richtigen Ordner wählen"',
        '  POSIX path of f',
        'on error number -128',
        '  return ""',
        'end try',
      ].join("\n");
      cmd = "osascript"; args = ["-e", script];
    } else if (process.platform === "win32") {
      const ps = WIN_TOPMOST_OWNER
        + "$d = New-Object System.Windows.Forms.FolderBrowserDialog; "
        + "$d.Description = 'Richtigen Ordner waehlen'; "
        + "$r = $d.ShowDialog($o); $o.Close(); "
        + "if ($r -eq [System.Windows.Forms.DialogResult]::OK) { [Console]::Out.Write($d.SelectedPath) }";
      cmd = "powershell.exe"; args = WIN_PS_ARGS(ps);
    } else {
      resolve({ error: "Plattform nicht unterstützt" }); return;
    }
    let out = "", errOut = "";
    try {
      const p = spawn(cmd, args);
      p.stdout.on("data", (d) => { out += d.toString(); });
      p.stderr.on("data", (d) => { errOut += d.toString(); });
      p.on("error", (e) => resolve({ error: e.message }));
      p.on("close", () => {
        const pth = out.trim();
        if (pth) return resolve({ path: pth });
        resolve(errOut.trim() ? { error: errOut.trim() } : { canceled: true });
      });
    } catch (e) { resolve({ error: e.message }); }
  });
}

app.post("/api/pick-folder", async (_req, res) => {
  try { res.json(await pickFolderDialog()); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

// Mitgelieferte Beispiel-Bauteilliste (EF1157) — damit Tabelle 24 sofort testbar ist,
// auch ohne echtes P:\PCS. Liefert den absoluten Pfad zum Parsen.
app.get("/api/tabelle24/sample", (_req, res) => {
  const p = path.join(__dirname, "samples", "EF1157-Bauteilliste-Beispiel.docx");
  if (!fs.existsSync(p)) return res.status(404).json({ error: "Beispiel-Datei nicht gefunden" });
  res.json({ path: p, name: path.basename(p) });
});

// Ordner-Info (PDF-Anzahl, Dateien, Archiv) für einen frei gewählten Pfad —
// liefert dieselbe Form wie die Analyse, damit das Badge danach korrekt aussieht.
app.get("/api/tabelle24/folder-info", (req, res) => {
  const p = req.query?.path;
  if (!p || typeof p !== "string") return res.status(400).json({ error: "missing ?path=" });
  const resolved = path.resolve(p);
  const exists = fs.existsSync(resolved);
  const files = exists ? listFilesDeep(resolved, 2) : [];
  const pdfs = files.filter((f) => /\.pdf$/i.test(f));
  res.json({
    targetFolder: resolved,
    targetExists: exists,
    hasArchiv: exists && fs.existsSync(path.join(resolved, "Archiv")),
    pdfCount: pdfs.length,
    fileCount: files.length,
    oldPdfs: pdfs.map((f) => ({ name: path.basename(f), path: f })),
  });
});

app.get("/api/tabelle24/files", async (req, res, next) => {
 try {
  const resolved = await resolveTabelle24Root(req.query);
  if (!resolved) return res.status(400).json({ error: "missing or unresolved ?root=, ?href=, or ?projectId=" });
  const searchRoot = await normalizeTabelle24SearchRoot(resolved);

  const allCandidates = [];
  async function walk(dir, depth) {
    if (depth > 4) return;
    let entries;
    try { entries = await fsReaddir(dir, { withFileTypes: true }); } catch { return; }
    for (const e of entries) {
      const full = path.join(dir, e.name);
      if (e.isDirectory()) {
        if (/^archiv$/i.test(e.name)) continue;
        await walk(full, depth + 1);
      } else if (/\.(docx?|docm)$/i.test(e.name) && /intern/i.test(e.name) && !/^~\$/.test(e.name)) {
        allCandidates.push(full);
      }
    }
  }
  await walk(searchRoot, 0);

  // Yield between per-file marker reads so a slow share stays responsive
  // (fileContainsTabelle24 itself is unchanged / still synchronous).
  const files = [];
  for (const p of allCandidates) {
    if (cachedMarker(p, "t24", () => fileContainsTabelle24(p))) {
      const stat = await fsStat(p);
      files.push({ path: p, name: path.basename(p), size: stat.size, mtime: stat.mtime.toISOString() });
    }
    await new Promise((r) => setImmediate(r));
  }
  files.sort((a, b) => b.mtime.localeCompare(a.mtime));
  res.json({ root: searchRoot, files });
 } catch (e) { next(e); }
});

// Look up a VDE certificate, return PDF list + extracted date + comparison verdict.
app.post("/api/tabelle24/vde-lookup", async (req, res) => {
  const { certNumber, currentDate } = req.body || {};
  if (!certNumber || typeof certNumber !== "string") return res.status(400).json({ error: "missing certNumber" });
  try {
    res.json(await lookupVde({ certNumber, currentDate }));
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Serve a file from anywhere under ~/Desktop/ (used by Tabelle 24 PDF thumbnails).
// Path is given as an absolute filesystem path; we resolve + verify it stays inside Desktop.
app.get("/api/tabelle24/file", (req, res) => {
  const p = req.query.path;
  if (!p || typeof p !== "string") return res.status(400).send("missing ?path");
  const safe = path.resolve(p);
  const desktop = path.resolve(os.homedir(), "Desktop");
  if (!isWithin(desktop, safe) || !fs.existsSync(safe)) return res.status(404).send("not found");
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
  // Only allow paths inside the OS temp dir's vde-lookups subdir (where VDE lookup writes).
  const allowed = path.resolve(os.tmpdir(), "vde-lookups");
  if (!isWithin(allowed, safe) || !fs.existsSync(safe)) return res.status(404).send("not found");
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

  try {
    const lookup = await lookupVde({ certNumber, currentDate });
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
app.post("/api/tabelle24/save", (req, res) => {
  const { sourceFile, outputFile, changes } = req.body || {};
  if (!sourceFile || typeof sourceFile !== "string") return res.status(400).json({ error: "missing sourceFile" });
  if (!Array.isArray(changes) || !changes.length) return res.status(400).json({ error: "missing changes[]" });
  if (!fs.existsSync(sourceFile)) return res.status(404).json({ error: `source not found: ${sourceFile}` });
  try {
    res.json(updateTabelle24({ sourceFile, outputFile, changes }));
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Tabelle 30 (VDE Typenprüfung — "Resistance to heat and fire") ─────────────
const TABELLE30_MARKER = "Resistance to heat and fire";
const TABELLE30_MARKER_ALT = "TABLE: Resistance to heat";

function fileContainsTabelle30(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  try {
    if (ext === ".docx" || ext === ".docm") {
      const xml = readDocumentXml(filePath);
      return xml.includes(TABELLE30_MARKER) || xml.includes(TABELLE30_MARKER_ALT);
    }
    if (ext === ".doc") {
      const raw = fs.readFileSync(filePath);
      return raw.includes(Buffer.from(TABELLE30_MARKER, "utf16le"))
        || raw.includes(Buffer.from(TABELLE30_MARKER, "utf8"))
        || raw.includes(Buffer.from(TABELLE30_MARKER_ALT, "utf16le"));
    }
  } catch {
    return false;
  }
  return false;
}

// List Word files (anywhere under the project root) that actually contain a Tabelle 30
// section. Two-stage: walk → collect all .doc/.docx/.docm → pipe paths through
// a pure-Node marker scan (fast zip read for .docx/.docm, raw-byte search for
// .doc — no Python or LibreOffice needed for the file picker).
// Resolve the project root folder for Tabelle 30 file discovery. Accepts an explicit
// ?root= (back-compat) or a ?projectId= that we resolve under PCS_ROOT (P:\PCS on
// Windows, ~/Desktop in Mac dev). This keeps the frontend free of any hardcoded OS
// path, so the same build finds the report/Excel on both macOS and Windows.
async function resolveTabelle30Root(query) {
  if (query.root && typeof query.root === "string" && query.root.trim()) {
    const r = path.resolve(query.root);
    return (await fsExists(r)) ? r : null;
  }
  const projectId = typeof query.projectId === "string" ? query.projectId.trim() : "";
  if (!projectId) return null;
  const numeric = projectId.replace(/^EF/i, "").trim();
  const direct = [
    path.join(PCS_ROOT, projectId),
    path.join(PCS_ROOT, numeric),
    path.join(PCS_ROOT, `EF${numeric}`),
    path.join(PCS_ROOT, `EF ${numeric}`),
  ];
  for (const c of direct) { if (await fsIsDirectory(c)) return c; }
  // Fall back to scanning PCS_ROOT for a folder beginning with the project number
  // (handles names like "1157 EF1157 CoffeeB Pluto" or "EF1157 …").
  try {
    const re = new RegExp(`^(ef[\\s_-]*)?0*${numeric}(\\D|$)`, "i");
    for (const e of await fsReaddir(PCS_ROOT, { withFileTypes: true })) {
      if (e.isDirectory() && re.test(e.name)) return path.join(PCS_ROOT, e.name);
    }
  } catch {}
  return null;
}

app.get("/api/tabelle30/files", async (req, res, next) => {
 try {
  const resolved = await resolveTabelle30Root(req.query);
  if (!resolved) return res.status(400).json({ error: "missing/unresolved ?root= or ?projectId=" });

  const allCandidates = [];
  async function walk(dir, depth) {
    if (depth > 5) return;
    let entries;
    try { entries = await fsReaddir(dir, { withFileTypes: true }); } catch { return; }
    for (const e of entries) {
      const full = path.join(dir, e.name);
      if (e.isDirectory()) {
        await walk(full, depth + 1);
      } else if (/\.(docx?|docm)$/i.test(e.name) && !/^~\$/.test(e.name)) {
        allCandidates.push(full);
      }
    }
  }
  await walk(resolved, 0);

  // Yield between per-file marker reads so a slow share doesn't monopolise the
  // event loop (fileContainsTabelle30 itself is unchanged / still synchronous).
  const files = [];
  for (const p of allCandidates) {
    if (cachedMarker(p, "t30", () => fileContainsTabelle30(p))) {
      const stat = await fsStat(p);
      files.push({ path: p, name: path.basename(p), size: stat.size, mtime: stat.mtime.toISOString() });
    }
    await new Promise((r) => setImmediate(r));
  }
  files.sort((a, b) => b.mtime.localeCompare(a.mtime));
  res.json({ root: resolved, files });
 } catch (e) { next(e); }
});

// ── Tabelle 30 Phase 2: Excel-Vergleich ──────────────────────
// List candidate "Ergänzung" Excel files under the project root — typically in
// "02 Änderungen/" with "ergänzung" or "ergaenzung" in the filename.
app.get("/api/tabelle30/excels", async (req, res, next) => {
 try {
  const showAll = req.query.all === "1" || req.query.all === "true";
  const resolved = await resolveTabelle30Root(req.query);
  if (!resolved) return res.status(400).json({ error: "missing/unresolved ?root= or ?projectId=" });

  const found = [];
  async function walk(dir, depth) {
    if (depth > 5) return;
    let entries;
    try { entries = await fsReaddir(dir, { withFileTypes: true }); } catch { return; }
    for (const e of entries) {
      const full = path.join(dir, e.name);
      if (e.isDirectory()) {
        await walk(full, depth + 1);
      } else if (/\.(xlsx|xlsm)$/i.test(e.name) && !/^~\$/.test(e.name) && (showAll || /(erg[äa]nzung|material|change|alternativ)/i.test(e.name + dir))) {
        const stat = await fsStat(full);
        found.push({ path: full, name: e.name, size: stat.size, mtime: stat.mtime.toISOString() });
      }
    }
  }
  await walk(resolved, 0);
  const resultFiles = found;
  const excelPriority = (f) => {
    const hay = `${f.name} ${f.path}`.toLowerCase();
    if (hay.includes("alternative") || hay.includes("alternativ")) return 0;
    if (hay.includes("02 änderungen") || hay.includes("02 änderungen") || hay.includes("ergänzung") || hay.includes("ergaenzung") || hay.includes("change")) return 1;
    if (hay.includes("tabelle30") || hay.includes("tabelle 30")) return 2;
    if (hay.includes("material")) return 3;
    return 4;
  };
  resultFiles.sort((a, b) => excelPriority(a) - excelPriority(b) || b.mtime.localeCompare(a.mtime));
  res.json({ root: resolved, files: resultFiles });
 } catch (e) { next(e); }
});

// Compare an Excel Ergänzung file against the current Tabelle 30 from a Word report.
// Returns categorized rows: new (in Excel only), matched (both), and skipped Excel rows
// like "Keep ..." entries that aren't real material changes.
app.post("/api/tabelle30/compare", (req, res) => {
  const { wordFile, excelFile } = req.body || {};
  if (!wordFile || !excelFile) return res.status(400).json({ error: "missing { wordFile, excelFile }" });
  if (!fs.existsSync(wordFile))  return res.status(404).json({ error: `wordFile not found: ${wordFile}` });
  if (!fs.existsSync(excelFile)) return res.status(404).json({ error: `excelFile not found: ${excelFile}` });

  try {
    const wordData = cachedParse(parsedTabelle30Cache, wordFile, () => parseTabelle30(wordFile));
    const excelData = cachedParse(parsedExcelCache, excelFile, () => parseExcelErgaenzung(excelFile));
    if (excelData.format === "part-list" && !excelData.hasTab30Markers && !excelData.hasChangeMarkers) {
      return res.status(400).json({
        error: `Dieses Excel hat keine Tab30-Markierungen und keine erkennbaren "Add the part number"-Änderungen: ${path.basename(excelFile)}.`,
        excelFile,
        excelRowCount: excelData.rowCount,
      });
    }

    // Matching strategy: Word and Excel name the same components differently
    // ("Panel Rear CC, Pos. 10" vs "Panel Rear CC CM"). We use a "stem":
    // first segment before comma / "Pos." / multi-word suffix, normalized + truncated.
    // A Word row matches an Excel row when stems match AND material has substring overlap.
    const norm = (s) => String(s || "")
      .toLowerCase()
      .replace(/\s+/g, " ")
      .replace(/[^\w\s]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
    const stem = (s) => {
      const n = norm(s);
      // Strip "pos N" / "pos. N" suffixes and anything after a comma
      return n
        .replace(/\b(pos|position)\b.*$/, "")
        .split(",")[0]
        .trim()
        .split(/\s+/)
        .filter(Boolean)
        .slice(0, 4)
        .join(" ");
    };
    const materialKey = (s) => {
      // Material identifiers like "ABS Terluran GP22" — keep first 3-4 meaningful tokens
      return norm(s).split(/\s+/).filter((t) => t.length >= 2 && !/^\d+$/.test(t)).slice(0, 4).join(" ");
    };
    const colorKey = (s) => {
      const words = norm(s).split(/\s+/);
      return ["black", "white", "green", "red", "grey", "gray", "transparent"].find((c) => words.includes(c)) || "";
    };
    const positionsFromText = (s) => {
      const positions = new Set();
      const text = String(s || "");
      const re = /\bpos\.?\s*([0-9][0-9\-]*(?:\/[0-9\-]+)?)/gi;
      let match;
      while ((match = re.exec(text))) {
        const raw = match[1];
        positions.add(raw);
        if (raw.includes("/")) {
          const [base, suffix] = raw.split("/");
          positions.add(base);
          const parts = base.split("-");
          if (parts.length > 1 && suffix) {
            positions.add([...parts.slice(0, -1), suffix].join("-"));
          }
        }
      }
      return positions;
    };

    // Index Word rows by stem of partName
    const wordByStem = new Map();
    const wordByPosition = new Map();
    for (const row of wordData.rows) {
      const partName = row.cells?.[0] || "";
      const material = row.cells?.[2] || "";
      const refLast  = row.cells?.[row.cells.length - 1] || "";
      const partStem = stem(partName);
      const indexed = {
        rowIdx: row.rowIdx,
        partName: partName.trim(),
        material: material.trim(),
        matKey: materialKey(material),
        color: colorKey(material),
        refLast: refLast.trim(),
      };
      if (!partStem) continue;
      if (!wordByStem.has(partStem)) wordByStem.set(partStem, []);
      wordByStem.get(partStem).push(indexed);
      for (const pos of positionsFromText(partName)) {
        if (!wordByPosition.has(pos)) wordByPosition.set(pos, []);
        wordByPosition.get(pos).push(indexed);
      }
    }

    const matched = [];
    const newEntries = [];
    const skipped = [];
    const changeMarkerEntries = [];
    const isPlaceholderMaterial = (s) => {
      const compact = String(s || "").replace(/[\s/.,;:_-]+/g, "");
      return !compact || /^x+$/i.test(compact) || /^na$/i.test(compact);
    };

    for (const excelRow of excelData.rows) {
      const changeTo = excelRow.changeTo || "";
      // Skip rows that are "Keep ...", empty, or placeholder-only ("--- --- ---").
      if (isPlaceholderMaterial(changeTo) || /^keep\b/i.test(changeTo) || changeTo === excelRow.todayUsed) {
        skipped.push({ excelRow, reason: isPlaceholderMaterial(changeTo) ? "empty / placeholder material" : "no change / Keep" });
        continue;
      }

      if (excelData.hasChangeMarkers && !excelData.hasTab30Markers) {
        // Supplier change-list: rows are flagged "new" by the supplier, but we still
        // cross-check each against the existing Tabelle 30 so only genuinely missing
        // parts are added (see tabelle30-match.js). Classified after the loop.
        changeMarkerEntries.push({ excelRow, candidatePartMatches: [] });
        continue;
      }

      const partStem = stem(excelRow.partName);
      const excelMatKey = materialKey(changeTo);
      const excelColor = colorKey(changeTo);
      const excelPosition = String(excelRow.position || "").trim();
      // Prefer exact exploded-position matches; Tabelle 30 often groups several
      // object names in one Word row, so name stems alone miss existing rows.
      let candidates = excelPosition ? (wordByPosition.get(excelPosition) || []) : [];
      const hadPositionCandidates = candidates.length > 0;
      // Direct stem hit
      if (!candidates.length) candidates = wordByStem.get(partStem) || [];
      // If no direct stem hit, allow looser: any Word stem that shares ≥2 tokens with this Excel stem
      if (!candidates.length) {
        const excelTokens = partStem.split(/\s+/).filter(Boolean);
        for (const [wStem, rows] of wordByStem.entries()) {
          const wTokens = wStem.split(/\s+/).filter(Boolean);
          const shared = excelTokens.filter((t) => wTokens.includes(t)).length;
          if (shared >= 2) candidates.push(...rows);
        }
      }

      // Check material overlap among candidates
      const exact = candidates.find((c) => {
        if (!c.matKey || !excelMatKey) return false;
        if (c.color && excelColor && c.color !== excelColor) return false;
        const cTokens = c.matKey.split(/\s+/).filter(Boolean);
        const eTokens = excelMatKey.split(/\s+/).filter(Boolean);
        return cTokens.filter((t) => eTokens.includes(t)).length >= 2;
      });

      if (exact) {
        matched.push({ excelRow, wordRowIdx: exact.rowIdx, wordMaterial: exact.material });
      } else if (hadPositionCandidates) {
        // The object/position already exists in Tabelle 30. Do not create a
        // duplicate preview row just because material wording differs slightly
        // (e.g. PA-757 vs PA-757F+masterbatch); surface it as matched for now.
        const existing = candidates[0];
        matched.push({
          excelRow,
          wordRowIdx: existing.rowIdx,
          wordMaterial: existing.material,
          materialWarning: true,
        });
      } else {
        newEntries.push({
          excelRow,
          candidatePartMatches: candidates.map((c) => ({ rowIdx: c.rowIdx, partName: c.partName, material: c.material })),
        });
      }
    }

    // Supplier change-list entries: cross-check against the existing Tabelle 30 rows.
    // Only parts whose name AND material grade both strongly match an existing row are
    // treated as already present (excluded). Everything else stays on the add-list;
    // partial overlaps keep a non-blocking warning so the engineer can verify.
    if (changeMarkerEntries.length) {
      const { add, present } = classifyNewEntries(changeMarkerEntries, wordData.rows);
      for (const e of add) {
        newEntries.push({
          excelRow: e.excelRow,
          candidatePartMatches: [],
          warning: e._match.warning || null,
          matchRowIdx: e._match.matchRowIdx || null,
        });
      }
      for (const e of present) {
        matched.push({
          excelRow: e.excelRow,
          wordRowIdx: e._match.matchRowIdx,
          alreadyPresent: true,
          note: e._match.warning,
        });
      }
    }

    res.json({
      wordFile,
      excelFile,
      wordRowCount: wordData.rowCount,
      excelRowCount: excelData.rowCount,
      summary: {
        matched: matched.length,
        new: newEntries.length,
        skipped: skipped.length,
      },
      matched, new: newEntries, skipped,
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Apply selected Excel entries to Tabelle 30 in the Word file.
// Body: { sourceFile, outputFile?, entries: [{ partName, supplier, material, reference }] }
// Returns: { outputFile, applied, warnings } or { error }.
app.post("/api/tabelle30/apply", (req, res) => {
  const { sourceFile, outputFile, entries } = req.body || {};
  if (!sourceFile || typeof sourceFile !== "string") return res.status(400).json({ error: "missing sourceFile" });
  if (!Array.isArray(entries) || !entries.length) return res.status(400).json({ error: "no entries to apply" });
  if (!fs.existsSync(sourceFile)) return res.status(404).json({ error: `source not found: ${sourceFile}` });
  try {
    res.json(updateTabelle30({ sourceFile, outputFile, entries }));
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Parse Tabelle 30 from a Typenprüfung file.
app.post("/api/tabelle30/parse", (req, res) => {
  const file = req.body?.file;
  if (!file || typeof file !== "string") return res.status(400).json({ error: "missing { file }" });
  const resolved = path.resolve(file);
  if (!fs.existsSync(resolved)) return res.status(404).json({ error: `not found: ${resolved}` });
  try {
    res.json(cachedParse(parsedTabelle30Cache, resolved, () => parseTabelle30(resolved)));
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Parse one Bauteilliste file → Tabelle 24 rows as JSON.
app.post("/api/tabelle24/parse", (req, res) => {
  const file = req.body?.file;
  if (!file || typeof file !== "string") return res.status(400).json({ error: "missing { file }" });
  const resolved = path.resolve(file);
  if (!fs.existsSync(resolved)) return res.status(404).json({ error: `not found: ${resolved}` });
  try {
    res.json(cachedParse(parsedTabelle24Cache, resolved, () => parseTabelle24(resolved)));
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Build the automation worklist: row -> target folder -> M3/certificate tasks.
app.post("/api/tabelle24/analyze", async (req, res) => {
  const file = req.body?.file;
  if (!file || typeof file !== "string") return res.status(400).json({ error: "missing { file }" });
  const resolved = path.resolve(file);
  if (!fs.existsSync(resolved)) return res.status(404).json({ error: `not found: ${resolved}` });

  try {
    const parsed = cachedParse(parsedTabelle24Cache, resolved, () => parseTabelle24(resolved));
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

app.use((err, _req, res, _next) => {
  const status = err?.status || 500;
  const message = err instanceof FsTimeoutError
    ? "Netzlaufwerk antwortet nicht. Bitte P:/M: Verbindung prüfen oder später erneut versuchen."
    : (err?.message || "Interner Serverfehler");
  if (status >= 500) console.warn(`Request failed (${status}): ${err?.message || err}`);
  res.status(status).json({ error: message });
});

// ── Start ──────────────────────────────────────────────────
const server = app.listen(PORT, HOST, () => {
  console.log(`PCS Dashboard → http://${HOST}:${PORT}`);
  console.log(`Datenbank: pcs.db  |  Dokumente: evidence-1157/`);
  if (process.env.AUTO_ARCHIVE_SCAN === "1") {
    setTimeout(() => {
      runArchiveScan(ARCHIVE_EXCEL_PATH)
        .then((n) => console.log(`Archiv-Excel: ${n} Einträge synchronisiert`))
        .catch((e) => console.warn(`Archiv-Excel nicht geladen: ${e.message}`));
    }, 100);
  } else {
    console.log("Archiv-Excel: Auto-Sync deaktiviert (bei Bedarf im UI starten)");
  }
});
server.requestTimeout = FS_TIMEOUT_MS + 5000;
server.headersTimeout = FS_TIMEOUT_MS + 7000;
