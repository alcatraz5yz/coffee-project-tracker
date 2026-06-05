// ── DOM refs ────────────────────────────────────────────────
const machineList = document.querySelector("#machine-list");
const search = document.querySelector("#search");
const views = {
  dashboard: document.querySelector("#dashboard-view"),
  overview: document.querySelector("#overview-view"),
  subtopic: document.querySelector("#subtopic-view"),
  tasks: document.querySelector("#tasks-view"),
  docs: document.querySelector("#docs-view"),
  calculations: document.querySelector("#calculations-view"),
  tabelle24: document.querySelector("#tabelle24-view"),
  tabelle30: document.querySelector("#tabelle30-view"),
  freigabe: document.querySelector("#freigabe-view"),
  packaging: document.querySelector("#packaging-view")
};
const buttons = {
  overview: document.querySelector("#view-overview"),
  subtopic: document.querySelector("#view-subtopic"),
  tasks: document.querySelector("#view-tasks"),
  docs: document.querySelector("#view-docs"),
  calculations: document.querySelector("#view-calculations"),
  tabelle24: document.querySelector("#tabelle24-link"),
  tabelle30: document.querySelector("#tabelle30-link"),
  freigabe: document.querySelector("#view-freigabe"),
  packaging: document.querySelector("#view-packaging")
};
const tabelle24Link = document.querySelector("#tabelle24-link");
const tabelle24Frame = document.querySelector("#tabelle24-frame");
const tabelle30Link = document.querySelector("#tabelle30-link");
const tabelle30Frame = document.querySelector("#tabelle30-frame");
const taskFilter = document.querySelector("#task-filter");
const subtopicFilter = document.querySelector("#subtopic-filter");
const zifferTable = document.querySelector("#ziffer-table");
const calculationsContainer = document.querySelector("#calculations-container");
const calcReset = document.querySelector("#calc-reset");
const themeToggle = document.querySelector("#theme-toggle");
const dashboardLink = document.querySelector("#dashboard-link");

// ── PDF thumbnail rendering ───────────────────────────────────
// Lazy: an IntersectionObserver triggers PDF.js render only when the canvas scrolls into view.
// On failure (corrupt PDF, network issue), draws a red "PDF" badge so the row stays visible.
let _pdfThumbObserver = null;

function getPdfThumbObserver() {
  if (_pdfThumbObserver) return _pdfThumbObserver;
  if (typeof window.pdfjsLib === "undefined") return null;
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
  _pdfThumbObserver = new IntersectionObserver((entries) => {
    for (const e of entries) {
      if (!e.isIntersecting) continue;
      _pdfThumbObserver.unobserve(e.target);
      renderPdfThumb(e.target);
    }
  }, { rootMargin: "200px" });
  return _pdfThumbObserver;
}

async function renderPdfThumb(canvas) {
  canvas.dataset.pdfDone = "1";
  try {
    const pdf = await pdfjsLib.getDocument({ url: canvas.dataset.pdfHref }).promise;
    const page = await pdf.getPage(1);
    const SIZE = 40;
    const vp0 = page.getViewport({ scale: 1 });
    const scale = SIZE / Math.min(vp0.width, vp0.height);
    const vp = page.getViewport({ scale });
    canvas.width = vp.width;
    canvas.height = vp.height;
    await page.render({ canvasContext: canvas.getContext("2d"), viewport: vp }).promise;
  } catch {
    const ctx = canvas.getContext("2d");
    canvas.width = 40; canvas.height = 40;
    ctx.fillStyle = "#fee2e2"; ctx.fillRect(0, 0, 40, 40);
    ctx.fillStyle = "#dc2626"; ctx.font = "bold 9px sans-serif";
    ctx.textAlign = "center"; ctx.textBaseline = "middle";
    ctx.fillText("PDF", 20, 20);
  }
}

window._observePdfThumbs = () => {
  const io = getPdfThumbObserver();
  if (!io) return;
  document.querySelectorAll("canvas.file-thumb-pdf:not([data-pdf-done])").forEach(c => io.observe(c));
};

function fileThumbHtml(entry) {
  const name = entry.name || "";
  const href = entry.href || "";
  if (/\.(jpe?g|png|gif|webp|svg|bmp)$/i.test(name)) {
    return `<img src="${href}" loading="lazy" class="file-thumb-img" alt="">`;
  }
  if (/\.pdf$/i.test(name)) {
    return `<canvas class="file-thumb-img file-thumb-pdf" data-pdf-href="${href}"></canvas>`;
  }
  if (isExcelFile?.(name)) return `<span class="file-thumb-icon file-thumb-icon--excel">XLS</span>`;
  if (isWordFile?.(name))  return `<span class="file-thumb-icon file-thumb-icon--word">DOC</span>`;
  return `<span class="file-thumb-icon file-thumb-icon--file">FILE</span>`;
}

// ── Recent projects (localStorage) ──────────────────────────
const RECENT_KEY = "pcs-recent";
const RECENT_MAX = 6;

function getRecentProjects() {
  try { return JSON.parse(localStorage.getItem(RECENT_KEY) || "[]"); } catch { return []; }
}

function recordRecent(projectId) {
  const recent = getRecentProjects().filter((r) => r.id !== projectId);
  recent.unshift({ id: projectId });
  if (recent.length > RECENT_MAX) recent.length = RECENT_MAX;
  localStorage.setItem(RECENT_KEY, JSON.stringify(recent));
}

// ── State ────────────────────────────────────────────────────
let projectList = [];
let activeProject = null;
let activeSubtopic = "Approbation";
let activeBuild = "Alle";
let activeView = "overview";
let activeEvidenceGroup = null;
let activeMarketFilter = null;
const evidenceEntries = new Map();
const evidencePathEntries = new Map();
let selectedEvidenceHref = null;
let selectedEvidenceHrefs = new Set();
let evidenceSelectionAnchor = null;
let evidenceClipboard = null;
let evidenceDrag = null;
// Sortierung + Suche im aktuellen Ordner.
let evidenceSort = { key: "name", dir: "asc" };
let evidenceSearch = "";
// Pro Ordner (Gruppe) den zuletzt besuchten Unterordner merken, damit man beim
// Wechsel zwischen Ordnern in seinem jeweiligen Unterordner bleibt.
const lastFolderByGroup = new Map();
function sortEvidenceEntries(entries) {
  const { key, dir } = evidenceSort;
  const mul = dir === "asc" ? 1 : -1;
  const ext = (n) => { const i = String(n).lastIndexOf("."); return i > 0 ? n.slice(i + 1).toLowerCase() : ""; };
  return entries.slice().sort((a, b) => {
    // Ordner immer zuerst (wie im Finder/Explorer).
    const fa = a.type === "Ordner" ? 0 : 1;
    const fb = b.type === "Ordner" ? 0 : 1;
    if (fa !== fb) return fa - fb;
    let cmp = 0;
    if (key === "modified") cmp = (a.mtime || 0) - (b.mtime || 0);
    else if (key === "type") cmp = ext(a.name).localeCompare(ext(b.name), "de");
    else cmp = String(a.name).localeCompare(String(b.name), "de", { numeric: true, sensitivity: "base" });
    if (cmp === 0) cmp = String(a.name).localeCompare(String(b.name), "de", { numeric: true, sensitivity: "base" });
    return cmp * mul;
  });
}
function sortArrow(key) {
  if (evidenceSort.key !== key) return "";
  return evidenceSort.dir === "asc" ? " ▲" : " ▼";
}
function filterEvidenceRows() {
  const q = evidenceSearch.trim().toLowerCase();
  const pane = document.querySelector("#docs-detail-pane");
  if (!pane) return;
  let visible = 0, total = 0;
  pane.querySelectorAll(".evidence-file-row:not(.head)").forEach((row) => {
    total++;
    const name = (row.dataset.evidenceEntryName || "").toLowerCase();
    const hide = q && !name.includes(q);
    row.classList.toggle("evidence-row-hidden", !!hide);
    if (!hide) visible++;
  });
  const hint = pane.querySelector(".evidence-search-empty");
  if (hint) hint.classList.toggle("hidden", !(q && total && visible === 0));
}
// Spring-loaded folders: beim Draggen über einen Ordner verweilen → öffnet ihn.
let springTimer = null;
let springKey = null;
function clearSpring() {
  if (springTimer) { clearTimeout(springTimer); springTimer = null; }
  springKey = null;
}
function scheduleSpring(key, openFn) {
  if (springKey === key) return;       // gleiches Ziel → Timer weiterlaufen lassen
  clearSpring();
  springKey = key;
  springTimer = setTimeout(() => {
    springTimer = null; springKey = null;
    if (evidenceDrag) openFn();
  }, 650);
}

// ── Labels / helpers ─────────────────────────────────────────
const statusFlow = ["Open", "Done", "Not needed"];
const taskFlow = ["Open", "Done", "Blocked"];
const freigabeFlow = ["Offen", "Ja", "Nein"];

const statusLabels = {
  Available: "Vorhanden",
  Archived: "Archiviert",
  Blocked: "Blockiert",
  Current: "Aktuell",
  Done: "Erledigt",
  Draft: "Entwurf",
  Good: "Gut",
  "In progress": "In Arbeit",
  "In review": "In Prüfung",
  "Needs check": "Prüfen",
  "Not needed": "Nicht nötig",
  "Not started": "Nicht gestartet",
  Open: "Offen",
  Partial: "Teilweise",
  Planned: "Geplant",
  Reference: "Referenz",
  "Bei VDE": "Bei VDE",
  "Finale Dokumente fehlen": "Finale Dokumente fehlen",
  Started: "Gestartet",
  Watch: "Beobachten",
  Abgeschlossen: "Abgeschlossen",
  Ja: "Ja",
  Nein: "Nein",
  Offen: "Offen",
};

const termLabels = {
  Administration: "Administration",
  Approbation: "Approbation",
  Available: "Vorhanden",
  Blocked: "Blockiert",
  Brazil: "Brasilien",
  Build: "Build",
  Builds: "Builds",
  Certification: "Zulassung",
  Components: "Komponenten",
  Current: "Aktuell",
  Done: "Erledigt",
  Electronics: "Elektronik",
  Evidence: "Nachweis",
  Global: "Global",
  Open: "Offen",
  PCS: "PCS",
  Planned: "Geplant",
  Reference: "Referenz",
  Watch: "Beobachten"
};

const approbationText = {
  "IEC / VDE approval checklist by Ziffer. Status values are a first draft from the EF1157 folder and should be corrected as the team reviews evidence.":
    "IEC-/VDE-Checkliste nach Ziffer. Die Stati sind ein erster Entwurf aus dem EF1157-Ordner und sollen bei der fachlichen Prüfung korrigiert werden.",
  "General conditions / product scope": "Allgemeine Bedingungen / Produktscope",
  "Marking and instructions": "Kennzeichnung und Anleitungen",
  "Protection against live parts": "Schutz gegen Berührung spannungsführender Teile",
  "Input and current": "Leistungsaufnahme und Strom",
  Heating: "Erwärmung",
  "Leakage current / electric strength": "Ableitstrom / elektrische Festigkeit",
  "Moisture resistance": "Feuchtebeständigkeit",
  "Leakage after moisture": "Ableitstrom nach Feuchteprüfung",
  "Overload protection": "Überlastschutz",
  "Abnormal operation / motor heating": "Abnormaler Betrieb / Motorerwärmung",
  "Stability and mechanical hazards": "Standfestigkeit und mechanische Gefahren",
  Construction: "Konstruktion",
  Components: "Komponenten",
  "Supply connection / external cords": "Netzanschluss / externe Leitungen",
  "Earthing provision": "Schutzleiteranschluss",
  "Clearances / creepage distances": "Luft- und Kriechstrecken",
  "Resistance to heat and fire": "Beständigkeit gegen Wärme und Feuer",
  "Radiation / toxicity / similar hazards": "Strahlung / Toxizität / ähnliche Gefahren",
  "Project scope exists; final scope confirmation still useful.": "Produktscope vorhanden; finale Scope-Bestätigung noch sinnvoll.",
  "Need final rating plate and manual consistency check.": "Finaler Abgleich von Typenschild und Anleitung nötig.",
  "Needs explicit evidence set.": "Expliziter Nachweissatz nötig.",
  "Measurement files are present.": "Messdateien vorhanden.",
  "Folder contains heater/TCO measurement evidence.": "Ordner enthält Messnachweise zu Heizung/TCO.",
  "Need final result mapping.": "Finale Zuordnung der Resultate nötig.",
  "Check if test is done or not applicable.": "Prüfen, ob Test erledigt oder nicht anwendbar ist.",
  "Needs final report status.": "Finaler Berichtstatus nötig.",
  "Placeholder until standard matrix is reviewed.": "Platzhalter bis die Normmatrix geprüft ist.",
  "Folder explicitly contains Ziffer 19 measurement evidence.": "Ordner enthält explizite Messnachweise zu Ziffer 19.",
  "Need mechanical checklist.": "Mechanische Checkliste nötig.",
  "Needs final construction review.": "Finale Konstruktionsprüfung nötig.",
  "Many component certificates are present.": "Viele Komponentenzertifikate vorhanden.",
  "Cord approval folders are present.": "Nachweisordner für Netzkabel vorhanden.",
  "Need confirm appliance class and PE concept.": "Geräteklasse und Schutzleiterkonzept bestätigen.",
  "Needs PCB review record.": "PCB-Prüfnachweis nötig.",
  "Ziffer_30 folder exists.": "Ziffer-30-Ordner vorhanden.",
  "Likely not applicable, confirm in matrix.": "Voraussichtlich nicht anwendbar; in Matrix bestätigen."
};

function statusClass(value) {
  return String(value).toLowerCase().replaceAll(" ", "-").replaceAll("/", "-");
}
function statusLabel(value) { return statusLabels[value] || value; }
function termLabel(value) { return termLabels[value] || value; }
function approbationLabel(value) { return approbationText[value] || value; }
function freigabeStatus(value) {
  return freigabeFlow.includes(value) ? value : "Offen";
}
function evidenceCacheKey(projectId, groupName) {
  return `${projectId}:${groupName}`;
}
function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}


function applyTheme(theme) {
  document.body.dataset.theme = theme;
  document.documentElement.dataset.theme = theme;   // damit auch <html> die Theme-Farbe kennt
  localStorage.setItem("pcs-theme", theme);
  themeToggle.querySelector("strong").textContent = theme === "dark" ? "Light" : "Dark";
  themeToggle.setAttribute("aria-pressed", String(theme === "dark"));
}

applyTheme(localStorage.getItem("pcs-theme") || "light");

// ── API ──────────────────────────────────────────────────────
async function apiFetch(path, options = {}) {
  const res = await fetch(path, {
    headers: { "Content-Type": "application/json" },
    ...options
  });
  if (!res.ok) throw new Error(`API ${path}: ${res.status}`);
  return res.json();
}

function apiPut(path, body) {
  return apiFetch(path, { method: "PUT", body: JSON.stringify(body) });
}

async function reloadProject() {
  activeProject = await apiFetch(`/api/projects/${activeProject.id}`);
}

async function loadEvidenceEntries(group, browseHref) {
  evidenceSearch = "";   // Suche bei jedem Ordnerwechsel zurücksetzen
  const key = evidenceCacheKey(activeProject.id, group.primary);
  // Ohne expliziten Pfad: zuletzt besuchten Unterordner dieses Ordners wieder
  // öffnen (pro Ordner gemerkt), sonst die Wurzel. So bleibt man beim Wechsel
  // zwischen z.B. 09 und 10 in seinem jeweiligen Unterordner.
  const href = browseHref || lastFolderByGroup.get(key) || evidenceHref(group);
  const rootHref = evidenceHref(group);
  lastFolderByGroup.set(key, href);   // aktuellen Ordner pro Gruppe merken
  const pathKey = `${key}:${href}`;
  const cachedPath = evidencePathEntries.get(pathKey);
  if (cachedPath) {
    evidenceEntries.set(key, { ...cachedPath, loading: false, browseHref: href, rootHref });
    renderDocs(activeProject);
    return;
  }
  const prev = evidenceEntries.get(key);
  evidenceEntries.set(key, { loading: true, entries: prev?.entries || [], browseHref: href, rootHref });
  if (!prev?.entries?.length) renderDocs(activeProject);
  try {
    const data = await apiFetch(`/api/list-path?href=${encodeURIComponent(href)}`);
    const next = { loading: false, entries: data.entries || [], browseHref: href, rootHref };
    evidencePathEntries.set(pathKey, next);
    evidenceEntries.set(key, next);
  } catch (err) {
    evidencePathEntries.delete(pathKey);
    evidenceEntries.set(key, { loading: false, error: true, errorMessage: err.message, entries: [], browseHref: href, rootHref });
    console.error("Evidence list error:", err);
  }
  renderDocs(activeProject);
}

function toggleEvidenceGroup(groupName) {
  activeEvidenceGroup = activeEvidenceGroup === groupName ? null : groupName;
}

// ── View switching ───────────────────────────────────────────
function setView(name) {
  activeView = name;
  document.body.classList.toggle("dashboard-mode", name === "dashboard");
  Object.entries(views).forEach(([k, el]) => el.classList.toggle("hidden", k !== name));
  Object.entries(buttons).forEach(([k, btn]) => btn.classList.toggle("active", k === name));
  document.querySelector(".top-actions").classList.toggle("hidden", name === "dashboard");
  document.querySelector("#excel-sidebar").classList.toggle("hidden", name !== "docs");
  if (name !== "docs") document.querySelector("#excel-sidebar-list").classList.add("hidden");
  // In der Dateien-Ansicht scrollen die Bereiche intern → kein Seiten-Scrollen
  // nötig; das verhindert den leeren Streifen + Scrollbalken ganz unten.
  // (Die Seite scrollt auf <html>, nicht <body>.)
  document.documentElement.style.overflow = name === "docs" ? "hidden" : "";
  if (name === "docs") { requestAnimationFrame(fitDocsPanes); setTimeout(fitDocsPanes, 80); }
}

// Dateien-Bereiche dynamisch bis zum unteren Viewport-Rand füllen — statt einer
// festen calc(100vh - 270px)-Höhe, die je nach Bildschirm/Projekt einen leeren
// Streifen unten liess. Nur aktiv, wenn die Panes nebeneinander liegen.
function fitDocsPanes() {
  const two = document.querySelector(".docs-two-pane");
  const detail = document.querySelector("#docs-detail-pane");
  const grid = document.querySelector("#document-group-grid");
  if (!two || !detail || !grid || two.offsetParent === null) return;
  const sideBySide = Math.abs(detail.getBoundingClientRect().top - grid.getBoundingClientRect().top) < 5;
  if (!sideBySide) { detail.style.height = ""; grid.style.maxHeight = ""; return; }
  const h = Math.max(240, Math.round(window.innerHeight - two.getBoundingClientRect().top - 34));
  detail.style.height = `${h}px`;
  grid.style.maxHeight = `${h}px`;
}
window.addEventListener("resize", fitDocsPanes);

// ── Machine list ─────────────────────────────────────────────
function relatedProjects() {
  if (!activeProject || activeView === "dashboard") return projectList;
  if (activeProject.variant_group) {
    return projectList.filter((p) => p.variant_group === activeProject.variant_group);
  }
  // Group via variant_of: find the root, then show root + all its derivatives
  const rootId = activeProject.variant_of || activeProject.id;
  return projectList.filter((p) => p.id === rootId || p.variant_of === rootId);
}

function renderMachines() {
  const q = search.value.trim().toLowerCase();
  const source = activeView === "dashboard" ? projectList : relatedProjects();
  const filtered = source.filter((p) => {
    if (activeView === "dashboard" && !q) return true;
    return [p.id, p.name, p.owner, p.phase, p.market, p.variant_group, p.variant_of]
      .join(" ").toLowerCase().includes(q);
  });

  const IEC_STAGES = new Set(["PT1", "OOT", "TS1", "TS2", "TS3"]);
  const renderStages = (stages, p) =>
    stages.map((b) => `<span data-build-select="${b}" class="${b === activeBuild && p.id === activeProject?.id ? "active" : ""}">${b}</span>`).join("");

  const renderBtn = (p) => {
    const iec = (p.buildStages || []).filter((b) => IEC_STAGES.has(b));
    const ul  = (p.buildStages || []).filter((b) => !IEC_STAGES.has(b));
    const buildsHTML = (iec.length || ul.length) ? `
      <div class="sidebar-builds">
        ${iec.length ? `<span class="build-type-label">IEC</span>${renderStages(iec, p)}` : ""}
        ${iec.length && ul.length ? `<hr class="build-type-divider">` : ""}
        ${ul.length  ? `<span class="build-type-label">UL</span>${renderStages(ul, p)}` : ""}
      </div>` : "";
    return `
    <button class="${p.id === activeProject?.id ? "active" : ""}" data-id="${p.id}" type="button">
      <strong>${p.id}</strong>
      <em class="${statusClass(p.health)}">${statusLabel(p.health)}</em>
      <span>${p.market ? `${termLabel(p.market)} / ` : ""}${p.phase}</span>
      ${buildsHTML}
    </button>`;
  };

  machineList.innerHTML = filtered.map((p) => renderBtn(p)).join("");
}

const dashboardSearch = document.querySelector("#dashboard-search");
const recentlyOpenedEl = document.querySelector("#recently-opened");

function renderRecentlyOpened() {
  const recentIds = getRecentProjects().map((r) => r.id);
  const recent = recentIds.map((id) => projectList.find((p) => p.id === id)).filter(Boolean);
  recentlyOpenedEl.innerHTML = recent.length ? `
    <div class="recently-opened-row">
      <span class="recently-opened-label">Zuletzt geöffnet</span>
      ${recent.map((p) => `
        <button class="recent-card" type="button" data-dashboard-project="${p.id}">
          <strong>${p.id}</strong>
          <em class="${statusClass(p.health)}">${statusLabel(p.health)}</em>
          <span>${escapeHtml(p.family || p.name || "")}</span>
        </button>
      `).join("")}
    </div>
  ` : "";
}

function renderDashboard() {
  renderRecentlyOpened();
  const q = (dashboardSearch?.value || search.value).trim().toLowerCase();
  const projects = projectList.filter((p) =>
    [p.id, p.name, p.owner, p.family, p.market, p.phase, p.variant_group, p.variant_of, p.project_no, p.sw_version, p.hw_version, p.machine_type, p.machine_use]
      .join(" ").toLowerCase().includes(q)
  );
  document.querySelector("#project-family").textContent = "PCS Projektübersicht";
  document.querySelector("#project-title").textContent = "Kaffee Dashboard";
  document.querySelector("#dashboard-count").textContent = `${projects.length} Projekte`;
  document.querySelector("#dashboard-grid").innerHTML = projects.length ? projects.map((p) => `
    <button class="dashboard-card" type="button" data-dashboard-project="${p.id}">
      <div class="dashboard-card-top">
        <span class="dashboard-card-brand">
          ${escapeHtml(p.family || "PCS Maschine")}
          ${p.project_no ? `<em class="project-no-badge">${escapeHtml(p.project_no)}</em>` : `<em class="project-no-badge project-no-empty" title="Project No. eingeben">+ Nr.</em>`}
        </span>
      </div>
      <strong>${p.id}</strong>
      <p>${escapeHtml(p.name)}</p>
      <em class="${statusClass(p.health)}">${statusLabel(p.health)}</em>
      <small>${p.market ? `${termLabel(p.market)} / ` : ""}${p.phase || ""}</small>
      <div class="dashboard-card-versions">
        <span class="version-badge" data-version-field="sw" data-version-project="${p.id}" title="SW Version bearbeiten">
          <i>SW</i>${p.sw_version ? escapeHtml(p.sw_version) : `<em class="version-empty">—</em>`}
        </span>
        <span class="version-badge" data-version-field="hw" data-version-project="${p.id}" title="HW Version bearbeiten">
          <i>HW</i>${p.hw_version ? escapeHtml(p.hw_version) : `<em class="version-empty">—</em>`}
        </span>
      </div>
      <div class="dashboard-card-meta">
        <span class="meta-badge ${p.machine_type ? "meta-badge--type" : "meta-badge--empty"}" data-meta-field="machine_type" data-meta-project="${p.id}" title="Maschinentyp bearbeiten">
          ${p.machine_type ? escapeHtml(p.machine_type) : `<em class="version-empty">Typ?</em>`}
        </span>
        <span class="meta-badge ${p.machine_use === "Commercial Use" ? "meta-badge--commercial" : p.machine_use === "Private Use" ? "meta-badge--privat" : "meta-badge--empty"}" data-meta-field="machine_use" data-meta-project="${p.id}" title="Verwendung bearbeiten">
          ${p.machine_use ? escapeHtml(p.machine_use) : `<em class="version-empty">Verw.?</em>`}
        </span>
      </div>
    </button>
  `).join("") : `<p class="empty-state">Keine Projekte gefunden.</p>`;
}

function updateSidebarBuildSelection() {
  machineList.querySelectorAll("[data-build-select]").forEach((pill) => {
    const machine = pill.closest("button[data-id]");
    pill.classList.toggle("active", machine?.dataset.id === activeProject?.id && pill.dataset.buildSelect === activeBuild);
  });
}

// ── Product image ─────────────────────────────────────────────
function loadProductImage(projectId) {
  const n = projectId.replace(/\D/g, "");
  const img = document.querySelector("#product-image");
  const candidates = [
    `evidence-${n}/product.jpg`,
    `evidence-${n}/product.png`,
    `evidence-${n}/thumbnail.jpg`,
    `evidence-${n}/thumbnail.png`
  ];
  img.style.display = "none";
  img.src = "";
  let i = 0;
  function tryNext() {
    if (i >= candidates.length) return;
    const src = candidates[i++];
    fetch(src, { method: "HEAD" })
      .then((r) => { if (r.ok) { img.src = src; img.style.display = "block"; } else tryNext(); })
      .catch(tryNext);
  }
  tryNext();
}

// ── Summary cards ─────────────────────────────────────────────
function renderSummary(project) {
  const relation = project.variant_of ? ` / Variante von ${project.variant_of}` : "";
  document.querySelector("#project-family").textContent =
    `${project.family} / ${termLabel(project.market || "Global")}${relation}`;
  document.querySelector("#project-title").textContent = project.name;

  const archiveLoc = project.archive_location || "";
  document.querySelector("#summary-grid").innerHTML = `
    <article class="summary-card hero">
      <span>PCS Reifegrad</span>
      <strong>${project.progress}%</strong>
      <div><i style="width:${project.progress}%"></i></div>
      <p>${project.phase}</p>
    </article>
    <article class="summary-card archive-card">
      <span>Archiv</span>
      <strong class="archive-location-text ${archiveLoc ? "" : "muted"}">${archiveLoc ? archiveLoc.split("·").map(s => `<span>${escapeHtml(s.trim())}</span>`).join("") : "—"}</strong>
      <button class="archive-open-btn" type="button" data-archive-open="${escapeHtml(project.id)}" title="In Excel öffnen">↗ Excel</button>
    </article>
    ${(project.stats || []).filter(([title]) => title !== "Certification").map(([title, value, note]) => {
      const isSubtopic = Boolean(project.subtopics?.[title]);
      const tag = isSubtopic ? "button" : "article";
      const type = isSubtopic ? ' type="button"' : "";
      const dataset = isSubtopic ? ` data-subtopic="${title}"` : "";
      const cls = isSubtopic ? "summary-card clickable" : "summary-card";
      return `
        <${tag} class="${cls}"${type}${dataset}>
          <span>${termLabel(title)}</span>
          <strong>${value}</strong>
          <p>${note}</p>
        </${tag}>
      `;
    }).join("")}
  `;
}

// ── Overview panels ───────────────────────────────────────────
function renderOverview(project) {
  const freigabeIsJa = (label) =>
    (project.fachfreigabe?.gates || []).some((gate) => gate.label === label && gate.status === "Ja");
  const coveredByFreigabe = (item) => {
    const name = item.name || "";
    if (/EMC|ErP/i.test(name)) return freigabeIsJa("EMC / ErP Berichte akzeptiert");
    if (/Declaration|conformity|final/i.test(name)) return freigabeIsJa("Finale Dokumente vollständig");
    if (/IEC|60335|review/i.test(name)) return freigabeIsJa("Ziffern fachlich geprüft") || freigabeIsJa("VDE bestanden");
    if (/VDE/i.test(name)) return freigabeIsJa("VDE bestanden");
    return false;
  };
  const openCertifications = (project.certification || [])
    .filter((c) => !c.done && !["Done", "Not needed", "Abgeschlossen"].includes(c.state) && !coveredByFreigabe(c));
  const openBuilds = (project.builds || []).filter((b) => !["Done", "Not needed"].includes(b.state)).length;
  document.querySelector("#cert-progress").textContent = `${openCertifications.length} offen`;
  document.querySelector("#build-progress").textContent = `${openBuilds} offen`;
  document.querySelector("#risk-count").textContent = `${(project.risks || []).length} Risiken`;
  document.querySelector("#next-count").textContent =
    `${(project.tasks || []).filter((t) => t.status !== "Done").length} offen`;

  document.querySelector("#pcs-strip").classList.add("hidden");
  document.querySelector("#pcs-strip").innerHTML = "";

  document.querySelector("#cert-list").innerHTML = openCertifications.length
    ? openCertifications.map((item) => `
      <div>
        <b>Offen</b>
        <span>${item.name}</span>
        <em>${statusLabel(item.state)}</em>
      </div>
    `).join("")
    : `<p class="empty-state">Keine offenen Zulassungspunkte.</p>`;

  document.querySelector("#build-list").innerHTML = (project.builds || []).map((build) => `
    <article class="build-card">
      <time>${build.date}</time>
      <strong>${build.label}</strong>
      <em class="${statusClass(build.state)}">${statusLabel(build.state)}</em>
      <span>${build.samples || "Muster TBD"}</span>
      <p>${build.note}</p>
    </article>
  `).join("");

  document.querySelector("#risk-list").innerHTML = (project.risks || []).map((risk) => `
    <article class="${risk.level.toLowerCase()}">
      <strong>${risk.level}</strong>
      <p>${risk.text}</p>
    </article>
  `).join("");

  document.querySelector("#next-list").innerHTML = (project.tasks || [])
    .filter((t) => t.status !== "Done")
    .slice(0, 5)
    .map((task) => `
      <article>
        <strong>${task.task}</strong>
        <span>${termLabel(task.area)} / ${task.owner} / ${task.due}</span>
      </article>
    `).join("");
}

// ── Tasks table ───────────────────────────────────────────────
function renderBuildTags(value) {
  return String(value || "Alle").split(",").map((build) => build.trim()).filter(Boolean)
    .map((build) => `<em>${build}</em>`).join("");
}

function renderTasks(project) {
  const area = taskFilter.value;
  let tasks = area === "all"
    ? (project.tasks || [])
    : (project.tasks || []).filter((t) => t.area === area);
  if (activeBuild !== "Alle") {
    tasks = tasks.filter((task) => {
      const builds = String(task.builds || "Alle").split(",").map((b) => b.trim());
      return builds.includes("Alle") || builds.includes(activeBuild);
    });
  }
  const buildLabel = activeBuild !== "Alle" ? ` — ${activeBuild}` : "";
  document.querySelector("#task-title").textContent = `PCS Massnahmen${buildLabel}`;

  document.querySelector("#task-table").innerHTML = `
    <div class="row head">
      <span>Build</span><span>Bereich</span><span>PCS Massnahme</span>
      <span>Verantwortlich</span><span>Fällig</span><span>Status</span>
    </div>
    ${tasks.length ? tasks.map((task) => `
      <div class="row ${task.status === "Done" ? "row-done" : ""} ${task.status === "Blocked" ? "row-blocked" : ""}">
        <span class="task-builds">${renderBuildTags(task.builds)}</span>
        <span>${termLabel(task.area)}</span>
        <strong>${task.task}</strong>
        <span>${task.owner}</span>
        <span>${task.due}</span>
        <button class="status-toggle ${statusClass(task.status)}" type="button"
          data-task-id="${task.id}" aria-label="Status ändern">
          ${statusLabel(task.status)}
        </button>
      </div>
      ${task.status === "Blocked" ? `
        <div class="task-block-reason">
          <label for="block-reason-${task.id}">Blocker-Grund</label>
          <textarea id="block-reason-${task.id}" rows="2" data-task-block-reason="${task.id}"
            placeholder="Warum ist diese Massnahme blockiert?">${escapeHtml(task.block_reason)}</textarea>
        </div>
      ` : ""}
    `).join("") : `<p class="empty-state">Keine Massnahmen für ${activeBuild} in diesem Bereich.</p>`}
  `;
}

// ── Evidence cell ─────────────────────────────────────────────
function evidenceRelativePath(link) {
  if (!link.href?.startsWith(trackerConfig.localDocumentRoot)) return link.href || "";
  return link.href.replace(trackerConfig.localDocumentRoot, "");
}
function evidenceHref(link) {
  const rel = evidenceRelativePath(link);
  if (!rel) return "";
  if (/^(evidence-\d+|files)\//.test(rel)) return encodeURI(rel);
  if (trackerConfig.documentMode === "sharepoint") {
    return trackerConfig.sharePointRoot + rel.split("/").map(encodeURIComponent).join("/");
  }
  return encodeURI(trackerConfig.localDocumentRoot + rel);
}
function parentEvidenceHref(currentHref, rootHref) {
  const root = rootHref.endsWith("/") ? rootHref : `${rootHref}/`;
  const current = currentHref.endsWith("/") ? currentHref.slice(0, -1) : currentHref;
  if (current === root.slice(0, -1)) return null;
  const slash = current.lastIndexOf("/");
  const parent = slash >= 0 ? `${current.slice(0, slash)}/` : root;
  return parent.startsWith(root) ? parent : root;
}
function isWordFile(p) { return /\.(doc|docx|docm)$/i.test(p || ""); }
function isExcelFile(p) { return /\.(xls|xlsx|xlsm|xlsb|xlam)$/i.test(p || ""); }
function isOfficeFile(p) { return isWordFile(p) || isExcelFile(p); }
function officeApp(p) { return isExcelFile(p) ? "excel" : "word"; }
function officeEditHref(link) {
  const href = evidenceHref(link);
  if (!href || !isOfficeFile(href)) return "";
  const abs = href.startsWith("http") ? href : new URL(href, window.location.href).href;
  return isExcelFile(href) ? `ms-excel:ofe|u|${abs}` : `ms-word:ofe|u|${abs}`;
}
function shouldOpenOfficeLocally(link) {
  return trackerConfig.documentMode === "local" && isOfficeFile(evidenceHref(link));
}
function renderEvidenceCell(item) {
  if (!item.evidenceLinks?.length) return item.evidence || "";
  return `
    <div class="evidence-stack">
      <span>${item.evidence || ""}</span>
      ${item.evidenceLinks.map((link) => {
        if (shouldOpenOfficeLocally(link)) {
          return `<button class="inline-open-action" type="button" data-open-office-href="${evidenceHref(link)}">${link.label}</button>`;
        }
        const href = officeEditHref(link) || evidenceHref(link);
        return `<a href="${href}" target="_blank" rel="noreferrer">${link.label}</a>`;
      }).join("")}
    </div>
  `;
}

// ── Ziffer checklist ──────────────────────────────────────────
const BUILD_ORDER = ["PT1", "OOT", "TS1", "TS2", "TS3"];

function renderSubtopic(project) {
  const subtopic = project.subtopics?.[activeSubtopic];
  const table = document.querySelector("#ziffer-table");
  const buildLabel = activeBuild !== "Alle" ? ` — ${activeBuild}` : "";
  document.querySelector("#subtopic-title").textContent = `PCS Approbation / Ziffer-Checkliste${buildLabel}`;

  if (!subtopic) {
    document.querySelector("#subtopic-summary").textContent = "Für diesen Bereich ist noch keine Checkliste angelegt.";
    table.innerHTML = "";
    return;
  }

  document.querySelector("#subtopic-summary").textContent = approbationLabel(subtopic.summary);

  // Filter by build then by status
  let ziffern = subtopic.ziffern || [];
  if (activeBuild !== "Alle") {
    ziffern = ziffern.filter((z) => {
      const bl = (z.builds || "").split(",").map((s) => s.trim());
      return bl.includes(activeBuild) || bl.includes("Alle");
    });
  }
  const statusF = subtopicFilter.value;
  if (statusF !== "all") ziffern = ziffern.filter((z) => z.status === statusF);

  // Reports below ziffer table
  const reportsPanel = document.querySelector("#subtopic-reports-panel");
  if (reportsPanel) {
    const reportBuildMatches = (report) => {
      if (activeBuild === "Alle") return true;
      const build = String(report.build || "").toLowerCase();
      const file = String(report.file || "").toLowerCase();
      const selected = activeBuild.toLowerCase();
      if (build.includes(selected) || file.includes(selected)) return true;
      if (activeBuild === "PT1" && (build.includes("pt /") || build.includes("pre-approval") || build.includes("prototype"))) return true;
      return false;
    };
    const reports = (project.reportVersions || []).filter(reportBuildMatches);
    reportsPanel.innerHTML = reports.length ? `
      <div class="table-header">
        <h2>PCS Approbationsberichte${activeBuild !== "Alle" ? ` — ${activeBuild}` : ""}</h2>
      </div>
      <div class="report-table">
        <div class="report-row head">
          <span>Projekt</span><span>Build</span><span>Version</span>
          <span>Geändert</span><span>Status</span><span>Datei</span>
        </div>
        ${reports.map((r) => `
          <div class="report-row">
            <strong>${r.project || r.project_id || ""}</strong>
            <span>${r.build}</span><span>${r.version}</span>
            <span>${r.modified}<br>${r.size}</span>
            <em class="${statusClass(r.state)}">${statusLabel(r.state)}</em>
            <span class="report-file">
              ${isOfficeFile(r.file)
                ? `<button class="word-action" type="button" data-open-office-href="${evidenceHref(r)}">${r.file}</button>`
                : `<a href="${evidenceHref(r)}" target="_blank" rel="noreferrer">${r.file}</a>`}
            </span>
          </div>
        `).join("")}
      </div>
    ` : "";
  }

  table.innerHTML = `
    <div class="ziffer-row head">
      <span>Z.</span><span>Thema</span><span>Status</span>
    </div>
    ${ziffern.map((item) => `
      <div class="ziffer-row ${item.status === "Open" ? "ziffer-row-open" : ""} ${item.status === "Done" ? "ziffer-row-done" : ""} ${item.status === "Not needed" ? "ziffer-row-muted" : ""}">
        <span class="ziffer-number">${item.nr}</span>
        <span class="ziffer-topic">
          <strong>${approbationLabel(item.title)}</strong>
          ${item.note ? `<small>${approbationLabel(item.note)}</small>` : ""}
          ${item.status === "Not needed" ? `<textarea class="ziffer-reason" data-reason-nr="${item.nr}" placeholder="Warum nicht nötig?" rows="1">${escapeHtml(item.not_needed_reason || "")}</textarea>` : ""}
        </span>
        <button class="status-toggle ${statusClass(item.status)}" type="button"
          data-nr="${item.nr}" aria-label="Status ändern">
          ${statusLabel(item.status)}
        </button>
      </div>
    `).join("")}
  `;
}

// ── Docs / Evidence ───────────────────────────────────────────
function renderDocs(project) {
  const allGroups = (project.documentGroups || [])
    .slice()
    .sort((a, b) => String(a.primary || "").localeCompare(String(b.primary || "")));

  // Detect unique markets (prefix before " / " in area name).
  // Only treat a prefix as a real market if it looks like a market code (IEC, UL, EU, US, etc.)
  // — not IEC_FOLDER_MAP values like "Standards / Changes" or "Manual / Labels".
  function isMarketCode(s) {
    return /^(IEC|UL|EU|US|JP|CN|AU|IN|MX|TW|BR|KR|CH|DE|FR)([\s,]|$)/.test(s);
  }
  const markets = [...new Set(
    allGroups.map(g => {
      if (!g.area?.includes(" / ")) return null;
      const prefix = g.area.split(" / ")[0];
      return isMarketCode(prefix) ? prefix : null;
    }).filter(Boolean)
  )];

  // Auto-select first market if none active and markets exist
  if (markets.length > 0 && (!activeMarketFilter || !markets.includes(activeMarketFilter))) {
    activeMarketFilter = markets[0];
  }

  const groups = activeMarketFilter
    ? allGroups.filter(g => {
        if (!g.area?.includes(" / ")) return true;
        const prefix = g.area.split(" / ")[0];
        if (!isMarketCode(prefix)) return true; // IEC_FOLDER_MAP area — always show
        return g.area.startsWith(activeMarketFilter + " / ");
      })
    : allGroups;

  // Immer einen Ordner ausgewählt halten — die linke Seite soll nie leer sein.
  // Ist nichts (mehr) ausgewählt (z.B. beim Öffnen oder nach Markt-Wechsel),
  // automatisch den ersten Ordner mit Dateien wählen.
  const activeStillValid = activeEvidenceGroup && groups.some(g => g.primary === activeEvidenceGroup);
  if (!activeStillValid && groups.length) {
    const firstReal = groups.find(g => (parseInt(g.count) || 0) > 0) || groups[0];
    if (firstReal) {
      activeEvidenceGroup = firstReal.primary;
      loadEvidenceEntries(firstReal);   // merkt sich den zuletzt besuchten Unterordner
    }
  }

  const groupDetailMarkup = (group) => {
    const key = evidenceCacheKey(project.id, group.primary);
    const cached = evidenceEntries.get(key);
    const entries = cached?.entries || [];
    const isSubfolder = cached?.browseHref && cached.browseHref !== cached?.rootHref;
    const currentHref = cached?.browseHref || evidenceHref(group);
    const rootHref = cached?.rootHref || evidenceHref(group);
    const parentHref = isSubfolder ? parentEvidenceHref(currentHref, rootHref) : null;
    // Hrefs sind inkonsistent kodiert: evidenceHref() nutzt encodeURI (Leerzeichen
    // → %20) und liefert je nach Modus teils SOGAR doppelt kodierte Pfade
    // (z.B. "01%2520Administration" = "%20" nochmals kodiert), während die vom Server
    // gelieferten entry.href roh sind. Daher beide Seiten mehrfach voll dekodieren —
    // sonst greift das Prefix-Stripping nicht und der ganze Pfad bzw. ein kodierter
    // Ordnername (z.B. "08%20Schema%20Gera%CC%88t") landet im Breadcrumb/Titel.
    const decodeSafe = (s) => {
      let v = String(s || "");
      for (let i = 0; i < 5; i++) {
        try { const d = decodeURIComponent(v); if (d === v) break; v = d; } catch { break; }
      }
      return v;
    };
    const decRoot = decodeSafe(rootHref).replace(/\/$/, "");
    const decCurrent = decodeSafe(currentHref).replace(/\/$/, "");
    const relPath = decCurrent.startsWith(decRoot)
      ? decCurrent.slice(decRoot.length).replace(/^\//, "")
      : decCurrent;
    const pathParts = relPath ? relPath.split("/").filter(Boolean) : [];
    const locationTitle = pathParts.length ? pathParts[pathParts.length - 1] : group.primary;
    // Breadcrumb-Pfad: Wurzel (= Ordnername) › Unterordner › … — alle ausser dem
    // letzten sind anklickbar. Wurzel nutzt die exakte rootHref (gleiche Kodierung),
    // damit isSubfolder korrekt false wird; Zwischenebenen den dekodierten Pfad
    // (der Server dekodiert ohnehin).
    const crumbs = [{ label: group.primary, href: rootHref }];
    let crumbAcc = decRoot;
    pathParts.forEach((seg) => { crumbAcc = `${crumbAcc}/${seg}`; crumbs.push({ label: seg, href: `${crumbAcc}/` }); });
    const breadcrumbMarkup = `
      <nav class="evidence-breadcrumb" aria-label="Pfad">
        ${crumbs.map((c, i) => i === crumbs.length - 1
          ? `<span class="evidence-crumb current">${escapeHtml(c.label)}</span>`
          : `<button type="button" class="evidence-crumb" data-evidence-crumb="${escapeHtml(c.href)}" data-evidence-crumb-group="${escapeHtml(group.primary)}">${escapeHtml(c.label)}</button><span class="evidence-crumb-sep">›</span>`
        ).join("")}
      </nav>`;
    const selectedInThisFolder = entries.some((entry) => selectedEvidenceHrefs.has(entry.href));
    const selectedCountInThisFolder = entries.filter((entry) => selectedEvidenceHrefs.has(entry.href)).length;
    const pasteDisabled = evidenceClipboard ? "" : "disabled";
    const selectedDisabled = selectedInThisFolder ? "" : "disabled";
    const singleSelectedDisabled = selectedCountInThisFolder === 1 ? "" : "disabled";
    const backBtn = isSubfolder
      ? `<button class="evidence-back-btn" type="button" data-evidence-back="${group.primary}">← Zurück</button>` : "";
    const parentDropZone = parentHref
      ? `<div class="evidence-parent-drop" data-parent-drop-href="${escapeHtml(parentHref)}">
          <strong>In übergeordneten Ordner verschieben</strong>
          <span>${pathParts.length > 1 ? escapeHtml(pathParts[pathParts.length - 2]) : escapeHtml(group.primary)}</span>
        </div>`
      : "";
    const fileListMarkup = cached?.loading
      ? `<div class="evidence-crumb-bar">${backBtn}${breadcrumbMarkup}</div>${parentDropZone}<p class="empty-state">Ordnerinhalt wird geladen...</p>`
      : cached?.error
        ? `<div class="evidence-crumb-bar">${backBtn}${breadcrumbMarkup}</div>${parentDropZone}<p class="empty-state">Ordnerinhalt konnte nicht gelesen werden.
            <button class="finder-action" type="button" data-evidence-retry="${group.primary}">Erneut laden</button>
            ${cached.errorMessage ? `<small>${escapeHtml(cached.errorMessage)}</small>` : ""}
          </p>`
        : entries.length ? `
          <div class="evidence-crumb-bar">${backBtn}${breadcrumbMarkup}</div>
          <div class="evidence-location-bar">
            <button class="finder-action" type="button" data-file-action="mkdir" data-current-href="${escapeHtml(currentHref)}">Neuer Ordner</button>
            <button class="finder-action" type="button" data-file-action="rename" ${singleSelectedDisabled}>Umbenennen</button>
            <button class="finder-action" type="button" data-file-action="copy" ${selectedDisabled}>Kopieren</button>
            <button class="finder-action" type="button" data-file-action="cut" ${selectedDisabled}>Ausschneiden</button>
            <button class="finder-action" type="button" data-file-action="paste" data-current-href="${escapeHtml(currentHref)}" ${pasteDisabled}>Einfügen</button>
            <button class="finder-action danger" type="button" data-file-action="delete" ${selectedDisabled}>Löschen</button>
          </div>
          <div class="evidence-search-bar">
            <input type="search" class="evidence-search" data-evidence-search placeholder="Im Ordner suchen…" value="${escapeHtml(evidenceSearch)}" autocomplete="off">
          </div>
          ${parentDropZone}
          <div class="evidence-file-row head">
            <button type="button" class="evidence-sort${evidenceSort.key === "name" ? " active" : ""}" data-sort="name">Name${sortArrow("name")}</button>
            <button type="button" class="evidence-sort${evidenceSort.key === "type" ? " active" : ""}" data-sort="type">Typ${sortArrow("type")}</button>
            <button type="button" class="evidence-sort${evidenceSort.key === "modified" ? " active" : ""}" data-sort="modified">Geändert${sortArrow("modified")}</button>
            <span>Aktion</span>
          </div>
          <p class="evidence-search-empty empty-state hidden">Keine Treffer.</p>
          ${sortEvidenceEntries(entries).map((entry) => `
            <div class="evidence-file-row${entry.type === "Ordner" ? " evidence-file-row--folder" : ""}${selectedEvidenceHrefs.has(entry.href) ? " explorer-selected" : ""}"
              draggable="true"
              data-evidence-entry-href="${entry.href}"
              data-evidence-entry-name="${escapeHtml(entry.name)}"
              data-evidence-entry-type="${entry.type}"
              data-evidence-entry-group="${group.primary}">
              ${entry.type === "Ordner"
                ? entry.empty
                  ? `<span class="evidence-folder-btn evidence-folder-empty"><span class="folder-ico">📁</span> ${entry.name} <em>leer</em></span>`
                  : `<span class="evidence-folder-btn"><span class="folder-ico">📂</span> ${entry.name} <em class="folder-count">${entry.childCount} Dateien</em></span>`
                : `<div class="file-name-cell">${fileThumbHtml(entry)}${entry.type === "Datei" && isOfficeFile(entry.name)
                    ? `<span class="word-action word-action--name">${entry.name}</span>`
                    : `<span class="evidence-file-name">${entry.name}</span>`}</div>`}
              <span>${entry.size || entry.type}</span>
              <span>${entry.modified}</span>
              <span class="evidence-file-actions">
                ${entry.type !== "Ordner" ? `<button class="finder-action" type="button" data-open-href="${entry.href}">Öffnen</button>` : `<button class="finder-action" type="button" data-open-href="${entry.href}">Finder</button>`}
              </span>
            </div>
          `).join("")}
        ` : `<div class="evidence-crumb-bar">${backBtn}${breadcrumbMarkup}</div>${parentDropZone}<p class="empty-state">Dieser Ordner enthält keine sichtbaren Dateien.</p>`;

    return `
      <section class="document-group-detail">
        <div class="evidence-file-table">
          ${fileListMarkup}
          <div class="evidence-clear-space" data-clear-evidence-selection aria-hidden="true"></div>
        </div>
      </section>
    `;
  };

  // Market filter bar (IEC / UL / …) — only shown when a project spans >1 market
  const filterBar = document.querySelector("#market-filter-bar");
  if (filterBar) {
    if (markets.length > 1) {
      filterBar.innerHTML = markets.map(m =>
        `<button class="market-filter-btn ${activeMarketFilter === m ? "active" : ""}" data-market="${m}">${m}</button>`
      ).join("");
      filterBar.style.display = "flex";
    } else {
      filterBar.innerHTML = "";
      filterBar.style.display = "none";
    }
  }

  document.querySelector("#document-group-grid").innerHTML = groups.length
    ? groups.map((group) => {
        const num = parseInt(group.count) || 0;
        const folderLabel = group.primary || termLabel(group.area);
        const isEmpty = num === 0;

        const isSelected = activeEvidenceGroup === group.primary;
        return `
          <article class="document-group-card ${statusClass(group.status)} ${isSelected ? "selected" : ""} ${isEmpty ? "dgc-empty" : ""}"
            data-group-drop-href="${escapeHtml(evidenceHref(group))}"
            ${isEmpty ? "" : `data-evidence-group-card="${group.primary}" role="button" tabindex="0"`}>
            <div class="dgc-top">
              <strong>${folderLabel}</strong>
              <em class="status-toggle ${statusClass(group.status)}">${statusLabel(group.status)}</em>
            </div>
            <span class="dgc-area">${termLabel(group.area)}</span>
            <div class="dgc-count">${isEmpty ? `<em class="dgc-leer">leer</em>` : `${num}<span>Dateien</span>`}</div>
            ${isEmpty ? "" : `
            <div class="dgc-actions">
              <button class="finder-action" type="button" data-evidence-group="${group.primary}">${isSelected ? "Details ausblenden" : "Details anzeigen"}</button>
              <a class="dgc-link" href="${evidenceHref(group)}" target="_blank" rel="noreferrer">Ordner anzeigen →</a>
              <button class="finder-action" type="button" data-open-href="${evidenceHref(group)}">Im Finder öffnen</button>
            </div>`}
          </article>`;
      }).join("")
    : `<p class="empty-state">Für dieses Projekt ist noch kein PCS Nachweisindex angelegt.</p>`;

  // Detail pane (left): contents/subfolders of the currently selected top folder.
  const detailPane = document.querySelector("#docs-detail-pane");
  if (detailPane) {
    const selectedGroup = groups.find((g) => g.primary === activeEvidenceGroup);
    detailPane.innerHTML = selectedGroup
      ? groupDetailMarkup(selectedGroup)
      : `<p class="empty-state">Wähle rechts einen Ordner, um den Inhalt hier zu sehen.</p>`;
    if (evidenceSearch) filterEvidenceRows();   // aktive Suche nach Re-Render erneut anwenden
  }

  document.querySelector("#reports-panel")?.classList.add("hidden");
  window._observePdfThumbs?.();
  requestAnimationFrame(fitDocsPanes);
  setTimeout(fitDocsPanes, 80);
}

function currentEvidenceContext() {
  if (!activeProject || !activeEvidenceGroup) return null;
  const group = activeProject.documentGroups?.find((item) => item.primary === activeEvidenceGroup);
  if (!group) return null;
  const key = evidenceCacheKey(activeProject.id, group.primary);
  const cached = evidenceEntries.get(key);
  return {
    group,
    key,
    currentHref: cached?.browseHref || evidenceHref(group),
    selectedHref: selectedEvidenceHref,
    selectedHrefs: [...selectedEvidenceHrefs],
  };
}

function clearEvidenceSelection() {
  if (!selectedEvidenceHref && !selectedEvidenceHrefs.size && !evidenceSelectionAnchor) return false;
  selectedEvidenceHref = null;
  selectedEvidenceHrefs = new Set();
  evidenceSelectionAnchor = null;
  return true;
}

function clearEvidenceFolderCache(group, href) {
  const key = evidenceCacheKey(activeProject.id, group.primary);
  for (const cacheKey of [...evidencePathEntries.keys()]) {
    if (cacheKey.startsWith(`${key}:`)) evidencePathEntries.delete(cacheKey);
  }
  evidenceEntries.delete(key);
}

async function refreshEvidenceFolder(group, href) {
  clearEvidenceSelection();
  clearEvidenceFolderCache(group, href);
  await loadEvidenceEntries(group, href);
}

async function runFileAction(action, currentHref) {
  const ctx = currentEvidenceContext();
  if (!ctx) return;
  const folderHref = currentHref || ctx.currentHref;
  const selected = ctx.selectedHrefs || [];

  try {
    if (action === "copy" || action === "cut") {
      if (!selected.length) return;
      evidenceClipboard = { mode: action, hrefs: selected };
      renderDocs(activeProject);
      return;
    }

    if (action === "mkdir") {
      const name = prompt("Name für neuen Ordner:");
      if (!name) return;
      await apiFetch("/api/files/mkdir", {
        method: "POST",
        body: JSON.stringify({ parentHref: folderHref, name })
      });
      await refreshEvidenceFolder(ctx.group, folderHref);
      return;
    }

    if (action === "rename") {
      if (selected.length !== 1) return;
      const sourceHref = selected[0];
      const selectedRow = document.querySelector(`[data-evidence-entry-href="${CSS.escape(sourceHref)}"]`);
      const currentName = selectedRow?.dataset.evidenceEntryName
        || decodeURIComponent(sourceHref.split("/").filter(Boolean).pop() || "");
      const newName = prompt("Neuer Name:", currentName);
      if (newName === null) return;
      const trimmedName = newName.trim();
      if (!trimmedName || trimmedName === currentName) return;
      await apiFetch("/api/files/rename", {
        method: "POST",
        body: JSON.stringify({ href: sourceHref, newName: trimmedName })
      });
      await refreshEvidenceFolder(ctx.group, folderHref);
      return;
    }

    if (action === "paste") {
      if (!evidenceClipboard) return;
      const endpoint = evidenceClipboard.mode === "cut" ? "/api/files/move" : "/api/files/copy";
      for (const href of evidenceClipboard.hrefs || []) {
        await apiFetch(endpoint, {
          method: "POST",
          body: JSON.stringify({ sourceHref: href, destHref: folderHref })
        });
      }
      if (evidenceClipboard.mode === "cut") evidenceClipboard = null;
      await refreshEvidenceFolder(ctx.group, folderHref);
      return;
    }

    if (action === "delete") {
      if (!selected.length) return;
      if (!confirm(`${selected.length} Element(e) in den Papierkorb verschieben?`)) return;
      for (const href of selected) {
        await apiFetch("/api/files/delete", {
          method: "POST",
          body: JSON.stringify({ href })
        });
      }
      await refreshEvidenceFolder(ctx.group, folderHref);
    }
  } catch (err) {
    alert(`Dateiaktion fehlgeschlagen: ${err.message}`);
    console.error("File action error:", err);
  }
}

async function moveEvidenceEntriesToFolder(sourceHrefs, targetHref) {
  const ctx = currentEvidenceContext();
  if (!ctx || !targetHref || !sourceHrefs?.length) return;

  const uniqueSources = [...new Set(sourceHrefs)].filter((href) => href && href !== targetHref);
  if (!uniqueSources.length) return;

  try {
    for (const href of uniqueSources) {
      await apiFetch("/api/files/move", {
        method: "POST",
        body: JSON.stringify({ sourceHref: href, destHref: targetHref })
      });
    }
    // Ordner-Cache leeren, damit sowohl der Quell- als auch der Ziel-Ordner beim
    // nächsten Öffnen frisch geladen werden (sonst zeigt das Ziel den alten Stand).
    evidencePathEntries.clear();
    await refreshEvidenceFolder(ctx.group, ctx.currentHref);
  } catch (err) {
    alert(`Verschieben fehlgeschlagen: ${err.message}`);
    console.error("Drag move error:", err);
  }
}

// ── Fachfreigabe ──────────────────────────────────────────────
function freigabeGesamtstatus(gates) {
  if (!gates.length) return "Not started";
  const statuses = gates.map((g) => freigabeStatus(g.status));
  if (statuses.every((s) => s === "Ja")) return "Abgeschlossen";
  if (statuses.some((s) => s === "Nein")) return "Blocked";
  if (statuses.every((s) => s === "Offen")) return "Not started";
  return "In progress";
}

function renderFachfreigabe(project) {
  const gates = project.fachfreigabe?.gates || [];
  const meta = project.fachfreigabe?.meta || {};
  const gesamtstatus = freigabeGesamtstatus(gates);

  document.querySelector("#freigabe-container").innerHTML = `
    <div class="table-panel freigabe-panel">
      <div class="freigabe-header">
        <div>
          <h2>Fachliche Freigabe</h2>
          <p>Manuelle Bestätigung durch PCS-Chef / Verantwortliche nach Rücksprache mit VDE, Approbation oder Projektleitung. Diese Punkte können nicht automatisch erkannt werden.</p>
        </div>
        <em class="${statusClass(gesamtstatus)} freigabe-badge">${statusLabel(gesamtstatus)}</em>
      </div>
      <h3 class="freigabe-section-title">Bestätigung der Abschluss-Kriterien</h3>
      <div class="freigabe-gates">
        ${gates.length ? gates.map((gate) => `
          <div class="freigabe-gate-row">
            <span class="freigabe-gate-label">${gate.label}</span>
            <button class="freigabe-btn ${statusClass(freigabeStatus(gate.status))}" type="button"
              data-label="${gate.label}">${statusLabel(freigabeStatus(gate.status))}</button>
          </div>
        `).join("") : `<p class="empty-state">Keine Freigabe-Kriterien für dieses Projekt definiert.</p>`}
      </div>
      <div class="freigabe-divider"></div>
      <h3 class="freigabe-section-title">Rücksprache &amp; Dokumentation</h3>
      <div class="freigabe-meta">
        <div class="freigabe-field">
          <label for="ff-bestaetigt">Bestätigt von</label>
          <input id="ff-bestaetigt" type="text" placeholder="Name / Funktion" value="${meta.confirmed_by || ""}">
        </div>
        <div class="freigabe-field">
          <label for="ff-datum">Datum</label>
          <input id="ff-datum" type="date" value="${meta.datum || ""}">
        </div>
        <div class="freigabe-field freigabe-field-wide">
          <label for="ff-notiz">Notiz / Rücksprache</label>
          <textarea id="ff-notiz" rows="4"
            placeholder="z.B. VDE Rückfragen per E-Mail beantwortet am 15.05.2026, finaler Bericht erwartet bis 30.05.2026.">${meta.notiz || ""}</textarea>
        </div>
      </div>
    </div>
  `;
}

// ── Calculations ─────────────────────────────────────────────
const calcColumns = "ABCDEFGHIJKLMNOP".split("");
const calcStorageKey = "pcs-calculation-overrides";
const warmthStorageKey = "pcs-warmth-overrides";
const calculationWorkbooks = [
  {
    id: "ls",
    title: "Berechnung LS-Aufnahme",
    source: "berechnung_ls-aufnahme.xlsx",
    sheets: [{
      id: "uebersicht",
      title: "Übersicht",
      cols: 11,
      rows: {
        1: [null, null, "Messung 1", null, "Messung 2", null, "Messung 3", null, "Messung 4"],
        2: ["Abs. 10 ", "Leistungsaufnahme"],
        3: [null, "einzustellende Spannung : ", "=(C4+C5)/2", "Volt", "=(E4+E5)/2", "Volt", "=(G4+G5)/2", "Volt", "=(I4+I5)/2", "Volt"],
        4: [null, "Min. Bemessungsspannung: ", 220, "Volt", null, "Volt", null, "Volt", null, "Volt"],
        5: [null, "Max. Bemessungsspannung: ", 240, "Volt", null, "Volt", null, "Volt", null, "Volt"],
        6: [null, "Bemessungsaufnahme: ", 1380, "Watt", null, "Watt", null, "Watt", null, "Watt"],
        7: [null, "Messwert Leistungsaufnahme: ", 1380, "Watt", null, "Watt", null, "Watt", null, "Watt"],
        9: [null, "Abweichung:", "=100/C6*C7-100", "%", "=100/E6*E7-100", "%", "=100/G6*G7-100", "%", "=100/I6*I7-100", "%"],
        10: ["Abs. 11.8", "Erwärmung"],
        11: [null, "1,15-fache Bemessungsaufnahme: ", "=1.15*C6*((C5/C3)*(C5/C3))", "Watt", "=1.15*E6*((E5/E3)*(E5/E3))", "Watt", "=1.15*G6*((G5/G3)*(G5/G3))", "Watt", "=1.15*I6*((I5/I3)*(I5/I3))", "Watt"],
        12: [null, null, "=SQRT(((C3*C3)/C7)*C11)", "Volt", "=SQRT(((E3*E3)/E7)*E11)", "Volt", "=SQRT(((G3*G3)/G7)*G11)", "Volt", "=SQRT(((I3*I3)/I7)*I11)", "Volt"],
        13: ["Abs. 11.Zxx", "Berührbare Oberflächen"],
        14: [null, "1,0-fache Bemessungsaufnahme: ", "=C6*((C5/C3)*(C5/C3))", "Watt", "=E6*((E5/E3)*(E5/E3))", "Watt", "=G6*((G5/G3)*(G5/G3))", "Watt", "=I6*((I5/I3)*(I5/I3))", "Watt"],
        15: [null, null, "=SQRT(((C3*C3)/C7)*C14)", "Volt", "=SQRT(((E3*E3)/E7)*E14)", "Volt", "=SQRT(((G3*G3)/G7)*G14)", "Volt", "=SQRT(((I3*I3)/I7)*I14)", "Volt"],
        16: ["Abs. 16", "Ableitstrom"],
        17: [null, "1,06-fache Bemessungsspannung: ", "=C5*1.06", "Volt", "=E5*1.06", "Volt", "=G5*1.06", "Volt", "=I5*1.06", "Volt"],
        18: ["Abs. 19", "Unsachgemäßer Gebrauch"],
        19: [null, "0,85-fache Bemessungsaufnahme: ", "=0.85*C6*((C4/C3)*(C4/C3))", "Watt", "=0.85*E6*((E4/E3)*(E4/E3))", "Watt", "=0.85*G6*((G4/G3)*(G4/G3))", "Watt", "=0.85*I6*((I4/I3)*(I4/I3))", "Watt"],
        20: [null, "bei: ", "=SQRT(((C3*C3)/C7)*C19)", "Volt", "=SQRT(((E3*E3)/E7)*E19)", "Volt", "=SQRT(((G3*G3)/G7)*G19)", "Volt", "=SQRT(((I3*I3)/I7)*I19)", "Volt"],
        22: [null, "1,24-fache Bemessungsaufnahme: ", "=1.24*C6*((C5/C3)*(C5/C3))", "Watt", "=1.24*E6*((E5/E3)*(E5/E3))", "Watt", "=1.24*G6*((G5/G3)*(G5/G3))", "Watt", "=1.24*I6*((I5/I3)*(I5/I3))", "Watt"],
        23: [null, "bei: ", "=SQRT(((C3*C3)/C7)*C22)", "Volt", "=SQRT(((E3*E3)/E7)*E22)", "Volt", "=SQRT(((G3*G3)/G7)*G22)", "Volt", "=SQRT(((I3*I3)/I7)*I22)", "Volt"],
        24: ["Abs. 25", "Netzanschluß"],
        25: [null, "Berechneter Bemessungsstrom :", "=C6/C3+0.045", "Ampere", "=E6/E3+0.045", "Ampere", "=G6/G3+0.045", "Ampere", "=I6/I3+0.045", "Ampere"],
        27: ["Berechnung der 1,15-fachen Bemessungsaufnahme und die dafür benötigte Spannung:"],
        28: ["Annahme: min. Bemessungsspannung = 220V; max. Bemessungsspannung = 240V; einzustellende Spannung = 230V\nDer Widerstandswert für die Berechnung der 1,15-fachen Bemessungsaufnahme ist ein theoretischer Wert, der über die Bemessungsaufnahme bzw. -spannung berechnet wird!\nDer Widerstandswert für die Berechnung der Spannung, die für die 1,15-fache Bemessungsaufnahme benötigt wird, bezieht sich auf die tatsächlich aufgenommene Leistung des Gerätes."],
        38: [null, null, null, null, null, null, null, null, null, "Version 20150730"]
      }
    }]
  },
  {
    id: "ziff",
    title: "Formeln Ziffer 11 / 19",
    source: "formeln__ziff_11_19.xls",
    sheets: [
      {
        id: "tabelle1",
        title: "Tabelle1",
        cols: 16,
        rows: {
          3: [null, "Spannung", "Nennleistung", "gemessene Leistung", "Toleranz", "Pumpe"],
          4: [null, 220, 1450, 1310, "=ROUND(D4*100/C4,1)"],
          5: [null, 230, 1450, 1435, "=ROUND(D5*100/C5,1)"],
          6: [null, 240, 1450, 1562, "=ROUND(D6*100/C6,1)"],
          9: ["0026", null, null, 1210],
          10: ["0028", null, null, 1244],
          11: ["0029", null, null, 1223],
          12: ["0030", null, null, 1200],
          13: [null, null, null, "=SUM(D9:D12)"],
          14: [null, null, null, "=D13/4"],
          16: [null, "Spannung", "Nennleistung", "gemessene Leistung", "Toleranz", "Pumpe", "x 1,15", "effektiv", "x1,24", "x0,85", "Spannung x1,06"],
          17: [null, 220, 1299, 1134, "=(D17*100/C17)-100", null, "=C17*1.15", null, "=C17*1.24", "=C17*0.85", "=B17*1.06"],
          18: [null, 230, 1415, 1280, "=D18*100/C18-100", null, "=C18*1.15", null, "=C18*1.24", "=C18*0.85", "=B18*1.06"],
          19: [null, 240, 1300, 1360, "=D19*100/C19-100", null, "=C19*1.15", null, "=C19*1.24", "=C19*0.85", "=B19*1.06"],
          23: [null, null, "Runden"],
          24: [null, "Berechnung Erwärmung Kupferwindungen"],
          25: [null, "Temperatur Anfang", 23, null, 22.8],
          26: [null, "Temperatur Ende", 24.3, null, 23.4],
          27: [null, "Ohmwert Anfang", 179, null, 200.9],
          28: [null, "Ohmwert Ende", 331, null, 283.7],
          29: [null, "Konstante Kupfer", 234.5, null, 234.5],
          30: ["Ziff11.8", "Erwärmung in K", "=ROUND((C28-C27)/C27*(C29+C25)-(C26-C25),1)", null, "=MMULT((E28-E27)/E27*(E29+E25)-(E26-E25),1)"],
          31: ["Ziff19", "Absolut", "=C30+C26", null, "=E30+E26"],
          38: [null, "Berechnung Prüfspannung für CSA"],
          39: [null, "Wc = Wm x (125/Vm)2", null, "m=marked", "c=compensated"],
          40: [null, "Nennleistung", 1350],
          41: [null, "Nennspannung", 120],
          42: [null, "Gemessene Leistung", 1350],
          44: [null, "errechnet für CSA", "=C40*(125/C41)*(125/C41)"],
          45: [null, "Geprüft wird bei der Spannung, bei der die Leistung Wc erreicht wird"],
          48: [null, "Berechnung Prüfleistung AU wenn auf dem Typenschild die spannung als Bereich angegeben wird, jedoch die Leistung bezieht sich auf die 230V"],
          49: [null, "Nennleistung P", 1455, null, 220, 230, 240],
          50: [null, "unterer Bereichswert", 230],
          51: [null, "oberer Bereichswert", 240, null, "Ziff10 Watt (230)", "M.Bereich 220V", "(220/230)² *0.85*Pn="],
          52: [null, "Resultat (Prüfleistung)", "=C49*((C51/230)*(C51/230))*1.15", null, 1409, "=SUM(G52*H52*(E49/F49)*(E49/F49))", 0.85, 1455],
          53: [null, null, null, null, "Ziff10 Amp. (230)", "=SUM(H52*H53)", 1.24, 1.24],
          54: [null, "Resultat mal 0,85", "=C52*0.85", null, 6.12],
          55: [null, "Multipliziert mit 1,24", "=C56*1.24", null, "Result Mess.Span", "Result R Ohm"],
          56: [null, "Zurückgerechnet (geteilt durch 1,15)", "=C52/1.15", null, "=SUM(E52/E54)", "=SUM(E56/E54)", "=SQRT(F56*F52)"],
          59: ["Musster Formel", null, null, null, null, null, null, null, null, null, null, null, null, null, "Formel: P="],
          60: ["Sp.Bereich", "Ziff10 Pmax (W)", "Ziff10 Imax (A)", "Calc U=Ziff10", "R=Ziff10", "Pn(W)", "Bemessungsaufnahme", null, null, "V", "V", "V", "Prüfleistung", "Prüf=U"],
          61: ["220", 1409, 6.12, "=SUM(B61/C61)", "=SUM(D61/C61)", 1455, 0.85, 1.15, 1.24, 220, 230, 240, "=SUM(J61/K61*J61/K61*G61*F61)", "=SQRT(M61*E61)", "(220/230)² *0.85*Pn="],
          62: ["230", null, null, null, null, null, null, null, null, null, null, null, "=SUM(L61/K61*L61/K61*H61*F61)", null, "(240/230)² *1.15*Pn="],
          63: ["240", null, null, null, null, null, null, null, null, null, null, null, "=SUM(K61/J61*K61/J61*G61*F61)", "=SQRT(M63*E61)", "(230/220)² *0.85*Pn="],
          66: ["To Calc.", "bitte eingeben", "bitte eingeben", null, null, "bitte eingeben"],
          67: [null, 1296, 5.8902, "=SUM(B67/C67)", "=SUM(D67/C67)", 1300, 0.85, 1.15, 1.24, 220, 230, 240, "=SUM(J67/K67*J67/K67*G67*F67)", "=SQRT(M67*E67)", "(220/230)² *0.85*Pn="],
          68: [null, null, null, null, null, null, null, null, null, null, null, null, "=SUM(L67/K67*L67/K67*H67*F67)", null, "(240/230)² *1.15*Pn="],
          69: [null, null, null, null, null, null, null, null, null, null, null, null, "=SUM(K67/J67*K67/J67*G67*F67)", "=SQRT(M69*E67)", "(230/220)² *0.85*Pn="]
        }
      },
      { id: "tabelle2", title: "Tabelle2", cols: 6, rows: { 3: [null, "R2", "R1", "T1", "T2", "delta"], 4: [null, 200.2, 153, 23.9, 23.9, "=(((B4-C4)/C4)*(234.5+D4))-(E4-D4)"], 5: [null, 262.3, 153, 23.9, 23.9, "=(((B5-C5)/C5)*(234.5+D5))-(E5-D5)"], 8: [null, "=60*130/400"] } },
      { id: "tabelle3", title: "Tabelle3", cols: 1, rows: {} }
    ]
  }
];

function loadCalcOverrides() {
  try { return JSON.parse(localStorage.getItem(calcStorageKey) || "{}"); } catch { return {}; }
}
function saveCalcOverrides(overrides) { localStorage.setItem(calcStorageKey, JSON.stringify(overrides)); }
function loadWarmthOverrides() {
  try { return JSON.parse(localStorage.getItem(warmthStorageKey) || "{}"); } catch { return {}; }
}
function saveWarmthOverrides(overrides) {
  localStorage.setItem(warmthStorageKey, JSON.stringify(overrides));
}
function cellName(colIndex, rowNum) { return `${calcColumns[colIndex]}${rowNum}`; }
function isCalcFormula(value) { return typeof value === "string" && value.startsWith("="); }
function calcParseInput(value) {
  const trimmed = String(value ?? "").trim().replace(",", ".");
  if (trimmed === "") return "";
  const num = Number(trimmed);
  return Number.isFinite(num) ? num : value;
}
function calcFormat(value) {
  if (value === null || value === undefined || value === "") return "";
  if (typeof value !== "number") return value;
  if (!Number.isFinite(value)) return "";
  return Number.isInteger(value) ? String(value) : String(Math.round(value * 1000) / 1000);
}
function formatWarmth(value) {
  if (value === null || value === undefined || value === "") return "";
  if (typeof value !== "number" || !Number.isFinite(value)) return "";
  return String(Math.round(value * 1000) / 1000);
}
function parseWarmthValue(value) {
  const trimmed = String(value ?? "").trim().replace(",", ".");
  if (trimmed === "") return "";
  const num = Number(trimmed);
  return Number.isFinite(num) ? num : "";
}

function makeCalcEvaluator(sheet, overrides) {
  const memo = {};
  const rawValue = (ref) => {
    const key = `${sheet.id}:${ref}`;
    if (Object.prototype.hasOwnProperty.call(overrides, key)) return overrides[key];
    const col = calcColumns.indexOf(ref.match(/[A-Z]+/)[0]);
    const row = Number(ref.match(/\d+/)[0]);
    return sheet.rows[row]?.[col] ?? "";
  };
  const evaluateRef = (ref) => {
    if (memo[ref] !== undefined) return memo[ref];
    const value = rawValue(ref);
    if (isCalcFormula(value)) {
      memo[ref] = evaluateFormula(value);
      return memo[ref];
    }
    const parsed = calcParseInput(value);
    return typeof parsed === "number" ? parsed : 0;
  };
  const rangeValues = (range) => {
    const [start, end] = range.split(":");
    const startCol = calcColumns.indexOf(start.match(/[A-Z]+/)[0]);
    const endCol = calcColumns.indexOf(end.match(/[A-Z]+/)[0]);
    const startRow = Number(start.match(/\d+/)[0]);
    const endRow = Number(end.match(/\d+/)[0]);
    const values = [];
    for (let row = startRow; row <= endRow; row += 1) {
      for (let col = startCol; col <= endCol; col += 1) values.push(evaluateRef(cellName(col, row)));
    }
    return values;
  };
  const evaluateFormula = (formula) => {
    const ranges = [];
    let expr = formula.slice(1).replace(/\^/g, "**");
    expr = expr.replace(/\b([A-Z]{1,3}\d+:[A-Z]{1,3}\d+)\b/g, (_, range) => {
      const token = `__RANGE_${ranges.length}__`;
      ranges.push(range);
      return `"${token}"`;
    });
    expr = expr.replace(/\b([A-Z]{1,3}\d+)\b/g, (_, ref) => `CELL("${ref}")`);
    expr = expr.replace(/"__RANGE_(\d+)__"/g, (_, index) => `"${ranges[Number(index)]}"`);
    const CELL = evaluateRef;
    const RANGE = rangeValues;
    const SUM = (...args) => args.flatMap((arg) => typeof arg === "string" && arg.includes(":") ? RANGE(arg) : [arg]).reduce((sum, value) => sum + (Number(value) || 0), 0);
    const ROUND = (value, digits = 0) => {
      const factor = 10 ** Number(digits);
      return Math.round((Number(value) || 0) * factor) / factor;
    };
    const SQRT = (value) => Math.sqrt(Number(value) || 0);
    const MMULT = (a, b) => (Number(a) || 0) * (Number(b) || 0);
    try {
      return Function("CELL", "SUM", "ROUND", "SQRT", "MMULT", `"use strict"; return (${expr});`)(CELL, SUM, ROUND, SQRT, MMULT);
    } catch (err) {
      console.warn("Calculation formula failed", formula, err);
      return "";
    }
  };
  return { evaluateFormula };
}
function warmthKey(setIndex, field) {
  return `set${setIndex}:${field}`;
}
function warmthSet(overrides, setIndex, fallback) {
  return {
    tempStart: parseWarmthValue(overrides[warmthKey(setIndex, "tempStart")] ?? fallback.tempStart),
    tempEnd: parseWarmthValue(overrides[warmthKey(setIndex, "tempEnd")] ?? fallback.tempEnd),
    ohmStart: parseWarmthValue(overrides[warmthKey(setIndex, "ohmStart")] ?? fallback.ohmStart),
    ohmEnd: parseWarmthValue(overrides[warmthKey(setIndex, "ohmEnd")] ?? fallback.ohmEnd),
    kConst: parseWarmthValue(overrides[warmthKey(setIndex, "kConst")] ?? fallback.kConst),
  };
}
function warmthResult(set) {
  const values = [set.tempStart, set.tempEnd, set.ohmStart, set.ohmEnd, set.kConst];
  if (values.some((v) => v === "")) return { ziff118: "", ziff19: "" };
  const tempStart = Number(set.tempStart);
  const tempEnd = Number(set.tempEnd);
  const ohmStart = Number(set.ohmStart);
  const ohmEnd = Number(set.ohmEnd);
  const kConst = Number(set.kConst);
  const ziff118 = Math.round((((ohmEnd - ohmStart) / ohmStart) * (kConst + tempStart) - (tempEnd - tempStart)) * 10) / 10;
  return { ziff118, ziff19: ziff118 + tempEnd };
}
function renderCalculations() {
  const overrides = loadCalcOverrides();
  const warmthOverrides = loadWarmthOverrides();
  const value = (key, fallback) => Number(overrides[`simple:${key}`] ?? fallback);
  const minV = value("minV", 220);
  const maxV = value("maxV", 240);
  const ratedW = value("ratedW", 1380);
  const measuredW = value("measuredW", 1380);
  const setV = (minV + maxV) / 2;
  const powerAtMax = ratedW * ((maxV / setV) ** 2);
  const powerAtMin = ratedW * ((minV / setV) ** 2);
  const voltageForPower = (power) => Math.sqrt(((setV * setV) / measuredW) * power);
  const groups = [
    {
      ziffer: "Abs. 10",
      title: "Leistungsaufnahme",
      rows: [
        ["einzustellende Spannung", setV, "Volt"],
        ["Abweichung", (100 / ratedW * measuredW) - 100, "%"]
      ]
    },
    {
      ziffer: "Abs. 11.8",
      title: "Erwärmung",
      rows: [
        ["1,15-fache Bemessungsaufnahme", 1.15 * powerAtMax, "Watt"],
        ["Prüfspannung für 1,15-fach", voltageForPower(1.15 * powerAtMax), "Volt"]
      ]
    },
    {
      ziffer: "Abs. 11.Zxx",
      title: "Berührbare Oberflächen",
      rows: [
        ["1,0-fache Bemessungsaufnahme", powerAtMax, "Watt"],
        ["Prüfspannung für 1,0-fach", voltageForPower(powerAtMax), "Volt"]
      ]
    },
    {
      ziffer: "Abs. 16",
      title: "Ableitstrom",
      rows: [["1,06-fache Bemessungsspannung", maxV * 1.06, "Volt"]]
    },
    {
      ziffer: "Abs. 19",
      title: "Unsachgemäßer Gebrauch",
      rows: [
        ["0,85-fache Bemessungsaufnahme", 0.85 * powerAtMin, "Watt"],
        ["Prüfspannung für 0,85-fach", voltageForPower(0.85 * powerAtMin), "Volt"],
        ["1,24-fache Bemessungsaufnahme", 1.24 * powerAtMax, "Watt"],
        ["Prüfspannung für 1,24-fach", voltageForPower(1.24 * powerAtMax), "Volt"]
      ]
    },
    {
      ziffer: "Abs. 25",
      title: "Netzanschluß",
      rows: [["Berechneter Bemessungsstrom", ratedW / setV + 0.045, "Ampere"]]
    }
  ];
  const input = (key, label, unit, fallback) => `
    <label class="simple-calc-field">
      <span>${label}</span>
      <input data-simple-calc="${key}" type="number" step="0.01" value="${escapeHtml(value(key, fallback))}">
      <em>${unit}</em>
    </label>`;

  calculationsContainer.innerHTML = `
    <section class="simple-calc">
      <header class="simple-calc-header">
        <div>
          <h3>Leistungsaufnahme</h3>
          <p>Nur diese vier Werte eingeben. Alle Resultate darunter kommen aus den Excel-Formeln.</p>
        </div>
      </header>
      <div class="simple-calc-inputs">
        ${input("minV", "Min. Bemessungsspannung", "Volt", 220)}
        ${input("maxV", "Max. Bemessungsspannung", "Volt", 240)}
        ${input("ratedW", "Bemessungsaufnahme", "Watt", 1380)}
        ${input("measuredW", "Messwert Leistungsaufnahme", "Watt", 1380)}
      </div>
      <div class="simple-calc-results">
      ${groups.map((group) => `
          <section class="simple-calc-group">
            <header>
              <b>${group.ziffer}</b>
              <span>${group.title}</span>
            </header>
            ${group.rows.map(([label, result, unit]) => `
              <div class="simple-calc-result">
                <span>${label}</span>
                <strong>${escapeHtml(calcFormat(result))}</strong>
                <em>${unit}</em>
              </div>
            `).join("")}
          </section>
        `).join("")}
      </div>
      <div class="simple-calc-formula">
        <p><strong>Berechnung der 1,15-fachen Bemessungsaufnahme und die dafür benötigte Spannung:</strong></p>
        <p>Annahme: min. Bemessungsspannung = ${escapeHtml(calcFormat(minV))}V; max. Bemessungsspannung = ${escapeHtml(calcFormat(maxV))}V; einzustellende Spannung = ${escapeHtml(calcFormat(setV))}V</p>
        <p>Der Widerstandswert für die Berechnung der 1,15-fachen Bemessungsaufnahme ist ein theoretischer Wert, der über die Bemessungsaufnahme bzw. -spannung berechnet wird!</p>
        <p>Der Widerstandswert für die Berechnung der Spannung, die für die 1,15-fache Bemessungsaufnahme benötigt wird, bezieht sich auf die tatsächlich aufgenommene Leistung des Gerätes.</p>
        <img class="formula-image" src="assets/ls-formula.png" alt="Formelbereich aus dem Excel-Sheet">
      </div>
    </section>
    <section class="warmth-calc">
      <div class="warmth-title">Runden</div>
      <div class="warmth-sheet">
        <h3>Berechnung Erwärmung Kupferwindungen</h3>
        <table class="warmth-table">
          <thead>
            <tr>
              <th></th>
              <th>Set 1</th>
              <th>Set 2</th>
              <th>Set 3</th>
            </tr>
          </thead>
          <tbody>
            ${[
              ["Temperatur Anfang", "tempStart", 1, 23, "°C"],
              ["Temperatur Ende", "tempEnd", 1, 24.3, "°C"],
              ["Ohmwert Anfang", "ohmStart", 1, 179, "Ω"],
              ["Ohmwert Ende", "ohmEnd", 1, 331, "Ω"],
              ["Konstante Kupfer", "kConst", 1, 234.5, ""]
            ].map(([label, field, step, fallback, unit]) => `
              <tr>
                <th>${label}</th>
                ${[1,2,3].map((setIndex) => {
                  const fallbackValues = [
                    { tempStart: 23, tempEnd: 24.3, ohmStart: 179, ohmEnd: 331, kConst: 234.5 },
                    { tempStart: 22.8, tempEnd: 23.4, ohmStart: 200.9, ohmEnd: 283.7, kConst: 234.5 },
                    { tempStart: 22.8, tempEnd: 23.4, ohmStart: 200.9, ohmEnd: 283.7, kConst: 234.5 }
                  ][setIndex - 1];
                  const set = warmthSet(warmthOverrides, setIndex, fallbackValues);
                  const val = set[field];
                  const locked = field === "kConst" && setIndex !== 0;
                  return `
                    <td>
                      <input type="number" step="${step}" data-warmth-set="${setIndex}" data-warmth-field="${field}" value="${escapeHtml(val)}" ${field === "kConst" ? "readonly" : ""}>
                      ${unit ? `<span>${unit}</span>` : ""}
                    </td>
                  `;
                }).join("")}
              </tr>
            `).join("")}
            <tr class="warmth-result-row">
              <th>Erwärmung in Kelvin</th>
              ${[1,2,3].map((setIndex) => {
                const fallbackValues = [
                  { tempStart: 23, tempEnd: 24.3, ohmStart: 179, ohmEnd: 331, kConst: 234.5 },
                  { tempStart: 22.8, tempEnd: 23.4, ohmStart: 200.9, ohmEnd: 283.7, kConst: 234.5 },
                  { tempStart: 22.8, tempEnd: 23.4, ohmStart: 200.9, ohmEnd: 283.7, kConst: 234.5 }
                ][setIndex - 1];
                const set = warmthSet(warmthOverrides, setIndex, fallbackValues);
                const result = warmthResult(set);
                return `<td class="warmth-result">${escapeHtml(formatWarmth(result.ziff118))}</td>`;
              }).join("")}
            </tr>
            <tr class="warmth-result-row">
              <th>Absolut</th>
              ${[1,2,3].map((setIndex) => {
                const fallbackValues = [
                  { tempStart: 23, tempEnd: 24.3, ohmStart: 179, ohmEnd: 331, kConst: 234.5 },
                  { tempStart: 22.8, tempEnd: 23.4, ohmStart: 200.9, ohmEnd: 283.7, kConst: 234.5 },
                  { tempStart: 22.8, tempEnd: 23.4, ohmStart: 200.9, ohmEnd: 283.7, kConst: 234.5 }
                ][setIndex - 1];
                const set = warmthSet(warmthOverrides, setIndex, fallbackValues);
                const result = warmthResult(set);
                return `<td class="warmth-result">${escapeHtml(formatWarmth(result.ziff19))}</td>`;
              }).join("")}
            </tr>
          </tbody>
        </table>
        <div class="warmth-notes">
          <p><strong>Bemerkung:</strong></p>
          <p><strong>Annahme:</strong></p>
          <p>mit k = 225 für Aluminiumwicklungen und Kupfer/Aluminiumwicklungen mit einem Aluminiumanteil ≥ 85%</p>
          <p>mit k = 229,75 für Kupfer/Aluminiumwicklungen mit einem Kupferanteil &gt; 15% und &lt; 85%</p>
          <p>mit k = 234.5 für Kupferwicklungen</p>
        </div>
      </div>
    </section>
  `;
}

function renderTabelle24View() {
  if (!tabelle24Frame) return;
  if (!activeProject) {
    tabelle24Frame.removeAttribute("src");
    return;
  }
  const theme = document.body.dataset.theme || "light";
  const url = `/tabelle24.html?projectId=${encodeURIComponent(activeProject.id)}&theme=${theme}`;
  if (tabelle24Frame.getAttribute("src") !== url) {
    tabelle24Frame.setAttribute("src", url);
  }
}

function renderTabelle30View() {
  if (!tabelle30Frame) return;
  if (!activeProject) {
    tabelle30Frame.removeAttribute("src");
    return;
  }
  const theme = document.body.dataset.theme || "light";
  const url = `/tabelle30.html?projectId=${encodeURIComponent(activeProject.id)}&theme=${theme}`;
  if (tabelle30Frame.getAttribute("src") !== url) {
    tabelle30Frame.setAttribute("src", url);
  }
}

// ── Full render ───────────────────────────────────────────────
function renderAll() {
  if (!activeProject) return;
  renderMachines();
  loadProductImage(activeProject.id);
  renderSummary(activeProject);
  renderOverview(activeProject);
  renderSubtopic(activeProject);
  renderTasks(activeProject);
  renderDocs(activeProject);
  renderFachfreigabe(activeProject);
  renderCalculations();
  renderTabelle24View();
  renderTabelle30View();
}

async function openProject(projectId, view = "overview", pushHistory = true) {
  recordRecent(projectId);
  activeProject = await apiFetch(`/api/projects/${projectId}`);
  activeSubtopic = "Approbation";
  activeBuild = "Alle";
  activeEvidenceGroup = null;
  activeMarketFilter = null;
  excelFilesCache = null;
  subtopicFilter.value = "all";
  setView(view);
  renderAll();
  if (pushHistory) history.pushState({ projectId, view }, "", `#${projectId}`);
}

function goToDashboard(pushHistory = true) {
  activeProject = null;
  setView("dashboard");
  renderDashboard();
  if (pushHistory) history.pushState({ dashboard: true }, "", location.pathname);
}

async function ensureProjectData(projectId) {
  if (activeProject?.id !== projectId) {
    activeProject = await apiFetch(`/api/projects/${projectId}`);
  }
}

window.addEventListener("popstate", async (e) => {
  if (e.state?.subfolderGroup && e.state?.projectId) {
    // Level 3: inside subfolder
    await ensureProjectData(e.state.projectId);
    activeEvidenceGroup = e.state.openGroup;
    setView("docs");
    const group = activeProject.documentGroups?.find((g) => g.primary === e.state.subfolderGroup);
    if (group) loadEvidenceEntries(group, e.state.subfolderHref);
  } else if (e.state?.projectId && e.state?.view === "docs" && e.state?.openGroup) {
    // Level 2: panel open at root
    await ensureProjectData(e.state.projectId);
    activeEvidenceGroup = e.state.openGroup;
    setView("docs");
    const group = activeProject.documentGroups?.find((g) => g.primary === e.state.openGroup);
    if (group) loadEvidenceEntries(group, evidenceHref(group));
    else renderDocs(activeProject);
  } else if (e.state?.projectId && e.state?.view === "docs") {
    // Level 1: docs tab, panel closed
    await ensureProjectData(e.state.projectId);
    activeEvidenceGroup = null;
    setView("docs");
    renderDocs(activeProject);
  } else if (e.state?.projectId && e.state?.view === "tabelle24") {
    await ensureProjectData(e.state.projectId);
    setView("tabelle24");
    renderTabelle24View();
  } else if (e.state?.projectId && e.state?.view === "tabelle30") {
    await ensureProjectData(e.state.projectId);
    setView("tabelle30");
    renderTabelle30View();
  } else if (e.state?.projectId) {
    openProject(e.state.projectId, e.state.view || "overview", false);
  } else {
    activeProject = null;
    setView("dashboard");
    renderDashboard();
  }
});

function renderBuildChange() {
  updateSidebarBuildSelection();
  activeEvidenceGroup = null;
  renderSubtopic(activeProject);
  renderTasks(activeProject);
  renderDocs(activeProject);
}

// ── Events ────────────────────────────────────────────────────
machineList.addEventListener("click", async (event) => {
  const pill = event.target.closest("[data-build-select]");
  const btn = event.target.closest("button[data-id]");
  if (!btn) return;

  if (pill) {
    event.stopPropagation();
    if (btn.dataset.id !== activeProject?.id) {
      recordRecent(btn.dataset.id);
      activeProject = await apiFetch(`/api/projects/${btn.dataset.id}`);
      activeBuild = pill.dataset.buildSelect;
      activeSubtopic = "Approbation";
      activeEvidenceGroup = null;
      subtopicFilter.value = "all";
      renderAll();
      setView(activeView === "dashboard" ? "subtopic" : activeView);
      return;
    }
    activeBuild = pill.dataset.buildSelect;
    activeSubtopic = "Approbation";
    subtopicFilter.value = "all";
    renderBuildChange();
    setView(activeView === "dashboard" ? "subtopic" : activeView);
    return;
  }

  openProject(btn.dataset.id, activeView === "dashboard" ? "overview" : activeView);
});

dashboardLink.addEventListener("click", () => {
  dashboardSearch.value = "";
  goToDashboard();
  renderMachines();
});

tabelle24Link.addEventListener("click", () => {
  if (!activeProject) return;
  setView("tabelle24");
  renderTabelle24View();
  history.pushState({ projectId: activeProject.id, view: "tabelle24" }, "", `#${activeProject.id}/tabelle24`);
});

tabelle30Link?.addEventListener("click", () => {
  if (!activeProject) return;
  setView("tabelle30");
  renderTabelle30View();
  history.pushState({ projectId: activeProject.id, view: "tabelle30" }, "", `#${activeProject.id}/tabelle30`);
});

document.querySelector("#dashboard-grid").addEventListener("click", (event) => {
  // Click on project-no badge → inline edit, don't navigate
  const badge = event.target.closest(".project-no-badge");
  if (badge) {
    event.stopPropagation();
    const card = badge.closest("[data-dashboard-project]");
    const projectId = card.dataset.dashboardProject;
    const project = projectList.find((p) => p.id === projectId);
    if (!project) return;
    const current = project.project_no || "";
    const input = document.createElement("input");
    input.className = "project-no-input";
    input.value = current;
    input.placeholder = "z.B. 1234567";
    badge.replaceWith(input);
    input.focus();
    input.select();
    const save = async () => {
      const val = input.value.trim();
      project.project_no = val;
      await apiPut(`/api/projects/${projectId}/project-no`, { project_no: val }).catch(console.error);
      renderDashboard();
    };
    input.addEventListener("blur", save);
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); input.blur(); }
      if (e.key === "Escape") { input.removeEventListener("blur", save); renderDashboard(); }
    });
    return;
  }

  // Click on machine type/use badge → dropdown select
  const metaBadge = event.target.closest("[data-meta-field]");
  if (metaBadge) {
    event.stopPropagation();
    const field = metaBadge.dataset.metaField;       // "machine_type" or "machine_use"
    const projectId = metaBadge.dataset.metaProject;
    const project = projectList.find((p) => p.id === projectId);
    if (!project) return;
    const options = field === "machine_type"
      ? ["", "Nespresso", "Kapselmaschine / Nespresso Vertuo", "Kapselmaschine / CoffeeB", "Cafissimo", "Carogusto", "Kapselmaschine", "Vollautomatisch"]
      : ["", "Commercial Use", "Private Use"];
    const sel = document.createElement("select");
    sel.className = "meta-select";
    options.forEach((opt) => {
      const o = document.createElement("option");
      o.value = opt;
      o.textContent = opt || "— wählen —";
      if ((project[field] || "") === opt) o.selected = true;
      sel.appendChild(o);
    });
    metaBadge.replaceWith(sel);
    sel.focus();
    const save = async () => {
      const val = sel.value;
      project[field] = val;
      const endpoint = field === "machine_type" ? "machine-type" : "machine-use";
      const body = field === "machine_type" ? { machine_type: val } : { machine_use: val };
      await apiPut(`/api/projects/${projectId}/${endpoint}`, body).catch(console.error);
      renderDashboard();
    };
    sel.addEventListener("change", save);
    sel.addEventListener("blur", () => { if (document.activeElement !== sel) save(); });
    sel.addEventListener("keydown", (e) => {
      if (e.key === "Escape") { sel.removeEventListener("change", save); renderDashboard(); }
    });
    return;
  }

  // Click on SW/HW version badge → inline edit
  const versionBadge = event.target.closest("[data-version-field]");
  if (versionBadge) {
    event.stopPropagation();
    const field = versionBadge.dataset.versionField;   // "sw" or "hw"
    const projectId = versionBadge.dataset.versionProject;
    const project = projectList.find((p) => p.id === projectId);
    if (!project) return;
    const current = (field === "sw" ? project.sw_version : project.hw_version) || "";
    const input = document.createElement("input");
    input.className = "version-input";
    input.value = current;
    input.placeholder = field === "sw" ? "z.B. V1.2.3" : "z.B. Rev B";
    versionBadge.replaceWith(input);
    input.focus();
    input.select();
    const save = async () => {
      const val = input.value.trim();
      if (field === "sw") { project.sw_version = val; await apiPut(`/api/projects/${projectId}/sw-version`, { sw_version: val }).catch(console.error); }
      else                { project.hw_version = val; await apiPut(`/api/projects/${projectId}/hw-version`, { hw_version: val }).catch(console.error); }
      renderDashboard();
    };
    input.addEventListener("blur", save);
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); input.blur(); }
      if (e.key === "Escape") { input.removeEventListener("blur", save); renderDashboard(); }
    });
    return;
  }
  const card = event.target.closest("[data-dashboard-project]");
  if (!card) return;
  openProject(card.dataset.dashboardProject, "overview");
});

recentlyOpenedEl.addEventListener("click", (event) => {
  const card = event.target.closest("[data-dashboard-project]");
  if (!card) return;
  openProject(card.dataset.dashboardProject, "overview");
});

search.addEventListener("input", () => {
  renderMachines();
  if (activeView === "dashboard") renderDashboard();
});
dashboardSearch.addEventListener("input", () => renderDashboard());
taskFilter.addEventListener("change", () => renderTasks(activeProject));
subtopicFilter.addEventListener("change", () => renderSubtopic(activeProject));

zifferTable.addEventListener("click", async (event) => {
  const btn = event.target.closest("[data-nr]");
  if (!btn || btn.dataset.reasonNr) return;
  const ziffer = activeProject.subtopics?.[activeSubtopic]?.ziffern.find((z) => z.nr === btn.dataset.nr);
  if (!ziffer) return;
  const next = statusFlow[(statusFlow.indexOf(ziffer.status) + 1) % statusFlow.length];
  ziffer.status = next; // optimistic update
  renderSubtopic(activeProject);
  apiPut(`/api/projects/${activeProject.id}/ziffern/${activeSubtopic}/${btn.dataset.nr}`, { status: next })
    .catch(console.error);
});

zifferTable.addEventListener("input", (event) => {
  const ta = event.target.closest("[data-reason-nr]");
  if (!ta) return;
  ta.style.height = "auto";
  ta.style.height = ta.scrollHeight + "px";
});

zifferTable.addEventListener("focusout", (event) => {
  const ta = event.target.closest("[data-reason-nr]");
  if (!ta) return;
  const nr = ta.dataset.reasonNr;
  const ziffer = activeProject.subtopics?.[activeSubtopic]?.ziffern.find((z) => z.nr === nr);
  if (ziffer) ziffer.not_needed_reason = ta.value;
  apiPut(`/api/projects/${activeProject.id}/ziffern/${activeSubtopic}/${nr}/reason`, { reason: ta.value })
    .catch(console.error);
});

document.querySelector("#summary-grid").addEventListener("click", (event) => {
  const archiveBtn = event.target.closest("[data-archive-open]");
  if (archiveBtn) {
    apiFetch(`/api/open-archive?project_id=${encodeURIComponent(archiveBtn.dataset.archiveOpen)}`).catch(console.error);
    return;
  }
  const btn = event.target.closest("[data-subtopic]");
  if (!btn) return;
  activeSubtopic = btn.dataset.subtopic;
  subtopicFilter.value = "all";
  renderSubtopic(activeProject);
  setView("subtopic");
});

// Fachfreigabe gate toggle
document.querySelector("#freigabe-view").addEventListener("click", (event) => {
  const btn = event.target.closest(".freigabe-btn[data-label]");
  if (!btn) return;
  const gate = activeProject.fachfreigabe?.gates?.find((g) => g.label === btn.dataset.label);
  if (!gate) return;
  const next = freigabeFlow[(freigabeFlow.indexOf(freigabeStatus(gate.status)) + 1) % freigabeFlow.length];
  gate.status = next; // optimistic update
  renderFachfreigabe(activeProject);
  apiPut(`/api/projects/${activeProject.id}/fachfreigabe/gates`, { label: btn.dataset.label, status: next })
    .catch(console.error);
});

// Fachfreigabe meta fields — debounced save
let ffMetaTimer = null;
document.querySelector("#freigabe-view").addEventListener("input", (event) => {
  if (!["ff-bestaetigt", "ff-datum", "ff-notiz"].includes(event.target.id)) return;
  clearTimeout(ffMetaTimer);
  ffMetaTimer = setTimeout(() => {
    const meta = {
      confirmed_by: document.querySelector("#ff-bestaetigt")?.value || "",
      datum: document.querySelector("#ff-datum")?.value || "",
      notiz: document.querySelector("#ff-notiz")?.value || ""
    };
    apiPut(`/api/projects/${activeProject.id}/fachfreigabe/meta`, meta).catch(console.error);
  }, 600);
});

// Task status toggle
document.querySelector("#task-table").addEventListener("click", async (event) => {
  const btn = event.target.closest("[data-task-id]");
  if (!btn) return;
  const taskId = btn.dataset.taskId;
  const task = activeProject.tasks?.find((t) => String(t.id) === taskId);
  if (!task) return;
  const next = taskFlow[(taskFlow.indexOf(task.status) + 1) % taskFlow.length];
  task.status = next;
  renderTasks(activeProject);
  renderOverview(activeProject);
  apiPut(`/api/projects/${activeProject.id}/tasks/${taskId}`, {
    status: next,
    block_reason: task.block_reason || ""
  }).catch(console.error);
});

const taskBlockReasonTimers = new Map();
function saveTaskBlockReason(projectId, task) {
  if (!projectId || !task) return;
  apiPut(`/api/projects/${projectId}/tasks/${task.id}`, {
    status: task.status,
    block_reason: task.block_reason || ""
  }).catch(console.error);
}

function handleTaskBlockReasonInput(field, immediate = false) {
  const taskId = field.dataset.taskBlockReason;
  const task = activeProject.tasks?.find((t) => String(t.id) === taskId);
  if (!task) return;
  const projectId = activeProject.id;
  task.block_reason = field.value;
  clearTimeout(taskBlockReasonTimers.get(taskId));
  if (immediate) {
    saveTaskBlockReason(projectId, task);
    taskBlockReasonTimers.delete(taskId);
    return;
  }
  taskBlockReasonTimers.set(taskId, setTimeout(() => {
    saveTaskBlockReason(projectId, task);
    taskBlockReasonTimers.delete(taskId);
  }, 300));
}

document.querySelector("#task-table").addEventListener("input", (event) => {
  const field = event.target.closest("[data-task-block-reason]");
  if (!field) return;
  handleTaskBlockReasonInput(field);
});

document.querySelector("#task-table").addEventListener("focusout", (event) => {
  const field = event.target.closest("[data-task-block-reason]");
  if (!field) return;
  handleTaskBlockReasonInput(field, true);
});

document.querySelector("#docs-view").addEventListener("dragstart", (event) => {
  const row = event.target.closest("[data-evidence-entry-href]");
  if (!row) return;

  const href = row.dataset.evidenceEntryHref;
  const sources = selectedEvidenceHrefs.has(href) ? [...selectedEvidenceHrefs] : [href];
  evidenceDrag = { sources, group: row.dataset.evidenceEntryGroup };
  row.classList.add("dragging");
  document.querySelector("#docs-view").classList.add("evidence-drag-active");
  if (event.dataTransfer) {
    event.dataTransfer.effectAllowed = "move";
    event.dataTransfer.setData("text/plain", sources.join("\n"));
  }
});

document.querySelector("#docs-view").addEventListener("dragover", (event) => {
  if (!evidenceDrag) return;

  const parentZone = event.target.closest("[data-parent-drop-href]");
  if (parentZone) {
    event.preventDefault();
    if (event.dataTransfer) event.dataTransfer.dropEffect = "move";
    parentZone.classList.add("drop-target");
    clearSpring();
    return;
  }

  // Rechte Top-Ordner-Kachel als Drop-Ziel — Verweilen öffnet den Ordner links
  const card = event.target.closest(".document-group-card[data-group-drop-href]");
  if (card) {
    event.preventDefault();
    if (event.dataTransfer) event.dataTransfer.dropEffect = "move";
    card.classList.add("drop-target");
    const cardPrimary = card.dataset.evidenceGroupCard;
    if (cardPrimary) {
      scheduleSpring("card:" + cardPrimary, () => {
        const g = activeProject.documentGroups?.find((x) => x.primary === cardPrimary);
        if (!g || activeEvidenceGroup === cardPrimary) return;
        activeEvidenceGroup = cardPrimary;
        renderDocs(activeProject);
        loadEvidenceEntries(g);   // gemerkten Unterordner wieder öffnen
      });
    } else clearSpring();
    return;
  }

  // Unterordner-Zeile links — Verweilen navigiert hinein
  const row = event.target.closest("[data-evidence-entry-href]");
  if (row && row.dataset.evidenceEntryType === "Ordner" && !evidenceDrag.sources.includes(row.dataset.evidenceEntryHref)) {
    event.preventDefault();
    if (event.dataTransfer) event.dataTransfer.dropEffect = "move";
    row.classList.add("drop-target");
    const rowHref = row.dataset.evidenceEntryHref;
    scheduleSpring("row:" + rowHref, () => {
      const ctx = currentEvidenceContext();
      if (!ctx) return;
      const cached = evidenceEntries.get(ctx.key);
      if (cached?.browseHref === rowHref) return;
      loadEvidenceEntries(ctx.group, rowHref);
    });
    return;
  }

  // über kein Ordner-Ziel → geplantes Öffnen abbrechen
  clearSpring();
});

document.querySelector("#docs-view").addEventListener("dragleave", (event) => {
  const parentZone = event.target.closest("[data-parent-drop-href]");
  if (parentZone && !parentZone.contains(event.relatedTarget)) {
    parentZone.classList.remove("drop-target");
    return;
  }

  const card = event.target.closest(".document-group-card[data-group-drop-href]");
  if (card && !card.contains(event.relatedTarget)) {
    card.classList.remove("drop-target");
    return;
  }

  const row = event.target.closest("[data-evidence-entry-href]");
  if (!row || row.contains(event.relatedTarget)) return;
  row.classList.remove("drop-target");
});

document.querySelector("#docs-view").addEventListener("drop", async (event) => {
  clearSpring();
  const parentZone = event.target.closest("[data-parent-drop-href]");
  const card = event.target.closest(".document-group-card[data-group-drop-href]");
  const row = event.target.closest("[data-evidence-entry-href]");
  document.querySelector("#docs-view").classList.remove("evidence-drag-active");
  document.querySelectorAll(".evidence-parent-drop.drop-target, .document-group-card.drop-target").forEach((el) => el.classList.remove("drop-target"));
  document.querySelectorAll(".evidence-file-row.drop-target").forEach((el) => el.classList.remove("drop-target"));
  document.querySelectorAll(".evidence-file-row.dragging").forEach((el) => el.classList.remove("dragging"));
  const isFolderRow = row && row.dataset.evidenceEntryType === "Ordner";
  if (!evidenceDrag || (!parentZone && !card && !isFolderRow)) {
    evidenceDrag = null;
    return;
  }

  event.preventDefault();
  const targetHref = parentZone ? parentZone.dataset.parentDropHref
    : card ? card.dataset.groupDropHref
    : row.dataset.evidenceEntryHref;
  const sources = evidenceDrag.sources;
  evidenceDrag = null;
  await moveEvidenceEntriesToFolder(sources, targetHref);
});

document.querySelector("#docs-view").addEventListener("dragend", () => {
  evidenceDrag = null;
  clearSpring();
  document.querySelector("#docs-view").classList.remove("evidence-drag-active");
  document.querySelectorAll(".evidence-parent-drop.drop-target, .document-group-card.drop-target").forEach((el) => el.classList.remove("drop-target"));
  document.querySelectorAll(".evidence-file-row.drop-target, .evidence-file-row.dragging")
    .forEach((el) => el.classList.remove("drop-target", "dragging"));
});

// Live-Suche im aktuellen Ordner (filtert die Zeilen ohne Re-Render → Fokus bleibt)
document.querySelector("#docs-view").addEventListener("input", (event) => {
  const box = event.target.closest("[data-evidence-search]");
  if (!box) return;
  evidenceSearch = box.value;
  filterEvidenceRows();
});

// Open local folder in Finder/Explorer via local backend
document.querySelector("#docs-view").addEventListener("click", async (event) => {
  const fileActionBtn = event.target.closest("[data-file-action]");
  if (fileActionBtn) {
    await runFileAction(fileActionBtn.dataset.fileAction, fileActionBtn.dataset.currentHref);
    return;
  }

  const entryRow = event.target.closest("[data-evidence-entry-href]");
  if (entryRow && !event.target.closest("[data-open-href]")) {
    const href = entryRow.dataset.evidenceEntryHref;
    const type = entryRow.dataset.evidenceEntryType;
    const group = activeProject.documentGroups?.find((g) => g.primary === entryRow.dataset.evidenceEntryGroup);
    const rows = [...document.querySelectorAll(`.evidence-file-row[data-evidence-entry-group="${CSS.escape(entryRow.dataset.evidenceEntryGroup)}"]`)];
    const alreadySelected = selectedEvidenceHrefs.has(href) && selectedEvidenceHrefs.size === 1;

    if (event.shiftKey && evidenceSelectionAnchor) {
      const start = rows.findIndex((row) => row.dataset.evidenceEntryHref === evidenceSelectionAnchor);
      const end = rows.indexOf(entryRow);
      selectedEvidenceHrefs = new Set();
      if (start >= 0 && end >= 0) {
        const [from, to] = start < end ? [start, end] : [end, start];
        rows.slice(from, to + 1).forEach((row) => selectedEvidenceHrefs.add(row.dataset.evidenceEntryHref));
      } else {
        selectedEvidenceHrefs.add(href);
        evidenceSelectionAnchor = href;
      }
    } else {
      selectedEvidenceHrefs = new Set([href]);
      evidenceSelectionAnchor = href;
    }
    selectedEvidenceHref = href;
    // Beim Auswählen einer Datei die Scroll-Position der Liste erhalten, damit
    // sie nicht nach oben springt. Der Scroll-Container ist .evidence-file-table,
    // den renderDocs neu erzeugt → Wert sichern und auf dem neuen Element setzen.
    const savedScroll = document.querySelector("#docs-detail-pane .evidence-file-table")?.scrollTop || 0;
    renderDocs(activeProject);
    const newScroller = document.querySelector("#docs-detail-pane .evidence-file-table");
    if (newScroller) newScroller.scrollTop = savedScroll;

    if (type === "Ordner") {
      if (!group) return;
      if (!alreadySelected || event.shiftKey) return;
      history.pushState({ projectId: activeProject.id, view: "docs", openGroup: group.primary, subfolderGroup: group.primary, subfolderHref: href }, "", `#${activeProject.id}`);
      clearEvidenceSelection();
      loadEvidenceEntries(group, href);
      return;
    }

    // If the preview panel is already open, a single click on a different file should swap
    // the preview to that file (no need to click twice). Without preview, keep "click-twice"
    // behaviour so the first click only selects.
    const previewActive = document.querySelector("#file-preview-panel")?.classList.contains("active");
    if (event.shiftKey || !group) return;
    if (!alreadySelected && !previewActive) return;

    if (isOfficeFile(entryRow.dataset.evidenceEntryName)) {
      try {
        await apiFetch("/api/open-path", {
          method: "POST",
          body: JSON.stringify({ href: new URL(href, window.location.href).pathname, app: "word" })
        });
      } catch (err) {
        console.error("Open office file error:", err);
      }
      return;
    }

    if (/\.(jpe?g|png|gif|webp|svg|bmp|pdf)$/i.test(href)) {
      const previewRow = document.querySelector(`[data-evidence-entry-href="${CSS.escape(href)}"]`);
      if (previewRow) previewRow.dispatchEvent(new CustomEvent("evidence-preview-open", { bubbles: true, detail: { href, name: entryRow.dataset.evidenceEntryName } }));
      return;
    }

    window.open(href, "_blank", "noreferrer");
    return;
  }

  // Breadcrumb-Navigation: zu einer Ebene im Pfad springen
  const crumb = event.target.closest("[data-evidence-crumb]");
  if (crumb) {
    const group = activeProject.documentGroups?.find((g) => g.primary === crumb.dataset.evidenceCrumbGroup);
    if (group) {
      const href = crumb.dataset.evidenceCrumb;
      history.pushState({ projectId: activeProject.id, view: "docs", openGroup: group.primary, subfolderGroup: group.primary, subfolderHref: href }, "", `#${activeProject.id}`);
      loadEvidenceEntries(group, href);
    }
    return;
  }

  // Sortierung umschalten (Spaltenkopf)
  const sortBtn = event.target.closest("[data-sort]");
  if (sortBtn) {
    const sortKey = sortBtn.dataset.sort;
    if (evidenceSort.key === sortKey) evidenceSort.dir = evidenceSort.dir === "asc" ? "desc" : "asc";
    else { evidenceSort.key = sortKey; evidenceSort.dir = "asc"; }
    const saved = document.querySelector("#docs-detail-pane .evidence-file-table")?.scrollTop || 0;
    renderDocs(activeProject);
    const sc = document.querySelector("#docs-detail-pane .evidence-file-table");
    if (sc) sc.scrollTop = saved;
    return;
  }

  // Subfolder navigation
  const subfolderBtn = event.target.closest("[data-browse-subfolder]");
  if (subfolderBtn) {
    const group = activeProject.documentGroups?.find((g) => g.primary === subfolderBtn.dataset.browseGroup);
    if (group) {
      const href = subfolderBtn.dataset.browseSubfolder;
      history.pushState({ projectId: activeProject.id, view: "docs", openGroup: group.primary, subfolderGroup: group.primary, subfolderHref: href }, "", `#${activeProject.id}`);
      loadEvidenceEntries(group, href);
    }
    return;
  }
  // Back button
  const backBtn = event.target.closest("[data-evidence-back]");
  if (backBtn) {
    history.back();
    return;
  }

  const groupBtn = event.target.closest("[data-evidence-group]");
  if (groupBtn) {
    const wasOpen = activeEvidenceGroup === groupBtn.dataset.evidenceGroup;
    toggleEvidenceGroup(groupBtn.dataset.evidenceGroup);
    if (!wasOpen && activeEvidenceGroup) {
      history.replaceState({ projectId: activeProject.id, view: "docs" }, "", `#${activeProject.id}`);
      history.pushState({ projectId: activeProject.id, view: "docs", openGroup: activeEvidenceGroup }, "", `#${activeProject.id}`);
    }
    const group = activeProject.documentGroups?.find((item) => item.primary === groupBtn.dataset.evidenceGroup);
    renderDocs(activeProject);
    if (activeEvidenceGroup && group) loadEvidenceEntries(group);
    return;
  }

  const retryBtn = event.target.closest("[data-evidence-retry]");
  if (retryBtn) {
    const group = activeProject.documentGroups?.find((item) => item.primary === retryBtn.dataset.evidenceRetry);
    if (!group) return;
    const key = evidenceCacheKey(activeProject.id, group.primary);
    const cached = evidenceEntries.get(key);
    const href = cached?.browseHref || evidenceHref(group);
    evidencePathEntries.delete(`${key}:${href}`);
    loadEvidenceEntries(group, href);
    return;
  }

  const groupCard = event.target.closest("[data-evidence-group-card]");
  if (groupCard && !event.target.closest("a, button")) {
    const cardGroup = groupCard.dataset.evidenceGroupCard;
    // Klick auf eine Kachel wählt sie immer aus (kein Umschalten/Abwählen),
    // damit ein Doppelklick auf denselben Ordner nicht zum ersten zurückspringt.
    if (activeEvidenceGroup === cardGroup) return;
    const hadSelection = !!activeEvidenceGroup;
    activeEvidenceGroup = cardGroup;
    if (!hadSelection) {
      history.replaceState({ projectId: activeProject.id, view: "docs" }, "", `#${activeProject.id}`);
    }
    history.pushState({ projectId: activeProject.id, view: "docs", openGroup: activeEvidenceGroup }, "", `#${activeProject.id}`);
    const group = activeProject.documentGroups?.find((item) => item.primary === cardGroup);
    renderDocs(activeProject);
    if (group) loadEvidenceEntries(group);
    return;
  }

  const clearSpace = event.target.closest("[data-clear-evidence-selection]");
  if (clearSpace) {
    if (clearEvidenceSelection()) renderDocs(activeProject);
    return;
  }

  const btn = event.target.closest("[data-open-href]");
  if (!btn) return;
  btn.disabled = true;
  try {
    await apiFetch("/api/open-path", {
      method: "POST",
      body: JSON.stringify({ href: new URL(btn.dataset.openHref, window.location.href).pathname })
    });
  } catch (err) {
    console.error("Open path error:", err);
  } finally {
    btn.disabled = false;
  }
});

document.addEventListener("keydown", async (event) => {
  if (activeView !== "docs") return;
  if (event.target.closest("input, textarea, select, [contenteditable='true']")) return;
  const isMod = event.metaKey || event.ctrlKey;
  const key = event.key.toLowerCase();

  if (isMod && key === "c") {
    event.preventDefault();
    await runFileAction("copy");
  } else if (isMod && key === "x") {
    event.preventDefault();
    await runFileAction("cut");
  } else if (isMod && key === "v") {
    event.preventDefault();
    await runFileAction("paste");
  } else if (isMod && event.shiftKey && key === "n") {
    event.preventDefault();
    await runFileAction("mkdir");
  } else if (event.key === "Delete" || event.key === "Backspace") {
    event.preventDefault();
    await runFileAction("delete");
  }
});

// ── File preview panel ────────────────────────────────────────
(function () {
  const panel = document.querySelector("#file-preview-panel");
  const nameEl = document.querySelector("#file-preview-name");
  const body = document.querySelector("#file-preview-body");
  const zoomBar = document.querySelector("#file-preview-zoom");
  const closeBtn = document.querySelector("#file-preview-close");
  const IMAGE_EXT = /\.(jpe?g|png|gif|webp|svg|bmp)$/i;
  const PDF_EXT = /\.pdf$/i;
  let zoomLevel = 1;
  let selectedRow = null;
  let hoveredRow = null;

  function canPreview(href) { return IMAGE_EXT.test(href) || PDF_EXT.test(href); }

  function setZoom(z) {
    zoomLevel = Math.min(5, Math.max(0.2, z));
    const img = body.querySelector("img");
    if (img) img.style.transform = `scale(${zoomLevel})`;
  }

  function openPreview(row, href, name) {
    if (selectedRow) selectedRow.classList.remove("preview-selected");
    tx = 0; ty = 0;
    panel.style.transform = "translate(0,0)";
    selectedRow = row;
    row.classList.add("preview-selected");
    nameEl.textContent = name;
    body.innerHTML = "";
    zoomLevel = 1;
    if (IMAGE_EXT.test(href)) {
      zoomBar.style.display = "flex";
      const img = document.createElement("img");
      img.src = href; img.alt = name;
      body.appendChild(img);
    } else {
      zoomBar.style.display = "none";
      const iframe = document.createElement("iframe");
      iframe.src = href;
      body.appendChild(iframe);
    }
    panel.classList.add("active");
  }

  function closePreview() {
    panel.classList.remove("active");
    body.innerHTML = "";
    if (selectedRow) selectedRow.classList.remove("preview-selected");
    selectedRow = null;
  }

  closeBtn.addEventListener("click", closePreview);

  const fullscreenBtn = document.querySelector("#file-preview-fullscreen");
  fullscreenBtn.addEventListener("click", () => {
    const isFs = panel.classList.toggle("fullscreen");
    fullscreenBtn.textContent = isFs ? "⤡" : "⤢";
    fullscreenBtn.title = isFs ? "Vollbild beenden" : "Vollbild";
  });

  // Resize from left edge
  const resizeHandle = document.querySelector("#file-preview-resize");
  resizeHandle.addEventListener("mousedown", (e) => {
    e.preventDefault();
    const startX = e.clientX;
    const startWidth = panel.offsetWidth;
    let raf = null;
    panel.classList.add("dragging");
    const onMove = (ev) => {
      if (raf) return;
      raf = requestAnimationFrame(() => {
        raf = null;
        panel.style.width = Math.max(320, Math.min(window.innerWidth - 40, startWidth + (startX - ev.clientX))) + "px";
      });
    };
    const onUp = () => {
      panel.classList.remove("dragging");
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
    };
    document.addEventListener("mousemove", onMove, { passive: true });
    document.addEventListener("mouseup", onUp);
  });

  // Drag to move via header — GPU-accelerated via transform
  const header = document.querySelector(".file-preview-header");
  let tx = 0, ty = 0;
  header.addEventListener("mousedown", (e) => {
    if (e.target.closest("button")) return;
    e.preventDefault();
    const startX = e.clientX - tx;
    const startY = e.clientY - ty;
    panel.classList.add("dragging");
    let raf = null;
    const onMove = (ev) => {
      if (raf) return;
      raf = requestAnimationFrame(() => {
        raf = null;
        const rect = panel.getBoundingClientRect();
        tx = Math.max(-rect.left, Math.min(window.innerWidth - rect.right, ev.clientX - startX));
        ty = Math.max(-rect.top, Math.min(window.innerHeight - rect.bottom + rect.height - 60, ev.clientY - startY));
        panel.style.transform = `translate(${tx}px, ${ty}px)`;
      });
    };
    const onUp = () => {
      panel.classList.remove("dragging");
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
    };
    document.addEventListener("mousemove", onMove, { passive: true });
    document.addEventListener("mouseup", onUp);
  });
  document.querySelector("#preview-zoom-in").addEventListener("click", () => setZoom(zoomLevel * 1.25));
  document.querySelector("#preview-zoom-out").addEventListener("click", () => setZoom(zoomLevel * 0.8));
  document.querySelector("#preview-zoom-reset").addEventListener("click", () => setZoom(1));

  document.querySelector("#docs-view").addEventListener("mouseover", (e) => {
    const row = e.target.closest(".evidence-file-row:not(.evidence-file-row--folder):not(.head)");
    hoveredRow = row || null;
  });

  document.querySelector("#docs-view").addEventListener("evidence-preview-open", (e) => {
    const row = e.target.closest(".evidence-file-row:not(.head)");
    const { href, name } = e.detail || {};
    if (row && href && canPreview(href)) openPreview(row, href, name || row.textContent.trim());
  });

  window.addEventListener("keydown", (e) => {
    if (e.key === " ") {
      if (panel.classList.contains("active")) { e.preventDefault(); closePreview(); return; }
      const selectedRow = document.querySelector(".evidence-file-row.explorer-selected:not(.head)") || hoveredRow;
      if (selectedRow) {
        e.preventDefault();
        const href = selectedRow.dataset.evidenceEntryHref;
        const name = selectedRow.dataset.evidenceEntryName;
        if (href && canPreview(href)) openPreview(selectedRow, href, name);
        return;
      }
    }
    if (!panel.classList.contains("active")) return;
    if (e.key === "Escape") { closePreview(); return; }
    if (e.key === "+" || e.key === "=") setZoom(zoomLevel * 1.25);
    if (e.key === "-") setZoom(zoomLevel * 0.8);
    if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      e.preventDefault();
      const rows = [...document.querySelectorAll(".evidence-file-row:not(.evidence-file-row--folder):not(.head)")];
      const idx = rows.indexOf(selectedRow);
      const next = e.key === "ArrowDown" ? rows[idx + 1] : rows[idx - 1];
      if (next) {
        const link = next.querySelector("a[href]");
        if (link && canPreview(link.href)) { next.scrollIntoView({ block: "nearest" }); openPreview(next, link.href, link.textContent.trim()); }
      }
    }
  }, true);
})();

document.querySelector("#docs-view").addEventListener("click", (event) => {
  const btn = event.target.closest("[data-market]");
  if (!btn) return;
  const market = btn.dataset.market;
  if (!market || activeMarketFilter === market) return;
  const prevGroup = activeEvidenceGroup;
  activeMarketFilter = market;
  // War ein Ordner offen, den gleichen Ordner im neuen Markt (IEC↔UL) öffnen.
  // Abgleich über die führende Nummer, da IEC und UL denselben Abschnitt oft
  // unterschiedlich benennen (z.B. IEC "03 User Manual" vs UL "03 Offerte …").
  const folderNo = (s) => (String(s || "").match(/^\s*0*(\d+)/) || [])[1];
  const prevNo = folderNo(prevGroup);
  const sameFolder = prevNo && (activeProject.documentGroups || []).find(
    (g) => (g.area || "").startsWith(market + " / ") && folderNo(g.primary) === prevNo
  );
  if (sameFolder) {
    // Auswahl auf den Ordner des neuen Markts setzen (anderer Name, gleiche Nummer)
    activeEvidenceGroup = sameFolder.primary;
    // frisch vom markt-spezifischen href laden (per-href gecacht)
    evidenceEntries.delete(evidenceCacheKey(activeProject.id, sameFolder.primary));
    renderDocs(activeProject);
    loadEvidenceEntries(sameFolder, evidenceHref(sameFolder));
  } else {
    activeEvidenceGroup = null;
    renderDocs(activeProject);
  }
});

// Sidebar (ganz links): ein-/ausblenden und Breite verschieben
(() => {
  const toggle = document.querySelector("#sidebar-toggle");
  const resizer = document.querySelector("#sidebar-resizer");
  if (!toggle || !resizer) return;
  const setVar = (px) => document.documentElement.style.setProperty("--sb-w", `${px}px`);

  toggle.addEventListener("click", () => {
    const collapsed = document.body.classList.toggle("sb-collapsed");
    toggle.textContent = collapsed ? "»" : "«";
  });

  let dragging = false;
  resizer.addEventListener("mousedown", (event) => {
    if (document.body.classList.contains("sb-collapsed")) return;
    dragging = true;
    document.body.style.userSelect = "none";
    document.body.style.cursor = "col-resize";
    event.preventDefault();
  });
  document.addEventListener("mousemove", (event) => {
    if (!dragging) return;
    // Breite = Mausposition von links; begrenzt auf einen sinnvollen Bereich
    const w = Math.max(180, Math.min(520, event.clientX));
    setVar(w);
  });
  document.addEventListener("mouseup", () => {
    if (!dragging) return;
    dragging = false;
    document.body.style.userSelect = "";
    document.body.style.cursor = "";
  });
})();

document.querySelector("#docs-view").addEventListener("keydown", (event) => {
  if (!["Enter", " "].includes(event.key)) return;
  const groupCard = event.target.closest("[data-evidence-group-card]");
  if (!groupCard) return;
  event.preventDefault();
  const group = activeProject.documentGroups?.find((item) => item.primary === groupCard.dataset.evidenceGroupCard);
  toggleEvidenceGroup(groupCard.dataset.evidenceGroupCard);
  renderDocs(activeProject);
  if (activeEvidenceGroup && group) loadEvidenceEntries(group);
});

document.body.addEventListener("click", async (event) => {
  const btn = event.target.closest("[data-open-office-href]");
  if (!btn) return;
  event.preventDefault();
  btn.disabled = true;
  try {
    const href = new URL(btn.dataset.openOfficeHref, window.location.href).pathname;
    await apiFetch("/api/open-path", {
      method: "POST",
      body: JSON.stringify({ href, app: officeApp(href) })
    });
  } catch (err) {
    console.error("Open Office error:", err);
  } finally {
    btn.disabled = false;
  }
});

// Tab buttons
// ── Packaging / Versand ───────────────────────────────────────────────────
let packagingShipments = [];
let packagingLabs = [];
let openShipments = [];

const SHIP_STATUSES = ["Geplant", "Versendet", "Unterwegs", "Angekommen", "Zurückgesendet"];
const SHIP_STATUS_CLASS = {
  Geplant: "planned", Versendet: "sent", Unterwegs: "transit",
  Angekommen: "arrived", Zurückgesendet: "returned"
};

const CHECKLIST_OUT = [
  "Proforma-Rechnung / Commercial Invoice",
  "Packing List (Inhalt, Gewicht, Abmessungen)",
  "Technische Dokumentation (Safety, EMC, BOM)",
  "Versandanweisung an Vertrieb / PM",
  "Muster korrekt beschriftet (EF-Nr., Build, Datum)",
  "Gerät funktionsfähig und vollständig",
  "Rücksendeinformationen beilegen (Adresse, Kontakt)",
  "Zollwert deklariert (Muster = kein Handelswert)",
];
const CHECKLIST_IN = [
  "Tracking-Nummer erfasst",
  "Empfang bestätigt (Labor kontaktieren falls nötig)",
  "Muster auf Transportschäden geprüft",
  "Prüfmuster-Status im Projekt aktualisieren",
  "Eingangsbestätigung an PM / Projektleitung",
];

async function loadPackagingData() {
  [packagingShipments, packagingLabs, openShipments] = await Promise.all([
    apiFetch(`/api/projects/${activeProject.id}/shipments`),
    apiFetch("/api/labs"),
    apiFetch("/api/shipments/open")
  ]);
}

function renderPackaging() {
  const el = document.querySelector("#packaging-container");

  const statusBadge = (s) =>
    `<span class="ship-status ${SHIP_STATUS_CLASS[s] || ""}">${s}</span>`;

  const dirIcon = (d) => d === "in" ? "← Eingehend" : "→ Ausgehend";

  // ── Open shipments across all projects ──────────────────
  const openRows = openShipments.length
    ? openShipments.map((s) => `
        <div class="ship-row">
          <span class="ship-project">${escapeHtml(s.project_id)}</span>
          <span>${escapeHtml(s.build || "—")}</span>
          <span class="dir-tag ${s.direction === "in" ? "dir-in" : "dir-out"}">${dirIcon(s.direction)}</span>
          <span>${escapeHtml(s.destination || "—")}</span>
          <span class="muted">${escapeHtml(s.sent_date || "—")}</span>
          <span class="muted">${escapeHtml(s.expected_date || "—")}</span>
          ${statusBadge(s.status)}
          <span class="muted ship-notes">${escapeHtml(s.notes || "")}</span>
        </div>`).join("")
    : `<p class="empty-state">Keine offenen Versände.</p>`;

  // ── Shipments for current project ────────────────────────
  const projRows = packagingShipments.length
    ? packagingShipments.map((s) => `
        <div class="ship-row" data-ship-id="${s.id}">
          <span>${escapeHtml(s.build || "—")}</span>
          <span class="dir-tag ${s.direction === "in" ? "dir-in" : "dir-out"}">${dirIcon(s.direction)}</span>
          <span>${escapeHtml(s.destination || "—")}</span>
          <span class="muted">${escapeHtml(s.sent_date || "—")}</span>
          <span class="muted">${escapeHtml(s.expected_date || "—")}</span>
          ${statusBadge(s.status)}
          <span class="muted ship-notes">${escapeHtml(s.notes || "")}</span>
          <span class="ship-actions">
            <button class="finder-action" type="button" data-ship-edit="${s.id}">Bearbeiten</button>
            <button class="finder-action" type="button" data-ship-delete="${s.id}">Löschen</button>
          </span>
        </div>`).join("")
    : `<p class="empty-state">Noch keine Versände für dieses Projekt.</p>`;

  // ── Lab cards ─────────────────────────────────────────────
  const labCards = packagingLabs.map((lab) => `
    <div class="lab-card">
      <div class="lab-top">
        <strong>${escapeHtml(lab.short_name)}</strong>
        <span class="lab-country">${escapeHtml(lab.country)}</span>
      </div>
      <p class="lab-name">${escapeHtml(lab.name)}</p>
      ${lab.address ? `<p class="lab-addr">${escapeHtml(lab.address)}</p>` : ""}
      ${lab.contact ? `<p class="lab-contact">Kontakt: ${escapeHtml(lab.contact)}</p>` : ""}
      ${lab.email   ? `<a class="lab-link" href="mailto:${escapeHtml(lab.email)}">${escapeHtml(lab.email)}</a>` : ""}
      ${lab.phone   ? `<p class="lab-contact">${escapeHtml(lab.phone)}</p>` : ""}
      ${lab.notes   ? `<p class="lab-notes">${escapeHtml(lab.notes)}</p>` : ""}
      <button class="finder-action" type="button" data-lab-edit="${lab.id}" style="margin-top:8px">Bearbeiten</button>
    </div>`).join("");

  el.innerHTML = `
    <div class="packaging-layout">

      <div class="pkg-panel pkg-panel-wide">
        <div class="panel-header"><h2>Offene Versände — Alle Projekte</h2><span class="muted">${openShipments.length} offen</span></div>
        <div class="ship-table">
          ${openShipments.length ? `
          <div class="ship-row ship-head">
            <span>Projekt</span><span>Muster</span><span>Richtung</span><span>Labor</span>
            <span>Versendet</span><span>Erwartet</span><span>Status</span><span>Notiz</span>
          </div>` : ""}
          ${openRows}
        </div>
      </div>

      <div class="pkg-two-col">
        <div class="pkg-panel">
          <div class="panel-header">
            <h2>Musterversand — ${escapeHtml(activeProject.id)}</h2>
            <button class="finder-action" type="button" id="add-shipment-btn">+ Versand erfassen</button>
          </div>
          <div id="shipment-form-wrap"></div>
          <div class="ship-table">
            ${packagingShipments.length ? `
            <div class="ship-row ship-head">
              <span>Muster</span><span>Richtung</span><span>Labor</span>
              <span>Versendet</span><span>Erwartet</span><span>Status</span><span>Notiz</span><span></span>
            </div>` : ""}
            ${projRows}
          </div>
        </div>

        <div class="pkg-panel">
          <div class="panel-header"><h2>Versand-Checkliste</h2></div>
          <div class="checklist-section">
            <h3>Ausgehend → Labor</h3>
            <ul class="pkg-checklist">
              ${CHECKLIST_OUT.map((item) => `<li><label><input type="checkbox"> ${escapeHtml(item)}</label></li>`).join("")}
            </ul>
          </div>
          <div class="checklist-section">
            <h3>Eingehend ← Rücksendung</h3>
            <ul class="pkg-checklist">
              ${CHECKLIST_IN.map((item) => `<li><label><input type="checkbox"> ${escapeHtml(item)}</label></li>`).join("")}
            </ul>
          </div>
        </div>
      </div>

      <div class="pkg-panel">
        <div class="panel-header">
          <h2>Laborkontakte & Adressen</h2>
          <button class="finder-action" type="button" id="add-lab-btn">+ Labor hinzufügen</button>
        </div>
        <div id="lab-form-wrap"></div>
        <div class="lab-grid">${labCards}</div>
      </div>

    </div>`;
}

function buildShipmentForm(existing = null) {
  const s = existing || {};
  return `
    <form class="pkg-form" id="shipment-form">
      <div class="pkg-form-grid">
        <label>Muster / Build<input name="build" value="${escapeHtml(s.build || "")}" placeholder="z.B. PT1, TS1"></label>
        <label>Richtung
          <select name="direction">
            <option value="out" ${s.direction !== "in" ? "selected" : ""}>→ Ausgehend</option>
            <option value="in"  ${s.direction === "in"  ? "selected" : ""}>← Eingehend</option>
          </select>
        </label>
        <label>Labor / Ziel<input name="destination" value="${escapeHtml(s.destination || "")}" placeholder="z.B. VDE, UL-IT"></label>
        <label>Status
          <select name="status">
            ${SHIP_STATUSES.map((st) => `<option ${s.status === st ? "selected" : ""}>${st}</option>`).join("")}
          </select>
        </label>
        <label>Versanddatum<input name="sent_date" type="date" value="${s.sent_date || ""}"></label>
        <label>Erwartet am<input name="expected_date" type="date" value="${s.expected_date || ""}"></label>
        <label>Erhalten am<input name="received_date" type="date" value="${s.received_date || ""}"></label>
        <label>Tracking<input name="tracking" value="${escapeHtml(s.tracking || "")}" placeholder="Sendungsnummer"></label>
        <label class="full-width">Notiz<input name="notes" value="${escapeHtml(s.notes || "")}" placeholder="Weitere Infos"></label>
      </div>
      <div class="pkg-form-actions">
        <button type="submit" class="status-toggle done">${existing ? "Speichern" : "Erfassen"}</button>
        <button type="button" id="cancel-shipment-btn" class="finder-action">Abbrechen</button>
      </div>
      ${existing ? `<input type="hidden" name="id" value="${existing.id}">` : ""}
    </form>`;
}

function buildLabForm(existing = null) {
  const l = existing || {};
  return `
    <form class="pkg-form" id="lab-form">
      <div class="pkg-form-grid">
        <label>Kürzel<input name="short_name" value="${escapeHtml(l.short_name || "")}" placeholder="z.B. VDE"></label>
        <label>Land<input name="country" value="${escapeHtml(l.country || "")}" placeholder="DE, IT, TW …"></label>
        <label class="full-width">Name<input name="name" value="${escapeHtml(l.name || "")}" placeholder="Offizieller Name des Labors"></label>
        <label class="full-width">Adresse<input name="address" value="${escapeHtml(l.address || "")}" placeholder="Strasse, PLZ, Ort"></label>
        <label>Kontaktperson<input name="contact" value="${escapeHtml(l.contact || "")}"></label>
        <label>E-Mail<input name="email" type="email" value="${escapeHtml(l.email || "")}"></label>
        <label>Telefon<input name="phone" value="${escapeHtml(l.phone || "")}"></label>
        <label class="full-width">Notiz<input name="notes" value="${escapeHtml(l.notes || "")}" placeholder="Welche Prüfungen, Besonderheiten …"></label>
      </div>
      <div class="pkg-form-actions">
        <button type="submit" class="status-toggle done">${existing ? "Speichern" : "Hinzufügen"}</button>
        <button type="button" id="cancel-lab-btn" class="finder-action">Abbrechen</button>
      </div>
      ${existing ? `<input type="hidden" name="id" value="${existing.id}">` : ""}
    </form>`;
}

function formData(form) {
  return Object.fromEntries(new FormData(form).entries());
}

document.querySelector("#packaging-view").addEventListener("click", async (e) => {
  // Add shipment
  if (e.target.id === "add-shipment-btn") {
    document.querySelector("#shipment-form-wrap").innerHTML = buildShipmentForm();
    return;
  }
  if (e.target.id === "cancel-shipment-btn") {
    document.querySelector("#shipment-form-wrap").innerHTML = "";
    return;
  }
  // Edit shipment
  const editId = e.target.dataset.shipEdit;
  if (editId) {
    const s = packagingShipments.find((x) => String(x.id) === editId);
    if (s) document.querySelector("#shipment-form-wrap").innerHTML = buildShipmentForm(s);
    return;
  }
  // Delete shipment
  const delId = e.target.dataset.shipDelete;
  if (delId) {
    await apiFetch(`/api/projects/${activeProject.id}/shipments/${delId}`, { method: "DELETE" });
    await loadPackagingData();
    renderPackaging();
    return;
  }
  // Add lab
  if (e.target.id === "add-lab-btn") {
    document.querySelector("#lab-form-wrap").innerHTML = buildLabForm();
    return;
  }
  if (e.target.id === "cancel-lab-btn") {
    document.querySelector("#lab-form-wrap").innerHTML = "";
    return;
  }
  // Edit lab
  const labEditId = e.target.dataset.labEdit;
  if (labEditId) {
    const lab = packagingLabs.find((x) => String(x.id) === labEditId);
    if (lab) document.querySelector("#lab-form-wrap").innerHTML = buildLabForm(lab);
    return;
  }
});

document.querySelector("#packaging-view").addEventListener("submit", async (e) => {
  e.preventDefault();
  if (e.target.id === "shipment-form") {
    const data = formData(e.target);
    if (data.id) {
      await apiFetch(`/api/projects/${activeProject.id}/shipments/${data.id}`, { method: "PUT", body: JSON.stringify(data) });
    } else {
      await apiFetch(`/api/projects/${activeProject.id}/shipments`, { method: "POST", body: JSON.stringify(data) });
    }
    await loadPackagingData();
    renderPackaging();
    return;
  }
  if (e.target.id === "lab-form") {
    const data = formData(e.target);
    if (data.id) {
      await apiFetch(`/api/labs/${data.id}`, { method: "PUT", body: JSON.stringify(data) });
    } else {
      await apiFetch("/api/labs", { method: "POST", body: JSON.stringify(data) });
    }
    await loadPackagingData();
    renderPackaging();
  }
});

buttons.overview.addEventListener("click", () => setView("overview"));
buttons.subtopic.addEventListener("click", () => setView("subtopic"));
buttons.tasks.addEventListener("click", () => setView("tasks"));
buttons.docs.addEventListener("click", () => setView("docs"));
buttons.calculations.addEventListener("click", () => {
  setView("calculations");
  renderCalculations();
});
buttons.freigabe.addEventListener("click", () => setView("freigabe"));
buttons.packaging.addEventListener("click", async () => {
  setView("packaging");
  await loadPackagingData();
  renderPackaging();
});
function saveSimpleCalcInput(input) {
  if (!input) return;
  const overrides = loadCalcOverrides();
  overrides[`simple:${input.dataset.simpleCalc}`] = calcParseInput(input.value);
  saveCalcOverrides(overrides);
  renderCalculations();
}

function saveWarmthInput(input) {
  if (!input) return;
  const overrides = loadWarmthOverrides();
  overrides[warmthKey(Number(input.dataset.warmthSet), input.dataset.warmthField)] = parseWarmthValue(input.value);
  saveWarmthOverrides(overrides);
  renderCalculations();
}

calculationsContainer.addEventListener("change", (event) => {
  const input = event.target.closest("[data-simple-calc]");
  if (!input) return;
  saveSimpleCalcInput(input);
});
calculationsContainer.addEventListener("keydown", (event) => {
  if (event.key !== "Enter") return;
  const input = event.target.closest("[data-simple-calc]");
  if (!input) return;
  event.preventDefault();
  saveSimpleCalcInput(input);
});
calcReset.addEventListener("click", () => {
  if (!confirm("Alle Berechnungs-Eingaben zurücksetzen?")) return;
  localStorage.removeItem(calcStorageKey);
  renderCalculations();
});
calculationsContainer.addEventListener("change", (event) => {
  const warmthInput = event.target.closest("[data-warmth-set][data-warmth-field]");
  if (warmthInput) {
    saveWarmthInput(warmthInput);
  }
});
calculationsContainer.addEventListener("keydown", (event) => {
  const warmthInput = event.target.closest("[data-warmth-set][data-warmth-field]");
  if (!warmthInput || event.key !== "Enter") return;
  event.preventDefault();
  saveWarmthInput(warmthInput);
});
themeToggle.addEventListener("click", () => {
  applyTheme(document.body.dataset.theme === "dark" ? "light" : "dark");
  // Push the new theme to the Tabelle 24/30 iframes so they repaint without losing state.
  const msg = { type: "theme", theme: document.body.dataset.theme };
  for (const frame of [tabelle24Frame, tabelle30Frame]) {
    if (frame?.contentWindow) {
      try { frame.contentWindow.postMessage(msg, window.location.origin); } catch {}
    }
  }
});

// ── Scanner UI ────────────────────────────────────────────────
// ── Excel sidebar ─────────────────────────────────────────────
const excelSidebarBtn = document.querySelector("#excel-sidebar-btn");
const excelSidebarList = document.querySelector("#excel-sidebar-list");
let excelFilesCache = null;

excelSidebarBtn.addEventListener("click", async () => {
  const isOpen = !excelSidebarList.classList.contains("hidden");
  if (isOpen) { excelSidebarList.classList.add("hidden"); return; }
  excelSidebarList.classList.remove("hidden");
  if (excelFilesCache !== null) return;
  excelSidebarList.innerHTML = `<span class="excel-loading">Wird geladen…</span>`;
  try {
    const files = await apiFetch(`/api/projects/${activeProject.id}/excel-files`);
    excelFilesCache = files;
    excelSidebarList.innerHTML = files.length
      ? files.map((f) => `<button class="excel-file-btn" type="button" title="${escapeHtml(f.href)}">${escapeHtml(f.name)}</button>`).join("")
      : `<span class="excel-loading">Keine Excel-Dateien gefunden.</span>`;
    // Store hrefs on the buttons directly
    excelSidebarList.querySelectorAll(".excel-file-btn").forEach((btn, i) => {
      btn._excelHref = files[i].href;
    });
  } catch {
    excelSidebarList.innerHTML = `<span class="excel-loading">Fehler beim Laden.</span>`;
  }
});

excelSidebarList.addEventListener("click", async (e) => {
  const btn = e.target.closest(".excel-file-btn");
  if (!btn || !btn._excelHref) return;
  btn.style.opacity = "0.5";
  try {
    await apiFetch("/api/open-path", {
      method: "POST",
      body: JSON.stringify({ href: btn._excelHref, app: "excel" })
    });
  } catch (err) {
    console.error("Excel open error:", err);
  } finally {
    btn.style.opacity = "";
  }
});

const scanBtn = document.querySelector("#scan-btn");
const scanStatus = document.querySelector("#scan-status");

function showScanDone(s) {
  if (!s?.scanned_at) { scanStatus.textContent = "Noch nicht gescannt"; return; }
  const d = new Date(s.scanned_at);
  const fmt = d.toLocaleDateString("de-CH", { day: "2-digit", month: "2-digit" })
    + " " + d.toLocaleTimeString("de-CH", { hour: "2-digit", minute: "2-digit" });
  scanStatus.textContent = `${fmt} · ${s.projects_found} Proj.`;
}

async function startScanStream() {
  // In einem Projekt → nur dieses scannen; auf dem Dashboard → alle.
  const onlyId = activeView !== "dashboard" && activeProject ? activeProject.id : null;
  scanBtn.disabled = true;
  scanStatus.textContent = onlyId ? `Scanne ${onlyId}…` : "Scanne…";
  await apiFetch("/api/scan/start", {
    method: "POST",
    body: JSON.stringify(onlyId ? { projectId: onlyId } : {})
  });
  while (true) {
    await new Promise((r) => setTimeout(r, 500));
    const s = await apiFetch("/api/scan/status");
    if (s.in_progress) {
      if (s.progress?.total > 0) {
        scanStatus.textContent = `${s.progress.done}/${s.progress.total}  ${s.progress.current}`;
      }
    } else {
      scanBtn.disabled = false;
      projectList = await apiFetch("/api/projects");
      if (activeProject) activeProject = await apiFetch(`/api/projects/${activeProject.id}`);
      if (activeView === "dashboard") { renderMachines(); renderDashboard(); } else renderAll();
      if (onlyId) scanStatus.textContent = `✓ ${onlyId} aktualisiert`;
      else showScanDone(s);
      break;
    }
  }
}

scanBtn.addEventListener("click", startScanStream);

// ── Archiv-Excel sync ─────────────────────────────────────────
const archiveSyncBtn = document.querySelector("#archive-sync-btn");

async function runArchiveSync() {
  archiveSyncBtn.disabled = true;
  archiveSyncBtn.textContent = "Sync…";
  try {
    const result = await apiFetch("/api/scan-archive", { method: "POST" });
    archiveSyncBtn.textContent = `✓ ${result.updated} Einträge`;
    // Reload current project to pick up new archive_location
    if (activeProject) {
      activeProject = await apiFetch(`/api/projects/${activeProject.id}`);
      renderSummary(activeProject);
    }
    setTimeout(() => { archiveSyncBtn.textContent = "Archiv sync"; archiveSyncBtn.disabled = false; }, 3000);
  } catch (e) {
    archiveSyncBtn.textContent = "Fehler";
    archiveSyncBtn.disabled = false;
    console.error(e);
  }
}

archiveSyncBtn.addEventListener("click", runArchiveSync);

// ── Sitzungsexport ────────────────────────────────────────────
document.querySelectorAll("[id^='shortcut-btn-']").forEach((btn) => {
  const idx = btn.id.split("-").pop();
  const pathKey = `pcs-shortcut-path-${idx}`;
  const labelKey = `pcs-shortcut-label-${idx}`;
  const savedLabel = localStorage.getItem(labelKey);
  if (savedLabel) btn.textContent = `📁 ${savedLabel}`;

  btn.addEventListener("click", () => {
    const storedPath = localStorage.getItem(pathKey);
    if (storedPath) {
      fetch("/api/open-folder", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ path: storedPath }) })
        .catch(console.error);
      return;
    }
    // First use — ask for label then path
    const labelInput = document.createElement("input");
    labelInput.style.cssText = "font-size:12px;padding:4px 8px;border-radius:6px;border:1.5px solid var(--accent);background:var(--panel);color:var(--ink);width:120px;outline:none;";
    labelInput.placeholder = "Name…";
    labelInput.value = localStorage.getItem(labelKey) || "";
    const pathInput = document.createElement("input");
    pathInput.style.cssText = "font-size:12px;padding:4px 8px;border-radius:6px;border:1.5px solid var(--accent);background:var(--panel);color:var(--ink);width:260px;outline:none;";
    pathInput.placeholder = "Pfad…";
    const wrap = document.createElement("span");
    wrap.style.cssText = "display:inline-flex;gap:4px;align-items:center;";
    wrap.appendChild(labelInput);
    wrap.appendChild(pathInput);
    btn.replaceWith(wrap);
    labelInput.focus();
    const confirm = () => {
      const label = labelInput.value.trim();
      const path = pathInput.value.trim();
      if (label) localStorage.setItem(labelKey, label);
      if (path) {
        localStorage.setItem(pathKey, path);
        fetch("/api/open-folder", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ path }) })
          .catch(console.error);
      }
      if (label) btn.textContent = `📁 ${label}`;
      wrap.replaceWith(btn);
    };
    pathInput.addEventListener("blur", confirm);
    pathInput.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); pathInput.blur(); }
      if (e.key === "Escape") { wrap.replaceWith(btn); }
    });
    labelInput.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); pathInput.focus(); }
      if (e.key === "Escape") { wrap.replaceWith(btn); }
    });
  });
});

document.querySelectorAll("[id^='link-btn-']").forEach((btn) => {
  const idx = btn.id.split("-").pop();
  const urlKey = `pcs-link-url-${idx}`;
  const labelKey = `pcs-link-label-${idx}`;
  const savedLabel = localStorage.getItem(labelKey);
  if (savedLabel) btn.textContent = `🔗 ${savedLabel}`;

  btn.addEventListener("click", () => {
    const storedUrl = localStorage.getItem(urlKey);
    if (storedUrl) { window.open(storedUrl, "_blank"); return; }
    // First use — ask for label then URL
    const labelInput = document.createElement("input");
    labelInput.style.cssText = "font-size:12px;padding:4px 8px;border-radius:6px;border:1.5px solid var(--accent);background:var(--panel);color:var(--ink);width:120px;outline:none;";
    labelInput.placeholder = "Name…";
    labelInput.value = localStorage.getItem(labelKey) || "";
    const urlInput = document.createElement("input");
    urlInput.style.cssText = "font-size:12px;padding:4px 8px;border-radius:6px;border:1.5px solid var(--accent);background:var(--panel);color:var(--ink);width:260px;outline:none;";
    urlInput.placeholder = "https://…";
    const wrap = document.createElement("span");
    wrap.style.cssText = "display:inline-flex;gap:4px;align-items:center;";
    wrap.appendChild(labelInput);
    wrap.appendChild(urlInput);
    btn.replaceWith(wrap);
    labelInput.focus();
    const confirm = () => {
      const label = labelInput.value.trim();
      const url = urlInput.value.trim();
      if (label) { localStorage.setItem(labelKey, label); btn.textContent = `🔗 ${label}`; }
      if (url) { localStorage.setItem(urlKey, url); window.open(url, "_blank"); }
      wrap.replaceWith(btn);
    };
    urlInput.addEventListener("blur", confirm);
    urlInput.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); urlInput.blur(); }
      if (e.key === "Escape") { wrap.replaceWith(btn); }
    });
    labelInput.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); urlInput.focus(); }
      if (e.key === "Escape") { wrap.replaceWith(btn); }
    });
  });
});

const exportBtn = document.querySelector("#export-btn");

exportBtn.addEventListener("click", async () => {
  exportBtn.disabled = true;
  exportBtn.textContent = "Wird erstellt…";
  try {
    const projects = await Promise.all(projectList.map((p) => apiFetch(`/api/projects/${p.id}`)));
    const win = window.open("", "_blank");
    win.document.write(buildExportHTML(projects));
    win.document.close();
  } finally {
    exportBtn.disabled = false;
    exportBtn.textContent = "Sitzungsexport";
  }
});

function buildExportHTML(projects) {
  const now = new Date();
  const dateStr = now.toLocaleDateString("de-CH", { weekday: "long", year: "numeric", month: "long", day: "numeric" });
  const isoDate = now.toISOString().slice(0, 10);
  const active = projects.filter((p) => !["Abgeschlossen", "Archiviert"].includes(p.phase));

  const healthColor = { "Good": "#22754b", "Watch": "#9a6615", "Risk": "#a33d37" };
  const statusDE = { "Open": "Offen", "Done": "Erledigt", "Blocked": "Blockiert" };

  // Editable inline field
  const field = (id, placeholder, multiline = false, value = "") => {
    if (multiline) return `<div class="editable" contenteditable="true" data-field="${id}" data-placeholder="${placeholder}">${escapeHtml(value)}</div>`;
    return `<span class="editable inline" contenteditable="true" data-field="${id}" data-placeholder="${placeholder}">${escapeHtml(value)}</span>`;
  };

  const projectHTML = active.map((p, idx) => {
    const allTasks = (p.tasks || []);
    const openTasks = allTasks.filter((t) => t.status !== "Done");
    const risks = p.risks || [];
    const openCerts = (p.certification || []).filter((c) => !c.done && !["Done", "Not needed", "Abgeschlossen"].includes(c.state));
    const builds = p.builds || [];
    const hColor = healthColor[p.health] || "#9a6615";
    const traktandumNr = idx + 1;

    const buildsText = builds.map((b) => `${b.label} (${b.date || "TBD"}) — ${b.state}`).join(", ");

    const tasksTable = openTasks.length ? `
      <table class="tasks-table">
        <thead>
          <tr><th>Massnahme</th><th>Verantwortlich</th><th>Termin</th><th>Status</th></tr>
        </thead>
        <tbody>
          ${openTasks.map((t) => `
            <tr class="${t.status === "Blocked" ? "row-bad" : t.status === "Done" ? "row-done" : ""}">
              <td>${escapeHtml(t.task)}${t.block_reason ? `<div class="block-reason">Blocker: ${escapeHtml(t.block_reason)}</div>` : ""}</td>
              <td class="cell-owner">${escapeHtml(t.owner || "—")}</td>
              <td class="cell-due">${escapeHtml(t.due || "—")}</td>
              <td class="cell-status"><span class="stag ${t.status === "Blocked" ? "bad" : t.status === "Done" ? "good" : "watch"}">${statusDE[t.status] || t.status}</span></td>
            </tr>`).join("")}
        </tbody>
      </table>` : `<p class="none">Keine offenen Massnahmen.</p>`;

    const risksHTML = risks.length ? `
      <div class="sub-section">
        <h4>PCS Blocker / Risiken</h4>
        ${risks.map((r) => `<div class="risk-row"><span class="stag bad">${r.level}</span> ${escapeHtml(r.text)}</div>`).join("")}
      </div>` : "";

    const certsHTML = openCerts.length ? `
      <div class="sub-section">
        <h4>Zulassung offen (${openCerts.length})</h4>
        <ul>${openCerts.map((c) => `<li>${escapeHtml(c.name)}</li>`).join("")}</ul>
      </div>` : "";

    return `
      <div class="project" id="proj-${p.id}">
        <div class="proj-header" style="border-left-color:${hColor}">
          <div class="proj-title">
            <span class="trak-nr">${traktandumNr}.</span>
            <span class="proj-id">${escapeHtml(p.id)}</span>
            <span class="proj-name">${escapeHtml(p.name || "")}</span>
          </div>
          <div class="proj-meta">
            <span class="badge" style="background:${hColor}22;color:${hColor}">${p.health || "Watch"}</span>
            <span class="muted">${escapeHtml(p.phase || "")}</span>
            ${p.target ? `<span class="muted">Ziel: ${escapeHtml(p.target)}</span>` : ""}
            ${buildsText ? `<span class="muted">${escapeHtml(buildsText)}</span>` : ""}
          </div>
        </div>
        <div class="proj-body">
          <div class="verlauf-section">
            <h3>Verlauf / Status</h3>
            ${field(`verlauf-${p.id}`, "Statusbericht, Verlauf, Rückfragen, Entscheidungen …", true)}
          </div>
          <div class="massnahmen-section">
            <h3>PCS Massnahmen (${openTasks.length} offen)</h3>
            ${tasksTable}
          </div>
          ${risksHTML}${certsHTML}
        </div>
      </div>`;
  }).join("");

  const teamMembers = ["A. Kaufmann", "B. Meier", "C. Müller", "D. Schneider"];
  const ferienRows = teamMembers.map((name) => `
    <tr>
      <td class="cell-owner">${name}</td>
      <td>${field(`ferien-${name.replace(/[^a-z]/gi, "")}`, "z.B. 25.–29. Aug.", false)}</td>
    </tr>`).join("");

  return `<!doctype html>
<html lang="de">
<head>
  <meta charset="utf-8">
  <title>PCS Aktivitätensitzung ${dateStr}</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: "Segoe UI", Arial, sans-serif; font-size: 11px; color: #172027; background: white; padding: 18mm 20mm; max-width: 260mm; margin: 0 auto; }
    h1 { font-size: 20px; font-weight: 800; margin-bottom: 2px; }
    h2 { font-size: 13px; font-weight: 800; margin: 0 0 10px; }
    h3 { font-size: 10px; font-weight: 800; text-transform: uppercase; letter-spacing: 0.05em; color: #63707a; margin-bottom: 6px; }
    h4 { font-size: 10px; font-weight: 700; color: #63707a; margin-bottom: 4px; }
    .doc-meta { color: #63707a; font-size: 12px; margin-bottom: 4px; }
    .doc-date { color: #172027; font-size: 13px; font-weight: 600; margin-bottom: 20px; }

    /* Editable fields */
    .editable { border-bottom: 1px dashed #b0bec5; min-height: 18px; outline: none; padding: 2px 2px; border-radius: 2px; width: 100%; display: block; }
    .editable:empty::before { content: attr(data-placeholder); color: #aab2ba; font-style: italic; }
    .editable:focus { background: #f0f8ff; border-bottom-color: #4a90d9; }
    .editable.inline { display: inline-block; width: auto; min-width: 80px; }

    /* Participants block */
    .participants-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-bottom: 24px; padding: 12px 16px; border: 1px solid #d6dde2; border-radius: 6px; background: #f9fafb; }
    .participants-grid h3 { margin-bottom: 4px; }

    /* Summary bar */
    .summary { display: flex; gap: 16px; margin-bottom: 24px; padding: 10px 16px; background: #f3f6f8; border-radius: 6px; }
    .summary div { text-align: center; flex: 1; }
    .summary strong { display: block; font-size: 20px; font-weight: 800; }
    .summary span { font-size: 9px; color: #63707a; text-transform: uppercase; letter-spacing: 0.04em; }

    /* Projects */
    .project { margin-bottom: 24px; border: 1px solid #d6dde2; border-radius: 6px; overflow: hidden; page-break-inside: avoid; }
    .proj-header { display: flex; justify-content: space-between; align-items: flex-start; padding: 8px 14px; background: #f3f6f8; border-left: 5px solid #9a6615; gap: 12px; }
    .proj-title { display: flex; align-items: baseline; gap: 8px; }
    .trak-nr { font-size: 13px; font-weight: 800; color: #8a96a0; }
    .proj-id { font-size: 14px; font-weight: 800; }
    .proj-name { font-size: 12px; color: #52606b; }
    .proj-meta { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; flex-shrink: 0; font-size: 10px; }
    .badge { border-radius: 999px; padding: 2px 8px; font-size: 10px; font-weight: 700; }
    .muted { color: #63707a; }
    .proj-body { padding: 0; }
    .verlauf-section { padding: 8px 14px; border-top: 1px solid #eef2f4; }
    .verlauf-section .editable { min-height: 36px; font-size: 11px; line-height: 1.5; }
    .massnahmen-section { padding: 8px 14px; border-top: 1px solid #eef2f4; }
    .sub-section { padding: 8px 14px; border-top: 1px solid #eef2f4; }

    /* Tasks table */
    .tasks-table { width: 100%; border-collapse: collapse; font-size: 10.5px; }
    .tasks-table thead tr { background: #eef2f4; }
    .tasks-table th { text-align: left; padding: 3px 6px; font-size: 9.5px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.03em; color: #63707a; border-bottom: 1px solid #d6dde2; }
    .tasks-table td { padding: 3px 6px 3px 0; vertical-align: top; border-bottom: 1px solid #f0f3f5; }
    .tasks-table tr:last-child td { border-bottom: none; }
    .tasks-table .cell-owner { white-space: nowrap; color: #52606b; width: 100px; }
    .tasks-table .cell-due { white-space: nowrap; color: #52606b; width: 80px; }
    .tasks-table .cell-status { width: 70px; }
    .tasks-table .row-bad td { background: #fff8f8; }
    .tasks-table .row-done td { opacity: 0.5; }
    .block-reason { font-size: 10px; color: #a33d37; margin-top: 2px; font-style: italic; }
    .stag { display: inline-block; border-radius: 4px; padding: 1px 5px; font-size: 9.5px; font-weight: 700; white-space: nowrap; }
    .stag.bad { background: #fce8e7; color: #a33d37; }
    .stag.good { background: #e6f4ee; color: #22754b; }
    .stag.watch { background: #fdf3e3; color: #9a6615; }
    .none { font-size: 10px; color: #9aa5ae; font-style: italic; }
    .risk-row { margin-bottom: 4px; }
    ul { padding-left: 16px; }
    li { margin-bottom: 2px; }

    /* Freetext sections */
    .free-section { margin-bottom: 24px; border: 1px solid #d6dde2; border-radius: 6px; overflow: hidden; page-break-inside: avoid; }
    .free-section-header { padding: 8px 14px; background: #f3f6f8; border-left: 5px solid #c5ced5; }
    .free-section-body { padding: 10px 14px; }
    .free-row { display: grid; grid-template-columns: 160px 1fr; gap: 8px; margin-bottom: 8px; align-items: start; }
    .free-row label { font-weight: 600; font-size: 10.5px; padding-top: 2px; }

    /* Ferien table */
    .ferien-table { width: 100%; border-collapse: collapse; font-size: 10.5px; }
    .ferien-table th { text-align: left; padding: 3px 6px; font-size: 9.5px; font-weight: 700; text-transform: uppercase; color: #63707a; border-bottom: 1px solid #d6dde2; background: #eef2f4; }
    .ferien-table td { padding: 4px 6px 4px 0; border-bottom: 1px solid #f0f3f5; vertical-align: middle; }
    .ferien-table tr:last-child td { border-bottom: none; }

    /* Divider */
    .section-divider { margin: 28px 0 18px; border: none; border-top: 2px solid #d6dde2; }

    @media print {
      body { padding: 10mm 14mm; }
      .editable { border-bottom: 1px solid #d0d8de; }
      .editable:empty::before { content: ""; }
      .no-print { display: none !important; }
    }
  </style>
</head>
<body>
  <p class="doc-meta">PCS Kaffee Dashboard — Aktivitätensitzung</p>
  <h1>PCS Aktivitätensitzung</h1>
  <p class="doc-date">${dateStr}</p>

  <!-- Teilnehmer -->
  <div class="participants-grid">
    <div>
      <h3>Gesprächsteilnehmer</h3>
      ${field("teilnehmer", "Namen eingeben, z.B. A. Kaufmann, B. Meier …", true)}
    </div>
    <div>
      <h3>Entschuldigt</h3>
      ${field("entschuldigt", "Entschuldigte Personen …", true)}
    </div>
  </div>

  <!-- Statistik -->
  <div class="summary">
    <div><strong>${active.length}</strong><span>Projekte</span></div>
    <div><strong>${active.reduce((n, p) => n + (p.tasks || []).filter((t) => t.status !== "Done").length, 0)}</strong><span>Offene Massnahmen</span></div>
    <div><strong>${active.reduce((n, p) => n + (p.risks || []).length, 0)}</strong><span>Risiken / Blocker</span></div>
    <div><strong>${active.filter((p) => p.health === "Risk").length}</strong><span>Projekte im Risiko</span></div>
    <div><strong>${active.filter((p) => p.health === "Good").length}</strong><span>Projekte OK</span></div>
  </div>

  <!-- Traktanden (EF-Projekte) -->
  ${projectHTML}

  <!-- UL Allgemein -->
  <div class="free-section">
    <div class="free-section-header">
      <h2>${active.length + 1}. UL / Allgemein</h2>
    </div>
    <div class="free-section-body">
      ${field("ul-allgemein", "UL-Rückfragen, Laborkontakte, allgemeine Zulassungsthemen …", true)}
    </div>
  </div>

  <!-- Administratives -->
  <div class="free-section">
    <div class="free-section-header">
      <h2>${active.length + 2}. Administratives</h2>
    </div>
    <div class="free-section-body">
      <div class="free-row">
        <label>Schulungen</label>
        ${field("admin-schulungen", "Geplante oder durchgeführte Schulungen …", true)}
      </div>
      <div class="free-row">
        <label>Neue Normen / IDM</label>
        ${field("admin-normen", "Neue Normen, IDM-Einträge, Normänderungen …", true)}
      </div>
      <div class="free-row">
        <label>Diverses</label>
        ${field("admin-diverses", "Sonstige administrative Punkte …", true)}
      </div>
    </div>
  </div>

  <hr class="section-divider">

  <!-- Ferien -->
  <div class="free-section">
    <div class="free-section-header">
      <h2>Ferien / Abwesenheiten</h2>
    </div>
    <div class="free-section-body">
      <table class="ferien-table">
        <thead><tr><th>Person</th><th>Abwesenheit</th></tr></thead>
        <tbody>
          ${["Person 1", "Person 2", "Person 3", "Person 4", "Person 5"].map((name, i) => `
            <tr>
              <td>${field(`ferien-name-${i}`, name, false)}</td>
              <td>${field(`ferien-dates-${i}`, "z.B. 25.–29. Aug.", false)}</td>
            </tr>`).join("")}
        </tbody>
      </table>
    </div>
  </div>

  <!-- Nächste Sitzung -->
  <div class="free-section">
    <div class="free-section-header">
      <h2>Sitzungstermine</h2>
    </div>
    <div class="free-section-body">
      <div class="free-row">
        <label>Nächste Sitzung</label>
        ${field("naechste-sitzung", "z.B. Montag, 09. Juni 2026, 14:00 Uhr", false)}
      </div>
      <div class="free-row">
        <label>Weitere Termine</label>
        ${field("weitere-termine", "z.B. Quartalssitzung, Review-Termin …", true)}
      </div>
    </div>
  </div>

  <p class="muted no-print" style="margin-top:24px;font-size:10px">
    Drucken: Cmd+P / Ctrl+P → Als PDF speichern &nbsp;·&nbsp;
    Felder sind direkt bearbeitbar vor dem Drucken.
  </p>
</body>
</html>`;
}

// ── Init ──────────────────────────────────────────────────────
(async () => {
  projectList = await apiFetch("/api/projects");
  if (projectList.length) {
    const preferred = projectList.find((project) => project.id === "EF1157") || projectList[0];
    activeProject = await apiFetch(`/api/projects/${preferred.id}`);
  }
  setView("dashboard");
  renderMachines();
  renderDashboard();
  document.body.classList.remove("app-loading");
  history.replaceState({ dashboard: true }, "", location.pathname);
})();
