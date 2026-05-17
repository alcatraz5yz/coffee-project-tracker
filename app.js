// ── DOM refs ────────────────────────────────────────────────
const machineList = document.querySelector("#machine-list");
const search = document.querySelector("#search");
const views = {
  dashboard: document.querySelector("#dashboard-view"),
  overview: document.querySelector("#overview-view"),
  subtopic: document.querySelector("#subtopic-view"),
  tasks: document.querySelector("#tasks-view"),
  docs: document.querySelector("#docs-view"),
  freigabe: document.querySelector("#freigabe-view"),
  packaging: document.querySelector("#packaging-view")
};
const buttons = {
  overview: document.querySelector("#view-overview"),
  subtopic: document.querySelector("#view-subtopic"),
  tasks: document.querySelector("#view-tasks"),
  docs: document.querySelector("#view-docs"),
  freigabe: document.querySelector("#view-freigabe"),
  packaging: document.querySelector("#view-packaging")
};
const taskFilter = document.querySelector("#task-filter");
const subtopicFilter = document.querySelector("#subtopic-filter");
const zifferTable = document.querySelector("#ziffer-table");
const themeToggle = document.querySelector("#theme-toggle");
const dashboardLink = document.querySelector("#dashboard-link");

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
let IS_WIN = false;
let FILE_MANAGER_LABEL = "Finder";
apiFetch("/api/platform").then(d => { IS_WIN = d.isWin; FILE_MANAGER_LABEL = d.isWin ? "Explorer" : "Finder"; }).catch(() => {});
let activeSubtopic = "Approbation";
let activeBuild = "Alle";
let activeView = "overview";
let activeEvidenceGroup = null;
const evidenceEntries = new Map();

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
  const key = evidenceCacheKey(activeProject.id, group.primary);
  const href = browseHref || evidenceHref(group);
  const rootHref = evidenceHref(group);
  const prev = evidenceEntries.get(key);
  evidenceEntries.set(key, { loading: true, entries: prev?.entries || [], browseHref: href, rootHref });
  if (!prev?.entries?.length) renderDocs(activeProject);
  try {
    const data = await apiFetch(`/api/list-path?href=${encodeURIComponent(href)}`);
    evidenceEntries.set(key, { loading: false, entries: data.entries || [], browseHref: href, rootHref });
  } catch (err) {
    evidenceEntries.set(key, { loading: false, error: true, entries: [], browseHref: href, rootHref });
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
}

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
  const groups = (project.documentGroups || [])
    .slice()
    .sort((a, b) => String(a.primary || "").localeCompare(String(b.primary || "")));
  const groupDetailMarkup = (group) => {
    const key = evidenceCacheKey(project.id, group.primary);
    const cached = evidenceEntries.get(key);
    const entries = cached?.entries || [];
    const isSubfolder = cached?.browseHref && cached.browseHref !== cached?.rootHref;
    const backBtn = isSubfolder
      ? `<button class="evidence-back-btn" type="button" data-evidence-back="${group.primary}">← Zurück</button>` : "";
    const fileListMarkup = cached?.loading
      ? `${backBtn}<p class="empty-state">Ordnerinhalt wird geladen...</p>`
      : cached?.error
        ? `${backBtn}<p class="empty-state">Ordnerinhalt konnte nicht gelesen werden.</p>`
        : entries.length ? `
          ${backBtn}
          <div class="evidence-file-row head">
            <span>Name</span><span>Typ</span><span>Geändert</span><span>Aktion</span>
          </div>
          ${entries.map((entry) => `
            <div class="evidence-file-row${entry.type === "Ordner" ? " evidence-file-row--folder" : ""}">
              ${entry.type === "Ordner"
                ? entry.empty
                  ? `<span class="evidence-folder-btn evidence-folder-empty">📁 ${entry.name} <em>leer</em></span>`
                  : `<button class="evidence-folder-btn" type="button" data-browse-subfolder="${entry.href}" data-browse-group="${group.primary}">📂 ${entry.name} <em class="folder-count">${entry.childCount} Dateien</em></button>`
                : entry.type === "Datei" && isOfficeFile(entry.name)
                  ? `<button class="word-action word-action--name" type="button" data-preview-href="${entry.href}">${entry.name}</button>`
                  : `<a href="${entry.href}" target="_blank" rel="noreferrer">${entry.name}</a>`}
              <span>${entry.size || entry.type}</span>
              <span>${entry.modified}</span>
              <span class="evidence-file-actions">
                ${entry.type !== "Ordner" ? `<button class="finder-action" type="button" data-open-href="${entry.href}">Öffnen</button>` : `<button class="finder-action" type="button" data-open-href="${entry.href}">${FILE_MANAGER_LABEL}</button>`}
              </span>
            </div>
          `).join("")}
        ` : `${backBtn}<p class="empty-state">Dieser Ordner enthält keine sichtbaren Dateien.</p>`;

    return `
      <section class="document-group-detail">
        <div class="evidence-file-table">${fileListMarkup}</div>
      </section>
    `;
  };

  document.querySelector("#document-group-grid").innerHTML = groups.length
    ? groups.map((group) => {
        const num = parseInt(group.count) || 0;
        const folderLabel = group.primary || termLabel(group.area);
        const isEmpty = num === 0;

        const isSelected = activeEvidenceGroup === group.primary;
        return `
          <article class="document-group-card ${statusClass(group.status)} ${isSelected ? "selected" : ""} ${isEmpty ? "dgc-empty" : ""}"
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
              <button class="finder-action" type="button" data-open-href="${evidenceHref(group)}">Im ${FILE_MANAGER_LABEL} öffnen</button>
            </div>`}
          </article>
          ${isSelected ? groupDetailMarkup(group) : ""}`;
      }).join("")
    : `<p class="empty-state">Für dieses Projekt ist noch kein PCS Nachweisindex angelegt.</p>`;

  document.querySelector("#reports-panel")?.classList.add("hidden");
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
}

async function openProject(projectId, view = "overview", pushHistory = true) {
  recordRecent(projectId);
  activeProject = await apiFetch(`/api/projects/${projectId}`);
  activeSubtopic = "Approbation";
  activeBuild = "Alle";
  activeEvidenceGroup = null;
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

// Open local folder in Finder/Explorer via local backend
document.querySelector("#docs-view").addEventListener("click", async (event) => {
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

  const groupCard = event.target.closest("[data-evidence-group-card]");
  if (groupCard && !event.target.closest("a, button")) {
    const wasOpen = activeEvidenceGroup === groupCard.dataset.evidenceGroupCard;
    toggleEvidenceGroup(groupCard.dataset.evidenceGroupCard);
    if (!wasOpen && activeEvidenceGroup) {
      history.replaceState({ projectId: activeProject.id, view: "docs" }, "", `#${activeProject.id}`);
      history.pushState({ projectId: activeProject.id, view: "docs", openGroup: activeEvidenceGroup }, "", `#${activeProject.id}`);
    }
    const group = activeProject.documentGroups?.find((item) => item.primary === groupCard.dataset.evidenceGroupCard);
    renderDocs(activeProject);
    if (activeEvidenceGroup && group) loadEvidenceEntries(group);
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

// ── File preview panel ────────────────────────────────────────
(function () {
  const panel = document.querySelector("#file-preview-panel");
  const nameEl = document.querySelector("#file-preview-name");
  const body = document.querySelector("#file-preview-body");
  const zoomBar = document.querySelector("#file-preview-zoom");
  const closeBtn = document.querySelector("#file-preview-close");
  const IMAGE_EXT = /\.(jpe?g|png|gif|webp|svg|bmp)$/i;
  const PDF_EXT = /\.pdf$/i;
  const OFFICE_EXT = /\.(docx|xlsx|xls|xlsm|xlsb)$/i;
  let zoomLevel = 1;
  let previewType = "image";
  let selectedRow = null;
  let hoveredRow = null;

  function canPreview(href) { return IMAGE_EXT.test(href) || PDF_EXT.test(href) || OFFICE_EXT.test(href); }

  function setZoom(z) {
    zoomLevel = Math.min(5, Math.max(0.2, z));
    const img = body.querySelector("img");
    if (img) { img.style.transform = `scale(${zoomLevel})`; return; }
    if (previewType === "office") {
      const iframe = body.querySelector("iframe");
      try { iframe.contentDocument.body.style.zoom = zoomLevel; } catch (_) {}
    }
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
      previewType = "image";
      zoomBar.style.display = "flex";
      const img = document.createElement("img");
      img.src = href; img.alt = name;
      body.appendChild(img);
    } else if (OFFICE_EXT.test(href)) {
      previewType = "office";
      zoomBar.style.display = "flex";
      const iframe = document.createElement("iframe");
      iframe.src = `/api/preview-file?href=${encodeURIComponent(href)}`;
      iframe.addEventListener("load", () => {
        try { iframe.contentDocument.body.style.zoom = zoomLevel; } catch (_) {}
      });
      body.appendChild(iframe);
    } else {
      previewType = "pdf";
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

  document.querySelector("#docs-view").addEventListener("click", (e) => {
    const btn = e.target.closest(".evidence-file-row:not(.evidence-file-row--folder) [data-preview-href]");
    if (!btn) return;
    const row = btn.closest(".evidence-file-row");
    openPreview(row, btn.dataset.previewHref, btn.textContent.trim());
  });

  window.addEventListener("keydown", (e) => {
    if (e.key === " ") {
      if (panel.classList.contains("active")) { e.preventDefault(); closePreview(); return; }
      if (hoveredRow) {
        e.preventDefault();
        const link = hoveredRow.querySelector("a[href]");
        const previewBtn = hoveredRow.querySelector("[data-preview-href]");
        if (link && canPreview(link.href)) { openPreview(hoveredRow, link.href, link.textContent.trim()); return; }
        if (previewBtn) openPreview(hoveredRow, previewBtn.dataset.previewHref, previewBtn.textContent.trim());
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
        const previewBtn = next.querySelector("[data-preview-href]");
        next.scrollIntoView({ block: "nearest" });
        if (link && canPreview(link.href)) { openPreview(next, link.href, link.textContent.trim()); }
        else if (previewBtn) { openPreview(next, previewBtn.dataset.previewHref, previewBtn.textContent.trim()); }
      }
    }
  }, true);
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
buttons.freigabe.addEventListener("click", () => setView("freigabe"));
buttons.packaging.addEventListener("click", async () => {
  setView("packaging");
  await loadPackagingData();
  renderPackaging();
});
themeToggle.addEventListener("click", () => {
  applyTheme(document.body.dataset.theme === "dark" ? "light" : "dark");
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
  scanBtn.disabled = true;
  scanStatus.textContent = "Scanne…";
  await apiFetch("/api/scan/start", { method: "POST" });
  while (true) {
    await new Promise((r) => setTimeout(r, 500));
    const s = await apiFetch("/api/scan/status");
    if (s.in_progress) {
      if (s.progress?.total > 0) {
        scanStatus.textContent = `${s.progress.done}/${s.progress.total}  ${s.progress.current}`;
      }
    } else {
      showScanDone(s);
      scanBtn.disabled = false;
      projectList = await apiFetch("/api/projects");
      if (activeProject) activeProject = await apiFetch(`/api/projects/${activeProject.id}`);
      if (activeView === "dashboard") { renderMachines(); renderDashboard(); } else renderAll();
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
  startScanStream();
})();
