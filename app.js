// ── DOM refs ────────────────────────────────────────────────
const machineList = document.querySelector("#machine-list");
const search = document.querySelector("#search");
const views = {
  dashboard: document.querySelector("#dashboard-view"),
  overview: document.querySelector("#overview-view"),
  subtopic: document.querySelector("#subtopic-view"),
  tasks: document.querySelector("#tasks-view"),
  docs: document.querySelector("#docs-view"),
  freigabe: document.querySelector("#freigabe-view")
};
const buttons = {
  overview: document.querySelector("#view-overview"),
  subtopic: document.querySelector("#view-subtopic"),
  tasks: document.querySelector("#view-tasks"),
  docs: document.querySelector("#view-docs"),
  freigabe: document.querySelector("#view-freigabe")
};
const taskFilter = document.querySelector("#task-filter");
const subtopicFilter = document.querySelector("#subtopic-filter");
const zifferTable = document.querySelector("#ziffer-table");
const themeToggle = document.querySelector("#theme-toggle");
const dashboardLink = document.querySelector("#dashboard-link");

// ── State ────────────────────────────────────────────────────
let projectList = [];
let activeProject = null;
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

async function loadEvidenceEntries(group) {
  const key = evidenceCacheKey(activeProject.id, group.primary);
  if (evidenceEntries.has(key)) return;
  evidenceEntries.set(key, { loading: true, entries: [] });
  renderDocs(activeProject);
  try {
    const data = await apiFetch(`/api/list-path?href=${encodeURIComponent(evidenceHref(group))}`);
    evidenceEntries.set(key, { loading: false, entries: data.entries || [] });
  } catch (err) {
    evidenceEntries.set(key, { loading: false, error: true, entries: [] });
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
  document.querySelector("#product-image").style.display = name === "dashboard" ? "none" : document.querySelector("#product-image").style.display;
}

// ── Machine list ─────────────────────────────────────────────
function relatedProjects() {
  if (!activeProject || activeView === "dashboard") return projectList;
  return projectList.filter((p) => {
    if (p.family && p.family === activeProject.family) return true;
    if (p.variant_group && p.variant_group === activeProject.variant_group) return true;
    return p.id === activeProject.id;
  });
}

function renderMachines() {
  const q = search.value.trim().toLowerCase();
  const source = activeView === "dashboard" ? projectList : relatedProjects();
  const filtered = source.filter((p) => {
    if (activeView === "dashboard" && !q) return true;
    return [p.id, p.name, p.owner, p.phase, p.market, p.variant_group, p.variant_of]
      .join(" ").toLowerCase().includes(q);
  });

  const renderBtn = (p) => `
    <button class="${p.id === activeProject?.id ? "active" : ""}" data-id="${p.id}" type="button">
      <strong>${p.id}</strong>
      <em class="${statusClass(p.health)}">${statusLabel(p.health)}</em>
      <span>${p.market ? `${termLabel(p.market)} / ` : ""}${p.phase}</span>
      ${p.buildStages?.length ? `
        <div class="sidebar-builds">
          ${p.buildStages.map((b) => `<span data-build-select="${b}" class="${b === activeBuild && p.id === activeProject?.id ? "active" : ""}">${b}</span>`).join("")}
        </div>` : ""}
    </button>`;

  machineList.innerHTML = filtered.map((p) => renderBtn(p)).join("");
}

const dashboardSearch = document.querySelector("#dashboard-search");

function renderDashboard() {
  const q = (dashboardSearch?.value || search.value).trim().toLowerCase();
  const projects = projectList.filter((p) =>
    [p.id, p.name, p.owner, p.family, p.market, p.phase, p.variant_group, p.variant_of]
      .join(" ").toLowerCase().includes(q)
  );
  document.querySelector("#project-family").textContent = "PCS Projektübersicht";
  document.querySelector("#project-title").textContent = "Kaffee Dashboard";
  document.querySelector("#dashboard-count").textContent = `${projects.length} Projekte`;
  document.querySelector("#dashboard-grid").innerHTML = projects.length ? projects.map((p) => `
    <button class="dashboard-card" type="button" data-dashboard-project="${p.id}">
      <span>${p.family || "PCS Maschine"}</span>
      <strong>${p.id}</strong>
      <p>${p.name}</p>
      <em class="${statusClass(p.health)}">${statusLabel(p.health)}</em>
      <small>${p.market ? `${termLabel(p.market)} / ` : ""}${p.phase || ""}</small>
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

  document.querySelector("#summary-grid").innerHTML = `
    <article class="summary-card hero">
      <span>PCS Reifegrad</span>
      <strong>${project.progress}%</strong>
      <div><i style="width:${project.progress}%"></i></div>
      <p>${project.phase}</p>
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
function officeEditHref(link) {
  const href = evidenceHref(link);
  if (!href || !isWordFile(href)) return "";
  const abs = href.startsWith("http") ? href : new URL(href, window.location.href).href;
  return `ms-word:ofe|u|${abs}`;
}
function shouldOpenWordLocally(link) {
  return trackerConfig.documentMode === "local" && isWordFile(evidenceHref(link));
}
function renderEvidenceCell(item) {
  if (!item.evidenceLinks?.length) return item.evidence || "";
  return `
    <div class="evidence-stack">
      <span>${item.evidence || ""}</span>
      ${item.evidenceLinks.map((link) => {
        if (shouldOpenWordLocally(link)) {
          return `<button class="inline-open-action" type="button" data-open-word-href="${evidenceHref(link)}">${link.label}</button>`;
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
  const reportLabel = activeBuild !== "Alle" ? ` — ${activeBuild}` : "";
  const reportMarkup = reports.length ? `
    <div class="report-row head">
      <span>Projekt</span><span>Build</span><span>Version</span>
      <span>Geändert</span><span>Status</span><span>Datei</span>
    </div>
    ${reports.map((r) => `
      <div class="report-row">
        <strong>${r.project || r.project_id || ""}</strong>
        <span>${r.build}</span>
        <span>${r.version}</span>
        <span>${r.modified}<br>${r.size}</span>
        <em class="${statusClass(r.state)}">${statusLabel(r.state)}</em>
        <span class="report-file">
          <a href="${evidenceHref(r)}" target="_blank" rel="noreferrer">${r.file}</a>
          ${isWordFile(r.file) ? `<button class="word-action" type="button" data-open-word-href="${evidenceHref(r)}">In Word öffnen</button>` : ""}
        </span>
      </div>
    `).join("")}
  ` : `<p class="empty-state">Für ${activeBuild} sind keine PCS Approbationsberichte indexiert.</p>`;

  const groupDetailMarkup = (group, isReportsGroup) => {
    const key = evidenceCacheKey(project.id, group.primary);
    const cached = evidenceEntries.get(key);
    const entries = cached?.entries || [];
    const fileListMarkup = cached?.loading
      ? `<p class="empty-state">Ordnerinhalt wird geladen...</p>`
      : cached?.error
        ? `<p class="empty-state">Ordnerinhalt konnte nicht gelesen werden.</p>`
        : entries.length ? `
          <div class="evidence-file-row head">
            <span>Name</span><span>Typ</span><span>Geändert</span><span>Aktion</span>
          </div>
          ${entries.map((entry) => `
            <div class="evidence-file-row">
              <a href="${entry.href}" target="_blank" rel="noreferrer">${entry.name}</a>
              <span>${entry.size || entry.type}</span>
              <span>${entry.modified}</span>
              <span class="evidence-file-actions">
                ${entry.type === "Datei" && isWordFile(entry.name) ? `<button class="word-action" type="button" data-open-word-href="${entry.href}">In Word öffnen</button>` : ""}
                <button class="finder-action" type="button" data-open-href="${entry.href}">${entry.type === "Ordner" ? "Finder" : "Öffnen"}</button>
              </span>
            </div>
          `).join("")}
        ` : `<p class="empty-state">Dieser Ordner enthält keine sichtbaren Dateien.</p>`;

    return `
      <section class="document-group-detail">
        <div class="table-header compact">
          <div>
            <h2>${group.primary}</h2>
            <p>${isReportsGroup ? "Approbationsberichte und kompletter Ordnerinhalt." : "Direkter Zugriff auf Dateien und Unterordner dieser Nachweisgruppe."}</p>
          </div>
        </div>
        ${isReportsGroup ? `
          <div class="report-subsection">
            <h3>PCS Approbationsberichte${reportLabel}</h3>
            <div class="report-table">${reportMarkup}</div>
          </div>
        ` : ""}
        <div class="evidence-file-table">${fileListMarkup}</div>
      </section>
    `;
  };

  document.querySelector("#document-group-grid").innerHTML = groups.length
    ? groups.map((group) => {
        const num = parseInt(group.count) || 0;
        const folderLabel = group.primary || termLabel(group.area);
        const isReportsGroup = String(group.primary || "").startsWith("12 ");
        const isSelected = activeEvidenceGroup === group.primary;
        return `
          <article class="document-group-card ${statusClass(group.status)} ${isSelected ? "selected" : ""}"
            data-evidence-group-card="${group.primary}" role="button" tabindex="0">
            <div class="dgc-top">
              <strong>${folderLabel}</strong>
              <em class="status-toggle ${statusClass(group.status)}">${statusLabel(group.status)}</em>
            </div>
            <span class="dgc-area">${termLabel(group.area)}</span>
            <div class="dgc-count">${num}<span>Dateien</span></div>
            <div class="dgc-actions">
              <button class="finder-action" type="button" data-evidence-group="${group.primary}">${isSelected ? "Details ausblenden" : "Details anzeigen"}</button>
              <a class="dgc-link" href="${evidenceHref(group)}" target="_blank" rel="noreferrer">Ordner anzeigen →</a>
              <button class="finder-action" type="button" data-open-href="${evidenceHref(group)}">Im Finder öffnen</button>
            </div>
          </article>
          ${isSelected ? groupDetailMarkup(group, isReportsGroup) : ""}`;
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

async function openProject(projectId, view = "overview") {
  activeProject = await apiFetch(`/api/projects/${projectId}`);
  activeSubtopic = "Approbation";
  activeBuild = "Alle";
  activeEvidenceGroup = null;
  subtopicFilter.value = "all";
  setView(view);
  renderAll();
}

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
  setView("dashboard");
  renderMachines();
  renderDashboard();
});

document.querySelector("#dashboard-grid").addEventListener("click", (event) => {
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
  if (!btn) return;
  const ziffer = activeProject.subtopics?.[activeSubtopic]?.ziffern.find((z) => z.nr === btn.dataset.nr);
  if (!ziffer) return;
  const next = statusFlow[(statusFlow.indexOf(ziffer.status) + 1) % statusFlow.length];
  ziffer.status = next; // optimistic update
  renderSubtopic(activeProject);
  apiPut(`/api/projects/${activeProject.id}/ziffern/${activeSubtopic}/${btn.dataset.nr}`, { status: next })
    .catch(console.error);
});

document.querySelector("#summary-grid").addEventListener("click", (event) => {
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
  const groupBtn = event.target.closest("[data-evidence-group]");
  if (groupBtn) {
    const group = activeProject.documentGroups?.find((item) => item.primary === groupBtn.dataset.evidenceGroup);
    toggleEvidenceGroup(groupBtn.dataset.evidenceGroup);
    renderDocs(activeProject);
    if (activeEvidenceGroup && group) loadEvidenceEntries(group);
    return;
  }

  const groupCard = event.target.closest("[data-evidence-group-card]");
  if (groupCard && !event.target.closest("a, button")) {
    const group = activeProject.documentGroups?.find((item) => item.primary === groupCard.dataset.evidenceGroupCard);
    toggleEvidenceGroup(groupCard.dataset.evidenceGroupCard);
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
  const btn = event.target.closest("[data-open-word-href]");
  if (!btn) return;
  event.preventDefault();
  btn.disabled = true;
  try {
    await apiFetch("/api/open-path", {
      method: "POST",
      body: JSON.stringify({
        href: new URL(btn.dataset.openWordHref, window.location.href).pathname,
        app: "word"
      })
    });
  } catch (err) {
    console.error("Open Word error:", err);
  } finally {
    btn.disabled = false;
  }
});

// Tab buttons
buttons.overview.addEventListener("click", () => setView("overview"));
buttons.subtopic.addEventListener("click", () => setView("subtopic"));
buttons.tasks.addEventListener("click", () => setView("tasks"));
buttons.docs.addEventListener("click", () => setView("docs"));
buttons.freigabe.addEventListener("click", () => setView("freigabe"));
themeToggle.addEventListener("click", () => {
  applyTheme(document.body.dataset.theme === "dark" ? "light" : "dark");
});

// ── Scanner UI ────────────────────────────────────────────────
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

// ── Sitzungsexport ────────────────────────────────────────────
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
  startScanStream();
})();
