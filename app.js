// ── DOM refs ────────────────────────────────────────────────
const machineList = document.querySelector("#machine-list");
const search = document.querySelector("#search");
const views = {
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

// ── State ────────────────────────────────────────────────────
let projectList = [];
let activeProject = null;
let activeSubtopic = "Approbation";
let activeBuild = "Alle";
let activeView = "overview";
let activeEvidenceGroup = null;
const evidenceEntries = new Map();

// ── Labels / helpers ─────────────────────────────────────────
const statusFlow = ["Open", "Done", "Not needed", "Blocked"];
const taskFlow = ["Open", "In progress", "Done", "Blocked"];
const freigabeFlow = ["Offen", "Ja", "Nein", "Teilweise"];

const statusLabels = {
  Available: "Vorhanden",
  Archived: "Archiviert",
  Blocked: "Blockiert",
  Current: "Aktuell",
  Done: "Erledigt",
  Draft: "Entwurf",
  Good: "Gut",
  "In progress": "In Arbeit",
  "In review": "In Pruefung",
  "Needs check": "Pruefen",
  "Not needed": "Nicht noetig",
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
  Teilweise: "Teilweise"
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

function statusClass(value) {
  return String(value).toLowerCase().replaceAll(" ", "-").replaceAll("/", "-");
}
function statusLabel(value) { return statusLabels[value] || value; }
function termLabel(value) { return termLabels[value] || value; }
function evidenceCacheKey(projectId, groupName) {
  return `${projectId}:${groupName}`;
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
  Object.entries(views).forEach(([k, el]) => el.classList.toggle("hidden", k !== name));
  Object.entries(buttons).forEach(([k, btn]) => btn.classList.toggle("active", k === name));
}

// ── Machine list ─────────────────────────────────────────────
function renderMachines() {
  const q = search.value.trim().toLowerCase();
  machineList.innerHTML = projectList
    .filter((p) =>
      [p.id, p.name, p.owner, p.phase, p.market, p.variant_group, p.variant_of]
        .join(" ").toLowerCase().includes(q)
    )
    .map((p) => `
      <button class="${p.id === activeProject?.id ? "active" : ""}" data-id="${p.id}" type="button">
        <strong>${p.id}</strong>
        <em class="${statusClass(p.health)}">${statusLabel(p.health)}</em>
        <span>${p.market ? `${termLabel(p.market)} / ` : ""}${p.phase}</span>
        ${p.buildStages?.length ? `
          <div class="sidebar-builds">
            ${p.buildStages.map((b) => `<span data-build-select="${b}" class="${b === activeBuild && p.id === activeProject?.id ? "active" : ""}">${b}</span>`).join("")}
          </div>` : ""}
      </button>
    `).join("");
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
    `PCS Steuerung / ${project.family} / ${termLabel(project.market || "Global")}${relation}`;
  document.querySelector("#project-title").textContent = project.name;

  document.querySelector("#summary-grid").innerHTML = `
    <article class="summary-card hero">
      <span>PCS Reifegrad</span>
      <strong>${project.progress}%</strong>
      <div><i style="width:${project.progress}%"></i></div>
      <p>${project.phase}</p>
    </article>
    ${(project.stats || []).map(([title, value, note]) => {
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
  const openCertifications = (project.certification || [])
    .filter((c) => !c.done && !["Done", "Not needed", "Abgeschlossen"].includes(c.state));
  const certDone = (project.certification || []).filter((c) => c.done).length;
  const openBuilds = (project.builds || []).filter((b) => !["Done", "Not needed"].includes(b.state)).length;
  document.querySelector("#cert-progress").textContent = `${certDone}/${(project.certification || []).length}`;
  document.querySelector("#build-progress").textContent = `${openBuilds} offen`;
  document.querySelector("#risk-count").textContent = `${(project.risks || []).length} Risiken`;
  document.querySelector("#next-count").textContent =
    `${(project.tasks || []).filter((t) => t.status !== "Done").length} offen`;

  document.querySelector("#pcs-strip").innerHTML = `
    <article class="closeout-gates">
      <span>Abschluss-Gates</span>
      ${(project.closeout?.gates || []).map((gate) => `
        <div>
          <b>${gate.label}</b>
          <em class="${statusClass(gate.status)}">${statusLabel(gate.status)}</em>
        </div>
      `).join("")}
    </article>
  `;

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
function renderTasks(project) {
  const area = taskFilter.value;
  const tasks = area === "all"
    ? (project.tasks || [])
    : (project.tasks || []).filter((t) => t.area === area);

  document.querySelector("#task-table").innerHTML = `
    <div class="row head">
      <span>Bereich</span><span>PCS Massnahme</span>
      <span>Verantwortlich</span><span>Faellig</span><span>Status</span>
    </div>
    ${tasks.map((task) => `
      <div class="row ${task.status === "Done" ? "row-done" : ""}">
        <span>${termLabel(task.area)}</span>
        <strong>${task.task}</strong>
        <span>${task.owner}</span>
        <span>${task.due}</span>
        <button class="status-toggle ${statusClass(task.status)}" type="button"
          data-task-id="${task.id}" aria-label="Status aendern">
          ${statusLabel(task.status)}
        </button>
      </div>
    `).join("")}
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
  document.querySelector("#subtopic-title").textContent = `PCS ${activeSubtopic} / Ziffer Checkliste${buildLabel}`;

  if (!subtopic) {
    document.querySelector("#subtopic-summary").textContent = "Fuer diesen Bereich ist noch keine Checkliste angelegt.";
    table.innerHTML = "";
    return;
  }

  document.querySelector("#subtopic-summary").textContent = subtopic.summary;

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
      <div class="ziffer-row">
        <span class="ziffer-number">${item.nr}</span>
        <span class="ziffer-topic">
          <strong>${item.title}</strong>
          ${item.note ? `<small>${item.note}</small>` : ""}
          ${item.evidenceLinks?.length ? `
            <span class="ziffer-links">
              ${item.evidenceLinks.slice(0, 3).map((link) => {
                const href = officeEditHref(link) || evidenceHref(link);
                return `<a href="${href}" target="_blank" rel="noreferrer">${link.label}</a>`;
              }).join("")}
            </span>
          ` : ""}
        </span>
        <button class="status-toggle ${statusClass(item.status)}" type="button"
          data-nr="${item.nr}" aria-label="Status aendern">
          ${statusLabel(item.status)}
        </button>
      </div>
    `).join("")}
  `;
}

// ── Docs / Evidence ───────────────────────────────────────────
function renderDocs(project) {
  const groups = project.documentGroups || [];
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
      <span>Geaendert</span><span>Status</span><span>Datei</span>
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
          ${isWordFile(r.file) ? `<button class="word-action" type="button" data-open-word-href="${evidenceHref(r)}">In Word oeffnen</button>` : ""}
        </span>
      </div>
    `).join("")}
  ` : `<p class="empty-state">Fuer ${activeBuild} sind keine PCS Approbationsberichte indexiert.</p>`;

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
            <span>Name</span><span>Typ</span><span>Geaendert</span><span>Aktion</span>
          </div>
          ${entries.map((entry) => `
            <div class="evidence-file-row">
              <a href="${entry.href}" target="_blank" rel="noreferrer">${entry.name}</a>
              <span>${entry.size || entry.type}</span>
              <span>${entry.modified}</span>
              <span class="evidence-file-actions">
                ${entry.type === "Datei" && isWordFile(entry.name) ? `<button class="word-action" type="button" data-open-word-href="${entry.href}">In Word oeffnen</button>` : ""}
                <button class="finder-action" type="button" data-open-href="${entry.href}">${entry.type === "Ordner" ? "Finder" : "Oeffnen"}</button>
              </span>
            </div>
          `).join("")}
        ` : `<p class="empty-state">Dieser Ordner enthaelt keine sichtbaren Dateien.</p>`;

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
              <button class="finder-action primary" type="button" data-evidence-group="${group.primary}">${isSelected ? "Details ausblenden" : "Details anzeigen"}</button>
              <a class="dgc-link" href="${evidenceHref(group)}" target="_blank" rel="noreferrer">Ordner anzeigen →</a>
              <button class="finder-action" type="button" data-open-href="${evidenceHref(group)}">Im Finder öffnen</button>
            </div>
          </article>
          ${isSelected ? groupDetailMarkup(group, isReportsGroup) : ""}`;
      }).join("")
    : `<p class="empty-state">Fuer dieses Projekt ist noch kein PCS Nachweisindex angelegt.</p>`;

  document.querySelector("#reports-panel")?.classList.add("hidden");
}

// ── Fachfreigabe ──────────────────────────────────────────────
function freigabeGesamtstatus(gates) {
  if (!gates.length) return "Not started";
  const statuses = gates.map((g) => g.status || "Offen");
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
          <p>Manuelle Bestaetigung durch PCS-Chef / Verantwortliche nach Ruecksprache mit VDE, Approbation oder Projektleitung. Diese Punkte koennen nicht automatisch erkannt werden.</p>
        </div>
        <em class="${statusClass(gesamtstatus)} freigabe-badge">${statusLabel(gesamtstatus)}</em>
      </div>
      <h3 class="freigabe-section-title">Bestaetigung der Abschluss-Kriterien</h3>
      <div class="freigabe-gates">
        ${gates.length ? gates.map((gate) => `
          <div class="freigabe-gate-row">
            <span class="freigabe-gate-label">${gate.label}</span>
            <button class="freigabe-btn ${statusClass(gate.status || "Offen")}" type="button"
              data-label="${gate.label}">${statusLabel(gate.status || "Offen")}</button>
          </div>
        `).join("") : `<p class="empty-state">Keine Freigabe-Kriterien fuer dieses Projekt definiert.</p>`}
      </div>
      <div class="freigabe-divider"></div>
      <h3 class="freigabe-section-title">Ruecksprache &amp; Dokumentation</h3>
      <div class="freigabe-meta">
        <div class="freigabe-field">
          <label for="ff-bestaetigt">Bestaetigt von</label>
          <input id="ff-bestaetigt" type="text" placeholder="Name / Funktion" value="${meta.confirmed_by || ""}">
        </div>
        <div class="freigabe-field">
          <label for="ff-datum">Datum</label>
          <input id="ff-datum" type="date" value="${meta.datum || ""}">
        </div>
        <div class="freigabe-field freigabe-field-wide">
          <label for="ff-notiz">Notiz / Ruecksprache</label>
          <textarea id="ff-notiz" rows="4"
            placeholder="z.B. VDE Rueckfragen per E-Mail beantwortet am 15.05.2026, finaler Bericht erwartet bis 30.05.2026.">${meta.notiz || ""}</textarea>
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

function renderBuildChange() {
  updateSidebarBuildSelection();
  activeEvidenceGroup = null;
  renderSubtopic(activeProject);
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
      setView(activeView);
      return;
    }
    activeBuild = pill.dataset.buildSelect;
    activeSubtopic = "Approbation";
    subtopicFilter.value = "all";
    renderBuildChange();
    setView(activeView);
    return;
  }

  activeProject = await apiFetch(`/api/projects/${btn.dataset.id}`);
  activeSubtopic = "Approbation";
  activeBuild = "Alle";
  activeEvidenceGroup = null;
  subtopicFilter.value = "all";
  renderAll();
});

search.addEventListener("input", renderMachines);
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
  const next = freigabeFlow[(freigabeFlow.indexOf(gate.status || "Offen") + 1) % freigabeFlow.length];
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
  apiPut(`/api/projects/${activeProject.id}/tasks/${taskId}`, { status: next }).catch(console.error);
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

async function loadScanStatus() {
  try {
    const s = await apiFetch("/api/scan/status");
    if (s.scanned_at) {
      const d = new Date(s.scanned_at);
      const fmt = d.toLocaleDateString("de-CH", { day: "2-digit", month: "2-digit" })
        + " " + d.toLocaleTimeString("de-CH", { hour: "2-digit", minute: "2-digit" });
      scanStatus.textContent = `${fmt} · ${s.projects_found} Proj.`;
    } else {
      scanStatus.textContent = "Noch nicht gescannt";
    }
  } catch { /* ignore */ }
}

scanBtn.addEventListener("click", async () => {
  scanBtn.disabled = true;
  scanBtn.textContent = "Scanne...";
  scanStatus.textContent = "";
  try {
    const result = await apiFetch("/api/scan", { method: "POST" });
    const n = result.projects?.length || 0;
    const f = result.totalFiles || 0;
    scanStatus.textContent = `${n} Proj. · ${f} Dateien`;
    // Reload project list and active project after scan
    projectList = await apiFetch("/api/projects");
    if (activeProject) {
      activeProject = await apiFetch(`/api/projects/${activeProject.id}`);
    }
    renderAll();
  } catch (err) {
    scanStatus.textContent = "Fehler";
    console.error("Scan error:", err);
  } finally {
    scanBtn.disabled = false;
    scanBtn.textContent = "P:\\PCS scannen";
  }
});

// ── Init ──────────────────────────────────────────────────────
(async () => {
  projectList = await apiFetch("/api/projects");
  if (projectList.length) {
    const preferred = projectList.find((project) => project.id === "EF1157") || projectList[0];
    activeProject = await apiFetch(`/api/projects/${preferred.id}`);
  }
  renderAll();
  loadScanStatus();
})();
