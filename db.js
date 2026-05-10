const Database = require("better-sqlite3");
const path = require("path");

const db = new Database(path.join(__dirname, "pcs.db"));
db.pragma("journal_mode = WAL");
db.pragma("foreign_keys = ON");

db.exec(`
  CREATE TABLE IF NOT EXISTS projects (
    id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    family TEXT,
    market TEXT,
    variant_of TEXT,
    variant_group TEXT,
    owner TEXT,
    phase TEXT,
    target TEXT,
    health TEXT DEFAULT 'Watch',
    progress INTEGER DEFAULT 0,
    closeout_status TEXT,
    closeout_summary TEXT,
    updated_at TEXT
  );

  CREATE TABLE IF NOT EXISTS project_stats (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id TEXT REFERENCES projects(id) ON DELETE CASCADE,
    label TEXT, value TEXT, note TEXT, sort_order INTEGER DEFAULT 0
  );

  CREATE TABLE IF NOT EXISTS builds (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id TEXT REFERENCES projects(id) ON DELETE CASCADE,
    label TEXT, date TEXT, state TEXT DEFAULT 'Open',
    samples TEXT, note TEXT
  );

  CREATE TABLE IF NOT EXISTS tasks (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id TEXT REFERENCES projects(id) ON DELETE CASCADE,
    area TEXT, task TEXT, owner TEXT, due TEXT, status TEXT DEFAULT 'Open',
    builds TEXT DEFAULT 'Alle',
    block_reason TEXT DEFAULT ''
  );

  CREATE TABLE IF NOT EXISTS risks (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id TEXT REFERENCES projects(id) ON DELETE CASCADE,
    level TEXT, text TEXT
  );

  CREATE TABLE IF NOT EXISTS certification (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id TEXT REFERENCES projects(id) ON DELETE CASCADE,
    name TEXT, state TEXT, done INTEGER DEFAULT 0
  );

  CREATE TABLE IF NOT EXISTS subtopic_summaries (
    project_id TEXT,
    subtopic TEXT,
    summary TEXT,
    PRIMARY KEY (project_id, subtopic),
    FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE
  );

  CREATE TABLE IF NOT EXISTS ziffern (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id TEXT REFERENCES projects(id) ON DELETE CASCADE,
    subtopic TEXT DEFAULT 'Approbation',
    nr TEXT, title TEXT,
    status TEXT DEFAULT 'Open',
    owner TEXT, evidence TEXT, note TEXT
  );

  CREATE TABLE IF NOT EXISTS evidence_links (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ziffer_id INTEGER REFERENCES ziffern(id) ON DELETE CASCADE,
    label TEXT, href TEXT
  );

  CREATE TABLE IF NOT EXISTS closeout_gates (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id TEXT REFERENCES projects(id) ON DELETE CASCADE,
    label TEXT, status TEXT DEFAULT 'Open', sort_order INTEGER DEFAULT 0
  );

  CREATE TABLE IF NOT EXISTS fachfreigabe_gates (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id TEXT REFERENCES projects(id) ON DELETE CASCADE,
    label TEXT, status TEXT DEFAULT 'Offen', sort_order INTEGER DEFAULT 0
  );

  CREATE TABLE IF NOT EXISTS fachfreigabe_meta (
    project_id TEXT PRIMARY KEY REFERENCES projects(id) ON DELETE CASCADE,
    confirmed_by TEXT DEFAULT '',
    datum TEXT DEFAULT '',
    notiz TEXT DEFAULT '',
    updated_at TEXT
  );

  CREATE TABLE IF NOT EXISTS document_groups (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id TEXT REFERENCES projects(id) ON DELETE CASCADE,
    area TEXT, status TEXT, count TEXT,
    summary TEXT, primary_doc TEXT, href TEXT,
    last_scanned TEXT
  );

  CREATE TABLE IF NOT EXISTS report_versions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id TEXT REFERENCES projects(id) ON DELETE CASCADE,
    build TEXT, version TEXT, modified TEXT, size TEXT,
    state TEXT DEFAULT 'Current', file TEXT, href TEXT,
    last_scanned TEXT
  );

  CREATE TABLE IF NOT EXISTS scan_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    scanned_at TEXT,
    projects_found INTEGER,
    files_found INTEGER,
    notes TEXT
  );
`);

// ── Migrations ────────────────────────────────────────────────────────────
try { db.exec("ALTER TABLE ziffern ADD COLUMN builds TEXT DEFAULT ''"); } catch { /* exists */ }
try { db.exec("ALTER TABLE tasks ADD COLUMN builds TEXT DEFAULT 'Alle'"); } catch { /* exists */ }
try { db.exec("ALTER TABLE tasks ADD COLUMN block_reason TEXT DEFAULT ''"); } catch { /* exists */ }
try { db.exec("ALTER TABLE document_groups ADD COLUMN folder_mtime TEXT"); } catch { /* exists */ }

{
  const setBuilds = db.prepare(
    "UPDATE ziffern SET builds = ? WHERE project_id = ? AND nr = ? AND (builds IS NULL OR builds = '')"
  );
  const ef1157 = {
    "4": "Alle", "7": "PT1,TS1", "8": "PT1,TS1", "10": "PT1,TS1,TS2",
    "11": "PT1,TS1", "13": "PT1,TS1", "15": "TS1", "16": "TS1",
    "17": "PT1,TS1", "19": "PT1,OOT,TS1", "20": "PT1", "22": "PT1,TS1",
    "24": "Alle", "25": "PT1,TS1", "27": "PT1,TS1", "29": "PT1,TS1",
    "30": "PT1,TS1", "32": "PT1,TS1"
  };
  for (const [nr, builds] of Object.entries(ef1157)) setBuilds.run(builds, "EF1157", nr);
  // EF1234 and EF1107 default to PT1,TS1
  db.prepare("UPDATE ziffern SET builds = 'PT1,TS1' WHERE project_id != 'EF1157' AND (builds IS NULL OR builds = '')").run();
}

{
  const setTaskBuilds = db.prepare(
    "UPDATE tasks SET builds = ? WHERE project_id = ? AND task = ? AND (builds IS NULL OR builds = '' OR builds = 'Alle')"
  );
  const ef1157 = [
    ["Confirm latest VDE order scope", "Alle"],
    ["Review EF1157 mainboard Rev B against wiring diagram", "PT1"],
    ["Link TS1 / TS2 dispatch lists to actual samples", "TS1,TS2"],
    ["Create sample evidence matrix", "PT1,OOT,TS1,TS2"],
    ["Check safety, EMC, ErP report completeness", "TS1"],
    ["Prepare final conformity document list", "TS2"],
    ["Clarify missing final CB / EMC / ErP report package", "TS2"],
    ["Create clean TS1 / TS2 sample and dispatch overview", "TS1,TS2"],
    ["Confirm country implementation folder scope", "TS1,TS2"]
  ];
  ef1157.forEach(([task, builds]) => setTaskBuilds.run(builds, "EF1157", task));
  const ef1234 = [
    ["Confirm Brazil approval route and required standards", "Alle"],
    ["Define PT1 Brazil sample configuration", "PT1"],
    ["Create EF1157 to EF1234 delta matrix", "PT1,TS1"],
    ["Check mains, cord, plug, and PCB carry-over", "PT1,TS1"],
    ["Link Brazil samples to evidence package", "PT1,TS1"]
  ];
  ef1234.forEach(([task, builds]) => setTaskBuilds.run(builds, "EF1234", task));
  setTaskBuilds.run("PT1,TS1", "EF1107", "Compare shared approval documents with EF1157");
}

// ── Queries ────────────────────────────────────────────────────────────────

const BUILD_ORDER = ["PT1", "OOT", "TS1", "TS2", "TS3"];

function getProjects() {
  const projects = db.prepare("SELECT * FROM projects ORDER BY id").all();
  const getZifferBuilds = db.prepare(
    "SELECT builds FROM ziffern WHERE project_id = ? AND builds != '' AND builds != 'Alle'"
  );
  return projects.map((p) => {
    const buildSet = new Set();
    getZifferBuilds.all(p.id).forEach((r) =>
      r.builds.split(",").map((b) => b.trim()).filter((b) => b && b !== "Alle").forEach((b) => buildSet.add(b))
    );
    p.buildStages = [...BUILD_ORDER.filter((b) => buildSet.has(b)), ...[...buildSet].filter((b) => !BUILD_ORDER.includes(b))];
    return p;
  });
}

function getProject(id) {
  const project = db.prepare("SELECT * FROM projects WHERE id = ?").get(id);
  if (!project) return null;

  project.stats = db.prepare(
    "SELECT label, value, note FROM project_stats WHERE project_id = ? ORDER BY sort_order"
  ).all(id).map((r) => [r.label, r.value, r.note]);

  project.builds = db.prepare("SELECT * FROM builds WHERE project_id = ?").all(id);
  project.tasks = db.prepare("SELECT * FROM tasks WHERE project_id = ? ORDER BY due").all(id);
  project.risks = db.prepare("SELECT * FROM risks WHERE project_id = ?").all(id);
  project.certification = db.prepare("SELECT * FROM certification WHERE project_id = ?").all(id)
    .map((r) => ({ ...r, done: Boolean(r.done) }));

  project.documentGroups = db.prepare("SELECT * FROM document_groups WHERE project_id = ? ORDER BY primary_doc").all(id)
    .map((r) => ({ area: r.area, status: r.status, count: r.count, summary: r.summary, primary: r.primary_doc, href: r.href }));

  project.reportVersions = db.prepare("SELECT * FROM report_versions WHERE project_id = ?").all(id);

  // Ziffern with evidence links grouped by subtopic
  const rawZiffern = db.prepare(
    "SELECT * FROM ziffern WHERE project_id = ? ORDER BY subtopic, CAST(nr AS INTEGER)"
  ).all(id);

  const summaries = db.prepare(
    "SELECT subtopic, summary FROM subtopic_summaries WHERE project_id = ?"
  ).all(id);
  const summaryMap = Object.fromEntries(summaries.map((s) => [s.subtopic, s.summary]));

  project.subtopics = {};
  rawZiffern.forEach((z) => {
    const links = db.prepare("SELECT label, href FROM evidence_links WHERE ziffer_id = ?").all(z.id);
    const ziffer = { ...z, evidenceLinks: links.length ? links : undefined };
    if (!project.subtopics[z.subtopic]) {
      project.subtopics[z.subtopic] = { summary: summaryMap[z.subtopic] || "", ziffern: [] };
    }
    project.subtopics[z.subtopic].ziffern.push(ziffer);
  });

  // Closeout gates
  const closeoutGates = db.prepare(
    "SELECT label, status FROM closeout_gates WHERE project_id = ? ORDER BY sort_order"
  ).all(id);
  project.closeout = {
    status: project.closeout_status || "Open",
    summary: project.closeout_summary || "",
    gates: closeoutGates
  };

  // Fachfreigabe
  const ffGates = db.prepare(
    "SELECT label, status FROM fachfreigabe_gates WHERE project_id = ? ORDER BY sort_order"
  ).all(id);
  const ffMeta = db.prepare("SELECT * FROM fachfreigabe_meta WHERE project_id = ?").get(id) || {};
  project.fachfreigabe = { gates: ffGates, meta: ffMeta };

  return project;
}

function updateZifferStatus(projectId, subtopic, nr, status) {
  db.prepare(
    "UPDATE ziffern SET status = ? WHERE project_id = ? AND subtopic = ? AND nr = ?"
  ).run(status, projectId, subtopic, nr);
}

function updateFachfreigabeGate(projectId, label, status) {
  db.prepare(
    "UPDATE fachfreigabe_gates SET status = ? WHERE project_id = ? AND label = ?"
  ).run(status, projectId, label);
}

function updateFachfreigabeMeta(projectId, confirmed_by, datum, notiz) {
  db.prepare(`
    INSERT INTO fachfreigabe_meta (project_id, confirmed_by, datum, notiz, updated_at)
    VALUES (?, ?, ?, ?, datetime('now'))
    ON CONFLICT(project_id) DO UPDATE SET
      confirmed_by = excluded.confirmed_by,
      datum = excluded.datum,
      notiz = excluded.notiz,
      updated_at = excluded.updated_at
  `).run(projectId, confirmed_by || "", datum || "", notiz || "");
}

function updateTaskStatus(projectId, taskId, status, blockReason) {
  if (typeof blockReason === "string") {
    db.prepare("UPDATE tasks SET status = ?, block_reason = ? WHERE id = ? AND project_id = ?")
      .run(status, blockReason, taskId, projectId);
    return;
  }
  db.prepare("UPDATE tasks SET status = ? WHERE id = ? AND project_id = ?")
    .run(status, taskId, projectId);
}

module.exports = {
  getProjects,
  getProject,
  updateZifferStatus,
  updateFachfreigabeGate,
  updateFachfreigabeMeta,
  updateTaskStatus,
  db
};
