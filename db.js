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

  CREATE TABLE IF NOT EXISTS shipments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id TEXT REFERENCES projects(id) ON DELETE CASCADE,
    build TEXT,
    direction TEXT DEFAULT 'out',
    destination TEXT,
    sent_date TEXT,
    expected_date TEXT,
    received_date TEXT,
    status TEXT DEFAULT 'Geplant',
    tracking TEXT,
    notes TEXT
  );

  CREATE TABLE IF NOT EXISTS labs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    short_name TEXT,
    name TEXT,
    country TEXT,
    address TEXT,
    contact TEXT,
    email TEXT,
    phone TEXT,
    notes TEXT
  );

  CREATE TABLE IF NOT EXISTS archive_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id TEXT,
    device_name TEXT,
    build TEXT,
    quantity INTEGER DEFAULT 1,
    location TEXT,
    status TEXT DEFAULT 'Vorhanden',
    date_stored TEXT,
    notes TEXT,
    source TEXT DEFAULT 'manual',
    last_scanned TEXT
  );
`);

// ── Seed labs (once) ──────────────────────────────────────────────────────
{
  const hasLabs = db.prepare("SELECT COUNT(*) as n FROM labs").get().n;
  if (!hasLabs) {
    const ins = db.prepare(`INSERT INTO labs (short_name, name, country, address, contact, email, phone, notes)
      VALUES (@short_name, @name, @country, @address, @contact, @email, @phone, @notes)`);
    const seedLabs = db.transaction(() => {
      [
        { short_name: "VDE",      name: "VDE Institut GmbH",           country: "DE", address: "Merianstrasse 28, 63069 Offenbach am Main", contact: "Stefan Debes", email: "", phone: "", notes: "IEC / EN Prüfungen, CB-Scheme" },
        { short_name: "UL-IT",    name: "UL LLC – Labor Italien",      country: "IT", address: "Via Carducci 125/A, 20099 Sesto San Giovanni (MI)", contact: "A. Rownicki", email: "", phone: "", notes: "cULus Household, Prüfmuster hierhin" },
        { short_name: "UL-PL",    name: "UL LLC – Labor Polen",        country: "PL", address: "ul. Gładka 1, 31-421 Kraków", contact: "A. Rownicki", email: "", phone: "", notes: "cULus Household, Prüfmuster hierhin" },
        { short_name: "TÜV-TW",   name: "TÜV SÜD Taiwan Ltd.",         country: "TW", address: "5F., No.31, Sec. 1, Dunhua S. Rd., Taipei", contact: "", email: "", phone: "", notes: "BSMI Taiwan – Erneuerungen und Updates" },
        { short_name: "Eurofins", name: "Eurofins E&E Germany GmbH",   country: "DE", address: "Handwerkstrasse 29, 70565 Stuttgart", contact: "", email: "", phone: "", notes: "EMC / ErP Prüfungen" },
        { short_name: "EFE",      name: "Eugster/Frismag AG – intern", country: "CH", address: "Fabrikstrasse 30, 9200 Gossau SG", contact: "", email: "", phone: "", notes: "Internes Labor (LSVA), Safety-Vorprüfungen" },
      ].forEach((l) => ins.run(l));
    });
    seedLabs();
  }
}

// ── Migrations ────────────────────────────────────────────────────────────
try { db.exec("ALTER TABLE ziffern ADD COLUMN builds TEXT DEFAULT ''"); } catch { /* exists */ }
try { db.exec("ALTER TABLE tasks ADD COLUMN builds TEXT DEFAULT 'Alle'"); } catch { /* exists */ }
try { db.exec("ALTER TABLE tasks ADD COLUMN block_reason TEXT DEFAULT ''"); } catch { /* exists */ }
try { db.exec("ALTER TABLE document_groups ADD COLUMN folder_mtime TEXT"); } catch { /* exists */ }
try { db.exec("ALTER TABLE projects ADD COLUMN archive_location TEXT DEFAULT ''"); } catch { /* exists */ }
try { db.exec("ALTER TABLE projects ADD COLUMN project_no TEXT DEFAULT ''"); } catch { /* exists */ }
db.prepare("UPDATE fachfreigabe_gates SET status = 'Offen' WHERE status = 'Teilweise'").run();

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
  const projects = db.prepare("SELECT * FROM projects ORDER BY CAST(SUBSTR(id, 3) AS INTEGER)").all();
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

function updateArchiveLocation(projectId, location) {
  db.prepare("UPDATE projects SET archive_location = ? WHERE id = ?").run(location || "", projectId);
}

function updateProjectNo(projectId, projectNo) {
  db.prepare("UPDATE projects SET project_no = ? WHERE id = ?").run(projectNo || "", projectId);
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

function getShipments(projectId) {
  return db.prepare("SELECT * FROM shipments WHERE project_id = ? ORDER BY sent_date DESC, id DESC").all(projectId);
}

function getAllOpenShipments() {
  return db.prepare(`
    SELECT s.*, p.name AS project_name
    FROM shipments s
    LEFT JOIN projects p ON p.id = s.project_id
    WHERE s.status NOT IN ('Angekommen', 'Zurückgesendet')
    ORDER BY s.expected_date ASC, s.id DESC
  `).all();
}

function upsertShipment(projectId, data) {
  if (data.id) {
    db.prepare(`UPDATE shipments SET build=@build, direction=@direction, destination=@destination,
      sent_date=@sent_date, expected_date=@expected_date, received_date=@received_date,
      status=@status, tracking=@tracking, notes=@notes WHERE id=@id AND project_id=@project_id`)
      .run({ ...data, project_id: projectId });
    return data.id;
  }
  const r = db.prepare(`INSERT INTO shipments
    (project_id,build,direction,destination,sent_date,expected_date,received_date,status,tracking,notes)
    VALUES (@project_id,@build,@direction,@destination,@sent_date,@expected_date,@received_date,@status,@tracking,@notes)`)
    .run({ ...data, project_id: projectId });
  return r.lastInsertRowid;
}

function deleteShipment(projectId, shipmentId) {
  db.prepare("DELETE FROM shipments WHERE id = ? AND project_id = ?").run(shipmentId, projectId);
}

function getLabs() {
  return db.prepare("SELECT * FROM labs ORDER BY id").all();
}

function upsertLab(data) {
  if (data.id) {
    db.prepare(`UPDATE labs SET short_name=@short_name, name=@name, country=@country,
      address=@address, contact=@contact, email=@email, phone=@phone, notes=@notes WHERE id=@id`)
      .run(data);
    return data.id;
  }
  return db.prepare(`INSERT INTO labs (short_name,name,country,address,contact,email,phone,notes)
    VALUES (@short_name,@name,@country,@address,@contact,@email,@phone,@notes)`)
    .run(data).lastInsertRowid;
}

function getArchiveItems(projectId) {
  if (projectId) {
    return db.prepare("SELECT * FROM archive_items WHERE project_id = ? ORDER BY location, id").all(projectId);
  }
  return db.prepare("SELECT * FROM archive_items ORDER BY project_id, location, id").all();
}

function upsertArchiveItem(data) {
  if (data.id) {
    db.prepare(`UPDATE archive_items SET project_id=@project_id, device_name=@device_name,
      build=@build, quantity=@quantity, location=@location, status=@status,
      date_stored=@date_stored, notes=@notes WHERE id=@id`).run(data);
    return data.id;
  }
  return db.prepare(`INSERT INTO archive_items
    (project_id,device_name,build,quantity,location,status,date_stored,notes,source)
    VALUES (@project_id,@device_name,@build,@quantity,@location,@status,@date_stored,@notes,@source)`)
    .run({ source: "manual", ...data }).lastInsertRowid;
}

function deleteArchiveItem(id) {
  db.prepare("DELETE FROM archive_items WHERE id = ?").run(id);
}

module.exports = {
  getProjects,
  getProject,
  updateArchiveLocation,
  updateProjectNo,
  updateZifferStatus,
  updateFachfreigabeGate,
  updateFachfreigabeMeta,
  updateTaskStatus,
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
};
