/**
 * Einmalig ausführen: node seed.js
 * Liest data.js und schreibt alle Projekte in die SQLite-Datenbank.
 * Bestehende Daten werden vollständig ersetzt.
 */
const vm = require("vm");
const fs = require("fs");
const { db } = require("./db");

// const is block-scoped in vm — rewrite to var so it lands on the sandbox
const src = fs.readFileSync("./data.js", "utf8")
  .replace(/\bconst\s+projects\b/, "var projects");
const sandbox = {};
vm.createContext(sandbox);
vm.runInContext(src, sandbox);
const { projects } = sandbox;

if (!projects?.length) {
  console.error("Keine Projekte in data.js gefunden.");
  process.exit(1);
}

const insertProject = db.prepare(`
  INSERT OR REPLACE INTO projects
    (id, name, family, market, variant_of, variant_group, owner, phase, target, health, progress, closeout_status, closeout_summary, updated_at)
  VALUES
    (@id, @name, @family, @market, @variantOf, @variantGroup, @owner, @phase, @target, @health, @progress, @closeoutStatus, @closeoutSummary, @updatedAt)
`);

const insertStat = db.prepare(
  "INSERT INTO project_stats (project_id, label, value, note, sort_order) VALUES (?, ?, ?, ?, ?)"
);
const insertBuild = db.prepare(
  "INSERT INTO builds (project_id, label, date, state, samples, note) VALUES (?, ?, ?, ?, ?, ?)"
);
const insertTask = db.prepare(
  "INSERT INTO tasks (project_id, area, task, owner, due, status, builds, block_reason) VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
);
const insertRisk = db.prepare(
  "INSERT INTO risks (project_id, level, text) VALUES (?, ?, ?)"
);
const insertCert = db.prepare(
  "INSERT INTO certification (project_id, name, state, done) VALUES (?, ?, ?, ?)"
);
const insertSubtopicSummary = db.prepare(`
  INSERT OR REPLACE INTO subtopic_summaries (project_id, subtopic, summary) VALUES (?, ?, ?)
`);
const insertZiffer = db.prepare(
  "INSERT INTO ziffern (project_id, subtopic, nr, title, status, owner, evidence, note) VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
);
const insertEvidenceLink = db.prepare(
  "INSERT INTO evidence_links (ziffer_id, label, href) VALUES (?, ?, ?)"
);
const insertCloseoutGate = db.prepare(
  "INSERT INTO closeout_gates (project_id, label, status, sort_order) VALUES (?, ?, ?, ?)"
);
const insertFfGate = db.prepare(
  "INSERT INTO fachfreigabe_gates (project_id, label, status, sort_order) VALUES (?, ?, ?, ?)"
);
const insertDocGroup = db.prepare(
  "INSERT INTO document_groups (project_id, area, status, count, summary, primary_doc, href) VALUES (?, ?, ?, ?, ?, ?, ?)"
);
const insertReport = db.prepare(
  "INSERT INTO report_versions (project_id, build, version, modified, size, state, file, href) VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
);

// Wipe and re-seed in a transaction
const seed = db.transaction(() => {
  // Clear all dependent tables first (CASCADE would handle it but let's be explicit)
  const tables = [
    "evidence_links", "ziffern", "subtopic_summaries",
    "project_stats", "builds", "tasks", "risks", "certification",
    "closeout_gates", "fachfreigabe_gates", "fachfreigabe_meta",
    "document_groups", "report_versions", "projects"
  ];
  tables.forEach((t) => db.prepare(`DELETE FROM ${t}`).run());

  for (const p of projects) {
    insertProject.run({
      id: p.id,
      name: p.name,
      family: p.family || "",
      market: p.market || "",
      variantOf: p.variantOf || null,
      variantGroup: p.variantGroup || "",
      owner: p.owner || "",
      phase: p.phase || "",
      target: p.target || "",
      health: p.health || "Watch",
      progress: p.progress || 0,
      closeoutStatus: p.closeout?.status || "Open",
      closeoutSummary: p.closeout?.summary || "",
      updatedAt: p.updated || new Date().toISOString()
    });

    (p.stats || []).forEach(([label, value, note], i) =>
      insertStat.run(p.id, label, value, note, i)
    );

    (p.builds || []).forEach((b) =>
      insertBuild.run(p.id, b.label, b.date, b.state, b.samples || "", b.note || "")
    );

    (p.tasks || []).forEach((t) =>
      insertTask.run(p.id, t.area, t.task, t.owner, t.due, t.status, t.builds || "Alle", t.block_reason || "")
    );

    (p.risks || []).forEach((r) =>
      insertRisk.run(p.id, r.level, r.text)
    );

    (p.certification || []).forEach((c) =>
      insertCert.run(p.id, c.name, c.state, c.done ? 1 : 0)
    );

    Object.entries(p.subtopics || {}).forEach(([subtopic, data]) => {
      insertSubtopicSummary.run(p.id, subtopic, data.summary || "");
      (data.ziffern || []).forEach((z) => {
        const result = insertZiffer.run(
          p.id, subtopic, z.nr, z.title, z.status, z.owner, z.evidence || "", z.note || ""
        );
        (z.evidenceLinks || []).forEach((link) =>
          insertEvidenceLink.run(result.lastInsertRowid, link.label, link.href)
        );
      });
    });

    (p.closeout?.gates || []).forEach((g, i) =>
      insertCloseoutGate.run(p.id, g.label, g.status, i)
    );

    (p.fachfreigabe?.gates || []).forEach((g, i) =>
      insertFfGate.run(p.id, g.label, g.status || "Offen", i)
    );

    (p.documentGroups || []).forEach((g) =>
      insertDocGroup.run(p.id, g.area, g.status, g.count || "", g.summary || "", g.primary || "", g.href || "")
    );

    (p.reportVersions || []).forEach((r) =>
      insertReport.run(p.id, r.build, r.version, r.modified, r.size, r.state, r.file, r.href || "")
    );
  }
});

seed();

console.log(`Datenbank bereit. ${projects.length} Projekte importiert:`);
projects.forEach((p) => console.log(`  ${p.id}  ${p.name}`));
