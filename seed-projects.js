// One-time seed: insert all known PCS projects from the P:\Projekte\PCS folder list
// and from the Aktivitätensitzungen 2025/2026.
// Uses INSERT OR IGNORE so it's safe to run multiple times.

const { db } = require("./db.js");

const insert = db.prepare(`
  INSERT OR IGNORE INTO projects
    (id, name, family, market, phase, health, progress, owner, target, variant_group, variant_of)
  VALUES
    (@id, @name, @family, @market, @phase, @health, @progress, @owner, @target, @variant_group, @variant_of)
`);

const setGroup = db.prepare(`
  UPDATE projects SET variant_group = @group
  WHERE id = @id AND (variant_group IS NULL OR variant_group = '')
`);
const setPrimary = db.prepare(`
  UPDATE projects SET variant_of = @primary
  WHERE id = @id AND (variant_of IS NULL OR variant_of = '')
`);

// Shorthand builder
function p(id, name, family, market, phase, health, progress = 0, owner = "", target = "") {
  return { id, name, family, market, phase, health, progress, owner, target, variant_group: "", variant_of: "" };
}

const projects = [
  // ── Abgeschlossen – from P:\Projekte\PCS image ─────────────────────────────
  p("EF221",  "Freezio 221",                   "Freezio",   "EU",     "Abgeschlossen", "Good", 100),
  p("EF253",  "Nivona 253",                    "Nivona",    "EU",     "Abgeschlossen", "Good", 100),
  p("EF254",  "Nivona 254",                    "Nivona",    "EU",     "Abgeschlossen", "Good", 100, "PCS9_E"),
  p("EF258",  "Melitta 258",                   "Melitta",   "EU",     "Abgeschlossen", "Good", 100),
  p("EF259",  "Melitta 259",                   "Melitta",   "EU",     "Abgeschlossen", "Good", 100),
  p("EF265",  "Delica DC265",                  "Delica",    "EU",     "In progress",   "Watch", 40, "PCS9_E", "2026-11"),
  p("EF267",  "Delica DC267",                  "Delica",    "EU",     "Abgeschlossen", "Good", 100),
  p("EF269",  "Delica DC269",                  "Delica",    "EU",     "Abgeschlossen", "Good", 100),
  p("EF275",  "NNSA Maschinenpartner 275",      "NNSA",      "EU",     "Abgeschlossen", "Good", 100),
  p("EF276",  "NNSA Maschinenpartner 276",      "NNSA",      "EU",     "Abgeschlossen", "Good", 100),
  p("EF310",  "Jura 310",                      "Jura",      "EU",     "Abgeschlossen", "Good", 100),
  p("EF359",  "Krüger 359",                    "Krüger",    "EU",     "Abgeschlossen", "Good", 100),
  p("EF361",  "Krüger 361",                    "Krüger",    "EU",     "Abgeschlossen", "Good", 100),
  p("EF375",  "Carogusto 375",                 "Carogusto", "EU",     "Abgeschlossen", "Good", 100),
  p("EF376",  "Carogusto 376",                 "Carogusto", "EU",     "Abgeschlossen", "Good", 100, "PCS2_E"),
  p("EF384",  "Melitta 384",                   "Melitta",   "EU",     "Abgeschlossen", "Good", 100),
  p("EF385",  "Melitta 385",                   "Melitta",   "EU",     "Abgeschlossen", "Good", 100),
  p("EF392",  "Miele 392",                     "Miele",     "EU",     "Abgeschlossen", "Good", 100),
  p("EF393",  "Miele 393",                     "Miele",     "EU",     "Abgeschlossen", "Good", 100),
  p("EF395",  "Miele 395",                     "Miele",     "EU",     "Abgeschlossen", "Good", 100),
  p("EF396",  "Miele 396",                     "Miele",     "EU",     "Abgeschlossen", "Good", 100),
  p("EF397",  "Miele 397",                     "Miele",     "EU",     "Abgeschlossen", "Good", 100),
  p("EF398",  "Miele 398",                     "Miele",     "EU",     "Abgeschlossen", "Good", 100),
  p("EF503",  "Miele 503",                     "Miele",     "EU",     "Abgeschlossen", "Good", 100),
  p("EF504",  "Miele 504",                     "Miele",     "EU",     "Abgeschlossen", "Good", 100),
  p("EF505",  "Miele 505",                     "Miele",     "EU",     "Abgeschlossen", "Good", 100),
  p("EF534",  "Jura 534",                      "Jura",      "EU",     "Abgeschlossen", "Good", 100),
  p("EF539",  "Jura 539",                      "Jura",      "EU",     "Abgeschlossen", "Good", 100),
  p("EF545",  "Jura 545",                      "Jura",      "EU",     "Abgeschlossen", "Good", 100),
  p("EF555",  "Jura 555",                      "Jura",      "EU",     "Abgeschlossen", "Good", 100),
  p("EF565",  "Jura 565",                      "Jura",      "EU",     "Abgeschlossen", "Good", 100),
  p("EF566",  "Jura 566",                      "Jura",      "EU",     "Abgeschlossen", "Good", 100),
  p("EF567",  "Jura 567",                      "Jura",      "EU",     "Abgeschlossen", "Good", 100),
  p("EF568",  "Nivona 568",                    "Nivona",    "EU",     "In progress",   "Watch", 55, "PCS8_E", "2026-11"),
  p("EF573",  "Nivona 573",                    "Nivona",    "EU",     "Abgeschlossen", "Watch", 90, "PCS9_E"),
  p("EF577",  "Franke 577",                    "Franke",    "EU",     "Abgeschlossen", "Good", 100),
  p("EF578",  "Franke 578",                    "Franke",    "EU",     "Abgeschlossen", "Good", 100),
  p("EF590",  "Melitta 590",                   "Melitta",   "EU",     "Abgeschlossen", "Good", 100),
  p("EF592",  "Melitta 592",                   "Melitta",   "EU",     "Abgeschlossen", "Good", 100),
  p("EF595",  "Krüger 595",                    "Krüger",    "EU",     "Abgeschlossen", "Good", 100),
  p("EF677",  "Melitta 677",                   "Melitta",   "EU",     "Abgeschlossen", "Good", 100),
  p("EF691",  "Franke 691",                    "Franke",    "EU",     "Abgeschlossen", "Good", 100),
  p("EF693",  "Melitta 693",                   "Melitta",   "EU",     "Abgeschlossen", "Good", 100),
  p("EF695",  "Melitta 695",                   "Melitta",   "EU",     "Abgeschlossen", "Good", 100),
  p("EF696",  "Melitta 696",                   "Melitta",   "EU",     "Abgeschlossen", "Good", 100),
  p("EF698",  "Nivona 698",                    "Nivona",    "EU",     "Abgeschlossen", "Watch", 90, "PCS9_E"),
  p("EF704",  "NNSA 704",                      "NNSA",      "EU",     "Abgeschlossen", "Good", 100),
  p("EF709",  "Nivona 709",                    "Nivona",    "EU",     "Abgeschlossen", "Good", 100),
  p("EF715",  "Jura 715",                      "Jura",      "EU",     "Abgeschlossen", "Good", 100),
  p("EF1013", "Jura 1013",                     "Jura",      "EU",     "Abgeschlossen", "Good", 100),
  p("EF1015", "Jura 1015",                     "Jura",      "EU",     "Abgeschlossen", "Good", 100),

  // ── NNSA – Länderzulassung Funk, active ────────────────────────────────────
  p("EF230",  "NNSA Proxy 220–240V",           "NNSA",      "Global", "In progress",   "Watch", 70, "PCS2_E"),
  p("EF231",  "NNSA Proxy 120–127V",           "NNSA",      "Global", "In progress",   "Watch", 70, "PCS2_E"),
  p("EF232",  "NNSA Proxy 100V",               "NNSA",      "Global", "In progress",   "Watch", 70, "PCS2_E"),

  // ── Nespresso TW / D140 ────────────────────────────────────────────────────
  p("EF1006", "Nespresso TW 1006",             "Nespresso", "Global", "In progress",   "Watch", 60, "PCS9_E"),
  p("EF1007", "Nespresso TW 1007",             "Nespresso", "Global", "In progress",   "Watch", 60, "PCS9_E"),
  p("EF1008", "Nespresso TW 1008",             "Nespresso", "Global", "Abgeschlossen", "Good", 100, "PCS9_E"),
  p("EF1009", "Nespresso TW 1009",             "Nespresso", "Global", "Abgeschlossen", "Good", 100, "PCS9_E"),
  p("EF1016", "Nespresso D140 TW",             "Nespresso", "Global", "In progress",   "Watch", 65, "PCS9_E"),
  p("EF1017", "Nespresso D140 Swarovski",      "Nespresso", "Global", "In progress",   "Watch", 65, "PCS9_E"),

  // ── Delica – VDE-GS Verlängerung ──────────────────────────────────────────
  p("EF1052", "Delica DC1052",                 "Delica",    "EU",     "In progress",   "Watch", 30, "PCS9_E", "2026-07"),

  // ── Jura cULus ────────────────────────────────────────────────────────────
  p("EF1106", "Jura cULus 1106",               "Jura",      "US",     "Abgeschlossen", "Good", 100, "PCS8_E"),
  p("EF1119", "Jura cULus 1119",               "Jura",      "US",     "Abgeschlossen", "Good", 100, "PCS8_E"),
  p("EF1120", "Jura cULus 1120",               "Jura",      "US",     "Abgeschlossen", "Good", 100, "PCS8_E"),
  p("EF1121", "Jura cULus 1121",               "Jura",      "US",     "In progress",   "Watch", 50, "PCS8_E"),
  p("EF1123", "Jura cULus Household 1123",     "Jura",      "US",     "In progress",   "Watch", 50, "PCS2_E"),
  p("EF1171", "Jura cULus Household 1171",     "Jura",      "US",     "In progress",   "Watch", 25, "PCS8_E"),

  // ── Delica 1156 (BU) – related to EF1157 ─────────────────────────────────
  p("EF1156", "Delica Brüheinheit 1156",       "Delica",    "EU",     "In progress",   "Watch", 40, "PCS2_E"),

  // ── NNSA Horeca ───────────────────────────────────────────────────────────
  // EF1175 already in DB (scanner found it), seed updates family if needed
  p("EF1175", "NNSA MH200 Horeca",             "NNSA",      "Global", "In progress",   "Watch", 70, "PCS2_E"),
  p("EF1197", "NNSA Horeca UL-Variante",       "NNSA",      "US",     "Abgeschlossen", "Good", 100, "PCS2_E"),

  // ── Nivona VDE ────────────────────────────────────────────────────────────
  p("EF1187", "Nivona 1187",                   "Nivona",    "EU",     "In progress",   "Watch", 50, "PCS8_E"),
  p("EF1190", "Nivona 1190",                   "Nivona",    "EU",     "In progress",   "Watch", 50, "PCS8_E"),

  // ── Nano – Nachfolge EF720 (neu) ─────────────────────────────────────────
  p("EF1181", "Nano (Nachfolge EF720)",        "Nano",      "Global", "Approbation",   "Watch", 10, "PCS2_E", "2026-Q4"),

  // ── NDG Genio-S Re-Design (neu) ───────────────────────────────────────────
  p("EF1216", "NDG Genio-S 220–240V",          "NDG",       "EU",     "Approbation",   "Watch",  5, "PCS"),
  p("EF1217", "NDG Genio-S 120–127V",          "NDG",       "Global", "Approbation",   "Watch",  5, "PCS"),
  p("EF1218", "NDG Genio-S 100V",              "NDG",       "Global", "Approbation",   "Watch",  5, "PCS"),
  p("EF1221", "NDG Genio-S Variante",          "NDG",       "Global", "Approbation",   "Watch",  5, "PCS"),
];

// Variant groups: [all IDs in group, primary ID]
const variantGroups = [
  [["EF230", "EF231", "EF232"],                     "EF230"],
  [["EF254", "EF698"],                               "EF254"],
  [["EF258", "EF259"],                               "EF258"],
  [["EF275", "EF276"],                               "EF275"],
  [["EF375", "EF376"],                               "EF375"],
  [["EF384", "EF385"],                               "EF384"],
  [["EF392", "EF393"],                               "EF392"],
  [["EF395", "EF396"],                               "EF395"],
  [["EF397", "EF398"],                               "EF397"],
  [["EF503", "EF504"],                               "EF503"],
  [["EF565", "EF566"],                               "EF565"],
  [["EF577", "EF578"],                               "EF577"],
  [["EF590", "EF592"],                               "EF590"],
  [["EF693", "EF695", "EF696"],                      "EF693"],
  [["EF1006", "EF1007", "EF1008", "EF1009"],         "EF1006"],
  [["EF1016", "EF1017"],                             "EF1016"],
  [["EF1106", "EF1119", "EF1120", "EF1121", "EF1123"], "EF1106"],
  [["EF1156", "EF1157", "EF1234"],                   "EF1157"],
  [["EF1171"],                                        "EF1171"],   // standalone for now
  [["EF1175", "EF1197"],                              "EF1175"],
  [["EF1187", "EF1190"],                              "EF1187"],
  [["EF1216", "EF1217", "EF1218", "EF1221"],          "EF1216"],
];

// Also update family for already-existing projects the scanner found
const updateFamily = db.prepare(`
  UPDATE projects SET family = @family, market = @market, owner = @owner
  WHERE id = @id AND (family IS NULL OR family = '' OR family = 'CoffeeB M4' OR family = 'Reference project' OR family = 'Neu gefunden')
`);

const familyFixes = [
  { id: "EF1107", family: "Jura",   market: "EU",  owner: "PCS8_E" },
  { id: "EF1157", family: "Delica", market: "EU",  owner: "PCS2_E" },
  { id: "EF1175", family: "NNSA",   market: "Global", owner: "PCS2_E" },
  { id: "EF1234", family: "Delica", market: "Global", owner: "PCS2_E" },
];

const tx = db.transaction(() => {
  let inserted = 0, skipped = 0;

  for (const proj of projects) {
    const changes = insert.run(proj);
    if (changes.changes > 0) inserted++;
    else skipped++;
  }

  // Fix families of scanner-found projects
  for (const fix of familyFixes) {
    updateFamily.run(fix);
  }

  // Apply variant groups
  for (const [ids, primary] of variantGroups) {
    const group = ids.join("+");
    for (const id of ids) {
      setGroup.run({ group, id });
      if (id !== primary) setPrimary.run({ primary, id });
    }
  }

  console.log(`Inserted: ${inserted}, skipped (already exists): ${skipped}`);
  console.log("Variant groups applied:", variantGroups.length);
});

tx();
