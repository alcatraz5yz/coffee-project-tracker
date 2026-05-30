/**
 * tabelle30-match.js — Cross-check new Excel parts against the parts already listed
 * in the current Tabelle 30 (Resistance to heat and fire) of the Word report.
 *
 * Goal (compliance-critical): only the *genuinely* new parts should be added.
 *   - A part is treated as ALREADY PRESENT only when BOTH its part identity
 *     (same object noun + same part number) AND its material grade strongly
 *     match an existing row. Those are excluded from the add-list.
 *   - Anything else stays on the add-list ("new"), but if it partially overlaps
 *     an existing row (shares a part number, or shares a material grade) it gets
 *     a non-blocking warning pointing at the similar row, so the engineer can
 *     double-check. Nothing is ever silently dropped.
 *
 * Why not a looser match? A naive "share ≥2 words / same polymer family" check
 * wrongly fuses different parts: e.g. "Carrier 268 thermobl" (new, PP GFPP-30)
 * would collide with the existing "Holder 268 thermobl" (PP GH43) purely on the
 * shared words "268 thermobl" and the generic "PP … black". That would drop a
 * real new part. We therefore match on *distinctive* material grades (trade
 * names / grade codes), not polymer family + colour, and require the primary
 * object noun to match for a strong part hit.
 */

"use strict";

// Polymer families and colours are NOT distinctive — every ABS-black part shares them.
const POLYMER = new Set([
  "abs", "pp", "pc", "pa", "pom", "ptfe", "nbr", "tpe", "tpu", "pe", "pvc",
  "pa6", "pa66", "pa6t", "pa6i", "san", "asa", "pmma", "pbt", "pet", "pps",
  "peek", "ppo", "pa12", "silicone", "silikon", "rubber", "gummi", "lacquer", "lack",
]);
const COLOR = new Set([
  "black", "schwarz", "white", "weiss", "grey", "gray", "grau", "red", "rot",
  "green", "grun", "blue", "blau", "transparent", "natural", "natur", "yellow",
  "gelb", "silver", "silber", "brown", "braun", "orange",
]);
const GENERIC = new Set([
  "pos", "position", "alternative", "alternativ", "alternate", "cpl", "complete",
  "komplett", "und", "and", "der", "die", "das", "mit", "for", "the", "new",
  "old", "ef", "type", "model", "material", "teil", "part",
]);

function norm(s) {
  return String(s || "")
    .toLowerCase()
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}
function tokens(s) {
  return norm(s).split(" ").filter(Boolean);
}
function isVariant(t) {
  return /^v\d+$/.test(t); // V2, V3 …
}

// Distinctive material identifiers: trade names (alpha, length ≥ 4) or grade codes
// (mixed letters+digits like "121h", "gfpp", "6265x"). Polymer family, colour and
// pure numbers (colour numbers, weights) are excluded as too generic.
function gradeTokens(material) {
  const out = new Set();
  for (const t of tokens(material)) {
    if (POLYMER.has(t) || COLOR.has(t) || GENERIC.has(t)) continue;
    if (/^\d+$/.test(t)) continue;
    const isGradeCode = /[a-z]/.test(t) && /\d/.test(t);
    if (t.length >= 4 || isGradeCode) out.add(t);
  }
  return out;
}

// Part numbers in a part name: standalone 3–4 digit groups (1157, 1068, 268 …).
function partNumbers(part) {
  const out = new Set();
  for (const m of norm(part).matchAll(/\b(\d{3,4})\b/g)) out.add(m[1]);
  return out;
}

// Primary object noun = the first meaningful alpha token (Holder, Cover, Carrier …).
function primaryObject(part) {
  for (const t of tokens(part)) {
    if (/^\d+$/.test(t) || isVariant(t)) continue;
    if (GENERIC.has(t) || COLOR.has(t)) continue;
    if (t.length >= 3) return t;
  }
  return "";
}

function setsShare(a, b) {
  for (const x of a) if (b.has(x)) return true;
  return false;
}
function sharedList(a, b) {
  return [...a].filter((x) => b.has(x));
}

/**
 * Build a lookup over the existing Tabelle 30 rows (from parse_tabelle30 output).
 * Word columns: cells[0] = Object/Part, cells[2] = Type/Model (material).
 * Footnote / blank rows (no material) are ignored.
 */
function indexWordRows(wordRows) {
  const idx = [];
  for (const r of wordRows || []) {
    const cells = r.cells || [];
    const partText = cells[0] || "";
    const materialText = cells[2] || "";
    if (!materialText.trim()) continue; // footnote / supplementary rows
    idx.push({
      rowIdx: r.rowIdx,
      part: String(partText).replace(/\s+/g, " ").trim(),
      material: String(materialText).replace(/\s+/g, " ").trim(),
      object: primaryObject(partText),
      numbers: partNumbers(partText),
      grades: gradeTokens(materialText),
    });
  }
  return idx;
}

/**
 * Classify one Excel new-part entry against the existing rows.
 * Returns { status: "new" | "present", warning?: string, matchRowIdx?: number }.
 */
function classifyEntry(excelEntry, wordIndex) {
  const r = excelEntry.excelRow || excelEntry;
  const partName = (r.partName || "").split("\n")[0];
  const material = r.changeTo || "";

  const obj = primaryObject(partName);
  const nums = partNumbers(partName);
  const grades = gradeTokens(material);

  let strongMatch = null; // same object + same number AND shared grade → already present
  let gradeMatch = null;  // same material grade on some row
  let numberMatch = null; // shares a part number with some row

  for (const w of wordIndex) {
    const sameObject = obj && w.object && obj === w.object;
    const sharedNumber = setsShare(nums, w.numbers);
    const sharedGrade = setsShare(grades, w.grades);

    if (sameObject && sharedNumber && sharedGrade && !strongMatch) strongMatch = w;
    if (sharedGrade && !gradeMatch) gradeMatch = w;
    if (sharedNumber && !numberMatch) numberMatch = w;
  }

  if (strongMatch) {
    return {
      status: "present",
      matchRowIdx: strongMatch.rowIdx,
      warning: `bereits in Tabelle 30 (Zeile ${strongMatch.rowIdx}: ${strongMatch.part} — ${strongMatch.material})`,
    };
  }
  if (gradeMatch) {
    const shared = sharedList(grades, gradeMatch.grades).join(", ");
    return {
      status: "new",
      matchRowIdx: gradeMatch.rowIdx,
      warning: `gleiche Materialsorte (${shared}) wie Zeile ${gradeMatch.rowIdx}: ${gradeMatch.part} — prüfen`,
    };
  }
  if (numberMatch) {
    return {
      status: "new",
      matchRowIdx: numberMatch.rowIdx,
      warning: `ähnliche Position wie Zeile ${numberMatch.rowIdx}: ${numberMatch.part} — prüfen`,
    };
  }
  return { status: "new" };
}

/**
 * Classify a list of new Excel entries. Returns:
 *   { add: [{ ...entry, _match }], present: [{ ...entry, _match }] }
 * where add = stays on the green add-list (may carry a warning), present = excluded.
 */
function classifyNewEntries(newEntries, wordRows) {
  const wordIndex = indexWordRows(wordRows);
  const add = [];
  const present = [];
  for (const entry of newEntries || []) {
    const match = classifyEntry(entry, wordIndex);
    const tagged = { ...entry, _match: match };
    if (match.status === "present") present.push(tagged);
    else add.push(tagged);
  }
  return { add, present, wordRowCount: wordIndex.length };
}

module.exports = {
  classifyNewEntries,
  classifyEntry,
  indexWordRows,
  // exported for testing
  gradeTokens,
  partNumbers,
  primaryObject,
};
