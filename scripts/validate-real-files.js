/*
 * validate-real-files.js
 * -------------------------
 * Headless validation of Parts 1-3 (ask-once-remember column resolution,
 * two-tier memory, rate sanity backstop) against the REAL client files
 * that motivated them — no live Excel, no live proxy server. Uses
 * src/sheet-classifier/xlsx-test-helper.js (already used by regression-
 * test.js/sheet-memory-test.js for the committed synthetic demo file) to
 * read real .xlsx files, then drives the same pure pipeline pieces the
 * live add-in uses.
 *
 * NOT part of `npm run test:*` / CI — this hard-depends on gitignored
 * real client files ("Real company GST data used for local validation
 * only — never committed", see .gitignore) that will never exist in a
 * fresh clone. Skips (not fails) any file that isn't present.
 *
 * Run with: node scripts/validate-real-files.js
 */

// mentor-gst-reconciliation-ui.js transitively requires mentor-column-
// memory-ui.js and mentor-sheet-memory-ui.js, both of which instantiate a
// BrowserSheetMemoryStore at module load time -- its constructor calls
// window.localStorage.getItem(...), which throws in plain Node. This
// minimal stub (no document/Excel/fetch needed -- those are only
// referenced inside function bodies this script never calls) lets those
// files load so mentorComputeGstReconciliation (the REAL production
// composition of Step 1 + every Step 2 detector + rate-mismatch) can be
// reused directly, rather than re-derived here.
global.window = { localStorage: { getItem: () => null, setItem: () => {} } };

const fs = require("fs");
const path = require("path");
const { readWorkbookSheets } = require("../src/sheet-classifier/xlsx-test-helper.js");
const { recognizeGstSheets } = require("../src/gst-reconciliation/gst-reconciliation.js");
const { identifyGstWorkflow } = require("../src/gst-reconciliation/gst-workflows.js");
const { fieldsInPlay } = require("../src/gst-reconciliation/gst-column-ambiguity-rules.js");
const { resolveColumnField, rememberColumnField } = require("../src/gst-reconciliation/column-memory.js");
const { normalizeHeaderText } = require("../src/sheet-classifier/structural-signature.js");
const { JsonFileSheetMemoryStore } = require("../src/sheet-classifier/sheet-memory-store.js");
const { ColumnPatternStore } = require("../server/column-pattern-store.js");
const { mentorComputeGstReconciliation } = require("../src/taskpane/mentor-gst-reconciliation-ui.js");
const { buildStepTwoTotals } = require("../src/taskpane/gst-report-writer.js");
const { buildGstSummary } = require("../src/gst-reconciliation/gst-summary-builder.js");

const DEMO_DIR = path.join(__dirname, "..", "src", "demo-data");
const APR_AUG_EXPECTED_BOOKS_TOTAL = 23180361.87; // ₹2,31,80,361.87

const REAL_FILES = [
  { label: "RN Metals", file: "GST Reco Sheeet up to Dec.xlsx" },
  { label: "Basai Steel Traders full-year", file: "2A Reco till Feb.xlsx" },
  { label: "Basai Steel Traders Apr-Aug", file: "2A reco(April-Aug21) BST.xlsx" },
];

// Mirrors mentor-column-memory-ui.js's mentorResolveGstColumns() loop --
// explicitly duplicated here (not imported) rather than refactoring that
// already-shipped, Office.js-bound file to be Node-reusable: its Tier-1
// calls are fetch-based and its Tier-2 store is Excel-backed, neither of
// which runs in plain Node without scaffolding this one-off script
// doesn't warrant.
async function resolveColumnsForSheet({ role, sheetName, sheetSignals, columns, clientId, tier2Store, tier1Store }) {
  const inPlay = fieldsInPlay(role, columns.candidates);
  const allColumnIndices = sheetSignals.columns.map((c) => c.index);
  const resolved = {};
  const needsInput = [];

  for (const item of inPlay) {
    const r = await resolveColumnField({ clientId, sheetName, sheetSignals, fieldName: item.field, candidateIndices: item.candidateIndices, allColumnIndices, store: tier2Store });
    if (r.status === "resolved_from_memory") {
      resolved[item.field] = r.fieldIndex;
      continue;
    }

    let tier1Resolved = false;
    if (item.reason === "ambiguous" && tier1Store) {
      const candidateHeaders = r.candidates.map((c) => c.header);
      const tier1 = tier1Store.lookup(item.field, candidateHeaders);
      if (tier1.found) {
        const match = sheetSignals.columns.find((c) => normalizeHeaderText(c.header) === tier1.chosenHeader);
        if (match) {
          resolved[item.field] = match.index;
          tier1Resolved = true;
        }
      }
    }

    if (!tier1Resolved) needsInput.push({ field: item.field, reason: item.reason, candidates: r.candidates });
  }

  return { resolved, needsInput, inPlay };
}

function recognizeWorkflow(file) {
  const filePath = path.join(DEMO_DIR, file);
  if (!fs.existsSync(filePath)) return { skip: true };

  const sheets = readWorkbookSheets(filePath);
  const roleResults = recognizeGstSheets(sheets);
  const workflow = identifyGstWorkflow(roleResults);
  return { skip: false, roleResults, workflow };
}

async function runPass1(label, file) {
  const { skip, roleResults, workflow } = recognizeWorkflow(file);
  console.log(`\n=== ${label} (${file}) ===`);
  if (skip) {
    console.log(`  SKIP -- file not present in this environment`);
    return null;
  }

  for (const [name, r] of Object.entries(roleResults)) {
    console.log(`  '${name}' -> role: ${r.role}${r.role !== "unknown" ? " (confidence: " + r.confidence + ")" : ""}`);
  }

  if (!workflow || workflow.workflowId !== "gstr2a_vs_books") {
    console.log("  No gstr2a_vs_books workflow recognized -- nothing further to validate for this file.");
    return null;
  }

  const gstr2aSheet = workflow.sheets.gstr2a;
  const booksSheet = workflow.sheets.purchaseRegister;
  const clientId = "workbook:" + file;

  const tier2Store = new JsonFileSheetMemoryStore({ filePath: null });
  const tier1Store = new ColumnPatternStore({ filePath: null });

  const gstr2aRes = await resolveColumnsForSheet({ role: "gstr2a", sheetName: gstr2aSheet.sheetName, sheetSignals: gstr2aSheet.result.sheetSignals, columns: gstr2aSheet.result.columns, clientId, tier2Store, tier1Store });
  const booksRes = await resolveColumnsForSheet({ role: "purchase_register", sheetName: booksSheet.sheetName, sheetSignals: booksSheet.result.sheetSignals, columns: booksSheet.result.columns, clientId, tier2Store, tier1Store });

  console.log("  2A needs-input fields:   ", gstr2aRes.needsInput.length === 0 ? "(none)" : gstr2aRes.needsInput.map((n) => n.field + "(" + n.reason + ")").join(", "));
  console.log("  Books needs-input fields:", booksRes.needsInput.length === 0 ? "(none)" : booksRes.needsInput.map((n) => n.field + "(" + n.reason + ")").join(", "));

  for (const n of booksRes.needsInput.concat(gstr2aRes.needsInput)) {
    if (n.reason === "ambiguous") {
      console.log(`    -- '${n.field}' candidates: ${n.candidates.map((c) => JSON.stringify(c.header)).join(", ")}`);
    }
  }

  return { workflow, gstr2aSheet, booksSheet, gstr2aRes, booksRes, clientId };
}

// Resolves any still-ambiguous field using a documented, sensible
// heuristic, via the SAME rememberColumnField/Tier-1-remember calls a real
// user's answer would trigger. Prints what it picked -- this is a real,
// inspectable choice, not a hidden shortcut.
//
// gstin is special-cased: it's a CONTENT-based field (identifyGstColumns
// scores it by how many rows actually look like a valid GSTIN, via
// computeGstinRatios -- see gst-column-identifier.js), not a header-text
// field, so the highest-content-match-ratio candidate is the correct
// signal to pick by -- NOT header length. A first attempt at this script
// used "shortest header" uniformly, which picked a column literally
// labeled "vlookup" for gstin (shorter than "GSTIN of supplier") purely
// because both columns' raw content happened to score above the
// ambiguity margin -- exactly the kind of wrong-guess-by-a-shallow-
// heuristic this whole feature exists to prevent a human from doing
// blindly. Every other (header-based) field keeps the shortest/most-exact
// heading heuristic (e.g. bare "CGST" over "Input CGST@9%").
async function seedResolution(pass1Result) {
  const { gstr2aSheet, booksSheet, gstr2aRes, booksRes, clientId } = pass1Result;
  const tier2Store = new JsonFileSheetMemoryStore({ filePath: null });
  const tier1Store = new ColumnPatternStore({ filePath: null });

  for (const { sheet, res } of [
    { sheet: gstr2aSheet, res: gstr2aRes },
    { sheet: booksSheet, res: booksRes },
  ]) {
    for (const n of res.needsInput) {
      if (n.reason !== "ambiguous" || n.candidates.length === 0) continue;

      let chosen;
      if (n.field === "gstin") {
        // Ratio-sorted descending already, by identifyGstColumns.
        const byRatio = sheet.result.columns.candidates.gstin;
        const best = byRatio[0];
        chosen = n.candidates.find((c) => c.index === best.index);
        console.log(`  Seeding '${sheet.sheetName}'.gstin -> "${chosen.header}" (content match ratio ${best.ratio.toFixed(2)}, vs ${byRatio.slice(1).map((r) => r.ratio.toFixed(2)).join(", ")} for the other candidate(s))`);
      } else {
        chosen = n.candidates.slice().sort((a, b) => String(a.header).length - String(b.header).length)[0];
        console.log(`  Seeding '${sheet.sheetName}'.${n.field} -> "${chosen.header}" (shortest of: ${n.candidates.map((c) => JSON.stringify(c.header)).join(", ")})`);
      }

      await rememberColumnField({ clientId, sheetName: sheet.sheetName, sheetSignals: sheet.result.sheetSignals, fieldName: n.field, chosenColumnIndex: chosen.index, store: tier2Store });
      tier1Store.remember(n.field, n.candidates.map((c) => c.header), chosen.header);
    }
  }

  return { tier2Store, tier1Store };
}

async function runPass2(label, pass1Result, seededStores) {
  if (!pass1Result) return;
  const { gstr2aSheet, booksSheet, clientId } = pass1Result;
  const { tier2Store, tier1Store } = seededStores;

  const gstr2aRes = await resolveColumnsForSheet({ role: "gstr2a", sheetName: gstr2aSheet.sheetName, sheetSignals: gstr2aSheet.result.sheetSignals, columns: gstr2aSheet.result.columns, clientId, tier2Store, tier1Store });
  const booksRes = await resolveColumnsForSheet({ role: "purchase_register", sheetName: booksSheet.sheetName, sheetSignals: booksSheet.result.sheetSignals, columns: booksSheet.result.columns, clientId, tier2Store, tier1Store });

  console.log(`\n--- ${label}: pass 2 (after seeding) ---`);
  console.log("  2A needs-input fields:   ", gstr2aRes.needsInput.length === 0 ? "(none)" : JSON.stringify(gstr2aRes.needsInput.map((n) => n.field)));
  console.log("  Books needs-input fields:", booksRes.needsInput.length === 0 ? "(none)" : JSON.stringify(booksRes.needsInput.map((n) => n.field)));

  if (gstr2aRes.needsInput.length > 0 || booksRes.needsInput.length > 0) {
    console.log("  Still unresolved after seeding -- cannot compute a trustworthy total.");
    return;
  }

  const resolvedGstr2aSheet = { ...gstr2aSheet, result: { ...gstr2aSheet.result, columns: { ...gstr2aSheet.result.columns, ...gstr2aRes.resolved } } };
  const resolvedBooksSheet = { ...booksSheet, result: { ...booksSheet.result, columns: { ...booksSheet.result.columns, ...booksRes.resolved } } };

  const proposal = mentorComputeGstReconciliation(resolvedGstr2aSheet, resolvedBooksSheet);
  const stepTwoTotals = buildStepTwoTotals(proposal.wrongHeadResult, proposal.crossSheetWrongHeadResult, proposal.rcmBySource, proposal.itcBySource, proposal.dupBySource, proposal.rateMismatchBySource);
  const summary = buildGstSummary(proposal.comparisonSummary, stepTwoTotals, proposal.invoiceLevelExtrasResult);

  console.log(`  Total tax per 2A:    Rs ${summary.totalTaxPer2A.toLocaleString("en-IN")}`);
  console.log(`  Total tax per Books: Rs ${summary.totalTaxPerBooks.toLocaleString("en-IN")}`);

  if (label === "Basai Steel Traders Apr-Aug") {
    const diff = Math.abs(summary.totalTaxPerBooks - APR_AUG_EXPECTED_BOOKS_TOTAL);
    if (diff < 1) {
      console.log(`  MATCH: Books total matches the expected Rs ${APR_AUG_EXPECTED_BOOKS_TOTAL.toLocaleString("en-IN")}`);
    } else {
      console.log(`  MISMATCH: expected Rs ${APR_AUG_EXPECTED_BOOKS_TOTAL.toLocaleString("en-IN")}, got Rs ${summary.totalTaxPerBooks.toLocaleString("en-IN")} (diff Rs ${diff.toFixed(2)})`);
    }
  }
}

async function main() {
  console.log("--- Pass 1: fresh stores, no prior memory ---");
  const pass1Results = {};
  for (const { label, file } of REAL_FILES) {
    pass1Results[label] = await runPass1(label, file);
  }

  console.log("\n\n--- Seed step (Basai Steel Traders Apr-Aug only) ---");
  const aprAug = pass1Results["Basai Steel Traders Apr-Aug"];
  if (aprAug) {
    const seededStores = await seedResolution(aprAug);
    await runPass2("Basai Steel Traders Apr-Aug", aprAug, seededStores);
  } else {
    console.log("  SKIP -- file not present.");
  }

  const fullYear = pass1Results["Basai Steel Traders full-year"];
  if (fullYear) {
    const seededStores = await seedResolution(fullYear);
    await runPass2("Basai Steel Traders full-year", fullYear, seededStores);
  }
}

main().catch((err) => {
  console.error("validate-real-files.js error:", err);
  process.exitCode = 1;
});
