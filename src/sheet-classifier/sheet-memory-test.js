/*
 * sheet-memory-test.js
 * ---------------------
 * Proves the sheet-memory layer does what it's for:
 *   1. An unrecognized sheet triggers a prompt, not silence.
 *   2. Once answered, the SAME shape is recognized again even if the sheet
 *      is renamed and has grown more rows.
 *   3. A shape that's drifted slightly (one column added) still matches via
 *      the fuzzy fallback, not just an exact string match.
 *   4. Memory is scoped per client — another client with an identically-
 *      shaped sheet is still asked.
 *   5. A sheet the rule-based classifier (Step 2) already recognized never
 *      touches the memory store at all.
 * Then repeats the "unrecognized -> answer -> recognized on rename" flow
 * against two of the REAL demo workbook's genuinely unknown sheets, so this
 * isn't only proven on hand-built fixtures.
 *
 * Finally, a cross-type collision investigation: two genuinely different
 * sheet types (a Vendor List and a Payroll List) for the SAME client,
 * constructed to be close in size/shape but not identical — does the fuzzy
 * match confuse them? This is a reporting exercise, not just pass/fail: see
 * the FINDING lines in the output and the summary printed at the end.
 *
 * Run with: node src/sheet-classifier/sheet-memory-test.js
 * Uses an in-memory store (no file writes) so repeated runs stay deterministic.
 */

const path = require("path");
const { readWorkbookSheets } = require("./xlsx-test-helper.js");
const { extractSheetSignals, extractWorkbookSignals } = require("./signal-extractor.js");
const { classifySheet, classifyWorkbook } = require("./classifier.js");
const { JsonFileSheetMemoryStore } = require("./sheet-memory-store.js");
const { resolveSheetLabel, rememberSheetLabel } = require("./sheet-memory.js");
const { computeStructuralSignature, tagSequenceSimilarity } = require("./structural-signature.js");

const DEMO_FILE = path.join(__dirname, "..", "demo-data", "Three_Sisters_Kitchen_MENTOR_Demo.xlsx");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

// A custom sheet shape none of the four known rule sets recognize, and with
// no date column at all — guaranteed "unknown" from classifySheet, so this
// test isolates memory behavior from classifier-margin edge cases.
function buildInventoryValues(rowCount, options) {
  const opts = options || {};
  const categories = ["Produce", "Dairy", "Dry Goods"];
  const header = ["SKU", "Item Name", "Category", "Qty On Hand", "Reorder Level"];
  if (opts.includeSupplierColumn) header.push("Supplier");

  const rows = [header];
  for (let i = 1; i <= rowCount; i++) {
    const row = ["SKU-" + String(i).padStart(4, "0"), "Item " + i, categories[i % categories.length], 50 + (i % 40), [5, 10, 20][i % 3]];
    if (opts.includeSupplierColumn) row.push(["North Farms Ltd", "Green Valley Co", "Dairy Direct"][i % 3]);
    rows.push(row);
  }
  return rows;
}

async function runSyntheticScenarios() {
  console.log("-- Synthetic scenarios --\n");

  const store = new JsonFileSheetMemoryStore({ filePath: null }); // in-memory only

  const v1 = extractSheetSignals("Inventory Snapshot", buildInventoryValues(20));
  const v1Classification = classifySheet(v1);
  assert(v1Classification.type === "unknown", "custom inventory shape (no date column) classifies as unknown");

  // 1. First encounter — should prompt, and must not silently guess.
  const first = await resolveSheetLabel({ clientId: "client-A", sheetName: "Inventory Snapshot", sheetSignals: v1, classification: v1Classification, store });
  assert(first.status === "needs_input", "unseen shape returns needs_input");
  assert(first.prompt === "I don't recognize this sheet ('Inventory Snapshot') — what is it?", "prompt text matches the expected wording");

  // User answers the prompt.
  await rememberSheetLabel({ clientId: "client-A", sheetName: "Inventory Snapshot", sheetSignals: v1, userProvidedLabel: "Weekly inventory count", store });
  assert((await store.list("client-A")).length === 1, "answer is persisted to the store");

  // 2. Same client, sheet renamed AND grown by 15 rows — still the same shape.
  const v2 = extractSheetSignals("Stock Levels", buildInventoryValues(35));
  const v2Classification = classifySheet(v2);
  const renamed = await resolveSheetLabel({ clientId: "client-A", sheetName: "Stock Levels", sheetSignals: v2, classification: v2Classification, store });
  assert(renamed.status === "remembered", "renamed + grown sheet is recognized without re-asking");
  assert(renamed.label === "Weekly inventory count", "recognized sheet returns the previously given label");
  assert(renamed.matchedVia === "exact", "identical column shape matches via the exact fast path");

  // 3. Same client, one new column added ("Supplier") — minor structural drift.
  const v3 = extractSheetSignals("Stock Levels (New)", buildInventoryValues(28, { includeSupplierColumn: true }));
  const v3Classification = classifySheet(v3);
  const drifted = await resolveSheetLabel({ clientId: "client-A", sheetName: "Stock Levels (New)", sheetSignals: v3, classification: v3Classification, store });
  assert(drifted.status === "remembered", "sheet with one added column still matches via fuzzy fallback");
  assert(drifted.matchedVia === "fuzzy", "drifted shape is reported as a fuzzy match, not exact");
  assert(typeof drifted.similarity === "number" && drifted.similarity < 1, "fuzzy match reports a similarity below 1.0");

  // 4. A DIFFERENT client with the identical original shape — must not see client-A's answer.
  const otherClient = await resolveSheetLabel({ clientId: "client-B", sheetName: "Inventory Snapshot", sheetSignals: v1, classification: classifySheet(v1), store });
  assert(otherClient.status === "needs_input", "identically-shaped sheet for a different client is still asked (no cross-client leakage)");

  // 5. A sheet the classifier already recognized should never touch the store.
  const listBefore = (await store.list("client-A")).length;
  const classified = await resolveSheetLabel({
    clientId: "client-A",
    sheetName: "Raw - Supplier Invoices",
    sheetSignals: v1, // signals are irrelevant here — classification short-circuits before signature lookup
    classification: { type: "invoice_list", displayLabel: "Invoice List" },
    store,
  });
  assert(classified.status === "classified" && classified.label === "Invoice List", "already-classified sheet returns immediately");
  assert((await store.list("client-A")).length === listBefore, "already-classified sheet never touches the memory store");

  console.log("");
}

async function runRealDemoDataFlow() {
  console.log("-- Real demo-data flow (genuinely unknown sheets from Step 3) --\n");

  const sheets = readWorkbookSheets(DEMO_FILE);
  const signals = extractWorkbookSignals(sheets);
  const classifications = classifyWorkbook(signals);
  const signalsByName = new Map(signals.map((s) => [s.sheetName, s]));
  const classificationByName = new Map(classifications.map((c) => [c.sheetName, c]));

  const store = new JsonFileSheetMemoryStore({ filePath: null });
  const clientId = "three-sisters-kitchen";

  const targets = ["Monthly P&L (Answer Key)", "Variance Analysis"];

  for (const sheetName of targets) {
    const sheetSignals = signalsByName.get(sheetName);
    const classification = classificationByName.get(sheetName);
    assert(classification.type === "unknown", `'${sheetName}' is unknown per the Step 3 regression check`);

    const first = await resolveSheetLabel({ clientId, sheetName, sheetSignals, classification, store });
    assert(first.status === "needs_input", `'${sheetName}' prompts on first encounter`);

    await rememberSheetLabel({ clientId, sheetName, sheetSignals, userProvidedLabel: `${sheetName} (client-defined report)`, store });

    // Simulate the client renaming the tab next month — same signals, new name.
    const renamedResult = await resolveSheetLabel({
      clientId,
      sheetName: sheetName + " (renamed)",
      sheetSignals,
      classification,
      store,
    });
    assert(renamedResult.status === "remembered" && renamedResult.matchedVia === "exact", `'${sheetName}' is recognized after being renamed, without re-prompting`);
  }

  console.log("");
}

// ---------------------------------------------------------------
// CROSS-TYPE COLLISION INVESTIGATION
//
// Two genuinely different sheet types for the same client, constructed to
// be close in size and column mix (not identical) — does the fuzzy-match
// path confuse them? Uses the exact production path (store + resolveSheetLabel/
// rememberSheetLabel), not a shortcut call into tagSequenceSimilarity, so
// this reflects what would actually happen to a real client's workbook.
// ---------------------------------------------------------------
const findings = [];
function reportFinding(text) {
  findings.push(text);
  console.log(`FINDING  ${text}`);
}

function buildVendorListValues(rowCount) {
  const categories = ["Food", "Beverage", "Cleaning", "Equipment"];
  const paymentTerms = [14, 30, 45, 60];
  const header = ["Vendor ID", "Vendor Name", "Category", "Payment Terms (days)", "Active", "Primary Contact Email"];
  const rows = [header];
  for (let i = 1; i <= rowCount; i++) {
    rows.push(["V-" + String(i).padStart(3, "0"), "Vendor " + i + " Ltd", categories[i % categories.length], paymentTerms[i % paymentTerms.length], i % 5 === 0 ? "N" : "Y", "contact" + i + "@vendor" + i + ".com"]);
  }
  return rows;
}

// Close variant: differs from the vendor list in exactly ONE column's shape
// (Hourly Rate is a varied numeric amount where Vendor's last column is a
// unique text field). Everything else — ID/Name/Category/low-cardinality-
// numeric/status shape — is the same, because that's a genuinely common
// pattern across many unrelated record types, not a contrived edge case.
function buildPayrollListValuesCloseVariant(rowCount) {
  const departments = ["Kitchen", "Front of House", "Management"];
  const contractedHours = [20, 30, 40];
  const employmentTypes = ["Full-time", "Part-time", "Zero-hours"];
  const header = ["Employee ID", "Employee Name", "Department", "Weekly Contracted Hours", "Employment Type", "Hourly Rate (£)"];
  const rows = [header];
  for (let i = 1; i <= rowCount; i++) {
    rows.push(["E-" + String(i).padStart(3, "0"), "Employee " + i, departments[i % departments.length], contractedHours[i % contractedHours.length], employmentTypes[i % employmentTypes.length], (9 + (i % 12) * 0.75).toFixed(2)]);
  }
  return rows;
}

// More clearly differentiated variant: TWO columns diverge from the vendor
// list instead of one (Contract Notes is free text where Vendor's Active
// column is a fixed Y/N category, in addition to the Hourly Rate difference).
function buildPayrollListValuesDistinctVariant(rowCount) {
  const departments = ["Kitchen", "Front of House", "Management"];
  const contractedHours = [20, 30, 40];
  const notes = ["Reviewed annually, no flexibility clause", "Seasonal contract, ends after summer", "Union-negotiated rate, see HR file", "Standard contract", "Probation period until month 3"];
  const header = ["Employee ID", "Employee Name", "Department", "Weekly Contracted Hours", "Contract Notes", "Hourly Rate (£)"];
  const rows = [header];
  for (let i = 1; i <= rowCount; i++) {
    rows.push(["E-" + String(i).padStart(3, "0"), "Employee " + i, departments[i % departments.length], contractedHours[i % contractedHours.length], notes[i % notes.length] + " (" + i + ")", (9 + (i % 12) * 0.75).toFixed(2)]);
  }
  return rows;
}

async function runCrossTypeCollisionInvestigation() {
  console.log('-- Cross-type collision investigation (Vendor List vs. Payroll List, same client) --\n');

  const clientId = "collision-test-client";
  const vendorSignals = extractSheetSignals("Approved Vendors", buildVendorListValues(16));
  const vendorClassification = classifySheet(vendorSignals);
  assert(vendorClassification.type === "unknown", "Vendor List classifies as unknown (reaches the memory layer, as intended)");

  // --- Case 1: the "close" variant — differs in exactly one column's shape ---
  {
    const store = new JsonFileSheetMemoryStore({ filePath: null });
    await rememberSheetLabel({ clientId, sheetName: "Approved Vendors", sheetSignals: vendorSignals, userProvidedLabel: "Approved supplier list", store });

    const payrollSignals = extractSheetSignals("Employee Payroll", buildPayrollListValuesCloseVariant(16));
    const payrollClassification = classifySheet(payrollSignals);
    assert(payrollClassification.type === "unknown", "Payroll List (close variant) classifies as unknown (reaches the memory layer, as intended)");

    // Compute the raw similarity directly, not just whatever resolveSheetLabel
    // happens to surface — we want the actual number regardless of outcome.
    const similarity = tagSequenceSimilarity(computeStructuralSignature(vendorSignals).tags, computeStructuralSignature(payrollSignals).tags);
    const result = await resolveSheetLabel({ clientId, sheetName: "Employee Payroll", sheetSignals: payrollSignals, classification: payrollClassification, store });

    console.log(`  Close variant (1 column differs): status=${result.status}, similarity=${(similarity * 100).toFixed(1)}%`);
    console.log(`  Cutoff is 80% — margin: ${((similarity - 0.8) * 100).toFixed(1)} points ${similarity >= 0.8 ? "OVER" : "under"} the threshold`);

    if (result.status === "remembered") {
      reportFinding(
        `Vendor List vs. Payroll List, differing in only 1 of 6 columns, scored ${(similarity * 100).toFixed(1)}% similarity — ` +
          `at/above the 80% cutoff. Payroll List was fuzzy-matched to Vendor's remembered label ("${result.label}"), which is wrong — ` +
          "these are genuinely different sheet types that happen to share a common ID/Name/Category/low-cardinality-numeric/status shape."
      );
    } else {
      console.log("  No collision for this variant.");
    }
  }

  // --- Case 2: the "clearly differentiated" variant — differs in two columns ---
  {
    const store = new JsonFileSheetMemoryStore({ filePath: null });
    await rememberSheetLabel({ clientId, sheetName: "Approved Vendors", sheetSignals: vendorSignals, userProvidedLabel: "Approved supplier list", store });

    const payrollSignals = extractSheetSignals("Employee Payroll", buildPayrollListValuesDistinctVariant(16));
    const payrollClassification = classifySheet(payrollSignals);

    const similarity = tagSequenceSimilarity(computeStructuralSignature(vendorSignals).tags, computeStructuralSignature(payrollSignals).tags);
    const result = await resolveSheetLabel({ clientId, sheetName: "Employee Payroll", sheetSignals: payrollSignals, classification: payrollClassification, store });

    console.log(`\n  Distinct variant (2 columns differ): status=${result.status}, similarity=${(similarity * 100).toFixed(1)}%`);
    console.log(`  Cutoff is 80% — margin: ${((similarity - 0.8) * 100).toFixed(1)} points ${similarity >= 0.8 ? "OVER" : "under"} the threshold`);
    assert(result.status === "needs_input", "Payroll List (distinct variant, 2 columns differ) is correctly NOT fuzzy-matched to Vendor's label");
  }

  // --- Case 3: characterize the crossover mathematically across column counts ---
  // A single differing column among n total columns produces similarity = 1 - 1/n.
  // This finds exactly where that crosses the 80% cutoff.
  console.log("\n  Single-column-difference similarity by sheet size (same fuzzy-match function, synthetic tag sequences):");
  const crossoverRows = [];
  for (const n of [3, 4, 5, 6, 7, 8, 10, 12]) {
    const a = Array.from({ length: n }, (_, i) => "TAG" + i);
    const b = a.slice();
    b[0] = "DIFFERENT";
    const sim = tagSequenceSimilarity(a, b);
    crossoverRows.push({ n, sim });
    console.log(`    n=${String(n).padStart(2)} columns: ${(sim * 100).toFixed(1)}%  ${sim >= 0.8 ? "(>= 80% cutoff — WOULD fuzzy-match)" : ""}`);
  }
  const crossoverPoint = crossoverRows.find((r) => r.sim >= 0.8);
  reportFinding(
    `Because similarity = 1 - editDistance/columnCount, a SINGLE differing column already reaches or exceeds the 80% cutoff for any ` +
      `sheet with ${crossoverPoint.n}+ columns (exact crossover: ${crossoverPoint.n} columns = ${(crossoverPoint.sim * 100).toFixed(1)}%). Every sheet in our own ` +
      "demo dataset (Bank Statement: 8 cols, Payroll Summary: 8 cols, Supplier Invoices: 11 cols) is well past that crossover — " +
      "meaning at realistic sheet sizes, ANY two sheets differing by only one column's shape will be treated as the same signature."
  );

  // --- Case 4: the sharpest evidence — legitimate drift and illegitimate collision produce the SAME number ---
  const driftSimilarity = tagSequenceSimilarity(
    ["TEXT_UNIQUE", "TEXT_UNIQUE", "TEXT_CATEGORICAL", "NUM_LOW", "NUM_LOW"],
    ["TEXT_UNIQUE", "TEXT_UNIQUE", "TEXT_CATEGORICAL", "NUM_LOW", "NUM_LOW", "TEXT_MODERATE"]
  );
  reportFinding(
    `The legitimate "client added one column" drift case (Step 2's own regression fixture) scores ${(driftSimilarity * 100).toFixed(1)}% similarity — ` +
      "statistically indistinguishable from the illegitimate Vendor-vs-Payroll collision above. No single similarity threshold can separate " +
      "'same sheet, minor edit' from 'different sheet, coincidental overlap' using column-shape tags alone — they can and do produce identical scores."
  );

  console.log("");
}

async function run() {
  await runSyntheticScenarios();
  await runRealDemoDataFlow();
  await runCrossTypeCollisionInvestigation();

  if (findings.length > 0) {
    console.log(`\n${findings.length} finding(s) from the cross-type collision investigation (see FINDING lines above) — design gap, not a broken test.`);
  }

  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All sheet-memory checks passed.");
  }
}

run();
