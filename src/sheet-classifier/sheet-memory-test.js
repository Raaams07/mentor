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
 * constructed to be close in size/shape but not identical. This originally
 * found a real gap — column-shape similarity alone couldn't tell a
 * legitimate one-column-drift edit apart from a coincidental cross-type
 * shape overlap (both landed at the same 83.3%). header_signature +
 * headerSequenceSimilarity was added as a required second signal to fix
 * it (see sheet-memory-store.js) — this section proves the fix: the
 * close-variant collision now correctly returns needs_input, while the
 * legitimate drift case still fuzzy-matches. See the FINDING line for the
 * exact numbers.
 *
 * Run with: node src/sheet-classifier/sheet-memory-test.js
 * Uses an in-memory store (no file writes) so repeated runs stay deterministic.
 */

const path = require("path");
const { readWorkbookSheets } = require("./xlsx-test-helper.js");
const { extractSheetSignals, extractWorkbookSignals } = require("./signal-extractor.js");
const { classifySheet, classifyWorkbook } = require("./classifier.js");
const { JsonFileSheetMemoryStore } = require("./sheet-memory-store.js");
const { resolveSheetLabel, rememberSheetLabel, updateStoredSheetLabel, forgetSheetLabel } = require("./sheet-memory.js");
const { computeStructuralSignature, tagSequenceSimilarity, headerSequenceSimilarity } = require("./structural-signature.js");

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
  const vendorSig = computeStructuralSignature(vendorSignals);
  assert(vendorClassification.type === "unknown", "Vendor List classifies as unknown (reaches the memory layer, as intended)");

  // --- Case 1: the "close" variant — differs in exactly one column's shape ---
  // This is the exact case that used to collide (structural similarity 83.3%,
  // above the 80% cutoff) before the header-text corroboration was added.
  {
    const store = new JsonFileSheetMemoryStore({ filePath: null });
    await rememberSheetLabel({ clientId, sheetName: "Approved Vendors", sheetSignals: vendorSignals, userProvidedLabel: "Approved supplier list", store });

    const payrollSignals = extractSheetSignals("Employee Payroll", buildPayrollListValuesCloseVariant(16));
    const payrollClassification = classifySheet(payrollSignals);
    assert(payrollClassification.type === "unknown", "Payroll List (close variant) classifies as unknown (reaches the memory layer, as intended)");

    const payrollSig = computeStructuralSignature(payrollSignals);
    const structuralSimilarity = tagSequenceSimilarity(vendorSig.tags, payrollSig.tags);
    const headerSimilarity = headerSequenceSimilarity(vendorSig.headerTokens, payrollSig.headerTokens);
    const result = await resolveSheetLabel({ clientId, sheetName: "Employee Payroll", sheetSignals: payrollSignals, classification: payrollClassification, store });

    console.log(`  Close variant (1 column differs): status=${result.status}`);
    console.log(`    structural similarity=${(structuralSimilarity * 100).toFixed(1)}%  (cutoff 80%, ${((structuralSimilarity - 0.8) * 100).toFixed(1)} pts ${structuralSimilarity >= 0.8 ? "OVER" : "under"})`);
    console.log(`    header similarity=${(headerSimilarity * 100).toFixed(1)}%  (cutoff 50%, ${((headerSimilarity - 0.5) * 100).toFixed(1)} pts ${headerSimilarity >= 0.5 ? "OVER" : "under"})`);

    assert(result.status === "needs_input", "Payroll List (close variant) is correctly NOT fuzzy-matched to Vendor's label — header check blocks it despite structural similarity clearing 80%");

    if (result.status === "remembered") {
      reportFinding(
        `Vendor List vs. Payroll List, differing in only 1 of 6 columns, scored ${(structuralSimilarity * 100).toFixed(1)}% structural similarity ` +
          `and ${(headerSimilarity * 100).toFixed(1)}% header similarity — both cleared their cutoffs. Payroll List was fuzzy-matched to Vendor's remembered ` +
          `label ("${result.label}"), which is wrong.`
      );
    }
  }

  // --- Case 2: the "clearly differentiated" variant — differs in two columns ---
  {
    const store = new JsonFileSheetMemoryStore({ filePath: null });
    await rememberSheetLabel({ clientId, sheetName: "Approved Vendors", sheetSignals: vendorSignals, userProvidedLabel: "Approved supplier list", store });

    const payrollSignals = extractSheetSignals("Employee Payroll", buildPayrollListValuesDistinctVariant(16));
    const payrollClassification = classifySheet(payrollSignals);
    const payrollSig = computeStructuralSignature(payrollSignals);

    const structuralSimilarity = tagSequenceSimilarity(vendorSig.tags, payrollSig.tags);
    const headerSimilarity = headerSequenceSimilarity(vendorSig.headerTokens, payrollSig.headerTokens);
    const result = await resolveSheetLabel({ clientId, sheetName: "Employee Payroll", sheetSignals: payrollSignals, classification: payrollClassification, store });

    console.log(`\n  Distinct variant (2 columns differ): status=${result.status}`);
    console.log(`    structural similarity=${(structuralSimilarity * 100).toFixed(1)}%  (cutoff 80%, ${((structuralSimilarity - 0.8) * 100).toFixed(1)} pts ${structuralSimilarity >= 0.8 ? "OVER" : "under"})`);
    console.log(`    header similarity=${(headerSimilarity * 100).toFixed(1)}%  (cutoff 50%, ${((headerSimilarity - 0.5) * 100).toFixed(1)} pts ${headerSimilarity >= 0.5 ? "OVER" : "under"})`);
    assert(result.status === "needs_input", "Payroll List (distinct variant, 2 columns differ) is correctly NOT fuzzy-matched to Vendor's label");
  }

  // --- Case 3: the legitimate drift case, re-verified against the REAL fixture
  // used elsewhere in this suite (added one column, nothing else changed) ---
  // Confirms the header check doesn't collateral-damage the case it has to keep working.
  {
    const store = new JsonFileSheetMemoryStore({ filePath: null });
    const v1Signals = extractSheetSignals("Inventory Snapshot", buildInventoryValues(20));
    await rememberSheetLabel({ clientId, sheetName: "Inventory Snapshot", sheetSignals: v1Signals, userProvidedLabel: "Weekly inventory count", store });

    const v3Signals = extractSheetSignals("Stock Levels (New)", buildInventoryValues(28, { includeSupplierColumn: true }));
    const v1Sig = computeStructuralSignature(v1Signals);
    const v3Sig = computeStructuralSignature(v3Signals);
    const structuralSimilarity = tagSequenceSimilarity(v1Sig.tags, v3Sig.tags);
    const headerSimilarity = headerSequenceSimilarity(v1Sig.headerTokens, v3Sig.headerTokens);
    const result = await resolveSheetLabel({ clientId, sheetName: "Stock Levels (New)", sheetSignals: v3Signals, classification: classifySheet(v3Signals), store });

    console.log(`\n  Legitimate drift (1 column added, same sheet): status=${result.status}`);
    console.log(`    structural similarity=${(structuralSimilarity * 100).toFixed(1)}%   header similarity=${(headerSimilarity * 100).toFixed(1)}%`);
    assert(result.status === "remembered" && result.matchedVia === "fuzzy", "legitimate one-column-added drift still fuzzy-matches correctly with the header gate in place");
  }

  reportFinding(
    "Structural-tag similarity alone is unchanged and still can't tell these two cases apart on its own (both land at 83.3%) — " +
      "that math didn't change. What changed is that a fuzzy match now ALSO requires header-text similarity to clear 50%: the " +
      "legitimate drift case (headers unchanged except one insertion) scores ~83% there too, while the Vendor/Payroll collision " +
      "scores ~25% — different wording throughout, despite the coincidentally-matching shape. The two-signal AND-gate separates " +
      "them cleanly where either signal alone could not."
  );

  console.log("");
}

async function runReviewPanelEditAndReset() {
  console.log("-- Review panel: updateStoredSheetLabel() corrects a label in place, forgetSheetLabel() clears it --\n");
  const store = new JsonFileSheetMemoryStore({ filePath: null });
  const clientId = "workbook:review-panel-test.xlsx";
  // The real incident this feature exists for: a GST pivot table wrongly
  // labeled "Journal entries" (see the "PT 2A" mislabel report).
  const sheetSignals = extractSheetSignals("PT 2A", [
    ["GSTIN", "Taxable Value", "CGST", "SGST", "IGST"],
    ["29AAAPL2356Q1Z8", 10000, 900, 900, 0],
  ]);
  const { signature } = computeStructuralSignature(sheetSignals);

  await rememberSheetLabel({ clientId, sheetName: "PT 2A", sheetSignals, userProvidedLabel: "Journal entries", store });

  const updated = await updateStoredSheetLabel({ clientId, structuralSignature: signature, newLabel: "GST purchase register", store });
  assert(updated && updated.user_provided_label === "GST purchase register", "updateStoredSheetLabel() overwrites the stored label in place");

  const resolvedAfterEdit = await resolveSheetLabel({ clientId, sheetName: "PT 2A", sheetSignals, classification: { type: "unknown" }, store });
  assert(resolvedAfterEdit.status === "remembered" && resolvedAfterEdit.label === "GST purchase register", "the CORRECTED label is what the next scan resolves, not the original wrong one");

  const forgotten = await forgetSheetLabel({ clientId, structuralSignature: signature, store });
  assert(forgotten === true, "forgetSheetLabel() removes the record");

  const resolvedAfterForget = await resolveSheetLabel({ clientId, sheetName: "PT 2A", sheetSignals, classification: { type: "unknown" }, store });
  assert(resolvedAfterForget.status === "needs_input", "after forgetting, the next scan asks again instead of reusing anything stale");

  console.log("");
}

async function run() {
  await runSyntheticScenarios();
  await runRealDemoDataFlow();
  await runCrossTypeCollisionInvestigation();
  await runReviewPanelEditAndReset();

  if (findings.length > 0) {
    console.log(`\n${findings.length} note(s) from the cross-type collision investigation (see FINDING lines above).`);
  }

  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All sheet-memory checks passed.");
  }
}

run();
