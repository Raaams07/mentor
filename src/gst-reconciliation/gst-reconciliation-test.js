/*
 * gst-reconciliation-test.js
 * ----------------------------
 * Proves the GSTR-2A vs Books reconciliation pipeline end-to-end:
 *   1. Shape-based role recognition tells a GSTR-2A-style sheet apart from
 *      a purchase-register-style sheet, and correctly says "unknown" for
 *      a sheet that's neither.
 *   2. Workflow identification only fires once BOTH required roles are
 *      present — not on either sheet alone.
 *   3. Aggregation-by-GSTIN and comparison correctly produce all four
 *      outcomes: matched, discrepancy, missing-in-Books, missing-in-2A.
 *   4. The whole thing is read-only — running it doesn't write anything
 *      anywhere; the summary is a plain object for a UI to render.
 *
 * All data below is synthetic and fictional — made-up vendor names and
 * GSTINs constructed purely to satisfy the legal GSTIN format (see
 * gstin.js), not sourced from any real company or real test file. This
 * script exists to validate the shape-based rules, not to encode anything
 * learned from a specific dataset.
 *
 * Run with: node src/gst-reconciliation/gst-reconciliation-test.js
 */

const { recognizeGstRole } = require("./gst-role-recognizer.js");
const { identifyGstWorkflow } = require("./gst-workflows.js");
const { runGstReconciliation, recognizeGstSheets } = require("./gst-reconciliation.js");
const { extractSheetSignals } = require("../sheet-classifier/signal-extractor.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

// Fictional, structurally-valid GSTINs — constructed to satisfy the legal
// 15-character format, not copied from any real company.
const GSTIN_ALPHA = "29AAAPL2356Q1Z8"; // "Vendor Alpha Pvt Ltd" — matches exactly
const GSTIN_BETA = "27AAACT2727Q1ZW"; // "Vendor Beta Enterprises" — has a real discrepancy
const GSTIN_GAMMA = "07AABCU9603R1ZP"; // "Vendor Gamma Traders" — in 2A only
const GSTIN_DELTA = "33AAAAA0000A1Z9"; // "Vendor Delta Co" — in Books only

function buildGstr2aValues() {
  return [
    ["GSTIN of Supplier", "Trade/Legal Name", "Invoice Number", "Invoice Date", "Invoice Value", "Taxable Value", "Rate", "Integrated Tax", "Central Tax", "State/UT Tax", "Place of Supply"],
    [GSTIN_ALPHA, "Vendor Alpha Pvt Ltd", "INV-001", "2026-05-01", 11800, 10000, 18, 0, 900, 900, "Karnataka"],
    [GSTIN_BETA, "Vendor Beta Enterprises", "INV-002", "2026-05-03", 23600, 20000, 18, 3600, 0, 0, "Maharashtra"],
    [GSTIN_BETA, "Vendor Beta Enterprises", "INV-003", "2026-05-10", 5900, 5000, 18, 900, 0, 0, "Maharashtra"],
    [GSTIN_GAMMA, "Vendor Gamma Traders", "INV-004", "2026-05-15", 17700, 15000, 18, 2700, 0, 0, "Delhi"],
  ];
}

function buildBooksValues() {
  return [
    ["Date", "Particulars", "Voucher Number", "GSTIN", "Taxable Value", "IGST", "CGST", "SGST"],
    ["2026-05-01", "Vendor Alpha Pvt Ltd", "PV-101", GSTIN_ALPHA, 10000, 0, 900, 900],
    ["2026-05-03", "Vendor Beta Enterprises", "PV-102", GSTIN_BETA, 20000, 3600, 0, 0],
    ["2026-05-10", "Vendor Beta Enterprises", "PV-103", GSTIN_BETA, 4500, 810, 0, 0], // deliberately understated vs 2A
    ["2026-05-20", "Vendor Delta Co", "PV-104", GSTIN_DELTA, 8000, 0, 720, 720],
  ];
}

// A plain, non-GST table — reused to confirm the recognizer says "unknown"
// rather than force-fitting an unrelated sheet into one of the two roles.
function buildUnrelatedValues() {
  return [
    ["Employee ID", "Employee Name", "Department", "Salary"],
    ["E-001", "Employee 1", "Sales", 45000],
    ["E-002", "Employee 2", "Marketing", 52000],
  ];
}

function runRoleRecognitionTests() {
  console.log("-- Role recognition --\n");

  const gstr2aSignals = extractSheetSignals("2A", buildGstr2aValues());
  const gstr2aResult = recognizeGstRole(gstr2aSignals, buildGstr2aValues(), gstr2aSignals.headerRowIndex);
  console.log("  2A sheet ->", gstr2aResult.role, `(confidence: ${gstr2aResult.confidence}, scores: ${JSON.stringify(gstr2aResult.scores)})`);
  assert(gstr2aResult.role === "gstr2a", "GSTR-2A-shaped sheet recognized as 'gstr2a'");
  assert(gstr2aResult.columns.gstin !== null, "GSTIN column identified by content format");
  assert(gstr2aResult.columns.igst !== null && gstr2aResult.columns.cgst !== null && gstr2aResult.columns.sgst !== null, "all three tax columns identified by header vocabulary");

  const booksSignals = extractSheetSignals("Books", buildBooksValues());
  const booksResult = recognizeGstRole(booksSignals, buildBooksValues(), booksSignals.headerRowIndex);
  console.log("  Books sheet ->", booksResult.role, `(confidence: ${booksResult.confidence}, scores: ${JSON.stringify(booksResult.scores)})`);
  assert(booksResult.role === "purchase_register", "purchase-register-shaped sheet recognized as 'purchase_register'");
  assert(booksResult.columns.voucherNumber !== null, "Voucher Number column identified");
  assert(booksResult.columns.particulars !== null, "Particulars column identified");

  const unrelatedSignals = extractSheetSignals("Payroll", buildUnrelatedValues());
  const unrelatedResult = recognizeGstRole(unrelatedSignals, buildUnrelatedValues(), unrelatedSignals.headerRowIndex);
  console.log("  Unrelated sheet ->", unrelatedResult.role, `(reason: ${unrelatedResult.reason})`);
  assert(unrelatedResult.role === "unknown", "a non-GST sheet is NOT force-fit into either role");

  console.log("");
}

function runWorkflowIdentificationTests() {
  console.log("-- Workflow identification --\n");

  const gstr2aSignals = extractSheetSignals("2A", buildGstr2aValues());
  const gstr2aResult = recognizeGstRole(gstr2aSignals, buildGstr2aValues(), gstr2aSignals.headerRowIndex);

  const onlyOneSheet = identifyGstWorkflow({ "2A": gstr2aResult });
  assert(onlyOneSheet === null, "workflow is NOT identified when only one required role is present");

  const booksSignals = extractSheetSignals("Books", buildBooksValues());
  const booksResult = recognizeGstRole(booksSignals, buildBooksValues(), booksSignals.headerRowIndex);

  const bothSheets = identifyGstWorkflow({ "2A": gstr2aResult, Books: booksResult });
  assert(bothSheets !== null && bothSheets.workflowId === "gstr2a_vs_books", "workflow IS identified once both roles are present, regardless of sheet names");
  assert(bothSheets.sheets.gstr2a.sheetName === "2A" && bothSheets.sheets.purchaseRegister.sheetName === "Books", "workflow correctly maps each role to the right sheet name");

  console.log("");
}

function runReconciliationTests() {
  console.log("-- Aggregation, comparison, and the proposed summary --\n");

  const sheets = [
    { name: "2A", values: buildGstr2aValues() },
    { name: "Books", values: buildBooksValues() },
  ];

  const result = runGstReconciliation(sheets);
  assert(result.status === "reconciled", "reconciliation runs end-to-end from raw sheet values");
  assert(result.workflowId === "gstr2a_vs_books", "the correct workflow is reported");
  assert(result.summary.proposedAction === "review", "output is explicitly a proposal, not an auto-applied action");

  const byGstin = {};
  result.summary.vendors.forEach((v) => (byGstin[v.gstin] = v));

  console.log("  Summary counts:", JSON.stringify(result.summary.counts));

  const alpha = byGstin[GSTIN_ALPHA];
  assert(alpha && alpha.status === "matched", "Vendor Alpha (identical in both sheets) is reported as matched");

  const beta = byGstin[GSTIN_BETA];
  assert(beta && beta.status === "discrepancy", "Vendor Beta (understated in Books) is reported as a discrepancy");
  assert(beta.differences.taxableValue === 500 && beta.differences.igst === 90, "Vendor Beta's discrepancy amounts are computed correctly (2A minus Books)");

  const gamma = byGstin[GSTIN_GAMMA];
  assert(gamma && gamma.status === "missing_in_books", "Vendor Gamma (in 2A only) is reported as missing in Books");

  const delta = byGstin[GSTIN_DELTA];
  assert(delta && delta.status === "missing_in_gstr2a", "Vendor Delta (in Books only) is reported as missing in 2A");

  // Rounding-tolerance check: a sub-rupee difference should still count as matched.
  const almostMatchedSheets = [
    { name: "2A", values: [buildGstr2aValues()[0], [GSTIN_ALPHA, "Vendor Alpha Pvt Ltd", "INV-001", "2026-05-01", 11800.4, 10000.4, 18, 0, 900, 900, "Karnataka"]] },
    { name: "Books", values: [buildBooksValues()[0], ["2026-05-01", "Vendor Alpha Pvt Ltd", "PV-101", GSTIN_ALPHA, 10000, 0, 900, 900]] },
  ];
  const almostMatchedResult = runGstReconciliation(almostMatchedSheets);
  const alphaAlmost = almostMatchedResult.summary.vendors.find((v) => v.gstin === GSTIN_ALPHA);
  assert(alphaAlmost.status === "matched", "a sub-₹1 difference still counts as matched (rounding tolerance)");

  // Taxable Value is NOT a match criterion (validated against a real expert
  // reconciliation — see gstr2a-vs-books-reconciler.js) — a large taxable
  // value gap with tax fields matching exactly must still be "matched".
  const taxableValueOnlyMismatchSheets = [
    { name: "2A", values: [buildGstr2aValues()[0], [GSTIN_ALPHA, "Vendor Alpha Pvt Ltd", "INV-001", "2026-05-01", 23600, 20000, 18, 0, 900, 900, "Karnataka"]] },
    { name: "Books", values: [buildBooksValues()[0], ["2026-05-01", "Vendor Alpha Pvt Ltd", "PV-101", GSTIN_ALPHA, 10000, 0, 900, 900]] },
  ];
  const taxableValueOnlyResult = runGstReconciliation(taxableValueOnlyMismatchSheets);
  const alphaTaxableOnly = taxableValueOnlyResult.summary.vendors.find((v) => v.gstin === GSTIN_ALPHA);
  assert(alphaTaxableOnly.differences.taxableValue === 10000, "a large taxable-value gap is still computed and reported...");
  assert(alphaTaxableOnly.status === "matched", "...but does NOT by itself cause a 'discrepancy' status when tax fields match");

  console.log("");
}

function run() {
  runRoleRecognitionTests();
  runWorkflowIdentificationTests();
  runReconciliationTests();

  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All GST reconciliation checks passed.");
  }
}

run();
