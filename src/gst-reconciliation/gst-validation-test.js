/*
 * gst-validation-test.js
 * -------------------------
 * Proves runStandingValidators() is a faithful, source-agnostic wrapper
 * around the registered validators (currently just rate-mismatch) — same
 * results as calling the underlying detector directly, correct per-source
 * attribution across multiple sources, and the exact `...bySource` shape
 * existing callers (writeRateMismatchSheet, sumRateMismatchAmount,
 * buildStepTwoTotals in gst-report-writer.js) already depend on.
 *
 * All data is synthetic/fictional.
 *
 * Run with: node src/gst-reconciliation/gst-validation-test.js
 */

const { runStandingValidators, VALIDATORS } = require("./gst-validation.js");
const { detectRateMismatch } = require("./rate-mismatch-detector.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

function columnsWith(overrides) {
  return { gstin: 0, invoiceNumber: 1, taxableValue: 2, rate: 3, igst: null, cgst: 4, sgst: 5, ...overrides };
}

console.log("-- Exactly one validator registered today (rate-mismatch) --\n");
{
  assert(VALIDATORS.length === 1, "VALIDATORS currently has exactly one entry — update this test's count when a second is added");
  assert(VALIDATORS[0].name === "rateMismatch", "the one registered validator is named 'rateMismatch'");
}

console.log("\n-- Matches calling detectRateMismatch directly, for a clean row --\n");
{
  const values = [["GSTIN", "Inv No", "Taxable Value", "Rate", "CGST", "SGST"], ["29AAAPL2356Q1Z8", "INV-1", 10000, 18, 900, 900]];
  const source = { sheetName: "2A", role: "gstr2a", values, headerRowIndex: 0, columns: columnsWith({}) };

  const direct = detectRateMismatch(values, 0, columnsWith({}));
  const { rateMismatch } = runStandingValidators([source]);

  assert(rateMismatch.length === 1, "one source in, one result out");
  assert(JSON.stringify(rateMismatch[0].result) === JSON.stringify(direct), "result is byte-identical to calling detectRateMismatch directly — the module doesn't transform or reinterpret it");
  assert(rateMismatch[0].sheetName === "2A" && rateMismatch[0].role === "gstr2a", "sheetName/role are carried through onto the result object");
}

console.log("\n-- The real incident this backstop exists for: rate-tier column mistaken for the consolidated one --\n");
{
  // 10000 x 18% should be 1800, but the resolved CGST/SGST columns hold an
  // intermediate per-rate-tier figure (900 total instead of 1800) — the
  // exact shape of the historical Basai Steel Traders shortfall.
  const values = [["GSTIN", "Inv No", "Taxable Value", "Rate", "CGST", "SGST"], ["29AAAPL2356Q1Z8", "INV-1", 10000, 18, 450, 450]];
  const source = { sheetName: "Books", role: "purchase_register", values, headerRowIndex: 0, columns: columnsWith({}) };

  const { rateMismatch } = runStandingValidators([source]);
  assert(rateMismatch[0].result.applicable === true, "applicable — rate/taxableValue/tax columns all resolved");
  assert(rateMismatch[0].result.flagged.length === 1, "the understated total is flagged, not silently accepted");
  assert(rateMismatch[0].result.flagged[0].expectedTotalTax === 1800, "expected tax correctly computed from taxable value x rate%");
  assert(rateMismatch[0].result.flagged[0].actualTotalTax === 900, "actual tax correctly read off the (wrong) resolved columns");
}

console.log("\n-- Multiple sources: each gets its OWN result, never mixed up --\n");
{
  const cleanValues = [["GSTIN", "Inv No", "Taxable Value", "Rate", "CGST", "SGST"], ["29AAAPL2356Q1Z8", "INV-1", 10000, 18, 900, 900]];
  const flaggedValues = [["GSTIN", "Inv No", "Taxable Value", "Rate", "CGST", "SGST"], ["29AAAPL2356Q1Z8", "INV-2", 10000, 18, 450, 450]];

  const { rateMismatch } = runStandingValidators([
    { sheetName: "2A", role: "gstr2a", values: cleanValues, headerRowIndex: 0, columns: columnsWith({}) },
    { sheetName: "Books", role: "purchase_register", values: flaggedValues, headerRowIndex: 0, columns: columnsWith({}) },
  ]);

  assert(rateMismatch.length === 2, "two sources in, two results out, in the same order");
  assert(rateMismatch[0].sheetName === "2A" && rateMismatch[0].result.flagged.length === 0, "2A's own clean result stays clean");
  assert(rateMismatch[1].sheetName === "Books" && rateMismatch[1].result.flagged.length === 1, "Books' own flagged result stays flagged — not swapped or merged with 2A's");
}

console.log("\n-- A source with no resolved Rate column: not applicable, not an error --\n");
{
  const values = [["GSTIN", "Inv No", "Taxable Value", "CGST", "SGST"], ["29AAAPL2356Q1Z8", "INV-1", 10000, 900, 900]];
  const source = { sheetName: "2A", role: "gstr2a", values, headerRowIndex: 0, columns: columnsWith({ rate: null }) };

  const { rateMismatch } = runStandingValidators([source]);
  assert(rateMismatch[0].result.applicable === false, "no Rate column resolved -> not applicable, matches detectRateMismatch's own silent-skip behavior");
}

console.log("");
if (failures > 0) {
  console.log(`${failures} check(s) FAILED.`);
  process.exitCode = 1;
} else {
  console.log("All gst-validation checks passed.");
}
