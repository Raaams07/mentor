/*
 * rate-mismatch-detector-test.js
 * ---------------------------------
 * Proves the pattern-agnostic sanity-check backstop: taxable value x
 * rate% should approximately equal the tax amount actually picked up.
 *
 * The headline scenario is the one this backstop exists for: a row where
 * the resolved tax column(s) hold figures that don't reconcile with the
 * rate — exactly the kind of future bug in this same family that column-
 * name/ambiguity checks alone can't predict, since the column name may
 * have matched with full confidence.
 *
 * All data is synthetic/fictional.
 *
 * Run with: node src/gst-reconciliation/rate-mismatch-detector-test.js
 */

const { detectRateMismatch, normalizeRatePercent } = require("./rate-mismatch-detector.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

// columns object shape matches identifyGstColumns()'s post-resolution
// output — only the indices this detector actually reads are populated.
function columnsWith(overrides) {
  return { gstin: 0, invoiceNumber: 1, taxableValue: 2, rate: 3, igst: null, cgst: 4, sgst: 5, ...overrides };
}

function runCorrectRowNotFlagged() {
  console.log("-- Correct row: not flagged --\n");
  const values = [["GSTIN", "Inv No", "Taxable Value", "Rate", "CGST", "SGST"], ["29AAAPL2356Q1Z8", "INV-1", 10000, 18, 900, 900]];
  const result = detectRateMismatch(values, 0, columnsWith({}));
  assert(result.applicable === true, "sheet has rate + taxableValue + tax columns resolved -> applicable");
  assert(result.flagged.length === 0, "10000 x 18% = 1800, CGST 900 + SGST 900 = 1800 -> matches, not flagged");
  console.log("");
}

function runWrongColumnPickStillCaught() {
  console.log("-- Sanity check catches a wrong-but-confidently-matched column pick (the whole point of this backstop) --\n");
  // Simulates the Tally bug's aftermath: columns.cgst header-matched with
  // full confidence (no ambiguity at this layer), but happens to hold
  // figures computed off a DIFFERENT rate than the row's actual Rate cell
  // -- naming/ambiguity checks alone have no way to catch this.
  const values = [["GSTIN", "Inv No", "Taxable Value", "Rate", "CGST", "SGST"], ["29AAAPL2356Q1Z8", "INV-1", 10000, 9, 700, 700]]; // CGST/SGST computed off 14%, not the stated 9%
  const result = detectRateMismatch(values, 0, columnsWith({}));
  assert(result.flagged.length === 1, "expected tax at 9% (900) doesn't match actual tax picked up (1400) -> flagged");
  const f = result.flagged[0];
  assert(f.expectedTotalTax === 900 && f.actualTotalTax === 1400 && f.difference === 500, "flagged row carries the exact expected/actual/difference figures for review");
  console.log("");
}

function runNormalizeRatePercent() {
  console.log("-- normalizeRatePercent: fraction vs whole-number handling --\n");
  assert(normalizeRatePercent(0.18) === 18, "Excel %-formatted fraction (0.18) -> 18");
  assert(normalizeRatePercent(18) === 18, "already a whole number (18) -> unchanged");
  assert(normalizeRatePercent(0) === 0, "zero -> zero");
  assert(normalizeRatePercent("18%") === 18, "string '18%' -> 18 (parsed, % stripped)");
  console.log("");
}

function runUnresolvedRateSkipsSilently() {
  console.log("-- Unresolved rate column: applicable=false, silent skip, not an error --\n");
  const values = [["GSTIN", "Inv No", "Taxable Value", "CGST", "SGST"], ["29AAAPL2356Q1Z8", "INV-1", 10000, 900, 900]];
  const result = detectRateMismatch(values, 0, columnsWith({ rate: null }));
  assert(result.applicable === false, "no resolved rate column -> not applicable");
  assert(result.flagged.length === 0, "no flags when not applicable");
  console.log("");
}

function runToleranceBoundary() {
  console.log("-- Tolerance: max(₹1, 1% of expected) --\n");
  // expectedTotalTax = 180000 * 100 / 100 ... use taxableValue=1000000, rate=18 -> expected=180000, tolerance = max(1, 1800) = 1800
  const withinTolerance = [["GSTIN", "Inv No", "Taxable Value", "Rate", "CGST", "SGST"], ["29AAAPL2356Q1Z8", "INV-1", 1000000, 18, 90025, 90025]]; // actual 180050, diff 50, within 1800
  const r1 = detectRateMismatch(withinTolerance, 0, columnsWith({}));
  assert(r1.flagged.length === 0, "difference of ₹50 against a ₹1800 tolerance -> not flagged");

  const beyondTolerance = [["GSTIN", "Inv No", "Taxable Value", "Rate", "CGST", "SGST"], ["29AAAPL2356Q1Z8", "INV-1", 1000000, 18, 92500, 92500]]; // actual 185000, diff 5000, beyond 1800
  const r2 = detectRateMismatch(beyondTolerance, 0, columnsWith({}));
  assert(r2.flagged.length === 1, "difference of ₹5000 against a ₹1800 tolerance -> flagged");
  console.log("");
}

function runNilRatedRowNotFlagged() {
  console.log("-- Nil-rated/exempt row (rate=0, taxable=0, tax=0): not flagged --\n");
  const values = [["GSTIN", "Inv No", "Taxable Value", "Rate", "CGST", "SGST"], ["29AAAPL2356Q1Z8", "INV-1", 0, 0, 0, 0]];
  const result = detectRateMismatch(values, 0, columnsWith({}));
  assert(result.flagged.length === 0, "a genuinely blank/nil row produces no flag");
  console.log("");
}

function run() {
  runCorrectRowNotFlagged();
  runWrongColumnPickStillCaught();
  runNormalizeRatePercent();
  runUnresolvedRateSkipsSilently();
  runToleranceBoundary();
  runNilRatedRowNotFlagged();

  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All rate-mismatch-detector checks passed.");
  }
}

run();
