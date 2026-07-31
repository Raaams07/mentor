/*
 * wrong-head-detector-test.js
 * ------------------------------
 * Unit tests for wrong-head-detector.js, using only fictional synthetic
 * GSTINs (constructed to satisfy the legal GSTIN format — see gstin.js)
 * and made-up place-of-supply values. No real vendor or invoice data.
 *
 * Run with: node src/gst-reconciliation/wrong-head-detector-test.js
 */

const { detectWrongHead } = require("./wrong-head-detector.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

// Fictional GSTINs, one per state used in these tests.
const GSTIN_MAHARASHTRA = "27ZZAPL5432Q1Z9"; // state code 27
const GSTIN_DELHI = "07ZZAPL7788Q1ZO"; // state code 07
const GSTIN_KARNATAKA = "29ZZAPL1122Q1ZI"; // state code 29

// columns index layout used throughout: [GSTIN, Invoice No, Place of Supply, IGST, CGST, SGST]
const COLUMNS = { gstin: 0, invoiceNumber: 1, placeOfSupply: 2, igst: 3, cgst: 4, sgst: 5, taxableValue: null };
const HEADER_ROW = ["GSTIN", "Invoice No", "Place of Supply", "IGST", "CGST", "SGST"];

function run(rows) {
  return detectWrongHead([HEADER_ROW, ...rows], 0, COLUMNS);
}

function runTests() {
  console.log("-- Correct cases (should NOT be flagged) --\n");

  const correctIntraState = run([[GSTIN_MAHARASHTRA, "INV-1", "Maharashtra", 0, 900, 900]]);
  assert(correctIntraState.flagged.length === 0, "same-state supplier + same-state Place of Supply, charged CGST+SGST -> not flagged");

  const correctInterState = run([[GSTIN_MAHARASHTRA, "INV-2", "Karnataka", 1800, 0, 0]]);
  assert(correctInterState.flagged.length === 0, "different-state supplier + Place of Supply, charged IGST -> not flagged");

  const codePrefixedPos = run([[GSTIN_DELHI, "INV-3", "07-Delhi", 0, 500, 500]]);
  assert(codePrefixedPos.flagged.length === 0, "'<code>-<name>' style Place of Supply is correctly parsed");

  const nilRated = run([[GSTIN_KARNATAKA, "INV-4", "Karnataka", 0, 0, 0]]);
  assert(nilRated.flagged.length === 0, "no tax charged at all (nil-rated/exempt) is not flagged — nothing to judge");

  console.log("\n-- Wrong Head cases (SHOULD be flagged) --\n");

  const wrongHeadIgstChargedForIntraState = run([[GSTIN_MAHARASHTRA, "INV-5", "Maharashtra", 1800, 0, 0]]);
  assert(wrongHeadIgstChargedForIntraState.flagged.length === 1, "same-state supplier/PoS but IGST charged -> flagged");
  assert(wrongHeadIgstChargedForIntraState.flagged[0].issue === "wrong_head" && wrongHeadIgstChargedForIntraState.flagged[0].expected === "cgst_sgst" && wrongHeadIgstChargedForIntraState.flagged[0].actual === "igst", "...with the correct expected/actual heads recorded");

  const wrongHeadCgstSgstChargedForInterState = run([[GSTIN_MAHARASHTRA, "INV-6", "Karnataka", 0, 900, 900]]);
  assert(wrongHeadCgstSgstChargedForInterState.flagged.length === 1, "different-state supplier/PoS but CGST+SGST charged -> flagged");
  assert(wrongHeadCgstSgstChargedForInterState.flagged[0].expected === "igst" && wrongHeadCgstSgstChargedForInterState.flagged[0].actual === "cgst_sgst", "...with the correct expected/actual heads recorded");

  const mixedHeads = run([[GSTIN_MAHARASHTRA, "INV-7", "Maharashtra", 900, 450, 450]]);
  assert(mixedHeads.flagged.length === 1 && mixedHeads.flagged[0].issue === "mixed_tax_heads", "both IGST and CGST/SGST populated on one line -> flagged as mixed_tax_heads, not wrong_head");

  console.log("\n-- Multiple rows, mixed correctness --\n");

  const multiRow = run([
    [GSTIN_MAHARASHTRA, "INV-8", "Maharashtra", 0, 900, 900], // correct
    [GSTIN_MAHARASHTRA, "INV-9", "Karnataka", 0, 900, 900], // wrong (should be IGST)
    [GSTIN_DELHI, "INV-10", "Delhi", 0, 500, 500], // correct
  ]);
  assert(multiRow.flagged.length === 1 && multiRow.flagged[0].invoiceNumber === "INV-9", "only the genuinely wrong row is flagged among several correct ones");

  console.log("\n-- Sheets without the needed columns --\n");

  const noPlaceOfSupply = detectWrongHead([HEADER_ROW, [GSTIN_MAHARASHTRA, "INV-11", "Maharashtra", 0, 900, 900]], 0, { ...COLUMNS, placeOfSupply: null });
  assert(noPlaceOfSupply.applicable === false, "sheet without a Place of Supply column is correctly reported as not applicable, not force-evaluated");

  console.log("");
}

function run_() {
  runTests();
  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All Wrong Head detector checks passed.");
  }
}

run_();
