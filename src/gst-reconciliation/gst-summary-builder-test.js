/*
 * gst-summary-builder-test.js
 * -------------------------------
 * Unit tests for gst-summary-builder.js, using only fictional synthetic
 * GSTINs and made-up figures. No real vendor or invoice data.
 *
 * Run with: node src/gst-reconciliation/gst-summary-builder-test.js
 */

const { buildGstSummary, emptyStepTwoTotals } = require("./gst-summary-builder.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

function vendor(gstin, status, gstr2a, books) {
  return {
    gstin,
    status,
    gstr2a: { taxableValue: 0, igst: 0, cgst: 0, sgst: 0, rowCount: 1, ...gstr2a },
    books: { taxableValue: 0, igst: 0, cgst: 0, sgst: 0, rowCount: 1, ...books },
  };
}

function runTests() {
  console.log("-- Matched vendor: fully supportable, nothing netted off --\n");

  const matchedOnly = {
    tolerance: 1,
    totalGstins: 1,
    counts: { matched: 1, discrepancy: 0, missing_in_books: 0, missing_in_gstr2a: 0 },
    vendors: [vendor("A", "matched", { igst: 1000 }, { igst: 1000 })],
  };
  const s1 = buildGstSummary(matchedOnly, emptyStepTwoTotals());
  assert(s1.totalTaxPer2A === 1000 && s1.totalTaxPerBooks === 1000, "totals reflect the single matched vendor on both sides");
  assert(s1.provisionallySupportable === 1000, "fully matched vendor -> the whole amount is supportable (min of equal figures)");
  assert(s1.netProvisionallyEligible === 1000, "no Step 2 findings -> net equals the supportable base, nothing subtracted");

  console.log("\n-- Extra in 2A (missing_in_books): NOT counted as supportable, reported separately --\n");

  const extraIn2A = {
    tolerance: 1,
    totalGstins: 1,
    counts: { matched: 0, discrepancy: 0, missing_in_books: 1, missing_in_gstr2a: 0 },
    vendors: [vendor("B", "missing_in_books", { igst: 500 }, {})],
  };
  const s2 = buildGstSummary(extraIn2A, emptyStepTwoTotals());
  assert(s2.provisionallySupportable === 0, "vendor present only in 2A -> min(500, 0) = 0, correctly excluded from the supportable base");
  assert(s2.extraIn2A.count === 1 && s2.extraIn2A.amount === 500, "...but still reported as its own 'Extra in 2A' line, not silently dropped");

  console.log("\n-- Extra in Books (missing_in_gstr2a): same treatment, mirrored --\n");

  const extraInBooks = {
    tolerance: 1,
    totalGstins: 1,
    counts: { matched: 0, discrepancy: 0, missing_in_books: 0, missing_in_gstr2a: 1 },
    vendors: [vendor("C", "missing_in_gstr2a", {}, { cgst: 200, sgst: 200 })],
  };
  const s3 = buildGstSummary(extraInBooks, emptyStepTwoTotals());
  assert(s3.provisionallySupportable === 0, "vendor present only in Books -> min(0, 400) = 0, correctly excluded");
  assert(s3.extraInBooks.count === 1 && s3.extraInBooks.amount === 400, "...reported as its own 'Extra in Books' line");

  console.log("\n-- Discrepancy: supportable is capped at the LOWER side, gap reported as mismatch --\n");

  const discrepancy = {
    tolerance: 1,
    totalGstins: 1,
    counts: { matched: 0, discrepancy: 1, missing_in_books: 0, missing_in_gstr2a: 0 },
    vendors: [vendor("D", "discrepancy", { igst: 900 }, { igst: 700 })],
  };
  const s4 = buildGstSummary(discrepancy, emptyStepTwoTotals());
  assert(s4.provisionallySupportable === 700, "2A shows 900 but Books only supports 700 -> capped at the lower, more conservative figure");
  assert(s4.mismatch.count === 1 && s4.mismatch.unsupportedAmount === 200, "the 200 gap is reported as mismatch, not silently absorbed either way");

  console.log("\n-- Ineligible ITC and duplicate over-claims are subtracted from the base; Wrong Head/RCM are not --\n");

  const base = {
    tolerance: 1,
    totalGstins: 1,
    counts: { matched: 1, discrepancy: 0, missing_in_books: 0, missing_in_gstr2a: 0 },
    vendors: [vendor("E", "matched", { igst: 10000 }, { igst: 10000 })],
  };
  const stepTwo = {
    wrongHead: { rowCount: 3 },
    rcm: { rowCount: 2, amount: 1500 },
    ineligibleItc: { rowCount: 1, amount: 800 },
    duplicates: { clusterCount: 1, overclaimAmount: 300 },
  };
  const s5 = buildGstSummary(base, stepTwo);
  assert(s5.provisionallySupportable === 10000, "base supportable figure unaffected by Step 2 findings");
  assert(s5.lessIneligibleItc.amount === 800 && s5.lessDuplicateOverclaim.amount === 300, "ineligible ITC and duplicate over-claim amounts carried through as deductions");
  assert(s5.netProvisionallyEligible === 10000 - 800 - 300, "net = base - ineligible - duplicate overclaim, i.e. 8900");
  assert(s5.informational.wrongHeadRowCount === 3 && s5.informational.rcmRowCount === 2 && s5.informational.rcmAmount === 1500, "Wrong Head and RCM are carried through as informational counts/amounts only");
  assert(s5.netProvisionallyEligible !== 10000 - 800 - 300 - 1500, "sanity check: RCM's 1500 must NOT have been subtracted from the net figure");

  console.log("\n-- Multiple vendors combine correctly across all four statuses at once --\n");

  const mixed = {
    tolerance: 1,
    totalGstins: 4,
    counts: { matched: 1, discrepancy: 1, missing_in_books: 1, missing_in_gstr2a: 1 },
    vendors: [
      vendor("F", "matched", { cgst: 100, sgst: 100 }, { cgst: 100, sgst: 100 }),
      vendor("G", "discrepancy", { igst: 500 }, { igst: 400 }),
      vendor("H", "missing_in_books", { igst: 300 }, {}),
      vendor("I", "missing_in_gstr2a", {}, { igst: 250 }),
    ],
  };
  const s6 = buildGstSummary(mixed, emptyStepTwoTotals());
  assert(s6.totalTaxPer2A === 200 + 500 + 300 + 0, "2A total sums correctly across all four vendors");
  assert(s6.totalTaxPerBooks === 200 + 400 + 0 + 250, "Books total sums correctly across all four vendors");
  assert(s6.provisionallySupportable === 200 + 400 + 0 + 0, "supportable = matched(200) + discrepancy-lower(400) + extraIn2A(0) + extraInBooks(0) = 600");
  assert(s6.extraIn2A.amount === 300 && s6.extraInBooks.amount === 250, "extra-in-2A and extra-in-books amounts correctly isolated to their own vendors");

  console.log("");
}

function run() {
  runTests();
  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All GST summary builder checks passed.");
  }
}

run();
