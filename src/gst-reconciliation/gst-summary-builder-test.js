/*
 * gst-summary-builder-test.js
 * -------------------------------
 * Unit tests for gst-summary-builder.js, using only fictional synthetic
 * GSTINs and made-up figures. No real vendor or invoice data.
 *
 * Run with: node src/gst-reconciliation/gst-summary-builder-test.js
 */

const { buildGstSummary, emptyStepTwoTotals, emptyInvoiceLevelExtras } = require("./gst-summary-builder.js");

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

// A single invoice-level extra row, as extra-invoice-detector.js's
// detectInvoiceLevelExtras() produces one.
function extraRow(gstin, identifier, igst, cgst, sgst) {
  return { rowIndex: 0, gstin, identifier, taxableValue: 0, igst, cgst, sgst, placeOfSupply: null };
}

function runTests() {
  console.log("-- Matched vendor: fully supportable, nothing netted off --\n");

  const matchedOnly = {
    tolerance: 1,
    totalGstins: 1,
    counts: { matched: 1, discrepancy: 0, missing_in_books: 0, missing_in_gstr2a: 0 },
    vendors: [vendor("A", "matched", { igst: 1000 }, { igst: 1000 })],
  };
  const s1 = buildGstSummary(matchedOnly, emptyStepTwoTotals(), emptyInvoiceLevelExtras());
  assert(s1.totalTaxPer2A === 1000 && s1.totalTaxPerBooks === 1000, "totals reflect the single matched vendor on both sides");
  assert(s1.provisionallySupportable === 1000, "fully matched vendor -> the whole amount is supportable (min of equal figures)");
  assert(s1.netProvisionallyEligible === 1000, "no Step 2 findings -> net equals the supportable base, nothing subtracted");
  assert(s1.extraIn2A.count === 0 && s1.extraInBooks.count === 0, "no invoice-level extras supplied -> zero on both Extra lines");

  console.log("\n-- Extra in 2A / Extra in Books now come from invoice-level data, not vendor status --\n");

  const someVendorSummary = {
    tolerance: 1,
    totalGstins: 1,
    counts: { matched: 0, discrepancy: 0, missing_in_books: 1, missing_in_gstr2a: 0 },
    vendors: [vendor("B", "missing_in_books", { igst: 500 }, {})],
  };
  const invoiceExtras = { applicable: true, extraIn2A: [extraRow("B", "INV-1", 500, 0, 0)], extraInBooks: [] };
  const s2 = buildGstSummary(someVendorSummary, emptyStepTwoTotals(), invoiceExtras);
  assert(s2.provisionallySupportable === 0, "vendor present only in 2A -> min(500, 0) = 0, correctly excluded from the supportable base (still vendor-level, unaffected by this fix)");
  assert(s2.extraIn2A.count === 1 && s2.extraIn2A.amount === 500, "Extra in 2A count/amount now come from the supplied invoice-level rows, not the vendor status filter");

  const s2b = buildGstSummary(someVendorSummary, emptyStepTwoTotals(), emptyInvoiceLevelExtras());
  assert(s2b.extraIn2A.count === 0, "sanity check: with NO invoice-level extras supplied, Extra in 2A is zero even though a vendor has missing_in_books status -- proves the summary no longer derives this from vendor status at all");

  console.log("\n-- The exact bug this fix targets: a vendor Step 1 calls 'matched' can still have invoice-level extras --\n");

  const matchedVendorWithHiddenExtras = {
    tolerance: 1,
    totalGstins: 1,
    counts: { matched: 1, discrepancy: 0, missing_in_books: 0, missing_in_gstr2a: 0 },
    vendors: [vendor("C", "matched", { igst: 1000 }, { igst: 1000 })], // GSTIN-aggregate totals happen to match
  };
  // But 4 of this same vendor's 5 Books invoices never appear in 2A --
  // Step 1's GSTIN-level view can't see this at all.
  const hiddenExtras = {
    applicable: true,
    extraIn2A: [],
    extraInBooks: [extraRow("C", "INV-2", 0, 225, 225), extraRow("C", "INV-3", 0, 225, 225), extraRow("C", "INV-4", 0, 225, 225), extraRow("C", "INV-5", 0, 225, 225)],
  };
  const s3 = buildGstSummary(matchedVendorWithHiddenExtras, emptyStepTwoTotals(), hiddenExtras);
  assert(s3.vendorCounts.matched === 1 && s3.vendorCounts.missing_in_gstr2a === 0, "Step 1's vendor-level status still calls this vendor 'matched' -- that count is intentionally unchanged");
  assert(s3.extraInBooks.count === 4 && s3.extraInBooks.amount === 1800, "but the Extra in Books line correctly surfaces all 4 hidden invoices (900 CGST + 900 SGST = 1800), not zero");

  console.log("\n-- Discrepancy: supportable is capped at the LOWER side, gap reported as mismatch (still vendor-level, unchanged) --\n");

  const discrepancy = {
    tolerance: 1,
    totalGstins: 1,
    counts: { matched: 0, discrepancy: 1, missing_in_books: 0, missing_in_gstr2a: 0 },
    vendors: [vendor("D", "discrepancy", { igst: 900 }, { igst: 700 })],
  };
  const s4 = buildGstSummary(discrepancy, emptyStepTwoTotals(), emptyInvoiceLevelExtras());
  assert(s4.provisionallySupportable === 700, "2A shows 900 but Books only supports 700 -> capped at the lower, more conservative figure");
  assert(s4.mismatch.count === 1 && s4.mismatch.unsupportedAmount === 200, "the 200 gap is reported as mismatch, not silently absorbed either way -- Mismatch stays vendor-level, out of scope for this fix");

  console.log("\n-- Ineligible ITC and duplicate over-claims are subtracted from the base; Wrong Head/RCM are not --\n");

  const base = {
    tolerance: 1,
    totalGstins: 1,
    counts: { matched: 1, discrepancy: 0, missing_in_books: 0, missing_in_gstr2a: 0 },
    vendors: [vendor("E", "matched", { igst: 10000 }, { igst: 10000 })],
  };
  const stepTwo = {
    wrongHead: { rowCount: 3, crossSheetRowCount: 2 },
    rcm: { rowCount: 2, amount: 1500 },
    ineligibleItc: { rowCount: 1, amount: 800 },
    duplicates: { clusterCount: 1, overclaimAmount: 300 },
  };
  const s5 = buildGstSummary(base, stepTwo, emptyInvoiceLevelExtras());
  assert(s5.provisionallySupportable === 10000, "base supportable figure unaffected by Step 2 findings");
  assert(s5.lessIneligibleItc.amount === 800 && s5.lessDuplicateOverclaim.amount === 300, "ineligible ITC and duplicate over-claim amounts carried through as deductions");
  assert(s5.netProvisionallyEligible === 10000 - 800 - 300, "net = base - ineligible - duplicate overclaim, i.e. 8900");
  assert(s5.informational.wrongHeadRowCount === 3 && s5.informational.rcmRowCount === 2 && s5.informational.rcmAmount === 1500, "Wrong Head and RCM are carried through as informational counts/amounts only");
  assert(s5.informational.wrongHeadCrossSheetRowCount === 2, "the cross-sheet (2A vs Books, same invoice) Wrong Head count is carried through separately from the single-sheet count");
  assert(s5.netProvisionallyEligible !== 10000 - 800 - 300 - 1500, "sanity check: RCM's 1500 must NOT have been subtracted from the net figure");
  assert(s5.netProvisionallyEligible === 10000 - 800 - 300, "sanity check: the cross-sheet Wrong Head count must NOT have changed the net figure either -- informational only, like its single-sheet counterpart");

  console.log("\n-- Multiple vendors combine correctly; invoice-level extras combine independently of vendor status --\n");

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
  const mixedExtras = {
    applicable: true,
    extraIn2A: [extraRow("H", "H-1", 300, 0, 0)],
    extraInBooks: [extraRow("I", "I-1", 250, 0, 0)],
  };
  const s6 = buildGstSummary(mixed, emptyStepTwoTotals(), mixedExtras);
  assert(s6.totalTaxPer2A === 200 + 500 + 300 + 0, "2A total sums correctly across all four vendors");
  assert(s6.totalTaxPerBooks === 200 + 400 + 0 + 250, "Books total sums correctly across all four vendors");
  assert(s6.provisionallySupportable === 200 + 400 + 0 + 0, "supportable = matched(200) + discrepancy-lower(400) + extraIn2A(0) + extraInBooks(0) = 600");
  assert(s6.extraIn2A.amount === 300 && s6.extraInBooks.amount === 250, "extra-in-2A and extra-in-books amounts correctly isolated, sourced from the invoice-level rows supplied");

  console.log("\n-- Possible matches (Fix 4) are informational only -- never netted, never counted as Extra --\n");

  const withPossibleMatches = {
    tolerance: 1,
    totalGstins: 1,
    counts: { matched: 1, discrepancy: 0, missing_in_books: 0, missing_in_gstr2a: 0 },
    vendors: [vendor("J", "matched", { igst: 5000 }, { igst: 5000 })],
  };
  const possibleMatchExtras = {
    applicable: true,
    extraIn2A: [],
    extraInBooks: [],
    possibleMatches: [
      {
        gstin: "J",
        gstr2aIdentifier: "2018",
        gstr2a: { taxableValue: 165738.16, igst: 0, cgst: 14916.43, sgst: 14916.43 },
        booksIdentifier: "1997",
        books: { taxableValue: 165737.5, igst: 0, cgst: 14916.38, sgst: 14916.38 },
        reason: "Invoice numbers don't match ('2018' vs '1997')",
      },
    ],
  };
  const s7 = buildGstSummary(withPossibleMatches, emptyStepTwoTotals(), possibleMatchExtras);
  assert(s7.informational.possibleMatchCount === 1 && s7.informational.possibleMatchAmount === 29832.86, "possible match surfaced as its own informational count/amount (14916.43+14916.43 CGST+SGST, from the 2A side)");
  assert(s7.extraIn2A.count === 0 && s7.extraInBooks.count === 0, "NOT counted toward either Extra line -- the whole point of this being a separate, lower-confidence category");
  assert(s7.netProvisionallyEligible === 5000, "NOT netted into the supportable/net figure either -- purely informational, same treatment as Wrong Head and RCM");

  const s7b = buildGstSummary(withPossibleMatches, emptyStepTwoTotals(), emptyInvoiceLevelExtras());
  assert(s7b.informational.possibleMatchCount === 0 && s7b.informational.possibleMatchAmount === 0, "with no possibleMatches supplied at all, the count/amount default cleanly to zero (emptyInvoiceLevelExtras stays a safe default)");

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
