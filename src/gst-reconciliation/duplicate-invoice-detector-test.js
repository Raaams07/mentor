/*
 * duplicate-invoice-detector-test.js
 * --------------------------------------
 * Unit tests for duplicate-invoice-detector.js, using only fictional
 * synthetic GSTINs and made-up invoice data. No real vendor or invoice
 * data.
 *
 * Run with: node src/gst-reconciliation/duplicate-invoice-detector-test.js
 */

const { findDuplicateInvoices } = require("./duplicate-invoice-detector.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

const GSTIN_A = "27ZZAPL5432Q1Z9";
const GSTIN_B = "07ZZAPL7788Q1ZO";

// columns index layout: [GSTIN, Date, Invoice No, Taxable Value]
const COLUMNS = { gstin: 0, invoiceNumber: 2, voucherNumber: null, taxableValue: 3, dateColumns: [1] };
const HEADER_ROW = ["GSTIN", "Date", "Invoice No", "Taxable Value"];

function run(rows, options) {
  return findDuplicateInvoices([HEADER_ROW, ...rows], 0, COLUMNS, options);
}

function runTests() {
  console.log("-- Should be flagged --\n");

  const sameInvoiceNumberClose = run([
    [GSTIN_A, "2026-05-01", "INV-100", 10000],
    [GSTIN_A, "2026-05-03", "INV-100", 10000],
  ]);
  assert(sameInvoiceNumberClose.clusters.length === 1, "same GSTIN, same invoice number, 2 days apart -> one cluster");
  assert(sameInvoiceNumberClose.clusters[0].matchReason === "same_invoice_number_and_amount", "...matched on both number and amount");

  const sameAmountDifferentNumber = run([
    [GSTIN_A, "2026-05-01", "INV-200", 5000],
    [GSTIN_A, "2026-05-04", "INV-201", 5000],
  ]);
  assert(sameAmountDifferentNumber.clusters.length === 1, "same GSTIN, same amount, different invoice numbers, 3 days apart -> one cluster");
  assert(sameAmountDifferentNumber.clusters[0].matchReason === "same_amount", "...matched on amount only");

  const threeEntriesOneCluster = run([
    [GSTIN_A, "2026-05-01", "INV-300", 8000],
    [GSTIN_A, "2026-05-02", "INV-300", 8000],
    [GSTIN_A, "2026-05-05", "INV-300", 8000],
  ]);
  assert(threeEntriesOneCluster.clusters.length === 1, "the same invoice entered three times forms ONE cluster, not three overlapping pairs");
  assert(threeEntriesOneCluster.clusters[0].members.length === 3, "...containing all three entries");

  const monthlyReentryPattern = run([
    [GSTIN_A, "2026-05-01", "INV-999", 45000],
    [GSTIN_A, "2026-05-31", "INV-999", 45000], // 30 days apart — a bookkeeping-close-cycle re-entry, the real pattern this rule was widened for
  ]);
  assert(monthlyReentryPattern.clusters.length === 1, "same invoice number, 30 days apart (a monthly re-entry pattern, validated against real data) -> flagged under the widened 45-day identifier window");

  console.log("\n-- Should NOT be flagged --\n");

  const sameInvoiceNumberFarApart = run([
    [GSTIN_A, "2026-01-01", "INV-400", 10000],
    [GSTIN_A, "2026-08-01", "INV-400", 10000],
  ]);
  assert(sameInvoiceNumberFarApart.clusters.length === 0, "same invoice number, same amount, but ~7 months (212 days) apart -> NOT flagged — well outside even the widened 45-day identifier window (likely a distinct re-supply or an invoice-numbering reset, not a duplicate entry)");

  const recurringMonthlyCharge = run([
    [GSTIN_A, "2026-05-01", "RENT-MAY", 20000],
    [GSTIN_A, "2026-06-01", "RENT-JUN", 20000],
  ]);
  assert(recurringMonthlyCharge.clusters.length === 0, "two consecutive months' rent — same amount, ~31 days apart, DIFFERENT invoice numbers — is NOT flagged: amount-only matching uses a tight 3-day window, not the wider identifier window");

  console.log("\n-- Regression: the specific over-flagging bug this file was fixed for --\n");

  const manySameAmountAcrossTheYear = run([
    [GSTIN_A, "2026-01-15", "BF-001", 500],
    [GSTIN_A, "2026-02-15", "BF-002", 500],
    [GSTIN_A, "2026-03-15", "BF-003", 500],
    [GSTIN_A, "2026-04-15", "BF-004", 500],
    [GSTIN_A, "2026-05-15", "BF-005", 500],
    [GSTIN_A, "2026-06-15", "BF-006", 500],
    [GSTIN_A, "2026-07-15", "BF-007", 500],
    [GSTIN_A, "2026-08-15", "BF-008", 500],
    [GSTIN_A, "2026-09-15", "BF-009", 500],
  ]);
  assert(
    manySameAmountAcrossTheYear.clusters.length === 0,
    "a vendor with a legitimately-recurring identical amount (e.g. a standard ₹500 monthly bank/service fee), each occurrence ~30 days from its neighbor and different invoice numbers -> NOT flagged, even though the whole SET shares one amount"
  );

  const chainedButShouldNotMergeIntoOneWideCluster = run([
    [GSTIN_A, "2026-01-01", "C-001", 339],
    [GSTIN_A, "2026-01-03", "C-002", 339], // 2 days from the previous -- within the tight amount-only window
    [GSTIN_A, "2026-01-05", "C-003", 339], // 2 days from the previous
    [GSTIN_A, "2026-01-07", "C-004", 339], // 2 days from the previous -- but 6 days from the FIRST entry
    [GSTIN_A, "2026-01-09", "C-005", 339], // 2 days from the previous -- but 8 days from the first entry
  ]);
  const totalMembersClustered = chainedButShouldNotMergeIntoOneWideCluster.clusters.reduce((sum, c) => sum + c.members.length, 0);
  const maxClusterSpread = chainedButShouldNotMergeIntoOneWideCluster.clusters.reduce((max, c) => Math.max(max, c.dateSpreadDays), 0);
  assert(
    !chainedButShouldNotMergeIntoOneWideCluster.clusters.some((c) => c.members.length === 5),
    "five same-amount entries, each 2 days from its immediate neighbor but spanning 8 days total, do NOT get merged into one 5-member cluster via chaining"
  );
  assert(maxClusterSpread <= 3, "no resulting cluster's own date spread exceeds the 3-day amount-only window, even though the whole chain spans 8 days");
  console.log("  (this scenario produced " + chainedButShouldNotMergeIntoOneWideCluster.clusters.length + " cluster(s) covering " + totalMembersClustered + " of 5 entries, each individually within the tight window — not one 8-day-wide cluster)");

  const differentGstins = run([
    [GSTIN_A, "2026-05-01", "INV-500", 7000],
    [GSTIN_B, "2026-05-02", "INV-500", 7000],
  ]);
  assert(differentGstins.clusters.length === 0, "same invoice number/amount but DIFFERENT GSTINs -> not flagged (grouping is per-GSTIN)");

  const genuinelyDifferent = run([
    [GSTIN_A, "2026-05-01", "INV-600", 3000],
    [GSTIN_A, "2026-05-02", "INV-601", 4500],
  ]);
  assert(genuinelyDifferent.clusters.length === 0, "different invoice numbers AND different amounts -> not flagged");

  console.log("\n-- Configurability --\n");

  const widerIdentifierWindowCatchesIt = run(
    [
      [GSTIN_A, "2026-01-01", "INV-700", 10000],
      [GSTIN_A, "2026-03-01", "INV-700", 10000], // 59 days apart -- exceeds even the widened 45-day default
    ],
    { identifierWindowDays: 60 }
  );
  assert(widerIdentifierWindowCatchesIt.clusters.length === 1, "a caller-supplied wider identifier window (60 days) catches a pair the 45-day default would have missed at 59 days apart");

  const tighterAmountOnlyWindowMissesIt = run(
    [
      [GSTIN_A, "2026-05-01", "INV-800", 9000],
      [GSTIN_A, "2026-05-03", "INV-801", 9000], // 2 days apart -- caught by the 3-day default, but not by a 1-day override
    ],
    { amountOnlyWindowDays: 1 }
  );
  assert(tighterAmountOnlyWindowMissesIt.clusters.length === 0, "a caller-supplied tighter amount-only window (1 day) correctly misses a pair 2 days apart that the 3-day default would catch");

  console.log("\n-- Sheets without the needed columns --\n");

  const noGstin = findDuplicateInvoices([HEADER_ROW, [GSTIN_A, "2026-05-01", "INV-800", 1000]], 0, { ...COLUMNS, gstin: null });
  assert(noGstin.applicable === false, "sheet without a GSTIN column is correctly reported as not applicable");

  const noDate = findDuplicateInvoices([HEADER_ROW, [GSTIN_A, "2026-05-01", "INV-900", 1000]], 0, { ...COLUMNS, dateColumns: [] });
  assert(noDate.applicable === false, "sheet without a date column is correctly reported as not applicable — can't apply a proximity window without dates");

  console.log("");
}

function run_() {
  runTests();
  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All Duplicate Invoices detector checks passed.");
  }
}

run_();
