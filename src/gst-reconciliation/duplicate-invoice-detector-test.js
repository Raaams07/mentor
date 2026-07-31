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
  assert(sameAmountDifferentNumber.clusters.length === 1, "same GSTIN, same amount, different invoice numbers, close dates -> one cluster");
  assert(sameAmountDifferentNumber.clusters[0].matchReason === "same_amount", "...matched on amount only");

  const threeEntriesOneCluster = run([
    [GSTIN_A, "2026-05-01", "INV-300", 8000],
    [GSTIN_A, "2026-05-02", "INV-300", 8000],
    [GSTIN_A, "2026-05-05", "INV-300", 8000],
  ]);
  assert(threeEntriesOneCluster.clusters.length === 1, "the same invoice entered three times forms ONE cluster, not three overlapping pairs");
  assert(threeEntriesOneCluster.clusters[0].members.length === 3, "...containing all three entries");

  console.log("\n-- Should NOT be flagged --\n");

  const sameInvoiceNumberFarApart = run([
    [GSTIN_A, "2026-01-01", "INV-400", 10000],
    [GSTIN_A, "2026-08-01", "INV-400", 10000],
  ]);
  assert(sameInvoiceNumberFarApart.clusters.length === 0, "same invoice number, same amount, but ~7 months apart -> NOT flagged (outside the proximity window — likely a distinct re-supply, not a duplicate entry)");

  const recurringMonthlyCharge = run([
    [GSTIN_A, "2026-05-01", "RENT-MAY", 20000],
    [GSTIN_A, "2026-06-01", "RENT-JUN", 20000],
  ]);
  assert(recurringMonthlyCharge.clusters.length === 0, "two consecutive months' rent — same amount, ~31 days apart, different invoice numbers — is NOT flagged as a duplicate with the default 15-day window");

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

  const widerWindowCatchesIt = run(
    [
      [GSTIN_A, "2026-01-01", "INV-700", 10000],
      [GSTIN_A, "2026-01-25", "INV-700", 10000],
    ],
    { windowDays: 30 }
  );
  assert(widerWindowCatchesIt.clusters.length === 1, "a caller-supplied wider window (30 days) catches a pair the tighter default (15 days) would have missed at 24 days apart");

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
