/*
 * duplicate-invoice-detector-test.js
 * --------------------------------------
 * Unit tests for duplicate-invoice-detector.js, using only fictional
 * synthetic GSTINs and made-up invoice data. No real vendor or invoice
 * data.
 *
 * Run with: node src/gst-reconciliation/duplicate-invoice-detector-test.js
 */

const { findDuplicateInvoices, nearSequentialInvoiceNumbers } = require("./duplicate-invoice-detector.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

const GSTIN_A = "27ZZAPL5432Q1Z9";
const GSTIN_B = "07ZZAPL7788Q1ZO";

// columns index layout: [GSTIN, Date, Invoice No, Taxable Value] — no
// IGST/CGST/SGST identified (explicitly null, matching identifyGstColumns'
// own default shape, not merely omitted) for tests that don't care about
// the tax-rate check.
const COLUMNS = { gstin: 0, invoiceNumber: 2, voucherNumber: null, taxableValue: 3, igst: null, cgst: null, sgst: null, dateColumns: [1] };
const HEADER_ROW = ["GSTIN", "Date", "Invoice No", "Taxable Value"];

// columns index layout: [GSTIN, Date, Invoice No, Taxable Value, IGST, CGST, SGST]
const COLUMNS_WITH_TAX = { gstin: 0, invoiceNumber: 2, voucherNumber: null, taxableValue: 3, igst: 4, cgst: 5, sgst: 6, dateColumns: [1] };
const HEADER_ROW_WITH_TAX = ["GSTIN", "Date", "Invoice No", "Taxable Value", "IGST", "CGST", "SGST"];

function run(rows, options) {
  return findDuplicateInvoices([HEADER_ROW, ...rows], 0, COLUMNS, options);
}

function runWithTax(rows, options) {
  return findDuplicateInvoices([HEADER_ROW_WITH_TAX, ...rows], 0, COLUMNS_WITH_TAX, options);
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
    [GSTIN_A, "2026-05-01", "REF-A", 5000],
    [GSTIN_A, "2026-05-04", "MISC-Z", 5000], // deliberately non-sequential-looking identifiers so this isolates the plain same-amount path from the near-sequential exclusion (see point 4 tests below)
  ]);
  assert(sameAmountDifferentNumber.clusters.length === 1, "same GSTIN, same amount, different (non-sequential) invoice numbers, 3 days apart -> one cluster");
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

  console.log("\n-- Regression: the specific over-flagging bug this file was originally fixed for (date chaining) --\n");

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

  // Identifiers here are deliberately non-numeric (no digit run at all) so
  // this test isolates the DATE-chaining span constraint from the newer
  // near-sequential-invoice-number exclusion (point 4) — both would
  // independently prevent this cluster from merging, and this test wants
  // to verify the date-span constraint specifically still works on its own.
  const chainedButShouldNotMergeIntoOneWideCluster = run([
    [GSTIN_A, "2026-01-01", "FEE-ALPHA", 339],
    [GSTIN_A, "2026-01-03", "FEE-BRAVO", 339], // 2 days from the previous -- within the tight amount-only window
    [GSTIN_A, "2026-01-05", "FEE-CHARLIE", 339], // 2 days from the previous
    [GSTIN_A, "2026-01-07", "FEE-DELTA", 339], // 2 days from the previous -- but 6 days from the FIRST entry
    [GSTIN_A, "2026-01-09", "FEE-ECHO", 339], // 2 days from the previous -- but 8 days from the first entry
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

  console.log("\n-- Point 1: same invoice number, but a legitimate multi-rate line split (differing taxable value and/or tax rate) --\n");

  const multiRateSplitDifferentValue = runWithTax([
    [GSTIN_A, "2026-05-01", "INV-900", 10000, 1800, 0, 0], // 18% IGST line
    [GSTIN_A, "2026-05-01", "INV-900", 5000, 250, 0, 0], // 5% IGST line, SAME invoice, DIFFERENT taxable value
  ]);
  assert(
    multiRateSplitDifferentValue.clusters.length === 0,
    "same invoice number, same date, but a different taxable value per line (a legitimate multi-rate GSTR-2A split) -> NOT flagged as a duplicate"
  );

  const multiRateSplitDifferentRateOnly = runWithTax([
    [GSTIN_A, "2026-05-01", "INV-950", 10000, 1800, 0, 0], // 18% IGST
    [GSTIN_A, "2026-05-01", "INV-950", 10000, 500, 0, 0], // 5% IGST — same taxable value, but a different rate
  ]);
  assert(
    multiRateSplitDifferentRateOnly.clusters.length === 0,
    "same invoice number, same date, SAME taxable value but a different effective tax rate -> NOT flagged (still a legitimate split, not a duplicate)"
  );

  const genuineDuplicateWithMatchingTaxData = runWithTax([
    [GSTIN_A, "2026-05-01", "INV-970", 10000, 1800, 0, 0],
    [GSTIN_A, "2026-05-03", "INV-970", 10000, 1800, 0, 0], // same value AND same rate
  ]);
  assert(
    genuineDuplicateWithMatchingTaxData.clusters.length === 1,
    "same invoice number, same taxable value, AND same tax rate -> still correctly flagged as a genuine duplicate (the new checks don't suppress real duplicates when tax data is available and consistent)"
  );

  console.log("\n-- Point 2: same-amount recurrence elsewhere in the file --\n");

  const isolatedPairNoRecurrenceElsewhere = run([
    [GSTIN_A, "2026-05-01", "ALPHA", 500],
    [GSTIN_A, "2026-05-02", "ZULU", 500],
  ]);
  assert(
    isolatedPairNoRecurrenceElsewhere.clusters.length === 1,
    "a close same-amount pair with NO other occurrence of that amount anywhere else in the file -> still flagged (isolated pairs remain the strongest same_amount signal)"
  );

  const pairWithTwoExtraOccurrencesElsewhere = run([
    [GSTIN_A, "2026-05-01", "ALPHA", 500],
    [GSTIN_A, "2026-05-02", "ZULU", 500],
    [GSTIN_A, "2026-01-01", "FARJAN", 500], // far away in time, but the SAME amount — counts as recurrence evidence regardless of date
    [GSTIN_A, "2026-09-01", "FARSEP", 500], // a second far-away occurrence
  ]);
  assert(
    pairWithTwoExtraOccurrencesElsewhere.clusters.length === 0,
    "the same close pair, but this amount ALSO recurs on 2 more rows elsewhere in the file (far from either date) -> excluded entirely as a routine recurring amount, not a duplicate"
  );

  const pairWithOnlyOneExtraOccurrenceElsewhere = run([
    [GSTIN_A, "2026-05-01", "ALPHA", 500],
    [GSTIN_A, "2026-05-02", "ZULU", 500],
    [GSTIN_A, "2026-01-01", "FARJAN", 500], // just ONE extra occurrence elsewhere — meets the (lowered) 1+ exclusion threshold on its own
  ]);
  assert(
    pairWithOnlyOneExtraOccurrenceElsewhere.clusters.length === 0,
    "even a single extra occurrence elsewhere (a low-volume routine charge — 3 total occurrences, not just a high-volume one) -> excluded, same as the higher-volume case"
  );

  console.log("\n-- Regression: the specific known real-data false positive this whole fix targets (many identical same-day amounts) --\n");

  const manySameDayIdenticalAmounts = run([
    [GSTIN_A, "2026-04-10", "BANK-A", 500],
    [GSTIN_A, "2026-04-10", "BANK-B", 500],
    [GSTIN_A, "2026-04-10", "BANK-C", 500],
    [GSTIN_A, "2026-04-10", "BANK-D", 500],
    [GSTIN_A, "2026-04-10", "BANK-E", 500],
    [GSTIN_A, "2026-04-10", "BANK-F", 500],
    [GSTIN_A, "2026-04-10", "BANK-G", 500],
    [GSTIN_A, "2026-04-10", "BANK-H", 500],
    [GSTIN_A, "2026-04-10", "BANK-I", 500],
  ]);
  assert(
    manySameDayIdenticalAmounts.clusters.length === 0,
    "nine same-amount rows on the SAME day (date span trivially satisfies the tight window) -> NOT flagged: each pair has many other same-amount rows elsewhere in the group, so the recurrence-elsewhere exclusion (point 2) correctly recognizes this as routine, not a 9-way duplicate cluster"
  );

  console.log("\n-- Point 3: amount drifting through a chain, all on one date (isolating amount-span capping from date-span capping) --\n");

  const amountDriftChain = run(
    [
      [GSTIN_A, "2026-06-01", "FEE-ONE", 100.0],
      [GSTIN_A, "2026-06-01", "FEE-TWO", 100.9], // 0.9 from FEE-ONE -- within ₹1 tolerance
      [GSTIN_A, "2026-06-01", "FEE-THREE", 101.8], // 0.9 from FEE-TWO, but 1.8 from FEE-ONE -- exceeds tolerance from the far end of the chain
      [GSTIN_A, "2026-06-01", "FEE-FOUR", 102.7], // 0.9 from FEE-THREE, but 2.7 from FEE-ONE
    ],
    { amountTolerance: 1 }
  );
  assert(
    !amountDriftChain.clusters.some((c) => c.members.length > 2),
    "amounts drifting ₹0.9 per adjacent step must NOT chain into one cluster spanning ₹2.7 (well past the ₹1 tolerance) just because each individual step was small"
  );
  const maxAmountSpread = amountDriftChain.clusters.reduce((max, c) => {
    const amounts = c.members.map((m) => m.amount);
    return Math.max(max, Math.max(...amounts) - Math.min(...amounts));
  }, 0);
  assert(maxAmountSpread <= 1, "every resulting cluster's own amount spread stays within the ₹1 tolerance");
  console.log(
    "  (this scenario produced " +
      amountDriftChain.clusters.length +
      " cluster(s): " +
      amountDriftChain.clusters.map((c) => "{" + c.members.map((m) => m.amount).join(",") + "}").join(", ") +
      ")"
  );

  console.log("\n-- Point 4: near-sequential invoice numbers with a small date gap read as consecutive invoicing, not re-entry --\n");

  const nearSequentialNumbersNotFlagged = run([
    [GSTIN_A, "2026-05-01", "SEQ-501", 7000],
    [GSTIN_A, "2026-05-03", "SEQ-502", 7000], // one apart, same format, same amount, 2 days apart
  ]);
  assert(
    nearSequentialNumbersNotFlagged.clusters.length === 0,
    "same amount, same format, invoice numbers exactly ONE apart, close in time -> NOT flagged (reads as two genuinely consecutive invoices)"
  );

  const farApartNumbersStillFlagged = run([
    [GSTIN_A, "2026-05-01", "SEQ-100", 7000],
    [GSTIN_A, "2026-05-03", "SEQ-999", 7000], // same format, but 899 apart -- NOT near-sequential
  ]);
  assert(
    farApartNumbersStillFlagged.clusters.length === 1,
    "same amount, same format, but invoice numbers far apart numerically -> still flagged (the near-sequential exclusion doesn't overfire on a large numeric gap)"
  );

  console.log("\n-- Point 1 follow-up: near-sequential threshold scales to the vendor's OWN typical numbering gap --\n");

  const vendorTypicalGapExcludesLargerNumericGap = run([
    // Establishes this vendor's own typical invoice-number spacing (~50 apart) — different amounts so these don't themselves interact with the candidate pair via same_amount.
    [GSTIN_A, "2026-01-01", "V-100", 999],
    [GSTIN_A, "2026-02-01", "V-150", 888],
    [GSTIN_A, "2026-03-01", "V-200", 777],
    [GSTIN_A, "2026-04-01", "V-250", 666],
    // Candidate pair: 45 apart -- well past the fixed default (5), but well within this vendor's own established ~50 spacing.
    [GSTIN_A, "2026-05-01", "V-300", 5000],
    [GSTIN_A, "2026-05-02", "V-345", 5000],
  ]);
  assert(
    !vendorTypicalGapExcludesLargerNumericGap.clusters.some((c) => c.members.some((m) => m.amount === 5000)),
    "a 45-apart pair is NOT flagged once the vendor's own invoicing history shows ~50-apart numbering is normal for them — the fixed default(5) alone would have wrongly flagged this"
  );

  const burstyVendorRateJustifiesLargerGapOverLongerDateSpan = run([
    // A bursty vendor: a tight same-day run (gap 1) establishes a small MEDIAN gap, but the vendor's numbering also spans much further over ten days -- a real overall pace of ~10/day.
    [GSTIN_A, "2026-01-01", "W-100", 111],
    [GSTIN_A, "2026-01-01", "W-101", 222],
    [GSTIN_A, "2026-01-01", "W-102", 333],
    [GSTIN_A, "2026-01-11", "W-200", 444],
    // Candidate pair: gap 10, 3 days apart. The vendor's median neighbor-gap (1) alone doesn't justify this, but the vendor's overall PACE (~10/day) easily does over a 3-day span.
    [GSTIN_A, "2026-01-01", "W-104", 5000],
    [GSTIN_A, "2026-01-04", "W-114", 5000],
  ]);
  assert(
    !burstyVendorRateJustifiesLargerGapOverLongerDateSpan.clusters.some((c) => c.members.some((m) => m.amount === 5000)),
    "a bursty vendor's tight same-day runs alone would under-estimate a normal gap over a several-day span — the vendor's overall issuance PACE (value range / date range) correctly justifies it instead"
  );

  const gapTooLargeForEitherSignalStillFlagged = run([
    [GSTIN_A, "2026-01-01", "V2-100", 999],
    [GSTIN_A, "2026-02-01", "V2-150", 888],
    [GSTIN_A, "2026-03-01", "V2-200", 777],
    [GSTIN_A, "2026-04-01", "V2-250", 666],
    // 5000 apart, 1 day gap -- far beyond anything this vendor's own history (~50-apart) or pace would justify.
    [GSTIN_A, "2026-05-01", "V2-300", 6000],
    [GSTIN_A, "2026-05-02", "V2-5300", 6000],
  ]);
  assert(
    gapTooLargeForEitherSignalStillFlagged.clusters.some((c) => c.members.some((m) => m.amount === 6000)),
    "a gap far beyond even this vendor's own established typical spacing is still flagged — the vendor-scaled check isn't unconditionally permissive"
  );

  console.log("\n-- Point 3: identifiers that are the same invoice number with different zero-padding --\n");

  const leadingZeroSameInvoice = run([
    [GSTIN_A, "2026-05-01", "6", 57825],
    [GSTIN_A, "2026-05-01", "006", 57825],
  ]);
  assert(
    leadingZeroSameInvoice.clusters.length === 1 && leadingZeroSameInvoice.clusters[0].matchReason === "same_invoice_number_and_amount",
    "\"6\" and \"006\" are the same invoice number written with different zero-padding -> matched via the identifier path (same_invoice_number_and_amount), not misread as two near-sequential-but-different invoices"
  );

  console.log("\n-- nearSequentialInvoiceNumbers() directly --\n");

  assert(nearSequentialInvoiceNumbers("INV-100", "INV-101", 5) === true, "adjacent numbers, identical structure -> near-sequential");
  assert(nearSequentialInvoiceNumbers("INV-100", "INV-200", 5) === false, "100 apart, well past a default max diff of 5 -> not near-sequential");
  assert(nearSequentialInvoiceNumbers("ST/23/2021-22", "ST/24/2021-22", 5) === true, "one running-number segment apart, financial-year segment unchanged -> near-sequential");
  assert(nearSequentialInvoiceNumbers("ST/23/2021-22", "ST/23/2022-23", 5) === false, "the running number is UNCHANGED but the financial year rolled over -> not the running-number pattern this check targets");
  assert(nearSequentialInvoiceNumbers("INV-100", "BILL-101", 5) === false, "different non-digit structure entirely -> not comparably sequential");
  assert(nearSequentialInvoiceNumbers("", "INV-101", 5) === false, "a blank identifier never counts as near-sequential");
  assert(nearSequentialInvoiceNumbers("INV-100", "INV-100", 5) === false, "identical identifiers aren't this function's concern (that's the same_invoice_number path)");

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
      [GSTIN_A, "2026-05-01", "MISC-A", 9000],
      [GSTIN_A, "2026-05-03", "MISC-B", 9000], // 2 days apart -- caught by the 3-day default, but not by a 1-day override
    ],
    { amountOnlyWindowDays: 1 }
  );
  assert(tighterAmountOnlyWindowMissesIt.clusters.length === 0, "a caller-supplied tighter amount-only window (1 day) correctly misses a pair 2 days apart that the 3-day default would catch");

  const higherElsewhereThresholdKeepsFlagging = run(
    [
      [GSTIN_A, "2026-05-01", "ALPHA", 500],
      [GSTIN_A, "2026-05-02", "ZULU", 500],
      [GSTIN_A, "2026-01-01", "FARJAN", 500],
      [GSTIN_A, "2026-09-01", "FARSEP", 500],
    ],
    { minElsewhereOccurrencesToExclude: 3 } // the default (1) would exclude this pair; requiring 3 does not
  );
  assert(
    higherElsewhereThresholdKeepsFlagging.clusters.length === 1,
    "a caller-supplied higher elsewhere-occurrence threshold (3) keeps flagging a pair that has only 2 extra occurrences elsewhere, which the default (1) would exclude"
  );

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
