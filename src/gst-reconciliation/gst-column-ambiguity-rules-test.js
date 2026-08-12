/*
 * gst-column-ambiguity-rules-test.js
 * --------------------------------------
 * Proves fieldsInPlay()'s two deliberately-different triggers:
 *   - Ambiguity (2+ candidates) fires for EVERY header-based field,
 *     regardless of whether the current role strictly needs it.
 *   - Zero-match only fires for the small set of fields that are actually
 *     load-bearing once a sheet has already been recognized as a role —
 *     everything else (Place of Supply, Filing Period, etc.) stays silent
 *     on zero-match, since their absence is normal on many real sheets.
 *
 * candidates objects below are hand-built plain objects (not run through
 * identifyGstColumns()) — this file tests fieldsInPlay()'s decision logic
 * in isolation, deliberately independent of the column-identification
 * mechanics already covered by gst-column-identifier-test.js.
 *
 * Run with: node src/gst-reconciliation/gst-column-ambiguity-rules-test.js
 */

const { fieldsInPlay, GSTIN_MARGIN_TO_RESOLVE, AMBIGUITY_CHECKED_FIELDS } = require("./gst-column-ambiguity-rules.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

// Baseline: every field has exactly one candidate, all required fields
// present — a "nothing wrong" sheet, so fieldsInPlay() should return [].
function baseCandidates(overrides) {
  const base = {
    taxableValue: [1],
    igst: [2],
    cgst: [3],
    sgst: [4],
    rate: [5],
    invoiceNumber: [6],
    voucherNumber: [],
    particulars: [7],
    tradeLegalName: [8],
    placeOfSupply: [],
    invoiceType: [],
    filingPeriod: [],
    reverseCharge: [],
    gstin: [{ index: 0, ratio: 0.95 }],
  };
  return { ...base, ...(overrides || {}) };
}

function fieldNames(inPlay) {
  return inPlay.map((i) => i.field);
}

function runNothingInPlayOnCleanSheet() {
  console.log("-- A clean, unambiguous sheet has nothing in play --\n");
  assert(fieldsInPlay("gstr2a", baseCandidates()).length === 0, "gstr2a: no ambiguity, no missing required field -> []");
  assert(fieldsInPlay("purchase_register", baseCandidates({ voucherNumber: [9] })).length === 0, "purchase_register with voucherNumber present -> []");
  console.log("");
}

function runAmbiguityAppliesToEveryField() {
  console.log("-- Ambiguity (2+ candidates) fires for every header-based field, not just tax columns --\n");

  for (const field of AMBIGUITY_CHECKED_FIELDS) {
    // voucherNumber present (unless IT is the field under test) so
    // purchase_register's own required-field rule doesn't also fire and
    // confuse the assertion.
    const overrides = { [field]: [1, 2] };
    if (field !== "voucherNumber") overrides.voucherNumber = [9];
    const candidates = baseCandidates(overrides);
    const inPlay = fieldsInPlay("purchase_register", candidates);
    assert(fieldNames(inPlay).includes(field), `2+ candidates for '${field}' -> flagged as ambiguous`);
    const ambiguousEntry = inPlay.find((i) => i.field === field);
    assert(ambiguousEntry.reason === "ambiguous" && JSON.stringify(ambiguousEntry.candidateIndices) === JSON.stringify([1, 2]), `'${field}' ambiguous entry carries the real candidate indices`);
  }

  // The Tally CGST bug itself, at this layer.
  const tallyLike = baseCandidates({ cgst: [2, 3, 4, 5], voucherNumber: [9] });
  const inPlay = fieldsInPlay("purchase_register", tallyLike);
  assert(fieldNames(inPlay).includes("cgst"), "Tally-style 4-candidate CGST scenario is flagged in-play");

  console.log("");
}

function runZeroMatchScopedToRequiredFields() {
  console.log("-- Zero-match only asks for load-bearing fields; everything else stays silent --\n");

  const placeOfSupplyAbsent = baseCandidates({ placeOfSupply: [], voucherNumber: [9] });
  assert(fieldNames(fieldsInPlay("purchase_register", placeOfSupplyAbsent)).indexOf("placeOfSupply") === -1, "placeOfSupply zero-match on purchase_register stays silent (normal, expected absence)");

  const filingPeriodAbsent = baseCandidates({ filingPeriod: [], voucherNumber: [9] });
  assert(fieldNames(fieldsInPlay("purchase_register", filingPeriodAbsent)).indexOf("filingPeriod") === -1, "filingPeriod zero-match stays silent");

  const reverseChargeAbsent = baseCandidates({ reverseCharge: [], voucherNumber: [9] });
  assert(fieldNames(fieldsInPlay("gstr2a", reverseChargeAbsent)).indexOf("reverseCharge") === -1, "reverseCharge zero-match stays silent even on gstr2a");

  const taxableValueAbsent = baseCandidates({ taxableValue: [], voucherNumber: [9] });
  assert(fieldNames(fieldsInPlay("purchase_register", taxableValueAbsent)).includes("taxableValue"), "taxableValue zero-match IS flagged (required for every role)");

  console.log("");
}

function runInvoiceVoucherNumberEitherOr() {
  console.log("-- purchase_register: invoiceNumber/voucherNumber required as an 'either' pair --\n");

  const onlyVoucher = baseCandidates({ invoiceNumber: [], voucherNumber: [9] });
  const inPlay1 = fieldsInPlay("purchase_register", onlyVoucher);
  assert(!fieldNames(inPlay1).includes("invoiceNumber") && !fieldNames(inPlay1).includes("voucherNumber"), "voucherNumber alone satisfies the requirement -> neither field in play");

  const onlyInvoice = baseCandidates({ invoiceNumber: [6], voucherNumber: [] });
  const inPlay2 = fieldsInPlay("purchase_register", onlyInvoice);
  assert(!fieldNames(inPlay2).includes("invoiceNumber") && !fieldNames(inPlay2).includes("voucherNumber"), "invoiceNumber alone also satisfies the requirement");

  const neither = baseCandidates({ invoiceNumber: [], voucherNumber: [] });
  const inPlay3 = fieldsInPlay("purchase_register", neither);
  assert(fieldNames(inPlay3).includes("invoiceNumber") && fieldNames(inPlay3).includes("voucherNumber"), "both absent -> BOTH flagged in play");

  // gstr2a only ever requires invoiceNumber (no voucherNumber concept for a government return).
  const gstr2aNoInvoice = baseCandidates({ invoiceNumber: [] });
  const inPlayGstr2a = fieldsInPlay("gstr2a", gstr2aNoInvoice);
  assert(fieldNames(inPlayGstr2a).includes("invoiceNumber") && !fieldNames(inPlayGstr2a).includes("voucherNumber"), "gstr2a: only invoiceNumber is flagged, never voucherNumber");

  console.log("");
}

function runAllThreeTaxFieldsZero() {
  console.log("-- All-three-tax-fields-simultaneously-zero is the defensive trigger; one-of-three-zero stays silent --\n");

  const allZero = baseCandidates({ igst: [], cgst: [], sgst: [], voucherNumber: [9] });
  const inPlayAll = fieldsInPlay("purchase_register", allZero);
  assert(fieldNames(inPlayAll).includes("igst") && fieldNames(inPlayAll).includes("cgst") && fieldNames(inPlayAll).includes("sgst"), "all three tax fields zero -> all three flagged");

  // A legitimate IGST-only vendor register — cgst/sgst are correctly
  // absent, must NOT be flagged.
  const igstOnly = baseCandidates({ igst: [2], cgst: [], sgst: [], voucherNumber: [9] });
  const inPlayIgstOnly = fieldsInPlay("purchase_register", igstOnly);
  assert(!fieldNames(inPlayIgstOnly).some((f) => f === "igst" || f === "cgst" || f === "sgst"), "IGST-only vendor (cgst/sgst legitimately absent) -> nothing flagged");

  console.log("");
}

function runRateNeverAsksOnZeroMatch() {
  console.log("-- rate zero-match never asks, regardless of role --\n");
  const noRate = baseCandidates({ rate: [], voucherNumber: [9] });
  assert(!fieldNames(fieldsInPlay("purchase_register", noRate)).includes("rate"), "rate absent on purchase_register -> never flagged");
  assert(!fieldNames(fieldsInPlay("gstr2a", baseCandidates({ rate: [] }))).includes("rate"), "rate absent on gstr2a -> never flagged");
  console.log("");
}

function runGstinMarginThresholding() {
  console.log("-- GSTIN near-tie margin thresholding --\n");

  const nearTie = baseCandidates({
    gstin: [
      { index: 0, ratio: 0.95 },
      { index: 1, ratio: 0.92 },
    ],
    voucherNumber: [9],
  });
  assert(fieldNames(fieldsInPlay("purchase_register", nearTie)).includes("gstin"), `margin 0.03 < GSTIN_MARGIN_TO_RESOLVE (${GSTIN_MARGIN_TO_RESOLVE}) -> flagged ambiguous`);

  const clearWinner = baseCandidates({
    gstin: [
      { index: 0, ratio: 0.95 },
      { index: 1, ratio: 0.65 },
    ],
    voucherNumber: [9],
  });
  assert(!fieldNames(fieldsInPlay("purchase_register", clearWinner)).includes("gstin"), "margin 0.3, well clear -> NOT flagged, silent as today");

  console.log("");
}

function run() {
  runNothingInPlayOnCleanSheet();
  runAmbiguityAppliesToEveryField();
  runZeroMatchScopedToRequiredFields();
  runInvoiceVoucherNumberEitherOr();
  runAllThreeTaxFieldsZero();
  runRateNeverAsksOnZeroMatch();
  runGstinMarginThresholding();

  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All gst-column-ambiguity-rules checks passed.");
  }
}

run();
