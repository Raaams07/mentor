/*
 * cross-sheet-wrong-head-detector-test.js
 * ---------------------------------------------
 * Unit tests for cross-sheet-wrong-head-detector.js, using only
 * fictional synthetic GSTINs and made-up invoice data. No real vendor or
 * invoice data.
 *
 * Run with: node src/gst-reconciliation/cross-sheet-wrong-head-detector-test.js
 */

const { detectCrossSheetWrongHead } = require("./cross-sheet-wrong-head-detector.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

// Supplier registered in Andhra Pradesh (state code 37), Place of Supply
// "Telangana" (36) -> different states -> IGST is the legally correct head.
const GSTIN_INTERSTATE = "37ZZAPL5432Q1Z9";
// Supplier registered in Telangana (36), Place of Supply "Telangana" too
// -> same state -> CGST+SGST is the legally correct head.
const GSTIN_INTRASTATE = "36ZZAPL7788Q1ZO";

const GSTR2A_HEADER = ["GSTIN", "Invoice Number", "IGST", "CGST", "SGST", "Taxable Value", "Place of Supply"];
const GSTR2A_COLUMNS = { gstin: 0, invoiceNumber: 1, voucherNumber: null, igst: 2, cgst: 3, sgst: 4, taxableValue: 5, placeOfSupply: 6 };

const BOOKS_HEADER = ["GSTIN", "Voucher No.", "IGST", "CGST", "SGST", "Taxable Value"];
const BOOKS_COLUMNS = { gstin: 0, invoiceNumber: null, voucherNumber: 1, igst: 2, cgst: 3, sgst: 4, taxableValue: 5, placeOfSupply: null };

function gstr2aInput(rows) {
  return { values: [GSTR2A_HEADER, ...rows], headerRowIndex: 0, columns: GSTR2A_COLUMNS };
}
function booksInput(rows) {
  return { values: [BOOKS_HEADER, ...rows], headerRowIndex: 0, columns: BOOKS_COLUMNS };
}

function runTests() {
  console.log("-- Clean matched pair: correct head on both sides -> NOT flagged --\n");

  const clean = detectCrossSheetWrongHead(
    gstr2aInput([[GSTIN_INTERSTATE, "INV-001", 1000, 0, 0, 10000, "Telangana"]]),
    booksInput([[GSTIN_INTERSTATE, "INV-001", 1000, 0, 0, 10000]])
  );
  assert(clean.applicable === true, "applicable when both sheets have GSTIN + identifier columns");
  assert(clean.matchedInvoiceCount === 1, "the one invoice pair was confidently matched");
  assert(clean.flagged.length === 0, "same head (IGST) on both sides for the same invoice -> not flagged");

  console.log("\n-- The real-world case this rule was built for: 2A right, Books wrong -- FLAGGED --\n");

  const twoARightBooksWrong = detectCrossSheetWrongHead(
    gstr2aInput([[GSTIN_INTERSTATE, "INV-002", 6800, 0, 0, 136000, "Telangana"]]),
    booksInput([[GSTIN_INTERSTATE, "INV-002", 0, 3400, 3400, 136000]])
  );
  assert(twoARightBooksWrong.flagged.length === 1, "2A correctly IGST, Books wrongly CGST+SGST for the SAME invoice -> flagged");
  const f1 = twoARightBooksWrong.flagged[0];
  assert(f1.gstr2aHead === "igst" && f1.booksHead === "cgst_sgst", "both sides' actual heads reported correctly");
  assert(f1.expectedHead === "igst", "expected head correctly derived from 2A's own supplier-state vs Place of Supply data");
  assert(f1.incorrectSide === "books", "correctly identifies BOOKS as the side at fault, not 2A");
  assert(f1.gstin === GSTIN_INTERSTATE.toUpperCase() && f1.identifier === "INV-002", "flagged entry identifies the specific invoice, not just a vendor-level gap");

  console.log("\n-- Mirror case: Books right, 2A wrong -- FLAGGED, correct side identified --\n");

  const booksRightGstr2aWrong = detectCrossSheetWrongHead(
    gstr2aInput([[GSTIN_INTRASTATE, "INV-003", 2000, 0, 0, 20000, "Telangana"]]), // wrongly IGST -- same state should be CGST+SGST
    booksInput([[GSTIN_INTRASTATE, "INV-003", 0, 1000, 1000, 20000]]) // correctly CGST+SGST
  );
  assert(booksRightGstr2aWrong.flagged.length === 1, "2A wrongly IGST, Books correctly CGST+SGST -> flagged");
  assert(booksRightGstr2aWrong.flagged[0].incorrectSide === "gstr2a", "correctly identifies 2A (not Books) as the side at fault this time");

  console.log("\n-- Invoice numbers don't match cleanly -- must NOT force a pairing --\n");

  const differentIdentifiers = detectCrossSheetWrongHead(
    gstr2aInput([[GSTIN_INTERSTATE, "INV-100", 1000, 0, 0, 10000, "Telangana"]]),
    booksInput([[GSTIN_INTERSTATE, "INV-100A", 0, 500, 500, 10000]]) // looks related but is NOT an exact match, and neither side is a bare number for the fuzzy fallback to bridge
  );
  assert(differentIdentifiers.flagged.length === 0, "'INV-100' vs 'INV-100A' are not treated as the same invoice — no match, exact or fuzzy");
  assert(differentIdentifiers.matchedInvoiceCount === 0, "correctly reports zero confidently-matched invoices for this pair");

  console.log("\n-- Fuzzy fallback: same invoice, formatted differently on each side (validated against real data) --\n");

  const fuzzyPrefixMatch = detectCrossSheetWrongHead(
    gstr2aInput([[GSTIN_INTERSTATE, "GSTINV/163", 6800, 0, 0, 136000, "Telangana"]]), // 2A prefixes the running number — a real government-portal export pattern
    booksInput([[GSTIN_INTERSTATE, "163", 0, 3400, 3400, 136000]]) // Books just has the bare number
  );
  assert(fuzzyPrefixMatch.matchedInvoiceCount === 1, "'GSTINV/163' (2A) and '163' (Books) are recognized as the SAME invoice via the fuzzy fallback -- not left unpaired");
  assert(fuzzyPrefixMatch.flagged.length === 1 && fuzzyPrefixMatch.flagged[0].incorrectSide === "books", "the fuzzy-matched pair is compared for Wrong Head exactly like an exact match would be");

  const fuzzySuffixMatch = detectCrossSheetWrongHead(
    gstr2aInput([[GSTIN_INTERSTATE, "5015/21-22", 1000, 0, 0, 10000, "Telangana"]]), // 2A appends a financial-year suffix
    booksInput([[GSTIN_INTERSTATE, "5015", 1000, 0, 0, 10000]]) // Books has the bare running number
  );
  assert(fuzzySuffixMatch.matchedInvoiceCount === 1, "'5015/21-22' (2A) and '5015' (Books) are recognized as the same invoice -- the year suffix doesn't block the match");
  assert(fuzzySuffixMatch.flagged.length === 0, "same head both sides once correctly paired -> not flagged");

  const fuzzyTooShortToBridge = detectCrossSheetWrongHead(
    gstr2aInput([[GSTIN_INTERSTATE, "PO-21-22", 1000, 0, 0, 10000, "Telangana"]]),
    booksInput([[GSTIN_INTERSTATE, "22", 0, 500, 500, 10000]]) // a bare 2-digit number -- too short to trust as a real invoice serial, could just be a date/year fragment
  );
  assert(fuzzyTooShortToBridge.matchedInvoiceCount === 0, "a bare 2-digit identifier is NOT bridged via the fuzzy fallback -- too likely to coincidentally collide with a date/year fragment elsewhere in a longer identifier");

  const fuzzyAmbiguousNotGuessed = detectCrossSheetWrongHead(
    gstr2aInput([
      [GSTIN_INTERSTATE, "GSTINV/500", 1000, 0, 0, 10000, "Telangana"],
      [GSTIN_INTERSTATE, "OLDINV/500", 1000, 0, 0, 10000, "Telangana"], // TWO different 2A identifiers both embed "500" for the same GSTIN
    ]),
    booksInput([[GSTIN_INTERSTATE, "500", 0, 500, 500, 10000]])
  );
  assert(fuzzyAmbiguousNotGuessed.matchedInvoiceCount === 0, "two different complex identifiers both embedding the same number for one GSTIN -> ambiguous, the fuzzy fallback does NOT guess which one is the real match");

  console.log("\n-- Same (GSTIN, identifier) ambiguous on one side -- skipped, not guessed --\n");

  const ambiguousOnGstr2a = detectCrossSheetWrongHead(
    gstr2aInput([
      [GSTIN_INTERSTATE, "INV-200", 1000, 0, 0, 10000, "Telangana"],
      [GSTIN_INTERSTATE, "INV-200", 1000, 0, 0, 10000, "Telangana"], // same key appears twice in 2A -- ambiguous
    ]),
    booksInput([[GSTIN_INTERSTATE, "INV-200", 0, 500, 500, 10000]])
  );
  assert(ambiguousOnGstr2a.flagged.length === 0, "a (GSTIN, identifier) key that's ambiguous on the 2A side is skipped, even though it could look like a mismatch");
  assert(ambiguousOnGstr2a.ambiguousCount === 1, "the ambiguity is explicitly counted/reported, not silently dropped");

  console.log("\n-- Case not present in Books at all -- not this rule's concern --\n");

  const onlyInGstr2a = detectCrossSheetWrongHead(gstr2aInput([[GSTIN_INTERSTATE, "INV-300", 1000, 0, 0, 10000, "Telangana"]]), booksInput([]));
  assert(onlyInGstr2a.flagged.length === 0, "an invoice present only in 2A (not in Books at all) is not flagged here — that's Step 1's 'Extra in 2A' territory");

  console.log("\n-- 'mixed' head on one side -- deferred to the existing single-sheet rule, not double-flagged here --\n");

  const mixedOnGstr2a = detectCrossSheetWrongHead(
    gstr2aInput([[GSTIN_INTERSTATE, "INV-400", 1000, 500, 500, 10000, "Telangana"]]), // both IGST and CGST/SGST populated -- invalid on its own
    booksInput([[GSTIN_INTERSTATE, "INV-400", 1000, 0, 0, 10000]])
  );
  assert(mixedOnGstr2a.flagged.length === 0, "a 'mixed' (both IGST and CGST/SGST populated) row is skipped by this cross-sheet rule -- wrong-head-detector.js's own single-sheet check already covers that row");

  console.log("\n-- 'none' (no tax at all) on one side -- a presence/amount question, not a head question --\n");

  const noneOnBooks = detectCrossSheetWrongHead(
    gstr2aInput([[GSTIN_INTERSTATE, "INV-500", 1000, 0, 0, 10000, "Telangana"]]),
    booksInput([[GSTIN_INTERSTATE, "INV-500", 0, 0, 0, 10000]]) // no tax recorded at all
  );
  assert(noneOnBooks.flagged.length === 0, "one side showing zero tax at all is not a 'wrong head' finding -- that's an amount/presence gap, not a head mismatch");

  console.log("\n-- Heads differ but Place of Supply can't be resolved -- incorrectSide reported as undetermined, not guessed --\n");

  const unresolvablePlaceOfSupply = detectCrossSheetWrongHead(
    gstr2aInput([[GSTIN_INTERSTATE, "INV-600", 1000, 0, 0, 10000, "Somewhere Unrecognizable"]]),
    booksInput([[GSTIN_INTERSTATE, "INV-600", 0, 500, 500, 10000]])
  );
  assert(unresolvablePlaceOfSupply.flagged.length === 1, "heads genuinely differ, so it's still flagged for review...");
  assert(unresolvablePlaceOfSupply.flagged[0].incorrectSide === null, "...but WHICH side is at fault is reported as undetermined (null), not guessed, since Place of Supply couldn't be resolved to a state code");
  assert(unresolvablePlaceOfSupply.flagged[0].expectedHead === null, "expectedHead is also null for the same reason");

  console.log("\n-- Same head, both sides -- but both arguably wrong per state comparison -- not this rule's job --\n");

  const bothSidesAgreeButBothWrong = detectCrossSheetWrongHead(
    gstr2aInput([[GSTIN_INTERSTATE, "INV-700", 0, 500, 500, 10000, "Telangana"]]), // wrongly CGST+SGST per its own state data (should be IGST) -- wrong-head-detector.js's job to catch, not this one
    booksInput([[GSTIN_INTERSTATE, "INV-700", 0, 500, 500, 10000]]) // Books agrees with 2A
  );
  assert(bothSidesAgreeButBothWrong.flagged.length === 0, "when both sides AGREE with each other (even if both are wrong relative to law), this cross-sheet rule has nothing to flag — that's wrong-head-detector.js's existing single-sheet job");

  console.log("\n-- Different GSTINs, same identifier -- correctly NOT paired --\n");

  const differentGstins = detectCrossSheetWrongHead(
    gstr2aInput([[GSTIN_INTERSTATE, "INV-800", 1000, 0, 0, 10000, "Telangana"]]),
    booksInput([[GSTIN_INTRASTATE, "INV-800", 0, 500, 500, 10000]])
  );
  assert(differentGstins.flagged.length === 0, "same invoice identifier but different GSTINs -> never paired (matching key is GSTIN + identifier together)");
  assert(differentGstins.matchedInvoiceCount === 0, "correctly zero matched invoices");

  console.log("\n-- Sheets missing the needed columns --\n");

  const noIdentifierCol = detectCrossSheetWrongHead({ ...gstr2aInput([[GSTIN_INTERSTATE, "INV-900", 1000, 0, 0, 10000, "Telangana"]]), columns: { ...GSTR2A_COLUMNS, invoiceNumber: null, voucherNumber: null } }, booksInput([]));
  assert(noIdentifierCol.applicable === false, "no invoice/voucher number column on 2A -> not applicable, not silently zero");

  console.log("");
}

function run() {
  runTests();
  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All cross-sheet Wrong Head detector checks passed.");
  }
}

run();
