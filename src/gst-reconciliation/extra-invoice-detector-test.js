/*
 * extra-invoice-detector-test.js
 * --------------------------------------
 * Unit tests for extra-invoice-detector.js, using only fictional synthetic
 * GSTINs and made-up invoice data. No real vendor or invoice data.
 *
 * Run with: node src/gst-reconciliation/extra-invoice-detector-test.js
 */

const { detectInvoiceLevelExtras, normalizeNumericIdentifier } = require("./extra-invoice-detector.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

const GSTIN_A = "27ZZAPL5432Q1Z9";
const GSTIN_B = "07ZZAPL7788Q1ZO";

// columns index layout: [GSTIN, Invoice No, IGST, CGST, SGST, Taxable Value]
const GSTR2A_COLUMNS = { gstin: 0, invoiceNumber: 1, voucherNumber: null, igst: 2, cgst: 3, sgst: 4, taxableValue: 5, placeOfSupply: null };
const GSTR2A_HEADER = ["GSTIN", "Invoice No", "IGST", "CGST", "SGST", "Taxable Value"];

// columns index layout: [GSTIN, Voucher No, IGST, CGST, SGST, Taxable Value]
const BOOKS_COLUMNS = { gstin: 0, invoiceNumber: null, voucherNumber: 1, igst: 2, cgst: 3, sgst: 4, taxableValue: 5, placeOfSupply: null };
const BOOKS_HEADER = ["GSTIN", "Voucher No", "IGST", "CGST", "SGST", "Taxable Value"];

function run(gstr2aRows, booksRows, options) {
  const gstr2aInput = { values: [GSTR2A_HEADER, ...gstr2aRows], headerRowIndex: 0, columns: GSTR2A_COLUMNS };
  const booksInput = { values: [BOOKS_HEADER, ...booksRows], headerRowIndex: 0, columns: BOOKS_COLUMNS };
  return detectInvoiceLevelExtras(gstr2aInput, booksInput, options);
}

// columns index layout WITH a Date column: [GSTIN, Invoice/Voucher No, IGST, CGST, SGST, Taxable Value, Date] — a separate fixture from the columns above (rather than adding a date to every existing row) since the short-identifier tiebreak needs real dates to corroborate against.
const GSTR2A_COLUMNS_WITH_DATE = { gstin: 0, invoiceNumber: 1, voucherNumber: null, igst: 2, cgst: 3, sgst: 4, taxableValue: 5, placeOfSupply: null, dateColumns: [6] };
const GSTR2A_HEADER_WITH_DATE = ["GSTIN", "Invoice No", "IGST", "CGST", "SGST", "Taxable Value", "Date"];
const BOOKS_COLUMNS_WITH_DATE = { gstin: 0, invoiceNumber: null, voucherNumber: 1, igst: 2, cgst: 3, sgst: 4, taxableValue: 5, placeOfSupply: null, dateColumns: [6] };
const BOOKS_HEADER_WITH_DATE = ["GSTIN", "Voucher No", "IGST", "CGST", "SGST", "Taxable Value", "Date"];

function runWithDate(gstr2aRows, booksRows, options) {
  const gstr2aInput = { values: [GSTR2A_HEADER_WITH_DATE, ...gstr2aRows], headerRowIndex: 0, columns: GSTR2A_COLUMNS_WITH_DATE };
  const booksInput = { values: [BOOKS_HEADER_WITH_DATE, ...booksRows], headerRowIndex: 0, columns: BOOKS_COLUMNS_WITH_DATE };
  return detectInvoiceLevelExtras(gstr2aInput, booksInput, options);
}

function runTests() {
  console.log("-- Core scenario: a vendor with 5 Books invoices, only 1 matching 2A --\n");

  const partiallyMatchedVendor = run(
    [[GSTIN_A, "INV-001", 0, 900, 900, 10000]], // only this one exists in 2A
    [
      [GSTIN_A, "INV-001", 0, 900, 900, 10000], // matches
      [GSTIN_A, "INV-002", 0, 450, 450, 5000],
      [GSTIN_A, "INV-003", 0, 450, 450, 5000],
      [GSTIN_A, "INV-004", 0, 450, 450, 5000],
      [GSTIN_A, "INV-005", 0, 450, 450, 5000],
    ]
  );
  assert(partiallyMatchedVendor.applicable, "applicable when both sheets have GSTIN + identifier columns");
  assert(partiallyMatchedVendor.extraInBooks.length === 4, "4 of the vendor's 5 Books invoices have no match in 2A -> surfaced as Extra in Books, even though the vendor itself has ONE matching invoice");
  assert(
    partiallyMatchedVendor.extraInBooks.every((r) => ["INV-002", "INV-003", "INV-004", "INV-005"].includes(r.identifier)),
    "...specifically the 4 unmatched invoice numbers, not the matched one"
  );
  assert(partiallyMatchedVendor.extraIn2A.length === 0, "the one 2A invoice for this vendor has a match in Books -> not Extra in 2A");

  console.log("\n-- A vendor where every invoice matches --\n");

  const fullyMatchedVendor = run(
    [
      [GSTIN_A, "INV-100", 0, 90, 90, 1000],
      [GSTIN_A, "INV-101", 0, 90, 90, 1000],
    ],
    [
      [GSTIN_A, "INV-100", 0, 90, 90, 1000],
      [GSTIN_A, "INV-101", 0, 90, 90, 1000],
    ]
  );
  assert(fullyMatchedVendor.extraIn2A.length === 0 && fullyMatchedVendor.extraInBooks.length === 0, "every invoice for this vendor matches on both sides -> zero Extra rows, either direction");

  console.log("\n-- A vendor entirely absent from one side --\n");

  const entirelyAbsentFromGstr2a = run(
    [], // nothing in 2A for this vendor at all
    [
      [GSTIN_B, "V-1", 0, 100, 100, 2000],
      [GSTIN_B, "V-2", 0, 200, 200, 3000],
      [GSTIN_B, "V-3", 0, 300, 300, 4000],
    ]
  );
  assert(entirelyAbsentFromGstr2a.extraInBooks.length === 3, "a vendor with zero 2A presence -> ALL of its Books invoices surface as Extra in Books (not collapsed into one vendor-level row)");
  assert(entirelyAbsentFromGstr2a.extraIn2A.length === 0, "nothing on the 2A side for this vendor -> nothing to report as Extra in 2A");

  console.log("\n-- Opposite direction: Extra in 2A --\n");

  const extraIn2AOnly = run(
    [
      [GSTIN_A, "G-1", 0, 50, 50, 500],
      [GSTIN_A, "G-2", 0, 50, 50, 500],
    ],
    [[GSTIN_A, "G-1", 0, 50, 50, 500]] // only G-1 exists in Books
  );
  assert(extraIn2AOnly.extraIn2A.length === 1 && extraIn2AOnly.extraIn2A[0].identifier === "G-2", "G-2 exists in 2A but not Books -> Extra in 2A, even though G-1 (same GSTIN) matches");
  assert(extraIn2AOnly.extraInBooks.length === 0, "nothing extra in Books for this vendor");

  console.log("\n-- Different GSTINs never cross-match --\n");

  const differentGstins = run([[GSTIN_A, "SAME-NUM", 0, 100, 100, 1000]], [[GSTIN_B, "SAME-NUM", 0, 100, 100, 1000]]);
  assert(differentGstins.extraIn2A.length === 1 && differentGstins.extraInBooks.length === 1, "same invoice number but different GSTINs -> matched on GSTIN+identifier together, so BOTH sides report as Extra (they're genuinely different invoices)");

  console.log("\n-- Ambiguous keys (duplicate invoice number within one GSTIN) --\n");

  // Confirmed real-world root cause (see module docstring): a repeated
  // identifier on one side isn't automatically "present" anymore -- each
  // row in the ambiguous group needs its OWN amount+tax corroboration.
  // Here NEITHER 2A row's amount corroborates with the one Books row, so
  // existence-by-identifier-alone is no longer enough; all 3 stay Extra.
  const ambiguousNoAmountCorroboration = run(
    [
      [GSTIN_A, "DUP-1", 0, 100, 100, 1000],
      [GSTIN_A, "DUP-1", 0, 200, 200, 2000], // same GSTIN+identifier appears twice in 2A
    ],
    [[GSTIN_A, "DUP-1", 0, 150, 150, 1500]] // matches NEITHER 2A row's amount
  );
  assert(
    ambiguousNoAmountCorroboration.extraIn2A.length === 2 && ambiguousNoAmountCorroboration.extraInBooks.length === 1,
    "an ambiguous identifier where NO row's amount corroborates with any counterpart in the group -> existence alone no longer suffices, all 3 rows stay Extra"
  );

  // The confirmed real pattern: a repeated identifier where ONE row
  // genuinely corroborates by amount and the OTHER doesn't (an accidental
  // duplicate identifier hiding a genuinely different invoice) -- the
  // corroborated row is correctly excluded, the surplus stays Extra.
  const ambiguousOneRowCorroborates = run(
    [
      [GSTIN_A, "DUP-3", 0, 90, 90, 1000],
      [GSTIN_A, "DUP-3", 0, 180, 180, 2000], // a DIFFERENT invoice that happens to share the same (erroneous) identifier
    ],
    [[GSTIN_A, "DUP-3", 0, 90, 90, 1000]] // genuinely matches only the FIRST 2A row
  );
  assert(
    ambiguousOneRowCorroborates.extraIn2A.length === 1 && ambiguousOneRowCorroborates.extraIn2A[0].taxableValue === 2000,
    "the genuinely-matched 2A row (taxable 1000) is correctly excluded from Extra; the surplus row sharing the SAME identifier by coincidence (taxable 2000) is NOT silently absorbed -- it stays a real candidate"
  );
  assert(ambiguousOneRowCorroborates.extraInBooks.length === 0, "the one Books row has real amount corroboration -> correctly not Extra");

  const ambiguousAndAbsentFromOtherSide = run(
    [
      [GSTIN_A, "DUP-2", 0, 100, 100, 1000],
      [GSTIN_A, "DUP-2", 0, 200, 200, 2000],
    ],
    [] // nothing in Books at all
  );
  assert(
    ambiguousAndAbsentFromOtherSide.extraIn2A.length === 2,
    "an ambiguous identifier that's ALSO completely absent from the other side -> both rows surface as Extra (ambiguity doesn't suppress a genuine gap)"
  );

  console.log("\n-- Fuzzy fallback: same invoice, formatted differently on each side (validated against real data) --\n");

  const fuzzyPrefixNotExtra = run([[GSTIN_A, "GSTINV/163", 0, 900, 900, 10000]], [[GSTIN_A, "163", 0, 900, 900, 10000]]);
  assert(
    fuzzyPrefixNotExtra.extraIn2A.length === 0 && fuzzyPrefixNotExtra.extraInBooks.length === 0,
    "'GSTINV/163' (2A) and '163' (Books) are the same invoice via the fuzzy fallback -- neither side is Extra"
  );

  const fuzzySuffixNotExtra = run([[GSTIN_A, "5015/21-22", 0, 450, 450, 5000]], [[GSTIN_A, "5015", 0, 450, 450, 5000]]);
  assert(
    fuzzySuffixNotExtra.extraIn2A.length === 0 && fuzzySuffixNotExtra.extraInBooks.length === 0,
    "'5015/21-22' (2A) and '5015' (Books) are the same invoice -- the financial-year suffix doesn't block the match"
  );

  const fuzzyTooShortStillExtra = run([[GSTIN_A, "PO-21-22", 0, 100, 100, 1000]], [[GSTIN_A, "22", 0, 100, 100, 1000]]);
  assert(
    fuzzyTooShortStillExtra.extraIn2A.length === 1 && fuzzyTooShortStillExtra.extraInBooks.length === 1,
    "a bare 2-digit identifier is NOT bridged via the fuzzy fallback -- both sides correctly remain Extra rather than risk a coincidental date/year-fragment match"
  );

  const fuzzyAmbiguousStillExtra = run(
    [
      [GSTIN_A, "GSTINV/500", 0, 100, 100, 1000],
      [GSTIN_A, "OLDINV/500", 0, 100, 100, 1000], // TWO different 2A identifiers both embed "500" for the same GSTIN
    ],
    [[GSTIN_A, "500", 0, 100, 100, 1000]]
  );
  assert(
    fuzzyAmbiguousStillExtra.extraIn2A.length === 2 && fuzzyAmbiguousStillExtra.extraInBooks.length === 1,
    "two different 2A identifiers both embedding the same number for one GSTIN -> ambiguous, so the fuzzy fallback does NOT pick one; all three rows (both 2A candidates AND the one Books invoice) stay Extra rather than guessing a pairing"
  );

  console.log("\n-- Row data carried through --\n");

  const rowDataCheck = run([], [[GSTIN_A, "ROWCHK", 0, 900, 900, 10000]]);
  const extra = rowDataCheck.extraInBooks[0];
  assert(extra.gstin === GSTIN_A && extra.identifier === "ROWCHK", "GSTIN and identifier carried through onto the extra row");
  assert(extra.taxableValue === 10000 && extra.igst === 0 && extra.cgst === 900 && extra.sgst === 900, "taxable value and tax fields carried through onto the extra row");
  assert(typeof extra.rowIndex === "number", "row index carried through so the sheet writer can point back to the source row");

  console.log("\n-- RCM rows excluded from Extra-in-2A/Extra-in-Books (the real-world bug: an invoice flagged BOTH ways at once) --\n");

  // 3 GSTIN_A invoices in 2A, none present in Books at all -- without
  // exclusion, all 3 would be Extra in 2A. Row 1 (identifier "RCM-2") is
  // also RCM-classified this run and must be excluded from the output.
  const rcmRowsExcludedFrom2A = run(
    [
      [GSTIN_A, "RCM-1", 0, 100, 100, 1000],
      [GSTIN_A, "RCM-2", 0, 200, 200, 2000],
      [GSTIN_A, "RCM-3", 0, 300, 300, 3000],
    ],
    [],
    { gstr2aExcludedRowIndices: new Set([1]) }
  );
  assert(rcmRowsExcludedFrom2A.extraIn2A.length === 2, "3 unmatched 2A invoices minus the 1 RCM-excluded one -> 2 remain Extra in 2A");
  assert(
    !rcmRowsExcludedFrom2A.extraIn2A.some((r) => r.identifier === "RCM-2"),
    "the specific RCM-excluded invoice does not appear in Extra in 2A at all"
  );
  assert(
    rcmRowsExcludedFrom2A.extraIn2A.some((r) => r.identifier === "RCM-1") && rcmRowsExcludedFrom2A.extraIn2A.some((r) => r.identifier === "RCM-3"),
    "the two NON-RCM invoices are still correctly reported as Extra in 2A -- exclusion is targeted, not blanket"
  );

  // Mirror direction: Books-side RCM exclusion.
  const rcmRowsExcludedFromBooks = run([], [[GSTIN_A, "RCM-B1", 0, 100, 100, 1000], [GSTIN_A, "RCM-B2", 0, 200, 200, 2000]], { booksExcludedRowIndices: new Set([0]) });
  assert(rcmRowsExcludedFromBooks.extraInBooks.length === 1 && rcmRowsExcludedFromBooks.extraInBooks[0].identifier === "RCM-B2", "Books-side RCM exclusion works the same way, mirrored");

  // The correctness risk this design specifically avoids: excluding an
  // RCM row from the OUTPUT must not break matching for a DIFFERENT row
  // that legitimately pairs with it. Here "RCM-4" is RCM-excluded on the
  // 2A side but ALSO genuinely exists in Books — Books' own "RCM-4" row
  // must still be recognized as matched (not wrongly become Extra in
  // Books) even though its 2A counterpart is excluded from being reported.
  const rcmExclusionDoesNotBreakOthersMatch = run(
    [[GSTIN_A, "RCM-4", 0, 500, 500, 5000]],
    [[GSTIN_A, "RCM-4", 0, 500, 500, 5000]],
    { gstr2aExcludedRowIndices: new Set([0]) }
  );
  assert(
    rcmExclusionDoesNotBreakOthersMatch.extraIn2A.length === 0,
    "the RCM-excluded 2A invoice is correctly absent from Extra in 2A (it has a real Books match, so it wouldn't have been extra anyway)"
  );
  assert(
    rcmExclusionDoesNotBreakOthersMatch.extraInBooks.length === 0,
    "critically: Books' matching invoice is NOT wrongly flagged as Extra in Books just because its 2A counterpart was excluded from the OUTPUT -- exclusion must not remove a row's ability to serve as a match target for the other side"
  );

  const noExclusionOptionIsANoOp = run([[GSTIN_A, "PLAIN-1", 0, 100, 100, 1000]], []);
  assert(noExclusionOptionIsANoOp.extraIn2A.length === 1, "omitting the options argument entirely still works exactly as before (no exclusions)");

  console.log("\n-- Short-identifier tiebreak: leading-zero mismatch, the real-world bug (2A '068' vs Books '68') --\n");

  const leadingZeroPair = runWithDate([[GSTIN_A, "068", 0, 15582, 15582, 111300, "2021-04-20"]], [[GSTIN_A, "68", 0, 15582, 15582, 111300, "2021-04-20"]]);
  assert(
    leadingZeroPair.extraIn2A.length === 0 && leadingZeroPair.extraInBooks.length === 0,
    "'068' (2A) and '68' (Books) -- too short for the digit-run fuzzy fallback, but the SAME taxable value and date confidently pair them -- neither side is Extra"
  );

  console.log("\n-- Short-identifier tiebreak: genuinely one-sided invoices must NOT be suppressed --\n");

  const genuinelyOneSided = runWithDate(
    [],
    [
      [GSTIN_A, "1025", 0, 169180.2, 169180.2, 1879780, "2021-11-15"],
      [GSTIN_A, "1026", 0, 170224.2, 170224.2, 1891380, "2021-11-15"],
      [GSTIN_A, "1030", 0, 150948, 150948, 1677200, "2021-11-23"],
    ]
  );
  assert(genuinelyOneSided.extraInBooks.length === 3, "invoices with no counterpart on the 2A side at all stay correctly flagged as Extra in Books -- the tiebreak has nothing to pair them against, so it doesn't suppress them");

  console.log("\n-- Short-identifier tiebreak: genuine repetition, resolved via amount+date --\n");

  // Identifier value 5 appears TWICE on each side for the same GSTIN
  // (real invoices, not a data error) -- amount+date must correctly pair
  // each occurrence to its OWN counterpart, not just the first match found.
  // NOTE: every identifier below is a DIFFERENT raw string on each side
  // (e.g. "05" vs "0005", never literally "5" on both) -- two rows that
  // happen to share the EXACT same string already exact-match via plain
  // existence (regardless of amount), before this tiebreak ever runs. This
  // test is specifically about rows that only coincide once NORMALIZED.
  const repeatedIdentifierResolved = runWithDate(
    [
      [GSTIN_A, "05", 0, 90, 90, 1000, "2021-01-01"],
      [GSTIN_A, "0005", 0, 180, 180, 2000, "2021-02-01"],
    ],
    [
      [GSTIN_A, "5", 0, 90, 90, 1000, "2021-01-01"],
      [GSTIN_A, "00005", 0, 180, 180, 2000, "2021-02-02"], // one day off -- still within the "near date" tolerance
    ]
  );
  assert(
    repeatedIdentifierResolved.extraIn2A.length === 0 && repeatedIdentifierResolved.extraInBooks.length === 0,
    "identifier value 5 repeats twice on BOTH sides (as different zero-padding variants), but amount+date correctly resolves each occurrence to its own distinct counterpart -- not a single arbitrary pairing"
  );

  console.log("\n-- Short-identifier tiebreak: genuine ambiguity is NOT guessed --\n");

  // Every raw string below is >= 3 digits long and distinct from every
  // other one (2A: "0009"/"00009", Books: "000009"/"0000009") so this
  // deliberately ALSO exercises the pre-existing digit-run fuzzy fallback
  // (cross-sheet-wrong-head-detector.js) — with 2 candidates registering
  // on each side for digit value 9, that mechanism already refuses to
  // pick one (ambiguous). This tiebreak must independently reach the same
  // "don't guess" conclusion once amount+date also fails to disambiguate.
  const trulyAmbiguousRepetition = runWithDate(
    [
      [GSTIN_A, "0009", 0, 90, 90, 1000, "2021-03-01"],
      [GSTIN_A, "00009", 0, 90, 90, 1000, "2021-03-01"], // SAME amount AND same date as the row above -- genuinely indistinguishable
    ],
    [
      [GSTIN_A, "000009", 0, 90, 90, 1000, "2021-03-01"],
      [GSTIN_A, "0000009", 0, 90, 90, 1000, "2021-03-01"],
    ]
  );
  assert(
    trulyAmbiguousRepetition.extraIn2A.length === 2 && trulyAmbiguousRepetition.extraInBooks.length === 2,
    "two 2A candidates and two Books candidates share identical amount+date once normalized to the same value -- no way to tell which pairs with which, so none are guessed; all 4 stay Extra"
  );

  console.log("\n-- Short-identifier tiebreak: amount alone, or date alone, is not enough --\n");

  const amountMatchesDateDoesNot = runWithDate([[GSTIN_A, "07", 0, 90, 90, 1000, "2021-01-01"]], [[GSTIN_A, "7", 0, 90, 90, 1000, "2021-06-01"]]);
  assert(amountMatchesDateDoesNot.extraIn2A.length === 1 && amountMatchesDateDoesNot.extraInBooks.length === 1, "same amount, but dates 5 months apart -- not corroborated, stays Extra on both sides");

  const dateMatchesAmountDoesNot = runWithDate([[GSTIN_A, "08", 0, 90, 90, 1000, "2021-01-01"]], [[GSTIN_A, "8", 0, 90, 90, 5000, "2021-01-01"]]);
  assert(dateMatchesAmountDoesNot.extraIn2A.length === 1 && dateMatchesAmountDoesNot.extraInBooks.length === 1, "same date, but very different taxable value -- not corroborated, stays Extra on both sides");

  console.log("\n-- normalizeNumericIdentifier() directly --\n");

  assert(normalizeNumericIdentifier("068") === "68", "leading zeros stripped");
  assert(normalizeNumericIdentifier("68") === "68", "already-bare numbers pass through unchanged");
  assert(normalizeNumericIdentifier("INV-68") === null, "a non-numeric identifier doesn't participate in this tiebreak at all");
  assert(normalizeNumericIdentifier("") === null, "an empty string is not treated as numeric");

  console.log("\n-- Invoice-number normalization: confirmed real patterns, gated by amount+tax --\n");

  const yearFormatPair = run([[GSTIN_A, "133/21-22", 0, 900, 900, 10000]], [[GSTIN_A, "133/2021-22", 0, 900, 900, 10000]]);
  assert(yearFormatPair.extraIn2A.length === 0 && yearFormatPair.extraInBooks.length === 0, "'133/21-22' vs '133/2021-22' -- year format difference, matching amount+tax -- pairs up");

  const truncatedSuffixPair = run([[GSTIN_A, "S-008", 0, 500, 500, 6000]], [[GSTIN_A, "S-008/2021-22", 0, 500, 500, 6000]]);
  assert(truncatedSuffixPair.extraIn2A.length === 0 && truncatedSuffixPair.extraInBooks.length === 0, "'S-008' vs 'S-008/2021-22' -- truncated suffix, matching amount+tax -- pairs up");

  const droppedDigitPair = run([[GSTIN_A, "5093/21-22", 0, 100449.45, 100449.45, 1116105]], [[GSTIN_A, "5093/21-2", 0, 100449.45, 100449.45, 1116105]]);
  assert(droppedDigitPair.extraIn2A.length === 0 && droppedDigitPair.extraInBooks.length === 0, "'5093/21-22' vs '5093/21-2' -- dropped trailing digit, matching amount+tax -- pairs up (the real confirmed case)");

  const yearTypoPair = run([[GSTIN_A, "5264/21-21", 0, 14508.9, 14508.9, 161210]], [[GSTIN_A, "5264/21-22", 0, 14508.9, 14508.9, 161210]]);
  assert(yearTypoPair.extraIn2A.length === 0 && yearTypoPair.extraInBooks.length === 0, "'5264/21-21' vs '5264/21-22' -- year typo (wrong value, not just format), matching amount+tax -- pairs up per the confirmed design decision");

  const differentFyPair = run([[GSTIN_A, "5024/2020-21", 0, 347.4, 347.4, 3860]], [[GSTIN_A, "5024/21-22", 0, 347.4, 347.4, 3860]]);
  assert(differentFyPair.extraIn2A.length === 0 && differentFyPair.extraInBooks.length === 0, "'5024/2020-21' vs '5024/21-22' -- entirely different financial year, matching amount+tax -- pairs up (real case: same running number, wrong year on one side)");

  console.log("\n-- Invoice-number normalization: the critical guardrail -- string similarity is NEVER enough on its own --\n");

  const stringMatchesAmountDoesNot = run([[GSTIN_A, "133/21-22", 0, 900, 900, 10000]], [[GSTIN_A, "133/2021-22", 0, 999, 999, 20000]]);
  assert(
    stringMatchesAmountDoesNot.extraIn2A.length === 1 && stringMatchesAmountDoesNot.extraInBooks.length === 1,
    "'133/21-22' vs '133/2021-22' normalizes to a clean string match, but the amounts genuinely differ -- NOT paired, exactly the guardrail this fix requires (string similarity alone is never sufficient)"
  );

  console.log("\n-- Invoice-number normalization: confirmed digit-typo pairs must NOT merge, even with matching amounts --\n");

  const digitSubstitutionTypo = run([[GSTIN_A, "944", 0, 90, 90, 1000]], [[GSTIN_A, "994", 0, 90, 90, 1000]]);
  assert(digitSubstitutionTypo.extraIn2A.length === 1 && digitSubstitutionTypo.extraInBooks.length === 1, "'944' vs '994' -- same amount, but a plain digit substitution in the serial number -- correctly stays unmatched");

  const digitTranspositionTypo = run([[GSTIN_A, "7081/21-22", 0, 43991.46, 43991.46, 488794]], [[GSTIN_A, "8071/21-22", 0, 43991.46, 43991.46, 488794]]);
  assert(
    digitTranspositionTypo.extraIn2A.length === 1 && digitTranspositionTypo.extraInBooks.length === 1,
    "'7081/21-22' vs '8071/21-22' -- the real confirmed case: same year format, same amount to the rupee, but a digit transposition in the running number -- correctly stays unmatched"
  );

  const singleDigitOffTypo = run([[GSTIN_A, "7239/21-22", 0, 26186.94, 26186.94, 290966]], [[GSTIN_A, "7238/21-22", 0, 26186.94, 26186.94, 290966]]);
  assert(
    singleDigitOffTypo.extraIn2A.length === 1 && singleDigitOffTypo.extraInBooks.length === 1,
    "'7239/21-22' vs '7238/21-22' -- a single-digit-off serial number, same species of risk as the transposition case above -- correctly stays unmatched even though amounts agree"
  );

  console.log("\n-- Invoice-number normalization: genuine ambiguity is not guessed --\n");

  const ambiguousNormalization = run(
    [
      [GSTIN_A, "133/21-22", 0, 900, 900, 10000],
      [GSTIN_A, "133/2021-23", 0, 900, 900, 10000], // deliberately a different (if oddly-formed) identifier, same GSTIN+amount -- creates a second candidate
    ],
    [[GSTIN_A, "133/2021-22", 0, 900, 900, 10000]]
  );
  assert(
    ambiguousNormalization.extraIn2A.length === 2 && ambiguousNormalization.extraInBooks.length === 1,
    "two 2A candidates both plausibly match the one Books candidate on string+amount -- ambiguous, none guessed, all 3 stay Extra"
  );

  console.log("\n-- Possible match (Fix 4): blank invoice number, confirmed real case --\n");

  // GSTIN 07AAACS5547H2Z9: Books is a Journal voucher with a BLANK
  // Supplier Invoice No. field -- buildInvoiceIndex skips blank
  // identifiers entirely, so without this tier the row would never even
  // reach extraInBooks, let alone get a chance to match.
  const blankBooksIdentifier = runWithDate([[GSTIN_A, "DE1MP2122273354", 12965.95, 0, 0, 72033.05, "2021-11-10"]], [[GSTIN_A, "", 12965.95, 0, 0, 72033.05, "2021-11-15"]]);
  assert(blankBooksIdentifier.applicable, "still applicable overall -- a blank identifier on ONE row doesn't make the whole sheet inapplicable");
  assert(blankBooksIdentifier.extraIn2A.length === 0 && blankBooksIdentifier.extraInBooks.length === 0, "matched via the possible-match fallback -- neither side reported as a confirmed Extra");
  assert(blankBooksIdentifier.possibleMatches.length === 1, "the pair is surfaced as its own 'possible match' category instead");
  assert(blankBooksIdentifier.possibleMatches[0].reason === "Books invoice/voucher number is blank", "reason correctly identifies which side is blank");
  assert(blankBooksIdentifier.possibleMatches[0].gstr2aIdentifier === "DE1MP2122273354" && blankBooksIdentifier.possibleMatches[0].booksIdentifier === "", "both sides' identifiers (including the blank one) are carried through for the reviewer to see");

  console.log("\n-- Possible match (Fix 4): completely different reference numbers, confirmed real case --\n");

  // GSTIN 36AAOFA8281B1ZE: 2A invoice "2018" vs Books invoice
  // "1997"/voucher 403 -- no normalization path exists between these
  // strings at all, but the amounts agree within a few paise.
  const totallyDifferentReferences = runWithDate([[GSTIN_A, "2018", 0, 14916.43, 14916.43, 165738.16, "2021-09-05"]], [[GSTIN_A, "1997", 0, 14916.38, 14916.38, 165737.5, "2021-09-20"]]);
  assert(totallyDifferentReferences.extraIn2A.length === 0 && totallyDifferentReferences.extraInBooks.length === 0, "'2018' vs '1997' -- no string relationship at all, but amount+tax agree within ₹1 and both fall in the same month -- possible match");
  assert(totallyDifferentReferences.possibleMatches.length === 1 && totallyDifferentReferences.possibleMatches[0].reason.includes("2018") && totallyDifferentReferences.possibleMatches[0].reason.includes("1997"), "reason names both mismatched identifiers");

  console.log("\n-- Possible match (Fix 4): the guardrail -- BOTH amount+tax AND same-month date required --\n");

  const sameAmountDifferentMonth = runWithDate([[GSTIN_A, "9001", 0, 500, 500, 5000, "2021-01-05"]], [[GSTIN_A, "9002", 0, 500, 500, 5000, "2021-06-25"]]);
  assert(
    sameAmountDifferentMonth.extraIn2A.length === 1 && sameAmountDifferentMonth.extraInBooks.length === 1,
    "same amount+tax, but 5 months apart -- NOT a possible match, stays Extra on both sides (date window is a real requirement, not decorative)"
  );

  const sameMonthDifferentAmount = runWithDate([[GSTIN_A, "9003", 0, 500, 500, 5000, "2021-01-05"]], [[GSTIN_A, "9004", 0, 700, 700, 9000, "2021-01-20"]]);
  assert(
    sameMonthDifferentAmount.extraIn2A.length === 1 && sameMonthDifferentAmount.extraInBooks.length === 1,
    "same month, but genuinely different amounts -- NOT a possible match, stays Extra on both sides"
  );

  console.log("\n-- Possible match (Fix 4): genuine ambiguity is not guessed --\n");

  const ambiguousPossibleMatch = runWithDate(
    [
      [GSTIN_A, "AAA", 0, 500, 500, 5000, "2021-04-10"],
      [GSTIN_A, "BBB", 0, 500, 500, 5000, "2021-04-12"], // same GSTIN, same amount, same month as the other 2A row too -- creates real ambiguity
    ],
    [[GSTIN_A, "CCC", 0, 500, 500, 5000, "2021-04-15"]]
  );
  assert(
    ambiguousPossibleMatch.extraIn2A.length === 2 && ambiguousPossibleMatch.extraInBooks.length === 1,
    "two 2A candidates both plausibly match the one Books candidate on amount+month -- ambiguous, none guessed, all 3 stay Extra"
  );

  console.log("\n-- Possible match (Fix 4): must NOT run before Fixes 1-3 -- a cleanly-normalizable pair is resolved by Fix 3, not demoted to 'possible match' --\n");

  const shouldResolveViaNormalizationNotPossibleMatch = run([[GSTIN_A, "133/21-22", 0, 900, 900, 10000]], [[GSTIN_A, "133/2021-22", 0, 900, 900, 10000]]);
  assert(
    shouldResolveViaNormalizationNotPossibleMatch.possibleMatches.length === 0,
    "'133/21-22' vs '133/2021-22' is cleanly resolved by invoice-number normalization (Fix 3) -- it must NOT also show up as a lower-confidence 'possible match', since Fix 3 already resolved it with much higher confidence"
  );

  console.log("\n-- Possible match (Fix 4): a genuinely one-sided blank-identifier row is not silently reported as Extra --\n");

  const unmatchedBlankIdentifier = runWithDate([], [[GSTIN_A, "", 0, 500, 500, 5000, "2021-04-10"]]);
  assert(
    unmatchedBlankIdentifier.extraInBooks.length === 0 && unmatchedBlankIdentifier.possibleMatches.length === 0,
    "a blank-identifier row with no possible counterpart on the other side is left out of scope entirely (deliberately not newly surfaced as Extra by this fix -- see module docstring)"
  );

  console.log("\n-- Sheets without the needed columns --\n");

  const noIdentifierColumn = detectInvoiceLevelExtras(
    { values: [["GSTIN"], [GSTIN_A]], headerRowIndex: 0, columns: { gstin: 0, invoiceNumber: null, voucherNumber: null, igst: null, cgst: null, sgst: null, taxableValue: null, placeOfSupply: null } },
    { values: [BOOKS_HEADER, [GSTIN_A, "V-1", 0, 100, 100, 1000]], headerRowIndex: 0, columns: BOOKS_COLUMNS }
  );
  assert(noIdentifierColumn.applicable === false, "2A sheet with no invoice/voucher number column identified -> not applicable, not silently empty");

  console.log("");
}

function run_() {
  runTests();
  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All extra-invoice detector checks passed.");
  }
}

run_();
