/*
 * gst-cross-sheet-column-search-test.js
 * -----------------------------------------
 * Proves findCrossSheetColumnMatch()'s three outcomes:
 *   1. A single OTHER same-role sheet with an unambiguous match for the
 *      field, whose header exactly matches one of THIS sheet's candidates
 *      -> silent "exact" resolution.
 *   2. No exact agreement, but a plausibly-similar header on another sheet
 *      -> "fuzzy" — a pre-fill SUGGESTION only, never silent.
 *   3. Two corroborating sheets disagreeing with each other -> never
 *      resolves silently, even if one of them would have matched exactly
 *      alone.
 * Plus the "nothing to corroborate with" cases: no other sheets, and an
 * other sheet where the field is itself ambiguous or absent.
 *
 * All data is synthetic/fictional, constructed to exercise the matching
 * rules — not sourced from any real company or real client file.
 *
 * Run with: node src/gst-reconciliation/gst-cross-sheet-column-search-test.js
 */

const { findCrossSheetColumnMatch, SUGGESTION_SIMILARITY_THRESHOLD } = require("./gst-cross-sheet-column-search.js");
const { identifyGstColumns } = require("./gst-column-identifier.js");
const { extractSheetSignals } = require("../sheet-classifier/signal-extractor.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

// Mirrors the shape roleResults[sheetName] actually carries (gst-
// reconciliation.js's recognizeGstSheets): { sheetName, columns, sheetSignals }.
function buildOtherSheetResult(sheetName, values) {
  const sheetSignals = extractSheetSignals(sheetName, values);
  const columns = identifyGstColumns(sheetSignals, values, sheetSignals.headerRowIndex);
  return { sheetName, columns, sheetSignals };
}

function runExactAgreement() {
  console.log("-- A single corroborating sheet, exact normalized-header match -> silent resolution --\n");

  // The sheet being resolved: "CGST Amount" and "CGST @9%" both look like
  // plausible CGST columns (both numeric, both header-rule matches) — this
  // is what makes cgst ambiguous in the first place.
  const candidateHeaders = [
    { index: 2, header: "CGST Amount" },
    { index: 3, header: "CGST @9%" },
  ];

  // Another purchase-register-role sheet in the SAME workbook (say, last
  // month's tab) has only one column matching the cgst header rule at all —
  // unambiguous on ITS OWN shape.
  const other = buildOtherSheetResult("Purchase Register - Nov", [
    ["GSTIN", "Taxable Value", "CGST Amount", "SGST Amount"],
    ["29AAAPL2356Q1Z8", 10000, 900, 900],
  ]);

  const result = findCrossSheetColumnMatch("cgst", candidateHeaders, [other]);
  assert(result.matchType === "exact", "resolves as an exact match");
  assert(result.matchedHeader === "CGST Amount", "picks the candidate whose normalized header matches the corroborating sheet, not the other candidate");
  assert(result.sourceSheetName === "Purchase Register - Nov", "reports which sheet corroborated it");

  console.log("");
}

function runFuzzySuggestionOnly() {
  console.log("-- Corroborating sheet's header is similar but not identical -> fuzzy SUGGESTION, not silent --\n");

  const candidateHeaders = [
    { index: 2, header: "CGST Amount" },
    { index: 3, header: "CGST @9%" },
  ];

  // "CGST Amt" is close in spelling to "CGST Amount" (one candidate) but
  // doesn't normalize to an exact match against it or the other candidate.
  const other = buildOtherSheetResult("Purchase Register - Nov", [
    ["GSTIN", "Taxable Value", "CGST Amt", "SGST Amount"],
    ["29AAAPL2356Q1Z8", 10000, 900, 900],
  ]);

  const result = findCrossSheetColumnMatch("cgst", candidateHeaders, [other]);
  assert(result.matchType === "fuzzy", "close-but-not-identical spelling surfaces as a fuzzy suggestion, never resolves silently");
  assert(result.matchedHeader === "CGST Amount", "suggests the candidate closest in spelling to the corroborating sheet's header, not the other one");
  assert(result.similarity >= SUGGESTION_SIMILARITY_THRESHOLD, "a surfaced fuzzy suggestion always clears the similarity threshold");

  console.log("");
}

function runDisagreementNeverResolvesSilently() {
  console.log("-- Two corroborating sheets disagree with each other -> never resolves silently, even though one alone would have matched exactly --\n");

  const candidateHeaders = [
    { index: 2, header: "CGST Amount" },
    { index: 3, header: "CGST @9%" },
  ];

  const agreesWithFirst = buildOtherSheetResult("Purchase Register - Nov", [
    ["GSTIN", "Taxable Value", "CGST Amount", "SGST Amount"],
    ["29AAAPL2356Q1Z8", 10000, 900, 900],
  ]);
  const agreesWithSecond = buildOtherSheetResult("Purchase Register - Oct", [
    ["GSTIN", "Taxable Value", "CGST @9%", "SGST Amount"],
    ["29AAAPL2356Q1Z8", 10000, 900, 900],
  ]);

  const result = findCrossSheetColumnMatch("cgst", candidateHeaders, [agreesWithFirst, agreesWithSecond]);
  assert(result.matchType !== "exact", "disagreement between corroborating sheets blocks the exact/silent path entirely");

  console.log("");
}

function runNothingToCorroborateWith() {
  console.log("-- No signal to corroborate with -> matchType: none --\n");

  const candidateHeaders = [
    { index: 2, header: "CGST Amount" },
    { index: 3, header: "CGST @9%" },
  ];

  const noOtherSheets = findCrossSheetColumnMatch("cgst", candidateHeaders, []);
  assert(noOtherSheets.matchType === "none", "no other sheets at all -> none");

  // Another sheet exists, but THIS field is itself ambiguous there too (two
  // cgst-looking columns) -- nothing unambiguous to corroborate with.
  const ambiguousElsewhereToo = buildOtherSheetResult("Purchase Register - Oct", [
    ["GSTIN", "Taxable Value", "CGST Amount", "CGST @9%"],
    ["29AAAPL2356Q1Z8", 10000, 900, 900],
  ]);
  const resultAmbiguousElsewhere = findCrossSheetColumnMatch("cgst", candidateHeaders, [ambiguousElsewhereToo]);
  assert(resultAmbiguousElsewhere.matchType === "none", "the other sheet having the SAME field ambiguous on its own shape contributes nothing -- correctly skipped, not guessed from");

  // Another sheet exists, but has no column matching this field's header
  // rule at all.
  const absentElsewhere = buildOtherSheetResult("Purchase Register - Oct", [
    ["GSTIN", "Taxable Value", "SGST Amount"],
    ["29AAAPL2356Q1Z8", 10000, 900],
  ]);
  const resultAbsentElsewhere = findCrossSheetColumnMatch("cgst", candidateHeaders, [absentElsewhere]);
  assert(resultAbsentElsewhere.matchType === "none", "the other sheet having zero candidates for this field contributes nothing");

  console.log("");
}

function run() {
  runExactAgreement();
  runFuzzySuggestionOnly();
  runDisagreementNeverResolvesSilently();
  runNothingToCorroborateWith();

  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All gst-cross-sheet-column-search checks passed.");
  }
}

run();
