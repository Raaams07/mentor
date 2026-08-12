/*
 * gst-column-identifier-test.js
 * --------------------------------
 * Proves identifyGstColumns()'s additive `candidates` output stays a
 * strict superset of information — `result.field` (the existing single-
 * best-guess value every downstream rule and gst-role-recognizer.js
 * already depends on) is byte-identical to its pre-candidates behavior,
 * while `result.candidates.field` exposes every plausible match so
 * ambiguity can be detected later (gst-column-ambiguity-rules.js).
 *
 * The first scenario below is the actual real-world bug this whole
 * feature was built to catch: a Tally-style purchase register with
 * multiple similarly-named CGST columns, where the OLD first-match-wins
 * behavior silently picked the wrong one with zero signal anything was
 * wrong.
 *
 * All data is synthetic/fictional — made-up headers and figures
 * constructed to exercise the matching rules, not sourced from any real
 * company or real client file.
 *
 * Run with: node src/gst-reconciliation/gst-column-identifier-test.js
 */

const { identifyGstColumns, GSTIN_CONTENT_THRESHOLD } = require("./gst-column-identifier.js");
const { extractSheetSignals } = require("../sheet-classifier/signal-extractor.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

function identify(values) {
  const signals = extractSheetSignals("Sheet1", values);
  return identifyGstColumns(signals, values, signals.headerRowIndex);
}

function runSingleMatchStaysUnambiguous() {
  console.log("-- Single unambiguous match: unchanged behavior, single-entry candidates --\n");

  const values = [
    ["GSTIN", "Taxable Value", "CGST", "SGST"],
    ["29AAAPL2356Q1Z8", 10000, 900, 900],
  ];
  const result = identify(values);

  assert(result.cgst === 2, "single CGST column resolves exactly as before");
  assert(result.candidates.cgst.length === 1 && result.candidates.cgst[0] === 2, "candidates.cgst has exactly the one match");
  assert(result.candidates.sgst.length === 1 && result.candidates.sgst[0] === 3, "candidates.sgst has exactly the one match");

  console.log("");
}

function runTallyBugRegression() {
  console.log("-- Tally CGST bug regression: multiple candidates surface, but first-match-wins stays byte-identical --\n");

  // The exact real-world scenario: a Tally purchase register with
  // "Input CGST@9%", "Input CGST@14%", "Output CGST@9%" all present
  // ahead of the real final "CGST" column.
  const values = [
    ["GSTIN", "Taxable Value", "Input CGST@9%", "Input CGST@14%", "Output CGST@9%", "CGST"],
    ["29AAAPL2356Q1Z8", 10000, 900, 1400, 100, 900],
  ];
  const result = identify(values);

  assert(result.candidates.cgst.length === 4, "all 4 CGST-like columns are captured as candidates");
  assert(JSON.stringify(result.candidates.cgst) === JSON.stringify([2, 3, 4, 5]), "candidates.cgst lists every match in column order");
  assert(result.cgst === 2, "result.cgst is STILL the first match (col 2, 'Input CGST@9%') — proves role-recognition scoring behavior is completely unchanged by this feature");

  console.log("");
}

function runZeroMatch() {
  console.log("-- Zero-match fields: empty candidates, null result --\n");

  const values = [
    ["GSTIN", "Taxable Value", "IGST"],
    ["29AAAPL2356Q1Z8", 10000, 1800],
  ];
  const result = identify(values);

  assert(result.cgst === null, "no CGST column -> result.cgst stays null");
  assert(Array.isArray(result.candidates.cgst) && result.candidates.cgst.length === 0, "no CGST column -> candidates.cgst is an empty array, not undefined");
  assert(Array.isArray(result.candidates.placeOfSupply) && result.candidates.placeOfSupply.length === 0, "an entirely absent optional field also gets an empty (not missing) candidates array");

  console.log("");
}

function runGstinNearTie() {
  console.log("-- GSTIN candidates: content-ratio based, sorted descending, only >= threshold --\n");

  const values = [
    ["Recipient GSTIN", "Supplier GSTIN", "Taxable Value"],
    ["29AAAPL2356Q1Z8", "27AAACT2727Q1ZW", 10000],
    ["29AAAPL2356Q1Z8", "27AAACT2727Q1ZW", 12000],
  ];
  const result = identify(values);

  assert(result.candidates.gstin.length === 2, "both columns clear the content threshold and appear as candidates");
  assert(result.candidates.gstin[0].ratio >= result.candidates.gstin[1].ratio, "candidates.gstin is sorted descending by ratio");
  assert(result.candidates.gstin.every((c) => c.ratio >= GSTIN_CONTENT_THRESHOLD), "every gstin candidate clears GSTIN_CONTENT_THRESHOLD");

  const singleGstinValues = [
    ["GSTIN", "Some Other Column"],
    ["29AAAPL2356Q1Z8", "not a gstin"],
  ];
  const singleResult = identify(singleGstinValues);
  assert(singleResult.candidates.gstin.length === 1, "a single genuine GSTIN column -> exactly one candidate");

  console.log("");
}

function runNewRateField() {
  console.log("-- New 'rate' field --\n");

  const values = [
    ["GSTIN", "Taxable Value", "Rate (%)", "CGST", "SGST"],
    ["29AAAPL2356Q1Z8", 10000, 18, 900, 900],
  ];
  const result = identify(values);

  assert(result.rate === 2, "'Rate (%)' header identified as the rate column");
  assert(result.candidates.rate.length === 1 && result.candidates.rate[0] === 2, "rate candidates populated the same way as every other header-based field");

  const noRateValues = [
    ["GSTIN", "Taxable Value", "CGST", "SGST"],
    ["29AAAPL2356Q1Z8", 10000, 900, 900],
  ];
  const noRateResult = identify(noRateValues);
  assert(noRateResult.rate === null && noRateResult.candidates.rate.length === 0, "no rate column present -> null/empty, not an error");

  console.log("");
}

function run() {
  runSingleMatchStaysUnambiguous();
  runTallyBugRegression();
  runZeroMatch();
  runGstinNearTie();
  runNewRateField();

  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All gst-column-identifier checks passed.");
  }
}

run();
