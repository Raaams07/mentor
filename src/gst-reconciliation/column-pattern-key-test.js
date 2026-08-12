/*
 * column-pattern-key-test.js
 * -----------------------------
 * Proves computeColumnPatternKey()/normalizeCandidateHeaders() are
 * deterministic regardless of input order/casing/whitespace, dedupe
 * correctly, and produce a different key per field name — the properties
 * that make the key a stable signature of "which software export shape is
 * this" rather than of any one client's specific workbook.
 *
 * Run with: node src/gst-reconciliation/column-pattern-key-test.js
 */

const { computeColumnPatternKey, normalizeCandidateHeaders } = require("./column-pattern-key.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

function runOrderIndependence() {
  console.log("-- Order-independence --\n");
  const a = computeColumnPatternKey("cgst", ["Input CGST@14%", "CGST", "Output CGST @ 9%"]);
  const b = computeColumnPatternKey("cgst", ["CGST", "Output CGST @ 9%", "Input CGST@14%"]);
  assert(a === b, "the same header set in a different order produces the identical key");
  console.log("");
}

function runCasingWhitespaceNormalization() {
  console.log("-- Casing/whitespace normalization --\n");
  const a = computeColumnPatternKey("cgst", ["  CGST  ", "Central Tax"]);
  const b = computeColumnPatternKey("cgst", ["cgst", "central tax"]);
  assert(a === b, "casing and surrounding whitespace don't affect the key");
  console.log("");
}

function runDifferentFieldNameDifferentKey() {
  console.log("-- Different field name -> different key --\n");
  const headers = ["Input CGST@9%", "CGST"];
  const cgstKey = computeColumnPatternKey("cgst", headers);
  const sgstKey = computeColumnPatternKey("sgst", headers);
  assert(cgstKey !== sgstKey, "identical header list, different field name -> different key (field name is part of the signature)");
  console.log("");
}

function runDedup() {
  console.log("-- Dedup of headers that normalize to the same token --\n");
  const normalized = normalizeCandidateHeaders(["CGST", "cgst ", "  CGST"]);
  assert(normalized.length === 1 && normalized[0] === "cgst", "three variants of the same header collapse to a single normalized entry");

  const withBlanks = normalizeCandidateHeaders(["CGST", "", null, undefined, "SGST"]);
  assert(withBlanks.length === 2, "blank/null/undefined headers are dropped, not counted as candidates");

  console.log("");
}

function runTallyScenarioFromTheBrief() {
  console.log("-- The exact Tally scenario from the brief: stable, deterministic key --\n");
  const orderingA = ["Input CGST @ 9%", "Input CGST@14%", "Output CGST @ 9%", "CGST"];
  const orderingB = ["CGST", "Output CGST @ 9%", "Input CGST @ 9%", "Input CGST@14%"];
  const keyA = computeColumnPatternKey("cgst", orderingA);
  const keyB = computeColumnPatternKey("cgst", orderingB);
  assert(keyA === keyB, "the motivating example's 4-column ambiguity set produces a stable key across reordering");
  assert(keyA.startsWith("cgst::"), "the key is prefixed by the field name for readability/debuggability");
  console.log("  key:", keyA);
  console.log("");
}

function run() {
  runOrderIndependence();
  runCasingWhitespaceNormalization();
  runDifferentFieldNameDifferentKey();
  runDedup();
  runTallyScenarioFromTheBrief();

  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All column-pattern-key checks passed.");
  }
}

run();
