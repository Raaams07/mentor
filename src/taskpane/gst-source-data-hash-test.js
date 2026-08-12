/*
 * gst-source-data-hash-test.js
 * -----------------------------------
 * Unit tests for gst-source-data-hash.js, using only fictional synthetic
 * data.
 *
 * Run with: node src/taskpane/gst-source-data-hash-test.js
 */

const { computeGstSourceDataHash } = require("./gst-source-data-hash.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

function runTests() {
  console.log("-- Determinism: same input -> same hash, every time --\n");

  const gstr2a = [
    ["GSTIN", "Invoice Number", "IGST"],
    ["27ZZAPL5432Q1Z9", "INV-001", 1000],
  ];
  const books = [
    ["GSTIN", "Voucher No.", "IGST"],
    ["27ZZAPL5432Q1Z9", "INV-001", 1000],
  ];
  const hash1 = computeGstSourceDataHash(gstr2a, books);
  const hash2 = computeGstSourceDataHash(gstr2a, books);
  assert(hash1 === hash2, "identical input produces the identical hash across repeated calls");

  console.log("\n-- Sensitivity: ANY cell change changes the hash --\n");

  const gstr2aEdited = [
    ["GSTIN", "Invoice Number", "IGST"],
    ["27ZZAPL5432Q1Z9", "INV-001", 1001], // one rupee different
  ];
  assert(computeGstSourceDataHash(gstr2aEdited, books) !== hash1, "a single changed cell value changes the hash");

  const booksEdited = [
    ["GSTIN", "Voucher No.", "IGST"],
    ["27ZZAPL5432Q1Z9", "INV-002", 1000], // Books-side change only
  ];
  assert(computeGstSourceDataHash(gstr2a, booksEdited) !== hash1, "a change on the Books side alone (2A unchanged) still changes the combined hash");

  console.log("\n-- New rows change the hash --\n");

  const gstr2aWithNewRow = gstr2a.concat([["27ZZAPL5432Q1Z9", "INV-002", 500]]);
  assert(computeGstSourceDataHash(gstr2aWithNewRow, books) !== hash1, "an added row changes the hash");

  console.log("\n-- Column-boundary collisions: no separator would make these collide, WITH separator they don't --\n");

  const rowSplitA = [["AB", "C"]];
  const rowSplitB = [["A", "BC"]];
  assert(computeGstSourceDataHash(rowSplitA, []) !== computeGstSourceDataHash(rowSplitB, []), "['AB','C'] and ['A','BC'] must NOT hash the same just because naive concatenation would read identically");

  console.log("\n-- Empty sheets --\n");

  assert(computeGstSourceDataHash([], []) === computeGstSourceDataHash([], []), "two empty inputs still hash deterministically (no crash, no NaN)");
  assert(typeof computeGstSourceDataHash([], []) === "string" && computeGstSourceDataHash([], []).length > 0, "empty input still produces a valid non-empty hash string");

  console.log("");
}

function run() {
  runTests();
  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All GST source-data hash checks passed.");
  }
}

run();
