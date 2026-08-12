/*
 * invoice-number-normalizer-test.js
 * --------------------------------------
 * Unit tests for invoice-number-normalizer.js, using only fictional
 * synthetic identifiers plus the EXACT real-world patterns confirmed
 * against Basai Steel Traders (values only — the identifiers/patterns
 * themselves are already documented in the module's own docstring and
 * this session's bug report, not sensitive).
 *
 * Run with: node src/gst-reconciliation/invoice-number-normalizer-test.js
 */

const { normalizeInvoiceNumber, invoiceNumbersMatch } = require("./invoice-number-normalizer.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

function runTests() {
  console.log("-- normalizeInvoiceNumber(): format normalization --\n");

  assert(normalizeInvoiceNumber("  inv-100  ") === "INV-100", "whitespace stripped, uppercased");
  assert(normalizeInvoiceNumber("133/21-22") === normalizeInvoiceNumber("133/2021-22"), "'133/21-22' and '133/2021-22' normalize identically -- year format unified");
  assert(normalizeInvoiceNumber("133/21-22") === "133-2021-22", "the 2-digit year pair is expanded to 4 digits, not left as-is");
  assert(normalizeInvoiceNumber("STPL/1181") === normalizeInvoiceNumber("STPL//1181"), "a doubled separator ('//') collapses the same as a single one");
  assert(normalizeInvoiceNumber("BLY/21-22-280") === normalizeInvoiceNumber("BLY/21-22/280"), "'-' and '/' are interchangeable as separators");

  console.log("\n-- normalizeInvoiceNumber(): does NOT touch a non-consecutive pair --\n");

  assert(normalizeInvoiceNumber("5264-21-21") === "5264-21-21", "'21-21' is not a consecutive FY pair (22 != 21+1) -- left unexpanded by the strict normalizer");
  assert(normalizeInvoiceNumber("944") === "944", "a bare number with no separators is unaffected");

  console.log("\n-- invoiceNumbersMatch(): confirmed real patterns (Basai Steel Traders) --\n");

  assert(invoiceNumbersMatch("133/21-22", "133/2021-22"), "year format difference -- matches via normalizeInvoiceNumber alone");
  assert(invoiceNumbersMatch("S-008", "S-008/2021-22"), "truncated suffix -- the shorter identifier's core is a prefix of the longer one's, once the year is stripped from both");
  assert(invoiceNumbersMatch("5093/21-2", "5093/21-22"), "dropped trailing digit within the year -- the running number '5093' still matches exactly once each side's year-shape is stripped");
  assert(invoiceNumbersMatch("7360/21-2", "7360/21-22"), "same dropped-trailing-digit pattern, a second real example");
  assert(invoiceNumbersMatch("BLY/21-22-280", "BLY/21-22/280"), "punctuation variant -- matches via normalizeInvoiceNumber (exact, no core fallback needed)");
  assert(invoiceNumbersMatch("STPL/1181", "STPL//1181"), "doubled-separator typo -- matches via normalizeInvoiceNumber");

  console.log("\n-- invoiceNumbersMatch(): year VALUE differs (typo), not just format -- still matches --\n");

  assert(invoiceNumbersMatch("5264/21-21", "5264/21-22"), "year typo ('21-21' vs '21-22') -- the running number '5264' matches exactly; the year segment is recognized by SHAPE and its VALUE is not compared");
  assert(invoiceNumbersMatch("5024/2020-21", "5024/21-22"), "different financial years entirely (2020-21 vs 2021-22) -- still matches on the running number '5024' alone, confirmed against real data (identical taxable value/tax on both sides)");

  console.log("\n-- invoiceNumbersMatch(): digit typos in the SERIAL number are NOT matched (the critical guardrail) --\n");

  assert(!invoiceNumbersMatch("944", "994"), "944 vs 994 -- a plain digit substitution in a bare serial number, never matched");
  assert(!invoiceNumbersMatch("8071/21-22", "7081/21-22"), "8071 vs 7081 -- a digit transposition in the running number, SAME year format on both sides -- still never matched, confirmed against the real case this guardrail exists for");
  assert(!invoiceNumbersMatch("7239/21-22", "7238/21-22"), "7239 vs 7238 -- a single-digit difference in the running number is structurally identical risk to 8071/7081 and is treated the same way, even though both sides' amounts happened to agree in the real file");

  console.log("\n-- invoiceNumbersMatch(): the 4-digit-serial-mistaken-for-a-year bug this file specifically guards against --\n");

  // "7081" and "8071" are both 4 digits -- if isYearShapeSegment used a
  // blanket \d{4} check (no realistic range restriction), BOTH would be
  // misread as a "year" and stripped from the core, leaving both sides
  // with the SAME leftover core and incorrectly matching. The 2000-2099
  // range restriction is what prevents this.
  assert(!invoiceNumbersMatch("7081/21-22", "9999/21-22"), "two clearly-different 4-digit serials, neither in the 2000-2099 'year' range -- not matched just because both happen to be followed by a real year-shaped pair");

  console.log("\n-- invoiceNumbersMatch(): genuinely different invoices, no relationship at all --\n");

  assert(!invoiceNumbersMatch("217-9", "6417"), "no normalization path turns one into the other -- correctly left unmatched by string logic alone (a case that needs amount+date corroboration elsewhere, not invoice-number normalization)");
  assert(!invoiceNumbersMatch("INV-100", "BILL-200"), "completely unrelated identifiers");
  assert(!invoiceNumbersMatch("ABC-123", "XYZ-123"), "same numeric suffix, different prefix -- prefix must match exactly, not just the tail");

  console.log("\n-- invoiceNumbersMatch(): edge cases --\n");

  assert(invoiceNumbersMatch("INV-100", "INV-100"), "identical strings trivially match");
  assert(!invoiceNumbersMatch("21-22", "23-24"), "two DIFFERENT identifiers that are each ENTIRELY a year-shaped pair both reduce to an empty core -- never treated as a match, nothing left to distinguish either one");
  assert(!invoiceNumbersMatch("", "21-22"), "an empty identifier never matches anything");

  console.log("");
}

function run() {
  runTests();
  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All invoice-number normalizer checks passed.");
  }
}

run();
