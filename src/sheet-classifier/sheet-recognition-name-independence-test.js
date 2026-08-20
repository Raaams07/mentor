/*
 * sheet-recognition-name-independence-test.js
 * -----------------------------------------------
 * Regression test for a claim made explicitly to a user asking whether a
 * real workbook with a completely unfamiliar filename would still work:
 * MENTOR's sheet recognition — GST role scoring, general sheet-type
 * classification, and the "remembers this shape" structural fingerprint —
 * depends ONLY on a sheet's internal structure (column headers, cell-value
 * shape), never on the sheet's tab name or the workbook's .xlsx filename.
 *
 * Proves it by calling the real functions with IDENTICAL data under
 * DIFFERENT sheet names and asserting byte-identical results, rather than
 * just inspecting the source for the absence of a filename parameter —
 * this would catch a future regression where someone "helpfully" threads
 * sheetName into a scoring rule as a tie-breaker.
 *
 * (Workbook filename itself — as opposed to sheet tab name — is a
 * separate, deliberately different concern: it's used ONLY to scope
 * learned-answer memory per client, e.g. mentor-sheet-memory-ui.js's
 * mentorGetSheetMemoryClientId(context.workbook.name). That's not
 * recognition, and a never-before-seen filename is EXPECTED to start with
 * an empty memory scope — it doesn't affect whether a sheet is recognized,
 * only whether an answer is already remembered for it.)
 *
 * Run with: node src/sheet-classifier/sheet-recognition-name-independence-test.js
 */

const { extractSheetSignals } = require("./signal-extractor.js");
const { classifySheet } = require("./classifier.js");
const { computeStructuralSignature } = require("./structural-signature.js");
const { recognizeGstRole } = require("../gst-reconciliation/gst-role-recognizer.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

// A realistic GSTR-2A-shaped sheet — the SAME values are reused under two
// wildly different sheet names below. Nothing about the data ever changes;
// only the name does.
const GSTR2A_VALUES = [
  ["GSTIN of supplier", "Trade/Legal Name", "Invoice Number", "Place of Supply", "Taxable Value", "IGST", "CGST", "SGST"],
  ["11AAAAA0000A1Z1", "Vendor A Traders", "INV-1001", "Karnataka", 100000, 18000, 0, 0],
  ["22BBBBB1111B1Z2", "Vendor B Enterprises", "INV-1002", "Maharashtra", 200000, 0, 18000, 18000],
  ["33CCCCC2222C1Z3", "Vendor C Industries", "INV-1003", "Telangana", 150000, 27000, 0, 0],
  ["44DDDDD3333D1Z4", "Vendor D Tube Co.", "INV-1004", "Karnataka", 75000, 13500, 0, 0],
  ["55EEEEE4444E1Z5", "Vendor E Traders", "INV-1005", "Gujarat", 50000, 9000, 0, 0],
];

const FAMILIAR_LOOKING_NAME = "2A";
const NEVER_SEEN_SHEET_NAME = "Sheet7 (a name MENTOR has never encountered before)";

console.log("-- Role recognition (gst-role-recognizer.js) is identical regardless of sheet name --\n");
{
  const signalsFamiliar = extractSheetSignals(FAMILIAR_LOOKING_NAME, GSTR2A_VALUES, { headerRowIndex: 0 });
  const signalsUnfamiliar = extractSheetSignals(NEVER_SEEN_SHEET_NAME, GSTR2A_VALUES, { headerRowIndex: 0 });

  const resultFamiliar = recognizeGstRole(signalsFamiliar, GSTR2A_VALUES, 0);
  const resultUnfamiliar = recognizeGstRole(signalsUnfamiliar, GSTR2A_VALUES, 0);

  assert(resultFamiliar.role === "gstr2a", "sheet named '2A' is recognized as gstr2a (sanity check the fixture itself is valid)");
  assert(resultUnfamiliar.role === "gstr2a", "the SAME data under a completely unfamiliar sheet name is STILL recognized as gstr2a");
  assert(resultFamiliar.confidence === "high", "confidence is 'high', not just a marginal pass (sanity check the fixture is unambiguous)");
  assert(resultFamiliar.role === resultUnfamiliar.role, "role is identical regardless of sheet name");
  assert(resultFamiliar.confidence === resultUnfamiliar.confidence, "confidence is identical regardless of sheet name");
  assert(JSON.stringify(resultFamiliar.scores) === JSON.stringify(resultUnfamiliar.scores), "every role's score is identical regardless of sheet name");
}

console.log("\n-- Sheet-type classification (classifier.js) never lets sheetName affect scoring --\n");
{
  const signalsFamiliar = extractSheetSignals(FAMILIAR_LOOKING_NAME, GSTR2A_VALUES, { headerRowIndex: 0 });
  const signalsUnfamiliar = extractSheetSignals(NEVER_SEEN_SHEET_NAME, GSTR2A_VALUES, { headerRowIndex: 0 });

  const classifiedFamiliar = classifySheet(signalsFamiliar);
  const classifiedUnfamiliar = classifySheet(signalsUnfamiliar);

  assert(classifiedFamiliar.sheetName === FAMILIAR_LOOKING_NAME, "sheetName is carried through into the result for display purposes...");
  assert(classifiedUnfamiliar.sheetName === NEVER_SEEN_SHEET_NAME, "...and reflects whatever name was actually passed in, per sheet");
  assert(classifiedFamiliar.type === classifiedUnfamiliar.type, "classified TYPE is identical regardless of sheet name — sheetName is a passthrough field, not a scoring input");
  assert(classifiedFamiliar.confidence === classifiedUnfamiliar.confidence, "confidence is identical regardless of sheet name");
  assert(JSON.stringify(classifiedFamiliar.scores) === JSON.stringify(classifiedUnfamiliar.scores), "every type's score is identical regardless of sheet name");
}

console.log("\n-- The 'remembers this shape' fingerprint (structural-signature.js) survives a rename, per its own docstring's claim --\n");
{
  const signalsFamiliar = extractSheetSignals(FAMILIAR_LOOKING_NAME, GSTR2A_VALUES, { headerRowIndex: 0 });
  const signalsUnfamiliar = extractSheetSignals(NEVER_SEEN_SHEET_NAME, GSTR2A_VALUES, { headerRowIndex: 0 });

  const sigFamiliar = computeStructuralSignature(signalsFamiliar);
  const sigUnfamiliar = computeStructuralSignature(signalsUnfamiliar);

  assert(sigFamiliar.signature === sigUnfamiliar.signature, "structural signature is byte-identical regardless of sheet name — proves 'same shape, different name' actually holds, not just claimed in the docstring");
  assert(sigFamiliar.headerSignature === sigUnfamiliar.headerSignature, "header signature is byte-identical regardless of sheet name");
}

console.log("\n-- The exact practical claim a real user would care about --\n");
{
  // Not a familiar-sounding name like "2A", not a clean name at all — the
  // kind of messy, auto-exported, never-seen-before name a real download
  // actually has. Recognition must not treat this as a weaker match than
  // a MENTOR-familiar-looking name.
  const messyName = "gstr2a_download_FINAL_v3(1) - Copy.xlsx — Sheet1";
  const signals = extractSheetSignals(messyName, GSTR2A_VALUES, { headerRowIndex: 0 });
  const result = recognizeGstRole(signals, GSTR2A_VALUES, 0);
  assert(result.role === "gstr2a" && result.confidence === "high", "a deliberately messy, never-before-seen sheet name recognizes with the exact same confidence as a familiar-looking one");
}

console.log("");
if (failures > 0) {
  console.log(`${failures} check(s) FAILED.`);
  process.exitCode = 1;
} else {
  console.log("All sheet-recognition name-independence checks passed.");
}
