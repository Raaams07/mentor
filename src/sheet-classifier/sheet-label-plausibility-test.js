/*
 * sheet-label-plausibility-test.js
 * -----------------------------------
 * Run with: node src/sheet-classifier/sheet-label-plausibility-test.js
 */

const { checkLabelPlausibility } = require("./sheet-label-plausibility.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

function sheetSignalsWithHeaders(headerLabels) {
  return { headerLabels };
}

console.log("-- The real incident: 'Journal entries' on a GSTIN/CGST/SGST/IGST pivot table --\n");
{
  const result = checkLabelPlausibility({
    label: "Journal entries",
    sheetSignals: sheetSignalsWithHeaders(["GSTIN of supplier", "Taxable Value", "CGST", "SGST", "IGST"]),
  });
  assert(result.plausible === false, "'Journal entries' is flagged as implausible against GST headers — the exact case that motivated this check");
  assert(result.reason === "no_shared_vocabulary", "reason is no_shared_vocabulary");
}

console.log("\n-- A correct label for the same sheet --\n");
{
  const result = checkLabelPlausibility({
    label: "GST purchase register",
    sheetSignals: sheetSignalsWithHeaders(["GSTIN of supplier", "Taxable Value", "CGST", "SGST", "IGST"]),
  });
  assert(result.plausible === true, "'GST purchase register' is plausible against GST headers (GST ~ GSTIN)");
  assert(result.sharedTokens.includes("gst"), "shared token includes 'gst'");
}

console.log("\n-- A legitimate label that happens to share a header word --\n");
{
  const result = checkLabelPlausibility({
    label: "Stock levels",
    sheetSignals: sheetSignalsWithHeaders(["Item Code", "Stock Qty", "Warehouse"]),
  });
  assert(result.plausible === true, "'Stock levels' is plausible against Item/Stock Qty/Warehouse headers");
}

console.log("\n-- Plural/singular variants still match (loose stemming) --\n");
{
  const result = checkLabelPlausibility({
    label: "Invoice list",
    sheetSignals: sheetSignalsWithHeaders(["Invoices", "Date", "Amount"]),
  });
  assert(result.plausible === true, "'Invoice' matches 'Invoices' via loose stemming");
}

console.log("\n-- No header row detected — not enough signal, don't false-alarm --\n");
{
  const result = checkLabelPlausibility({
    label: "Journal entries",
    sheetSignals: sheetSignalsWithHeaders([]),
  });
  assert(result.plausible === true, "no headers to compare against -> plausible (fails open, not closed)");
  assert(result.reason === "not_enough_signal", "reason is not_enough_signal");
}

console.log("\n-- Very short/generic label — not enough signal, don't false-alarm --\n");
{
  const result = checkLabelPlausibility({
    label: "P&L",
    sheetSignals: sheetSignalsWithHeaders(["Account", "Debit", "Credit"]),
  });
  assert(result.plausible === true, "a label with no tokenizable words (all < 3 chars) doesn't get flagged");
}

console.log("\n-- A genuinely legitimate label with zero header overlap is still flagged (soft warning, expected false positive) --\n");
{
  const result = checkLabelPlausibility({
    label: "Marketing spend log",
    sheetSignals: sheetSignalsWithHeaders(["Date", "Campaign", "Amount"]),
  });
  assert(result.plausible === false, "flagged despite being a legitimate label — this is why the UI treats it as an overridable warning, not a block");
}

console.log("");
if (failures > 0) {
  console.log(`${failures} check(s) FAILED.`);
  process.exitCode = 1;
} else {
  console.log("All sheet-label-plausibility checks passed.");
}
