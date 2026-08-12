/*
 * gst-column-resolution-integration-test.js
 * ---------------------------------------------
 * Composes identifyGstColumns() + fieldsInPlay() + resolveColumnField()/
 * rememberColumnField() end-to-end, pure (no Excel, no UI) — the same
 * pipeline mentor-column-memory-ui.js's mentorResolveGstColumns() drives
 * live, minus the Office.js plumbing.
 *
 * The headline scenario is the full Tally CGST bug, start to finish:
 * unresolved on first scan (with the real column among the candidates),
 * remembered once, resolved silently forever after — proving the whole
 * three-state model (confident / ambiguous-needs-input / absent) actually
 * closes the gap the bug report described.
 *
 * All data is synthetic/fictional.
 *
 * Run with: node src/gst-reconciliation/gst-column-resolution-integration-test.js
 */

const { identifyGstColumns } = require("./gst-column-identifier.js");
const { fieldsInPlay } = require("./gst-column-ambiguity-rules.js");
const { resolveColumnField, rememberColumnField } = require("./column-memory.js");
const { BaseSheetMemoryStore } = require("../sheet-classifier/sheet-memory-store-base.js");
const { extractSheetSignals } = require("../sheet-classifier/signal-extractor.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

const CLIENT_ID = "workbook:test.xlsx";

// end-to-end: values -> identifyGstColumns -> fieldsInPlay -> per-field
// resolveColumnField, exactly the shape mentorResolveGstColumns() drives.
async function resolveAll(sheetName, role, values, store) {
  const sheetSignals = extractSheetSignals(sheetName, values);
  const columns = identifyGstColumns(sheetSignals, values, sheetSignals.headerRowIndex);
  const inPlay = fieldsInPlay(role, columns.candidates);
  const allColumnIndices = sheetSignals.columns.map((c) => c.index);

  const resolved = {};
  const pending = [];
  for (const item of inPlay) {
    const result = await resolveColumnField({
      clientId: CLIENT_ID,
      sheetName,
      sheetSignals,
      fieldName: item.field,
      candidateIndices: item.candidateIndices,
      allColumnIndices,
      store,
    });
    if (result.status === "resolved_from_memory") resolved[item.field] = result.fieldIndex;
    else pending.push({ ...result, sheetSignals });
  }
  return { columns, inPlay, resolved, pending, sheetSignals };
}

function tallyLikeBooksValues() {
  return [
    ["GSTIN", "Voucher No.", "Taxable Value", "Input CGST@9%", "Input CGST@14%", "Output CGST@9%", "CGST", "SGST"],
    ["29AAAPL2356Q1Z8", "PV-101", 10000, 900, 1400, 100, 900, 900],
  ];
}

async function runFullTallyBugRegression() {
  console.log("-- Full Tally CGST bug, end-to-end: unresolved -> remembered -> resolved silently --\n");
  const store = new BaseSheetMemoryStore();
  const values = tallyLikeBooksValues();

  const firstScan = await resolveAll("Books", "purchase_register", values, store);
  const cgstPending = firstScan.pending.find((p) => p.fieldName === "cgst");
  assert(!!cgstPending, "on first scan, cgst is genuinely unresolved (needs_input), not silently guessed");
  assert(cgstPending.candidates.length === 4, "the prompt offers all 4 CGST-like candidates");
  const realCgstCandidate = cgstPending.candidates.find((c) => c.header === "CGST");
  assert(!!realCgstCandidate && realCgstCandidate.index === 6, "the REAL final CGST column (index 6) is among the offered candidates, correctly identified by header text");

  // User answers once.
  await rememberColumnField({ clientId: CLIENT_ID, sheetName: "Books", sheetSignals: firstScan.sheetSignals, fieldName: "cgst", chosenColumnIndex: realCgstCandidate.index, store });

  // Re-run the exact same pipeline on the identical sheet.
  const secondScan = await resolveAll("Books", "purchase_register", values, store);
  assert(secondScan.pending.find((p) => p.fieldName === "cgst") === undefined, "second scan: cgst no longer pending");
  assert(secondScan.resolved.cgst === 6, "second scan: cgst resolves silently to the REMEMBERED real column, not the first-match guess (which would have been index 3)");

  console.log("");
}

async function runExactOneMatchNeverTouchesStore() {
  console.log("-- Exact-one-match never even reaches the store (fieldsInPlay excludes it entirely) --\n");
  const store = new BaseSheetMemoryStore();
  const values = [
    ["GSTIN", "Voucher No.", "Taxable Value", "CGST", "SGST"],
    ["29AAAPL2356Q1Z8", "PV-101", 10000, 900, 900],
  ];
  const scan = await resolveAll("Books", "purchase_register", values, store);
  assert(scan.inPlay.length === 0, "a fully unambiguous sheet has nothing in play at all");
  assert(scan.pending.length === 0 && Object.keys(scan.resolved).length === 0, "nothing pending, nothing resolved-from-memory -- the existing silent path is completely untouched");
  const records = await store.list(CLIENT_ID);
  assert(records.length === 0, "the store was never even queried");

  console.log("");
}

async function runZeroMatchOnRequiredFallsBackToAllColumns() {
  console.log("-- Zero-match on a required field falls back to offering every column on the sheet --\n");
  const store = new BaseSheetMemoryStore();
  // gstr2a-shaped sheet with genuinely no Invoice Number column anywhere.
  const values = [
    ["GSTIN of Supplier", "Trade/Legal Name", "Taxable Value", "Integrated Tax", "Central Tax", "State/UT Tax", "Place of Supply"],
    ["29AAAPL2356Q1Z8", "Vendor Alpha Pvt Ltd", 10000, 0, 900, 900, "Karnataka"],
  ];
  const scan = await resolveAll("2A", "gstr2a", values, store);
  const invoiceNumberPending = scan.pending.find((p) => p.fieldName === "invoiceNumber");
  assert(!!invoiceNumberPending, "invoiceNumber genuinely absent on a recognized gstr2a-shaped sheet -> flagged, not silently skipped");
  assert(invoiceNumberPending.candidates.length === 7, "with zero header-rule matches, the fallback offers every column in the sheet (7 columns here)");

  console.log("");
}

async function runZeroMatchOnLegitimatelyOptionalTaxFieldStaysSilent() {
  console.log("-- An IGST-only vendor register (cgst/sgst legitimately absent) stays silent, no ask --\n");
  const store = new BaseSheetMemoryStore();
  const values = [
    ["GSTIN", "Voucher No.", "Taxable Value", "IGST"],
    ["27AAACT2727Q1ZW", "PV-201", 20000, 3600],
  ];
  const scan = await resolveAll("Books", "purchase_register", values, store);
  assert(scan.pending.find((p) => p.fieldName === "cgst") === undefined && scan.pending.find((p) => p.fieldName === "sgst") === undefined, "cgst/sgst absent (this vendor is genuinely IGST-only) -> not flagged, not asked");

  console.log("");
}

async function run() {
  await runFullTallyBugRegression();
  await runExactOneMatchNeverTouchesStore();
  await runZeroMatchOnRequiredFallsBackToAllColumns();
  await runZeroMatchOnLegitimatelyOptionalTaxFieldStaysSilent();

  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All gst-column-resolution integration checks passed.");
  }
}

run();
