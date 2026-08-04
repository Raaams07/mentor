/*
 * ineligible-itc-vendor-extraction-test.js
 * -----------------------------------------------
 * Unit tests for ineligible-itc-vendor-extraction.js, using only
 * fictional synthetic GSTINs/names first, then a check against the real
 * local demo file's already-known extraction result (see the earlier
 * live investigation in this session, which found "Mazda Motors & Sons",
 * "Tata AIG General Insurance Co Ltd", etc. — this test re-verifies that
 * exact result programmatically instead of relying on memory of it).
 *
 * Run with: node src/taskpane/ineligible-itc-vendor-extraction-test.js
 */

const { extractVendorNamesFromSheet, extractDistinctVendors } = require("./ineligible-itc-vendor-extraction.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

const GSTIN_A = "27ZZAPL5432Q1Z9";
const GSTIN_B = "07ZZAPL7788Q1ZO";
const GSTIN_C = "29ZZAPL1122Q1ZI"; // description-field.js needs 3+ multi-row GSTINs before it will draw a conclusion either way

function runSyntheticTests() {
  console.log("-- 2A-style sheet: Trade/Legal Name column present --\n");

  const gstr2aValues = [
    ["GSTIN", "Trade/Legal Name", "Taxable Value"],
    [GSTIN_A, "Speedway Motors Pvt Ltd", 10000],
    [GSTIN_A, "Speedway Motors Pvt Ltd", 5000],
    [GSTIN_B, "Bharat General Insurance Co", 8000],
  ];
  const gstr2aColumns = { gstin: 0, tradeLegalName: 1, particulars: null, taxableValue: 2 };
  const fromGstr2a = extractVendorNamesFromSheet(gstr2aValues, 0, gstr2aColumns);
  assert(fromGstr2a.length === 2, "one entry per distinct GSTIN, not per row");
  assert(fromGstr2a.find((v) => v.gstin === GSTIN_A).vendorName === "Speedway Motors Pvt Ltd", "correct name pulled from Trade/Legal Name");

  console.log("\n-- Books-style sheet: no Trade/Legal Name, Particulars behaving as vendor name (constant per GSTIN) --\n");

  const booksValues = [
    ["GSTIN", "Particulars", "TV"],
    [GSTIN_A, "Speedway Motors Pvt Ltd", 10000],
    [GSTIN_A, "Speedway Motors Pvt Ltd", 5000],
    [GSTIN_A, "Speedway Motors Pvt Ltd", 3000],
    [GSTIN_B, "Bharat General Insurance Co", 8000],
    [GSTIN_B, "Bharat General Insurance Co", 4000],
    [GSTIN_B, "Bharat General Insurance Co", 2000],
    [GSTIN_C, "Sunrise Retail Traders", 1500],
    [GSTIN_C, "Sunrise Retail Traders", 2500],
  ];
  const booksColumns = { gstin: 0, tradeLegalName: null, particulars: 1, taxableValue: 2 };
  const fromBooks = extractVendorNamesFromSheet(booksValues, 0, booksColumns);
  assert(fromBooks.length === 3, "Particulars correctly used as the vendor-name source when it's structurally constant per GSTIN");
  assert(fromBooks.find((v) => v.gstin === GSTIN_B).vendorName === "Bharat General Insurance Co", "correct name pulled from Particulars acting as vendor name");

  console.log("\n-- Books-style sheet: Particulars is a GENUINE description (varies per transaction) — must NOT be used as vendor name --\n");

  const genuineDescValues = [
    ["GSTIN", "Particulars", "TV"],
    [GSTIN_A, "Office Rent - Monthly", 10000],
    [GSTIN_A, "Office Rent - Late Fee", 500],
    [GSTIN_B, "Legal Fees - Contract Review", 8000],
    [GSTIN_B, "Legal Fees - Litigation", 4000],
    [GSTIN_C, "Security Services - Guard Deployment", 3000],
    [GSTIN_C, "Security Services - Overtime Charges", 1200],
  ];
  const genuineDescResult = extractVendorNamesFromSheet(genuineDescValues, 0, booksColumns);
  assert(genuineDescResult.length === 0, "when Particulars is a genuine per-transaction description, it must NOT be harvested as a vendor name (would misuse a real description as a fake name)");

  console.log("\n-- No usable name column at all (e.g. no GSTIN, or neither field present) --\n");

  const noNameColumns = { gstin: 0, tradeLegalName: null, particulars: null, taxableValue: 1 };
  assert(extractVendorNamesFromSheet([["GSTIN", "TV"], [GSTIN_A, 1000]], 0, noNameColumns).length === 0, "no Trade/Legal Name and no Particulars -> nothing extractable");

  console.log("\n-- extractDistinctVendors merges 2A + Books, dedupes by GSTIN, prefers 2A's name --\n");

  const gstr2aWithOverlap = [
    ["GSTIN", "Trade/Legal Name"],
    [GSTIN_A, "Speedway Motors Private Limited"], // slightly different spelling than Books
  ];
  const merged = extractDistinctVendors(
    { values: gstr2aWithOverlap, headerRowIndex: 0, columns: { gstin: 0, tradeLegalName: 1, particulars: null } },
    { values: booksValues, headerRowIndex: 0, columns: booksColumns }
  );
  assert(merged.length === 3, "GSTIN_A (present in both, counted once) plus GSTIN_B and GSTIN_C from Books only");
  assert(merged.find((v) => v.gstin === GSTIN_A).vendorName === "Speedway Motors Private Limited", "when the same GSTIN appears on both sides, 2A's (government-recognized) name wins over Books'");

  console.log("");
}

function runRealFileCheck() {
  console.log("-- Real local demo file: re-verifying the exact vendor names found during the earlier live investigation --\n");
  let fileAvailable = true;
  let readWorkbookSheets, extractSheetSignals, recognizeGstSheets, identifyGstWorkflow;
  try {
    ({ readWorkbookSheets } = require("../sheet-classifier/xlsx-test-helper.js"));
    ({ recognizeGstSheets } = require("../gst-reconciliation/gst-reconciliation.js"));
    ({ identifyGstWorkflow } = require("../gst-reconciliation/gst-workflows.js"));
    var sheets = readWorkbookSheets("./src/demo-data/GST_2A_Books_Demo.xlsx");
  } catch (error) {
    fileAvailable = false;
  }

  if (!fileAvailable) {
    console.log("(local demo file not available in this environment — skipping; synthetic checks above already cover the logic)");
    return;
  }

  const roleResults = recognizeGstSheets(sheets);
  const workflow = identifyGstWorkflow(roleResults);
  const gstr2aSheet = workflow.sheets.gstr2a;
  const booksSheet = workflow.sheets.purchaseRegister;

  const vendors = extractDistinctVendors(
    { values: gstr2aSheet.result.values, headerRowIndex: gstr2aSheet.result.headerRowIndex, columns: gstr2aSheet.result.columns },
    { values: booksSheet.result.values, headerRowIndex: booksSheet.result.headerRowIndex, columns: booksSheet.result.columns }
  );

  const names = vendors.map((v) => v.vendorName.toLowerCase());
  assert(vendors.length > 0, "distinct vendors were actually extracted from the real file");
  assert(names.some((n) => n.includes("mazda motors")), "'Mazda Motors & Sons' present, matching the earlier manual investigation");
  assert(names.some((n) => n.includes("tata aig")), "'Tata AIG General Insurance' present");
  assert(names.some((n) => n.includes("indian hotels")), "'The Indian Hotels Company Ltd' present");
  assert(names.some((n) => n.includes("golden jubilee")), "'Golden Jubilee Hotels' present");
  console.log("  (" + vendors.length + " distinct vendors extracted total)");
  console.log("");
}

function run() {
  runSyntheticTests();
  runRealFileCheck();
  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All Ineligible ITC vendor-extraction checks passed.");
  }
}

run();
