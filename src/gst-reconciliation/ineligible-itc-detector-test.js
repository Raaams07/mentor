/*
 * ineligible-itc-detector-test.js
 * -----------------------------------
 * Unit tests for ineligible-itc-detector.js, using only fictional
 * synthetic GSTINs and made-up expense descriptions. No real vendor or
 * invoice data.
 *
 * Run with: node src/gst-reconciliation/ineligible-itc-detector-test.js
 */

const { detectIneligibleItcForRow, detectIneligibleItcForSheet } = require("./ineligible-itc-detector.js");
const { ITC_INELIGIBLE_CATEGORIES } = require("./ineligible-itc-config.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

const GSTIN_A = "27ZZAPL5432Q1Z9";
const GSTIN_B = "07ZZAPL7788Q1ZO";
const GSTIN_C = "29ZZAPL1122Q1ZI";
const DESCRIPTION_COL = 1; // [GSTIN, Particulars]

function runTests() {
  console.log("-- Keyword matches per blocked-credit category (given a resolved description column) --\n");

  const motorVehicle = detectIneligibleItcForRow([GSTIN_A, "Car Insurance Premium - Company Vehicle"], DESCRIPTION_COL);
  assert(motorVehicle.isIneligible && motorVehicle.matchedCategory.id === "motor_vehicles", "motor vehicle insurance correctly flagged");
  assert(motorVehicle.matchedCategory.section.includes("17(5)"), "flagged item carries its CGST Act section reference");

  const foodBeverage = detectIneligibleItcForRow([GSTIN_A, "Staff Canteen Catering Expenses"], DESCRIPTION_COL);
  assert(foodBeverage.isIneligible && foodBeverage.matchedCategory.id === "food_beverage", "canteen/catering expense correctly flagged");

  const csr = detectIneligibleItcForRow([GSTIN_A, "CSR Donation - Education Trust"], DESCRIPTION_COL);
  assert(csr.isIneligible && csr.matchedCategory.id === "csr_expenditure", "CSR expenditure correctly flagged");
  assert(csr.matchedCategory.section.includes("2023"), "CSR category cites the 2023 amendment that introduced it, not an older section alone");

  const stationery = detectIneligibleItcForRow([GSTIN_A, "Office Stationery Purchase"], DESCRIPTION_COL);
  assert(stationery.isIneligible === false, "an ordinary, unrelated expense is not flagged");

  console.log("\n-- No description column resolved (never falls back to a vendor name) --\n");

  const noColumn = detectIneligibleItcForRow(["27AAAPL5432Q1Z9 Foods Pvt Ltd — this looks like a vendor name containing 'food'"], null);
  assert(noColumn.isIneligible === false, "with descriptionColumnIndex=null, nothing is ever matched — there is no fallback to a vendor-name-shaped string");

  console.log("\n-- Known limitation: composition-scheme purchases aren't keyword-detectable --\n");

  const compositionCategory = ITC_INELIGIBLE_CATEGORIES.find((c) => c.id === "composition_scheme_purchase");
  assert(compositionCategory.keywords.length === 0, "composition-scheme category deliberately has no keywords — it depends on supplier registration type, not item description, and is documented as such rather than silently guessed at");

  console.log("\n-- Sheet-level scan: genuine description field present --\n");

  const genuineColumns = { gstin: 0, particulars: 1, invoiceNumber: null, voucherNumber: null };
  const genuineSheetValues = [
    ["GSTIN", "Particulars"],
    [GSTIN_A, "Office Stationery Purchase"],
    [GSTIN_A, "Car Insurance Premium"],
    [GSTIN_B, "Staff Canteen Expenses"],
    [GSTIN_B, "Raw Material Purchase"],
  ];
  const genuineResult = detectIneligibleItcForSheet(genuineSheetValues, 0, genuineColumns);
  assert(genuineResult.applicable === true, "sheet with a genuinely-varying Particulars column is applicable");
  assert(genuineResult.flagged.length === 2, "...and correctly flags the 2 matching rows");

  console.log("\n-- Sheet-level scan: Particulars is actually vendor names (not applicable) --\n");

  const vendorNameColumns = { gstin: 0, particulars: 1, invoiceNumber: null, voucherNumber: null };
  const vendorNameSheetValues = [
    ["GSTIN", "Particulars"],
    [GSTIN_A, "Steel Traders Ltd"],
    [GSTIN_A, "Steel Traders Ltd"],
    [GSTIN_A, "Steel Traders Ltd"],
    [GSTIN_B, "Ispat Foods Udyog Pvt Ltd"], // coincidentally contains "food" — must NOT be matched
    [GSTIN_B, "Ispat Foods Udyog Pvt Ltd"],
    [GSTIN_B, "Ispat Foods Udyog Pvt Ltd"],
    [GSTIN_C, "Metal Corporation"],
    [GSTIN_C, "Metal Corporation"],
  ];
  const vendorNameResult = detectIneligibleItcForSheet(vendorNameSheetValues, 0, vendorNameColumns);
  assert(vendorNameResult.applicable === false, "sheet where Particulars is structurally a vendor name (constant per GSTIN) is reported as NOT applicable");
  assert(vendorNameResult.flagged.length === 0, "...and produces zero flagged rows rather than a false 'food' match against the vendor's own name");
  assert(typeof vendorNameResult.reason === "string" && vendorNameResult.reason.length > 0, "...with an explicit reason, not a silent empty result");

  console.log("\n-- Sheet-level scan: no Particulars column at all --\n");

  const noParticularsColumns = { gstin: 0, particulars: null, invoiceNumber: null, voucherNumber: null };
  const noParticularsResult = detectIneligibleItcForSheet([["GSTIN"], [GSTIN_A]], 0, noParticularsColumns);
  assert(noParticularsResult.applicable === false, "sheet with no Particulars/Narration column at all (e.g. a typical GSTR-2A sheet) is reported as not applicable, not silently checked against Trade/Legal Name");

  console.log("");
}

function run() {
  runTests();
  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All Ineligible ITC detector checks passed.");
  }
}

run();
