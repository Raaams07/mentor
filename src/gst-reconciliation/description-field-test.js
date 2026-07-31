/*
 * description-field-test.js
 * -----------------------------
 * Unit tests for description-field.js, using only fictional synthetic
 * GSTINs and made-up data. No real vendor or invoice data.
 *
 * Run with: node src/gst-reconciliation/description-field-test.js
 */

const { findDescriptionColumn } = require("./description-field.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

const GSTIN_A = "27ZZAPL5432Q1Z9";
const GSTIN_B = "07ZZAPL7788Q1ZO";
const GSTIN_C = "29ZZAPL1122Q1ZI";
const GSTIN_D = "33ZZAPL9988Q1Z9";
const GSTIN_E = "36ZZAPL4455Q1ZR";

const COLUMNS_NO_PARTICULARS = { gstin: 0, particulars: null };

function runTests() {
  console.log("-- No Particulars column at all --\n");

  const noColumn = findDescriptionColumn([["GSTIN", "Vendor"], [GSTIN_A, "Vendor A"]], 0, COLUMNS_NO_PARTICULARS);
  assert(noColumn.available === false, "sheet with no Particulars/Narration column -> not available");
  assert(noColumn.reason.includes("no Particulars"), "...with a clear reason why");

  console.log("\n-- Genuine description field (varies across a vendor's own transactions) --\n");

  const columns = { gstin: 0, particulars: 1 };
  const genuineValues = [
    ["GSTIN", "Particulars"],
    [GSTIN_A, "Office Rent - Monthly"],
    [GSTIN_A, "Office Rent - Late Fee"],
    [GSTIN_B, "Legal Fees - Contract Review"],
    [GSTIN_B, "Legal Fees - Litigation"],
    [GSTIN_C, "Security Services - Guard Deployment"],
    [GSTIN_C, "Security Services - Overtime Charges"],
    [GSTIN_D, "Raw Material Purchase"],
    [GSTIN_D, "Freight Charges"],
  ];
  const genuine = findDescriptionColumn(genuineValues, 0, columns);
  assert(genuine.available === true, "Particulars content that genuinely varies per vendor's own transactions -> trusted as a real description field");

  console.log("\n-- Vendor-name-disguised-as-Particulars (constant per GSTIN) --\n");

  const vendorNameValues = [
    ["GSTIN", "Particulars"],
    [GSTIN_A, "Steel Traders Ltd"],
    [GSTIN_A, "Steel Traders Ltd"],
    [GSTIN_A, "Steel Traders Ltd"],
    [GSTIN_B, "Ispat Udyog Pvt Ltd"],
    [GSTIN_B, "Ispat Udyog Pvt Ltd"],
    [GSTIN_C, "Metal Corporation"],
    [GSTIN_C, "Metal Corporation"],
    [GSTIN_D, "Rukmini Steels"],
    [GSTIN_D, "Rukmini Steels"],
  ];
  const vendorNameLike = findDescriptionColumn(vendorNameValues, 0, columns);
  assert(vendorNameLike.available === false, "Particulars content that's constant per GSTIN across repeat transactions -> correctly identified as a vendor name, not a description field");
  assert(vendorNameLike.reason.includes("vendor name"), "...with a reason that explains why");

  console.log("\n-- Insufficient repeat-transaction data to judge --\n");

  const sparseValues = [
    ["GSTIN", "Particulars"],
    [GSTIN_A, "Whatever Text"],
    [GSTIN_B, "Something Else"],
  ];
  const sparse = findDescriptionColumn(sparseValues, 0, columns);
  assert(sparse.available === true, "too few multi-row vendors to statistically judge -> defaults to trusting the header identification rather than discarding it without evidence");

  console.log("\n-- No GSTIN column to cross-check against --\n");

  const noGstinValues = [
    ["Particulars"],
    ["Office Rent"],
  ];
  const noGstin = findDescriptionColumn(noGstinValues, 0, { gstin: null, particulars: 0 });
  assert(noGstin.available === true, "no GSTIN column available to group by -> can't verify further, trusts the header identification");

  console.log("");
}

function run() {
  runTests();
  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All description-field checks passed.");
  }
}

run();
