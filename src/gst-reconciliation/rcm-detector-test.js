/*
 * rcm-detector-test.js
 * -----------------------
 * Unit tests for rcm-detector.js, using only fictional synthetic GSTINs
 * and made-up expense descriptions. No real vendor or invoice data.
 *
 * Run with: node src/gst-reconciliation/rcm-detector-test.js
 */

const { detectRcmForRow, detectRcmForSheet } = require("./rcm-detector.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

const GSTIN_FICTIONAL = "27ZZAPL5432Q1Z9";

function runTests() {
  console.log("-- Primary signal: the 2A 'Supply Attract Reverse Charge' flag --\n");

  const withYesFlag = { gstin: 0, particulars: 1, reverseCharge: 2, invoiceNumber: null, voucherNumber: null };
  const yesResult = detectRcmForRow([GSTIN_FICTIONAL, "Office stationery purchase", "Y"], withYesFlag, 1);
  assert(yesResult.isRcm === true && yesResult.source === "reverse_charge_flag", "flag='Y' -> RCM true, sourced from the flag (even though the description has no RCM keywords)");

  const noResult = detectRcmForRow([GSTIN_FICTIONAL, "GTA freight charges for inbound material", "N"], withYesFlag, 1);
  assert(noResult.isRcm === false && noResult.source === "reverse_charge_flag", "flag='N' -> RCM false, TRUSTING the flag even though the description looks like a GTA keyword match — flag is primary, not just a tiebreaker");

  console.log("\n-- Fallback: keyword match against a RESOLVED description column --\n");

  const noFlagColumn = { gstin: 0, particulars: 1, reverseCharge: null, invoiceNumber: null, voucherNumber: null };
  const gtaMatch = detectRcmForRow([GSTIN_FICTIONAL, "GTA Freight Charges - Inbound Material"], noFlagColumn, 1);
  assert(gtaMatch.isRcm === true && gtaMatch.source === "keyword_match" && gtaMatch.matchedCategory.id === "gta_freight", "no flag column -> falls back to keyword match against the resolved description column, correctly identifies GTA freight");

  const legalMatch = detectRcmForRow([GSTIN_FICTIONAL, "Legal Fees - Advocate Retainer"], noFlagColumn, 1);
  assert(legalMatch.isRcm === true && legalMatch.matchedCategory.id === "legal_services", "keyword match correctly identifies legal services from an advocate");

  const noMatch = detectRcmForRow([GSTIN_FICTIONAL, "Office Stationery Purchase"], noFlagColumn, 1);
  assert(noMatch.isRcm === false && noMatch.source === "no_signal", "no flag, no keyword match -> not RCM, source is 'no_signal'");

  const blankFlagFallsThrough = detectRcmForRow([GSTIN_FICTIONAL, "Security Services - Guard Deployment", ""], withYesFlag, 1);
  assert(blankFlagFallsThrough.isRcm === true && blankFlagFallsThrough.source === "keyword_match" && blankFlagFallsThrough.matchedCategory.id === "security_services", "flag column present but blank -> falls through to keyword matching instead of being treated as 'No'");

  console.log("\n-- No description column resolved (never falls back to a vendor name) --\n");

  const noDescriptionColumn = detectRcmForRow([GSTIN_FICTIONAL, "27AAAPL5432Q1Z9 GTA Freight Corp — a vendor name that happens to contain a keyword"], noFlagColumn, null);
  assert(noDescriptionColumn.isRcm === false && noDescriptionColumn.source === "no_signal", "with descriptionColumnIndex=null, nothing is matched — there is no fallback to a vendor-name-shaped string, even one containing an RCM keyword");

  console.log("\n-- Weak-signal category (requires manual confirmation) --\n");

  const rentMatch = detectRcmForRow([GSTIN_FICTIONAL, "Office Rent - Monthly"], noFlagColumn, 1);
  assert(rentMatch.isRcm === true && rentMatch.matchedCategory.id === "rent_unregistered_landlord", "rent keyword matches the rent category");
  assert(rentMatch.matchedCategory.requiresManualConfirmation === true, "...and is explicitly flagged as needing manual confirmation, since 'rent' alone can't confirm the landlord is unregistered");

  console.log("\n-- Notification reference is present for traceability --\n");
  assert(typeof gtaMatch.matchedCategory.notification === "string" && gtaMatch.matchedCategory.notification.length > 0, "matched category carries its notification reference for the user to verify");

  console.log("\n-- Conditional category: metal scrap RCM only applies to an UNREGISTERED supplier --\n");

  const columnsWithGstin = { gstin: 0, particulars: 1, reverseCharge: null, invoiceNumber: null, voucherNumber: null };

  const metalScrapUnregistered = detectRcmForRow(["", "Purchase of Metal Scrap - Mixed Ferrous"], columnsWithGstin, 1);
  assert(metalScrapUnregistered.isRcm === true && metalScrapUnregistered.matchedCategory.id === "metal_scrap", "metal scrap keyword + BLANK GSTIN (unregistered supplier) -> correctly flagged as RCM");

  const metalScrapRegistered = detectRcmForRow([GSTIN_FICTIONAL, "Purchase of Metal Scrap - Mixed Ferrous"], columnsWithGstin, 1);
  assert(metalScrapRegistered.isRcm === false && metalScrapRegistered.source === "no_signal", "metal scrap keyword + POPULATED GSTIN (registered supplier) -> NOT flagged as RCM (forward-charge + Section 51 TDS applies instead, not this rule's concern)");

  console.log("\n-- Sheet-level scan: flag column present, description also genuine --\n");

  const headerRow = ["GSTIN", "Particulars", "Reverse Charge"];
  const sheetValues = [
    headerRow,
    [GSTIN_FICTIONAL, "Office Stationery Purchase", "N"],
    [GSTIN_FICTIONAL, "GTA Freight Charges", ""],
    [GSTIN_FICTIONAL, "Director Sitting Fees", ""],
  ];
  const sheetColumns = { gstin: 0, particulars: 1, reverseCharge: 2, invoiceNumber: null, voucherNumber: null };
  const sheetResult = detectRcmForSheet(sheetValues, 0, sheetColumns);
  assert(sheetResult.flagged.length === 2, "sheet-level scan returns only the RCM rows (2 of 3), not the non-RCM one");
  assert(sheetResult.flagSignalAvailable === true, "flag signal correctly reported as available");

  console.log("\n-- Sheet-level scan: unregistered-supplier row (blank GSTIN) is still reachable --\n");

  const sheetWithUnregisteredSupplier = [
    headerRow,
    ["", "Purchase of Metal Scrap - Mixed Ferrous", ""], // blank GSTIN — unregistered supplier
    [GSTIN_FICTIONAL, "Purchase of Metal Scrap - Mixed Ferrous", ""], // registered supplier
  ];
  const scrapScanResult = detectRcmForSheet(sheetWithUnregisteredSupplier, 0, sheetColumns);
  assert(scrapScanResult.flagged.length === 1, "sheet-level scan correctly reaches the blank-GSTIN row (previously would have been skipped entirely) and correctly excludes the registered-supplier row");
  assert(scrapScanResult.flagged[0].gstin === null, "the flagged row's GSTIN is reported as null (unregistered), not silently dropped or mistaken for missing data");

  console.log("\n-- Sheet-level scan: no flag column AND Particulars is structurally a vendor name --\n");

  const gstinX = "07ZZAPL7788Q1ZO";
  const gstinY = "29ZZAPL1122Q1ZI";
  const gstinZ = "33ZZAPL9988Q1Z9";
  const vendorNameSheet = [
    ["GSTIN", "Particulars"],
    [GSTIN_FICTIONAL, "GTA Freight Transport Co"], // vendor name that happens to contain "freight" — must NOT be matched
    [GSTIN_FICTIONAL, "GTA Freight Transport Co"],
    [GSTIN_FICTIONAL, "GTA Freight Transport Co"],
    [gstinX, "Security Solutions Pvt Ltd"], // contains "security" — must NOT be matched
    [gstinX, "Security Solutions Pvt Ltd"],
    [gstinX, "Security Solutions Pvt Ltd"],
    [gstinY, "Metal Corporation"],
    [gstinY, "Metal Corporation"],
    [gstinZ, "Steel Traders Ltd"],
    [gstinZ, "Steel Traders Ltd"],
  ];
  const vendorNameSheetColumns = { gstin: 0, particulars: 1, reverseCharge: null, invoiceNumber: null, voucherNumber: null };
  const vendorNameSheetResult = detectRcmForSheet(vendorNameSheet, 0, vendorNameSheetColumns);
  assert(vendorNameSheetResult.keywordFallbackAvailable === false, "no flag column AND Particulars is constant-per-GSTIN (a vendor name) -> keyword fallback correctly reported as unavailable");
  assert(vendorNameSheetResult.flagged.length === 0, "...so zero rows are flagged, despite 'GTA Freight' and 'Security' both appearing in vendor names on this sheet");

  console.log("");
}

function run() {
  runTests();
  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All RCM detector checks passed.");
  }
}

run();
