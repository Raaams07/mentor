/*
 * regression-test.js
 * -------------------
 * MENTOR Phase 2, Step 3 — regression check against the Three Sisters
 * Kitchen demo dataset.
 *
 * Before this generalized signal-extraction + rule-based classifier
 * (Step 1 + Step 2) is trusted on an unfamiliar client workbook, it has
 * to at least still recognize the four sheets it was designed against.
 * This is that check: read the real demo .xlsx, classify every sheet,
 * and assert the four known raw-data sheets land on the right type.
 *
 * Run with: node src/sheet-classifier/regression-test.js
 * Exits 1 on any failed assertion (wired to `npm run test:classifier`).
 *
 * Uses the `xlsx` (SheetJS) package — a devDependency added solely for
 * this test harness, not part of the shipped add-in bundle.
 */

const path = require("path");
const { readWorkbookSheets } = require("./xlsx-test-helper.js");
const { extractWorkbookSignals } = require("./signal-extractor.js");
const { classifyWorkbook } = require("./classifier.js");

const DEMO_FILE = path.join(__dirname, "..", "demo-data", "Three_Sisters_Kitchen_MENTOR_Demo.xlsx");

// The four raw-data sheets Step 2's categories are meant to recognize.
const EXPECTED_KNOWN = {
  "Raw - Supplier Invoices": "invoice_list",
  "Raw - Bank Statement": "bank_statement",
  "Raw - POS Sales Report": "sales_pos",
  "Raw - Payroll Summary": "payroll",
};

// Report/summary sheets that aren't any of the four raw-data shapes — these
// should come back "unknown" ("Not recognized yet" to the user) rather than
// being force-fit into whichever category happens to score highest.
const EXPECTED_UNKNOWN = ["README - Ground Truth", "Monthly P&L (Demo - Start Here)", "Monthly P&L (Answer Key)", "Variance Analysis"];

function run() {
  const sheets = readWorkbookSheets(DEMO_FILE);
  const signals = extractWorkbookSignals(sheets);
  const classifications = classifyWorkbook(signals);
  const bySheetName = new Map(classifications.map((c) => [c.sheetName, c]));

  console.log(`Classified ${classifications.length} sheet(s) in ${path.basename(DEMO_FILE)}\n`);

  let failures = 0;
  const totalChecks = Object.keys(EXPECTED_KNOWN).length + EXPECTED_UNKNOWN.length;

  function reportCheck(sheetName, expectedType, result) {
    if (!result) {
      console.log(`FAIL  ${sheetName} — sheet not found in workbook`);
      failures++;
      return;
    }

    const pass = result.type === expectedType;
    if (!pass) failures++;

    console.log(`${pass ? "PASS" : "FAIL"}  ${sheetName}`);
    console.log(`      expected: ${expectedType}   actual: ${result.type} ("${result.displayLabel}")   confidence: ${result.confidence}`);
    console.log(`      scores: ${JSON.stringify(result.scores)}`);

    if (!pass) {
      console.log("      matched rules for actual type:");
      (result.matchedRules || []).forEach((r) => console.log(`        +${r.weight} ${r.label}`));
      if (result.reason) console.log(`      reason: ${result.reason}`);
    }
    console.log("");
  }

  console.log("-- Four known raw-data sheets (should classify) --\n");
  for (const [sheetName, expectedType] of Object.entries(EXPECTED_KNOWN)) {
    reportCheck(sheetName, expectedType, bySheetName.get(sheetName));
  }

  console.log("-- Report/summary sheets (should NOT be force-fit — expect \"unknown\") --\n");
  for (const sheetName of EXPECTED_UNKNOWN) {
    reportCheck(sheetName, "unknown", bySheetName.get(sheetName));
  }

  if (failures > 0) {
    console.log(`${failures} of ${totalChecks} regression check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log(`All ${totalChecks} regression checks passed.`);
  }
}

run();
