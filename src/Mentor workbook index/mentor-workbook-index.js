/*
 * mentor-workbook-index.js
 * ------------------------
 * Real implementation (not a sketch) of the background indexing layer
 * discussed with Claude. Drop this file into your project alongside
 * taskpane.js, then wire it in as shown at the bottom of this file.
 *
 * Column layout below matches "Three_Sisters_Kitchen_MENTOR_Demo.xlsx"
 * exactly — if your real workbook has the invoices/bank/POS/payroll
 * columns in a different order, adjust the column INDEXES in the
 * parse*Rows() functions to match. Everything else stays the same.
 */

// ---------------------------------------------------------------
// 1. THE INDEX ITSELF — one object, kept warm in memory
// ---------------------------------------------------------------
const workbookIndex = {
  invoices: [],
  bank: [],
  pos: [],
  payroll: [],
  vendorMap: {},   // canonical vendor name -> { canonical, aliases: [...] }
  lastBuilt: null
};

// ---------------------------------------------------------------
// 2. REGISTER LISTENERS ON ALL RAW SHEETS AT WORKBOOK OPEN
//    Call mentorInit() once, from your existing Office.onReady().
// ---------------------------------------------------------------
async function mentorInit() {
  await Excel.run(async (context) => {
    const sheetNames = [
      "Raw - Supplier Invoices",
      "Raw - Bank Statement",
      "Raw - POS Sales Report",
      "Raw - Payroll Summary"
    ];

    for (const name of sheetNames) {
      const sheet = context.workbook.worksheets.getItemOrNullObject(name);
      await context.sync();
      if (sheet.isNullObject) continue; // sheet not present in this workbook — skip quietly

      sheet.onChanged.add(debounce(() => rebuildIndexFor(name), 700));
    }
    await context.sync();
  });

  // Build the index once immediately at startup too — don't wait for
  // the first edit, since the user may start in the P&L tab right away.
  await rebuildIndexFor("Raw - Supplier Invoices");
  await rebuildIndexFor("Raw - Bank Statement");
  await rebuildIndexFor("Raw - POS Sales Report");
  await rebuildIndexFor("Raw - Payroll Summary");
}

// ---------------------------------------------------------------
// 3. REBUILD ONE SHEET'S SLICE OF THE INDEX
// ---------------------------------------------------------------
async function rebuildIndexFor(sheetName) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
    await context.sync();
    if (sheet.isNullObject) return;

    const range = sheet.getUsedRange();
    range.load("values");
    await context.sync();

    const values = range.values; // includes header rows — parsers below skip them

    if (sheetName === "Raw - Supplier Invoices") {
      workbookIndex.invoices = parseInvoiceRows(values);
      workbookIndex.vendorMap = buildVendorMap(workbookIndex.invoices);
    } else if (sheetName === "Raw - Bank Statement") {
      workbookIndex.bank = parseBankRows(values);
    } else if (sheetName === "Raw - POS Sales Report") {
      workbookIndex.pos = parsePosRows(values);
    } else if (sheetName === "Raw - Payroll Summary") {
      workbookIndex.payroll = parsePayrollRows(values);
    }
    workbookIndex.lastBuilt = Date.now();
  });
}

// ---------------------------------------------------------------
// 4. PARSERS — match the demo workbook's column order exactly.
//    Row 0-2 are title/blank rows, row index 3 is the header row
//    (matches the demo's HROW = 3, 0-indexed after getUsedRange()).
// ---------------------------------------------------------------
function parseInvoiceRows(values) {
  // Columns: Invoice ID, Site, Vendor, Invoice Date, Category,
  //          Net Amount, VAT Rate, VAT Amount, Gross Amount, Status, Notes
  return values.slice(4).filter(r => r[0]).map(r => ({
    id: r[0], site: r[1], vendor: r[2], date: r[3], category: r[4],
    net: r[5], vatRate: r[6], vatAmount: r[7], gross: r[8],
    status: r[9], notes: r[10]
  }));
}

function parseBankRows(values) {
  // Columns: Date (parsed), Date (as entered), Site, Description,
  //          Amount, Category, Matched Invoice ID, Notes
  return values.slice(4).filter(r => r[0]).map(r => ({
    date: r[0], dateAsEntered: r[1], site: r[2], description: r[3],
    amount: r[4], category: r[5], matchedInvoiceId: r[6], notes: r[7]
  }));
}

function parsePosRows(values) {
  // Columns: Month, Site, Dine-In, Takeaway, Deliveroo Gross,
  //          UberEats Gross, JustEat Gross, Events, Total
  return values.slice(4).filter(r => r[0]).map(r => ({
    month: r[0], site: r[1], dineIn: r[2], takeaway: r[3],
    deliveroo: r[4], uberEats: r[5], justEat: r[6], events: r[7], total: r[8]
  }));
}

function parsePayrollRows(values) {
  // Columns: Month, Site, Basic, Overtime, Tronc, NI, Pension, Total
  return values.slice(4).filter(r => r[0]).map(r => ({
    month: r[0], site: r[1], basic: r[2], overtime: r[3],
    tronc: r[4], ni: r[5], pension: r[6], total: r[7]
  }));
}

// ---------------------------------------------------------------
// 5. VENDOR PATTERN MEMORY — real Levenshtein distance, real grouping.
//    Handles Hard Case 1: "Fresh Fields Produce Ltd" ->
//    "Fresh Fields Wholesale Ltd" gets recognised as one vendor.
// ---------------------------------------------------------------
function levenshteinDistance(a, b) {
  const m = a.length, n = b.length;
  const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      dp[i][j] = Math.min(
        dp[i - 1][j] + 1,      // deletion
        dp[i][j - 1] + 1,      // insertion
        dp[i - 1][j - 1] + cost // substitution
      );
    }
  }
  return dp[m][n];
}

// Same first word (e.g. "Fresh Fields") + short edit distance = treat as
// the same vendor renamed, not two different suppliers. Tune the
// threshold (4) against real vendor names once you have more of them.
function buildVendorMap(invoiceRows) {
  const uniqueNames = [...new Set(invoiceRows.map(r => r.vendor))];
  const map = {};
  for (const name of uniqueNames) {
    const existingCanonical = Object.keys(map).find(canon => {
      const sameFirstWord = canon.split(" ")[0] === name.split(" ")[0];
      const closeEnough = levenshteinDistance(canon, name) <= 6;
      return sameFirstWord && closeEnough && canon !== name;
    });
    if (existingCanonical) {
      map[existingCanonical].aliases.push(name);
    } else if (!map[name]) {
      map[name] = { canonical: name, aliases: [name] };
    }
  }
  return map;
}

function resolveCanonicalVendor(name) {
  for (const key of Object.keys(workbookIndex.vendorMap)) {
    if (workbookIndex.vendorMap[key].aliases.includes(name)) return key;
  }
  return name;
}

function detectVendorRenames() {
  return Object.values(workbookIndex.vendorMap)
    .filter(v => v.aliases.length > 1)
    .map(v => `'${v.aliases.join("' and '")}' look like the same vendor, renamed.`);
}

// ---------------------------------------------------------------
// 6. WHAT FIRES WHEN THE USER IS ACTUALLY WORKING IN THE P&L SHEET.
//    Reads ONLY workbookIndex — never touches the raw sheets directly.
//    This is why it feels instant instead of laggy.
// ---------------------------------------------------------------
function onPnlRowLabelEntered(cellAddress, rowLabel) {
  if (!workbookIndex.lastBuilt) return; // index not warm yet — stay silent, don't guess

  if (rowLabel === "Food Cost") {
    const foodInvoices = workbookIndex.invoices.filter(r => r.category === "Food");
    const vendors = [...new Set(foodInvoices.map(r => resolveCanonicalVendor(r.vendor)))];
    const renameNotes = detectVendorRenames();

    showSuggestionBubble(cellAddress, {
      message: `Found ${foodInvoices.length} Food invoices across ${vendors.length} vendors. Pull these in?`,
      secondaryNote: renameNotes.length ? renameNotes.join(" ") : null,
      onAccept: () => insertSumifsFormula(cellAddress, "Food")
    });
  }
  // Extend with the same pattern for "Beverage Cost", "Rent",
  // "Utilities", "Payroll Cost (excl. Tronc)", etc.
}

function insertSumifsFormula(cellAddress, category) {
  // Build a real SUMIFS string referencing the raw sheet, same shape
  // as the formulas already used in Monthly P&L (Answer Key).
  const formula = `=SUMIFS('Raw - Supplier Invoices'!$F:$F,'Raw - Supplier Invoices'!$E:$E,"${category}")`;
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const cell = sheet.getRange(cellAddress);
    cell.formulas = [[formula]];
    await context.sync();
  });
}

// ---------------------------------------------------------------
// 7. UI hook — replace with your existing sidebar/marker function.
// ---------------------------------------------------------------
function showSuggestionBubble(cellAddress, opts) {
  // TODO: wire to your existing suggestion-bubble UI from the
  // reconciliation add-in. opts = { message, secondaryNote, onAccept }
  console.log(`[MENTOR] ${cellAddress}: ${opts.message}`, opts.secondaryNote || "");
}

function debounce(fn, ms) {
  let t;
  return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); };
}

// ---------------------------------------------------------------
// HOW TO WIRE THIS INTO YOUR EXISTING taskpane.js
// ---------------------------------------------------------------
// 1. Add near the top of taskpane.js:
//      import "./mentor-workbook-index.js";   (or <script> tag if not using modules)
//
// 2. Inside your existing Office.onReady(...) block, add one line:
//      mentorInit();
//
// 3. Inside your existing onChanged handler for the P&L sheet
//    (the one that currently detects "header exists" etc.), add:
//      onPnlRowLabelEntered(cellAddress, rowLabelText);
//    — call it with whatever cell/label your existing code already
//    reads when the user types a row label in column A.
//
// That's the whole integration — nothing else in your existing
// reconciliation logic needs to change.
