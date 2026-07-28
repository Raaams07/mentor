/*
 * MENTOR — Background WorkbookIndex sketch
 * -----------------------------------------
 * Goal: stop scanning sheets reactively when the user types.
 * Instead, keep a standing "index" of every raw sheet, rebuilt
 * whenever THAT sheet changes — regardless of which sheet the
 * user currently has open. When the user is later typing into
 * the P&L sheet, MENTOR consults the index instantly instead of
 * reading the workbook live.
 *
 * This is illustrative structure, not drop-in production code —
 * adapt names/ranges to your actual add-in.
 */

// One in-memory index per raw sheet, kept fresh in the background.
let workbookIndex = {
  invoices: null,   // parsed rows from "Raw - Supplier Invoices"
  bank: null,       // parsed rows from "Raw - Bank Statement"
  pos: null,        // parsed rows from "Raw - POS Sales Report"
  payroll: null,    // parsed rows from "Raw - Payroll Summary"
  vendorMap: null,  // built FROM invoices: canonical vendor -> known aliases
  lastBuilt: null
};

// 1. Register a change listener on EACH raw sheet, independent of
//    whatever sheet is currently active. This is the "read overnight"
//    behaviour — it runs whether or not the user is looking at it.
async function registerBackgroundListeners() {
  await Excel.run(async (context) => {
    const sheets = ["Raw - Supplier Invoices", "Raw - Bank Statement",
                     "Raw - POS Sales Report", "Raw - Payroll Summary"];
    for (const name of sheets) {
      const sheet = context.workbook.worksheets.getItem(name);
      sheet.onChanged.add(debounce(() => rebuildIndexFor(name), 700));
    }
    await context.sync();
  });
}

// 2. Rebuild just the piece of the index that changed — cheap, incremental,
//    not a full workbook rescan every keystroke.
async function rebuildIndexFor(sheetName) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getUsedRange();
    range.load("values");
    await context.sync();

    const rows = parseSheetRows(sheetName, range.values); // your own parser

    if (sheetName === "Raw - Supplier Invoices") {
      workbookIndex.invoices = rows;
      workbookIndex.vendorMap = buildVendorMap(rows); // fuzzy-match pass, Levenshtein reuse
    } else if (sheetName === "Raw - Bank Statement") {
      workbookIndex.bank = rows;
    } else if (sheetName === "Raw - POS Sales Report") {
      workbookIndex.pos = rows;
    } else if (sheetName === "Raw - Payroll Summary") {
      workbookIndex.payroll = rows;
    }
    workbookIndex.lastBuilt = Date.now();
  });
}

// 3. Vendor pattern-memory (Hard Case 1 — Fresh Fields rename).
//    Groups near-duplicate vendor names using your existing Levenshtein
//    distance function from the reconciliation add-in.
function buildVendorMap(invoiceRows) {
  const vendorNames = [...new Set(invoiceRows.map(r => r.vendor))];
  const map = {};
  for (const name of vendorNames) {
    const existingCanonical = Object.keys(map).find(
      canon => levenshteinDistance(canon, name) <= 4 // tune threshold
    );
    if (existingCanonical) {
      map[existingCanonical].aliases.push(name);
    } else {
      map[name] = { canonical: name, aliases: [name] };
    }
  }
  return map;
}

// 4. THIS is what actually fires when the user is in the P&L sheet.
//    Notice: it never touches the raw sheets directly — it only reads
//    workbookIndex, which is already warm. This is why it feels instant
//    instead of laggy, and why it works even if the user never opens
//    the raw sheets themselves that session.
async function onPnlCellEdited(cellAddress, rowLabel) {
  if (!workbookIndex.invoices) return; // index not warm yet — stay silent

  if (rowLabel === "Food Cost") {
    const foodInvoices = workbookIndex.invoices.filter(r => r.category === "Food");
    const vendorCount = new Set(foodInvoices.map(r =>
      resolveCanonicalVendor(r.vendor, workbookIndex.vendorMap)
    )).size;

    showSuggestionBubble(cellAddress, {
      message: `Found ${foodInvoices.length} Food invoices across ${vendorCount} vendors. Pull these in?`,
      onAccept: () => insertFormula(cellAddress, buildSumifsFormula(foodInvoices)),
      // Secondary, lower-priority note — same bubble, expandable, not blocking:
      secondaryNote: detectVendorRenames(workbookIndex.vendorMap)
    });
  }
}

// Placeholder stubs for illustration — these already exist in some form
// in your reconciliation add-in and can be reused directly.
function levenshteinDistance(a, b) { /* ...existing implementation... */ }
function resolveCanonicalVendor(name, map) { /* ...lookup... */ }
function detectVendorRenames(map) { /* ...returns a short note if aliases found... */ }
function parseSheetRows(sheetName, values) { /* ...your column mapping per sheet... */ }
function buildSumifsFormula(rows) { /* ...construct the formula string... */ }
function insertFormula(cellAddress, formula) { /* ...write to the cell... */ }
function showSuggestionBubble(cellAddress, opts) { /* ...your existing sidebar/marker UI... */ }
function debounce(fn, ms) { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; }
