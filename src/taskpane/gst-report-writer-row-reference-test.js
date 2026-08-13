/*
 * gst-report-writer-row-reference-test.js
 * -------------------------------------------
 * Regression test for a real incident: every "Row"/"Row (in X)" column
 * printed on a generated GST report sheet showed a detector's internal,
 * RELATIVE row index (0-based, counting from the first data row after the
 * header) instead of the row's actual 1-based Excel row number. A person
 * manually navigating to "row N" on the source sheet landed on completely
 * unrelated data — confirmed on a real client file, where a genuine
 * duplicate-invoice pair was printed as "row 387/390" when the actual
 * matching rows were 395/398 (headerRowIndex was 6, so every printed "Row"
 * value was off by exactly headerRowIndex + 2).
 *
 * For every sheet type that has a "Row" column, this test builds the same
 * row array production code builds (using the REAL, exported row-builder
 * functions from gst-report-writer.js — not a re-derived copy), takes the
 * "Row" value it produced, dereferences it against the raw source `values`
 * array EXACTLY the way a human manually checking a flagged row in Excel
 * would (1-based Excel row -> 0-based array index), and confirms that's
 * the SAME row the flagged item's own identifying fields describe — not
 * some other row nearby.
 *
 * Source sheets deliberately use a header row NOT at index 0, and the two
 * sheets use DIFFERENT header positions from each other — this both
 * mirrors real files (the incident's real file had headerRowIndex=6) and
 * would catch a variant of the same bug class where one sheet's
 * headerRowIndex gets used for the other sheet's row reference.
 *
 * Run with: node src/taskpane/gst-report-writer-row-reference-test.js
 */

// mentor-gst-reconciliation-ui.js (transitively required by gst-report-
// writer.js's own requires) instantiates a BrowserSheetMemoryStore at
// module load time, whose constructor calls window.localStorage — throws
// in plain Node without this stub. Same pattern already used by
// scripts/validate-real-files.js.
global.window = { localStorage: { getItem: () => null, setItem: () => {} } };

const {
  toAbsoluteExcelRow,
  extraInvoiceRow,
  possibleMatchRow,
  wrongHeadRow,
  crossSheetWrongHeadRow,
  rcmRow,
  ineligibleItcRow,
  duplicateInvoiceRow,
  rateMismatchRow,
} = require("./gst-report-writer.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

// Dereferences a "Row" value the exact way a human manually checking a
// flagged row in Excel would: 1-based Excel row -> 0-based array index.
function rowAt(sourceValues, reportedExcelRow) {
  return sourceValues[reportedExcelRow - 1];
}

// Combines dereferencing + the identity check into one safe assertion — if
// a regression makes the reported Row wildly wrong (out of the source
// sheet's bounds entirely, e.g. the raw relative index on a sheet with few
// data rows), this reports a clean FAIL instead of crashing the whole run
// on an undefined access, so every other check still gets a chance to run
// and report its own result.
function assertRowMatches(sourceValues, reportedRow, checkFn, description) {
  const actual = rowAt(sourceValues, reportedRow);
  assert(actual !== undefined && checkFn(actual), description);
}

// --- Synthetic "2A" sheet: 2 rows above the header (headerRowIndex = 2) ---
const GSTR2A_HEADER_ROW_INDEX = 2;
const SHEET_2A = [
  ["ACME CORP — GSTR-2A Extract"],
  [],
  ["GSTIN of supplier", "Invoice details", "Place of Supply", "Taxable Value (₹)", "IGST", "CGST", "SGST"],
  ["11AAAAA0000A1Z1", "2A-INV-0", "Karnataka", 1000, 180, 0, 0],
  ["22BBBBB1111B1Z2", "2A-INV-1", "Maharashtra", 2000, 0, 180, 180],
  ["33CCCCC2222C1Z3", "2A-INV-2", "Telangana", 3000, 540, 0, 0],
  ["44DDDDD3333D1Z4", "2A-INV-3", "Karnataka", 4000, 720, 0, 0],
];

// --- Synthetic "Books" sheet: 4 rows above the header (headerRowIndex = 4)
// — deliberately DIFFERENT from 2A's, to catch a swapped-headerRowIndex bug ---
const BOOKS_HEADER_ROW_INDEX = 4;
const SHEET_BOOKS = [
  ["Company Books"],
  [],
  [],
  [],
  ["Voucher No.", "Supplier Invoice No.", "GSTIN/UIN", "Taxable Value", "IGST", "CGST", "SGST"],
  ["V-100", "BOOKS-INV-0", "11AAAAA0000A1Z1", 1000, 180, 0, 0],
  ["V-101", "BOOKS-INV-1", "55EEEEE4444E1Z5", 5000, 0, 450, 450],
  ["V-102", "BOOKS-INV-2", "66FFFFF5555F1Z6", 6000, 0, 540, 540],
];

function runToAbsoluteExcelRow() {
  console.log("-- toAbsoluteExcelRow: the core conversion every row-builder below routes through --\n");
  assert(toAbsoluteExcelRow(GSTR2A_HEADER_ROW_INDEX, 0) === 4, "headerRowIndex=2, relative 0 -> Excel row 4 (the first data row)");
  assert(toAbsoluteExcelRow(GSTR2A_HEADER_ROW_INDEX, 3) === 7, "headerRowIndex=2, relative 3 -> Excel row 7");
  assert(toAbsoluteExcelRow(BOOKS_HEADER_ROW_INDEX, 0) === 6, "headerRowIndex=4, relative 0 -> Excel row 6 (different header position, still correct)");
  assert(toAbsoluteExcelRow(-1, 0) === 1, "headerRowIndex=-1 (no header row found) -> relative row 0 is literally Excel row 1");
  // The exact class of bug this whole test file exists to catch: the raw
  // relative index is NEVER the correct Excel row on its own (previously,
  // this is exactly what got printed).
  assert(toAbsoluteExcelRow(GSTR2A_HEADER_ROW_INDEX, 3) !== 3, "the relative index itself (3) is NOT the correct Excel row (7) -- confirms the bug this test guards against would actually be caught");
  console.log("");
}

function runExtraInvoiceRow() {
  console.log("-- extraInvoiceRow (GST - Extra in Books / GST - Extra in 2A) --\n");
  const r = { rowIndex: 1, gstin: "55EEEEE4444E1Z5", identifier: "BOOKS-INV-1", taxableValue: 5000, igst: 0, cgst: 450, sgst: 450 };
  const row = extraInvoiceRow(r, BOOKS_HEADER_ROW_INDEX);
  const reportedRow = row[2];
  assert(reportedRow === 7, "reported Row is the correct absolute Excel row (7), not the relative index (1)");
  assertRowMatches(SHEET_BOOKS, reportedRow, (actual) => actual[1] === r.identifier && actual[2] === r.gstin, "dereferencing the reported Row on the REAL source sheet lands on the row this report row actually describes (Invoice \"" + r.identifier + "\", GSTIN \"" + r.gstin + "\")");
  console.log("");
}

function runPossibleMatchRow() {
  console.log("-- possibleMatchRow (GST - Possible Matches) --\n");
  const m = {
    gstin: "cross-vendor-check",
    gstr2aIdentifier: "2A-INV-2",
    gstr2aRowIndex: 2,
    booksIdentifier: "BOOKS-INV-2",
    booksRowIndex: 2,
    gstr2a: { taxableValue: 3000, igst: 540, cgst: 0, sgst: 0 },
    books: { taxableValue: 6000, igst: 0, cgst: 540, sgst: 540 },
    reason: "amount+GSTIN, same month",
  };
  const row = possibleMatchRow(m, GSTR2A_HEADER_ROW_INDEX, BOOKS_HEADER_ROW_INDEX);
  const reportedGstr2aRow = row[2];
  const reportedBooksRow = row[4];
  assert(reportedGstr2aRow === 6, "reported 2A Row is the correct absolute Excel row (6), not the relative index (2)");
  assert(reportedBooksRow === 8, "reported Books Row is the correct absolute Excel row (8), not the relative index (2)");
  assertRowMatches(SHEET_2A, reportedGstr2aRow, (actual) => actual[1] === m.gstr2aIdentifier, "dereferencing the reported 2A Row lands on Invoice \"" + m.gstr2aIdentifier + "\", not some other row");
  assertRowMatches(SHEET_BOOKS, reportedBooksRow, (actual) => actual[1] === m.booksIdentifier, "dereferencing the reported Books Row lands on Invoice \"" + m.booksIdentifier + "\", not some other row");
  console.log("");
}

function runWrongHeadRow() {
  console.log("-- wrongHeadRow (GST - Wrong Head, single-sheet section) --\n");
  const f = { rowIndex: 3, gstin: "44DDDDD3333D1Z4", invoiceNumber: "2A-INV-3", issue: "wrong_head", expected: "igst", actual: "cgst_sgst", supplierStateCode: "44", placeOfSupplyStateCode: "27", igst: 0, cgst: 360, sgst: 360 };
  const row = wrongHeadRow(f, GSTR2A_HEADER_ROW_INDEX);
  const reportedRow = row[0];
  assert(reportedRow === 7, "reported Row is the correct absolute Excel row (7), not the relative index (3)");
  assertRowMatches(SHEET_2A, reportedRow, (actual) => actual[0] === f.gstin && actual[1] === f.invoiceNumber, "dereferencing the reported Row lands on GSTIN \"" + f.gstin + "\", Invoice \"" + f.invoiceNumber + "\"");
  console.log("");
}

function runCrossSheetWrongHeadRow() {
  console.log("-- crossSheetWrongHeadRow (GST - Wrong Head, cross-sheet section) --\n");
  const f = {
    gstin: "11AAAAA0000A1Z1",
    identifier: "SHARED-INV-0",
    gstr2aRowIndex: 0,
    booksRowIndex: 0,
    gstr2aHead: "igst",
    booksHead: "cgst_sgst",
    expectedHead: "igst",
    incorrectSide: "books",
    gstr2a: { igst: 180, cgst: 0, sgst: 0 },
    books: { igst: 0, cgst: 90, sgst: 90 },
  };
  const row = crossSheetWrongHeadRow(f, GSTR2A_HEADER_ROW_INDEX, BOOKS_HEADER_ROW_INDEX, { gstr2a: "2A", books: "Books" });
  const reportedGstr2aRow = row[2];
  const reportedBooksRow = row[3];
  assert(reportedGstr2aRow === 4, "reported 2A Row is the correct absolute Excel row (4), not the relative index (0)");
  assert(reportedBooksRow === 6, "reported Books Row is the correct absolute Excel row (6), not the relative index (0)");
  assertRowMatches(SHEET_2A, reportedGstr2aRow, (actual) => actual[0] === f.gstin, "dereferencing the reported 2A Row lands on the same GSTIN (\"" + f.gstin + "\")");
  assertRowMatches(SHEET_BOOKS, reportedBooksRow, (actual) => actual[2] === f.gstin, "dereferencing the reported Books Row also lands on the same GSTIN, on the correct (different) sheet layout");
  console.log("");
}

function runRcmRow() {
  console.log("-- rcmRow (GST - RCM) --\n");
  const source = { sheetName: "2A", headerRowIndex: GSTR2A_HEADER_ROW_INDEX };
  const f = { rowIndex: 1, gstin: "22BBBBB1111B1Z2", invoiceNumber: "2A-INV-1", source: "2a_flag", matchedCategory: null };
  const amounts = { taxableValue: 2000, igst: 0, cgst: 180, sgst: 180 };
  const row = rcmRow(source, f, amounts);
  const reportedRow = row[1];
  assert(reportedRow === 5, "reported Row is the correct absolute Excel row (5), not the relative index (1)");
  assertRowMatches(SHEET_2A, reportedRow, (actual) => actual[0] === f.gstin && actual[1] === f.invoiceNumber, "dereferencing the reported Row lands on the row this RCM flag actually describes");
  console.log("");
}

function runIneligibleItcRow() {
  console.log("-- ineligibleItcRow (GST - Ineligible ITC) --\n");
  const source = { sheetName: "Books", headerRowIndex: BOOKS_HEADER_ROW_INDEX };
  const f = { rowIndex: 1, gstin: "55EEEEE4444E1Z5", invoiceNumber: "BOOKS-INV-1", matchedCategory: { label: "Motor vehicle", section: "17(5)(a)" }, matchedKeyword: "car" };
  const amounts = { taxableValue: 5000, igst: 0, cgst: 450, sgst: 450 };
  const row = ineligibleItcRow(source, f, amounts);
  const reportedRow = row[1];
  assert(reportedRow === 7, "reported Row is the correct absolute Excel row (7), not the relative index (1)");
  assertRowMatches(SHEET_BOOKS, reportedRow, (actual) => actual[1] === f.invoiceNumber && actual[2] === f.gstin, "dereferencing the reported Row lands on the row this Ineligible ITC flag actually describes");
  console.log("");
}

function runDuplicateInvoiceRow() {
  console.log("-- duplicateInvoiceRow (GST - Duplicate Invoices) -- the exact sheet from the real incident --\n");
  const source = { sheetName: "Books", headerRowIndex: BOOKS_HEADER_ROW_INDEX };
  const cluster = { gstin: "55EEEEE4444E1Z5", matchReason: "same_invoice_number_and_amount", dateSpreadDays: 0 };
  const m = { rowIndex: 1, identifier: "BOOKS-INV-1", amount: 5000 };
  const amounts = { taxableValue: 5000, igst: 0, cgst: 450, sgst: 450 };
  const row = duplicateInvoiceRow(source, 0, cluster, m, amounts);
  const reportedRow = row[3];
  assert(reportedRow === 7, "reported Row is the correct absolute Excel row (7), not the relative index (1) -- this is the exact bug that misled a manual check on a real client file");
  assertRowMatches(SHEET_BOOKS, reportedRow, (actual) => actual[1] === m.identifier && actual[2] === cluster.gstin, "dereferencing the reported Row lands on the row this duplicate-cluster member actually describes, not an unrelated row");
  console.log("");
}

function runRateMismatchRow() {
  console.log("-- rateMismatchRow (GST - Rate Mismatch) --\n");
  const source = { sheetName: "2A", headerRowIndex: GSTR2A_HEADER_ROW_INDEX };
  const f = { rowIndex: 3, gstin: "44DDDDD3333D1Z4", invoiceNumber: "2A-INV-3", taxableValue: 4000, ratePercent: 18, expectedTotalTax: 720, actualTotalTax: 720, difference: 0 };
  const row = rateMismatchRow(source, f);
  const reportedRow = row[1];
  assert(reportedRow === 7, "reported Row is the correct absolute Excel row (7), not the relative index (3)");
  assertRowMatches(SHEET_2A, reportedRow, (actual) => actual[0] === f.gstin && actual[1] === f.invoiceNumber, "dereferencing the reported Row lands on the row this rate-mismatch flag actually describes");
  console.log("");
}

function run() {
  runToAbsoluteExcelRow();
  runExtraInvoiceRow();
  runPossibleMatchRow();
  runWrongHeadRow();
  runCrossSheetWrongHeadRow();
  runRcmRow();
  runIneligibleItcRow();
  runDuplicateInvoiceRow();
  runRateMismatchRow();

  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All gst-report-writer row-reference checks passed.");
  }
}

run();
