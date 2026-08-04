/* global Excel */

/*
 * gst-report-writer.js
 * -----------------------
 * Turns the already-computed, already-validated output of the GST
 * reconciliation pipeline (src/gst-reconciliation/*.js — Step 1's vendor
 * match/mismatch, Step 2's four line-item rules, and the Summary netting)
 * into real worksheets inside the user's workbook — the same set of sheets
 * a senior analyst manually produces, each one self-explanatory (why a row
 * was flagged, which law/notification it rests on), so the user can review
 * the whole thing without rebuilding it from scratch.
 *
 * This file does NOT recompute or alter any reconciliation logic — it only
 * reads the pure functions' return values and formats them onto sheets.
 * Consistent with MENTOR's "propose, don't auto-execute" rule: nothing
 * here runs until the user has already seen a proposal and clicked accept
 * (see mentor-gst-reconciliation-ui.js) — this module is the "write it for
 * real" step, not the decision to do so.
 */

const GST_REPORT_SHEET_NAMES = [
  "GST - Extra in Books",
  "GST - Extra in 2A",
  "GST - Mismatch",
  "GST - Wrong Head",
  "GST - RCM",
  "GST - Ineligible ITC",
  "GST - Duplicate Invoices",
  "GST - Summary",
];

function toNumber(value) {
  if (typeof value === "number") return value;
  if (value === "" || value === null || value === undefined) return 0;
  const parsed = parseFloat(String(value).replace(/[,$£€¥%()]/g, ""));
  return isNaN(parsed) ? 0 : parsed;
}

function roundTo2(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

// Builds a row's identity key from its key columns. NOT `row[c] || ""` —
// a key column is very often a rowIndex, and rowIndex 0 (the first flagged
// row) is a legitimate, common value that `||` would wrongly treat as
// falsy and silently collapse to an empty string, corrupting that row's
// key and breaking the match.
function buildRowKey(row, keyColumnIndices) {
  return keyColumnIndices.map((c) => (row[c] === null || row[c] === undefined ? "" : String(row[c]).trim())).join("|");
}

// "cgst_sgst" -> "CGST+SGST", "igst" -> "IGST", "mixed" -> "MIXED"
function formatTaxHead(head) {
  return String(head).replace(/_/g, "+").toUpperCase();
}

// Turns a sheet-writer's per-item metadata (built alongside its `rows`
// array) into Top-Issues candidates with real, absolute output-sheet row
// positions — using `result.headerRowIdx` (returned by writeReportSheet)
// so a later click can select the EXACT row/rows this item ended up on,
// without the caller needing to re-derive row math itself. rowSpanStart
// is an index into the sheet's own `rows` array (0-based); rowSpan is 1
// for a normal row, or >1 for something like a Duplicate Invoices cluster
// that occupies several contiguous rows.
function attachTopIssueCandidates(result, sheetName, itemMetaList) {
  const topIssueCandidates = itemMetaList.map((m) => ({
    sheetName,
    rowIndex: result.headerRowIdx + 1 + m.rowSpanStart,
    rowSpan: m.rowSpan || 1,
    amount: m.amount,
    reason: m.reason,
    category: m.category,
    tier: m.tier,
  }));
  return { ...result, topIssueCandidates };
}

function formatDate(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return "";
  return d.toISOString().slice(0, 10);
}

// Looks up a flagged row's real tax amounts back on its source sheet — the
// four Step 2 detectors deliberately return only rowIndex + what each rule
// itself needed (a category, a flag source, an identifier), not full
// amounts, since that's not their concern. The report-writer's job is
// display, so it re-reads the row here rather than asking any detector to
// carry data it doesn't need for its own logic.
function getRowAmounts(values, headerRowIndex, columns, rowIndex) {
  const dataRows = headerRowIndex === -1 ? values : values.slice(headerRowIndex + 1);
  const row = dataRows[rowIndex];
  if (!row) return { taxableValue: 0, igst: 0, cgst: 0, sgst: 0, totalTax: 0 };
  const taxableValue = columns.taxableValue !== null ? toNumber(row[columns.taxableValue]) : 0;
  const igst = columns.igst !== null ? toNumber(row[columns.igst]) : 0;
  const cgst = columns.cgst !== null ? toNumber(row[columns.cgst]) : 0;
  const sgst = columns.sgst !== null ? toNumber(row[columns.sgst]) : 0;
  return { taxableValue, igst, cgst, sgst, totalTax: roundTo2(igst + cgst + sgst) };
}

/* ============================================================
   Generic sheet writer
   ============================================================ */

async function getOrCreateCleanSheet(context, sheetName) {
  let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
  await context.sync();
  if (!sheet.isNullObject) {
    sheet.delete(); // regenerate from scratch each run rather than merging with stale content
    await context.sync();
  }
  sheet = context.workbook.worksheets.add(sheetName);
  return sheet;
}

// Two columns appended to every review sheet, always MENTOR-blank on first
// write, never computed FROM anything — the only columns MENTOR treats as
// the user's own. Because MENTOR never writes real content into them,
// there's no ambiguity to resolve about whether a cell was "changed by the
// user" vs. "recomputed differently this run" (the same ambiguity that
// makes preserving arbitrary edited computed cells unsafe — a stale
// computed figure could get silently frozen in place and mistaken for
// fresh data). Whatever's in these two columns is unambiguously the
// user's, and is carried across regeneration by matching each row's key
// columns (see keyColumnIndices on writeReportSheet) — not by row
// position, which shifts every time the underlying data or rules change.
const REVIEWER_HEADERS = ["Reviewer Status", "Reviewer Note"];
const REVIEWER_STATUS_OPTIONS = ["", "Confirmed OK — no action needed", "Needs follow-up", "Resolved outside MENTOR", "Disputed — see note"];

// Reads whatever is currently in a sheet's Reviewer Status/Note columns
// (if the sheet exists and was previously written with them) BEFORE it
// gets deleted and recreated, keyed by keyColumnIndices — the same column
// positions the fresh rows will be keyed by. Finds the header row by
// matching its first cell against the known header text rather than
// assuming a fixed row index, since the explanation block above it can be
// one line longer/shorter between runs (e.g. a conditional "not
// applicable" note).
async function readPreservedReviewData(context, sheetName, headers, keyColumnIndices) {
  const empty = { preserved: new Map() };
  if (!keyColumnIndices) return empty;

  const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
  sheet.load("isNullObject");
  await context.sync();
  if (sheet.isNullObject) return empty;

  const usedRange = sheet.getUsedRangeOrNullObject();
  usedRange.load("values, isNullObject");
  await context.sync();
  if (usedRange.isNullObject) return empty;

  const values = usedRange.values;
  let headerRowIdx = -1;
  for (let r = 0; r < values.length; r++) {
    if (String(values[r][0]).trim() === headers[0]) {
      headerRowIdx = r;
      break;
    }
  }
  if (headerRowIdx === -1) return empty; // never written by a version of MENTOR that shares this header layout

  const oldHeaderRow = values[headerRowIdx];
  const statusColIdx = oldHeaderRow.indexOf("Reviewer Status");
  const noteColIdx = oldHeaderRow.indexOf("Reviewer Note");
  if (statusColIdx === -1 && noteColIdx === -1) return empty; // older sheet, predates this feature

  const preserved = new Map();
  for (let r = headerRowIdx + 1; r < values.length; r++) {
    const row = values[r];
    const status = statusColIdx !== -1 ? String(row[statusColIdx] || "").trim() : "";
    const note = noteColIdx !== -1 ? String(row[noteColIdx] || "").trim() : "";
    if (!status && !note) continue; // nothing reviewed on this row — nothing to preserve
    const key = buildRowKey(row, keyColumnIndices);
    preserved.set(key, { status, note });
  }
  return { preserved };
}

// explanationLines: array of strings shown as a header block (why this
// sheet exists, what law/notification it rests on). headers/rows: a plain
// table below that block — this function does no interpretation of the
// data itself. keyColumnIndices: column indices (into headers/rows, before
// the Reviewer columns are appended) that together uniquely identify a
// row across regenerations — e.g. GSTIN alone for a vendor-level sheet,
// or [Source Sheet, Row] for a line-item sheet, since row POSITION isn't
// stable (a row can disappear, or the set of flagged rows can reorder,
// between runs). Pass null/omit for sheets with no natural row identity
// (currently just the Summary sheet, which isn't row-per-item).
async function writeReportSheet(context, sheetName, explanationLines, headers, rows, keyColumnIndices) {
  const { preserved } = await readPreservedReviewData(context, sheetName, headers, keyColumnIndices);

  const fullHeaders = keyColumnIndices ? headers.concat(REVIEWER_HEADERS) : headers;
  let reappliedCount = 0;
  const fullRows = rows.map((row) => {
    if (!keyColumnIndices) return row;
    const key = buildRowKey(row, keyColumnIndices);
    const match = preserved.get(key);
    if (match) {
      reappliedCount++;
      preserved.delete(key); // consumed — whatever's left in `preserved` after this loop is orphaned
      return row.concat([match.status, match.note]);
    }
    return row.concat(["", ""]);
  });
  const orphanedCount = preserved.size;

  const sheet = await getOrCreateCleanSheet(context, sheetName);

  const explainRowCount = explanationLines.length;
  const headerRowIdx = explainRowCount + 1; // one blank row of separation

  // Header + data are written and column-autofit FIRST, before the
  // explanation block above them — sizing the explanation block's merged
  // row heights needs the REAL resulting column widths (read back below),
  // not a blind guess at how many characters fit per line. Row order on
  // the sheet is unaffected; this only changes which range gets written
  // to first.
  const headerRange = sheet.getRangeByIndexes(headerRowIdx, 0, 1, fullHeaders.length);
  headerRange.values = [fullHeaders];
  headerRange.format.font.bold = true;
  headerRange.format.fill.color = "#e9ecef";

  if (fullRows.length > 0) {
    const dataRange = sheet.getRangeByIndexes(headerRowIdx + 1, 0, fullRows.length, fullHeaders.length);
    dataRange.values = fullRows;

    if (keyColumnIndices) {
      const statusColIdx = fullHeaders.length - 2;
      const statusRange = sheet.getRangeByIndexes(headerRowIdx + 1, statusColIdx, fullRows.length, 1);
      try {
        statusRange.dataValidation.rule = { list: { inCellDropDown: true, source: REVIEWER_STATUS_OPTIONS.join(",") } };
      } catch (validationError) {
        console.error("MENTOR GST report: could not apply the Reviewer Status dropdown on '" + sheetName + "'", validationError);
      }
    }
  }

  // Autofit only over the header + data rows, NOT sheet.getUsedRange() —
  // the explanation block above still lives in column A, and Excel's
  // column-width autofit measures hidden-row content too, so including it
  // here would size column A to the longest explanation sentence instead
  // of the actual GSTIN/Row/Source-Sheet identifier values that column
  // holds on every one of these sheets.
  const contentRowCount = 1 + fullRows.length; // header row + data rows
  const contentRange = sheet.getRangeByIndexes(headerRowIdx, 0, contentRowCount, fullHeaders.length);
  contentRange.format.autofitColumns();

  if (explainRowCount > 0) {
    // Text goes in column A only; the merge below is what actually spans
    // it across columns — writing a single narrow column then merging
    // avoids needing to hand-construct a mergeWidth-wide values array.
    const explainTextRange = sheet.getRangeByIndexes(0, 0, explainRowCount, 1);
    explainTextRange.values = explanationLines.map((line) => [line]);

    // Merge each explanation line across A:C (or fewer if the sheet has
    // fewer than 3 data columns) so a full sentence has room to wrap
    // instead of overflowing into — or being clipped by — column A alone.
    // merge(true) merges each ROW of the range separately (one merged
    // cell per line), not the whole block into a single cell.
    const mergeWidth = Math.min(3, fullHeaders.length);
    const explainRange = sheet.getRangeByIndexes(0, 0, explainRowCount, mergeWidth);
    explainRange.merge(true);
    explainRange.format.font.italic = true;
    explainRange.format.font.color = "#555555";
    explainRange.format.wrapText = true;
    explainRange.format.verticalAlignment = "Top";

    // Row height, computed from the REAL merged column width (read back
    // just below) rather than autofitRows() alone — merged cells are a
    // known weak spot for Excel's row-height autofit (it can silently
    // no-op), and clipped legal text is exactly the failure this fix is
    // for. autofitRows() still runs first as a genuine best effort; the
    // computed height is then applied as a floor, never shrinking below
    // whatever autofitRows() already produced.
    explainRange.format.autofitRows();
    const widthColumns = [];
    for (let c = 0; c < mergeWidth; c++) widthColumns.push(sheet.getRangeByIndexes(headerRowIdx, c, 1, 1));
    widthColumns.forEach((r) => r.format.load("columnWidth"));
    // Read each row's height individually, NOT as one multi-row range —
    // autofitRows() may already have given each row a DIFFERENT height
    // based on its own wrapped content, so a single combined read across
    // explainRowCount non-uniform rows would be ambiguous.
    const rowHeightCells = [];
    for (let i = 0; i < explainRowCount; i++) {
      const r = sheet.getRangeByIndexes(i, 0, 1, 1);
      r.format.load("rowHeight");
      rowHeightCells.push(r);
    }
    await context.sync();

    const mergedWidthPoints = widthColumns.reduce((sum, r) => sum + r.format.columnWidth, 0);
    // Rough, deliberately conservative estimate of Calibri-11 characters
    // per point of width — erring toward MORE estimated lines (taller
    // rows) rather than fewer, since a little extra whitespace is a much
    // smaller problem than clipped legal citations.
    const estimatedCharsPerLine = Math.max(10, Math.floor(mergedWidthPoints / 6));
    const LINE_HEIGHT_POINTS = 15;

    for (let i = 0; i < explainRowCount; i++) {
      const lineCount = Math.max(1, Math.ceil(explanationLines[i].length / estimatedCharsPerLine));
      const computedHeight = lineCount * LINE_HEIGHT_POINTS;
      if (computedHeight > rowHeightCells[i].format.rowHeight) {
        sheet.getRangeByIndexes(i, 0, 1, 1).format.rowHeight = computedHeight;
      }
    }

    // Collapsed by default — the full explanation (law/notification basis)
    // stays available via Excel's own [+] outline control in the row
    // gutter, rather than pushing every row of real data down every time
    // the sheet is opened. Includes the blank separator row so the header
    // row sits right at the top once collapsed. Wrapped in try/catch since
    // row grouping isn't guaranteed to behave identically across every
    // Excel host — if it fails, the explanation is still there, just
    // uncollapsed, which is no worse than before this change.
    try {
      const collapsibleRows = sheet.getRangeByIndexes(0, 0, explainRowCount + 1, 1).getEntireRow();
      collapsibleRows.group(Excel.GroupOption.byRows);
      collapsibleRows.rowHidden = true;
    } catch (groupError) {
      console.error("MENTOR GST report: could not collapse explanation rows on '" + sheetName + "'", groupError);
    }
  }

  // Freeze the explanation block + header row so both stay visible while
  // scrolling through data below — freezing includes the (collapsed)
  // explanation rows too, so if the user expands them, they stay pinned
  // at the top rather than scrolling away immediately.
  sheet.freezePanes.freezeRows(headerRowIdx + 1);

  return { sheet, reappliedCount, orphanedCount, headerRowIdx };
}

/* ============================================================
   Step 1 vendor-level sheets: Extra in Books / Extra in 2A / Mismatch
   ============================================================ */

const VENDOR_HEADERS = ["GSTIN", "2A Taxable Value", "2A IGST", "2A CGST", "2A SGST", "Books Taxable Value", "Books IGST", "Books CGST", "Books SGST", "Diff IGST (2A − Books)", "Diff CGST (2A − Books)", "Diff SGST (2A − Books)"];

function vendorRow(v) {
  return [v.gstin, v.gstr2a.taxableValue, v.gstr2a.igst, v.gstr2a.cgst, v.gstr2a.sgst, v.books.taxableValue, v.books.igst, v.books.cgst, v.books.sgst, v.differences.igst, v.differences.cgst, v.differences.sgst];
}

async function writeExtraInBooksSheet(context, comparisonSummary, sheetNames) {
  const vendors = comparisonSummary.vendors.filter((v) => v.status === "missing_in_gstr2a");
  const explanation = [
    "Vendors present in '" + sheetNames.books + "' with no matching GSTIN in '" + sheetNames.gstr2a + "'.",
    "This means the supplier hasn't (yet) reported this invoice in their GSTR-1, or filed it under a different/incorrect GSTIN.",
    "Since ITC eligibility requires the credit to be reflected in the supplier's return (CGST Act Section 16(2)(aa) / Rule 36(4)), these amounts are NOT counted as currently supportable in the Summary sheet until the supplier's filing is confirmed.",
    "Not auto-applied — review each row and follow up with the supplier before claiming.",
  ];
  const result = await writeReportSheet(context, "GST - Extra in Books", explanation, VENDOR_HEADERS, vendors.map(vendorRow), [0]); // key: GSTIN
  const itemMeta = vendors.map((v, i) => ({
    rowSpanStart: i,
    amount: roundTo2(v.books.igst + v.books.cgst + v.books.sgst),
    reason: "Not yet reflected in 2A — supplier may not have filed, or filed under a different GSTIN",
    category: "Extra in Books",
    tier: 2,
  }));
  return attachTopIssueCandidates(result, "GST - Extra in Books", itemMeta);
}

async function writeExtraIn2ASheet(context, comparisonSummary, sheetNames) {
  const vendors = comparisonSummary.vendors.filter((v) => v.status === "missing_in_books");
  const explanation = [
    "Vendors present in '" + sheetNames.gstr2a + "' with no matching GSTIN in '" + sheetNames.books + "'.",
    "This means the government return shows an invoice this workbook's purchase register has no record of — it may not have been booked yet, may belong to a different GSTIN by mistake, or may not actually be this taxpayer's purchase.",
    "Not counted as currently supportable in the Summary sheet until investigated — claiming credit the books don't support risks a later reversal.",
    "Not auto-applied — review each row before taking any action.",
  ];
  const result = await writeReportSheet(context, "GST - Extra in 2A", explanation, VENDOR_HEADERS, vendors.map(vendorRow), [0]); // key: GSTIN
  const itemMeta = vendors.map((v, i) => ({
    rowSpanStart: i,
    amount: roundTo2(v.gstr2a.igst + v.gstr2a.cgst + v.gstr2a.sgst),
    reason: "Not yet reflected in Books — not booked yet, wrong GSTIN, or not this taxpayer's purchase",
    category: "Extra in 2A",
    tier: 2,
  }));
  return attachTopIssueCandidates(result, "GST - Extra in 2A", itemMeta);
}

async function writeMismatchSheet(context, comparisonSummary, sheetNames) {
  const vendors = comparisonSummary.vendors.filter((v) => v.status === "discrepancy");
  const explanation = [
    "Vendors present in BOTH '" + sheetNames.gstr2a + "' and '" + sheetNames.books + "', where IGST/CGST/SGST totals differ by more than ₹" + comparisonSummary.tolerance + " (Taxable Value alone is not used as a match criterion — see gstr2a-vs-books-reconciler.js for why).",
    "The Summary sheet treats the LOWER of the two sides as currently supportable, and reports the gap separately, rather than assuming either side is correct.",
    "Not auto-applied — review each row before adjusting any claim.",
  ];
  const result = await writeReportSheet(context, "GST - Mismatch", explanation, VENDOR_HEADERS, vendors.map(vendorRow), [0]); // key: GSTIN
  const itemMeta = vendors.map((v, i) => ({
    rowSpanStart: i,
    amount: roundTo2(Math.abs(v.differences.igst + v.differences.cgst + v.differences.sgst)),
    reason: "2A and Books tax totals differ by more than the ₹" + comparisonSummary.tolerance + " tolerance for this GSTIN",
    category: "Mismatch",
    tier: 2,
  }));
  return attachTopIssueCandidates(result, "GST - Mismatch", itemMeta);
}

/* ============================================================
   Step 2: Wrong Head
   ============================================================ */

async function writeWrongHeadSheet(context, wrongHeadResult, source) {
  const headers = ["Row (in " + source.sheetName + ")", "GSTIN", "Invoice Number", "Issue", "Expected Head", "Actual Head", "Supplier State", "Place of Supply State", "IGST", "CGST", "SGST"];
  const explanation = [
    "Rows in '" + source.sheetName + "' where the tax head charged (IGST vs CGST+SGST) doesn't match what the supplier's state and the invoice's Place of Supply require in law.",
    "Same state as Place of Supply -> CGST+SGST is required (IGST Act Section 8). Different state -> IGST is required (IGST Act Section 7).",
    "This does NOT change the amount of credit available — same rupees, wrong head — so it is reported here for correction, not subtracted anywhere in the Summary sheet.",
    "Not auto-applied — the correct filing/amendment is the user's call.",
  ];
  if (!wrongHeadResult.applicable) {
    const result = await writeReportSheet(context, "GST - Wrong Head", explanation.concat(["Not applicable to this workbook: " + wrongHeadResult.reason + "."]), headers, [], [0]);
    return { ...result, topIssueCandidates: [] };
  }
  const rows = wrongHeadResult.flagged.map((f) => [f.rowIndex, f.gstin, f.invoiceNumber, f.issue, f.expected, f.actual, f.supplierStateCode, f.placeOfSupplyStateCode, f.igst, f.cgst, f.sgst]);
  const result = await writeReportSheet(context, "GST - Wrong Head", explanation, headers, rows, [0]); // key: Row (source is always 2A)
  const itemMeta = wrongHeadResult.flagged.map((f, i) => ({
    rowSpanStart: i,
    amount: roundTo2(f.igst + f.cgst + f.sgst),
    reason: "Charged " + formatTaxHead(f.actual) + ", should be " + formatTaxHead(f.expected) + " per Place of Supply",
    category: "Wrong Head",
    tier: 3,
  }));
  return attachTopIssueCandidates(result, "GST - Wrong Head", itemMeta);
}

/* ============================================================
   Step 2: RCM (combined across both recognized sheets)
   ============================================================ */

async function writeRcmSheet(context, rcmBySource, materialityThreshold) {
  const headers = ["Source Sheet", "Row", "GSTIN", "Invoice/Voucher Number", "Taxable Value", "IGST", "CGST", "SGST", "Category", "Notification", "Basis", "Needs Manual Confirmation?"];
  const explanation = [
    "Rows across both recognized sheets flagged as likely Reverse Charge Mechanism (RCM) supplies — under Section 9(3)/9(4) CGST Act and Section 5(3)/5(4) IGST Act, the RECIPIENT (not the supplier) is liable to pay this tax.",
    "'Basis' shows whether this came from GSTR-2A's own 'Supply Attract Reverse Charge' flag (the government's own determination, trusted first) or a fallback keyword match against a genuine expense-description field (see rcm-config.js for the notified category list).",
    "RCM tax is self-assessed and paid by the recipient through a separate mechanism, not claimed as ordinary vendor-matched ITC — these rows are shown for review, NOT subtracted from the Summary sheet's net figure, to avoid double-counting.",
    "Not auto-applied — verify RCM liability is being discharged and claimed through the correct GSTR-3B table.",
  ];

  const rows = [];
  const itemMeta = [];
  let anyApplicable = false;
  for (const source of rcmBySource) {
    if (!source.result.flagSignalAvailable && !source.result.keywordFallbackAvailable) continue;
    anyApplicable = true;
    for (const f of source.result.flagged) {
      const amounts = getRowAmounts(source.values, source.headerRowIndex, source.columns, f.rowIndex);
      rows.push([
        source.sheetName,
        f.rowIndex,
        f.gstin || "(unregistered)",
        f.invoiceNumber,
        amounts.taxableValue,
        amounts.igst,
        amounts.cgst,
        amounts.sgst,
        f.matchedCategory ? f.matchedCategory.label : "(per 2A flag)",
        f.matchedCategory ? f.matchedCategory.notification : "GSTR-2A 'Supply Attract Reverse Charge' flag",
        f.source,
        f.matchedCategory && f.matchedCategory.requiresManualConfirmation ? "Yes" : "No",
      ]);
      itemMeta.push({
        rowSpanStart: rows.length - 1,
        amount: amounts.totalTax,
        reason: "RCM: " + (f.matchedCategory ? f.matchedCategory.label : "flagged by GSTR-2A's own reverse-charge indicator") + " (" + source.sheetName + ")",
        category: "RCM",
        tier: amounts.totalTax >= materialityThreshold ? 1 : 3,
      });
    }
  }
  if (!anyApplicable) {
    explanation.push("Not applicable to this workbook: neither sheet has an RCM flag column or a genuine expense-description field to check against.");
  }
  const result = await writeReportSheet(context, "GST - RCM", explanation, headers, rows, [0, 1]); // key: Source Sheet + Row
  return attachTopIssueCandidates(result, "GST - RCM", itemMeta);
}

/* ============================================================
   Step 2: Ineligible ITC (combined across both recognized sheets)
   ============================================================ */

async function writeIneligibleItcSheet(context, itcBySource, materialityThreshold) {
  const headers = ["Source Sheet", "Row", "GSTIN", "Invoice/Voucher Number", "Taxable Value", "IGST", "CGST", "SGST", "Category", "Section", "Matched Keyword"];
  const explanation = [
    "Rows across both recognized sheets whose expense description matches a blocked-credit category under Section 17(5) CGST Act — ITC is disallowed for these regardless of whether GST was otherwise validly charged (see ineligible-itc-config.js for the category list).",
    "Keyword matching only ever runs against a field verified to genuinely describe the transaction — never a vendor/company name field, even when a keyword coincidentally appears inside one.",
    "This IS subtracted from the Summary sheet's net provisionally-eligible figure — Section 17(5) actually blocks the credit, unlike Wrong Head/RCM above.",
    "Not auto-applied — every match should be reviewed against the actual invoice before adjusting any claim; keyword matching is not a legal determination on its own.",
  ];

  const rows = [];
  const itemMeta = [];
  const notApplicableReasons = [];
  for (const source of itcBySource) {
    if (!source.result.applicable) {
      notApplicableReasons.push(source.sheetName + ": " + source.result.reason);
      continue;
    }
    for (const f of source.result.flagged) {
      const amounts = getRowAmounts(source.values, source.headerRowIndex, source.columns, f.rowIndex);
      rows.push([source.sheetName, f.rowIndex, f.gstin, f.invoiceNumber, amounts.taxableValue, amounts.igst, amounts.cgst, amounts.sgst, f.matchedCategory.label, f.matchedCategory.section, f.matchedKeyword]);
      itemMeta.push({
        rowSpanStart: rows.length - 1,
        amount: amounts.totalTax,
        reason: "Ineligible ITC: " + f.matchedCategory.label + " (" + f.matchedCategory.section + ", " + source.sheetName + ")",
        category: "Ineligible ITC",
        tier: amounts.totalTax >= materialityThreshold ? 1 : 3,
      });
    }
  }
  if (notApplicableReasons.length > 0) {
    explanation.push("Not applicable on: " + notApplicableReasons.join("; ") + ".");
  }
  const result = await writeReportSheet(context, "GST - Ineligible ITC", explanation, headers, rows, [0, 1]); // key: Source Sheet + Row
  return attachTopIssueCandidates(result, "GST - Ineligible ITC", itemMeta);
}

/* ============================================================
   Step 2: Duplicate Invoices (combined across both recognized sheets)
   ============================================================ */

async function writeDuplicateInvoicesSheet(context, dupBySource) {
  const headers = ["Source Sheet", "Cluster #", "GSTIN", "Row", "Identifier", "Taxable Value", "IGST", "CGST", "SGST", "Match Reason", "Date Spread (days)"];
  const explanation = [
    "Rows across both recognized sheets that look like the same invoice entered more than once: same GSTIN, and either the same invoice/voucher number or the same amount, within a 15-day window (see duplicate-invoice-detector.js for why that window is deliberately tight).",
    "Each numbered cluster groups every row believed to be the same duplicated invoice.",
    "The Summary sheet subtracts the DUPLICATE sheet's over-claimed portion (every cluster member beyond the first) from the net provisionally-eligible figure — the '2A' sheet's own duplicates are shown here for review but not subtracted, since a duplicate filing by the supplier doesn't by itself mean the taxpayer over-claimed.",
    "Not auto-applied — confirm each cluster is genuinely a duplicate (not two separate, coincidentally similar invoices) before removing anything.",
  ];

  const rows = [];
  const itemMeta = []; // one entry per CLUSTER (not per member row) — a cluster is the reviewable unit here
  let anyApplicable = false;
  const notApplicableReasons = [];
  for (const source of dupBySource) {
    if (!source.result.applicable) {
      notApplicableReasons.push(source.sheetName + ": " + source.result.reason);
      continue;
    }
    anyApplicable = true;
    source.result.clusters.forEach((cluster, clusterIdx) => {
      const rowSpanStart = rows.length; // cluster members are pushed contiguously below, so this is a stable span start
      const memberTaxTotals = [];
      for (const m of cluster.members) {
        const amounts = getRowAmounts(source.values, source.headerRowIndex, source.columns, m.rowIndex);
        memberTaxTotals.push(amounts.totalTax);
        rows.push([source.sheetName, clusterIdx + 1, cluster.gstin, m.rowIndex, m.identifier, amounts.taxableValue, amounts.igst, amounts.cgst, amounts.sgst, cluster.matchReason, cluster.dateSpreadDays]);
      }
      // Same "every occurrence beyond the single legitimate one" over-claim
      // math as sumDuplicateOverclaim() below — used here just to size/rank
      // the issue, not to double-subtract anything from the Summary sheet.
      const overclaim = memberTaxTotals.reduce((a, b) => a + b, 0) - Math.max(...memberTaxTotals);
      itemMeta.push({
        rowSpanStart,
        rowSpan: cluster.members.length,
        amount: roundTo2(overclaim),
        reason:
          "Cluster of " + cluster.members.length + " entries, matched on " + cluster.matchReason.replace(/_/g, " ") + " (" + source.sheetName + (source.role === "purchase_register" ? " — reduces claimable ITC" : " — informational, supplier-side" ) + ")",
        category: "Duplicate Invoices",
        tier: 3,
      });
    });
  }
  if (!anyApplicable) {
    explanation.push("Not applicable to this workbook: " + (notApplicableReasons.join("; ") || "no GSTIN/date columns identified on either sheet") + ".");
  }
  const result = await writeReportSheet(context, "GST - Duplicate Invoices", explanation, headers, rows, [0, 3]); // key: Source Sheet + Row (NOT Cluster # — cluster numbering isn't stable across runs)
  return attachTopIssueCandidates(result, "GST - Duplicate Invoices", itemMeta);
}

/* ============================================================
   Step 2 totals for the Summary sheet (amounts the Summary needs,
   computed here from enriched rows rather than inside the pure
   gst-summary-builder.js, which only does arithmetic on numbers it's handed)
   ============================================================ */

function sumIneligibleItcAmount(itcBySource) {
  let amount = 0;
  let rowCount = 0;
  for (const source of itcBySource) {
    if (!source.result.applicable) continue;
    for (const f of source.result.flagged) {
      amount += getRowAmounts(source.values, source.headerRowIndex, source.columns, f.rowIndex).totalTax;
      rowCount++;
    }
  }
  return { amount: roundTo2(amount), rowCount };
}

function sumRcmAmount(rcmBySource) {
  let amount = 0;
  let rowCount = 0;
  for (const source of rcmBySource) {
    for (const f of source.result.flagged) {
      amount += getRowAmounts(source.values, source.headerRowIndex, source.columns, f.rowIndex).totalTax;
      rowCount++;
    }
  }
  return { amount: roundTo2(amount), rowCount };
}

// Only the Books-side clusters reduce the net figure — see the Duplicate
// Invoices sheet's explanation block for why.
function sumDuplicateOverclaim(dupBySource) {
  let amount = 0;
  let clusterCount = 0;
  const booksSource = dupBySource.find((s) => s.role === "purchase_register");
  if (!booksSource || !booksSource.result.applicable) return { amount: 0, clusterCount: 0 };

  for (const cluster of booksSource.result.clusters) {
    clusterCount++;
    const taxTotals = cluster.members.map((m) => getRowAmounts(booksSource.values, booksSource.headerRowIndex, booksSource.columns, m.rowIndex).totalTax);
    const sum = taxTotals.reduce((a, b) => a + b, 0);
    const max = Math.max(...taxTotals);
    amount += sum - max; // every occurrence beyond the single legitimate one
  }
  return { amount: roundTo2(amount), clusterCount };
}

function wrongHeadRowCount(wrongHeadResult) {
  return wrongHeadResult.applicable ? wrongHeadResult.flagged.length : 0;
}

/* ============================================================
   Step 2 totals bundle, for gst-summary-builder.js
   ============================================================ */

function buildStepTwoTotals(wrongHeadResult, rcmBySource, itcBySource, dupBySource) {
  const ineligible = sumIneligibleItcAmount(itcBySource);
  const rcm = sumRcmAmount(rcmBySource);
  const duplicates = sumDuplicateOverclaim(dupBySource);
  return {
    wrongHead: { rowCount: wrongHeadRowCount(wrongHeadResult) },
    rcm: { rowCount: rcm.rowCount, amount: rcm.amount },
    ineligibleItc: { rowCount: ineligible.rowCount, amount: ineligible.amount },
    duplicates: { clusterCount: duplicates.clusterCount, overclaimAmount: duplicates.amount },
  };
}

/* ============================================================
   Summary sheet
   ============================================================ */

async function writeSummarySheet(context, summary, sheetNames) {
  const sheet = await getOrCreateCleanSheet(context, "GST - Summary");

  const rows = [
    ["GSTR-2A vs Books — ITC Reconciliation Summary", ""],
    ["2A sheet", sheetNames.gstr2a],
    ["Books sheet", sheetNames.books],
    ["Tolerance applied to tax-field matching", "₹" + summary.tolerance],
    ["", ""],
    ["Vendor match status (Step 1)", ""],
    ["  Matched", summary.vendorCounts.matched],
    ["  Discrepancy (Mismatch)", summary.vendorCounts.discrepancy],
    ["  Missing in Books (Extra in 2A)", summary.vendorCounts.missing_in_books],
    ["  Missing in 2A (Extra in Books)", summary.vendorCounts.missing_in_gstr2a],
    ["  Total GSTINs", summary.totalGstins],
    ["", ""],
    ["Tax figures (IGST + CGST + SGST combined)", ""],
    ["Total tax per 2A", summary.totalTaxPer2A],
    ["Total tax per Books", summary.totalTaxPerBooks],
    ["", ""],
    ["Provisionally supportable base", ""],
    ["  = per vendor, per tax head, the LOWER of 2A and Books, summed", summary.provisionallySupportable],
    ["", ""],
    ["Less: Ineligible ITC (Section 17(5) matches, Step 2)", "-" + summary.lessIneligibleItc.amount + "  (" + summary.lessIneligibleItc.rowCount + " row(s) — see 'GST - Ineligible ITC')"],
    ["Less: Duplicate invoice over-claims (Books side, Step 2)", "-" + summary.lessDuplicateOverclaim.amount + "  (" + summary.lessDuplicateOverclaim.clusterCount + " cluster(s) — see 'GST - Duplicate Invoices')"],
    ["", ""],
    ["NET PROVISIONALLY ELIGIBLE ITC", summary.netProvisionallyEligible],
    ["", ""],
    ["Not netted above — needs investigation before any claim", ""],
    ["  Extra in 2A (not yet in Books)", summary.extraIn2A.amount + "  (" + summary.extraIn2A.count + " vendor(s) — see 'GST - Extra in 2A')"],
    ["  Extra in Books (not yet in 2A)", summary.extraInBooks.amount + "  (" + summary.extraInBooks.count + " vendor(s) — see 'GST - Extra in Books')"],
    ["  Mismatch gap (2A vs Books, unresolved)", summary.mismatch.unsupportedAmount + "  (" + summary.mismatch.count + " vendor(s) — see 'GST - Mismatch')"],
    ["", ""],
    ["Informational only — not netted, doesn't change the ITC amount", ""],
    ["  Wrong Head rows (correct head, same rupees)", summary.informational.wrongHeadRowCount + "  (see 'GST - Wrong Head')"],
    ["  RCM rows (self-assessed by recipient, separate mechanism)", summary.informational.rcmRowCount + " row(s), ₹" + summary.informational.rcmAmount + " total (see 'GST - RCM')"],
    ["", ""],
    ["This is a PROPOSED reconciliation for review — nothing here has been filed, claimed, or reversed automatically. Every figure traces back to one of the sheets referenced above.", ""],
  ];

  const range = sheet.getRangeByIndexes(0, 0, rows.length, 2);
  range.values = rows;
  sheet.getRange("A1").format.font.bold = true;
  sheet.getRange("A1").format.font.size = 13;
  sheet.getRange("A6").format.font.bold = true;
  sheet.getRange("A13").format.font.bold = true;
  sheet.getRange("A17").format.font.bold = true;
  sheet.getRange("A23:B23").format.font.bold = true;
  sheet.getRange("A23:B23").format.fill.color = "#d4edda";
  sheet.getRange("A26").format.font.bold = true;
  sheet.getRange("A31").format.font.bold = true;
  sheet.getUsedRange().format.autofitColumns();
  sheet.freezePanes.freezeRows(1); // title row stays visible while scrolling the summary
  return sheet;
}

/* ============================================================
   Top-level orchestration
   ============================================================ */

// data: {
//   sheetNames: { gstr2a, books },
//   comparisonSummary,                 // Step 1 output, unmodified
//   wrongHeadResult,                   // detectWrongHead() on the 2A sheet
//   rcmBySource, itcBySource, dupBySource,  // [{ sheetName, role, values, headerRowIndex, columns, result }, ...]
// }
async function writeGstReconciliationReport(context, data) {
  const stepTwoTotals = buildStepTwoTotals(data.wrongHeadResult, data.rcmBySource, data.itcBySource, data.dupBySource);
  const { buildGstSummary } = require("../gst-reconciliation/gst-summary-builder.js");
  const summary = buildGstSummary(data.comparisonSummary, stepTwoTotals);

  // Materiality threshold for Priority-1 RCM/Ineligible-ITC gating in the
  // Top Issues ranking (see mentor-gst-reconciliation-ui.js's sidebar) — a
  // PERCENTAGE of total tax per 2A, not a fixed rupee figure, so it scales
  // with the size of the actual dataset instead of an arbitrary number
  // baked in for one company. 1% is a conservative, commonly-used starting
  // point for review materiality. RCM/Ineligible-ITC items below it still
  // appear (Priority 3, reachable via "Show all") — they just don't crowd
  // out the headline top-5 the way a genuinely material item should.
  const MATERIALITY_PERCENT = 0.01;
  const materialityThreshold = roundTo2(summary.totalTaxPer2A * MATERIALITY_PERCENT);

  const results = [
    await writeExtraInBooksSheet(context, data.comparisonSummary, data.sheetNames),
    await writeExtraIn2ASheet(context, data.comparisonSummary, data.sheetNames),
    await writeMismatchSheet(context, data.comparisonSummary, data.sheetNames),
    await writeWrongHeadSheet(context, data.wrongHeadResult, { sheetName: data.sheetNames.gstr2a }),
    await writeRcmSheet(context, data.rcmBySource, materialityThreshold),
    await writeIneligibleItcSheet(context, data.itcBySource, materialityThreshold),
    await writeDuplicateInvoicesSheet(context, data.dupBySource),
  ];
  await writeSummarySheet(context, summary, data.sheetNames); // no row-level identity — not part of review-note preservation or Top Issues
  await context.sync();

  // Reviewer Status/Note preservation totals across every sheet that has
  // them — surfaced to the caller so the "Regenerate" confirmation can
  // tell the user how many of their prior review notes actually survived,
  // rather than silently carrying them forward with no visible confirmation.
  const reviewPreservation = results.reduce(
    (acc, r) => ({ reappliedCount: acc.reappliedCount + r.reappliedCount, orphanedCount: acc.orphanedCount + r.orphanedCount }),
    { reappliedCount: 0, orphanedCount: 0 }
  );

  // Rank: tier ascending (1 = most urgent), then amount descending within
  // a tier. Sub-threshold RCM/Ineligible-ITC items fall to tier 3 alongside
  // Wrong Head/Duplicate Invoices rather than disappearing — see each
  // write*Sheet function above for exactly how tier is assigned.
  const topIssues = results.flatMap((r) => r.topIssueCandidates).sort((a, b) => a.tier - b.tier || b.amount - a.amount);

  return { summary, reviewPreservation, topIssues, materialityThreshold };
}

/* ============================================================
   Point-and-explain: given a sheet name and a row's cell values (read
   directly off the live sheet — nothing recomputed, nothing re-derived
   from source data), returns which rule fired, its legal basis, and the
   row's Reviewer Status/Note if any. Purely a read/format function; no
   Excel calls of its own, so the caller decides when/how to read the row.

   Column indices below mirror each write*Sheet function's own `headers`
   array EXACTLY (VENDOR_HEADERS is a shared constant already; the other
   four are defined inline in their own write*Sheet function above) — if a
   header layout ever changes there, this must change with it. Reviewer
   Status/Note are always the last two columns on every sheet this
   function handles, since writeReportSheet always appends them last.
   ============================================================ */

const GST_EXPLAIN_SHEET_NAMES = new Set(["GST - Extra in Books", "GST - Extra in 2A", "GST - Mismatch", "GST - Wrong Head", "GST - RCM", "GST - Ineligible ITC", "GST - Duplicate Invoices"]);

function explainVendorRow(sheetName, row, kind) {
  const gstin = row[0];
  const diffIgst = row[9], diffCgst = row[10], diffSgst = row[11];
  if (kind === "extraInBooks") {
    return {
      ruleLabel: "Extra in Books",
      whichRuleFired: "GSTIN " + gstin + " appears in Books but not in 2A — the supplier may not have filed this invoice yet, or filed it under a different GSTIN.",
      legalCitation: "ITC eligibility requires the credit to be reflected in the supplier's GSTR-1/2A (CGST Act Section 16(2)(aa) / Rule 36(4)).",
    };
  }
  if (kind === "extraIn2A") {
    return {
      ruleLabel: "Extra in 2A",
      whichRuleFired: "GSTIN " + gstin + " appears in 2A but not in Books — not yet booked, filed under a different GSTIN, or not this taxpayer's purchase.",
      legalCitation: "Claiming credit the books don't independently support risks a later reversal (CGST Act Section 16).",
    };
  }
  return {
    ruleLabel: "Mismatch",
    whichRuleFired: "GSTIN " + gstin + " — 2A and Books tax totals differ (Diff IGST=" + diffIgst + ", CGST=" + diffCgst + ", SGST=" + diffSgst + ").",
    legalCitation: "Match is determined on IGST/CGST/SGST only, not Taxable Value — see gstr2a-vs-books-reconciler.js's docstring for why.",
  };
}

function explainWrongHeadRow(row) {
  const issue = row[3], expected = row[4], actual = row[5], supplierState = row[6], posState = row[7];
  return {
    ruleLabel: "Wrong Head",
    whichRuleFired: "Issue: " + issue + " — charged " + formatTaxHead(actual) + ", should be " + formatTaxHead(expected) + " (supplier state " + supplierState + " vs. Place of Supply " + posState + ").",
    legalCitation: "Same state as Place of Supply -> CGST+SGST required (IGST Act Section 8). Different state -> IGST required (IGST Act Section 7).",
  };
}

function explainRcmRow(row) {
  const category = row[8], notification = row[9], basis = row[10], needsConfirmation = row[11];
  return {
    ruleLabel: "RCM (Reverse Charge Mechanism)",
    whichRuleFired: "Matched category: " + category + " (basis: " + basis + ")" + (needsConfirmation === "Yes" ? " — needs manual confirmation, this keyword alone can't fully confirm the case" : "") + ".",
    legalCitation: notification + " — Section 9(3)/9(4) CGST Act, Section 5(3)/5(4) IGST Act (recipient, not supplier, is liable).",
  };
}

function explainIneligibleItcRow(row) {
  const category = row[8], section = row[9], keyword = row[10];
  return {
    ruleLabel: "Ineligible ITC",
    whichRuleFired: "Matched category: " + category + " (matched keyword: '" + keyword + "').",
    legalCitation: section + " — blocked credit regardless of whether GST was otherwise validly charged.",
  };
}

function explainDuplicateRow(row) {
  const clusterNum = row[1], matchReason = row[9], dateSpread = row[10];
  return {
    ruleLabel: "Duplicate Invoices",
    whichRuleFired: "Cluster #" + clusterNum + " — matched on " + String(matchReason).replace(/_/g, " ") + ", " + dateSpread + " day(s) apart.",
    legalCitation: "ITC may only be claimed once per genuine invoice (CGST Act Section 16) — confirm this is a real duplicate, not two separate invoices, before removing anything.",
  };
}

// Finds the header row the same way readPreservedReviewData() does — by
// matching known header text — rather than assuming a fixed row index,
// since the (collapsed) explanation block above it varies in length
// between sheets/runs. Wrong Head's own first header cell has a dynamic
// segment (the 2A sheet's actual name baked into "Row (in X)"), so it's
// matched on its stable SECOND column ("GSTIN") instead.
function findGstSheetHeaderRowIndex(sheetName, values) {
  const matchCol = sheetName === "GST - Wrong Head" ? 1 : 0;
  const matchValue = sheetName === "GST - RCM" || sheetName === "GST - Ineligible ITC" || sheetName === "GST - Duplicate Invoices" ? "Source Sheet" : "GSTIN";
  for (let r = 0; r < values.length; r++) {
    if (String(values[r][matchCol]).trim() === matchValue) return r;
  }
  return -1;
}

// values: the sheet's FULL used-range values, read fresh by the caller —
// so an edit the user just made (including to Reviewer Status/Note) shows
// up immediately, not a stale copy. selectedRowIndex: the 0-based row the
// user actually has selected right now, in that SAME values array (i.e.
// absolute sheet position, not relative to the data block). Returns null
// if the sheet isn't one of the 7 explainable ones, or the selection isn't
// actually on a data row (e.g. the header row itself, the — possibly
// collapsed — explanation block above it, or past the last data row).
function explainGstRow(sheetName, values, selectedRowIndex) {
  if (!GST_EXPLAIN_SHEET_NAMES.has(sheetName)) return null;

  const headerRowIndex = findGstSheetHeaderRowIndex(sheetName, values);
  if (headerRowIndex === -1) return null; // sheet predates this feature or layout doesn't match — nothing to explain
  if (selectedRowIndex <= headerRowIndex || selectedRowIndex >= values.length) return null;

  const row = values[selectedRowIndex];

  let rule;
  if (sheetName === "GST - Extra in Books") rule = explainVendorRow(sheetName, row, "extraInBooks");
  else if (sheetName === "GST - Extra in 2A") rule = explainVendorRow(sheetName, row, "extraIn2A");
  else if (sheetName === "GST - Mismatch") rule = explainVendorRow(sheetName, row, "mismatch");
  else if (sheetName === "GST - Wrong Head") rule = explainWrongHeadRow(row);
  else if (sheetName === "GST - RCM") rule = explainRcmRow(row);
  else if (sheetName === "GST - Ineligible ITC") rule = explainIneligibleItcRow(row);
  else if (sheetName === "GST - Duplicate Invoices") rule = explainDuplicateRow(row);

  const reviewerStatus = String(row[row.length - 2] || "").trim();
  const reviewerNote = String(row[row.length - 1] || "").trim();

  return { ...rule, reviewerStatus, reviewerNote };
}

/* ============================================================
   Vendor-name LLM suggestions (Ineligible ITC only) — a separate,
   optional, explicitly user-triggered addition, never part of the normal
   8-sheet generation. APPENDS to the existing 'GST - Ineligible ITC'
   sheet rather than going through writeReportSheet's delete-and-recreate
   flow, specifically so it never wipes out any Reviewer Status/Note the
   user already set on the description-based rows above it. Re-running
   this replaces only its own previously-appended section (found by
   searching for its own marker text), never anything above it.

   Output is clearly, visually separated from the validated description-
   based findings above (distinct amber styling, its own sub-header) and
   is never netted into the Summary sheet — this is a screening aid for
   manual review, not a finding with the same standing as the rest of the
   pipeline.
   ============================================================ */

const ITC_SUGGESTIONS_MARKER = "Vendor-name based suggestions (unconfirmed — review invoice details before treating as ineligible)";
const ITC_SUGGESTIONS_HEADERS = ["GSTIN", "Vendor Name", "Likely Category", "Sub-Clause", "Reasoning"];

// screeningResult: { results: [{gstin, vendorName, flagged, category, subClause, reasoning}, ...] }
// — the parsed output of screenVendorsForIneligibleItc() (see
// server/anthropic-vendor-itc-screening.js), passed through unmodified.
async function appendIneligibleItcVendorSuggestions(context, screeningResult) {
  const sheet = context.workbook.worksheets.getItemOrNullObject("GST - Ineligible ITC");
  sheet.load("isNullObject");
  await context.sync();
  if (sheet.isNullObject) {
    throw new Error("'GST - Ineligible ITC' sheet doesn't exist — generate the review sheets first");
  }

  const usedRange = sheet.getUsedRangeOrNullObject();
  usedRange.load("values, isNullObject");
  await context.sync();

  let insertAtRow;
  if (usedRange.isNullObject) {
    insertAtRow = 0;
  } else {
    const values = usedRange.values;
    let markerRow = -1;
    for (let r = 0; r < values.length; r++) {
      if (String(values[r][0]).trim() === ITC_SUGGESTIONS_MARKER) {
        markerRow = r;
        break;
      }
    }
    if (markerRow !== -1) {
      // A previous run's section is here — clear content AND formatting
      // from that row to the end of the sheet before rewriting, so a
      // shorter re-run doesn't leave stale rows behind from a longer one.
      const clearRange = sheet.getRangeByIndexes(markerRow, 0, values.length - markerRow, Math.max(values[0].length, ITC_SUGGESTIONS_HEADERS.length));
      clearRange.clear(Excel.ClearApplyTo.all);
      insertAtRow = markerRow;
    } else {
      insertAtRow = values.length + 1; // one blank row of separation below the existing description-based section
    }
  }

  const flaggedResults = screeningResult.results.filter((r) => r.flagged);
  const skippedCount = screeningResult.results.length - flaggedResults.length;

  const markerRange = sheet.getRangeByIndexes(insertAtRow, 0, 1, 1);
  markerRange.values = [[ITC_SUGGESTIONS_MARKER]];
  markerRange.format.font.bold = true;
  markerRange.format.font.color = "#b8860b"; // distinct amber — visually separates this from the validated rule sections above

  const noteRow = insertAtRow + 1;
  const noteText =
    "AI-generated screening based on vendor NAME ONLY (no invoice or description data) — " +
    screeningResult.results.length + " vendor(s) checked, " + flaggedResults.length + " flagged, " + skippedCount + " no clear match. " +
    "NOT netted into the Summary sheet's ITC total. Review each flagged row against the actual invoice before treating it as ineligible.";
  const noteRange = sheet.getRangeByIndexes(noteRow, 0, 1, 1);
  noteRange.values = [[noteText]];
  noteRange.format.font.italic = true;
  noteRange.format.font.color = "#8a6d1a";

  const headerRow = noteRow + 1;
  const headerRange = sheet.getRangeByIndexes(headerRow, 0, 1, ITC_SUGGESTIONS_HEADERS.length);
  headerRange.values = [ITC_SUGGESTIONS_HEADERS];
  headerRange.format.font.bold = true;
  headerRange.format.fill.color = "#fdf3d8"; // distinct from the description-based section's grey header fill

  if (flaggedResults.length > 0) {
    const dataRows = flaggedResults.map((r) => [r.gstin, r.vendorName, r.category || "", r.subClause || "", r.reasoning]);
    const dataRange = sheet.getRangeByIndexes(headerRow + 1, 0, dataRows.length, ITC_SUGGESTIONS_HEADERS.length);
    dataRange.values = dataRows;
    dataRange.format.wrapText = true;
  }

  sheet.getRangeByIndexes(headerRow, 0, 1 + flaggedResults.length, ITC_SUGGESTIONS_HEADERS.length).format.autofitColumns();
  await context.sync();

  return { flaggedCount: flaggedResults.length, skippedCount, totalChecked: screeningResult.results.length };
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { writeGstReconciliationReport, GST_REPORT_SHEET_NAMES, GST_EXPLAIN_SHEET_NAMES, buildStepTwoTotals, explainGstRow, appendIneligibleItcVendorSuggestions, ITC_SUGGESTIONS_MARKER };
}
