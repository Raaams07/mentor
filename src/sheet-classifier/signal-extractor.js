/*
 * signal-extractor.js
 * -------------------
 * MENTOR Phase 2, Step 1 — generalized signal extraction.
 *
 * Today MENTOR only understands a sheet if its name and column order
 * exactly match the Three Sisters Kitchen demo (see parseInvoiceRows,
 * parseBankRows etc. in "Mentor workbook index/mentor-workbook-index.js").
 * That breaks the moment a real client's workbook uses different sheet
 * names or column layouts.
 *
 * This module doesn't know what a "bank statement" or "invoice list" is.
 * It only answers one question, per column, for ANY sheet: what does the
 * data in this column actually look like? Classification (Step 2) is a
 * separate layer built on top of these numbers — kept separate so the
 * signals stay honest and reusable even if the classification rules change.
 *
 * Pure functions, no Office.js dependency — so this can be unit-tested
 * directly in Node against a real .xlsx file, not only inside Excel.
 * A thin Excel-aware wrapper (extractSheetSignalsFromRange) is included
 * at the bottom for when this gets wired into the live add-in.
 */

// ---------------------------------------------------------------
// 1. HEADER ROW DETECTION
//    Generic replacement for the old keyword-list approach
//    (mentorFindHeaderRowIdx looked for "invoice"/"amount"/"category").
//    A header row is identified by SHAPE, not by knowing the vocabulary
//    of any particular client's column names: mostly text, at least two
//    filled cells, and real data actually follows it.
// ---------------------------------------------------------------
function isBlankCell(value) {
  return value === "" || value === null || value === undefined;
}

function findHeaderRowIndex(values) {
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const nonBlank = row.filter((c) => !isBlankCell(c));
    if (nonBlank.length < 2) continue; // a lone title cell isn't a header row

    const textCells = nonBlank.filter((c) => typeof c === "string" && isNaN(Number(c)));
    const looksTextual = textCells.length / nonBlank.length >= 0.6;
    if (!looksTextual) continue;

    const hasDataBelow = values.slice(i + 1, i + 6).some((r) => r.some((c) => !isBlankCell(c)));
    if (!hasDataBelow) continue;

    return i;
  }
  return -1;
}

// ---------------------------------------------------------------
// 2. PER-CELL TYPE DETECTION
// ---------------------------------------------------------------
const DATE_STRING_PATTERNS = [
  /^\d{4}-\d{1,2}-\d{1,2}$/, // 2026-06-05 (ISO)
  /^\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}$/, // 5/6/2026, 05.06.2026, 5-6-26
  /^[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4}$/, // June 5, 2026 / June 5 2026
  /^\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}$/, // 5 June 2026
  /^[A-Za-z]{3,9}\s+\d{4}$/, // Jun 2026 (month-label columns, e.g. POS "Month")
];

const NUMERIC_STRING_PATTERN = /^\(?[-+]?[$£€¥]?\s?\d{1,3}(,\d{3})*(\.\d+)?%?\)?$/;

// Excel date format codes always use d/m/y letter tokens and never use
// '0' or '#' digit placeholders — that combination is what separates a
// date format from a plain number/currency format string.
function formatLooksLikeDate(numberFormat) {
  if (!numberFormat) return false;
  const stripped = numberFormat.replace(/\[[^\]]*\]/g, "");
  if (stripped.toLowerCase() === "general" || stripped === "@") return false;
  return /[dmy]/i.test(stripped) && !/[#0]/.test(stripped);
}

function formatLooksLikeNumeric(numberFormat) {
  if (!numberFormat) return false;
  const stripped = numberFormat.replace(/\[[^\]]*\]/g, "");
  return /[#0]/.test(stripped);
}

// Serial-date range roughly spanning 1990-01-01 to 2035-12-31. Only used
// as a FALLBACK when no number-format metadata is available (e.g. a plain
// values-only read) — Excel serials and plain currency amounts overlap in
// range, so this is a weak signal, not a confident one. Kept separate in
// the output (see `dateSerialFallbackCount`) so Step 2 can weigh it accordingly.
const PLAUSIBLE_DATE_SERIAL_MIN = 32874;
const PLAUSIBLE_DATE_SERIAL_MAX = 49674;

function classifyCell(value, numberFormat) {
  if (isBlankCell(value)) return { type: "blank" };
  if (value instanceof Date) return { type: "date" };

  if (numberFormat) {
    if (formatLooksLikeDate(numberFormat)) return { type: "date" };
    if (formatLooksLikeNumeric(numberFormat)) return { type: "numeric" };
  }

  if (typeof value === "number") {
    const isPlausibleDateSerial = value >= PLAUSIBLE_DATE_SERIAL_MIN && value <= PLAUSIBLE_DATE_SERIAL_MAX;
    if (!numberFormat && isPlausibleDateSerial) {
      return { type: "date", fromDateSerialFallback: true };
    }
    return { type: "numeric" };
  }

  if (typeof value === "string") {
    const trimmed = value.trim();
    if (trimmed === "") return { type: "blank" };
    if (DATE_STRING_PATTERNS.some((re) => re.test(trimmed))) return { type: "date" };
    if (NUMERIC_STRING_PATTERN.test(trimmed)) return { type: "numeric" };
    return { type: "text" };
  }

  return { type: "text" }; // booleans and anything else — rare, treat as categorical text
}

// ---------------------------------------------------------------
// 3. PER-COLUMN SIGNAL AGGREGATION
// ---------------------------------------------------------------
function normalizeForDistinctCount(value) {
  return typeof value === "string" ? value.trim() : String(value);
}

function computeMonotonicSignal(numericValuesInOrder) {
  if (numericValuesInOrder.length < 2) {
    return {
      numericSampleSize: numericValuesInOrder.length,
      increasingRatio: null,
      decreasingRatio: null,
      isMonotonicIncreasing: false,
      isMonotonicDecreasing: false,
    };
  }

  let increasingPairs = 0;
  let decreasingPairs = 0;
  const totalPairs = numericValuesInOrder.length - 1;

  for (let i = 1; i < numericValuesInOrder.length; i++) {
    if (numericValuesInOrder[i] >= numericValuesInOrder[i - 1]) increasingPairs++;
    if (numericValuesInOrder[i] <= numericValuesInOrder[i - 1]) decreasingPairs++;
  }

  const increasingRatio = increasingPairs / totalPairs;
  const decreasingRatio = decreasingPairs / totalPairs;

  return {
    numericSampleSize: numericValuesInOrder.length,
    increasingRatio,
    decreasingRatio,
    // Running-balance signature: every step moves the same direction (or holds).
    // Ratio-based rather than a strict "every single pair" check, so one bad
    // data-entry cell doesn't erase an otherwise-clear signal.
    isMonotonicIncreasing: increasingRatio === 1,
    isMonotonicDecreasing: decreasingRatio === 1,
  };
}

// Sign distribution of a numeric column — e.g. a bank statement's Amount
// column mixes positive and negative values (debits and credits in one
// column), while an invoice's Net Amount column is almost always all-positive.
function computeSignStats(numericValuesInOrder) {
  const n = numericValuesInOrder.length;
  if (n === 0) {
    return { positiveCount: 0, negativeCount: 0, zeroCount: 0, pctPositive: 0, pctNegative: 0, pctZero: 0, hasBothSigns: false };
  }

  let positiveCount = 0;
  let negativeCount = 0;
  let zeroCount = 0;
  for (const value of numericValuesInOrder) {
    if (value > 0) positiveCount++;
    else if (value < 0) negativeCount++;
    else zeroCount++;
  }

  return {
    positiveCount,
    negativeCount,
    zeroCount,
    pctPositive: positiveCount / n,
    pctNegative: negativeCount / n,
    pctZero: zeroCount / n,
    hasBothSigns: positiveCount > 0 && negativeCount > 0,
  };
}

function extractColumnSignals(colIndex, header, dataRows, numberFormatRows) {
  const typeCounts = { blank: 0, date: 0, numeric: 0, text: 0 };
  const valueCounts = new Map();
  const numericValuesInOrder = [];
  const sampleValues = [];
  let dateSerialFallbackCount = 0;

  for (let r = 0; r < dataRows.length; r++) {
    const value = dataRows[r][colIndex];
    const numberFormat = numberFormatRows ? numberFormatRows[r][colIndex] : undefined;
    const { type, fromDateSerialFallback } = classifyCell(value, numberFormat);

    typeCounts[type]++;
    if (fromDateSerialFallback) dateSerialFallbackCount++;

    if (type === "blank") continue;

    if (type === "numeric") {
      const numericValue = typeof value === "number" ? value : parseFloat(String(value).replace(/[,$£€¥%()]/g, ""));
      if (!isNaN(numericValue)) numericValuesInOrder.push(numericValue);
    }

    const key = normalizeForDistinctCount(value);
    valueCounts.set(key, (valueCounts.get(key) || 0) + 1);

    if (sampleValues.length < 5) sampleValues.push(value);
  }

  const nonBlankCount = dataRows.length - typeCounts.blank;
  const pct = (count) => (nonBlankCount === 0 ? 0 : count / nonBlankCount);

  const topValues = [...valueCounts.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5)
    .map(([value, count]) => ({ value, count }));

  return {
    index: colIndex,
    header: header === undefined ? null : header,
    nonBlankCount,
    blankCount: typeCounts.blank,
    pctBlank: dataRows.length === 0 ? 0 : typeCounts.blank / dataRows.length,
    pctDate: pct(typeCounts.date),
    pctNumeric: pct(typeCounts.numeric),
    pctText: pct(typeCounts.text),
    dateSerialFallbackCount,
    distinctCount: valueCounts.size,
    distinctRatio: nonBlankCount === 0 ? 0 : valueCounts.size / nonBlankCount,
    topValues,
    monotonic: computeMonotonicSignal(numericValuesInOrder),
    signStats: computeSignStats(numericValuesInOrder),
    sampleValues,
  };
}

// ---------------------------------------------------------------
// 4. TOP-LEVEL: ONE SHEET, MANY SHEETS
// ---------------------------------------------------------------
function extractSheetSignals(sheetName, values, options) {
  const opts = options || {};
  const numberFormats = opts.numberFormats || null;

  if (!values || values.length === 0) {
    return { sheetName, headerRowIndex: -1, headerLabels: [], dataRowCount: 0, columnCount: 0, columns: [] };
  }

  const headerRowIndex = opts.headerRowIndex !== undefined ? opts.headerRowIndex : findHeaderRowIndex(values);
  const dataStartRow = headerRowIndex === -1 ? 0 : headerRowIndex + 1;
  const dataRows = values.slice(dataStartRow);
  const numberFormatRows = numberFormats ? numberFormats.slice(dataStartRow) : null;

  const columnCount = values.reduce((max, row) => Math.max(max, row.length), 0);
  const headerLabels = headerRowIndex === -1 ? [] : values[headerRowIndex];

  const columns = [];
  for (let c = 0; c < columnCount; c++) {
    const header = headerLabels[c];
    const signals = extractColumnSignals(c, header, dataRows, numberFormatRows);
    columns.push(signals);
  }

  return {
    sheetName: sheetName === undefined ? null : sheetName,
    headerRowIndex,
    headerLabels,
    dataRowCount: dataRows.length,
    columnCount,
    columns,
  };
}

function extractWorkbookSignals(sheets) {
  // sheets: [{ name, values, numberFormats? }, ...]
  return sheets.map((sheet) => extractSheetSignals(sheet.name, sheet.values, { numberFormats: sheet.numberFormats }));
}

// ---------------------------------------------------------------
// 5. THIN OFFICE.JS BRIDGE — for when this gets wired into the live add-in.
//    Not used yet; kept here so Step 1's output shape is provable against
//    a real worksheet, not just arrays constructed by hand.
// ---------------------------------------------------------------
async function extractSheetSignalsFromWorksheet(context, sheet) {
  const range = sheet.getUsedRange();
  range.load("values, numberFormat");
  await context.sync();
  sheet.load("name");
  await context.sync();
  return extractSheetSignals(sheet.name, range.values, { numberFormats: range.numberFormat });
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    findHeaderRowIndex,
    classifyCell,
    extractSheetSignals,
    extractWorkbookSignals,
    extractSheetSignalsFromWorksheet,
  };
}
