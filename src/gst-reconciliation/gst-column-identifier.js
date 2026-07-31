/*
 * gst-column-identifier.js
 * -------------------------
 * Maps column indices in a GST-related sheet to their GST-legal meaning —
 * GSTIN, Taxable Value, IGST, CGST, SGST, Invoice/Voucher Number,
 * Particulars, Trade/Legal Name, Place of Supply, Invoice Type, Filing
 * Period — using header vocabulary that is standardized by GST law and
 * government e-filing terminology (these are the government's own terms
 * for these fields, not any one company's naming choice), corroborated by
 * the column's actual content shape.
 *
 * GSTIN identification is content-FIRST: the 15-character format is a
 * legal constant (see gstin.js), a stronger signal than any header label
 * a spreadsheet author might have typed. Everything else is header-first,
 * since tax type can't be told apart from raw numbers alone — corroborated
 * by requiring the column to actually look numeric where that's expected.
 */

const { normalizeHeaderText } = require("../sheet-classifier/structural-signature.js");
const { isNumericColumn, isDateColumn } = require("../sheet-classifier/classifier.js");
const { matchesGstinFormat } = require("./gstin.js");

function containsPhrase(normalizedHeader, phrase) {
  return normalizedHeader.includes(phrase);
}

function containsWord(normalizedHeader, word) {
  return normalizedHeader.split(" ").includes(word);
}

// Order matters where terms could overlap (e.g. "sgst" is checked
// independently of "cgst" even though both may combine "state"/"central"
// with "gst"/"tax").
const HEADER_RULES = {
  gstin: (h) => containsWord(h, "gstin") || containsPhrase(h, "gstin uin"),
  // Real Indian bookkeeping exports (Tally and similar tools) routinely
  // abbreviate this column — "Taxable Value" becomes "Taxable Val", "Txbl
  // Value", or just "TV". "Assessable Value" is also standard GST/customs
  // terminology for the same figure (the value tax is calculated on, per
  // CGST Act Section 15) — a different word, same concept, same column
  // role. The 2-letter forms are checked as a whole header token (not a
  // substring) so something like "tv" can't accidentally match inside an
  // unrelated longer word.
  taxableValue: (h) => {
    if (containsPhrase(h, "taxable")) return true; // "Taxable Value", "Taxable Amount", "Taxable Val"...
    if (containsPhrase(h, "assessable")) return true; // "Assessable Value" / "Assessable Amount"
    if (containsPhrase(h, "txbl")) return true; // "Txbl Value", "Txbl Val", "Txbl Amt"
    const words = h.split(" ");
    return words.includes("tv"); // bare initials for Taxable Value
  },
  igst: (h) => containsWord(h, "igst") || containsPhrase(h, "integrated tax") || containsPhrase(h, "integrated gst"),
  cgst: (h) => containsWord(h, "cgst") || containsPhrase(h, "central tax") || containsPhrase(h, "central gst"),
  // "State/UT Tax" — the actual official GST portal label, since GST law
  // applies the same tax uniformly to States and Union Territories — comes
  // out as "state ut tax" after normalization (the "/" becomes a space),
  // which breaks a contiguous "state tax" phrase match. Checked word-by-word
  // instead of as one fixed phrase for exactly that reason.
  sgst: (h) => {
    const words = h.split(" ");
    if (words.includes("sgst") || words.includes("ugst")) return true;
    const mentionsStateOrUt = words.includes("state") || words.includes("ut") || h.includes("union territory");
    const mentionsTaxOrGst = words.includes("tax") || words.includes("gst");
    return mentionsStateOrUt && mentionsTaxOrGst;
  },
  invoiceNumber: (h) => containsPhrase(h, "invoice") && (containsWord(h, "no") || containsWord(h, "number") || containsWord(h, "num")),
  voucherNumber: (h) => containsPhrase(h, "voucher"),
  particulars: (h) => containsWord(h, "particulars") || containsWord(h, "narration"),
  tradeLegalName: (h) => containsPhrase(h, "trade") || containsPhrase(h, "legal name") || containsPhrase(h, "supplier name") || containsPhrase(h, "vendor name"),
  placeOfSupply: (h) => containsPhrase(h, "place of supply"),
  invoiceType: (h) => containsPhrase(h, "invoice type"),
  filingPeriod: (h) => containsPhrase(h, "filing period") || containsPhrase(h, "return period") || containsPhrase(h, "gstr 1") || containsPhrase(h, "gstr 2a"),
};

// Most non-blank cells in the column need to look like GSTINs, not all —
// tolerates a handful of blanks/typos without losing the column entirely.
const GSTIN_CONTENT_THRESHOLD = 0.6;

function computeGstinRatios(values, headerRowIndex) {
  const dataRows = headerRowIndex === -1 ? values : values.slice(headerRowIndex + 1);
  const columnCount = values.reduce((max, row) => Math.max(max, row.length), 0);
  const ratios = new Array(columnCount).fill(0);

  for (let c = 0; c < columnCount; c++) {
    let nonBlank = 0;
    let matches = 0;
    for (const row of dataRows) {
      const cell = row[c];
      if (cell === "" || cell === null || cell === undefined) continue;
      nonBlank++;
      if (matchesGstinFormat(cell)) matches++;
    }
    ratios[c] = nonBlank === 0 ? 0 : matches / nonBlank;
  }
  return ratios;
}

function identifyGstColumns(sheetSignals, values, headerRowIndex) {
  const columns = sheetSignals.columns;
  const gstinRatios = computeGstinRatios(values, headerRowIndex);

  const result = {
    gstin: null,
    taxableValue: null,
    igst: null,
    cgst: null,
    sgst: null,
    invoiceNumber: null,
    voucherNumber: null,
    particulars: null,
    tradeLegalName: null,
    placeOfSupply: null,
    invoiceType: null,
    filingPeriod: null,
    dateColumns: [],
  };

  let bestGstinIdx = -1;
  let bestGstinRatio = 0;
  for (let c = 0; c < gstinRatios.length; c++) {
    if (gstinRatios[c] >= GSTIN_CONTENT_THRESHOLD && gstinRatios[c] > bestGstinRatio) {
      bestGstinRatio = gstinRatios[c];
      bestGstinIdx = c;
    }
  }
  if (bestGstinIdx !== -1) result.gstin = bestGstinIdx;

  for (let c = 0; c < columns.length; c++) {
    const col = columns[c];
    const header = normalizeHeaderText(col.header);
    if (!header) continue;

    if (result.taxableValue === null && HEADER_RULES.taxableValue(header) && isNumericColumn(col)) result.taxableValue = c;
    if (result.igst === null && HEADER_RULES.igst(header) && isNumericColumn(col)) result.igst = c;
    if (result.cgst === null && HEADER_RULES.cgst(header) && isNumericColumn(col)) result.cgst = c;
    if (result.sgst === null && HEADER_RULES.sgst(header) && isNumericColumn(col)) result.sgst = c;
    if (result.invoiceNumber === null && HEADER_RULES.invoiceNumber(header)) result.invoiceNumber = c;
    if (result.voucherNumber === null && HEADER_RULES.voucherNumber(header)) result.voucherNumber = c;
    if (result.particulars === null && HEADER_RULES.particulars(header)) result.particulars = c;
    if (result.tradeLegalName === null && HEADER_RULES.tradeLegalName(header)) result.tradeLegalName = c;
    if (result.placeOfSupply === null && HEADER_RULES.placeOfSupply(header)) result.placeOfSupply = c;
    if (result.invoiceType === null && HEADER_RULES.invoiceType(header)) result.invoiceType = c;
    if (result.filingPeriod === null && HEADER_RULES.filingPeriod(header)) result.filingPeriod = c;

    if (isDateColumn(col)) result.dateColumns.push(c);
  }

  return result;
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { identifyGstColumns, computeGstinRatios, HEADER_RULES, GSTIN_CONTENT_THRESHOLD };
}
