/*
 * description-field.js
 * -----------------------
 * Determines whether a sheet has a genuine per-transaction expense/item
 * description field available for keyword-based content matching — shared
 * by rcm-detector.js and ineligible-itc-detector.js, which both need to
 * match against WHAT was purchased, never WHO it was purchased from.
 *
 * Two rules, in order:
 *
 *   1. Trade/Legal Name is NEVER a candidate. It's explicitly a vendor-
 *      identity field by definition (see gst-column-identifier.js) — a
 *      keyword coincidentally appearing inside a company's registered
 *      name (e.g. a steel-trading company whose name happens to contain
 *      "food") is not evidence about what was actually purchased, and
 *      must never be treated as if it were.
 *
 *   2. Particulars/Narration is only trusted if its CONTENT actually
 *      varies across a vendor's own repeat transactions. Some real
 *      purchase registers (this framework was corrected against one) use
 *      "Particulars" to hold the vendor's name, repeated identically on
 *      every row for that vendor, rather than a transaction-level
 *      description of what was bought. A field that's functionally
 *      constant per-GSTIN is behaving like a vendor attribute, not an
 *      item description — checked structurally, from the data itself,
 *      since the column's header label alone can't be trusted to reflect
 *      how a specific company actually uses it.
 *
 * A B2B raw-material purchase register may legitimately have no field
 * this rule can use — that's a correct, expected answer for some files,
 * not a limitation to engineer around by falling back to a weaker signal.
 */

const MIN_MULTI_ROW_GSTINS_TO_JUDGE = 3; // need this many repeat vendors before drawing a conclusion either way
const CONSTANT_FRACTION_THRESHOLD = 0.85; // this fraction of repeat vendors showing ZERO variation -> looks like a name field

function normalizeText(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim().toLowerCase();
}

// values/headerRowIndex: the sheet's raw data + header position.
// columns: identifyGstColumns() result for this sheet.
// Returns { columnIndex, available, reason } — available=false means
// "don't use this sheet for description-based matching," with reason
// explaining why (no column identified at all, or the column looks like
// a vendor-name field structurally).
function findDescriptionColumn(values, headerRowIndex, columns) {
  if (columns.particulars === null) {
    return { columnIndex: null, available: false, reason: "no Particulars/Narration column identified on this sheet" };
  }
  if (columns.gstin === null) {
    // Can't judge constancy-per-GSTIN without a GSTIN to group by — trust
    // the header identification, since there's no way to check further.
    return { columnIndex: columns.particulars, available: true, reason: "ok (no GSTIN column to cross-check content variability against)" };
  }

  const dataRows = headerRowIndex === -1 ? values : values.slice(headerRowIndex + 1);
  const valuesByGstin = new Map();
  dataRows.forEach((row) => {
    const gstinRaw = row[columns.gstin];
    if (!gstinRaw || String(gstinRaw).trim() === "") return;
    const gstin = String(gstinRaw).trim().toUpperCase();
    const text = normalizeText(row[columns.particulars]);
    if (!text) return;
    if (!valuesByGstin.has(gstin)) valuesByGstin.set(gstin, new Set());
    valuesByGstin.get(gstin).add(text);
  });

  // Only GSTINs with 2+ rows can tell us anything about variability.
  const rowCountByGstin = new Map();
  dataRows.forEach((row) => {
    const gstinRaw = row[columns.gstin];
    if (!gstinRaw || String(gstinRaw).trim() === "") return;
    const gstin = String(gstinRaw).trim().toUpperCase();
    rowCountByGstin.set(gstin, (rowCountByGstin.get(gstin) || 0) + 1);
  });

  let multiRowGstinCount = 0;
  let constantCount = 0;
  for (const [gstin, distinctValues] of valuesByGstin) {
    const rowCount = rowCountByGstin.get(gstin) || 0;
    if (rowCount < 2) continue;
    multiRowGstinCount++;
    if (distinctValues.size === 1) constantCount++;
  }

  if (multiRowGstinCount < MIN_MULTI_ROW_GSTINS_TO_JUDGE) {
    // Not enough repeat-transaction data to draw a conclusion either way —
    // trust the header identification rather than discard a column we
    // have no actual evidence against.
    return { columnIndex: columns.particulars, available: true, reason: "ok (insufficient repeat-transaction data to verify further — trusting the column header)" };
  }

  const constantFraction = constantCount / multiRowGstinCount;
  if (constantFraction >= CONSTANT_FRACTION_THRESHOLD) {
    return {
      columnIndex: null,
      available: false,
      reason: `Particulars column is constant for ${Math.round(constantFraction * 100)}% of vendors with multiple rows — behaves like a vendor name, not a per-transaction description; not used for keyword matching`,
    };
  }

  return { columnIndex: columns.particulars, available: true, reason: "ok (content varies across repeat transactions from the same vendor — behaves like a genuine description field)" };
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { findDescriptionColumn, MIN_MULTI_ROW_GSTINS_TO_JUDGE, CONSTANT_FRACTION_THRESHOLD };
}
