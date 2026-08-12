/*
 * rate-mismatch-detector.js
 * ---------------------------
 * Pattern-agnostic sanity-check backstop: taxable value × stated rate%
 * should approximately equal the tax amount actually picked up, for every
 * row. Independent of and complementary to the column-ambiguity resolution
 * in column-memory.js/gst-column-ambiguity-rules.js — this catches a wrong
 * column pick EVEN WHEN the header matched confidently (a single, non-
 * ambiguous "CGST" column that turns out to hold the wrong figures for
 * some structural reason ambiguity detection wouldn't predict), which is
 * exactly the kind of future bug in this same family that naming
 * conventions alone can't catch.
 *
 * Informational only — like Wrong Head/RCM, a flagged row here never
 * blocks the reconciliation proposal and never changes the net eligible
 * ITC figure. Only unresolved COLUMN IDENTITY (gst-column-ambiguity-
 * rules.js) blocks; this is a per-row data-quality flag for review.
 */

const TOLERANCE_FLOOR = 1; // ₹1 — matches this codebase's usual rounding-noise convention (see wrong-head-detector.js)
const TOLERANCE_PERCENT = 0.01; // 1% of the expected tax — reuses the existing materiality convention (gst-report-writer.js) rather than inventing a third constant; a flat ₹1 is too tight once legitimate per-line CGST/SGST rounding accumulates at scale.

function toNumber(value) {
  if (typeof value === "number") return value;
  if (value === "" || value === null || value === undefined) return 0;
  const parsed = parseFloat(String(value).replace(/[,$£€¥%()]/g, ""));
  return isNaN(parsed) ? 0 : parsed;
}

function roundTo2(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

// Real GST rates never fall strictly between 0 and 1 as a plain percent
// figure (the lowest statutory rate is 0.25%, rough/semi-precious stones
// only) EXCEPT when Excel stores a %-formatted cell as its underlying
// fraction (0.18 for "18%"). Treating anything in (0,1) as "already a
// fraction, multiply by 100" is a documented, accepted trade-off: it will
// misread a genuine 0.25% rate as 25%. That's a narrow real-world case,
// not fixed here via numberFormat plumbing through every detector call
// site for one edge case.
function normalizeRatePercent(raw) {
  const n = toNumber(raw);
  return n > 0 && n < 1 ? n * 100 : n;
}

// values/headerRowIndex: the sheet's raw data + header position. columns:
// identifyGstColumns() result (post-resolution) — needs rate, taxableValue,
// and at least one of igst/cgst/sgst resolved; otherwise there's nothing
// to check and the sheet is silently skipped (not an error — rate is
// optional, unlike the fields column-ambiguity-rules.js treats as
// required).
function detectRateMismatch(values, headerRowIndex, columns) {
  if (columns.rate === null || columns.taxableValue === null || (columns.igst === null && columns.cgst === null && columns.sgst === null)) {
    return { applicable: false, reason: "sheet has no resolved Rate column, Taxable Value column, or any tax column to check against", flagged: [] };
  }

  const dataRows = headerRowIndex === -1 ? values : values.slice(headerRowIndex + 1);
  const flagged = [];

  dataRows.forEach((row, i) => {
    const taxableValue = toNumber(row[columns.taxableValue]);
    const ratePercent = normalizeRatePercent(row[columns.rate]);
    const igst = columns.igst !== null ? toNumber(row[columns.igst]) : 0;
    const cgst = columns.cgst !== null ? toNumber(row[columns.cgst]) : 0;
    const sgst = columns.sgst !== null ? toNumber(row[columns.sgst]) : 0;
    const actualTotalTax = igst + cgst + sgst;

    if (taxableValue === 0 && ratePercent === 0 && actualTotalTax === 0) return; // blank/nil row, nothing to check

    const expectedTotalTax = roundTo2((taxableValue * ratePercent) / 100);
    if (expectedTotalTax === 0 && actualTotalTax === 0) return; // nil-rated/exempt

    const tolerance = Math.max(TOLERANCE_FLOOR, expectedTotalTax * TOLERANCE_PERCENT);
    const difference = roundTo2(actualTotalTax - expectedTotalTax);

    if (Math.abs(difference) > tolerance) {
      flagged.push({
        rowIndex: i,
        gstin: columns.gstin !== null ? row[columns.gstin] : null,
        invoiceNumber: columns.invoiceNumber !== null ? row[columns.invoiceNumber] : null,
        taxableValue,
        ratePercent,
        expectedTotalTax,
        actualTotalTax,
        difference,
      });
    }
  });

  return { applicable: true, flagged };
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { detectRateMismatch, normalizeRatePercent };
}
