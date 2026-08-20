/*
 * gstr2a-vs-books-reconciler.js
 * -------------------------------
 * Workflow-specific reconciliation logic for GSTR-2A vs Books.
 *
 * Aggregates each recognized sheet by GSTIN — sum of Taxable Value, IGST,
 * CGST, SGST per supplier, the same fast first pass a human does before
 * checking individual invoices — then compares the two aggregates per
 * GSTIN within a small rounding tolerance, and builds a PROPOSED summary.
 * Nothing here writes to the workbook or applies anything automatically;
 * MENTOR's existing "propose, don't auto-execute" rule (see
 * docs/ARCHITECTURE-PRINCIPLES.md) applies here too.
 *
 * This file is intentionally workflow-specific (not a generic "compare any
 * two GST sheets" function) — a future GSTR-1 vs GSTR-3B workflow compares
 * at the return-total level, not per-GSTIN, so it gets its own sibling
 * reconciler file rather than forcing this one to generalize.
 *
 * Match criterion is IGST/CGST/SGST only — NOT Taxable Value. This was
 * validated against a real expert-completed reconciliation: the expert's
 * own per-GSTIN comparison likewise never compares taxable value, only the
 * three tax fields. That's not an oversight — Input Tax Credit eligibility
 * (CGST Act Section 16) turns on the tax amount actually charged/claimed,
 * not the taxable value it was calculated on. A taxable-value gap with
 * matching tax can arise from rate splits or partial-invoice booking
 * without affecting the credit at stake, so it isn't treated as a
 * reconciliation exception on its own — though it's still computed and
 * reported (see `differences.taxableValue`) for context.
 *
 * Known limitation, not something this pipeline tries to force to zero:
 * even a careful manual reconciliation typically retains a small residual
 * of genuinely unexplained differences and presence gaps (timing, minor
 * uninvestigated items) — validated against the same real answer key,
 * whose own top-level summary shows an acknowledged unreconciled residual
 * after all of its categorization. A handful of remaining discrepancies/
 * presence-gaps after applying tax-field tolerance is expected, not
 * necessarily evidence of a bug in this pipeline.
 */

const DEFAULT_TOLERANCE = 1; // ₹1 — rounding/paise-level noise, not a real discrepancy

function toNumber(value) {
  if (typeof value === "number") return value;
  if (value === "" || value === null || value === undefined) return 0;
  const parsed = parseFloat(String(value).replace(/[,$£€¥%()]/g, ""));
  return isNaN(parsed) ? 0 : parsed;
}

function roundTo2(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

// values: raw 2D sheet array. headerRowIndex: from signal-extractor.
// columns: the identifyGstColumns() result for this sheet.
function aggregateByGstin(values, headerRowIndex, columns) {
  const dataRows = headerRowIndex === -1 ? values : values.slice(headerRowIndex + 1);
  const totals = new Map(); // normalized GSTIN -> { taxableValue, igst, cgst, sgst, rowCount }

  for (const row of dataRows) {
    const gstinRaw = columns.gstin !== null ? row[columns.gstin] : null;
    if (gstinRaw === null || gstinRaw === undefined || String(gstinRaw).trim() === "") continue;
    const gstin = String(gstinRaw).trim().toUpperCase();

    if (!totals.has(gstin)) {
      totals.set(gstin, { taxableValue: 0, igst: 0, cgst: 0, sgst: 0, rowCount: 0 });
    }
    const bucket = totals.get(gstin);
    bucket.taxableValue += columns.taxableValue !== null ? toNumber(row[columns.taxableValue]) : 0;
    bucket.igst += columns.igst !== null ? toNumber(row[columns.igst]) : 0;
    bucket.cgst += columns.cgst !== null ? toNumber(row[columns.cgst]) : 0;
    bucket.sgst += columns.sgst !== null ? toNumber(row[columns.sgst]) : 0;
    bucket.rowCount += 1;
  }

  return totals;
}

function emptyTotal() {
  return { taxableValue: 0, igst: 0, cgst: 0, sgst: 0, rowCount: 0 };
}

function compareAggregates(gstr2aTotals, booksTotals, options) {
  const tolerance = (options && options.tolerance) !== undefined ? options.tolerance : DEFAULT_TOLERANCE;
  const allGstins = new Set([...gstr2aTotals.keys(), ...booksTotals.keys()]);
  const results = [];

  for (const gstin of allGstins) {
    const inGstr2a = gstr2aTotals.has(gstin);
    const inBooks = booksTotals.has(gstin);
    const gstr2aTotal = inGstr2a ? gstr2aTotals.get(gstin) : emptyTotal();
    const booksTotal = inBooks ? booksTotals.get(gstin) : emptyTotal();

    const differences = {
      taxableValue: roundTo2(gstr2aTotal.taxableValue - booksTotal.taxableValue),
      igst: roundTo2(gstr2aTotal.igst - booksTotal.igst),
      cgst: roundTo2(gstr2aTotal.cgst - booksTotal.cgst),
      sgst: roundTo2(gstr2aTotal.sgst - booksTotal.sgst),
    };

    let status;
    if (!inGstr2a) status = "missing_in_gstr2a";
    else if (!inBooks) status = "missing_in_books";
    else {
      // Tax fields only — see the module docstring for why Taxable Value
      // is deliberately excluded from the match determination.
      const taxWithinTolerance = Math.abs(differences.igst) <= tolerance && Math.abs(differences.cgst) <= tolerance && Math.abs(differences.sgst) <= tolerance;
      status = taxWithinTolerance ? "matched" : "discrepancy";
    }

    results.push({
      gstin,
      status,
      gstr2a: { taxableValue: roundTo2(gstr2aTotal.taxableValue), igst: roundTo2(gstr2aTotal.igst), cgst: roundTo2(gstr2aTotal.cgst), sgst: roundTo2(gstr2aTotal.sgst), rowCount: gstr2aTotal.rowCount },
      books: { taxableValue: roundTo2(booksTotal.taxableValue), igst: roundTo2(booksTotal.igst), cgst: roundTo2(booksTotal.cgst), sgst: roundTo2(booksTotal.sgst), rowCount: booksTotal.rowCount },
      differences,
    });
  }

  // Discrepancies first (most actionable), then missing-in-*, matched last.
  const order = { discrepancy: 0, missing_in_books: 1, missing_in_gstr2a: 2, matched: 3 };
  results.sort((a, b) => order[a.status] - order[b.status] || a.gstin.localeCompare(b.gstin));

  return results;
}

function buildReconciliationSummary(comparisonResults, options) {
  const tolerance = (options && options.tolerance) !== undefined ? options.tolerance : DEFAULT_TOLERANCE;
  const counts = { matched: 0, discrepancy: 0, missing_in_books: 0, missing_in_gstr2a: 0 };
  for (const r of comparisonResults) counts[r.status]++;

  return {
    workflow: "gstr2a_vs_books",
    proposedAction: "review", // MENTOR proposes this for review; nothing here writes to the workbook or is auto-applied
    tolerance,
    totalGstins: comparisonResults.length,
    counts,
    vendors: comparisonResults,
  };
}

function reconcileGstr2aVsBooks(gstr2aInput, booksInput, options) {
  const gstr2aTotals = aggregateByGstin(gstr2aInput.values, gstr2aInput.headerRowIndex, gstr2aInput.columns);
  const booksTotals = aggregateByGstin(booksInput.values, booksInput.headerRowIndex, booksInput.columns);
  const comparisonResults = compareAggregates(gstr2aTotals, booksTotals, options);
  return buildReconciliationSummary(comparisonResults, options);
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { aggregateByGstin, compareAggregates, buildReconciliationSummary, reconcileGstr2aVsBooks, DEFAULT_TOLERANCE };
}
