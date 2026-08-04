/*
 * ineligible-itc-vendor-extraction.js
 * -----------------------------------------
 * Pulls the distinct (GSTIN, vendor name) pairs off the 2A and Books
 * sheets for the vendor-name LLM suggestion feature — used only when
 * ineligible-itc-detector.js found no genuine per-transaction description
 * field to check (see mentor-gst-reconciliation-ui.js for the trigger
 * condition). This file does NOT touch src/gst-reconciliation/*.js and
 * does not change how that detector behaves — it's a separate extraction
 * for a separate, optional, user-triggered feature.
 *
 * Column choice per sheet:
 *   - 2A: Trade/Legal Name (identifyGstColumns' own vendor-identity field).
 *   - Books: no Trade/Legal Name column exists structurally on a purchase
 *     register — but description-field.js has ALREADY determined, for the
 *     Ineligible ITC rule's own purposes, whether the Particulars column
 *     is actually behaving as a vendor-name field (constant per GSTIN)
 *     rather than a genuine description. Reusing that verdict here is
 *     correct, not a hack: we only ever read Particulars as a name source
 *     when it's already been found to structurally BE one.
 */

const { findDescriptionColumn } = require("../gst-reconciliation/description-field.js");

function resolveVendorNameColumn(values, headerRowIndex, columns) {
  if (columns.tradeLegalName !== null) return columns.tradeLegalName;

  if (columns.particulars !== null) {
    const descVerdict = findDescriptionColumn(values, headerRowIndex, columns);
    if (!descVerdict.available && /vendor name/.test(descVerdict.reason)) {
      return columns.particulars;
    }
  }
  return null;
}

// values/headerRowIndex/columns: same shape as everywhere else in the GST
// pipeline (identifyGstColumns() result). Returns [{gstin, vendorName}],
// one entry per distinct GSTIN on this sheet (first name seen wins if the
// same GSTIN's name spelling varies across rows).
function extractVendorNamesFromSheet(values, headerRowIndex, columns) {
  const nameColumnIndex = resolveVendorNameColumn(values, headerRowIndex, columns);
  if (nameColumnIndex === null || columns.gstin === null) return [];

  const dataRows = headerRowIndex === -1 ? values : values.slice(headerRowIndex + 1);
  const byGstin = new Map();
  for (const row of dataRows) {
    const gstinRaw = row[columns.gstin];
    if (!gstinRaw || String(gstinRaw).trim() === "") continue;
    const gstin = String(gstinRaw).trim().toUpperCase();
    const name = String(row[nameColumnIndex] || "").trim();
    if (!name) continue;
    if (!byGstin.has(gstin)) byGstin.set(gstin, name);
  }
  return [...byGstin.entries()].map(([gstin, vendorName]) => ({ gstin, vendorName }));
}

// gstr2aInput/booksInput: { values, headerRowIndex, columns } for each
// sheet. Merges and deduplicates by GSTIN — 2A's Trade/Legal Name (the
// government-recognized name) wins over Books' name when the same GSTIN
// appears on both sides.
function extractDistinctVendors(gstr2aInput, booksInput) {
  const fromBooks = extractVendorNamesFromSheet(booksInput.values, booksInput.headerRowIndex, booksInput.columns);
  const fromGstr2a = extractVendorNamesFromSheet(gstr2aInput.values, gstr2aInput.headerRowIndex, gstr2aInput.columns);

  const merged = new Map();
  for (const v of fromBooks) merged.set(v.gstin, v.vendorName);
  for (const v of fromGstr2a) merged.set(v.gstin, v.vendorName);
  return [...merged.entries()].map(([gstin, vendorName]) => ({ gstin, vendorName }));
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { extractVendorNamesFromSheet, extractDistinctVendors, resolveVendorNameColumn };
}
