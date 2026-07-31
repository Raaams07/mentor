/*
 * rcm-detector.js
 * -----------------
 * Reverse Charge Mechanism (RCM) detection, per Section 9(3)/9(4) CGST Act
 * and Section 5(3)/5(4) IGST Act: for notified categories, the RECIPIENT
 * (not the supplier) is liable to pay GST.
 *
 * Two independent signals, in priority order:
 *   1. GSTR-2A's own "Supply Attract Reverse Charge" (Y/N) flag, where the
 *      column is present and has a recognizable value — trusted as the
 *      primary signal, since it's the government's own determination.
 *   2. A fallback keyword match against a genuine per-transaction
 *      description field (see description-field.js), against
 *      rcm-config.js's editable list of notified categories.
 *
 * The fallback ONLY ever matches a field verified to actually describe
 * what was purchased — never a vendor/company name field, even when a
 * keyword coincidentally appears inside one. Losing the fallback (no
 * genuine description field on a sheet) does NOT disable the flag-based
 * path — they're independent signals, so a sheet with the flag column but
 * no usable description field is still fully checkable; a sheet with
 * NEITHER (e.g. a Books sheet, which never has the government's flag, and
 * whose Particulars column turns out to just repeat the vendor's name)
 * genuinely has nothing this rule can determine — reported explicitly via
 * keywordFallbackAvailable, not hidden behind a silent zero.
 */

const { RCM_CATEGORIES } = require("./rcm-config.js");
const { findDescriptionColumn } = require("./description-field.js");

const YES_VALUES = new Set(["y", "yes", "true", "1"]);
const NO_VALUES = new Set(["n", "no", "false", "0"]);

function normalizeFlag(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim().toLowerCase();
}

function readReverseChargeFlag(row, columns) {
  if (columns.reverseCharge === null) return null; // column not present on this sheet
  const raw = normalizeFlag(row[columns.reverseCharge]);
  if (YES_VALUES.has(raw)) return true;
  if (NO_VALUES.has(raw)) return false;
  return null; // present but blank/unrecognized — fall through to keyword matching
}

// Some categories (currently just metal scrap) are conditional on the
// SUPPLIER being unregistered, not just on the description keyword — see
// rcm-config.js. A populated, present GSTIN on the row means the supplier
// IS registered, so those categories must not confirm a match even if the
// keyword hits (ordinary forward-charge GST + Section 51 TDS applies
// instead there, not RCM).
function supplierIsUnregistered(row, columns) {
  if (columns.gstin === null) return true; // no GSTIN column at all on this sheet — can't contradict "unregistered"
  const gstinRaw = row[columns.gstin];
  return !gstinRaw || String(gstinRaw).trim() === "";
}

function matchRcmKeyword(descriptionText, row, columns) {
  if (!descriptionText) return null;
  const text = String(descriptionText).toLowerCase();
  for (const category of RCM_CATEGORIES) {
    if (!category.keywords.some((kw) => text.includes(kw))) continue;
    if (category.requiresUnregisteredSupplier && !supplierIsUnregistered(row, columns)) continue; // registered supplier — this category doesn't apply, keep checking others
    return category;
  }
  return null;
}

// row: raw row array. columns: identifyGstColumns() result for this sheet.
// descriptionColumnIndex: the column verified genuine by
// findDescriptionColumn() — null if no such field is available, in which
// case only the flag-based signal is checked.
function detectRcmForRow(row, columns, descriptionColumnIndex) {
  const flag = readReverseChargeFlag(row, columns);
  if (flag !== null) {
    return { isRcm: flag, source: "reverse_charge_flag", matchedCategory: null };
  }

  if (descriptionColumnIndex !== null) {
    const matchedCategory = matchRcmKeyword(row[descriptionColumnIndex], row, columns);
    if (matchedCategory) {
      return { isRcm: true, source: "keyword_match", matchedCategory };
    }
  }

  return { isRcm: false, source: "no_signal", matchedCategory: null };
}

// values/headerRowIndex: the sheet's raw data + header position.
// columns: identifyGstColumns() result for this sheet.
// Only returns rows determined to actually be RCM — callers wanting the
// full per-row breakdown (including non-RCM rows) should call
// detectRcmForRow directly per row instead.
//
// Deliberately does NOT require a populated GSTIN to consider a row — an
// unregistered-supplier purchase (the metal-scrap case, for one) is
// EXPECTED to have a blank GSTIN, so gating on GSTIN presence here would
// make that category unreachable through this scan. A row is skipped only
// if it has neither a GSTIN nor any description text to work with at all.
function detectRcmForSheet(values, headerRowIndex, columns) {
  const descField = findDescriptionColumn(values, headerRowIndex, columns);
  const keywordFallbackAvailable = descField.available;

  const dataRows = headerRowIndex === -1 ? values : values.slice(headerRowIndex + 1);
  const flagged = [];

  dataRows.forEach((row, i) => {
    const gstinRaw = columns.gstin !== null ? row[columns.gstin] : null;
    const hasGstin = gstinRaw && String(gstinRaw).trim() !== "";
    const hasDescription = keywordFallbackAvailable && row[descField.columnIndex] && String(row[descField.columnIndex]).trim() !== "";
    const hasFlagColumn = columns.reverseCharge !== null;
    if (!hasGstin && !hasDescription && !hasFlagColumn) return; // nothing to key this row on at all

    const result = detectRcmForRow(row, columns, keywordFallbackAvailable ? descField.columnIndex : null);
    if (!result.isRcm) return;

    flagged.push({
      rowIndex: i,
      gstin: hasGstin ? String(gstinRaw).trim().toUpperCase() : null, // null = unregistered supplier — a meaningful, expected value for this rule, not missing data
      invoiceNumber: columns.invoiceNumber !== null ? row[columns.invoiceNumber] : columns.voucherNumber !== null ? row[columns.voucherNumber] : null,
      source: result.source,
      matchedCategory: result.matchedCategory ? { id: result.matchedCategory.id, label: result.matchedCategory.label, notification: result.matchedCategory.notification, requiresManualConfirmation: !!result.matchedCategory.requiresManualConfirmation } : null,
    });
  });

  return {
    flagged,
    keywordFallbackAvailable,
    keywordFallbackReason: descField.reason,
    flagSignalAvailable: columns.reverseCharge !== null,
  };
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { detectRcmForRow, detectRcmForSheet, readReverseChargeFlag, matchRcmKeyword };
}
