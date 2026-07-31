/*
 * ineligible-itc-detector.js
 * -----------------------------
 * Flags line items likely blocked from Input Tax Credit under Section
 * 17(5) CGST Act, by matching a genuine per-transaction expense
 * description against ineligible-itc-config.js's editable category list.
 *
 * Only ever matches against a field verified (see description-field.js) to
 * actually be a transaction-level description — never a vendor/company
 * name field, even when a keyword coincidentally appears inside one (a
 * company literally named "... Food Products Pvt Ltd" is not evidence
 * about what was purchased). If no such field exists on a sheet, this
 * rule reports itself as not applicable rather than silently returning a
 * count that looks like a real finding — a B2B raw-material purchase
 * register may legitimately have nothing for this rule to check, and
 * that's a correct answer, not something to work around.
 *
 * Deliberately conservative about what it claims even when a genuine
 * description field IS found: this is keyword matching, not a
 * determination of actual ITC eligibility (which depends on facts this
 * data usually doesn't carry — e.g. whether a motor vehicle purchase was
 * for further supply, an excepted business use). Every match is returned
 * with the matched keyword and category so a human reviews the actual
 * claim, consistent with MENTOR's "propose, don't auto-execute" rule —
 * this never removes or reclassifies a credit on its own.
 */

const { ITC_INELIGIBLE_CATEGORIES } = require("./ineligible-itc-config.js");
const { findDescriptionColumn } = require("./description-field.js");

function matchIneligibleItcKeyword(descriptionText) {
  if (!descriptionText) return null;
  const text = String(descriptionText).toLowerCase();
  for (const category of ITC_INELIGIBLE_CATEGORIES) {
    const matchedKeyword = category.keywords.find((kw) => text.includes(kw));
    if (matchedKeyword) return { category, matchedKeyword };
  }
  return null;
}

// row: raw row array. descriptionColumnIndex: the column verified genuine
// by findDescriptionColumn() — null if no such field is available, in
// which case this always returns not-ineligible (nothing to check).
function detectIneligibleItcForRow(row, descriptionColumnIndex) {
  if (descriptionColumnIndex === null) return { isIneligible: false, matchedCategory: null, matchedKeyword: null };

  const match = matchIneligibleItcKeyword(row[descriptionColumnIndex]);
  if (!match) return { isIneligible: false, matchedCategory: null, matchedKeyword: null };

  return {
    isIneligible: true,
    matchedCategory: { id: match.category.id, label: match.category.label, section: match.category.section },
    matchedKeyword: match.matchedKeyword,
  };
}

// values/headerRowIndex: the sheet's raw data + header position.
// columns: identifyGstColumns() result for this sheet.
function detectIneligibleItcForSheet(values, headerRowIndex, columns) {
  const descField = findDescriptionColumn(values, headerRowIndex, columns);
  if (!descField.available) {
    return { applicable: false, reason: descField.reason, flagged: [] };
  }

  const dataRows = headerRowIndex === -1 ? values : values.slice(headerRowIndex + 1);
  const flagged = [];

  dataRows.forEach((row, i) => {
    const gstinRaw = columns.gstin !== null ? row[columns.gstin] : null;
    if (!gstinRaw || String(gstinRaw).trim() === "") return;

    const result = detectIneligibleItcForRow(row, descField.columnIndex);
    if (!result.isIneligible) return;

    flagged.push({
      rowIndex: i,
      gstin: String(gstinRaw).trim().toUpperCase(),
      invoiceNumber: columns.invoiceNumber !== null ? row[columns.invoiceNumber] : columns.voucherNumber !== null ? row[columns.voucherNumber] : null,
      matchedCategory: result.matchedCategory,
      matchedKeyword: result.matchedKeyword,
    });
  });

  return { applicable: true, flagged };
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { detectIneligibleItcForRow, detectIneligibleItcForSheet, matchIneligibleItcKeyword };
}
