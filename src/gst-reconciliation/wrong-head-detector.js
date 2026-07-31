/*
 * wrong-head-detector.js
 * ------------------------
 * "Wrong Head" detection: a supplier's registered state (the first 2
 * digits of their GSTIN — see gst-state-codes.js) and the invoice's
 * "Place of Supply" together determine, as a matter of law, which tax
 * head should have been charged:
 *
 *   - Same state as the Place of Supply  -> intra-state supply
 *     -> CGST + SGST (IGST Act Section 8)
 *   - Different state from Place of Supply -> inter-state supply
 *     -> IGST (IGST Act Section 7)
 *
 * This is row-level, not GSTIN-level — a single supplier can legitimately
 * ship to different places of supply on different invoices, so each
 * invoice line is checked against its OWN place of supply, not a single
 * assumption per vendor.
 *
 * Only meaningful for sheets where a Place of Supply column was
 * identified (see gst-column-identifier.js) — normally GSTR-2A, since
 * Place of Supply is a government-return-only field per the existing
 * role recognizer.
 */

const { stateCodeFromGstin, extractStateCodeFromPlaceOfSupply } = require("./gst-state-codes.js");

const TOLERANCE = 1; // ₹1 — treat sub-rupee tax amounts as effectively zero for head determination

function toNumber(value) {
  if (typeof value === "number") return value;
  if (value === "" || value === null || value === undefined) return 0;
  const parsed = parseFloat(String(value).replace(/[,$£€¥%()]/g, ""));
  return isNaN(parsed) ? 0 : parsed;
}

// Which tax head is ACTUALLY populated on this row, from the raw amounts —
// not what should have been charged, just what was.
function actualTaxHead(igst, cgst, sgst) {
  const hasIgst = Math.abs(igst) > TOLERANCE;
  const hasCgstSgst = Math.abs(cgst) > TOLERANCE || Math.abs(sgst) > TOLERANCE;
  if (hasIgst && hasCgstSgst) return "mixed"; // both populated — not valid under GST law for one invoice line
  if (hasIgst) return "igst";
  if (hasCgstSgst) return "cgst_sgst";
  return "none"; // no tax at all — likely nil-rated/exempt, nothing to judge here
}

// values/headerRowIndex: the sheet's raw data + header position.
// columns: identifyGstColumns() result — needs gstin, placeOfSupply, igst,
// cgst, sgst at minimum; invoiceNumber is used only to label output rows.
function detectWrongHead(values, headerRowIndex, columns) {
  if (columns.gstin === null || columns.placeOfSupply === null) {
    return { applicable: false, reason: "sheet has no GSTIN or Place of Supply column identified", flagged: [] };
  }

  const dataRows = headerRowIndex === -1 ? values : values.slice(headerRowIndex + 1);
  const flagged = [];

  dataRows.forEach((row, i) => {
    const gstinRaw = row[columns.gstin];
    if (!gstinRaw || String(gstinRaw).trim() === "") return;
    const gstin = String(gstinRaw).trim().toUpperCase();

    const placeOfSupplyRaw = row[columns.placeOfSupply];
    const supplierStateCode = stateCodeFromGstin(gstin);
    const placeOfSupplyStateCode = extractStateCodeFromPlaceOfSupply(placeOfSupplyRaw);
    if (!supplierStateCode || !placeOfSupplyStateCode) return; // can't judge without both

    const igst = columns.igst !== null ? toNumber(row[columns.igst]) : 0;
    const cgst = columns.cgst !== null ? toNumber(row[columns.cgst]) : 0;
    const sgst = columns.sgst !== null ? toNumber(row[columns.sgst]) : 0;

    const actual = actualTaxHead(igst, cgst, sgst);
    if (actual === "none") return; // nothing charged — not this rule's concern

    const expected = supplierStateCode === placeOfSupplyStateCode ? "cgst_sgst" : "igst";

    if (actual === "mixed") {
      flagged.push({
        rowIndex: i,
        gstin,
        invoiceNumber: columns.invoiceNumber !== null ? row[columns.invoiceNumber] : null,
        issue: "mixed_tax_heads",
        expected,
        actual,
        supplierStateCode,
        placeOfSupplyStateCode,
        igst,
        cgst,
        sgst,
      });
    } else if (actual !== expected) {
      flagged.push({
        rowIndex: i,
        gstin,
        invoiceNumber: columns.invoiceNumber !== null ? row[columns.invoiceNumber] : null,
        issue: "wrong_head",
        expected,
        actual,
        supplierStateCode,
        placeOfSupplyStateCode,
        igst,
        cgst,
        sgst,
      });
    }
  });

  return { applicable: true, flagged };
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { detectWrongHead, actualTaxHead };
}
