/*
 * gst-role-recognizer.js
 * ------------------------
 * Scores a sheet against known GST "roles" — a GSTR-2A-style export
 * (auto-drafted by the government from suppliers' own GSTR-1 filings, per
 * CGST Rule 60) or a purchase-register-style internal bookkeeping record
 * (the account a business is legally required to maintain under CGST Rule
 * 56) — using structural shape and standardized GST/accounting vocabulary,
 * never any file-specific data.
 *
 * Both roles share a lot of surface structure (GSTIN column, tax columns,
 * amounts) since they describe the same underlying purchases. What tells
 * them apart is which SIDE of the paperwork each one is: GSTR-2A is a
 * government report about what SUPPLIERS reported (so it carries
 * supplier-facing concepts — Trade/Legal Name, Place of Supply, Invoice
 * Type, Filing Period — that only make sense in a government e-return, and
 * never carries internal bookkeeping concepts like a Voucher Number, since
 * that's the buyer's own internal reference, not something a supplier's
 * filing would report). A purchase register is the mirror image — it's
 * internal bookkeeping, so it carries Voucher Number/Particulars and never
 * carries the government-report-only fields. Each rule below is grounded
 * in that real distinction, not in any one company's spreadsheet layout.
 *
 * Architected for extension: this file only knows about GST_ROLES named
 * here. A later workflow (e.g. GSTR-1 vs GSTR-3B) adds its own new role
 * names and rule sets alongside these, not a rewrite of this scoring
 * mechanism — see gst-workflows.js for how roles compose into workflows.
 */

const { identifyGstColumns } = require("./gst-column-identifier.js");

const GST_ROLES = ["gstr2a", "purchase_register"];

const MIN_SCORE_TO_RECOGNIZE = 5;
const MIN_MARGIN_TO_RECOGNIZE = 2;

const ROLE_LABELS = {
  gstr2a: "GSTR-2A (auto-drafted supplier data)",
  purchase_register: "Purchase Register (Books)",
  unknown: "Not recognized as GST data",
};

const ROLE_RULES = {
  gstr2a: [
    { weight: 3, label: "has a column matching the legal GSTIN format", test: (cols) => cols.gstin !== null },
    { weight: 2, label: "has a Taxable Value column", test: (cols) => cols.taxableValue !== null },
    {
      weight: 3,
      label: "has at least two of IGST/CGST/SGST identified (inter-state vs intra-state tax split required by GST law)",
      test: (cols) => [cols.igst, cols.cgst, cols.sgst].filter((v) => v !== null).length >= 2,
    },
    { weight: 2, label: "has an Invoice Number column", test: (cols) => cols.invoiceNumber !== null },
    { weight: 2, label: "has a Trade/Legal Name column (supplier-reported identity — only meaningful in a government return)", test: (cols) => cols.tradeLegalName !== null },
    {
      weight: 2,
      label: "has a government-return-only column (Place of Supply / Invoice Type / Filing Period)",
      test: (cols) => cols.placeOfSupply !== null || cols.invoiceType !== null || cols.filingPeriod !== null,
    },
    {
      // Negative: these are the buyer's OWN internal bookkeeping reference
      // numbers — a supplier's GSTR-1 filing (which is what 2A is drafted
      // from) has no way to know or report them.
      weight: -5,
      label: "has internal-bookkeeping-only columns (Voucher Number / Particulars) that a government return would never carry",
      test: (cols) => cols.voucherNumber !== null || cols.particulars !== null,
    },
  ],
  purchase_register: [
    { weight: 2, label: "has a date column", test: (cols) => cols.dateColumns.length > 0 },
    { weight: 3, label: "has a Voucher Number column (internal bookkeeping reference)", test: (cols) => cols.voucherNumber !== null },
    { weight: 2, label: "has a Particulars/Narration column (internal bookkeeping)", test: (cols) => cols.particulars !== null },
    { weight: 3, label: "has a column matching the legal GSTIN format", test: (cols) => cols.gstin !== null },
    { weight: 2, label: "has a Taxable Value column", test: (cols) => cols.taxableValue !== null },
    { weight: 1, label: "has at least one of IGST/CGST/SGST identified", test: (cols) => [cols.igst, cols.cgst, cols.sgst].some((v) => v !== null) },
    {
      // Negative: a business's own purchase ledger has no reason to record
      // fields that only exist because a return was filed on the GST portal.
      weight: -4,
      label: "has government-return-only columns (Place of Supply / Invoice Type / Filing Period) that an internal ledger would never carry",
      test: (cols) => cols.placeOfSupply !== null || cols.invoiceType !== null || cols.filingPeriod !== null,
    },
  ],
};

function scoreRole(identifiedColumns, role) {
  const rules = ROLE_RULES[role];
  const matchedRules = [];
  let score = 0;
  for (const rule of rules) {
    if (rule.test(identifiedColumns)) {
      score += rule.weight;
      matchedRules.push({ label: rule.label, weight: rule.weight });
    }
  }
  return { score, matchedRules };
}

// A sheet with no GSTIN column at all isn't either role — both a GSTR-2A
// export and a purchase register are, by definition, per-supplier records,
// and the GSTIN is the one field GST law requires to identify a supplier.
function recognizeGstRole(sheetSignals, values, headerRowIndex) {
  const columns = identifyGstColumns(sheetSignals, values, headerRowIndex);

  if (columns.gstin === null) {
    return { role: "unknown", label: ROLE_LABELS.unknown, confidence: "none", scores: {}, matchedRules: [], columns, reason: "no column matches the legal GSTIN format" };
  }

  const scored = GST_ROLES.map((role) => ({ role, ...scoreRole(columns, role) }));
  scored.sort((a, b) => b.score - a.score);

  const scores = {};
  scored.forEach((s) => (scores[s.role] = Math.round(s.score * 100) / 100));

  const best = scored[0];
  const runnerUp = scored[1];
  const margin = best.score - runnerUp.score;

  if (best.score < MIN_SCORE_TO_RECOGNIZE || margin < MIN_MARGIN_TO_RECOGNIZE) {
    return {
      role: "unknown",
      label: ROLE_LABELS.unknown,
      confidence: "none",
      scores,
      matchedRules: [],
      columns,
      reason:
        best.score < MIN_SCORE_TO_RECOGNIZE
          ? `no role scored at least ${MIN_SCORE_TO_RECOGNIZE} (best was ${best.role} at ${scores[best.role]})`
          : `${best.role} (${scores[best.role]}) didn't clear ${runnerUp.role} (${scores[runnerUp.role]}) by the required margin of ${MIN_MARGIN_TO_RECOGNIZE}`,
    };
  }

  const confidence = margin >= 4 ? "high" : "medium";

  return {
    role: best.role,
    label: ROLE_LABELS[best.role],
    confidence,
    scores,
    matchedRules: best.matchedRules,
    columns,
  };
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { recognizeGstRole, GST_ROLES, ROLE_LABELS, MIN_SCORE_TO_RECOGNIZE, MIN_MARGIN_TO_RECOGNIZE };
}
