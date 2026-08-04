/*
 * gst-summary-builder.js
 * -------------------------
 * Nets Step 1's vendor-level GSTR-2A vs Books comparison together with Step
 * 2's line-item categorization (Wrong Head / RCM / Ineligible ITC /
 * Duplicate Invoices) into one proposed "how much ITC does this actually
 * support" figure — the final synthesis a human reconciler does by hand
 * after finishing the individual checks.
 *
 * This is new logic (Step 1 and Step 2 themselves are untouched), so the
 * choices made here are documented explicitly:
 *
 * - Base figure ("provisionally supportable") = per vendor, per tax head
 *   (IGST/CGST/SGST), the LOWER of what 2A shows and what Books shows,
 *   summed across every recognized GSTIN. ITC can't legitimately exceed
 *   what the government return reflects (Section 16(2)(aa) / Rule 36(4))
 *   NOR exceed what was actually booked as an expense — so the min of the
 *   two is what's currently supportable by BOTH records at once. This one
 *   formula also correctly handles vendors missing from either side without
 *   a separate case: if a vendor is absent from 2A, its 2A total is 0, so
 *   min(...) is 0 — nothing from an unfiled supplier is counted as
 *   supportable yet, without needing a special-cased subtraction for it.
 *
 * - Ineligible ITC (Section 17(5) matches) and Duplicate Invoice over-claims
 *   are subtracted from that base — both actually shrink the credit a
 *   taxpayer can validly claim.
 *
 * - Wrong Head and RCM findings are reported INFORMATIONALLY ONLY, never
 *   subtracted here: Wrong Head doesn't change the amount of credit, only
 *   which head it should sit under (CGST+SGST vs IGST) — netting it would
 *   double-count the same rupees as a "loss." RCM tax is self-assessed and
 *   paid by the recipient through a separate mechanism, not claimed as
 *   ordinary vendor-matched ITC via this reconciliation at all — flagging
 *   it here is a prompt to verify it's handled correctly elsewhere, not a
 *   deduction from this total.
 *
 * - Extra in 2A / Extra in Books are reported as separate, NOT netted,
 *   lines — resolving them requires investigating the specific invoice
 *   (supplier hasn't filed yet vs. genuinely not the taxpayer's purchase),
 *   which this pipeline can't determine on its own. Consistent with the
 *   project's established rule: don't force a number where the honest
 *   answer is "needs investigation."
 *
 * Every figure here is explicitly the PROPOSED reconciliation for human
 * review — MENTOR's "propose, don't auto-execute" rule applies to this
 * synthesis exactly as it does to Step 1 and Step 2 individually.
 */

function roundTo2(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

function emptyStepTwoTotals() {
  return {
    wrongHead: { rowCount: 0 },
    rcm: { rowCount: 0, amount: 0 },
    ineligibleItc: { rowCount: 0, amount: 0 },
    duplicates: { clusterCount: 0, overclaimAmount: 0 },
  };
}

// comparisonSummary: buildReconciliationSummary()'s return value from
// gstr2a-vs-books-reconciler.js, unmodified.
// stepTwoTotals: pre-summed totals from the four Step 2 rules (see
// gst-report-writer.js, which enriches each rule's flagged rows with real
// rupee amounts before summing — this function does no sheet lookups of
// its own, purely arithmetic over numbers it's handed, so it's testable
// with plain synthetic figures).
function buildGstSummary(comparisonSummary, stepTwoTotals) {
  const totals = stepTwoTotals || emptyStepTwoTotals();
  const vendors = comparisonSummary.vendors || [];

  let totalTaxPer2A = 0;
  let totalTaxPerBooks = 0;
  let provisionallySupportable = 0;
  let extraIn2ACount = 0;
  let extraIn2AAmount = 0;
  let extraInBooksCount = 0;
  let extraInBooksAmount = 0;
  let mismatchCount = 0;
  let mismatchUnsupportedAmount = 0;

  for (const v of vendors) {
    const gTotal = v.gstr2a.igst + v.gstr2a.cgst + v.gstr2a.sgst;
    const bTotal = v.books.igst + v.books.cgst + v.books.sgst;
    totalTaxPer2A += gTotal;
    totalTaxPerBooks += bTotal;

    provisionallySupportable += Math.min(v.gstr2a.igst, v.books.igst) + Math.min(v.gstr2a.cgst, v.books.cgst) + Math.min(v.gstr2a.sgst, v.books.sgst);

    if (v.status === "missing_in_books") {
      extraIn2ACount++;
      extraIn2AAmount += gTotal;
    } else if (v.status === "missing_in_gstr2a") {
      extraInBooksCount++;
      extraInBooksAmount += bTotal;
    } else if (v.status === "discrepancy") {
      mismatchCount++;
      mismatchUnsupportedAmount += Math.abs(gTotal - bTotal);
    }
  }

  const lessIneligibleItc = totals.ineligibleItc.amount;
  const lessDuplicateOverclaim = totals.duplicates.overclaimAmount;
  const netProvisionallyEligible = provisionallySupportable - lessIneligibleItc - lessDuplicateOverclaim;

  return {
    tolerance: comparisonSummary.tolerance,
    totalGstins: comparisonSummary.totalGstins,
    vendorCounts: comparisonSummary.counts,

    totalTaxPer2A: roundTo2(totalTaxPer2A),
    totalTaxPerBooks: roundTo2(totalTaxPerBooks),
    provisionallySupportable: roundTo2(provisionallySupportable),

    lessIneligibleItc: { rowCount: totals.ineligibleItc.rowCount, amount: roundTo2(lessIneligibleItc) },
    lessDuplicateOverclaim: { clusterCount: totals.duplicates.clusterCount, amount: roundTo2(lessDuplicateOverclaim) },
    netProvisionallyEligible: roundTo2(netProvisionallyEligible),

    extraIn2A: { count: extraIn2ACount, amount: roundTo2(extraIn2AAmount) },
    extraInBooks: { count: extraInBooksCount, amount: roundTo2(extraInBooksAmount) },
    mismatch: { count: mismatchCount, unsupportedAmount: roundTo2(mismatchUnsupportedAmount) },

    informational: {
      wrongHeadRowCount: totals.wrongHead.rowCount,
      rcmRowCount: totals.rcm.rowCount,
      rcmAmount: roundTo2(totals.rcm.amount),
    },
  };
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { buildGstSummary, emptyStepTwoTotals };
}
