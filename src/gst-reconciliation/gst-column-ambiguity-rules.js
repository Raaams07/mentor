/*
 * gst-column-ambiguity-rules.js
 * -------------------------------
 * Decides which fields on an already-role-recognized GST sheet (gstr2a or
 * purchase_register) need a memory lookup / user prompt before their
 * resolved column can be trusted — the "third state" (ambiguous, don't
 * guess) that identifyGstColumns()'s first-match-wins loop never produced.
 *
 * Deliberately separate from gst-column-identifier.js: this only runs
 * AFTER a workflow has already identified the two specific sheets playing
 * the gstr2a/purchase_register roles (see mentor-gst-reconciliation-ui.js),
 * never during role recognition itself — role recognition scans every
 * sheet in the workbook on every scan, and firing an ask there would
 * prompt on random unrelated sheets that happen to have two columns
 * matching some header rule.
 *
 * Two different triggers, deliberately scoped differently:
 *   - Ambiguity (2+ candidates) applies to EVERY field with a header rule,
 *     regardless of whether this particular role/reconciliation needs it —
 *     this is what generalizes the Tally CGST bug fix to the full field
 *     set, not just tax columns.
 *   - Zero-match applies ONLY to fields that are actually load-bearing for
 *     the role already recognized. A purchase register legitimately has no
 *     Filing Period column, no Place of Supply column, etc. — asking about
 *     those on every zero-match would be constant, useless noise. Fields
 *     outside the required set only ever trigger on ambiguity, never on
 *     absence.
 */

// A genuine GSTIN column normally scores ~0.9-1.0 (nearly every row is a
// well-formed GSTIN); a coincidental partial match usually scores well
// under 0.8. Two columns landing within this margin of each other at the
// top is the specific signature of real ambiguity (e.g. both a "Recipient
// GSTIN" and "Supplier GSTIN" column present), not everyday scoring noise —
// mirrors how gst-role-recognizer.js's own MIN_MARGIN_TO_RECOGNIZE
// separates a genuine close call from routine noise.
const GSTIN_MARGIN_TO_RESOLVE = 0.1;

// Every field with a HEADER_RULES entry in gst-column-identifier.js except
// gstin (content-based, handled separately below).
const AMBIGUITY_CHECKED_FIELDS = [
  "taxableValue",
  "igst",
  "cgst",
  "sgst",
  "rate",
  "invoiceNumber",
  "voucherNumber",
  "particulars",
  "tradeLegalName",
  "placeOfSupply",
  "invoiceType",
  "filingPeriod",
  "reverseCharge",
];

// role: "gstr2a" | "purchase_register". candidates: identifyGstColumns()'s
// result.candidates. Returns [{ field, reason, candidateIndices }, ...] —
// every field that needs either a memory lookup or a user prompt before
// this sheet's columns can be trusted. candidateIndices is [] for a
// zero-match-required field — the caller falls back to offering every
// column in the sheet as a pickable candidate.
function fieldsInPlay(role, candidates) {
  const inPlay = [];

  for (const field of AMBIGUITY_CHECKED_FIELDS) {
    if (candidates[field].length >= 2) {
      inPlay.push({ field, reason: "ambiguous", candidateIndices: candidates[field] });
    }
  }

  if (candidates.gstin.length >= 2) {
    const margin = candidates.gstin[0].ratio - candidates.gstin[1].ratio;
    if (margin < GSTIN_MARGIN_TO_RESOLVE) {
      inPlay.push({ field: "gstin", reason: "ambiguous", candidateIndices: candidates.gstin.map((c) => c.index) });
    }
  }

  if (candidates.taxableValue.length === 0) {
    inPlay.push({ field: "taxableValue", reason: "zero_match_required", candidateIndices: [] });
  }

  if (role === "gstr2a" && candidates.invoiceNumber.length === 0) {
    inPlay.push({ field: "invoiceNumber", reason: "zero_match_required", candidateIndices: [] });
  }
  if (role === "purchase_register" && candidates.invoiceNumber.length === 0 && candidates.voucherNumber.length === 0) {
    inPlay.push({ field: "invoiceNumber", reason: "zero_match_required_either", candidateIndices: [] });
    inPlay.push({ field: "voucherNumber", reason: "zero_match_required_either", candidateIndices: [] });
  }

  // Defensive, expected to be rare in practice: role recognition already
  // requires evidence of tax columns to clear its score threshold
  // (gst-role-recognizer.js ROLE_RULES), so reaching this with all three
  // simultaneously zero on an already-recognized sheet is a pathology, not
  // a routine case. A single tax field being zero (e.g. an IGST-only
  // vendor register with no CGST/SGST) is normal and must NOT ask.
  if (candidates.igst.length === 0 && candidates.cgst.length === 0 && candidates.sgst.length === 0) {
    inPlay.push({ field: "igst", reason: "zero_match_all_tax_zero", candidateIndices: [] });
    inPlay.push({ field: "cgst", reason: "zero_match_all_tax_zero", candidateIndices: [] });
    inPlay.push({ field: "sgst", reason: "zero_match_all_tax_zero", candidateIndices: [] });
  }

  // rate zero-match deliberately NOT added here — the rate-sanity-check
  // backstop just silently skips a sheet with no resolved rate column,
  // since it's a nice-to-have, not a core reconciliation field.

  return inPlay;
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { fieldsInPlay, GSTIN_MARGIN_TO_RESOLVE, AMBIGUITY_CHECKED_FIELDS };
}
