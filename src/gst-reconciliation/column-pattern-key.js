/*
 * column-pattern-key.js
 * ------------------------
 * Pure, isomorphic key derivation for Tier 1 (shared software-pattern)
 * column-resolution memory — see column-memory.js for Tier 2 (per-client).
 *
 * Deliberately COARSER than Tier 2's structural signature
 * (structural-signature.js's computeStructuralSignature, which fingerprints
 * a whole sheet's shape): Tier 1's key is just a field name + the set of
 * candidate header texts that were ambiguous for it, order-independent and
 * content-free. This is what makes it a signature of the EXPORTING
 * SOFTWARE's layout (Tally, Zoho Books, Marg, ...) rather than any one
 * client's workbook shape — the same ambiguous header set recurs across
 * every workbook that software produces, regardless of what else is on
 * the sheet or what order the columns happen to be in.
 *
 * No sheet name, no column index, no sample values, no client identifier —
 * only header text. This file has zero Node-only or browser-only
 * dependencies (same as structural-signature.js/classifier.js, which it
 * builds on) — required identically by server/column-pattern-store.js
 * (plain `node` process) and this file's own test.
 */

const { normalizeHeaderText } = require("../sheet-classifier/structural-signature.js");

// Normalizes every header, drops blanks and post-normalization duplicates,
// sorts — deterministic regardless of input order, casing, or whitespace.
function normalizeCandidateHeaders(headers) {
  const seen = new Set();
  const out = [];
  for (const h of headers || []) {
    const norm = normalizeHeaderText(h);
    if (!norm || seen.has(norm)) continue;
    seen.add(norm);
    out.push(norm);
  }
  return out.sort();
}

function computeColumnPatternKey(fieldName, candidateHeaders) {
  return `${fieldName}::${normalizeCandidateHeaders(candidateHeaders).join("|")}`;
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { computeColumnPatternKey, normalizeCandidateHeaders };
}
