/*
 * gst-cross-sheet-column-search.js
 * -----------------------------------
 * MENTOR "internal search before asking" (Point 3 of the standing
 * architecture request). Before a "which column is X?" prompt fires for an
 * ambiguous field on one sheet, checks whether ANOTHER sheet in the SAME
 * workbook — already classified as playing the SAME GST role (gstr2a or
 * purchase_register; see gst-role-recognizer.js) — already has an
 * UNAMBIGUOUS answer for that same field. That's real corroborating
 * evidence from this exact workbook's own data, not a guess imported from
 * some other client's spreadsheet (Tier 1) or a decision this client made
 * on a prior visit (Tier 2) — it's this workbook confirming itself, right
 * now.
 *
 * Deliberately restricted to same-role sheets, not "any sheet in the
 * workbook" — a real client file (GST_Reco_Sheeet_up_to_Dec.xlsx) has two
 * different reconciliation types (GSTR-1-vs-3B and 3B-vs-Books) coexisting
 * in one workbook. Searching across role/type boundaries risks matching a
 * column from an unrelated reconciliation just because the header text
 * happens to look similar. The caller is responsible for the role filter
 * (it already has gst-role-recognizer.js's per-sheet results, computed
 * once per scan for every sheet in the workbook) — this module has no
 * notion of "role" itself, so it stays reusable unmodified for a future
 * workflow's own role names.
 *
 * Only ever invoked for genuinely AMBIGUOUS fields (2+ real candidates on
 * the sheet being resolved) — same scoping gst-column-ambiguity-rules.js
 * and Tier 1 already use, for the same reason: a zero-match-required
 * field's "candidates" are just every column on the sheet, not a real
 * signal to corroborate against.
 *
 * Confidence bar, matching the rest of this codebase's "don't silently
 * guess" posture: resolves SILENTLY only on an EXACT normalized-header
 * match — every corroborating sheet that has an opinion must agree, and
 * that agreed header must be one of this sheet's own candidates. Anything
 * less (no agreement, or agreement on a header this sheet doesn't have) is
 * surfaced only as a pre-filled SUGGESTION for the still-required prompt —
 * the caller renders it, but MUST still ask. Reuses structural-
 * signature.js's own header-similarity machinery and the same 0.5 cutoff
 * sheet-memory-store-base.js already established for "worth surfacing as a
 * fuzzy match" (HEADER_SIMILARITY_THRESHOLD), rather than inventing a new
 * number.
 *
 * A hit here is NEVER written back into column-memory's Tier 2 (per-client
 * store) or promoted into Tier 1 (shared store) — it's a live, software-
 * inferred corroboration, re-checked on every scan, exactly like a Tier-1
 * hit already is (see mentor-column-memory-ui.js's docstring). Only an
 * explicit human answer, via the actual prompt, ever populates either
 * memory tier.
 */

const { normalizeHeaderText, singleHeaderSimilarity } = require("../sheet-classifier/structural-signature.js");

const SUGGESTION_SIMILARITY_THRESHOLD = 0.5;

// otherSheetResults: [{ sheetName, columns, sheetSignals }, ...] — every
// OTHER sheet in the workbook the caller has already filtered down to the
// SAME role as the sheet currently being resolved. `columns` is
// identifyGstColumns()'s result (has `.candidates[field]`, an array of
// column indices); `sheetSignals` supplies the actual header text for
// those indices.
//
// candidateHeaders: this sheet's own ambiguous candidates for the field —
// [{ header, index, ... }, ...], the same shape already built for the
// prompt UI and for Tier 1 lookups.
//
// -> { matchType: "exact", matchedHeader, sourceSheetName }
//  | { matchType: "fuzzy", matchedHeader, sourceSheetName, similarity }
//  | { matchType: "none" }
function findCrossSheetColumnMatch(field, candidateHeaders, otherSheetResults) {
  const normalizedCandidates = candidateHeaders.map((c) => ({ ...c, normalized: normalizeHeaderText(c.header) }));

  // Every OTHER same-role sheet's own UNAMBIGUOUS answer for this field —
  // "unambiguous" meaning ITS OWN header rule matched exactly one column on
  // THAT sheet, not that a human confirmed it. A sheet where this field is
  // itself ambiguous (or wasn't found at all) has nothing to corroborate
  // with, so it's skipped rather than guessed from.
  const corroborations = [];
  for (const other of otherSheetResults || []) {
    const otherCandidateIndices = (other.columns && other.columns.candidates && other.columns.candidates[field]) || [];
    if (otherCandidateIndices.length !== 1) continue;
    const otherColumn = (other.sheetSignals.columns || []).find((c) => c.index === otherCandidateIndices[0]);
    if (!otherColumn) continue;
    corroborations.push({ sourceSheetName: other.sheetName, header: otherColumn.header, normalized: normalizeHeaderText(otherColumn.header) });
  }

  if (corroborations.length === 0) return { matchType: "none" };

  // Exact: every corroborating sheet that has an opinion must agree on the
  // SAME normalized header, AND that header must be one of this sheet's
  // own candidates. A lone corroborator is enough (most fields only ever
  // have one other same-role sheet in practice) — but disagreement between
  // two or more corroborators never resolves silently, it falls through to
  // the fuzzy suggestion below instead.
  const distinctNormalized = new Set(corroborations.map((c) => c.normalized));
  if (distinctNormalized.size === 1) {
    const [normalized] = distinctNormalized;
    const exactMatch = normalizedCandidates.find((c) => c.normalized === normalized);
    if (exactMatch) {
      return { matchType: "exact", matchedHeader: exactMatch.header, sourceSheetName: corroborations[0].sourceSheetName };
    }
  }

  // Fuzzy fallback — best single-header similarity across every
  // corroborator x candidate pair, regardless of whether corroborators
  // agreed with each other. Safe to take the single best score here since
  // this is only ever used as a pre-filled SUGGESTION for a prompt that
  // still fires, never to resolve silently.
  let best = null;
  for (const corroboration of corroborations) {
    for (const candidate of normalizedCandidates) {
      const similarity = singleHeaderSimilarity(corroboration.normalized, candidate.normalized);
      if (!best || similarity > best.similarity) {
        best = { matchedHeader: candidate.header, sourceSheetName: corroboration.sourceSheetName, similarity };
      }
    }
  }

  if (best && best.similarity >= SUGGESTION_SIMILARITY_THRESHOLD) {
    return { matchType: "fuzzy", matchedHeader: best.matchedHeader, sourceSheetName: best.sourceSheetName, similarity: best.similarity };
  }

  return { matchType: "none" };
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { findCrossSheetColumnMatch, SUGGESTION_SIMILARITY_THRESHOLD };
}
