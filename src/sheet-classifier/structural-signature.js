/*
 * structural-signature.js
 * ------------------------
 * MENTOR Phase 2, sheet memory — "same shape, different name" detection.
 *
 * Turns a sheet's Step 1 signals into a fingerprint that survives a rename
 * and normal data growth, but changes when the sheet's actual structure
 * changes. See the design discussion this was built from: it deliberately
 * keeps only a coarse per-column shape bucket (reusing classifier.js's own
 * threshold predicates, not new ones) and throws away anything expected to
 * drift over time — exact header text, row count, exact percentages.
 */

const {
  isDateColumn,
  isPeriodicDateColumn,
  isNumericColumn,
  isHighCardinalityNumeric,
  isNonNettingNumeric,
  isModerateCardinalityText,
  isHighCardinalityText,
  isCategoricalColumn,
} = require("./classifier.js");

// A column mostly missing carries no shape information worth fingerprinting.
function tagColumn(col) {
  if (!col || col.nonBlankCount === 0) return "BLANK";

  if (isPeriodicDateColumn(col)) return "DATE_PERIODIC";
  if (isDateColumn(col)) return "DATE_TRANSACTIONAL";

  if (isNumericColumn(col)) {
    if (!isHighCardinalityNumeric(col)) return "NUM_LOW"; // rate/flag/code-like, not an amount
    return isNonNettingNumeric(col) ? "NUM_HIGH_NONNETTING" : "NUM_HIGH_SIGNED";
  }

  // Text, checked most-specific first: unique-per-row, then small fixed set,
  // then the vendor-variety band, then a catch-all for anything left over.
  if (isHighCardinalityText(col)) return "TEXT_UNIQUE";
  if (isCategoricalColumn(col)) return "TEXT_CATEGORICAL";
  if (isModerateCardinalityText(col)) return "TEXT_MODERATE";
  return "TEXT_OTHER";
}

// Header text is normalized (lowercased, punctuation stripped) purely for
// COMPARISON purposes below — it's never used as a match key on its own,
// only as a secondary corroborating signal alongside the structural tags.
function normalizeHeaderText(header) {
  return String(header === undefined || header === null ? "" : header)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function computeStructuralSignature(sheetSignals) {
  const columns = (sheetSignals && sheetSignals.columns) || [];
  const tags = columns.map(tagColumn);
  const headerTokens = columns.map((c) => normalizeHeaderText(c && c.header));

  return {
    signature: `${tags.length}:${tags.join(",")}`,
    tags,
    columnCount: tags.length,
    headerTokens,
    // Placeholder for blank headers so position information survives the
    // join/split round-trip through storage — an empty string would collapse
    // against the delimiter.
    headerSignature: headerTokens.map((h) => h || "∅").join("|"),
  };
}

// Is this signature specific enough to safely promote to the SHARED,
// cross-client Tier-1 pattern store (sheet-label-pattern-store.js)?
//
// The coarseness that makes a structural signature rename-proof (e.g.
// "2:TEXT_UNIQUE,NUM_HIGH_SIGNED") also makes a SHORT or type-uniform one a
// poor cross-client key: plenty of unrelated real-world sheet shapes could
// legitimately produce the same few tags, so a label learned from one
// client's sheet could silently misapply to a structurally-coincidental
// but actually-different sheet from another client. This is a distinct
// question from classifier.js's own confidence scoring, which measures how
// closely a sheet matches one of the KNOWN types — it says nothing about
// how distinctive an UNKNOWN shape is, and sheet-memory only ever reaches
// the "needs_input"/promotion path for sheets the classifier already
// declined to place into any known type.
//
// Two-part bar, deliberately mirroring classifier.js's own two-part
// MIN_SCORE_TO_CLASSIFY + MIN_MARGIN_TO_CLASSIFY gate: an absolute floor
// (enough columns to carry real shape information) plus a diversity floor
// (not just N copies of the same column type, e.g. a single list of names
// padded out to five columns isn't evidence of a distinctive software-
// export shape). Both thresholds are deliberately conservative — false
// negatives here just mean "ask this one client once more than strictly
// necessary"; false positives mean a wrong label silently reaching every
// other client, which is the more expensive mistake.
const MIN_COLUMNS_FOR_SHARED_SHEET_PATTERN = 4;
const MIN_DISTINCT_TAGS_FOR_SHARED_SHEET_PATTERN = 2;

function isStructuralSignatureShareable(sheetSignals) {
  const { tags, columnCount } = computeStructuralSignature(sheetSignals);
  if (columnCount < MIN_COLUMNS_FOR_SHARED_SHEET_PATTERN) return false;
  const distinctTags = new Set(tags.filter((t) => t !== "BLANK"));
  return distinctTags.size >= MIN_DISTINCT_TAGS_FOR_SHARED_SHEET_PATTERN;
}

// Generic weighted edit distance: like classic Levenshtein, but the
// substitution cost between two items is whatever costFn returns (0 = same,
// 1 = completely different) instead of a hard equal/not-equal check. Both
// tag-sequence comparison (binary cost) and header-text comparison (fuzzy
// string-similarity cost) below are built on this one function — same
// technique this codebase already uses for fuzzy vendor-name matching (see
// levenshteinDistance in "Mentor workbook index/mentor-workbook-index.js"),
// generalized so a "substitution" can be partial instead of all-or-nothing.
function weightedEditDistance(itemsA, itemsB, costFn) {
  const m = itemsA.length;
  const n = itemsB.length;
  const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));

  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;

  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      const cost = costFn(itemsA[i - 1], itemsB[j - 1]);
      dp[i][j] = Math.min(
        dp[i - 1][j] + 1, // deletion
        dp[i][j - 1] + 1, // insertion
        dp[i - 1][j - 1] + cost // substitution (partial cost allowed)
      );
    }
  }
  return dp[m][n];
}

function tokenLevenshteinDistance(tokensA, tokensB) {
  return weightedEditDistance(tokensA, tokensB, (a, b) => (a === b ? 0 : 1));
}

// 1.0 = identical tag sequence, 0.0 = nothing alike. Used for the fuzzy
// fallback when no exact signature match exists — tolerates a column being
// added/removed/reordered, or one column's bucket shifting, without losing
// the match entirely.
function tagSequenceSimilarity(tagsA, tagsB) {
  const maxLen = Math.max(tagsA.length, tagsB.length);
  if (maxLen === 0) return 1;
  const distance = tokenLevenshteinDistance(tagsA, tagsB);
  return 1 - distance / maxLen;
}

// Character-level fuzzy similarity between two individual header strings
// (already normalized). Reuses weightedEditDistance directly — JS strings
// index like arrays, so no separate character-array conversion is needed.
function singleHeaderSimilarity(a, b) {
  if (a === b) return 1;
  const maxLen = Math.max(a.length, b.length);
  if (maxLen === 0) return 1;
  return 1 - weightedEditDistance(a, b, (x, y) => (x === y ? 0 : 1)) / maxLen;
}

// Sequence-level header similarity: aligns two sheets' header lists (via the
// same insert/delete/substitute edit-distance machinery used for tags, so a
// column insertion doesn't misalign everything after it), but each
// substitution's cost is graded by how similar the two header strings
// actually are, not just whether they're byte-identical. 1.0 = same wording
// throughout, 0.0 = nothing in common.
function headerSequenceSimilarity(headerTokensA, headerTokensB) {
  const maxLen = Math.max(headerTokensA.length, headerTokensB.length);
  if (maxLen === 0) return 1;
  const distance = weightedEditDistance(headerTokensA, headerTokensB, (a, b) => 1 - singleHeaderSimilarity(a, b));
  return 1 - distance / maxLen;
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    computeStructuralSignature,
    isStructuralSignatureShareable,
    tagColumn,
    tagSequenceSimilarity,
    tokenLevenshteinDistance,
    normalizeHeaderText,
    headerSequenceSimilarity,
    singleHeaderSimilarity,
  };
}
