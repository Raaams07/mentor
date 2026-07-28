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

function computeStructuralSignature(sheetSignals) {
  const columns = (sheetSignals && sheetSignals.columns) || [];
  const tags = columns.map(tagColumn);

  return {
    signature: `${tags.length}:${tags.join(",")}`,
    tags,
    columnCount: tags.length,
  };
}

// Token-level Levenshtein distance — identical technique to the character-level
// Levenshtein this codebase already uses for fuzzy vendor-name matching
// (see levenshteinDistance in "Mentor workbook index/mentor-workbook-index.js"),
// just operating on an array of shape tags instead of an array of characters.
function tokenLevenshteinDistance(tokensA, tokensB) {
  const m = tokensA.length;
  const n = tokensB.length;
  const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));

  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;

  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      const cost = tokensA[i - 1] === tokensB[j - 1] ? 0 : 1;
      dp[i][j] = Math.min(
        dp[i - 1][j] + 1, // deletion
        dp[i][j - 1] + 1, // insertion
        dp[i - 1][j - 1] + cost // substitution
      );
    }
  }
  return dp[m][n];
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

if (typeof module !== "undefined" && module.exports) {
  module.exports = { computeStructuralSignature, tagColumn, tagSequenceSimilarity, tokenLevenshteinDistance };
}
