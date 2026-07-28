/*
 * classifier.js
 * -------------
 * MENTOR Phase 2, Step 2 — rule-based sheet classification.
 *
 * Takes the raw per-column signals from signal-extractor.js (Step 1) and
 * scores a sheet against four known shapes: bank statement, invoice list,
 * sales/POS data, payroll. Every rule below reads only structural signals
 * (column types, cardinality, sign, monotonicity) — nothing here is tied
 * to Three Sisters Kitchen's sheet names or column order.
 *
 * The one exception is a small header-vocabulary bonus (e.g. "invoice",
 * "wages"): sales/POS and payroll sheets are often structurally identical
 * (a repeating month column + several positive numeric columns), so common
 * business vocabulary is the only real differentiator between them. That
 * bonus is deliberately weighted lower than the structural rules so shape
 * still dominates the score.
 *
 * Output includes which rules fired (`matchedRules`) so a wrong call is
 * debuggable, not a black box — needed for the Step 3 regression check
 * and for trusting this on an unfamiliar client file later.
 */

const SHEET_TYPES = ["bank_statement", "invoice_list", "sales_pos", "payroll"];

const THRESHOLDS = {
  dominantTypeShare: 0.6, // a column "is" date/numeric/text once it clears this share of non-blank cells
  highCardinalityRatio: 0.5, // numeric column where most values are distinct — "amount-like", not a rate/flag
  highCardinalityTextRatio: 0.85, // text column that's essentially unique per row — ID-like
  categoricalMaxDistinctCount: 12,
  categoricalMaxRatio: 0.4,
  moderateCardinalityMinCount: 4, // vendor-name-variety band: more than a handful of fixed categories...
  moderateCardinalityMinRatio: 0.05,
  moderateCardinalityMaxRatio: 0.9, // ...but not unique per row either
  periodicMaxDistinctCount: 15, // a date column repeating a small set of period labels (e.g. "Jun 2026" x 3)
  periodicMaxRatio: 0.6,
  nonNettingMinPositiveRatio: 0.9, // "non-netting amounts" — almost nothing is negative
  monotonicMinSampleSize: 3,
  monotonicMinRatio: 0.9,
};

const DEBIT_CREDIT_VOCAB = ["debit", "credit", "dr", "cr", "paid in", "paid out", "withdrawal", "deposit", "in", "out", "money in", "money out"];

const KEYWORD_VOCAB = {
  bank_statement: ["balance", "debit", "credit", "statement", "transaction", "withdrawal", "deposit", "sort code", "account"],
  invoice_list: ["invoice", "vendor", "supplier", "due date", "po number", "purchase order", "bill", "net amount", "vat", "gross amount"],
  sales_pos: ["sales", "revenue", "pos", "dine-in", "dine in", "takeaway", "delivery", "covers", "gross sales"],
  payroll: ["wage", "wages", "salary", "salaries", "payroll", "overtime", "ni", "national insurance", "pension", "tronc", "gross pay", "net pay", "paye"],
};

const MIN_SCORE_TO_CLASSIFY = 3;
// A sheet can clear MIN_SCORE_TO_CLASSIFY yet still not clearly resemble any
// one type more than another (e.g. a Variance/summary sheet that picks up a
// few incidental points toward two different types at once). Requiring a
// minimum margin over the runner-up, not just an absolute score, is what
// lets those sheets fall through to "unknown" instead of being force-fit
// into whichever type happened to score highest.
const MIN_MARGIN_TO_CLASSIFY = 2;

// Internal type -> what the add-in actually shows the user. Kept as a
// separate layer so the classifier's own logic and matchedRules tracing can
// keep using stable technical names, while the UI-facing wording is free to
// change independently.
const DISPLAY_LABELS = {
  bank_statement: "Bank Statement",
  invoice_list: "Invoice List",
  sales_pos: "Sales / POS Data",
  payroll: "Payroll",
  unknown: "Not recognized yet",
};

function getDisplayLabel(type) {
  return DISPLAY_LABELS[type] || DISPLAY_LABELS.unknown;
}

// ---------------------------------------------------------------
// 1. COLUMN ROLE HELPERS — read Step 1's signals, decide nothing about
//    sheet type. Each of these just answers "what shape is this column?"
// ---------------------------------------------------------------
function isDateColumn(col) {
  return col.nonBlankCount > 0 && col.pctDate >= THRESHOLDS.dominantTypeShare;
}

function isNumericColumn(col) {
  return col.nonBlankCount > 0 && col.pctNumeric >= THRESHOLDS.dominantTypeShare;
}

function isTextColumn(col) {
  return col.nonBlankCount > 0 && col.pctText >= THRESHOLDS.dominantTypeShare;
}

function isPeriodicDateColumn(col) {
  return isDateColumn(col) && col.distinctCount <= THRESHOLDS.periodicMaxDistinctCount && col.distinctRatio <= THRESHOLDS.periodicMaxRatio;
}

function isHighCardinalityNumeric(col) {
  return isNumericColumn(col) && col.distinctRatio >= THRESHOLDS.highCardinalityRatio;
}

function isNonNettingNumeric(col) {
  return isNumericColumn(col) && col.signStats.pctPositive >= THRESHOLDS.nonNettingMinPositiveRatio;
}

function isModerateCardinalityText(col) {
  return (
    isTextColumn(col) &&
    col.distinctCount >= THRESHOLDS.moderateCardinalityMinCount &&
    col.distinctRatio > THRESHOLDS.moderateCardinalityMinRatio &&
    col.distinctRatio < THRESHOLDS.moderateCardinalityMaxRatio
  );
}

function isHighCardinalityText(col) {
  return isTextColumn(col) && col.distinctRatio >= THRESHOLDS.highCardinalityTextRatio;
}

function isCategoricalColumn(col) {
  return col.nonBlankCount > 0 && col.distinctCount <= THRESHOLDS.categoricalMaxDistinctCount && col.distinctRatio <= THRESHOLDS.categoricalMaxRatio;
}

function isNearMonotonic(col) {
  const m = col.monotonic;
  return m.numericSampleSize >= THRESHOLDS.monotonicMinSampleSize && (m.isMonotonicIncreasing || m.increasingRatio >= THRESHOLDS.monotonicMinRatio);
}

function looksLikeDebitCreditTag(col) {
  if (col.distinctCount !== 2) return false;
  const values = col.topValues.map((t) => String(t.value).toLowerCase().trim());
  return values.some((v) => DEBIT_CREDIT_VOCAB.includes(v));
}

function countKeywordMatches(sheetSignals, keywords) {
  const headers = sheetSignals.columns.map((c) => String(c.header || "").toLowerCase());
  return headers.filter((h) => keywords.some((k) => h.includes(k))).length;
}

// ---------------------------------------------------------------
// 2. PRECOMPUTE COLUMN GROUPS ONCE PER SHEET — rules below just read these.
// ---------------------------------------------------------------
function buildColumnGroups(sheetSignals) {
  const columns = sheetSignals.columns;
  return {
    columns,
    dateColumns: columns.filter(isDateColumn),
    periodicDateColumns: columns.filter(isPeriodicDateColumn),
    numericColumns: columns.filter(isNumericColumn),
    highCardinalityNumericColumns: columns.filter(isHighCardinalityNumeric),
    nonNettingNumericColumns: columns.filter(isNonNettingNumeric),
    moderateCardinalityTextColumns: columns.filter(isModerateCardinalityText),
    highCardinalityTextColumns: columns.filter(isHighCardinalityText),
    categoricalColumns: columns.filter(isCategoricalColumn),
  };
}

// ---------------------------------------------------------------
// 3. STRUCTURAL RULES PER SHEET TYPE
//    Each rule is a boolean test over the precomputed column groups, plus
//    a human-readable label used for the matchedRules trace. Weights can be
//    negative — a penalty rule for a signal that's incompatible with a type.
// ---------------------------------------------------------------

// A column mixing positive AND negative monetary values is essentially a
// bank/ledger-only signature — invoice, POS, and payroll amounts are each
// one-directional. Previously bank_statement's genuine signal (this same
// mixed-sign column) had to outscore incidental overlaps elsewhere (e.g. a
// bank sheet's own "Matched Invoice ID" column tripping invoice_list's
// ID-column rule), leaving only a 7-vs-6 margin. Penalizing the other three
// types by the same magnitude the sign rule rewards bank_statement widens
// that margin on the signal that actually matters, instead of just adding
// more arbitrary weight to bank_statement's side.
const MIXED_SIGN_PENALTY_RULE = {
  weight: -3,
  label: "has a signed (mixed positive/negative) numeric column — inconsistent with amounts that should run one direction",
  test: (s, g) => g.highCardinalityNumericColumns.some((col) => col.signStats.hasBothSigns),
};

const STRUCTURAL_RULES = {
  bank_statement: [
    { weight: 2, label: "has a date column", test: (s, g) => g.dateColumns.length > 0 },
    {
      // Weighted well above the other rules deliberately: a single column
      // legitimately mixing positive and negative monetary values is rare
      // outside transaction/ledger data — invoices, POS, and payroll sheets
      // essentially never have this. When present, it's the single strongest
      // fingerprint available, strong enough to outweigh incidental overlaps
      // (e.g. a bank sheet with a reconciliation "Matched Invoice ID" column).
      weight: 5,
      label: "has a signed amount column (both positive and negative values in one column — debits and credits together)",
      test: (s, g) => g.highCardinalityNumericColumns.some((col) => col.signStats.hasBothSigns),
    },
    {
      weight: 3,
      label: "has a near-monotonic high-cardinality numeric column (possible running balance)",
      test: (s, g) => g.highCardinalityNumericColumns.some(isNearMonotonic),
    },
    { weight: 3, label: "has a binary Debit/Credit-style tag column", test: (s, g) => g.columns.some(looksLikeDebitCreditTag) },
  ],
  invoice_list: [
    { weight: 2, label: "has a date column (invoice/due date)", test: (s, g) => g.dateColumns.length > 0 },
    {
      weight: 1,
      label: "a date column header mentions 'due'",
      test: (s, g) => g.dateColumns.some((col) => String(col.header || "").toLowerCase().includes("due")),
    },
    {
      weight: 3,
      label: "has a non-netting amount column (almost entirely positive values — gross invoice amounts, not signed transactions)",
      test: (s, g) => g.highCardinalityNumericColumns.some((col) => isNonNettingNumeric(col)),
    },
    {
      weight: 2,
      label: "has a moderate-cardinality text column (vendor-name variety — more than a couple fixed categories, not unique per row)",
      test: (s, g) => g.moderateCardinalityTextColumns.length > 0,
    },
    {
      weight: 1,
      label: "has a high-cardinality text column (possible per-row invoice ID)",
      test: (s, g) => g.highCardinalityTextColumns.length > 0,
    },
    MIXED_SIGN_PENALTY_RULE,
  ],
  sales_pos: [
    {
      weight: 2,
      label: "has a repeating period-label date column (e.g. a Month column repeating a handful of values)",
      test: (s, g) => g.periodicDateColumns.length > 0,
    },
    {
      weight: 3,
      label: "has several non-netting numeric columns side by side (multiple revenue channels)",
      test: (s, g) => g.highCardinalityNumericColumns.filter(isNonNettingNumeric).length >= 3,
    },
    { weight: 1, label: "has a low-cardinality categorical column alongside period labels (e.g. Site/Location)", test: (s, g) => g.periodicDateColumns.length > 0 && g.categoricalColumns.length > 0 },
    MIXED_SIGN_PENALTY_RULE,
  ],
  payroll: [
    {
      weight: 2,
      label: "has a repeating period-label date column (e.g. a Month column repeating a handful of values)",
      test: (s, g) => g.periodicDateColumns.length > 0,
    },
    {
      weight: 2,
      label: "has several non-netting numeric columns side by side (wage/deduction components)",
      test: (s, g) => g.highCardinalityNumericColumns.filter(isNonNettingNumeric).length >= 3,
    },
    MIXED_SIGN_PENALTY_RULE,
  ],
};

const KEYWORD_RULE_WEIGHT = {
  bank_statement: 1,
  invoice_list: 1,
  sales_pos: 1,
  // Payroll and sales/POS are structurally near-identical (period column +
  // several positive numeric columns) — vocabulary is the main thing that
  // actually tells them apart, so it counts for more here.
  payroll: 1.5,
};
const KEYWORD_MATCH_CAP = 4;

// All four types are inherently date/period-anchored records — a bank
// statement has transaction dates, an invoice list has invoice/due dates, a
// POS or payroll report has a period column. Without ANY date column at all,
// no single other signal should be able to carry a classification — this is
// what stops a numbers-only summary sheet (e.g. a Variance Analysis, whose
// "vs last period" columns are legitimately mixed-sign) from being force-fit
// into bank_statement purely on the mixed-sign rule.
function hasAnyDateColumn(columnGroups) {
  return columnGroups.dateColumns.length > 0;
}

// ---------------------------------------------------------------
// 4. SCORE ONE SHEET AGAINST ALL FOUR TYPES
// ---------------------------------------------------------------
function scoreSheetType(sheetSignals, columnGroups, type) {
  if (!hasAnyDateColumn(columnGroups)) {
    return { score: 0, matchedRules: [{ label: "no date column present — ruled out (every recognized sheet type is date/period-anchored)", weight: 0 }] };
  }

  const rules = STRUCTURAL_RULES[type];
  const matchedRules = [];
  let score = 0;

  for (const rule of rules) {
    if (rule.test(sheetSignals, columnGroups)) {
      score += rule.weight;
      matchedRules.push({ label: rule.label, weight: rule.weight });
    }
  }

  const keywordMatches = Math.min(countKeywordMatches(sheetSignals, KEYWORD_VOCAB[type]), KEYWORD_MATCH_CAP);
  if (keywordMatches > 0) {
    const keywordScore = keywordMatches * KEYWORD_RULE_WEIGHT[type];
    score += keywordScore;
    matchedRules.push({
      label: `${keywordMatches} column header(s) match ${type.replace("_", " ")} vocabulary`,
      weight: keywordScore,
    });
  }

  return { score, matchedRules };
}

function classifySheet(sheetSignals) {
  if (!sheetSignals || !sheetSignals.columns || sheetSignals.columns.length === 0) {
    return {
      sheetName: sheetSignals && sheetSignals.sheetName,
      type: "unknown",
      displayLabel: getDisplayLabel("unknown"),
      confidence: "none",
      scores: {},
      matchedRules: [],
      reason: "no columns to score",
    };
  }

  const columnGroups = buildColumnGroups(sheetSignals);

  const scored = SHEET_TYPES.map((type) => ({ type, ...scoreSheetType(sheetSignals, columnGroups, type) }));
  scored.sort((a, b) => b.score - a.score);

  const scores = {};
  scored.forEach((s) => (scores[s.type] = Math.round(s.score * 100) / 100));

  const best = scored[0];
  const runnerUp = scored[1];
  const margin = best.score - runnerUp.score;

  if (best.score < MIN_SCORE_TO_CLASSIFY || margin < MIN_MARGIN_TO_CLASSIFY) {
    const reason =
      best.score < MIN_SCORE_TO_CLASSIFY
        ? `no sheet type scored at least ${MIN_SCORE_TO_CLASSIFY} points (best was ${best.type} at ${scores[best.type]})`
        : `${best.type} (${scores[best.type]}) didn't clear ${runnerUp.type} (${scores[runnerUp.type]}) by the required margin of ${MIN_MARGIN_TO_CLASSIFY}`;

    return {
      sheetName: sheetSignals.sheetName,
      type: "unknown",
      displayLabel: getDisplayLabel("unknown"),
      confidence: "none",
      scores,
      matchedRules: [],
      reason,
    };
  }

  const confidence = margin >= 4 ? "high" : "medium";

  return {
    sheetName: sheetSignals.sheetName,
    type: best.type,
    displayLabel: getDisplayLabel(best.type),
    confidence,
    scores,
    matchedRules: best.matchedRules,
  };
}

function classifyWorkbook(sheetSignalsList) {
  return sheetSignalsList.map(classifySheet);
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    classifySheet,
    classifyWorkbook,
    getDisplayLabel,
    SHEET_TYPES,
    // Column-role helpers, exported so other modules (e.g. structural-signature.js)
    // can reuse the exact same shape thresholds instead of redefining them.
    isDateColumn,
    isPeriodicDateColumn,
    isNumericColumn,
    isHighCardinalityNumeric,
    isNonNettingNumeric,
    isModerateCardinalityText,
    isHighCardinalityText,
    isCategoricalColumn,
  };
}
