/*
 * cross-sheet-wrong-head-detector.js
 * ---------------------------------------
 * Invoice-level Wrong Head check ACROSS 2A and Books: matches the SAME
 * invoice (GSTIN + Invoice/Voucher Number) between the two sheets and
 * compares the tax head each side actually used for it.
 *
 * This is a separate, complementary check to wrong-head-detector.js,
 * which only validates a single sheet's own internal consistency (does a
 * 2A row's tax head match what its own supplier-state vs Place of Supply
 * requires). That check is correct for what it does and is unchanged by
 * this file — but it structurally cannot catch a real class of error: 2A
 * itself can be internally correct (charged the right head per its own
 * state data) while Books recorded the SAME invoice under the wrong
 * head. Books has no Place of Supply column at all (a purchase register
 * never carries one — see gst-role-recognizer.js, which treats that
 * absence as one of the signals that a sheet IS a purchase register), so
 * there's nothing within Books alone to validate against; the only way
 * to catch this is comparing the two sides for the same invoice.
 *
 * Matching is primarily EXACT: GSTIN + Invoice/Voucher Number, normalized
 * (trim + uppercase) the same way duplicate-invoice-detector.js does. If a
 * (GSTIN, identifier) key isn't unique on EITHER side (more than one row
 * shares it), that pairing is ambiguous and is skipped entirely rather
 * than guessing which row corresponds to which.
 *
 * A single, narrow FUZZY fallback exists for a real, recurring formatting
 * gap validated against real data: the same invoice is often written
 * differently on the two sides — e.g. 2A shows "GSTINV/163" while Books
 * shows the bare number 163, or 2A shows "5015/21-22" (financial-year
 * suffix) while Books shows bare 5015. See findFuzzyMatchKey() below for
 * exactly how conservative this is: it only ever bridges a PURE number on
 * one side to a matching digit run (>= 3 digits, to avoid coincidentally
 * matching a 2-digit date/month fragment) embedded in a more complex
 * identifier on the other side, and only when that produces exactly ONE
 * candidate — an ambiguous or absent case is never guessed.
 *
 * "Which side is wrong" is determined the same way wrong-head-detector.js
 * determines it — from 2A's OWN supplier-state-vs-Place-of-Supply data,
 * since only 2A carries Place of Supply. If that can't be resolved (or
 * both sides disagree with the expected head, in different ways), the
 * incorrect side is reported as undetermined rather than guessed.
 */

const { actualTaxHead } = require("./wrong-head-detector.js");
const { stateCodeFromGstin, extractStateCodeFromPlaceOfSupply } = require("./gst-state-codes.js");
const { parseDateValue } = require("./duplicate-invoice-detector.js");

function toNumber(value) {
  if (typeof value === "number") return value;
  if (value === "" || value === null || value === undefined) return 0;
  const parsed = parseFloat(String(value).replace(/[,$£€¥%()]/g, ""));
  return isNaN(parsed) ? 0 : parsed;
}

function normalizeKey(gstin, identifier) {
  return String(gstin).trim().toUpperCase() + "::" + String(identifier).trim().toUpperCase();
}

// Builds a Map from normalized (GSTIN, identifier) key -> array of row
// records sharing that key (usually length 1; length > 1 means that key
// is ambiguous on this sheet and callers must not pick one arbitrarily).
function buildInvoiceIndex(values, headerRowIndex, columns) {
  const identifierCol = columns.invoiceNumber !== null ? columns.invoiceNumber : columns.voucherNumber;
  if (columns.gstin === null || identifierCol === null) return null;

  const dataRows = headerRowIndex === -1 ? values : values.slice(headerRowIndex + 1);
  const index = new Map();
  dataRows.forEach((row, i) => {
    const gstinRaw = row[columns.gstin];
    const idRaw = row[identifierCol];
    if (!gstinRaw || String(gstinRaw).trim() === "") return;
    if (!idRaw || String(idRaw).trim() === "") return;

    const key = normalizeKey(gstinRaw, idRaw);
    if (!index.has(key)) index.set(key, []);
    const dateCol = columns.dateColumns && columns.dateColumns.length > 0 ? columns.dateColumns[0] : null;
    index.get(key).push({
      rowIndex: i,
      gstin: String(gstinRaw).trim().toUpperCase(),
      identifier: String(idRaw).trim().toUpperCase(),
      taxableValue: columns.taxableValue !== null ? toNumber(row[columns.taxableValue]) : 0,
      igst: columns.igst !== null ? toNumber(row[columns.igst]) : 0,
      cgst: columns.cgst !== null ? toNumber(row[columns.cgst]) : 0,
      sgst: columns.sgst !== null ? toNumber(row[columns.sgst]) : 0,
      placeOfSupply: columns.placeOfSupply !== null ? row[columns.placeOfSupply] : null,
      date: dateCol !== null ? parseDateValue(row[dateCol]) : null,
    });
  });
  return index;
}

const MIN_FUZZY_DIGIT_RUN_LENGTH = 3; // shorter runs (day/month/2-digit-year fragments) are too likely to coincidentally collide with an unrelated invoice number

function digitRunsOfMinLength(identifier, minLength) {
  const runs = identifier.match(/\d+/g) || [];
  return runs.filter((r) => r.length >= minLength).map((r) => String(parseInt(r, 10))); // leading zeros stripped so "007" and "7" are the same run
}

// GSTIN::digitValue -> array of index KEYS (not rows) whose identifier
// contains that digit run somewhere (at least MIN_FUZZY_DIGIT_RUN_LENGTH
// digits long) — used to find a complex identifier ("GSTINV/163") that
// embeds a plain number ("163") from the other sheet.
function buildDigitRunIndex(invoiceIndex) {
  const digitIndex = new Map();
  for (const [key, rows] of invoiceIndex) {
    const gstin = rows[0].gstin;
    const runs = new Set(digitRunsOfMinLength(rows[0].identifier, MIN_FUZZY_DIGIT_RUN_LENGTH));
    for (const run of runs) {
      const digitKey = gstin + "::" + run;
      if (!digitIndex.has(digitKey)) digitIndex.set(digitKey, []);
      digitIndex.get(digitKey).push(key);
    }
  }
  return digitIndex;
}

// GSTIN::digitValue -> index KEY, but ONLY for identifiers that are
// PURELY numeric (the whole identifier, not just a run within it) — the
// "163" side of a "GSTINV/163" vs "163" pair.
function buildPureNumericIndex(invoiceIndex) {
  const pureIndex = new Map();
  for (const [key, rows] of invoiceIndex) {
    const identifier = rows[0].identifier;
    if (!/^\d+$/.test(identifier)) continue;
    if (identifier.length < MIN_FUZZY_DIGIT_RUN_LENGTH) continue;
    const digitKey = rows[0].gstin + "::" + String(parseInt(identifier, 10));
    // A pure-numeric identifier should be unique per GSTIN on a well-formed
    // sheet; if two different keys somehow map to the same digit value
    // (e.g. "007" and "7" — already collapsed to the same exact key
    // upstream — or a genuine data anomaly), don't pick one arbitrarily.
    if (pureIndex.has(digitKey)) {
      pureIndex.set(digitKey, null); // mark ambiguous
    } else {
      pureIndex.set(digitKey, key);
    }
  }
  return pureIndex;
}

// Tries to bridge `identifier` (GSTIN `gstin`, on one sheet) to a SINGLE,
// unambiguous key on the OTHER sheet, when no exact match exists — see
// the module docstring for exactly what this does and doesn't attempt.
// ownDigitRunIndex: the digit-run index built from the SAME sheet
// `identifier` came from — needed for the complex-identifier branch below
// to verify RECIPROCAL uniqueness (see its comment).
// Returns null (never guesses) if identifier isn't purely numeric AND
// doesn't contain a plain-numeric counterpart on the other side, or if
// more than one candidate would satisfy the match in EITHER direction.
function findFuzzyMatchKey(gstin, identifier, ownDigitRunIndex, otherDigitRunIndex, otherPureNumericIndex) {
  if (/^\d+$/.test(identifier)) {
    if (identifier.length < MIN_FUZZY_DIGIT_RUN_LENGTH) return null;
    const digitKey = gstin + "::" + String(parseInt(identifier, 10));
    const candidates = otherDigitRunIndex.get(digitKey);
    if (!candidates || candidates.length !== 1) return null; // absent, or ambiguous (2+ complex identifiers embed this same number) -> don't guess
    return candidates[0];
  }

  // identifier is complex (not purely numeric) — look for a PURE-numeric
  // counterpart on the other side matching one of ITS digit runs.
  // otherPureNumericIndex can map a digit value to `null` (marked
  // ambiguous when building the index) — `if (candidate)` skips both that
  // and a genuinely absent lookup, so only a clean single-key hit ever
  // gets added.
  const runs = new Set(digitRunsOfMinLength(identifier, MIN_FUZZY_DIGIT_RUN_LENGTH));
  const matchedKeys = new Set();
  for (const run of runs) {
    const candidate = otherPureNumericIndex.get(gstin + "::" + run);
    if (!candidate) continue;
    // Reciprocal check: this digit value must ALSO be embedded by only
    // ONE identifier on THIS (own) side — otherwise two different
    // identifiers here (e.g. "GSTINV/500" and "OLDINV/500") could both
    // claim the SAME single counterpart on the other side, which is
    // exactly the kind of guess this fallback must not make.
    const ownCandidatesForRun = ownDigitRunIndex.get(gstin + "::" + run);
    if (!ownCandidatesForRun || ownCandidatesForRun.length !== 1) continue;
    matchedKeys.add(candidate);
  }
  if (matchedKeys.size !== 1) return null; // no digit run matched, or different runs matched different keys -> ambiguous, don't guess
  return [...matchedKeys][0];
}

// gstr2aInput/booksInput: { values, headerRowIndex, columns } — same
// identifyGstColumns() shape used throughout the GST pipeline.
function detectCrossSheetWrongHead(gstr2aInput, booksInput) {
  const gstr2aIndex = buildInvoiceIndex(gstr2aInput.values, gstr2aInput.headerRowIndex, gstr2aInput.columns);
  const booksIndex = buildInvoiceIndex(booksInput.values, booksInput.headerRowIndex, booksInput.columns);

  if (!gstr2aIndex || !booksIndex) {
    return { applicable: false, reason: "no GSTIN + Invoice/Voucher Number column identified on both sheets — nothing to match invoices on", flagged: [] };
  }

  const gstr2aDigitRunIndex = buildDigitRunIndex(gstr2aIndex);
  const booksDigitRunIndex = buildDigitRunIndex(booksIndex);
  const booksPureNumericIndex = buildPureNumericIndex(booksIndex);

  const flagged = [];
  let matchedInvoiceCount = 0;
  let ambiguousCount = 0;

  for (const [key, gstr2aRows] of gstr2aIndex) {
    let booksRows = booksIndex.get(key);
    if (!booksRows) {
      const fuzzyKey = findFuzzyMatchKey(gstr2aRows[0].gstin, gstr2aRows[0].identifier, gstr2aDigitRunIndex, booksDigitRunIndex, booksPureNumericIndex);
      booksRows = fuzzyKey ? booksIndex.get(fuzzyKey) : null;
    }
    if (!booksRows) continue; // not present in Books at all — Step 1's Extra-in-2A territory, not this rule's concern

    if (gstr2aRows.length !== 1 || booksRows.length !== 1) {
      ambiguousCount++; // same (GSTIN, identifier) appears more than once on one side — can't confidently pair, skip rather than guess
      continue;
    }
    matchedInvoiceCount++;

    const g = gstr2aRows[0];
    const b = booksRows[0];

    const gstr2aHead = actualTaxHead(g.igst, g.cgst, g.sgst);
    const booksHead = actualTaxHead(b.igst, b.cgst, b.sgst);

    // Only compare when BOTH sides show one clean, unambiguous head.
    // "mixed" (both IGST and CGST/SGST populated on one row) is the
    // existing single-sheet rule's own concern for whichever side it
    // occurs on; "none" means nothing was charged on that side at all,
    // which is a presence/amount question (Step 1's Mismatch territory),
    // not a wrong-HEAD question.
    if (gstr2aHead === "none" || booksHead === "none" || gstr2aHead === "mixed" || booksHead === "mixed") continue;
    if (gstr2aHead === booksHead) continue; // same head both sides — correct, nothing to flag

    const supplierStateCode = stateCodeFromGstin(g.gstin);
    const placeOfSupplyStateCode = g.placeOfSupply !== null ? extractStateCodeFromPlaceOfSupply(g.placeOfSupply) : null;
    let expectedHead = null;
    let incorrectSide = null;
    if (supplierStateCode && placeOfSupplyStateCode) {
      expectedHead = supplierStateCode === placeOfSupplyStateCode ? "cgst_sgst" : "igst";
      if (gstr2aHead === expectedHead && booksHead !== expectedHead) incorrectSide = "books";
      else if (booksHead === expectedHead && gstr2aHead !== expectedHead) incorrectSide = "gstr2a";
      // else: neither side matches the expected head, or Place of Supply
      // couldn't be resolved — incorrectSide stays null (undetermined),
      // reported honestly rather than guessed.
    }

    flagged.push({
      gstin: g.gstin,
      identifier: g.identifier,
      gstr2aRowIndex: g.rowIndex,
      booksRowIndex: b.rowIndex,
      gstr2aHead,
      booksHead,
      expectedHead,
      incorrectSide,
      gstr2a: { igst: g.igst, cgst: g.cgst, sgst: g.sgst },
      books: { igst: b.igst, cgst: b.cgst, sgst: b.sgst },
    });
  }

  return { applicable: true, flagged, matchedInvoiceCount, ambiguousCount };
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { detectCrossSheetWrongHead, buildInvoiceIndex, buildDigitRunIndex, buildPureNumericIndex, findFuzzyMatchKey };
}
