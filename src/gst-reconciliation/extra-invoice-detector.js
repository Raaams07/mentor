/*
 * extra-invoice-detector.js
 * ---------------------------------------
 * Invoice-level Extra-in-2A / Extra-in-Books detection: matches the SAME
 * invoice (GSTIN + Invoice/Voucher Number) between 2A and Books, reusing
 * the invoice index cross-sheet-wrong-head-detector.js already builds and
 * validates, and reports any invoice present on ONE side with no
 * corresponding entry on the other — regardless of what Step 1's
 * GSTIN-aggregate status says for that vendor.
 *
 * Why this exists as a SEPARATE check from Step 1's vendor-level
 * missing_in_books/missing_in_gstr2a status (gstr2a-vs-books-reconciler.js,
 * unchanged and still the source of the Step 1 "Vendor match status"
 * counts): that check aggregates ALL of a vendor's rows into one GSTIN
 * total before comparing presence. A vendor with 5 invoices in Books where
 * even ONE also appears in 2A is "present" on both sides at the GSTIN
 * level — Step 1 calls that vendor matched/discrepancy, never missing —
 * even though the other 4 invoices from that same vendor may never appear
 * in 2A at all. Validated against a real manual reconciliation (Basai
 * Steel Traders): this exact pattern hid a real ~₹5,04,000.90 CGST +
 * ₹5,04,000.90 SGST of genuine Extra-in-Books invoices behind vendors Step
 * 1 had already marked matched/discrepancy.
 *
 * Matching philosophy differs slightly from the Wrong Head cross-sheet
 * check: that check compares VALUES for a specific pairing, so an
 * ambiguous (GSTIN, identifier) key — more than one row sharing it on
 * either side — is skipped entirely rather than guessing which row pairs
 * with which. This check only asks EXISTENCE ("does this invoice number
 * appear anywhere on the other side at all"), not "which specific row
 * corresponds to which" — so for an UNAMBIGUOUS key (exactly one row on
 * each side), the key existing on both sides is enough; only a key
 * completely ABSENT from the other side's index is reported as Extra.
 *
 * AMBIGUOUS-KEY AMOUNT CORROBORATION: when a key isn't unique on at
 * least one side (2+ rows share it), existence alone is no longer
 * trustworthy enough — confirmed against real data: Books had the SAME
 * "Supplier Invoice No." accidentally typed on two different vouchers,
 * one a genuine match to 2A, the other actually a different invoice
 * entirely (matching a DIFFERENT 2A invoice by amount, with a completely
 * different number). Treating the whole ambiguous group as "present"
 * once ANY row in it matched would have silently absorbed that second
 * voucher, hiding it from ever reaching the Possible-Match fallback
 * below. So within an ambiguous group specifically, a row only counts as
 * matched if it has a real amount+tax counterpart (±₹1) somewhere in the
 * SAME key's group; a row without one is treated as a normal candidate
 * for every fallback tier below, same as if the key had been absent
 * entirely. This does NOT change the simple, common case (exactly one
 * row on each side) — that still matches on identifier alone.
 *
 * Exact matches are also backed up by the SAME narrow fuzzy fallback used
 * by cross-sheet-wrong-head-detector.js (see findFuzzyMatchKey() there) —
 * bridging a bare number on one side ("163") to a matching digit run
 * embedded in a more complex identifier on the other ("GSTINV/163"),
 * only when that's unambiguous. Validated against real data: on one real
 * file, roughly a third of what LOOKED like genuine Extra-in-2A rows
 * (and the large majority of their rupee value) were actually this exact
 * formatting gap, not real presence gaps.
 *
 * RCM rows are excluded from the OUTPUT (options.gstr2aExcludedRowIndices /
 * booksExcludedRowIndices — 0-based data-row indices, the SAME convention
 * detectRcmForSheet()'s own `rowIndex` uses). Confirmed against real data:
 * an invoice self-assessed and correctly reported on 'GST - RCM' has no
 * reason to also independently exist in Books under the ordinary vendor-
 * matched ITC flow, so with no RCM-aware exclusion it was ALSO showing up
 * as a false "Extra in 2A" — the same invoice flagged twice under two
 * different, contradictory theories.
 *
 * Deliberately filtered at OUTPUT time, per individual row — not by
 * removing RCM rows from the matching indices up front. An RCM-flagged
 * row can still legitimately be the thing that makes some OTHER row on
 * the opposite sheet NOT extra (it's still a real invoice that exists);
 * removing it from the index entirely would make that counterpart wrongly
 * flagged as Extra too. Filtering only "would this SPECIFIC row have been
 * reported" keeps that other row's match intact while still keeping the
 * RCM row itself out of the Extra lists.
 *
 * SHORT-IDENTIFIER TIEBREAK (runs after exact + digit-run-fuzzy matching,
 * over whatever's still left in extraIn2A/extraInBooks): the digit-run
 * fuzzy fallback above deliberately refuses anything under
 * MIN_FUZZY_DIGIT_RUN_LENGTH (3) digits — too easy to coincidentally
 * collide with a date/month fragment with no other evidence. But a real,
 * confirmed pattern needs exactly that: 2A stores a short running number
 * zero-padded ("068"), Books stores the SAME invoice as a bare number
 * without the padding ("68") — same value, different string, and never
 * even reaches the digit-run index at all since "68"/"068" are only 2-3
 * characters long.
 *
 * This tiebreak is safe to apply at ANY identifier length specifically
 * BECAUSE it always requires independent corroboration before confirming
 * a pair — unlike the digit-run fallback (which trusts the digits alone),
 * this only pairs a still-extra 2A row with a still-extra Books row when
 * they share the same GSTIN + leading-zero-normalized numeric identifier
 * AND their taxable value matches (±₹1) AND their dates are the same or
 * within a few days. If more than one candidate on either side would
 * satisfy that — genuine repetition of the same short number for the same
 * vendor — no pairing is guessed; both stay reported as Extra rather than
 * picking one arbitrarily.
 *
 * INVOICE-NUMBER NORMALIZATION (runs after the short-identifier tiebreak,
 * over whatever's still extra): a real, recurring class of formatting
 * noise between 2A and Books — year written as "21-22" vs "2021-22", a
 * truncated suffix, a dropped trailing digit, "-" vs "/" as the
 * separator — confirmed against real data (see invoice-number-
 * normalizer.js, which does the actual string-level recognition). A
 * string-level match there is NECESSARY but never SUFFICIENT: this tier
 * additionally requires taxable value AND every tax field (IGST/CGST/
 * SGST) to agree within ₹1 before accepting a pair, and — like the
 * short-identifier tiebreak above — only ever resolves a pair that's
 * mutually unique; genuine ambiguity is left alone. This is the reverse
 * risk of the Duplicate Invoices same_amount false positive (amount
 * matching without invoice-number confirmation) — here the risk runs the
 * other way (invoice-string similarity without amount confirmation), so
 * the same "require both signals" principle applies.
 *
 * POSSIBLE MATCH — needs review (runs LAST, only on what survives every
 * tier above): some real rows have NO usable invoice number to match on
 * at all — a Books journal voucher with a blank Supplier Invoice No.
 * field, say — or use a completely different reference scheme than 2A
 * (2A invoice "2018" vs Books invoice "1997"/voucher 403 for the SAME
 * purchase), where amounts still agree almost exactly. There is no
 * string-level signal here at all, so this is deliberately NOT folded
 * into extraIn2A/extraInBooks as a confirmed match OR left as a silent
 * gap — it's reported as its OWN third category (possibleMatches),
 * requiring GSTIN + taxable value + every tax field to agree (±₹1) AND
 * the two dates to fall in the same calendar month, still never guessing
 * across 2+ ambiguous candidates the same way every tier above doesn't.
 * Running this LAST is deliberate: it must only catch what tiers 1-3
 * couldn't resolve on the strength of the invoice number itself, never
 * cases those tiers should have handled more precisely. A row with a
 * BLANK identifier never even reaches buildInvoiceIndex (which skips
 * blank identifiers entirely — see its own docstring), so this tier
 * pulls those rows in separately via collectRowsWithGstin(), specifically
 * so they get a chance at this fallback despite never being indexed.
 *
 * Rows that land in possibleMatches are removed from extraIn2A/
 * extraInBooks (per the fix's own instruction: NOT netted into either
 * total automatically) — the caller decides how to surface them (see
 * gst-report-writer.js's dedicated 'GST - Possible Matches' sheet).
 */

const { buildInvoiceIndex, buildDigitRunIndex, buildPureNumericIndex, findFuzzyMatchKey } = require("./cross-sheet-wrong-head-detector.js");
const { invoiceNumbersMatch } = require("./invoice-number-normalizer.js");
const { parseDateValue } = require("./duplicate-invoice-detector.js");

const TIEBREAK_AMOUNT_TOLERANCE = 1; // ₹1 — rounding/paise-level noise, consistent with the tolerance used elsewhere in this pipeline
const TIEBREAK_DATE_TOLERANCE_DAYS = 3; // "same/near date" — a small window for booking-date vs invoice-date drift, not a proximity match on its own

function normalizeNumericIdentifier(identifier) {
  return /^\d+$/.test(identifier) ? String(parseInt(identifier, 10)) : null; // leading zeros stripped; non-numeric identifiers don't participate in this tiebreak at all
}

function daysBetween(d1, d2) {
  return Math.abs(d1.getTime() - d2.getTime()) / 86400000;
}

// Both signals required — an amount match alone, or a date match alone,
// isn't "confident" on its own for a bare 2-3 digit identifier.
function amountAndDateCorroborate(a, b) {
  if (Math.abs(a.taxableValue - b.taxableValue) > TIEBREAK_AMOUNT_TOLERANCE) return false;
  if (!a.date || !b.date) return false; // can't confirm without a real date on both sides — don't guess
  return daysBetween(a.date, b.date) <= TIEBREAK_DATE_TOLERANCE_DAYS;
}

// Taxable value AND every tax field must agree (±₹1) — a stricter bar
// than the short-identifier tiebreak's amount+date, matching the Fix 3
// guardrail's explicit wording ("taxable value AND tax amounts").
function amountAndTaxCorroborate(a, b) {
  return (
    Math.abs(a.taxableValue - b.taxableValue) <= TIEBREAK_AMOUNT_TOLERANCE &&
    Math.abs(a.igst - b.igst) <= TIEBREAK_AMOUNT_TOLERANCE &&
    Math.abs(a.cgst - b.cgst) <= TIEBREAK_AMOUNT_TOLERANCE &&
    Math.abs(a.sgst - b.sgst) <= TIEBREAK_AMOUNT_TOLERANCE
  );
}

// Generic bipartite "mutually unique" pairing: among candidatesA and
// candidatesB, pairs (a, b) only when b is a's ONLY corroborating
// candidate AND a is b's ONLY corroborating candidate — i.e. never
// guesses when 2+ candidates on either side remain indistinguishable.
// isMatch(a, b): the caller's own corroboration predicate. `pairs`
// carries the actual (a, b) pairing, for callers that need to record
// BOTH sides together (not just which rows got resolved).
function resolveMutuallyUniquePairs(candidatesA, candidatesB, isMatch) {
  const matchesA = new Map();
  const matchesB = new Map();
  for (const a of candidatesA) {
    for (const b of candidatesB) {
      if (!isMatch(a, b)) continue;
      if (!matchesA.has(a)) matchesA.set(a, []);
      matchesA.get(a).push(b);
      if (!matchesB.has(b)) matchesB.set(b, []);
      matchesB.get(b).push(a);
    }
  }

  const resolvedA = new Set();
  const resolvedB = new Set();
  const pairs = [];
  for (const [a, matches] of matchesA) {
    if (matches.length !== 1) continue; // ambiguous on this side -- don't guess
    const b = matches[0];
    const backMatches = matchesB.get(b);
    if (!backMatches || backMatches.length !== 1) continue; // that candidate also matches some OTHER row -- ambiguous, don't guess
    resolvedA.add(a);
    resolvedB.add(b);
    pairs.push({ a, b });
  }
  return { resolvedA, resolvedB, pairs };
}

// Groups already-extra rows by GSTIN + normalized numeric identifier —
// the candidate pool the short-identifier tiebreak operates on.
function groupByNormalizedIdentifier(rows) {
  const groups = new Map();
  for (const r of rows) {
    const norm = normalizeNumericIdentifier(r.identifier);
    if (norm === null) continue;
    const groupKey = r.gstin + "::" + norm;
    if (!groups.has(groupKey)) groups.set(groupKey, []);
    groups.get(groupKey).push(r);
  }
  return groups;
}

// Applies the short-identifier tiebreak across every GSTIN+identifier
// group present in BOTH still-extra lists, returning the rows that should
// be REMOVED from extraIn2A/extraInBooks (i.e. successfully paired).
function resolveShortIdentifierTiebreaksAcrossGroups(extraIn2A, extraInBooks) {
  const gstr2aGroups = groupByNormalizedIdentifier(extraIn2A);
  const booksGroups = groupByNormalizedIdentifier(extraInBooks);

  const resolvedGstr2aRows = new Set();
  const resolvedBooksRows = new Set();
  for (const [groupKey, gstr2aCandidates] of gstr2aGroups) {
    const booksCandidates = booksGroups.get(groupKey);
    if (!booksCandidates) continue;
    const { resolvedA, resolvedB } = resolveMutuallyUniquePairs(gstr2aCandidates, booksCandidates, amountAndDateCorroborate);
    for (const r of resolvedA) resolvedGstr2aRows.add(r);
    for (const r of resolvedB) resolvedBooksRows.add(r);
  }
  return { resolvedGstr2aRows, resolvedBooksRows };
}

// Groups already-extra rows by GSTIN alone — invoiceNumbersMatch() isn't
// an equality-based key the way normalizeNumericIdentifier() is (it
// involves prefix/truncation logic), so candidates within one GSTIN are
// compared pairwise rather than grouped by a precomputed key.
function groupByGstin(rows) {
  const groups = new Map();
  for (const r of rows) {
    if (!groups.has(r.gstin)) groups.set(r.gstin, []);
    groups.get(r.gstin).push(r);
  }
  return groups;
}

// Applies invoice-number normalization + amount/tax corroboration across
// every GSTIN present in BOTH still-extra lists, returning the rows that
// should be REMOVED (i.e. successfully paired). See the module docstring's
// "INVOICE-NUMBER NORMALIZATION" paragraph for the two-signal requirement.
function resolveInvoiceNumberNormalizationAcrossGroups(extraIn2A, extraInBooks) {
  const gstr2aGroups = groupByGstin(extraIn2A);
  const booksGroups = groupByGstin(extraInBooks);

  const resolvedGstr2aRows = new Set();
  const resolvedBooksRows = new Set();
  for (const [gstin, gstr2aCandidates] of gstr2aGroups) {
    const booksCandidates = booksGroups.get(gstin);
    if (!booksCandidates) continue;
    const isMatch = (g, b) => invoiceNumbersMatch(g.identifier, b.identifier) && amountAndTaxCorroborate(g, b);
    const { resolvedA, resolvedB } = resolveMutuallyUniquePairs(gstr2aCandidates, booksCandidates, isMatch);
    for (const r of resolvedA) resolvedGstr2aRows.add(r);
    for (const r of resolvedB) resolvedBooksRows.add(r);
  }
  return { resolvedGstr2aRows, resolvedBooksRows };
}

function toNumber(value) {
  if (typeof value === "number") return value;
  if (value === "" || value === null || value === undefined) return 0;
  const parsed = parseFloat(String(value).replace(/[,$£€¥%()]/g, ""));
  return isNaN(parsed) ? 0 : parsed;
}

// Every row with a valid GSTIN, regardless of whether it has a usable
// identifier — unlike buildInvoiceIndex (which skips blank identifiers
// entirely), this tier specifically needs those rows as candidates. See
// the module docstring's "POSSIBLE MATCH" paragraph.
function collectRowsWithGstin(values, headerRowIndex, columns) {
  const identifierCol = columns.invoiceNumber !== null ? columns.invoiceNumber : columns.voucherNumber;
  const dateCol = columns.dateColumns && columns.dateColumns.length > 0 ? columns.dateColumns[0] : null;
  const dataRows = headerRowIndex === -1 ? values : values.slice(headerRowIndex + 1);
  const rows = [];
  dataRows.forEach((row, i) => {
    const gstinRaw = row[columns.gstin];
    if (!gstinRaw || String(gstinRaw).trim() === "") return;
    rows.push({
      rowIndex: i,
      gstin: String(gstinRaw).trim().toUpperCase(),
      identifier: identifierCol !== null ? String(row[identifierCol] || "").trim().toUpperCase() : "",
      taxableValue: columns.taxableValue !== null ? toNumber(row[columns.taxableValue]) : 0,
      igst: columns.igst !== null ? toNumber(row[columns.igst]) : 0,
      cgst: columns.cgst !== null ? toNumber(row[columns.cgst]) : 0,
      sgst: columns.sgst !== null ? toNumber(row[columns.sgst]) : 0,
      date: dateCol !== null ? parseDateValue(row[dateCol]) : null,
    });
  });
  return rows;
}

function sameCalendarMonth(d1, d2) {
  if (!d1 || !d2) return false; // can't confirm without a real date on both sides — don't guess
  return d1.getUTCFullYear() === d2.getUTCFullYear() && d1.getUTCMonth() === d2.getUTCMonth();
}

function possibleMatchCorroborate(a, b) {
  return amountAndTaxCorroborate(a, b) && sameCalendarMonth(a.date, b.date);
}

function possibleMatchReason(g, b) {
  if (!g.identifier && !b.identifier) return "Invoice/voucher number blank on both sides";
  if (!b.identifier) return "Books invoice/voucher number is blank";
  if (!g.identifier) return "2A invoice number is blank";
  return "Invoice numbers don't match ('" + g.identifier + "' vs '" + b.identifier + "')";
}

// gstr2aCandidates/booksCandidates: rows still extra after every prior
// tier, PLUS blank-identifier rows collected separately (see
// collectRowsWithGstin). Returns the rows to remove from extraIn2A/
// extraInBooks, plus the actual pairs (both sides' data + a human-
// readable reason) for the caller to report as its own category.
function resolvePossibleMatchesAcrossGroups(gstr2aCandidates, booksCandidates) {
  const gstr2aGroups = groupByGstin(gstr2aCandidates);
  const booksGroups = groupByGstin(booksCandidates);

  const resolvedGstr2aRows = new Set();
  const resolvedBooksRows = new Set();
  const possibleMatches = [];
  for (const [gstin, gCandidates] of gstr2aGroups) {
    const bCandidates = booksGroups.get(gstin);
    if (!bCandidates) continue;
    const { resolvedA, resolvedB, pairs } = resolveMutuallyUniquePairs(gCandidates, bCandidates, possibleMatchCorroborate);
    for (const r of resolvedA) resolvedGstr2aRows.add(r);
    for (const r of resolvedB) resolvedBooksRows.add(r);
    for (const { a, b } of pairs) {
      possibleMatches.push({
        gstin,
        gstr2aIdentifier: a.identifier,
        gstr2aRowIndex: a.rowIndex,
        gstr2a: { taxableValue: a.taxableValue, igst: a.igst, cgst: a.cgst, sgst: a.sgst },
        booksIdentifier: b.identifier,
        booksRowIndex: b.rowIndex,
        books: { taxableValue: b.taxableValue, igst: b.igst, cgst: b.cgst, sgst: b.sgst },
        reason: possibleMatchReason(a, b),
      });
    }
  }
  possibleMatches.sort((x, y) => x.gstin.localeCompare(y.gstin) || x.gstr2aRowIndex - y.gstr2aRowIndex);
  return { resolvedGstr2aRows, resolvedBooksRows, possibleMatches };
}

// gstr2aInput/booksInput: { values, headerRowIndex, columns } — same
// identifyGstColumns() shape used throughout the GST pipeline.
// options.gstr2aExcludedRowIndices / options.booksExcludedRowIndices:
// optional Set<number> of 0-based data-row indices (per sheet) already
// classified as RCM this run — kept OUT of the Extra-in-2A/Extra-in-Books
// output (see the module docstring's RCM paragraph) without affecting
// their ability to serve as a match target for the other sheet.
function detectInvoiceLevelExtras(gstr2aInput, booksInput, options) {
  const opts = options || {};
  const gstr2aExcluded = opts.gstr2aExcludedRowIndices || new Set();
  const booksExcluded = opts.booksExcludedRowIndices || new Set();

  const gstr2aIndex = buildInvoiceIndex(gstr2aInput.values, gstr2aInput.headerRowIndex, gstr2aInput.columns);
  const booksIndex = buildInvoiceIndex(booksInput.values, booksInput.headerRowIndex, booksInput.columns);

  if (!gstr2aIndex || !booksIndex) {
    return { applicable: false, reason: "no GSTIN + Invoice/Voucher Number column identified on both sheets — nothing to match invoices on", extraIn2A: [], extraInBooks: [], possibleMatches: [] };
  }

  const gstr2aDigitRunIndex = buildDigitRunIndex(gstr2aIndex);
  const gstr2aPureNumericIndex = buildPureNumericIndex(gstr2aIndex);
  const booksDigitRunIndex = buildDigitRunIndex(booksIndex);
  const booksPureNumericIndex = buildPureNumericIndex(booksIndex);

  const extraIn2A = [];
  for (const [key, rows] of gstr2aIndex) {
    const counterpartRows = booksIndex.get(key);
    if (!counterpartRows) {
      if (findFuzzyMatchKey(rows[0].gstin, rows[0].identifier, gstr2aDigitRunIndex, booksDigitRunIndex, booksPureNumericIndex)) continue; // e.g. 2A "GSTINV/163" bridges to Books "163" -> not extra
      for (const r of rows) {
        if (gstr2aExcluded.has(r.rowIndex)) continue; // already correctly classified as RCM this run — not also Extra in 2A
        extraIn2A.push(r);
      }
      continue;
    }
    if (rows.length === 1 && counterpartRows.length === 1) continue; // unambiguous exact match -> not extra, no change from existing behavior
    // Ambiguous on at least one side (see AMBIGUOUS-KEY AMOUNT CORROBORATION
    // in the module docstring) — a row only counts as matched if it has a
    // real amount+tax counterpart WITHIN this specific key's group; a
    // surplus row sharing the identifier by coincidence/data-entry error
    // stays a candidate for the later fallback tiers instead of being
    // silently absorbed as "present".
    for (const r of rows) {
      if (gstr2aExcluded.has(r.rowIndex)) continue;
      if (counterpartRows.some((b) => amountAndTaxCorroborate(r, b))) continue;
      extraIn2A.push(r);
    }
  }
  extraIn2A.sort((a, b) => a.gstin.localeCompare(b.gstin) || a.rowIndex - b.rowIndex);

  const extraInBooks = [];
  for (const [key, rows] of booksIndex) {
    const counterpartRows = gstr2aIndex.get(key);
    if (!counterpartRows) {
      if (findFuzzyMatchKey(rows[0].gstin, rows[0].identifier, booksDigitRunIndex, gstr2aDigitRunIndex, gstr2aPureNumericIndex)) continue;
      for (const r of rows) {
        if (booksExcluded.has(r.rowIndex)) continue;
        extraInBooks.push(r);
      }
      continue;
    }
    if (rows.length === 1 && counterpartRows.length === 1) continue;
    for (const r of rows) {
      if (booksExcluded.has(r.rowIndex)) continue;
      if (counterpartRows.some((g) => amountAndTaxCorroborate(r, g))) continue;
      extraInBooks.push(r);
    }
  }
  extraInBooks.sort((a, b) => a.gstin.localeCompare(b.gstin) || a.rowIndex - b.rowIndex);

  // Short-identifier tiebreak (see module docstring) — resolves whatever
  // it confidently can from what's STILL extra after exact + digit-run
  // fuzzy matching above; anything it can't confirm stays reported as-is.
  const shortIdTiebreak = resolveShortIdentifierTiebreaksAcrossGroups(extraIn2A, extraInBooks);
  const afterShortIdTiebreakIn2A = extraIn2A.filter((r) => !shortIdTiebreak.resolvedGstr2aRows.has(r));
  const afterShortIdTiebreakInBooks = extraInBooks.filter((r) => !shortIdTiebreak.resolvedBooksRows.has(r));

  // Invoice-number normalization (see module docstring) — runs after the
  // short-identifier tiebreak, over whatever's still extra.
  const invoiceNumberNormalization = resolveInvoiceNumberNormalizationAcrossGroups(afterShortIdTiebreakIn2A, afterShortIdTiebreakInBooks);
  const finalExtraIn2A = afterShortIdTiebreakIn2A.filter((r) => !invoiceNumberNormalization.resolvedGstr2aRows.has(r));
  const finalExtraInBooks = afterShortIdTiebreakInBooks.filter((r) => !invoiceNumberNormalization.resolvedBooksRows.has(r));

  // Possible match — needs review (see module docstring). Runs LAST, only
  // over what survives every tier above, PLUS rows with a blank
  // identifier (never indexed by buildInvoiceIndex in the first place, so
  // they'd otherwise never get a chance at any fallback at all).
  const blankGstr2aRows = collectRowsWithGstin(gstr2aInput.values, gstr2aInput.headerRowIndex, gstr2aInput.columns).filter(
    (r) => r.identifier === "" && !gstr2aExcluded.has(r.rowIndex)
  );
  const blankBooksRows = collectRowsWithGstin(booksInput.values, booksInput.headerRowIndex, booksInput.columns).filter(
    (r) => r.identifier === "" && !booksExcluded.has(r.rowIndex)
  );
  const possibleMatchResult = resolvePossibleMatchesAcrossGroups(finalExtraIn2A.concat(blankGstr2aRows), finalExtraInBooks.concat(blankBooksRows));
  // Blank-identifier rows were never part of finalExtraIn2A/finalExtraInBooks
  // to begin with, so filtering those arrays here only ever removes a row
  // that DID have an identifier and got resolved by this tier — an
  // unmatched blank-identifier row stays exactly as unreported as it was
  // before this fix (deliberately out of scope; see module docstring).
  const trulyFinalExtraIn2A = finalExtraIn2A.filter((r) => !possibleMatchResult.resolvedGstr2aRows.has(r));
  const trulyFinalExtraInBooks = finalExtraInBooks.filter((r) => !possibleMatchResult.resolvedBooksRows.has(r));

  return {
    applicable: true,
    extraIn2A: trulyFinalExtraIn2A,
    extraInBooks: trulyFinalExtraInBooks,
    possibleMatches: possibleMatchResult.possibleMatches,
    gstr2aInvoiceKeyCount: gstr2aIndex.size,
    booksInvoiceKeyCount: booksIndex.size,
  };
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { detectInvoiceLevelExtras, normalizeNumericIdentifier, TIEBREAK_AMOUNT_TOLERANCE, TIEBREAK_DATE_TOLERANCE_DAYS };
}
