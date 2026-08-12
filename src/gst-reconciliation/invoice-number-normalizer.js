/*
 * invoice-number-normalizer.js
 * ---------------------------------------
 * Recognizes when two invoice/voucher number STRINGS plausibly refer to
 * the same invoice, despite differing in ways that are common real-world
 * formatting noise rather than a genuinely different invoice. Validated
 * against a real file (Basai Steel Traders) that showed several distinct
 * patterns for the SAME invoice between 2A and Books:
 *
 *   - Year format:        "133/21-22"      vs "133/2021-22"
 *   - Truncated suffix:    "S-008"         vs "S-008/2021-22"
 *   - Dropped trailing digit: "5093/21-2"  vs "5093/21-22"
 *   - Punctuation variants: "BLY/21-22-280" vs "BLY/21-22/280",
 *                            "STPL/1181"    vs "STPL//1181"
 *   - Year VALUE differing (typo, not just format): "5264/21-21" vs
 *     "5264/21-22", "5024/2020-21" vs "5024/21-22" — confirmed against
 *     real data as the SAME invoice (identical taxable value and tax to
 *     the rupee) despite the year segment itself being wrong on one side.
 *
 * TWO deliberately separate mechanisms, in order:
 *
 *   1. normalizeInvoiceNumber() — a literal format normalization: strip
 *      whitespace, uppercase, unify "-"/"/" (and runs of them) into one
 *      separator, and expand a genuinely CONSECUTIVE 2-digit financial-
 *      year pair ("21-22") to 4 digits ("2021-22") so it lines up with an
 *      already-4-digit form. Two identifiers are the same invoice if this
 *      produces an EXACT match.
 *
 *   2. invoiceNumbersMatch() — a fallback for when normalization alone
 *      isn't enough (steps above still leave a difference): extracts each
 *      identifier's "core" segments by removing any FY-YEAR-SHAPED
 *      segment(s) — recognized by SHAPE, not by value, so a wrong or
 *      truncated year doesn't block a match — then requires the
 *      remaining core segments to line up: everything but the last must
 *      match EXACTLY, and the last may be a plain string prefix of the
 *      other (handles a dropped trailing digit, or one side simply
 *      lacking the year suffix at all).
 *
 * Deliberately does NOT do: fixing an arbitrary digit typo/transposition
 * in the SERIAL portion of an identifier ("944" vs "994", "8071" vs
 * "7081") via edit-distance or similar — confirmed against real data as
 * a genuinely different, much riskier class of change than a year/
 * separator/truncation difference. A single-digit difference in the part
 * of the identifier that actually DISTINGUISHES one invoice from another
 * is exactly the kind of thing this module must never paper over — see
 * invoice-number-normalizer-test.js for confirmed real cases this
 * correctly leaves unmatched.
 *
 * CRITICAL: neither function here is sufficient on its own to treat two
 * rows as the same invoice. A string-level match (either mechanism) is
 * NECESSARY but not SUFFICIENT — the caller (extra-invoice-detector.js)
 * must ALSO require taxable value and tax amounts to agree (within ₹1)
 * before accepting a pair. This mirrors the lesson from the Duplicate
 * Invoices false-positive fix (same_amount matching without invoice-
 * number confirmation caused false merges there) — here the risk runs
 * the other direction (invoice-string similarity without amount
 * confirmation), so the same "require both signals" principle applies.
 */

// A "year-shape" segment: recognized by SHAPE, not value, so a wrong or
// typo'd year doesn't block a match. A 2-digit segment is always
// eligible (this is the common short-FY-year form, "21"). A 4-digit
// segment is eligible ONLY within a realistic calendar-year range
// (2000-2099) — WITHOUT that restriction, a 4-digit SERIAL NUMBER like
// "7081" would be misread as a "year", stripping the very digits that
// distinguish one invoice from another and defeating the digit-typo
// guardrail above (confirmed: this was an actual bug caught while
// building this file, via "7081"/"8071" both incorrectly reducing to an
// empty-ish core before the range restriction was added).
const YEAR_4DIGIT_MIN = 2000;
const YEAR_4DIGIT_MAX = 2099;

function isYearShapeSegment(segment) {
  if (/^\d{2}$/.test(segment)) return true;
  if (/^\d{4}$/.test(segment)) {
    const n = parseInt(segment, 10);
    return n >= YEAR_4DIGIT_MIN && n <= YEAR_4DIGIT_MAX;
  }
  return false;
}

// Steps 1-2 of the spec: whitespace/case, then collapse any run of "-"
// or "/" characters into a single "-" — handles both a plain separator
// swap ("BLY/21-22-280" vs "BLY/21-22/280") and a doubled-separator typo
// ("STPL//1181" vs "STPL/1181") the same way.
function canonicalizeSeparators(identifier) {
  return String(identifier).trim().toUpperCase().replace(/[-/]+/g, "-");
}

function splitSegments(identifier) {
  return canonicalizeSeparators(identifier)
    .split("-")
    .filter((s) => s.length > 0);
}

// Step 3 of the spec: expand a GENUINELY CONSECUTIVE 2-digit FY pair
// ("21","22") to 4 digits ("2021","22") — deliberately narrower than
// isYearShapeSegment above (arithmetic consecutiveness required, not
// just shape), since this function's job is literal FORMAT
// normalization; tolerating a WRONG year value is invoiceNumbersMatch's
// separate, more lenient job below.
function normalizeInvoiceNumber(identifier) {
  const segments = splitSegments(identifier);
  const expanded = segments.map((seg, idx) => {
    const next = segments[idx + 1];
    if (/^\d{2}$/.test(seg) && next !== undefined && /^\d{2}$/.test(next)) {
      const yy = parseInt(seg, 10);
      const nextYy = parseInt(next, 10);
      if (nextYy === (yy + 1) % 100) return "20" + seg;
    }
    return seg;
  });
  return expanded.join("-");
}

// Removes every recognized year-shape segment PAIR (a year-shape segment
// immediately followed by a 1-2 digit segment — the 1-digit case covers
// a dropped trailing digit, e.g. "21-2"), leaving only the segments that
// actually identify WHICH invoice this is.
function extractCoreSegments(identifier) {
  const segments = splitSegments(identifier);
  const core = [];
  let i = 0;
  while (i < segments.length) {
    const seg = segments[i];
    const next = segments[i + 1];
    if (isYearShapeSegment(seg) && next !== undefined && /^\d{1,2}$/.test(next)) {
      i += 2;
      continue;
    }
    core.push(seg);
    i++;
  }
  return core;
}

// True if id1/id2 plausibly refer to the SAME invoice by string alone —
// see the module docstring: this is NECESSARY but never SUFFICIENT on
// its own; callers must also confirm amount/tax agreement.
function invoiceNumbersMatch(id1, id2) {
  if (normalizeInvoiceNumber(id1) === normalizeInvoiceNumber(id2)) return true;

  const core1 = extractCoreSegments(id1);
  const core2 = extractCoreSegments(id2);
  if (core1.length === 0 || core2.length === 0) return false; // nothing distinguishing left on one side -- never treat that as a match

  const minLen = Math.min(core1.length, core2.length);
  for (let idx = 0; idx < minLen - 1; idx++) {
    if (core1[idx] !== core2[idx]) return false; // every segment except the last compared one must match exactly
  }
  const last1 = core1[minLen - 1];
  const last2 = core2[minLen - 1];
  return last1 === last2 || last1.startsWith(last2) || last2.startsWith(last1); // last compared segment: exact, or a plain truncation of the other
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { normalizeInvoiceNumber, invoiceNumbersMatch, extractCoreSegments };
}
