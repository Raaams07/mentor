/*
 * duplicate-invoice-detector.js
 * ---------------------------------
 * Flags likely duplicate invoice entries: rows sharing the same GSTIN,
 * where EITHER the invoice/voucher number matches OR the amount matches
 * (within a small tolerance), AND the two dates fall within a proximity
 * window — but the two match types use DIFFERENT windows and different
 * chaining rules, because they're very different strengths of evidence:
 *
 *   - Same invoice/voucher NUMBER is strong, specific evidence on its
 *     own. Validated against a real expert-completed reconciliation, the
 *     genuine same-supplier duplicates in that file recurred roughly a
 *     MONTH apart (a bookkeeping-close-cycle re-entry pattern: the same
 *     invoice keyed in again during the following month's close), not
 *     "within days" as this module originally assumed. The identifier
 *     window is widened accordingly (45 days) — still bounded, to guard
 *     against a legitimate invoice-numbering reset recurring across a
 *     much longer span (e.g. a new financial year), but wide enough to
 *     actually catch the real-world pattern this rule exists to find.
 *
 *     A same-number pair is NOT treated as a duplicate, though, if the
 *     taxable value or effective tax rate differs between the two rows
 *     (beyond a small tolerance): GSTR-2A legitimately splits ONE invoice
 *     into several lines when it covers multiple HSN/tax rates, and each
 *     line carries the SAME invoice number with a DIFFERENT taxable
 *     value/rate. That's a multi-rate line split, not a re-entry.
 *
 *   - Same AMOUNT alone is much weaker evidence — round and common
 *     amounts (bank fees, standard charges) recur naturally and
 *     non-suspiciously across a ledger. This path uses a much tighter
 *     window (3 days) AND a clique constraint on BOTH date and amount:
 *     accepting a new amount-only match is rejected if it would push the
 *     WHOLE resulting cluster's date span past that same tight window,
 *     OR its amount span past the matching tolerance — not just what the
 *     one new pair looks like in isolation. A plain pairwise check alone
 *     isn't enough here — chained matches (A close to B, B close to C, C
 *     close to D...) can silently merge records spanning months, or
 *     whose amounts silently drift beyond the tolerance one small step at
 *     a time, even though no single pair does. Confirmed against real
 *     data: a vendor with a naturally recurring ₹339 charge was
 *     previously being reported as ONE 13-row cluster spanning 49 days,
 *     entirely via amount-only chaining.
 *
 *     A same-amount pair is also excluded when either signal below is
 *     present, since both are stronger evidence of "routine", not
 *     "duplicate":
 *       - the same (GSTIN, amount) combination recurs elsewhere in the
 *         file too (beyond just this one close pair) — a one-off
 *         accidental re-entry doesn't usually have extra siblings
 *         scattered through the rest of the ledger; a routine charge
 *         does.
 *       - the two invoice/voucher numbers are near-sequential (same
 *         format, one small numeric segment apart) — that's the
 *         signature of two genuinely different, consecutively-numbered
 *         invoices issued close together, not the same invoice entered
 *         twice.
 *
 * Clusters (not just pairs): if the same invoice was entered three or
 * more times, all of them are grouped into one cluster rather than
 * reported as several overlapping pairs.
 *
 * Works on either a GSTR-2A-style sheet or a purchase-register-style
 * sheet — whichever identifier column is available (Invoice Number or
 * Voucher Number) is used; this doesn't compare across the two sheets,
 * only within one at a time (the same invoice entered twice IN Books, or
 * the same invoice appearing twice in 2A due to a supplier filing error,
 * are both within-sheet phenomena).
 */

const DEFAULT_IDENTIFIER_WINDOW_DAYS = 45; // same invoice/voucher number — wide enough for a monthly re-entry pattern, still bounded
const DEFAULT_AMOUNT_ONLY_WINDOW_DAYS = 3; // same amount, no shared identifier — deliberately tight; also caps a merged cluster's TOTAL date span, not just each pair
const DEFAULT_AMOUNT_TOLERANCE = 1; // ₹1 — also caps a merged cluster's TOTAL amount spread for amount-only matches, not just each pair
const DEFAULT_RATE_TOLERANCE_PERCENT = 0.5; // percentage points; GST rate slabs are spaced >=2pp apart (0%, 0.25%, 3%, 5%, 12%, 18%, 28%), so this is a safe rounding cushion, not a slab boundary risk
const DEFAULT_MIN_ELSEWHERE_OCCURRENCES_TO_EXCLUDE = 1; // this many EXTRA rows (beyond the candidate pair itself) sharing the same (GSTIN, amount) anywhere else in the file is treated as evidence of a routine recurring charge — even a single low-volume repeat (3 total occurrences) is enough
const DEFAULT_NEAR_SEQUENTIAL_MAX_DIFF = 5; // floor for the near-sequential check below — used as-is when a vendor-specific typical gap can't be computed (too few data points), and as a lower bound even when it can
const NEAR_SEQUENTIAL_GAP_MULTIPLIER = 3; // how many multiples of a vendor's OWN typical same-period numbering gap still count as "roughly consecutive" rather than a coincidence
const MIN_DATA_POINTS_FOR_TYPICAL_GAP = 4; // need at least this many identifiers sharing a skeleton before trusting a computed "typical gap" over the fixed floor

function toNumber(value) {
  if (typeof value === "number") return value;
  if (value === "" || value === null || value === undefined) return null;
  const parsed = parseFloat(String(value).replace(/[,$£€¥%()]/g, ""));
  return isNaN(parsed) ? null : parsed;
}

function normalizeIdentifier(value) {
  if (value === null || value === undefined) return "";
  const trimmed = String(value).trim().toUpperCase();
  // A purely-numeric identifier ("6" vs "006") is the SAME invoice number
  // written with different zero-padding, not two different invoices —
  // strip leading zeros so both forms compare equal. Deliberately scoped
  // to purely-numeric strings only: an alphanumeric code like "OS0015..."
  // has meaningful leading zeros as part of a fixed-width series, and
  // stripping there would corrupt it.
  if (/^\d+$/.test(trimmed)) {
    return String(parseInt(trimmed, 10));
  }
  return trimmed;
}

function parseDateValue(value) {
  if (value instanceof Date) return isNaN(value.getTime()) ? null : value;
  if (typeof value === "number") {
    // Excel serial date (days since 1899-12-30, the platform's usual epoch quirk)
    const excelEpoch = Date.UTC(1899, 11, 30);
    return new Date(excelEpoch + value * 86400000);
  }
  if (typeof value === "string") {
    const trimmed = value.trim();
    let m = trimmed.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/); // ISO: YYYY-MM-DD
    if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    m = trimmed.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/); // DD-MM-YYYY or DD/MM/YYYY
    if (m) return new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
    const parsed = new Date(trimmed);
    if (!isNaN(parsed.getTime())) return parsed;
  }
  return null;
}

function daysBetween(d1, d2) {
  return Math.abs(d1.getTime() - d2.getTime()) / 86400000;
}

// Taxable-value comparison used to gate a same-invoice-number match (a
// multi-rate line split legitimately has DIFFERENT taxable values per
// line, same invoice number) — null on either side means we simply don't
// know (no Taxable Value column identified), so this doesn't block the
// match rather than guessing.
function amountsMatch(r1, r2, tolerance) {
  if (r1.amount === null || r2.amount === null) return true;
  return Math.abs(r1.amount - r2.amount) <= tolerance;
}

// Effective-tax-rate comparison, same purpose as amountsMatch: a
// multi-rate line split can coincidentally share a taxable value while
// differing in rate (or vice versa), so both are checked independently.
function ratesMatch(r1, r2, tolerancePercent) {
  if (r1.rate === null || r2.rate === null) return true;
  return Math.abs(r1.rate - r2.rate) <= tolerancePercent;
}

// null when NONE of IGST/CGST/SGST were identified at all (nothing to
// compute a rate from); missing individual columns otherwise treated as
// 0, the same way a real return would show an inapplicable tax head.
function computeTotalTax(row, columns) {
  // Loose null checks deliberately: a caller-supplied columns object that
  // simply omits these keys (leaves them undefined) must be treated the
  // same as "not identified" — not as "identified with a zero value".
  const hasAnyTaxColumn = columns.igst != null || columns.cgst != null || columns.sgst != null;
  if (!hasAnyTaxColumn) return null;
  const igst = columns.igst != null ? toNumber(row[columns.igst]) : 0;
  const cgst = columns.cgst != null ? toNumber(row[columns.cgst]) : 0;
  const sgst = columns.sgst != null ? toNumber(row[columns.sgst]) : 0;
  return (igst || 0) + (cgst || 0) + (sgst || 0);
}

function digitRuns(identifier) {
  return identifier.match(/\d+/g) || [];
}

// Replaces every digit run with one placeholder so two identifiers can be
// compared for "same non-numeric structure" regardless of how many
// digits are actually in each run (e.g. "INV-9" and "INV-10" both reduce
// to "INV-").
function skeletonOf(identifier) {
  return identifier.replace(/\d+/g, "#");
}

// True only when id1/id2 share an identical non-digit skeleton AND differ
// in EXACTLY one numeric segment, by no more than maxDiff — e.g.
// "ST/23/2021-22" vs "ST/24/2021-22" (the running number one apart, the
// financial-year segment unchanged). Two identifiers differing in more
// than one segment, or with a different structure entirely, aren't
// treated as comparably "sequential" — there's no reliable way to tell
// which segment is the running counter in that case.
//
// maxDiff is a FLOOR, not the only threshold used — see
// computeTypicalGapsBySkeleton below: a high-volume vendor's "next
// invoice to this customer" can legitimately be dozens or hundreds of
// numbers away (their numbering is shared across every customer they
// bill), so a single fixed small maxDiff systematically misreads those
// as duplicates. Callers that have a vendor-specific typical gap should
// pass the larger of maxDiff and that gap (already multiplied by
// NEAR_SEQUENTIAL_GAP_MULTIPLIER) here instead of the bare default.
function nearSequentialInvoiceNumbers(id1, id2, maxDiff) {
  if (!id1 || !id2) return false;
  if (skeletonOf(id1) !== skeletonOf(id2)) return false;
  const runs1 = digitRuns(id1);
  const runs2 = digitRuns(id2);
  if (runs1.length === 0 || runs1.length !== runs2.length) return false;

  let diffIndex = -1;
  for (let i = 0; i < runs1.length; i++) {
    if (runs1[i] !== runs2[i]) {
      if (diffIndex !== -1) return false; // more than one differing segment -> not a simple sequential increment
      diffIndex = i;
    }
  }
  if (diffIndex === -1) return false; // identical numbers entirely — not this function's concern

  const n1 = parseInt(runs1[diffIndex], 10);
  const n2 = parseInt(runs2[diffIndex], 10);
  if (isNaN(n1) || isNaN(n2)) return false;
  return Math.abs(n1 - n2) <= maxDiff;
}

// Two independent proxies for "how much does this vendor's invoice
// number typically advance, for this customer", per (GSTIN, skeleton) —
// used TOGETHER (the caller takes whichever is larger) because each
// alone misses a different real pattern:
//
//   - medianGap: the median gap between VALUE-SORTED consecutive
//     identifiers. Median is deliberately chosen over a simple min/max
//     range: some vendors mix multiple numbering series under one
//     shared prefix (e.g. "OS0010......", "OS0012......",
//     "OS0015......" all reducing to the same skeleton) — a min/max
//     range would be wrecked by the one or two huge jumps AT a series
//     boundary, but those are a tiny minority of the total gaps, so the
//     median stays representative of the dominant, real spacing.
//     This alone still under-catches a BURSTY vendor, though: one that
//     issues this customer several back-to-back numbers (gap 1) on some
//     days and a handful of genuinely-skipped numbers (gap 5-10, other
//     customers interleaved) on others — the dense runs pull the median
//     down near 1, well below a perfectly normal same-week gap.
//
//   - ratePerDay: total value range / total date range for the skeleton
//     — the vendor's overall issuance pace for this customer. This
//     catches the bursty case above (a multi-day date gap naturally
//     scales up the expected number gap) and, more importantly, a
//     genuinely HIGH-VOLUME vendor whose own historical gap-to-neighbor
//     never happened to reach as high as this specific pair's gap, but
//     whose overall pace easily explains it (e.g. a payment-settlement
//     feed issuing hundreds of sequence numbers a day, where this
//     customer's own observed neighbor-gaps topped out lower than one
//     particular pair's gap purely by chance of sampling).
//
// Only identifiers with EXACTLY ONE digit run are considered (matching
// nearSequentialInvoiceNumbers' own single-differing-segment
// requirement) — a multi-segment identifier like "ST/23/2021-22" isn't
// given a computed reference gap here; it falls back to the fixed floor.
function computeTypicalGapsBySkeleton(group) {
  const bySkeleton = new Map();
  for (const r of group) {
    if (!r.identifier || !r.date) continue;
    const runs = digitRuns(r.identifier);
    if (runs.length !== 1) continue;
    const value = parseInt(runs[0], 10);
    if (isNaN(value)) continue;
    const skel = skeletonOf(r.identifier);
    if (!bySkeleton.has(skel)) bySkeleton.set(skel, []);
    bySkeleton.get(skel).push({ value, date: r.date });
  }

  const typicalGapBySkeleton = new Map();
  for (const [skel, entries] of bySkeleton) {
    const distinctValues = Array.from(new Set(entries.map((e) => e.value))).sort((a, b) => a - b);
    if (distinctValues.length < MIN_DATA_POINTS_FOR_TYPICAL_GAP) continue;

    const gaps = [];
    for (let i = 1; i < distinctValues.length; i++) {
      gaps.push(distinctValues[i] - distinctValues[i - 1]);
    }
    gaps.sort((a, b) => a - b);
    const medianGap = gaps[Math.floor(gaps.length / 2)];

    const dateTimes = entries.map((e) => e.date.getTime());
    const dateRangeDays = (Math.max(...dateTimes) - Math.min(...dateTimes)) / 86400000;
    const valueRange = distinctValues[distinctValues.length - 1] - distinctValues[0];
    const ratePerDay = dateRangeDays > 0 ? valueRange / dateRangeDays : null;

    typicalGapBySkeleton.set(skel, { medianGap, ratePerDay });
  }
  return typicalGapBySkeleton;
}

// Combines two clusters' running [min,max] span for both date and
// amount. Amount span collapses to null (unknown, non-restricting) if
// EITHER side doesn't have a known amount, the same "don't block on
// missing information" rule used elsewhere in this module.
function mergeSpan(spanA, spanB) {
  const dateMin = Math.min(spanA.dateMin, spanB.dateMin);
  const dateMax = Math.max(spanA.dateMax, spanB.dateMax);
  const hasAmount = spanA.amountMin !== null && spanB.amountMin !== null;
  return {
    dateMin,
    dateMax,
    amountMin: hasAmount ? Math.min(spanA.amountMin, spanB.amountMin) : null,
    amountMax: hasAmount ? Math.max(spanA.amountMax, spanB.amountMax) : null,
  };
}

function determineReason(members, amountTolerance) {
  const allSameIdentifier = members.every((m) => m.identifier && m.identifier === members[0].identifier);
  const allSameAmount = members.every((m) => m.amount !== null && Math.abs(m.amount - members[0].amount) <= amountTolerance);
  if (allSameIdentifier && allSameAmount) return "same_invoice_number_and_amount";
  if (allSameIdentifier) return "same_invoice_number";
  if (allSameAmount) return "same_amount";
  return "chained"; // e.g. A matches B on number, B matches C on amount — grouped transitively
}

// Which matchReason values represent CONFIRMED evidence (every member of
// the cluster shares an identical invoice/voucher number) vs. a WEAKER
// signal that relies on amount alone somewhere in the cluster ("same_amount",
// or "chained", which by construction always includes at least one
// amount-only bridging link — see determineReason above: a cluster only
// reaches "chained" when NOT every member shares one identifier AND NOT
// every member shares one amount, which is only reachable via an
// amount-only union connecting members with different identifiers).
//
// Verified against a real answer-key file: a same_amount-only pair (two
// DIFFERENT NSDL e-Governance invoice numbers, coincidentally both
// ₹42.37) was independently confirmed to be a real "Extra in 2A" item,
// not a duplicate at all — same_amount alone is not reliable enough to
// present with the same confidence as a shared invoice number. Callers
// (gst-report-writer.js) use this to keep the two kinds of evidence in
// separate, differently-labeled sections rather than blending them.
const CONFIRMED_DUPLICATE_MATCH_REASONS = new Set(["same_invoice_number_and_amount", "same_invoice_number"]);
function isConfirmedDuplicateMatch(matchReason) {
  return CONFIRMED_DUPLICATE_MATCH_REASONS.has(matchReason);
}

// values/headerRowIndex: the sheet's raw data + header position.
// columns: identifyGstColumns() result for this sheet.
function findDuplicateInvoices(values, headerRowIndex, columns, options) {
  const opts = options || {};
  const identifierWindowDays = opts.identifierWindowDays !== undefined ? opts.identifierWindowDays : DEFAULT_IDENTIFIER_WINDOW_DAYS;
  const amountOnlyWindowDays = opts.amountOnlyWindowDays !== undefined ? opts.amountOnlyWindowDays : DEFAULT_AMOUNT_ONLY_WINDOW_DAYS;
  const amountTolerance = opts.amountTolerance !== undefined ? opts.amountTolerance : DEFAULT_AMOUNT_TOLERANCE;
  const rateTolerancePercent = opts.rateTolerancePercent !== undefined ? opts.rateTolerancePercent : DEFAULT_RATE_TOLERANCE_PERCENT;
  const minElsewhereOccurrencesToExclude =
    opts.minElsewhereOccurrencesToExclude !== undefined ? opts.minElsewhereOccurrencesToExclude : DEFAULT_MIN_ELSEWHERE_OCCURRENCES_TO_EXCLUDE;
  const nearSequentialMaxDiff = opts.nearSequentialMaxDiff !== undefined ? opts.nearSequentialMaxDiff : DEFAULT_NEAR_SEQUENTIAL_MAX_DIFF;

  if (columns.gstin === null) {
    return { applicable: false, reason: "no GSTIN column identified", clusters: [] };
  }
  const identifierCol = columns.invoiceNumber !== null ? columns.invoiceNumber : columns.voucherNumber;
  const dateCol = columns.dateColumns && columns.dateColumns.length > 0 ? columns.dateColumns[0] : null;
  if (identifierCol === null && columns.taxableValue === null) {
    return { applicable: false, reason: "no invoice/voucher number AND no taxable value column identified — nothing to match on", clusters: [] };
  }
  if (dateCol === null) {
    return { applicable: false, reason: "no date column identified — can't apply the proximity window", clusters: [] };
  }

  const dataRows = headerRowIndex === -1 ? values : values.slice(headerRowIndex + 1);
  const records = [];
  dataRows.forEach((row, i) => {
    const gstinRaw = row[columns.gstin];
    if (!gstinRaw || String(gstinRaw).trim() === "") return;
    const amount = columns.taxableValue !== null ? toNumber(row[columns.taxableValue]) : null;
    const totalTax = computeTotalTax(row, columns);
    records.push({
      rowIndex: i,
      gstin: String(gstinRaw).trim().toUpperCase(),
      identifier: identifierCol !== null ? normalizeIdentifier(row[identifierCol]) : "",
      amount,
      rate: amount !== null && amount > 0 && totalTax !== null ? (totalTax / amount) * 100 : null,
      date: parseDateValue(row[dateCol]),
    });
  });

  const byGstin = new Map();
  records.forEach((r) => {
    if (!byGstin.has(r.gstin)) byGstin.set(r.gstin, []);
    byGstin.get(r.gstin).push(r);
  });

  const clusters = [];
  for (const [gstin, group] of byGstin) {
    if (group.length < 2) continue;

    const typicalGapBySkeleton = computeTypicalGapsBySkeleton(group);
    const effectiveNearSequentialMaxDiff = (identifier1, identifier2, pairDateGapDays) => {
      const skel = skeletonOf(identifier1);
      const typical = typicalGapBySkeleton.get(skel);
      if (typical === undefined) return nearSequentialMaxDiff;
      const fromMedianGap = typical.medianGap * NEAR_SEQUENTIAL_GAP_MULTIPLIER;
      const fromRate = typical.ratePerDay !== null ? typical.ratePerDay * pairDateGapDays * NEAR_SEQUENTIAL_GAP_MULTIPLIER : 0;
      return Math.max(nearSequentialMaxDiff, fromMedianGap, fromRate);
    };

    // Union-find so 3+ entries of the same invoice group into ONE cluster,
    // not several overlapping pairs.
    const parent = group.map((_, i) => i);
    const find = (x) => {
      while (parent[x] !== x) {
        parent[x] = parent[parent[x]];
        x = parent[x];
      }
      return x;
    };
    const union = (a, b) => {
      const ra = find(a);
      const rb = find(b);
      if (ra !== rb) parent[ra] = rb;
      return find(a);
    };

    // Tracks each cluster's [min, max] for BOTH date and amount (epoch ms
    // / rupees) by current root — kept up to date on EVERY union
    // (identifier or amount) so a lookup is never stale after the
    // canonical root shifts, but only the amount-only path actually
    // REJECTS a union based on it.
    const clusterSpanByRoot = new Map();
    const currentSpan = (root, fallbackRecord) =>
      clusterSpanByRoot.get(root) || {
        dateMin: fallbackRecord.date.getTime(),
        dateMax: fallbackRecord.date.getTime(),
        amountMin: fallbackRecord.amount,
        amountMax: fallbackRecord.amount,
      };

    for (let a = 0; a < group.length; a++) {
      for (let b = a + 1; b < group.length; b++) {
        const r1 = group[a];
        const r2 = group[b];
        if (!r1.date || !r2.date) continue; // can't judge proximity without both dates

        const sameIdentifier = r1.identifier && r2.identifier && r1.identifier === r2.identifier;
        const sameAmount = r1.amount !== null && r2.amount !== null && Math.abs(r1.amount - r2.amount) <= amountTolerance;
        if (!sameIdentifier && !sameAmount) continue;

        const pairDays = daysBetween(r1.date, r2.date);

        if (sameIdentifier) {
          if (pairDays > identifierWindowDays) continue;
          if (!amountsMatch(r1, r2, amountTolerance)) continue; // multi-rate line split: different taxable value -> not a duplicate
          if (!ratesMatch(r1, r2, rateTolerancePercent)) continue; // multi-rate line split: different tax rate -> not a duplicate
          if (find(a) === find(b)) continue;
          const merged = mergeSpan(currentSpan(find(a), r1), currentSpan(find(b), r2));
          const newRoot = union(a, b);
          clusterSpanByRoot.set(newRoot, merged);
          continue;
        }

        // Amount-only match — see the module docstring for why this path
        // is deliberately much stricter than the identifier path.
        if (pairDays > amountOnlyWindowDays) continue;
        if (find(a) === find(b)) continue;

        // Recurs elsewhere in the file at all (beyond this one close
        // pair) -> a routine amount, not a one-off duplicate.
        const elsewhereCount = group.reduce((count, r, idx) => {
          if (idx === a || idx === b || r.amount === null) return count;
          return Math.abs(r.amount - r1.amount) <= amountTolerance ? count + 1 : count;
        }, 0);
        if (elsewhereCount >= minElsewhereOccurrencesToExclude) continue;

        // Near-sequential invoice numbers with this already-small date
        // gap is the signature of genuine consecutive invoicing, not the
        // same invoice re-entered. "Near" is scaled to THIS vendor's own
        // typical numbering gap where there's enough data to compute
        // one — a high-volume vendor's next invoice to this customer can
        // legitimately be dozens or hundreds of numbers away.
        if (nearSequentialInvoiceNumbers(r1.identifier, r2.identifier, effectiveNearSequentialMaxDiff(r1.identifier, r2.identifier, pairDays))) continue;

        const merged = mergeSpan(currentSpan(find(a), r1), currentSpan(find(b), r2));
        const mergedDateDays = (merged.dateMax - merged.dateMin) / 86400000;
        if (mergedDateDays > amountOnlyWindowDays) continue; // accepting this pair would blow the WHOLE resulting cluster's date span past the tight window — reject the bridge, not just re-check this one pair
        if (merged.amountMin !== null && merged.amountMax - merged.amountMin > amountTolerance) continue; // same idea, for amount: don't let chaining drift the WHOLE cluster's amount spread past tolerance

        const newRoot = union(a, b);
        clusterSpanByRoot.set(newRoot, merged);
      }
    }

    const clusterGroups = new Map();
    group.forEach((r, i) => {
      const root = find(i);
      if (!clusterGroups.has(root)) clusterGroups.set(root, []);
      clusterGroups.get(root).push(r);
    });

    for (const members of clusterGroups.values()) {
      if (members.length < 2) continue;
      const dates = members.map((m) => m.date.getTime());
      clusters.push({
        gstin,
        matchReason: determineReason(members, amountTolerance),
        dateSpreadDays: Math.round((Math.max(...dates) - Math.min(...dates)) / 86400000),
        members: members.map((m) => ({ rowIndex: m.rowIndex, identifier: m.identifier, amount: m.amount, date: m.date })),
      });
    }
  }

  return { applicable: true, clusters };
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    findDuplicateInvoices,
    parseDateValue,
    nearSequentialInvoiceNumbers,
    computeTypicalGapsBySkeleton,
    normalizeIdentifier,
    isConfirmedDuplicateMatch,
    CONFIRMED_DUPLICATE_MATCH_REASONS,
    DEFAULT_IDENTIFIER_WINDOW_DAYS,
    DEFAULT_AMOUNT_ONLY_WINDOW_DAYS,
    DEFAULT_AMOUNT_TOLERANCE,
    DEFAULT_RATE_TOLERANCE_PERCENT,
    DEFAULT_MIN_ELSEWHERE_OCCURRENCES_TO_EXCLUDE,
    DEFAULT_NEAR_SEQUENTIAL_MAX_DIFF,
    NEAR_SEQUENTIAL_GAP_MULTIPLIER,
    MIN_DATA_POINTS_FOR_TYPICAL_GAP,
  };
}
