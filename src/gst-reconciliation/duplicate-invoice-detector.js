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
 *   - Same AMOUNT alone is much weaker evidence — round and common
 *     amounts (bank fees, standard charges) recur naturally and
 *     non-suspiciously across a ledger. This path uses a much tighter
 *     window (3 days) AND a clique constraint: accepting a new amount-
 *     only match is rejected if it would push the WHOLE resulting
 *     cluster's date span past that same tight window, not just the one
 *     new pair. A plain pairwise check alone isn't enough here — chained
 *     matches (A close to B, B close to C, C close to D...) can silently
 *     merge records spanning months even though no single pair does, and
 *     that's exactly the false-positive pattern this rule was built to
 *     avoid. Confirmed against real data: a vendor with a naturally
 *     recurring ₹339 charge was previously being reported as ONE 13-row
 *     cluster spanning 49 days, entirely via amount-only chaining.
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
const DEFAULT_AMOUNT_ONLY_WINDOW_DAYS = 3; // same amount, no shared identifier — deliberately tight; also caps a merged cluster's TOTAL span, not just each pair
const DEFAULT_AMOUNT_TOLERANCE = 1; // ₹1

function toNumber(value) {
  if (typeof value === "number") return value;
  if (value === "" || value === null || value === undefined) return null;
  const parsed = parseFloat(String(value).replace(/[,$£€¥%()]/g, ""));
  return isNaN(parsed) ? null : parsed;
}

function normalizeIdentifier(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim().toUpperCase();
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

function determineReason(members, amountTolerance) {
  const allSameIdentifier = members.every((m) => m.identifier && m.identifier === members[0].identifier);
  const allSameAmount = members.every((m) => m.amount !== null && Math.abs(m.amount - members[0].amount) <= amountTolerance);
  if (allSameIdentifier && allSameAmount) return "same_invoice_number_and_amount";
  if (allSameIdentifier) return "same_invoice_number";
  if (allSameAmount) return "same_amount";
  return "chained"; // e.g. A matches B on number, B matches C on amount — grouped transitively
}

// values/headerRowIndex: the sheet's raw data + header position.
// columns: identifyGstColumns() result for this sheet.
function findDuplicateInvoices(values, headerRowIndex, columns, options) {
  const opts = options || {};
  const identifierWindowDays = opts.identifierWindowDays !== undefined ? opts.identifierWindowDays : DEFAULT_IDENTIFIER_WINDOW_DAYS;
  const amountOnlyWindowDays = opts.amountOnlyWindowDays !== undefined ? opts.amountOnlyWindowDays : DEFAULT_AMOUNT_ONLY_WINDOW_DAYS;
  const amountTolerance = opts.amountTolerance !== undefined ? opts.amountTolerance : DEFAULT_AMOUNT_TOLERANCE;

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
    records.push({
      rowIndex: i,
      gstin: String(gstinRaw).trim().toUpperCase(),
      identifier: identifierCol !== null ? normalizeIdentifier(row[identifierCol]) : "",
      amount: columns.taxableValue !== null ? toNumber(row[columns.taxableValue]) : null,
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

    // Tracks each cluster's [min, max] date (epoch ms) by current root —
    // kept up to date on EVERY union (identifier or amount) so a lookup
    // is never stale after the canonical root shifts, but only the
    // amount-only path actually REJECTS a union based on it.
    const clusterSpanByRoot = new Map();
    const currentSpan = (root, fallbackRecord) => clusterSpanByRoot.get(root) || { min: fallbackRecord.date.getTime(), max: fallbackRecord.date.getTime() };

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
          if (find(a) === find(b)) continue;
          const spanA = currentSpan(find(a), r1);
          const spanB = currentSpan(find(b), r2);
          const merged = { min: Math.min(spanA.min, spanB.min), max: Math.max(spanA.max, spanB.max) };
          const newRoot = union(a, b);
          clusterSpanByRoot.set(newRoot, merged);
          continue;
        }

        // Amount-only match — see the module docstring for why this path
        // is deliberately much stricter than the identifier path.
        if (pairDays > amountOnlyWindowDays) continue;
        if (find(a) === find(b)) continue;

        const spanA = currentSpan(find(a), r1);
        const spanB = currentSpan(find(b), r2);
        const merged = { min: Math.min(spanA.min, spanB.min), max: Math.max(spanA.max, spanB.max) };
        const mergedSpanDays = (merged.max - merged.min) / 86400000;
        if (mergedSpanDays > amountOnlyWindowDays) continue; // accepting this pair would blow the WHOLE resulting cluster's span past the tight window — reject the bridge, not just re-check this one pair

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
  module.exports = { findDuplicateInvoices, parseDateValue, DEFAULT_IDENTIFIER_WINDOW_DAYS, DEFAULT_AMOUNT_ONLY_WINDOW_DAYS, DEFAULT_AMOUNT_TOLERANCE };
}
