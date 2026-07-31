/*
 * duplicate-invoice-detector.js
 * ---------------------------------
 * Flags likely duplicate invoice entries: rows sharing the same GSTIN,
 * where EITHER the invoice/voucher number matches OR the amount matches
 * (within a small tolerance), AND the two dates fall within a configurable
 * proximity window.
 *
 * The date-proximity requirement is deliberate: a duplicate is normally a
 * DATA ENTRY error, which happens close in time (the same invoice keyed
 * in twice within days or weeks) — whereas a different, genuinely separate
 * invoice from the same vendor for a coincidentally identical amount many
 * months apart is far more likely to be a real, distinct transaction, not
 * a duplicate. Requiring proximity is what keeps this from over-flagging
 * routine recurring charges (e.g. an identical monthly rent amount).
 *
 * Works on either a GSTR-2A-style sheet or a purchase-register-style
 * sheet — whichever identifier column is available (Invoice Number or
 * Voucher Number) is used; this doesn't compare across the two sheets,
 * only within one at a time (the same invoice entered twice IN Books, or
 * the same invoice appearing twice in 2A due to a supplier filing error,
 * are both within-sheet phenomena).
 *
 * Clusters (not just pairs): if the same invoice was entered three or more
 * times, all of them are grouped into one cluster rather than reported as
 * several overlapping pairs.
 *
 * Default window is intentionally tight (15 days), not weeks — a genuine
 * data-entry duplicate is almost always keyed in within days of the
 * original, while a wide window (e.g. 60 days) would false-positive on
 * ordinary recurring charges of an identical amount (two consecutive
 * months' rent, ~30 days apart, from the same GSTIN). Callers who want a
 * wider window for a specific use case can pass options.windowDays.
 */

const DEFAULT_WINDOW_DAYS = 15;
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
  const windowDays = opts.windowDays !== undefined ? opts.windowDays : DEFAULT_WINDOW_DAYS;
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
    };

    for (let a = 0; a < group.length; a++) {
      for (let b = a + 1; b < group.length; b++) {
        const r1 = group[a];
        const r2 = group[b];
        if (!r1.date || !r2.date) continue; // can't judge proximity without both dates
        if (daysBetween(r1.date, r2.date) > windowDays) continue;

        const sameIdentifier = r1.identifier && r2.identifier && r1.identifier === r2.identifier;
        const sameAmount = r1.amount !== null && r2.amount !== null && Math.abs(r1.amount - r2.amount) <= amountTolerance;
        if (sameIdentifier || sameAmount) union(a, b);
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
  module.exports = { findDuplicateInvoices, parseDateValue, DEFAULT_WINDOW_DAYS, DEFAULT_AMOUNT_TOLERANCE };
}
