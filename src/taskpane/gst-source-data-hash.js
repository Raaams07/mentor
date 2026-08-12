/*
 * gst-source-data-hash.js
 * -----------------------------
 * A cheap, deterministic fingerprint of the 2A + Books sheets' raw data —
 * used to answer "has the underlying source data changed since MENTOR
 * last generated the GST review sheets," so a workbook close/reopen with
 * NO data change stays silent instead of re-prompting and re-running
 * everything (which previously produced near-duplicate Audit Log entries
 * on every session, even with zero real changes).
 *
 * This is NOT a cryptographic hash — no security property is needed here,
 * only "did this specific workbook's own data change since I last looked
 * at it." A 32-bit FNV-1a hash over both sheets' serialized values, plus
 * row-count and total-length as cheap extra collision resistance, is
 * more than sufficient for that — and runs in pure JS with no dependency
 * on Node's crypto module (unavailable in the browser-bundled task pane)
 * or the Web Crypto API (async, unnecessary overhead for this).
 */

function fnv1aHash(str) {
  let hash = 0x811c9dc5;
  for (let i = 0; i < str.length; i++) {
    hash ^= str.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(16);
}

// Unicode control characters (0x01/0x02/0x03) as column/row/sheet
// separators when serializing — they essentially never appear in real
// spreadsheet cell content, so e.g. a row ["A", "BC"] and a row ["AB",
// "C"] don't accidentally serialize to the same string the way naive
// no-separator concatenation would.
var COLUMN_SEPARATOR = String.fromCharCode(1);
var ROW_SEPARATOR = String.fromCharCode(2);
var SHEET_SEPARATOR = String.fromCharCode(3);

function serializeSheetValues(values) {
  return values.map((row) => row.join(COLUMN_SEPARATOR)).join(ROW_SEPARATOR);
}

// gstr2aValues/booksValues: each sheet's full used-range values (2D
// array), exactly as read via usedRange.values — no interpretation, just
// fingerprinting the raw content both sheets currently hold.
function computeGstSourceDataHash(gstr2aValues, booksValues) {
  const gstr2aText = serializeSheetValues(gstr2aValues);
  const booksText = serializeSheetValues(booksValues);
  const combined = gstr2aText + SHEET_SEPARATOR + booksText;
  return fnv1aHash(combined) + "-" + gstr2aValues.length + "x" + booksValues.length + "-" + combined.length;
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { computeGstSourceDataHash };
}
