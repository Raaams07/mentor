/*
 * gstin.js
 * --------
 * GST reconciliation — GSTIN format and check-digit validation.
 *
 * A GSTIN (Goods and Services Tax Identification Number) has a fixed
 * 15-character structure defined by GSTN's registration specification,
 * identical for every taxpayer in India:
 *   - 2 digits:  state code
 *   - 10 chars:  the taxpayer's PAN (5 letters, 4 digits, 1 letter)
 *   - 1 char:    entity/registration number of this PAN within the state
 *   - 1 letter:  default 'Z'
 *   - 1 char:    check digit
 *
 * Detecting this shape is a legally-grounded structural signal — it's the
 * same format for every company's GSTIN, not something specific to any
 * one file's data.
 */

const GSTIN_FORMAT = /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$/;
const GSTIN_CHECKSUM_ALPHABET = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";

function normalizeGstin(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim().toUpperCase();
}

function matchesGstinFormat(value) {
  return GSTIN_FORMAT.test(normalizeGstin(value));
}

// GSTN's published check-digit algorithm: a base-36 Luhn-style checksum
// over the first 14 characters, validated against the 15th. Used only as a
// CORROBORATING signal elsewhere in this module (see hasValidGstinChecksum)
// — a well-formed-but-checksum-failing value still counts as "looks like a
// GSTIN" for recognition purposes, so a subtle bug in this one function
// can't cause the recognizer to reject a genuinely GSTIN-shaped column.
function computeGstinCheckDigit(first14Chars) {
  const mod = GSTIN_CHECKSUM_ALPHABET.length;
  let factor = 2;
  let sum = 0;
  for (let i = first14Chars.length - 1; i >= 0; i--) {
    const codePoint = GSTIN_CHECKSUM_ALPHABET.indexOf(first14Chars[i]);
    let digit = factor * codePoint;
    digit = Math.floor(digit / mod) + (digit % mod);
    sum += digit;
    factor = factor === 2 ? 1 : 2;
  }
  const checkCodePoint = (mod - (sum % mod)) % mod;
  return GSTIN_CHECKSUM_ALPHABET[checkCodePoint];
}

function hasValidGstinChecksum(value) {
  const gstin = normalizeGstin(value);
  if (!matchesGstinFormat(gstin)) return false;
  return computeGstinCheckDigit(gstin.slice(0, 14)) === gstin[14];
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { matchesGstinFormat, hasValidGstinChecksum, computeGstinCheckDigit, normalizeGstin, GSTIN_FORMAT };
}
