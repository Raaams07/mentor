/*
 * gst-state-codes.js
 * --------------------
 * GST state/UT codes — the first 2 digits of every GSTIN (see gstin.js),
 * as published by GSTN. Used to compare a supplier's registered state
 * (from their GSTIN) against a "Place of Supply" field, which is what
 * determines IGST vs CGST+SGST under IGST Act Sections 7 & 8 — see
 * wrong-head-detector.js, the only current consumer of this table.
 *
 * Robustness note: real GSTR-2A exports very often write Place of Supply
 * as "<code>-<StateName>" (e.g. "27-Maharashtra"), which lets us compare
 * codes directly without touching this name table at all — see
 * extractStateCodeFromPlaceOfSupply() below, which tries that first. The
 * name-based lookup here is only the fallback for plain-text state names.
 *
 * This list should be verified against the current CBIC/GSTN state-code
 * notification before relying on it for anything beyond this framework's
 * own validation — kept as one small, easily-correctable table rather than
 * scattered through logic, for exactly that reason.
 */

const STATE_CODE_TO_NAME = {
  "01": "Jammu and Kashmir",
  "02": "Himachal Pradesh",
  "03": "Punjab",
  "04": "Chandigarh",
  "05": "Uttarakhand",
  "06": "Haryana",
  "07": "Delhi",
  "08": "Rajasthan",
  "09": "Uttar Pradesh",
  "10": "Bihar",
  "11": "Sikkim",
  "12": "Arunachal Pradesh",
  "13": "Nagaland",
  "14": "Manipur",
  "15": "Mizoram",
  "16": "Tripura",
  "17": "Meghalaya",
  "18": "Assam",
  "19": "West Bengal",
  "20": "Jharkhand",
  "21": "Odisha",
  "22": "Chhattisgarh",
  "23": "Madhya Pradesh",
  "24": "Gujarat",
  "26": "Dadra and Nagar Haveli and Daman and Diu",
  "27": "Maharashtra",
  "28": "Andhra Pradesh", // historical code retained by some legacy registrations pre-bifurcation
  "29": "Karnataka",
  "30": "Goa",
  "31": "Lakshadweep",
  "32": "Kerala",
  "33": "Tamil Nadu",
  "34": "Puducherry",
  "35": "Andaman and Nicobar Islands",
  "36": "Telangana",
  "37": "Andhra Pradesh",
  "38": "Ladakh",
  "97": "Other Territory",
  "99": "Centre Jurisdiction",
};

const NAME_TO_STATE_CODE = {};
for (const [code, name] of Object.entries(STATE_CODE_TO_NAME)) {
  NAME_TO_STATE_CODE[name.toLowerCase()] = code;
}
// A few common alternate spellings/abbreviations seen in real exports.
const NAME_ALIASES = {
  "j&k": "01",
  "jammu & kashmir": "01",
  "nct of delhi": "07",
  "orissa": "21", // pre-2011 official spelling of Odisha
  "pondicherry": "34", // pre-2006 official spelling of Puducherry
  "dnh and dd": "26",
  "dadra & nagar haveli and daman & diu": "26",
  "a&n islands": "35",
  "andaman & nicobar islands": "35",
};

function stateCodeFromGstin(gstin) {
  if (!gstin || gstin.length < 2) return null;
  const code = gstin.slice(0, 2);
  return STATE_CODE_TO_NAME[code] ? code : null;
}

// Tries the robust "<code>-<name>" prefix first (common in real GSTR-2A
// exports), then falls back to matching the text against the state name
// table directly.
function extractStateCodeFromPlaceOfSupply(placeOfSupply) {
  if (!placeOfSupply) return null;
  const text = String(placeOfSupply).trim();

  const prefixMatch = text.match(/^(\d{2})\s*[-–—]/);
  if (prefixMatch && STATE_CODE_TO_NAME[prefixMatch[1]]) return prefixMatch[1];

  const normalized = text.toLowerCase().replace(/^\d{2}\s*[-–—]\s*/, "").trim();
  if (NAME_TO_STATE_CODE[normalized]) return NAME_TO_STATE_CODE[normalized];
  if (NAME_ALIASES[normalized]) return NAME_ALIASES[normalized];

  return null;
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { STATE_CODE_TO_NAME, stateCodeFromGstin, extractStateCodeFromPlaceOfSupply };
}
