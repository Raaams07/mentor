/*
 * sheet-label-plausibility.js
 * -----------------------------
 * Basic sanity check for the sheet-memory "what is this sheet?" prompt:
 * before permanently remembering a free-text label a user typed, check
 * whether it shares any real vocabulary with the sheet's actual column
 * headers.
 *
 * Motivated by a real incident: a user answered "Journal entries" for a
 * sheet ("PT 2A") whose headers were GSTIN/Taxable Value/CGST/SGST/IGST —
 * a GST pivot table, not a journal — and MENTOR saved it verbatim with no
 * check, permanently mislabeling that sheet shape.
 *
 * Deliberately conservative in what it calls implausible: it only flags
 * ZERO shared vocabulary between the label and the headers, and even then
 * only WARNS (the UI layer lets the user save anyway) rather than
 * blocking — plenty of legitimate labels don't literally repeat a header
 * word (e.g. "Stock levels" for headers "Item Code, Qty, Warehouse" is a
 * fine label with no lexical overlap). This is a tripwire for the
 * "wildly, obviously wrong" case, not a strict validator — false alarms
 * cost one extra click; a missed bad label costs a silently wrong memory
 * entry, so it's tuned to err toward flagging.
 *
 * Pure function, no Office.js/DOM dependency — testable directly in Node,
 * reusable from the browser taskpane bundle unchanged.
 */

const STOPWORDS = new Set(["the", "a", "an", "of", "for", "and", "or", "in", "on", "at", "to", "this", "that", "is", "are", "sheet", "tab", "data", "workbook", "table", "file"]);

function tokenize(text) {
  return String(text || "")
    .toLowerCase()
    .split(/[^a-z0-9]+/)
    .filter((t) => t.length >= 3 && !STOPWORDS.has(t));
}

// Loose stem: strip a trailing "s"/"es"/"ies" so "invoices" ~ "invoice"
// matches without pulling in a real stemming library for one basic check.
function looseStem(token) {
  if (token.endsWith("ies") && token.length > 4) return token.slice(0, -3) + "y";
  if (token.endsWith("es") && token.length > 4) return token.slice(0, -2);
  if (token.endsWith("s") && token.length > 3) return token.slice(0, -1);
  return token;
}

function tokensOverlap(a, b) {
  if (a === b) return true;
  if (looseStem(a) === looseStem(b)) return true;
  // Substring containment either direction — catches "gst" vs "gstin",
  // "invoice" vs "invoices/invoicing", etc. without full stemming.
  return a.includes(b) || b.includes(a);
}

// label: the free-text string the user typed. sheetSignals: the same
// object already computed for this sheet (extractSheetSignals output) —
// uses .headerLabels, the raw header row.
function checkLabelPlausibility({ label, sheetSignals }) {
  const labelTokens = tokenize(label);
  const headerTexts = ((sheetSignals && sheetSignals.headerLabels) || []).filter((h) => h !== null && h !== undefined && String(h).trim() !== "");

  const headerTokens = [];
  for (const h of headerTexts) {
    for (const t of tokenize(h)) headerTokens.push(t);
  }

  if (labelTokens.length === 0 || headerTokens.length === 0) {
    // Nothing meaningful to compare (a very short/generic label like "Misc",
    // or a sheet with no detectable header row/text) — not enough signal
    // to flag anything, so don't produce a false alarm.
    return { plausible: true, reason: "not_enough_signal", sharedTokens: [], headerLabels: headerTexts };
  }

  const sharedTokens = new Set();
  for (const lt of labelTokens) {
    for (const ht of headerTokens) {
      if (tokensOverlap(lt, ht)) {
        sharedTokens.add(lt);
        break;
      }
    }
  }

  return {
    plausible: sharedTokens.size > 0,
    reason: sharedTokens.size > 0 ? "shares_vocabulary" : "no_shared_vocabulary",
    sharedTokens: [...sharedTokens],
    headerLabels: headerTexts,
  };
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { checkLabelPlausibility, tokenize };
}
