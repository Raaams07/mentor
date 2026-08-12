/*
 * column-memory.js
 * -----------------
 * Ask-once-remember for GST FIELD->COLUMN resolutions — the column-level
 * counterpart to sheet-classifier/sheet-memory.js's "what is this sheet"
 * memory. Same shape of decision ("have we seen this exact shape before,
 * and if not, hand back a ready-to-render prompt instead of guessing"),
 * applied to which column plays which GST-legal role instead of which
 * sheet plays which type.
 *
 * Storage design: ONE record per (client_id, structural_signature), whose
 * label string holds a JSON blob of every field resolved for that shape so
 * far — e.g. {"taxableValue":"taxable value","cgst":"cgst amount"}. This
 * reuses sheet-classifier/sheet-memory-store-base.js's BaseSheetMemoryStore
 * completely unmodified (findExact/findSimilar/remember all key on
 * (client_id, structural_signature) and assume exactly one record per
 * pair) — a "label" is just a string, and a JSON string is a valid string.
 * The alternative (one row per (signature, field_name)) doesn't fit that
 * assumption: resolving field #2 for a signature would look up "the"
 * existing record via findExact and silently overwrite field #1's row,
 * since findExact has no way to disambiguate by field.
 *
 * The blob stores NORMALIZED HEADER TEXT per field, not a raw column
 * index — indices shift as columns are added/removed, but header text is
 * stable. At resolution time, a stored header is only trusted if a column
 * with that exact normalized header still exists on the CURRENT sheet; if
 * the sheet's shape changed enough that it's gone, resolution falls
 * through to needs_input again rather than erroring or guessing.
 *
 * No Office.js/DOM dependency, same as sheet-memory.js — the UI layer
 * (mentor-column-memory-ui.js) renders `prompt`/`candidates` however it
 * likes and calls rememberColumnField() with whatever the user picks.
 */

const { computeStructuralSignature, normalizeHeaderText } = require("../sheet-classifier/structural-signature.js");

function parseBlob(record) {
  if (!record || !record.user_provided_label) return {};
  try {
    return JSON.parse(record.user_provided_label) || {};
  } catch (e) {
    return {};
  }
}

// Looks up fieldName's stored header text in `record`'s blob and checks
// whether a column with that exact normalized header still exists on the
// CURRENT sheet (sheetSignals). Returns the matching column index, or null
// if nothing was stored for this field, or the stored header is gone.
function findStoredColumnIndex(sheetSignals, record, fieldName) {
  const blob = parseBlob(record);
  const storedHeaderText = blob[fieldName];
  if (!storedHeaderText) return null;
  const match = sheetSignals.columns.find((c) => normalizeHeaderText(c.header) === storedHeaderText);
  return match ? match.index : null;
}

// candidateIndices: from gst-column-ambiguity-rules.js's fieldsInPlay() —
// [] for a zero-match-required field. allColumnIndices: every column index
// in the sheet, used ONLY as the pickable list when candidateIndices is
// empty (nothing matched the header rule at all, so there's no narrower
// list to offer).
async function resolveColumnField({ clientId, sheetName, sheetSignals, fieldName, candidateIndices, allColumnIndices, store }) {
  if (!clientId) throw new Error("resolveColumnField requires a clientId — column memory is scoped per client");

  const { signature, headerSignature } = computeStructuralSignature(sheetSignals);

  const exact = await store.findExact(clientId, signature);
  const exactMatch = findStoredColumnIndex(sheetSignals, exact, fieldName);
  if (exactMatch !== null) {
    return { status: "resolved_from_memory", fieldName, fieldIndex: exactMatch, matchedVia: "exact" };
  }

  const fuzzy = await store.findSimilar(clientId, signature, headerSignature);
  const fuzzyMatch = fuzzy ? findStoredColumnIndex(sheetSignals, fuzzy.record, fieldName) : null;
  if (fuzzyMatch !== null) {
    return { status: "resolved_from_memory", fieldName, fieldIndex: fuzzyMatch, matchedVia: "fuzzy", similarity: fuzzy.similarity };
  }

  const pickFrom = candidateIndices.length > 0 ? candidateIndices : allColumnIndices;
  const candidates = pickFrom.map((idx) => {
    const col = sheetSignals.columns.find((c) => c.index === idx);
    return { index: idx, header: col ? col.header : null, sampleValues: col ? col.sampleValues : [] };
  });

  return {
    status: "needs_input",
    fieldName,
    sheetName,
    structuralSignature: signature,
    prompt: `I can't confidently tell which column is "${fieldName}" on '${sheetName}' — which one is it?`,
    candidates,
  };
}

// Read-modify-write on the JSON blob so multiple fields accumulate
// independently for the same signature. Always writes under the sheet's
// FRESHLY COMPUTED exact signature — never back into a fuzzy-matched old
// record — mirroring sheet-memory.js's rememberSheetLabel(), which always
// remembers under the current signature, never the one a fuzzy hit
// happened to match.
async function rememberColumnField({ clientId, sheetName, sheetSignals, fieldName, chosenColumnIndex, store }) {
  if (!clientId) throw new Error("rememberColumnField requires a clientId — column memory is scoped per client");
  const col = sheetSignals.columns.find((c) => c.index === chosenColumnIndex);
  if (!col) throw new Error(`rememberColumnField: no column with index ${chosenColumnIndex} on '${sheetName}'`);

  const { signature, headerSignature } = computeStructuralSignature(sheetSignals);
  const headerText = normalizeHeaderText(col.header);

  const existing = await store.findExact(clientId, signature);
  const blob = parseBlob(existing);
  blob[fieldName] = headerText;

  return store.remember(clientId, { structuralSignature: signature, headerSignature, sheetName, userProvidedLabel: JSON.stringify(blob) });
}

// Hand-edits or clears ONE field within an existing record's blob, by
// structuralSignature — used by the "review learned answers" UI, where the
// user is correcting what's already on file rather than answering a fresh
// candidate-picker prompt (so this takes free-text header text, not a
// column index — there may be no live sheet selected to pick a column
// from). newHeaderText === null removes just that field; if the blob ends
// up empty, the whole record is forgotten so no orphan zero-field row is
// left behind. A header that no longer matches any column on the CURRENT
// sheet isn't an error here — same fail-safe fallback as a normal miss:
// findStoredColumnIndex() simply won't find it next time, and resolution
// falls through to needs_input again rather than guessing.
async function updateStoredColumnField({ clientId, structuralSignature, fieldName, newHeaderText, store }) {
  if (!clientId) throw new Error("updateStoredColumnField requires a clientId");
  const existing = await store.findExact(clientId, structuralSignature);
  if (!existing) return null;

  const blob = parseBlob(existing);
  if (newHeaderText === null) {
    delete blob[fieldName];
  } else {
    if (!newHeaderText.trim()) throw new Error("updateStoredColumnField requires a non-empty newHeaderText (pass null to clear the field instead)");
    blob[fieldName] = normalizeHeaderText(newHeaderText);
  }

  if (Object.keys(blob).length === 0) {
    await store.forget(clientId, structuralSignature);
    return { deleted: true };
  }
  const updated = await store.updateLabel(clientId, structuralSignature, JSON.stringify(blob));
  return { deleted: false, record: updated };
}

// Deletes every field resolved for this sheet shape at once — the next
// scan that hits this exact shape gets asked about all of them again.
async function forgetColumnMemoryRecord({ clientId, structuralSignature, store }) {
  if (!clientId) throw new Error("forgetColumnMemoryRecord requires a clientId");
  return store.forget(clientId, structuralSignature);
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { resolveColumnField, rememberColumnField, findStoredColumnIndex, parseBlob, updateStoredColumnField, forgetColumnMemoryRecord };
}
