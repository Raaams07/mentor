/*
 * sheet-memory.js
 * ---------------
 * MENTOR Phase 2, sheet memory — orchestration.
 *
 * Decides what should happen for a given sheet: if the rule-based classifier
 * (Step 2) already recognized it, there's nothing to do here. If it came
 * back "unknown", check whether this client has told MENTOR what this exact
 * shape is before (structural-signature.js + sheet-memory-store.js) — and
 * if not, hand back a ready-to-render prompt instead of guessing or staying
 * silent.
 *
 * This file has no Office.js/DOM dependency on purpose: it returns plain
 * data (a "prompt object" or a "remembered label" object) that a taskpane
 * UI layer can render however it likes. That wiring is a separate step.
 */

const { computeStructuralSignature } = require("./structural-signature.js");

async function resolveSheetLabel({ clientId, sheetName, sheetSignals, classification, store }) {
  if (!clientId) throw new Error("resolveSheetLabel requires a clientId — sheet memory is scoped per client");

  if (classification && classification.type && classification.type !== "unknown") {
    return {
      status: "classified",
      sheetName,
      type: classification.type,
      label: classification.displayLabel,
    };
  }

  const { signature, headerSignature } = computeStructuralSignature(sheetSignals);

  const exact = await store.findExact(clientId, signature);
  if (exact) {
    return {
      status: "remembered",
      sheetName,
      label: exact.user_provided_label,
      structuralSignature: signature,
      matchedVia: "exact",
    };
  }

  const fuzzy = await store.findSimilar(clientId, signature, headerSignature);
  if (fuzzy) {
    return {
      status: "remembered",
      sheetName,
      label: fuzzy.record.user_provided_label,
      structuralSignature: signature,
      matchedVia: "fuzzy",
      similarity: fuzzy.similarity,
      headerSimilarity: fuzzy.headerSimilarity,
    };
  }

  // Nothing on file for this shape yet — the UI layer renders `prompt` next
  // to a short text input, and calls rememberSheetLabel() with whatever the
  // user types, passing this same structuralSignature back unchanged.
  return {
    status: "needs_input",
    sheetName,
    structuralSignature: signature,
    prompt: `I don't recognize this sheet ('${sheetName}') — what is it?`,
  };
}

async function rememberSheetLabel({ clientId, sheetName, sheetSignals, userProvidedLabel, store }) {
  if (!clientId) throw new Error("rememberSheetLabel requires a clientId — sheet memory is scoped per client");
  if (!userProvidedLabel || !userProvidedLabel.trim()) throw new Error("rememberSheetLabel requires a non-empty userProvidedLabel");

  const { signature, headerSignature } = computeStructuralSignature(sheetSignals);
  return store.remember(clientId, { structuralSignature: signature, headerSignature, sheetName, userProvidedLabel: userProvidedLabel.trim() });
}

// Corrects an EXISTING record's label directly, by structuralSignature —
// used by the "review learned answers" UI, where the user is fixing what's
// already on file (e.g. a mislabeled sheet) rather than answering a fresh
// prompt. Unlike rememberSheetLabel, this doesn't need live sheetSignals
// (the sheet in question may not even be selected, or may have been
// renamed/removed since it was learned) — it operates purely on the stored
// record. Returns null if nothing was on file for that signature.
async function updateStoredSheetLabel({ clientId, structuralSignature, newLabel, store }) {
  if (!clientId) throw new Error("updateStoredSheetLabel requires a clientId");
  if (!newLabel || !newLabel.trim()) throw new Error("updateStoredSheetLabel requires a non-empty newLabel");
  return store.updateLabel(clientId, structuralSignature, newLabel.trim());
}

// Deletes a sheet-identity record outright — the next scan that hits this
// exact shape gets asked again instead of silently reusing what was here.
async function forgetSheetLabel({ clientId, structuralSignature, store }) {
  if (!clientId) throw new Error("forgetSheetLabel requires a clientId");
  return store.forget(clientId, structuralSignature);
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { resolveSheetLabel, rememberSheetLabel, updateStoredSheetLabel, forgetSheetLabel };
}
