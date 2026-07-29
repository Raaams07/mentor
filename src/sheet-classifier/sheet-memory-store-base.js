/*
 * sheet-memory-store-base.js
 * ----------------------------
 * Shared matching/record logic for every sheet-memory store implementation
 * (Node's JsonFileSheetMemoryStore, the browser's BrowserSheetMemoryStore,
 * and eventually a Supabase-backed one). Deliberately has NO platform-
 * specific code (no `fs`, no `localStorage`) — subclasses only need to
 * implement `_load()` and `_persist()`; findExact/findSimilar/remember/list
 * are identical everywhere and live here once, so the two fuzzy-match
 * thresholds can't drift out of sync between implementations.
 *
 * This file has no Node-only dependencies (only requires structural-
 * signature.js, which is itself pure JS) — safe to bundle into the browser
 * taskpane build.
 */

const { tagSequenceSimilarity, headerSequenceSimilarity } = require("./structural-signature.js");

// Below this many shared columns, comparing tag sequences isn't meaningful —
// skip the (relatively expensive) edit-distance computation entirely.
const MAX_COLUMN_COUNT_DRIFT = 3;
const FUZZY_SIMILARITY_THRESHOLD = 0.8;
// Calibrated against the two real cases that motivated this: the legitimate
// "client added one column" drift case scores ~83% header similarity (all
// other headers unchanged), while the Vendor-List-vs-Payroll-List collision
// scores ~25% (genuinely different wording throughout) — 0.5 sits with a
// wide margin on both sides of that gap.
const HEADER_SIMILARITY_THRESHOLD = 0.5;

class BaseSheetMemoryStore {
  constructor() {
    this.records = [];
  }

  // Subclasses override these two; everything else builds on `this.records`.
  _load() {
    return [];
  }
  _persist() {
    // no-op by default
  }

  async findExact(clientId, structuralSignature) {
    return this.records.find((r) => r.client_id === clientId && r.structural_signature === structuralSignature) || null;
  }

  // headerSignature is optional for backward compatibility, but omitting it
  // means header corroboration can't be checked — treated as a non-match
  // (fail closed) rather than silently falling back to structural-only
  // matching, since that's exactly the gap this two-signal check exists to close.
  async findSimilar(clientId, structuralSignature, headerSignature) {
    const [countStr, tagsStr] = structuralSignature.split(":");
    const columnCount = parseInt(countStr, 10);
    const tags = tagsStr ? tagsStr.split(",") : [];
    const headerTokens = headerSignature ? headerSignature.split("|").map((h) => (h === "∅" ? "" : h)) : null;

    let best = null;
    for (const record of this.records) {
      if (record.client_id !== clientId) continue;

      const [candCountStr, candTagsStr] = record.structural_signature.split(":");
      const candCount = parseInt(candCountStr, 10);
      if (Math.abs(candCount - columnCount) > MAX_COLUMN_COUNT_DRIFT) continue; // too different to bother comparing

      const candTags = candTagsStr ? candTagsStr.split(",") : [];
      const similarity = tagSequenceSimilarity(tags, candTags);
      if (similarity < FUZZY_SIMILARITY_THRESHOLD) continue;

      if (!headerTokens || !record.header_signature) continue; // can't corroborate — no match
      const candHeaderTokens = record.header_signature.split("|").map((h) => (h === "∅" ? "" : h));
      const headerSimilarity = headerSequenceSimilarity(headerTokens, candHeaderTokens);
      if (headerSimilarity < HEADER_SIMILARITY_THRESHOLD) continue;

      if (!best || similarity > best.similarity) {
        best = { record, similarity, headerSimilarity };
      }
    }
    return best;
  }

  async remember(clientId, { structuralSignature, headerSignature, sheetName, userProvidedLabel }) {
    // Re-labeling the same signature updates the existing record rather than
    // accumulating duplicates — the client's answer for "what is this shape"
    // is a single current fact, not a log of every time they were asked.
    const existing = await this.findExact(clientId, structuralSignature);
    if (existing) {
      existing.header_signature = headerSignature;
      existing.user_provided_label = userProvidedLabel;
      existing.sheet_name_at_creation = sheetName;
      this._persist();
      return existing;
    }

    const record = {
      client_id: clientId,
      structural_signature: structuralSignature,
      header_signature: headerSignature,
      sheet_name_at_creation: sheetName,
      user_provided_label: userProvidedLabel,
      created_at: new Date().toISOString(),
    };
    this.records.push(record);
    this._persist();
    return record;
  }

  async list(clientId) {
    return this.records.filter((r) => r.client_id === clientId);
  }
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { BaseSheetMemoryStore };
}
