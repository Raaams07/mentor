/*
 * column-pattern-store.js
 * -------------------------
 * Tier 1 (shared software-pattern) column-resolution memory — a flat
 * key/value store, NOT a BaseSheetMemoryStore subclass: this tier has no
 * per-client scoping and no fuzzy tag-sequence matching (the key,
 * column-pattern-key.js's computeColumnPatternKey, is already coarse —
 * normalized/sorted/deduped header text — so exact match is sufficient;
 * see column-pattern-key.js for why that coarseness is the whole point).
 *
 * JSON-file-backed, mirrors src/sheet-classifier/sheet-memory-store.js's
 * JsonFileSheetMemoryStore convention (filePath: null => pure in-memory,
 * used by tests; omitted => the real default path; anything else => that
 * path). No file-locking/concurrency handling — accepted simplification
 * for a lightweight local dev-tool store: writes only happen when a human
 * explicitly answers an ambiguous-column prompt, so write volume/overlap
 * risk is low, same trade-off spirit as this codebase's other local
 * single-file JSON stores.
 *
 * remember() is last-write-wins with no confidence tracking, by design —
 * one user's mistaken answer becomes every future user's silent
 * auto-resolution until someone else happens to hit the identical pattern
 * and overwrites it. Accepted trade-off for a store shared only across
 * users of one running local proxy (a machine, or a small firm's shared
 * server) — a future confidence-count enhancement would attach here.
 */

const fs = require("fs");
const path = require("path");
const { computeColumnPatternKey, normalizeCandidateHeaders } = require("../src/gst-reconciliation/column-pattern-key.js");
const { normalizeHeaderText } = require("../src/sheet-classifier/structural-signature.js");

const DEFAULT_STORE_PATH = path.join(__dirname, "data", "column-pattern-memory.json");

class ColumnPatternStore {
  // filePath: null => pure in-memory, never touches disk (handy for tests).
  // undefined => DEFAULT_STORE_PATH. Any other value => read/write that file.
  constructor(options) {
    const opts = options || {};
    this.filePath = opts.filePath === undefined ? DEFAULT_STORE_PATH : opts.filePath;
    this.records = this._load();
  }

  _load() {
    if (!this.filePath) return {};
    try {
      return JSON.parse(fs.readFileSync(this.filePath, "utf8"));
    } catch (err) {
      if (err.code === "ENOENT") return {}; // no store yet — first run
      throw err;
    }
  }

  _persist() {
    if (!this.filePath) return; // in-memory mode
    fs.mkdirSync(path.dirname(this.filePath), { recursive: true });
    fs.writeFileSync(this.filePath, JSON.stringify(this.records, null, 2), "utf8");
  }

  // -> { found: false } | { found: true, chosenHeader }
  lookup(fieldName, candidateHeaders) {
    const key = computeColumnPatternKey(fieldName, candidateHeaders);
    const record = this.records[key];
    return record ? { found: true, chosenHeader: record.chosenHeader } : { found: false };
  }

  remember(fieldName, candidateHeaders, chosenHeader) {
    const key = computeColumnPatternKey(fieldName, candidateHeaders);
    this.records[key] = {
      fieldName,
      candidateHeaders: normalizeCandidateHeaders(candidateHeaders), // stored normalized — never raw/original casing or client data
      chosenHeader: normalizeHeaderText(chosenHeader),
      updated_at: new Date().toISOString(),
    };
    this._persist();
    return { ok: true };
  }
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { ColumnPatternStore, DEFAULT_STORE_PATH };
}
