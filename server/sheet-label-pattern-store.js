/*
 * sheet-label-pattern-store.js
 * -------------------------------
 * Tier 1 (shared software-pattern) SHEET-IDENTITY memory — the sheet-level
 * counterpart to column-pattern-store.js's field-level one. Same shape of
 * problem: "has ANY client, anywhere, told MENTOR what a sheet shaped
 * exactly like this is called before" — e.g. a Tally purchase-register
 * export has a distinctive column-shape fingerprint regardless of which
 * client's data fills it, so the label one client gave it ("Purchase
 * register (Books)") is a genuinely reusable cross-client fact, not
 * something specific to that client.
 *
 * Key is the sheet's OWN structural signature (structural-signature.js's
 * computeStructuralSignature().signature — the same coarse per-column
 * type-tag fingerprint sheet-memory.js's Tier 2 already keys on) — no
 * separate key-normalization step needed here the way column-pattern-
 * key.js was needed for columns, because a sheet's structural signature is
 * already exactly that: a coarse, order-sensitive-but-drift-resistant
 * fingerprint of the WHOLE sheet shape, not an unordered candidate set
 * that needs sorting/deduping first. Exact match only, no fuzzy — same
 * "the key is already coarse enough" reasoning column-pattern-store.js
 * documents for its own exact-match choice.
 *
 * Deliberately stores the label VERBATIM (trimmed, not normalized) —
 * unlike a column header (used only internally for matching, never shown
 * to a user as-is), a sheet label IS the exact text redisplayed to a
 * future client as "MENTOR recognizes this as: ...". Lowercasing/
 * stripping punctuation the way normalizeHeaderText() does for columns
 * would visibly degrade what gets shown, not just how it's matched.
 *
 * JSON-file-backed, same convention as column-pattern-store.js and
 * src/sheet-classifier/sheet-memory-store.js (filePath: null => pure
 * in-memory for tests; omitted => the real default path). No file-
 * locking/concurrency handling — same accepted trade-off as its column
 * counterpart: writes only happen when a human explicitly answers a
 * "what is this sheet?" prompt, so write volume/overlap risk is low.
 *
 * remember() is last-write-wins with no confidence tracking, by design —
 * same trade-off column-pattern-store.js accepts for the identical reason.
 */

const fs = require("fs");
const path = require("path");

const DEFAULT_STORE_PATH = path.join(__dirname, "data", "sheet-label-pattern-memory.json");

class SheetLabelPatternStore {
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

  // -> { found: false } | { found: true, label }
  lookup(structuralSignature) {
    const record = this.records[structuralSignature];
    return record ? { found: true, label: record.label } : { found: false };
  }

  remember(structuralSignature, label) {
    this.records[structuralSignature] = {
      structuralSignature,
      label, // verbatim — see docstring for why this is NOT normalized
      updated_at: new Date().toISOString(),
    };
    this._persist();
    return { ok: true };
  }
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { SheetLabelPatternStore, DEFAULT_STORE_PATH };
}
