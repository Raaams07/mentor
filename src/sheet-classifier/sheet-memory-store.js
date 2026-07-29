/*
 * sheet-memory-store.js
 * ----------------------
 * MENTOR Phase 2, sheet memory — Node-side storage layer, used by the
 * regression-test / sheet-memory-test suites (and any future CLI tooling).
 * NOT used by the live add-in — see browser-sheet-memory-store.js for that;
 * Office Add-ins run in a browser sandbox with no filesystem access, so
 * this file's `fs`/`path` usage must never end up in the taskpane bundle.
 *
 * Stub implementation of what will eventually be a Supabase table:
 *
 *   sheet_memory(client_id, structural_signature, header_signature,
 *                sheet_name_at_creation, user_provided_label, created_at)
 *
 * All matching/record logic lives in sheet-memory-store-base.js and is
 * shared with browser-sheet-memory-store.js — this file only adds
 * filesystem persistence on top of it. Swapping either stub for the real
 * Supabase-backed store later means writing one new subclass of
 * BaseSheetMemoryStore — nothing else in the codebase needs to change.
 */

const fs = require("fs");
const path = require("path");
const { BaseSheetMemoryStore } = require("./sheet-memory-store-base.js");

const DEFAULT_STORE_PATH = path.join(__dirname, "data", "sheet-memory.local.json");

class JsonFileSheetMemoryStore extends BaseSheetMemoryStore {
  // filePath: null => pure in-memory, never touches disk (handy for tests).
  // Any other value => read/write that JSON file, persisting across restarts.
  constructor(options) {
    super();
    const opts = options || {};
    this.filePath = opts.filePath === undefined ? DEFAULT_STORE_PATH : opts.filePath;
    this.records = this._load();
  }

  _load() {
    if (!this.filePath) return [];
    try {
      const raw = fs.readFileSync(this.filePath, "utf8");
      return JSON.parse(raw);
    } catch (err) {
      if (err.code === "ENOENT") return []; // no store yet — first run
      throw err;
    }
  }

  _persist() {
    if (!this.filePath) return; // in-memory mode
    fs.mkdirSync(path.dirname(this.filePath), { recursive: true });
    fs.writeFileSync(this.filePath, JSON.stringify(this.records, null, 2), "utf8");
  }
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { JsonFileSheetMemoryStore, DEFAULT_STORE_PATH };
}
