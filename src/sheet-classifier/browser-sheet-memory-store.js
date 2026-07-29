/*
 * browser-sheet-memory-store.js
 * -------------------------------
 * MENTOR Phase 2, sheet memory — browser-side storage layer, used by the
 * live Excel add-in (taskpane.js). Office Add-ins run in a browser sandbox
 * with no filesystem access, so this uses localStorage instead of the JSON
 * file the Node test suite uses (sheet-memory-store.js) — same shared
 * matching logic either way (sheet-memory-store-base.js), just different
 * persistence.
 *
 * Also a stub, same as the Node one: real client data belongs in Supabase,
 * not a single browser's localStorage, which doesn't sync across machines
 * or survive a cleared profile/different device. Swapping this out later
 * means writing one new BaseSheetMemoryStore subclass — nothing else
 * (sheet-memory.js, taskpane.js) needs to change.
 */

const { BaseSheetMemoryStore } = require("./sheet-memory-store-base.js");

const DEFAULT_STORAGE_KEY = "mentorSheetMemory";

class BrowserSheetMemoryStore extends BaseSheetMemoryStore {
  constructor(options) {
    super();
    const opts = options || {};
    this.storageKey = opts.storageKey || DEFAULT_STORAGE_KEY;
    this.records = this._load();
  }

  _load() {
    try {
      const raw = window.localStorage.getItem(this.storageKey);
      return raw ? JSON.parse(raw) : [];
    } catch (err) {
      console.error("MENTOR sheet-memory: failed to read localStorage, starting empty", err);
      return [];
    }
  }

  _persist() {
    try {
      window.localStorage.setItem(this.storageKey, JSON.stringify(this.records));
    } catch (err) {
      console.error("MENTOR sheet-memory: failed to write localStorage", err);
    }
  }
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { BrowserSheetMemoryStore };
}
