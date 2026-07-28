/*
 * sheet-memory-store.js
 * ----------------------
 * MENTOR Phase 2, sheet memory — storage layer.
 *
 * Stub implementation of what will eventually be a Supabase table:
 *
 *   sheet_memory(client_id, structural_signature, sheet_name_at_creation,
 *                user_provided_label, created_at)
 *
 * Every store implementation — this stub, and later a real Supabase-backed
 * one — must expose exactly these four async methods:
 *
 *   findExact(clientId, structuralSignature)   -> record | null
 *   findSimilar(clientId, structuralSignature) -> { record, similarity } | null
 *   remember(clientId, structuralSignature, sheetName, userProvidedLabel) -> record
 *   list(clientId)                             -> record[]
 *
 * That's the whole contract. Swapping this stub for the real thing later
 * means writing one new file (e.g. supabase-sheet-memory-store.js) that
 * implements these same four methods against `supabase.from('sheet_memory')`
 * — nothing else in the codebase (sheet-memory.js, or whatever calls it)
 * needs to change.
 */

const fs = require("fs");
const path = require("path");
const { tagSequenceSimilarity } = require("./structural-signature.js");

const DEFAULT_STORE_PATH = path.join(__dirname, "data", "sheet-memory.local.json");

// Below this many shared columns, comparing tag sequences isn't meaningful —
// skip the (relatively expensive) edit-distance computation entirely.
const MAX_COLUMN_COUNT_DRIFT = 3;
const FUZZY_SIMILARITY_THRESHOLD = 0.8;

class JsonFileSheetMemoryStore {
  // filePath: null => pure in-memory, never touches disk (handy for tests).
  // Any other value => read/write that JSON file, persisting across restarts.
  constructor(options) {
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

  async findExact(clientId, structuralSignature) {
    return this.records.find((r) => r.client_id === clientId && r.structural_signature === structuralSignature) || null;
  }

  async findSimilar(clientId, structuralSignature) {
    const [countStr, tagsStr] = structuralSignature.split(":");
    const columnCount = parseInt(countStr, 10);
    const tags = tagsStr ? tagsStr.split(",") : [];

    let best = null;
    for (const record of this.records) {
      if (record.client_id !== clientId) continue;

      const [candCountStr, candTagsStr] = record.structural_signature.split(":");
      const candCount = parseInt(candCountStr, 10);
      if (Math.abs(candCount - columnCount) > MAX_COLUMN_COUNT_DRIFT) continue; // too different to bother comparing

      const candTags = candTagsStr ? candTagsStr.split(",") : [];
      const similarity = tagSequenceSimilarity(tags, candTags);
      if (similarity >= FUZZY_SIMILARITY_THRESHOLD && (!best || similarity > best.similarity)) {
        best = { record, similarity };
      }
    }
    return best;
  }

  async remember(clientId, structuralSignature, sheetName, userProvidedLabel) {
    // Re-labeling the same signature updates the existing record rather than
    // accumulating duplicates — the client's answer for "what is this shape"
    // is a single current fact, not a log of every time they were asked.
    const existing = await this.findExact(clientId, structuralSignature);
    if (existing) {
      existing.user_provided_label = userProvidedLabel;
      existing.sheet_name_at_creation = sheetName;
      this._persist();
      return existing;
    }

    const record = {
      client_id: clientId,
      structural_signature: structuralSignature,
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
  module.exports = { JsonFileSheetMemoryStore, DEFAULT_STORE_PATH };
}
