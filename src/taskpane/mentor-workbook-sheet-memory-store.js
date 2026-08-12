/* global Excel */

/*
 * mentor-workbook-sheet-memory-store.js
 * ----------------------------------------
 * Sheet-memory store backed by a dedicated worksheet inside the workbook
 * itself ("MENTOR Sheet Memory"), so "has this sheet already been
 * classified?" survives task pane close/reopen and travels with the file —
 * unlike localStorage, which turned out NOT to reliably persist in this
 * hosting context (Office's WebView2 task pane host uses multiple separate
 * storage profiles; a write in one session can land in a different profile
 * than the next session reads from).
 *
 * Deliberately a SEPARATE sheet from "MENTOR Audit Log", not a new column
 * on it: the Audit Log is a human-readable event history (one entry per
 * action, in prose); this is a structured lookup table the program re-reads
 * on every scan. Mixing them would mean parsing free-text descriptions back
 * into structured data on every classification check.
 *
 * Re-reads the sheet fresh before every operation rather than trusting an
 * in-memory cache from construction time — this sheet is meant to be THE
 * authoritative source of truth, so a stale local copy defeats the purpose.
 * Persists by rewriting the whole sheet on every write, same approach the
 * Node JSON-file store uses (record counts here are small — tens of shapes
 * per client, not thousands — so this stays cheap).
 */

const { BaseSheetMemoryStore } = require("../sheet-classifier/sheet-memory-store-base.js");

const SHEET_NAME = "MENTOR Sheet Memory";
const HEADER_ROW = ["Client ID", "Structural Signature", "Header Signature", "Sheet Name", "Label", "Created At"];

class WorkbookSheetMemoryStore extends BaseSheetMemoryStore {
  async _refreshFromExcel() {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAME);
      sheet.load("isNullObject");
      await context.sync();

      if (sheet.isNullObject) {
        // Sheet doesn't exist yet (first run, or removed via Undo) — no
        // records, not an error. Chaining getUsedRangeOrNullObject() off a
        // null sheet object here would throw at sync time instead of
        // resolving to another null object, so this check has to happen
        // (and sync) BEFORE going any further.
        console.log("MENTOR workbook sheet-memory: '" + SHEET_NAME + "' sheet does not exist yet — 0 records");
        this.records = [];
        return;
      }

      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("values, isNullObject");
      await context.sync();

      if (usedRange.isNullObject) {
        console.log("MENTOR workbook sheet-memory: '" + SHEET_NAME + "' sheet exists but is empty — 0 records");
        this.records = [];
        return;
      }

      this.records = usedRange.values
        .slice(1) // header row
        .filter((row) => row[0]) // skip blank trailing rows
        .map((row) => ({
          client_id: row[0],
          structural_signature: row[1],
          header_signature: row[2] || "",
          sheet_name_at_creation: row[3],
          user_provided_label: row[4],
          created_at: row[5],
        }));
      console.log("MENTOR workbook sheet-memory: loaded " + this.records.length + " record(s) from '" + SHEET_NAME + "'");
    });
  }

  async _persist() {
    console.log("MENTOR workbook sheet-memory: persisting " + this.records.length + " record(s) to '" + SHEET_NAME + "'");
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAME);
      sheet.load("isNullObject");
      await context.sync();

      if (sheet.isNullObject) {
        console.log("MENTOR workbook sheet-memory: creating '" + SHEET_NAME + "' sheet");
        sheet = context.workbook.worksheets.add(SHEET_NAME);
        await context.sync();
      }

      // Pure internal machine state, not meant for users to browse — set on
      // every persist (not just at creation) so it can't end up visible
      // again if something else unhides it.
      sheet.visibility = Excel.SheetVisibility.hidden;

      const existingUsedRange = sheet.getUsedRangeOrNullObject();
      existingUsedRange.load("isNullObject");
      await context.sync();
      if (!existingUsedRange.isNullObject) {
        existingUsedRange.clear();
        await context.sync();
      }

      const rows = this.records.map((r) => [
        r.client_id,
        r.structural_signature,
        r.header_signature || "",
        r.sheet_name_at_creation,
        r.user_provided_label,
        r.created_at,
      ]);
      const allRows = [HEADER_ROW, ...rows];

      const range = sheet.getRangeByIndexes(0, 0, allRows.length, HEADER_ROW.length);
      range.values = allRows;
      sheet.getRange("A1:F1").format.font.bold = true;

      await context.sync();
      console.log("MENTOR workbook sheet-memory: persist complete, wrote " + allRows.length + " row(s) (incl. header) to '" + SHEET_NAME + "'");
    });
  }

  async findExact(clientId, structuralSignature) {
    await this._refreshFromExcel();
    return super.findExact(clientId, structuralSignature);
  }

  async findSimilar(clientId, structuralSignature, headerSignature) {
    await this._refreshFromExcel();
    return super.findSimilar(clientId, structuralSignature, headerSignature);
  }

  async remember(clientId, info) {
    await this._refreshFromExcel();
    return super.remember(clientId, info);
  }

  async list(clientId) {
    await this._refreshFromExcel();
    return super.list(clientId);
  }

  async forget(clientId, structuralSignature) {
    await this._refreshFromExcel();
    return super.forget(clientId, structuralSignature);
  }

  async updateLabel(clientId, structuralSignature, newLabel) {
    await this._refreshFromExcel();
    return super.updateLabel(clientId, structuralSignature, newLabel);
  }
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { WorkbookSheetMemoryStore, SHEET_NAME };
}
