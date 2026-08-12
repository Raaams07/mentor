/* global Excel */

/*
 * mentor-workbook-column-memory-store.js
 * -----------------------------------------
 * Column-memory store backed by a dedicated worksheet inside the workbook
 * itself ("MENTOR Column Memory"), so "has this sheet shape's GST columns
 * already been confirmed?" survives task pane close/reopen and travels
 * with the file — same rationale and same mechanism as
 * mentor-workbook-sheet-memory-store.js (WorkbookSheetMemoryStore), just a
 * separate sheet/store instance so column resolutions and sheet-type
 * labels never collide.
 *
 * Reuses BaseSheetMemoryStore completely unmodified (see column-memory.js
 * for why: one record per (client_id, structural_signature), whose label
 * column holds a JSON blob of every field resolved for that shape) — this
 * class only supplies the Excel read/write plumbing, identical in
 * structure to WorkbookSheetMemoryStore.
 */

const { BaseSheetMemoryStore } = require("../sheet-classifier/sheet-memory-store-base.js");

const SHEET_NAME = "MENTOR Column Memory";
const HEADER_ROW = ["Client ID", "Structural Signature", "Header Signature", "Sheet Name", "Resolved Fields (JSON)", "Created At"];

class WorkbookColumnMemoryStore extends BaseSheetMemoryStore {
  async _refreshFromExcel() {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAME);
      sheet.load("isNullObject");
      await context.sync();

      if (sheet.isNullObject) {
        console.log("MENTOR workbook column-memory: '" + SHEET_NAME + "' sheet does not exist yet — 0 records");
        this.records = [];
        return;
      }

      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("values, isNullObject");
      await context.sync();

      if (usedRange.isNullObject) {
        console.log("MENTOR workbook column-memory: '" + SHEET_NAME + "' sheet exists but is empty — 0 records");
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
      console.log("MENTOR workbook column-memory: loaded " + this.records.length + " record(s) from '" + SHEET_NAME + "'");
    });
  }

  async _persist() {
    console.log("MENTOR workbook column-memory: persisting " + this.records.length + " record(s) to '" + SHEET_NAME + "'");
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAME);
      sheet.load("isNullObject");
      await context.sync();

      if (sheet.isNullObject) {
        console.log("MENTOR workbook column-memory: creating '" + SHEET_NAME + "' sheet");
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
      console.log("MENTOR workbook column-memory: persist complete, wrote " + allRows.length + " row(s) (incl. header) to '" + SHEET_NAME + "'");
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
  module.exports = { WorkbookColumnMemoryStore, SHEET_NAME };
}
