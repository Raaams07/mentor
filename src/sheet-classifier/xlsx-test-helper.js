/*
 * xlsx-test-helper.js
 * --------------------
 * Shared by regression-test.js and sheet-memory-test.js: reads a real
 * .xlsx file into the { name, values, numberFormats } shape signal-extractor.js
 * expects, using the `xlsx` (SheetJS) devDependency. Test-only tooling —
 * the live add-in reads sheets via Office.js `range.values`, not this.
 */

const XLSX = require("xlsx");

function readWorkbookSheets(filePath) {
  const wb = XLSX.readFile(filePath, { cellDates: false, cellNF: true });

  return wb.SheetNames.map((name) => {
    const ws = wb.Sheets[name];
    const values = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });

    const ref = ws["!ref"];
    let numberFormats = null;
    if (ref) {
      const range = XLSX.utils.decode_range(ref);
      numberFormats = [];
      for (let r = range.s.r; r <= range.e.r; r++) {
        const row = [];
        for (let c = range.s.c; c <= range.e.c; c++) {
          const cell = ws[XLSX.utils.encode_cell({ r, c })];
          row.push(cell && cell.z ? cell.z : undefined);
        }
        numberFormats.push(row);
      }
    }

    return { name, values, numberFormats };
  });
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { readWorkbookSheets };
}
