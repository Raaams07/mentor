/* global Excel */

/*
 * mentor-gst-summary-explain.js
 * ---------------------------------
 * Part 3 — extends point-and-explain to "GST - Summary": for every total
 * MENTOR reports (2A total, Books total, category subtotals), makes it
 * inspectable which resolved column(s) fed it. Reuses the EXACT SAME
 * interaction/render function every other GST sheet's point-and-explain
 * already uses (mentorRenderGstExplainHtml in mentor-gst-reconciliation-
 * ui.js) — this file only produces the {ruleLabel, whichRuleFired,
 * legalCitation} shape that function already knows how to render, just
 * pointed at column identity instead of legal citation. No new UI.
 *
 * The resolved `columns` object only exists in memory during one
 * generation run — it's never persisted anywhere queryable later (unlike
 * Wrong Head/RCM/etc. rows, which copy source row data onto their own
 * sheet and stay self-explanatory from a plain read). Rather than persist
 * a new snapshot, this file RE-DERIVES live on every click: the 2A/Books
 * source sheets are still in the workbook, and Tier 2/Tier 1 memory
 * (column-memory.js / the shared proxy) already hold the answer — re-
 * running recognition + resolution is cheap (small sheets, the same cost
 * every other point-and-explain already pays reading a sheet fresh on
 * selection) and always reflects CURRENT truth, not a stale generation-
 * time snapshot.
 */

const { extractSheetSignals } = require("../sheet-classifier/signal-extractor.js");
const { identifyGstColumns } = require("../gst-reconciliation/gst-column-identifier.js");
const { GST_SUMMARY_ROW_TRACEABILITY } = require("./gst-report-writer.js");
const { mentorResolveGstColumns } = require("./mentor-column-memory-ui.js");

// Column index (0-based) -> spreadsheet letter (0->A, 25->Z, 26->AA, ...).
// No existing helper in this codebase for this — the `xlsx` package's
// letter utilities are a Node/test-only devDependency (see xlsx-test-
// helper.js's own docstring), not appropriate to pull into the client
// bundle for one small conversion.
function columnIndexToLetter(index) {
  let n = index + 1;
  let letters = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

async function readSheetForResolution(context, sheetName) {
  const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
  sheet.load("isNullObject");
  await context.sync();
  if (sheet.isNullObject) return null;

  const usedRange = sheet.getUsedRangeOrNullObject();
  usedRange.load("values, numberFormat, isNullObject");
  await context.sync();
  if (usedRange.isNullObject) return null;

  const sheetSignals = extractSheetSignals(sheetName, usedRange.values, { numberFormats: usedRange.numberFormat });
  const columns = identifyGstColumns(sheetSignals, usedRange.values, sheetSignals.headerRowIndex);
  return { sheetName, result: { sheetSignals, columns } };
}

// rowLabel: the exact trimmed text from column A of the selected Summary
// row. gstr2aSheetName/booksSheetName: read by the caller off the Summary
// sheet's own "2A sheet"/"Books sheet" rows. Returns the same
// {ruleLabel, whichRuleFired, legalCitation} shape explainGstRow()
// produces for every other sheet — deliberately WITHOUT reviewerStatus/
// reviewerNote (a Summary total isn't a reviewable flagged row;
// mentorRenderGstExplainHtml is written to omit that section when those
// are left undefined). Returns null if this row isn't a traceable total
// (caller falls back to the existing static message) or if the 2A/Books
// sheets can't currently be read (renamed/deleted since generation).
async function mentorExplainGstSummaryRow(context, rowLabel, gstr2aSheetName, booksSheetName) {
  const descriptor = GST_SUMMARY_ROW_TRACEABILITY[rowLabel];
  if (!descriptor) return null;

  if (descriptor.derivedFrom) {
    return {
      ruleLabel: rowLabel,
      whichRuleFired: "Computed from the lines above, not directly from a single source column.",
      legalCitation: "See: " + descriptor.derivedFrom.join(", "),
    };
  }

  const [gstr2aSheet, booksSheet] = await Promise.all([readSheetForResolution(context, gstr2aSheetName), readSheetForResolution(context, booksSheetName)]);
  if (!gstr2aSheet || !booksSheet) {
    return {
      ruleLabel: rowLabel,
      whichRuleFired: "Can't currently trace this — '" + gstr2aSheetName + "' or '" + booksSheetName + "' no longer exists or is empty.",
      legalCitation: "Regenerate the review sheets to refresh this.",
    };
  }

  context.workbook.load("name");
  await context.sync();
  const clientId = "workbook:" + context.workbook.name;

  // Read-only: resolveColumnField/lookupColumnPattern never write, and
  // nothing here ever queues or shows a prompt — a field that comes back
  // unresolved is reported as such below, not blocked on or guessed at.
  const resolution = await mentorResolveGstColumns({ clientId, gstr2aSheet, booksSheet });

  const sheetsInScope = [];
  if (descriptor.sheet === "gstr2a" || descriptor.sheet === "both") sheetsInScope.push({ role: "gstr2a", sheet: gstr2aSheet });
  if (descriptor.sheet === "books" || descriptor.sheet === "both") sheetsInScope.push({ role: "books", sheet: booksSheet });

  const refs = [];
  for (const { role, sheet } of sheetsInScope) {
    for (const field of descriptor.fields) {
      const resolvedIdx = resolution.resolvedColumns[role][field];
      // resolvedColumns wins when the field was ambiguous/missing and got
      // resolved via Tier 2 or Tier 1 memory; otherwise it was never "in
      // play" at all, so the sheet's own plain (unambiguous) first-guess
      // is already correct.
      const idx = resolvedIdx !== undefined ? resolvedIdx : sheet.result.columns[field];
      if (idx === null || idx === undefined) {
        refs.push("'" + sheet.sheetName + "'." + field + ": not currently resolved");
      } else {
        const header = sheet.result.sheetSignals.columns[idx].header;
        refs.push("'" + sheet.sheetName + "'!" + columnIndexToLetter(idx) + " (header: \"" + header + "\")");
      }
    }
  }

  return {
    ruleLabel: rowLabel,
    whichRuleFired: refs.length + " column(s) feed this figure:",
    legalCitation: refs.join("; "),
  };
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { mentorExplainGstSummaryRow, columnIndexToLetter };
}
