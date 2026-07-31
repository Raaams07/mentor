/*
 * gst-reconciliation.js
 * -----------------------
 * Top-level entry point: given a workbook's sheets (as {name, values,
 * numberFormats}), recognizes each sheet's GST role, identifies which
 * registered workflow (if any) the recognized sheets complete, and — for
 * now, only gstr2a_vs_books — runs that workflow's reconciler and returns
 * the proposed summary.
 *
 * Adding a new workflow later means adding one more `if` branch here that
 * calls its own reconciler with its own sheets — not restructuring this
 * function or anything upstream of it.
 */

const { extractSheetSignals } = require("../sheet-classifier/signal-extractor.js");
const { recognizeGstRole } = require("./gst-role-recognizer.js");
const { identifyGstWorkflow } = require("./gst-workflows.js");
const { reconcileGstr2aVsBooks } = require("./gstr2a-vs-books-reconciler.js");

function recognizeGstSheets(sheets) {
  // sheets: [{ name, values, numberFormats? }, ...]
  const results = {};
  for (const sheet of sheets) {
    const sheetSignals = extractSheetSignals(sheet.name, sheet.values, { numberFormats: sheet.numberFormats });
    const headerRowIndex = sheetSignals.headerRowIndex;
    const roleResult = recognizeGstRole(sheetSignals, sheet.values, headerRowIndex);
    results[sheet.name] = { ...roleResult, sheetSignals, values: sheet.values, headerRowIndex };
  }
  return results;
}

function summarizeRoleResults(roleResults) {
  const summary = {};
  for (const [sheetName, r] of Object.entries(roleResults)) {
    summary[sheetName] = { role: r.role, label: r.label, confidence: r.confidence, scores: r.scores, reason: r.reason };
  }
  return summary;
}

function runGstReconciliation(sheets, options) {
  const roleResults = recognizeGstSheets(sheets);
  const workflow = identifyGstWorkflow(roleResults);

  if (!workflow) {
    return { status: "no_workflow_recognized", roleResults: summarizeRoleResults(roleResults) };
  }

  if (workflow.workflowId === "gstr2a_vs_books") {
    const gstr2aSheet = workflow.sheets.gstr2a;
    const booksSheet = workflow.sheets.purchaseRegister;

    const summary = reconcileGstr2aVsBooks(
      { values: gstr2aSheet.result.values, headerRowIndex: gstr2aSheet.result.headerRowIndex, columns: gstr2aSheet.result.columns },
      { values: booksSheet.result.values, headerRowIndex: booksSheet.result.headerRowIndex, columns: booksSheet.result.columns },
      options
    );

    return {
      status: "reconciled",
      workflowId: workflow.workflowId,
      workflowLabel: workflow.label,
      sheets: { gstr2a: gstr2aSheet.sheetName, purchaseRegister: booksSheet.sheetName },
      summary,
      roleResults: summarizeRoleResults(roleResults),
    };
  }

  // Unreachable while GST_WORKFLOWS has only one entry — kept explicit so a
  // future workflow registered without a matching reconciler call here
  // fails loudly instead of silently doing nothing.
  throw new Error(`Recognized workflow '${workflow.workflowId}' has no reconciler wired up yet`);
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { recognizeGstSheets, runGstReconciliation };
}
