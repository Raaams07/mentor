/* global Excel */

/*
 * mentor-gst-reconciliation-ui.js
 * -----------------------------------
 * Live task-pane wiring for GST reconciliation. Bridges the already-
 * validated src/gst-reconciliation/*.js pipeline (Step 1: GSTR-2A vs Books
 * vendor match/mismatch; Step 2: Wrong Head / RCM / Ineligible ITC /
 * Duplicate Invoices) to the actual add-in: detects GST-shaped sheets,
 * computes the full reconciliation read-only, and — following MENTOR's
 * "propose, don't auto-execute" rule — shows a summary card and only
 * writes the output sheets (via gst-report-writer.js) once the user
 * explicitly accepts.
 *
 * Nothing in src/gst-reconciliation/ is modified or reimplemented here —
 * this file only calls those pure functions and drives the UI/Excel-write
 * side of surfacing what they already compute.
 *
 * Mirrors mentor-sheet-memory-ui.js's structure: runs at workbook open and
 * whenever a sheet is added (a workbook's fundamental shape doesn't change
 * cell-by-cell), with its own dedicated prompt container rather than the
 * active-sheet-scoped suggestion badge, since "is there GST data in this
 * workbook" is a workbook-wide condition, not tied to whichever sheet
 * happens to be selected.
 */

const { extractSheetSignals } = require("../sheet-classifier/signal-extractor.js");
const { computeGstSourceDataHash } = require("./gst-source-data-hash.js");
const { recognizeGstSheets } = require("../gst-reconciliation/gst-reconciliation.js");
const { identifyGstWorkflow } = require("../gst-reconciliation/gst-workflows.js");
const { reconcileGstr2aVsBooks } = require("../gst-reconciliation/gstr2a-vs-books-reconciler.js");
const { detectWrongHead } = require("../gst-reconciliation/wrong-head-detector.js");
const { detectCrossSheetWrongHead } = require("../gst-reconciliation/cross-sheet-wrong-head-detector.js");
const { detectInvoiceLevelExtras } = require("../gst-reconciliation/extra-invoice-detector.js");
const { detectRcmForSheet } = require("../gst-reconciliation/rcm-detector.js");
const { detectIneligibleItcForSheet } = require("../gst-reconciliation/ineligible-itc-detector.js");
const { findDuplicateInvoices } = require("../gst-reconciliation/duplicate-invoice-detector.js");
const { detectRateMismatch } = require("../gst-reconciliation/rate-mismatch-detector.js");
const { writeGstReconciliationReport, GST_REPORT_SHEET_NAMES, GST_EXPLAIN_SHEET_NAMES, explainGstRow, appendIneligibleItcVendorSuggestions, readGstTopIssuesFromExistingSheets } = require("./gst-report-writer.js");
const { extractDistinctVendors } = require("./ineligible-itc-vendor-extraction.js");
const { MENTOR_OWNED_SHEET_NAMES } = require("./mentor-sheet-memory-ui.js");
const { mentorResolveGstColumns, mentorShowColumnMemoryQueue, mentorHideColumnMemoryQueue, COLUMN_MEMORY_SHEET_NAME } = require("./mentor-column-memory-ui.js");
const { mentorExplainGstSummaryRow } = require("./mentor-gst-summary-explain.js");

// Local proxy server (server/mentor-suggestions-proxy.js) — holds the
// Anthropic API key server-side so it's never embedded in this bundled
// client JS. Must be running separately (npm run gst-suggestions-proxy).
// If MENTOR_PROXY_PORT is ever changed on the server side, this needs to
// be updated too — there's no way for build-time client code to read that
// env var automatically.
const MENTOR_SUGGESTIONS_PROXY_URL = "http://localhost:3001/api/ineligible-itc-vendor-suggestions";

// MENTOR_OWNED_SHEET_NAMES is a shared, mutable Set — every module that
// requires mentor-sheet-memory-ui.js gets the same instance (module
// caching), so adding the GST report sheet names here makes BOTH the
// sheet-memory scanner (which already excludes MENTOR_OWNED_SHEET_NAMES)
// and taskpane.js's own onSheetChanged exclusion checks skip these sheets
// too, without editing either of those files.
GST_REPORT_SHEET_NAMES.forEach((name) => MENTOR_OWNED_SHEET_NAMES.add(name));
MENTOR_OWNED_SHEET_NAMES.add(COLUMN_MEMORY_SHEET_NAME);

// Placeholder client identifier, scoped to this workbook file — same
// mechanism and same rationale as mentor-sheet-memory-ui.js's own
// mentorGetSheetMemoryClientId(), duplicated here (not imported) since
// that function isn't exported and column-memory's clientId scoping is
// conceptually independent of sheet-memory's, even though the value is
// currently identical.
function mentorGetColumnMemoryClientId(workbookName) {
  return "workbook:" + workbookName;
}

let mentorGstLogAction = null;
let mentorGstAppendAuditLogRow = null;
let mentorPendingGstProposal = null; // { workbookKey, sheetNames, comparisonSummary, wrongHeadResult, rcmBySource, itcBySource, dupBySource, counts }
const mentorGstHandledThisSession = new Set(); // "gstr2aSheetName::booksSheetName" — accepted or dismissed

// Top Issues state for the post-generation confirmation card — kept at
// module scope (not just inside the accept handler) so the "Show all"
// toggle and each issue's click handler can re-render/act without needing
// to re-run the whole accept flow.
let mentorGstTopIssues = []; // ranked list from the most recent generate/regenerate
let mentorGstTopIssuesShowAll = false;
let mentorGstConfirmationPreamble = ""; // the "Reconciliation complete..." line + preservation note, rendered above the list

// Ineligible ITC vendor-name suggestion feature state — separate,
// optional, only ever populated when the description-based rule found
// nothing to check (see mentorAcceptGstReconciliation). null vendors list
// means the button doesn't show at all.
let mentorGstIneligibleItcVendors = null; // [{gstin, vendorName}, ...] computed once at generation time
let mentorGstIneligibleItcSuggestionState = "idle"; // "idle" | "loading" | "done" | "error"
let mentorGstIneligibleItcSuggestionResult = null; // {flaggedCount, skippedCount, totalChecked}
let mentorGstIneligibleItcSuggestionError = "";

function mentorGstWorkbookKey(gstr2aName, booksName) {
  return gstr2aName + "::" + booksName;
}

// context.workbook.settings is persisted IN the document itself (per
// file, per add-in) — survives close/reopen, unlike any in-memory Set.
// Storing the source-data hash here (not on a visible sheet) is what
// lets MENTOR tell "have 2A/Books actually changed since I last
// generated these sheets" across sessions, not just within one.
function mentorGstDataHashSettingKey(gstr2aName, booksName) {
  return "MENTOR_GST_DATA_HASH::" + mentorGstWorkbookKey(gstr2aName, booksName);
}

function mentorGstDebounce(fn, ms) {
  let timer = null;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), ms);
  };
}

async function mentorGstReconciliationInit(logAction, appendAuditLogRow) {
  mentorGstLogAction = logAction;
  mentorGstAppendAuditLogRow = appendAuditLogRow;

  try {
    await Excel.run(async (context) => {
      context.workbook.worksheets.onAdded.add(mentorScanForGstReconciliation);
      // Office.js allows multiple independent handlers on the same event —
      // this doesn't replace or conflict with taskpane.js's own
      // onSelectionChanged listener (that one is scoped to the P&L demo
      // sheet); both simply fire on every selection change.
      context.workbook.worksheets.onSelectionChanged.add(mentorGstDebounce(mentorGstOnSelectionChanged, 120));
      await context.sync();
    });
  } catch (error) {
    console.error("MENTOR GST reconciliation: failed to register onAdded/onSelectionChanged listeners", error);
  }

  await mentorScanForGstReconciliation();
}

async function mentorScanForGstReconciliation() {
  try {
    await Excel.run(async (context) => {
      context.workbook.load("name");
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      // Whether output sheets already exist only changes the PROPOSAL'S
      // WORDING (Generate vs Regenerate) below — it no longer skips the
      // scan outright. Skipping unconditionally meant a workbook that
      // already had generated sheets could never pick up a later fix to
      // this pipeline (freeze panes, formatting, a corrected column) short
      // of the user manually deleting 'GST - Summary' first, which is a
      // confusing dead end, not a feature.
      const alreadyGenerated = sheets.items.some((s) => s.name === "GST - Summary");

      const rawSheets = [];
      for (const s of sheets.items) {
        if (MENTOR_OWNED_SHEET_NAMES.has(s.name)) continue;
        try {
          const usedRange = s.getUsedRangeOrNullObject();
          usedRange.load("values, numberFormat, isNullObject");
          await context.sync();
          if (usedRange.isNullObject) continue;
          rawSheets.push({ name: s.name, values: usedRange.values, numberFormats: usedRange.numberFormat });
        } catch (sheetError) {
          console.error("MENTOR GST reconciliation: failed to read sheet '" + s.name + "' — skipping it this scan", sheetError);
        }
      }

      const roleResults = recognizeGstSheets(rawSheets);
      for (const [name, r] of Object.entries(roleResults)) {
        console.log("MENTOR GST reconciliation: '" + name + "' -> role: " + r.role + " (confidence: " + r.confidence + ")");
      }

      const workflow = identifyGstWorkflow(roleResults);
      if (!workflow || workflow.workflowId !== "gstr2a_vs_books") {
        console.log("MENTOR GST reconciliation: no complete workflow recognized this scan");
        return;
      }

      const gstr2aSheet = workflow.sheets.gstr2a;
      const booksSheet = workflow.sheets.purchaseRegister;
      const key = mentorGstWorkbookKey(gstr2aSheet.sheetName, booksSheet.sheetName);
      if (mentorGstHandledThisSession.has(key)) {
        console.log("MENTOR GST reconciliation: '" + key + "' already accepted/dismissed this session — skipping");
        return;
      }

      // Has the underlying 2A/Books data actually changed since MENTOR
      // last generated these sheets? If the output sheets still exist AND
      // the data hasn't moved, stay completely silent — no proposal card,
      // no re-scan noise, no Audit Log entry. This is what makes a plain
      // Excel close/reopen with zero edits a no-op instead of re-running
      // the whole pipeline and re-prompting every single session.
      const hashSettingKey = mentorGstDataHashSettingKey(gstr2aSheet.sheetName, booksSheet.sheetName);
      const currentDataHash = computeGstSourceDataHash(gstr2aSheet.result.values, booksSheet.result.values);
      const hashSetting = context.workbook.settings.getItemOrNullObject(hashSettingKey);
      hashSetting.load("value");
      await context.sync();
      const storedDataHash = hashSetting.isNullObject ? null : hashSetting.value;

      if (alreadyGenerated && storedDataHash === currentDataHash) {
        // Silent for the REGENERATE PROMPT specifically — nothing needs
        // regenerating. The persistent Top Issues/summary card is a
        // SEPARATE concern from that prompt and should still show
        // whenever generated sheets exist, read directly off those
        // sheets (no re-run, no write, no Audit Log entry) rather than
        // tied to the accept flow the way it used to be — that coupling
        // was the actual bug: the "stay silent" fix suppressed the
        // summary along with the prompt, since both used to only ever
        // get populated as a side effect of clicking Generate/Regenerate.
        console.log("MENTOR GST reconciliation: '" + key + "' — source data unchanged since the last generation; regenerate prompt stays silent, summary still shown");
        const quietProposal = mentorComputeGstReconciliation(gstr2aSheet, booksSheet);
        await mentorShowPersistentGstSummary(context, quietProposal, null);
        return;
      }

      console.log(
        "MENTOR GST reconciliation: workflow recognized — 2A='" + gstr2aSheet.sheetName + "', Books='" + booksSheet.sheetName + "'. " +
        (storedDataHash === null ? "No prior generation recorded." : storedDataHash === currentDataHash ? "Data unchanged, but output sheets are missing." : "Source data has changed since the last generation.") +
        " Resolving columns..."
      );

      // Every field a current or future rule depends on gets resolved here
      // — from column memory where possible, otherwise queued as a
      // blocking prompt. No Generate/Regenerate proposal is computed (let
      // alone shown) until every ambiguous/missing-required field on BOTH
      // sheets is resolved — this is what stops a repeat of the Tally CGST
      // bug from ever reaching a proposal silently.
      const clientId = mentorGetColumnMemoryClientId(context.workbook.name);
      const resolution = await mentorResolveGstColumns({ clientId, gstr2aSheet, booksSheet });

      if (resolution.blocked) {
        mentorHideGstProposalCard();
        mentorPendingGstProposal = null;
        console.log(
          "MENTOR GST reconciliation: '" + key + "' blocked — " + resolution.pendingPrompts.length +
          " column(s) need confirmation before reconciliation can run: " +
          resolution.pendingPrompts.map((p) => p.sheetName + "::" + p.fieldName).join(", ")
        );
        mentorShowColumnMemoryQueue(resolution.pendingPrompts, mentorScanForGstReconciliation, mentorGstAppendAuditLogRow, mentorGstLogAction);
        return;
      }
      mentorHideColumnMemoryQueue(); // a previously-blocked pair is now fully resolved — clear that UI if it was showing

      // gstr2aSheet/booksSheet's .result.columns are shallow-merged with
      // the newly resolved fields before being handed to
      // mentorComputeGstReconciliation() — every downstream detector and
      // gst-report-writer.js keeps reading a plain columns object,
      // identical shape to what identifyGstColumns() has always returned.
      const resolvedGstr2aSheet = { ...gstr2aSheet, result: { ...gstr2aSheet.result, columns: { ...gstr2aSheet.result.columns, ...resolution.resolvedColumns.gstr2a } } };
      const resolvedBooksSheet = { ...booksSheet, result: { ...booksSheet.result, columns: { ...booksSheet.result.columns, ...resolution.resolvedColumns.books } } };

      console.log("MENTOR GST reconciliation: '" + key + "' — columns resolved, computing...");
      const proposal = mentorComputeGstReconciliation(resolvedGstr2aSheet, resolvedBooksSheet);
      proposal.alreadyGenerated = alreadyGenerated;
      proposal.dataHashSettingKey = hashSettingKey;
      proposal.currentDataHash = currentDataHash;
      mentorPendingGstProposal = proposal;

      if (alreadyGenerated) {
        // Data has changed but sheets already exist — show the persistent
        // summary (from the LAST generation, clearly labeled as such) WITH
        // the regenerate banner on top, rather than replacing the summary
        // with a bare prompt.
        await mentorShowPersistentGstSummary(context, proposal, proposal);
      } else {
        // Never generated at all — nothing to read back yet.
        mentorShowGstProposalCard(proposal);
      }
    });
  } catch (error) {
    console.error("MENTOR GST reconciliation scan error:", error);
  }
}

// Pure computation — Step 1 + Step 2, read-only. Nothing here touches the
// workbook; this is exactly the fast, deterministic first pass, run
// automatically so the proposal card has real numbers, not a guess.
function mentorComputeGstReconciliation(gstr2aSheet, booksSheet) {
  const gstr2aInput = { values: gstr2aSheet.result.values, headerRowIndex: gstr2aSheet.result.headerRowIndex, columns: gstr2aSheet.result.columns };
  const booksInput = { values: booksSheet.result.values, headerRowIndex: booksSheet.result.headerRowIndex, columns: booksSheet.result.columns };

  const comparisonSummary = reconcileGstr2aVsBooks(gstr2aInput, booksInput);
  const wrongHeadResult = detectWrongHead(gstr2aInput.values, gstr2aInput.headerRowIndex, gstr2aInput.columns);
  const crossSheetWrongHeadResult = detectCrossSheetWrongHead(gstr2aInput, booksInput);

  // Computed BEFORE invoiceLevelExtrasResult (below) — an invoice already
  // correctly classified as RCM has no reason to also independently exist
  // in Books under the ordinary vendor-matched ITC flow, so it must be
  // excluded from Extra-in-2A/Extra-in-Books rather than flagged under
  // two contradictory theories at once. See extra-invoice-detector.js's
  // docstring for the real-data case this fixes.
  const rcmBySource = [
    { sheetName: gstr2aSheet.sheetName, role: "gstr2a", ...gstr2aInput, result: detectRcmForSheet(gstr2aInput.values, gstr2aInput.headerRowIndex, gstr2aInput.columns) },
    { sheetName: booksSheet.sheetName, role: "purchase_register", ...booksInput, result: detectRcmForSheet(booksInput.values, booksInput.headerRowIndex, booksInput.columns) },
  ];
  const gstr2aRcmRowIndices = new Set(rcmBySource[0].result.flagged.map((f) => f.rowIndex));
  const booksRcmRowIndices = new Set(rcmBySource[1].result.flagged.map((f) => f.rowIndex));

  // Invoice-level (GSTIN + Invoice/Voucher Number), NOT Step 1's
  // GSTIN-aggregate vendor status — see extra-invoice-detector.js's
  // docstring for why vendor-level presence hides real per-invoice gaps.
  // This is the source of truth for Extra in 2A / Extra in Books; Step
  // 1's comparisonSummary.counts.missing_in_books/missing_in_gstr2a stay
  // as a separate, higher-level vendor-status count only (Summary sheet's
  // "Vendor match status" section).
  const invoiceLevelExtrasResult = detectInvoiceLevelExtras(gstr2aInput, booksInput, {
    gstr2aExcludedRowIndices: gstr2aRcmRowIndices,
    booksExcludedRowIndices: booksRcmRowIndices,
  });

  const itcBySource = [
    { sheetName: gstr2aSheet.sheetName, role: "gstr2a", ...gstr2aInput, result: detectIneligibleItcForSheet(gstr2aInput.values, gstr2aInput.headerRowIndex, gstr2aInput.columns) },
    { sheetName: booksSheet.sheetName, role: "purchase_register", ...booksInput, result: detectIneligibleItcForSheet(booksInput.values, booksInput.headerRowIndex, booksInput.columns) },
  ];
  const dupBySource = [
    { sheetName: gstr2aSheet.sheetName, role: "gstr2a", ...gstr2aInput, result: findDuplicateInvoices(gstr2aInput.values, gstr2aInput.headerRowIndex, gstr2aInput.columns) },
    { sheetName: booksSheet.sheetName, role: "purchase_register", ...booksInput, result: findDuplicateInvoices(booksInput.values, booksInput.headerRowIndex, booksInput.columns) },
  ];

  // Pattern-agnostic sanity-check backstop (rate-mismatch-detector.js) —
  // informational only, same treatment as Wrong Head: never blocks, never
  // changes the net eligible ITC figure. Silently not-applicable on a
  // sheet with no resolved Rate column (rate is optional, never asked
  // about — see gst-column-ambiguity-rules.js).
  const rateMismatchBySource = [
    { sheetName: gstr2aSheet.sheetName, role: "gstr2a", ...gstr2aInput, result: detectRateMismatch(gstr2aInput.values, gstr2aInput.headerRowIndex, gstr2aInput.columns) },
    { sheetName: booksSheet.sheetName, role: "purchase_register", ...booksInput, result: detectRateMismatch(booksInput.values, booksInput.headerRowIndex, booksInput.columns) },
  ];

  // Cross-sheet Wrong Head findings count toward the SAME "Wrong Head"
  // category, not a separate one — it's a complementary check for the
  // same kind of issue (see cross-sheet-wrong-head-detector.js), not a
  // new category of finding.
  const hasWrongHeadFindings =
    (wrongHeadResult.applicable && wrongHeadResult.flagged.length > 0) || (crossSheetWrongHeadResult.applicable && crossSheetWrongHeadResult.flagged.length > 0);

  const hasExtraIn2AFindings = invoiceLevelExtrasResult.applicable && invoiceLevelExtrasResult.extraIn2A.length > 0;
  const hasExtraInBooksFindings = invoiceLevelExtrasResult.applicable && invoiceLevelExtrasResult.extraInBooks.length > 0;
  // Low-confidence amount+GSTIN fallback (Fix 4) — surfaced as its own
  // category so it's visible in the sidebar, but never netted into either
  // Extra total; see extra-invoice-detector.js's "POSSIBLE MATCH" docstring.
  const hasPossibleMatchFindings = invoiceLevelExtrasResult.applicable && invoiceLevelExtrasResult.possibleMatches.length > 0;

  const categoriesWithFindings = [
    comparisonSummary.counts.discrepancy > 0,
    hasExtraIn2AFindings,
    hasExtraInBooksFindings,
    hasPossibleMatchFindings,
    hasWrongHeadFindings,
    rcmBySource.some((s) => s.result.flagged.length > 0),
    itcBySource.some((s) => s.result.applicable && s.result.flagged.length > 0),
    dupBySource.some((s) => s.result.applicable && s.result.clusters.length > 0),
    rateMismatchBySource.some((s) => s.result.applicable && s.result.flagged.length > 0),
  ].filter(Boolean).length;

  const unresolvedItemCount =
    comparisonSummary.counts.discrepancy +
    (invoiceLevelExtrasResult.applicable ? invoiceLevelExtrasResult.extraIn2A.length : 0) +
    (invoiceLevelExtrasResult.applicable ? invoiceLevelExtrasResult.extraInBooks.length : 0) +
    (invoiceLevelExtrasResult.applicable ? invoiceLevelExtrasResult.possibleMatches.length : 0) +
    (wrongHeadResult.applicable ? wrongHeadResult.flagged.length : 0) +
    (crossSheetWrongHeadResult.applicable ? crossSheetWrongHeadResult.flagged.length : 0) +
    rcmBySource.reduce((sum, s) => sum + s.result.flagged.length, 0) +
    itcBySource.reduce((sum, s) => sum + (s.result.applicable ? s.result.flagged.length : 0), 0) +
    dupBySource.reduce((sum, s) => sum + (s.result.applicable ? s.result.clusters.length : 0), 0) +
    rateMismatchBySource.reduce((sum, s) => sum + (s.result.applicable ? s.result.flagged.length : 0), 0);

  return {
    key: mentorGstWorkbookKey(gstr2aSheet.sheetName, booksSheet.sheetName),
    sheetNames: { gstr2a: gstr2aSheet.sheetName, books: booksSheet.sheetName },
    comparisonSummary,
    invoiceLevelExtrasResult,
    wrongHeadResult,
    crossSheetWrongHeadResult,
    rcmBySource,
    itcBySource,
    dupBySource,
    rateMismatchBySource,
    totalVendors: comparisonSummary.totalGstins,
    matchedVendors: comparisonSummary.counts.matched,
    categoriesWithFindings,
    unresolvedItemCount,
  };
}

function mentorGstEscapeHtml(text) {
  return String(text).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

// Vendor-name suggestion trigger: only when Ineligible ITC's own
// description-based check found literally nothing (see
// ineligible-itc-detector.js — this never loosens that check, it's a
// separate, optional, explicitly user-triggered addition for exactly the
// case where there was nothing for it to work with). Shared by both the
// accept flow (fresh generation) and the read-back path (persistent
// summary on an unchanged session) — same proposal shape either way.
function mentorGstSetupIneligibleItcVendorTrigger(proposal) {
  const itcFlaggedCount = proposal.itcBySource.reduce((sum, s) => sum + (s.result.applicable ? s.result.flagged.length : 0), 0);
  if (itcFlaggedCount === 0) {
    const gstr2aSource = proposal.itcBySource.find((s) => s.role === "gstr2a");
    const booksSource = proposal.itcBySource.find((s) => s.role === "purchase_register");
    mentorGstIneligibleItcVendors = extractDistinctVendors(
      { values: gstr2aSource.values, headerRowIndex: gstr2aSource.headerRowIndex, columns: gstr2aSource.columns },
      { values: booksSource.values, headerRowIndex: booksSource.headerRowIndex, columns: booksSource.columns }
    );
  } else {
    mentorGstIneligibleItcVendors = null;
  }
  mentorGstIneligibleItcSuggestionState = "idle";
  mentorGstIneligibleItcSuggestionResult = null;
  mentorGstIneligibleItcSuggestionError = "";
}

// Populates and renders the PERSISTENT summary card by reading Top Issues
// directly off the already-generated sheets (readGstTopIssuesFromExisting
// Sheets — a pure read, no write, no Audit Log entry) — used both when
// staying silent (regenerateProposal: null, no banner) and when source
// data has changed but sheets still exist (regenerateProposal: the
// pending proposal, shown as a banner above the — still last-generation —
// summary). proposal: mentorComputeGstReconciliation()'s read-only result,
// used only for the vendor-suggestion trigger and (when regenerating) the
// banner's own counts — NOT for the Top Issues numbers themselves, which
// always come from the sheets.
async function mentorShowPersistentGstSummary(context, proposal, regenerateProposal) {
  const readBack = await readGstTopIssuesFromExistingSheets(context);
  if (!readBack) {
    // Sheets unexpectedly missing despite alreadyGenerated being true —
    // fall back to the normal propose/generate flow rather than showing
    // an empty or broken summary.
    mentorPendingGstProposal = proposal;
    mentorShowGstProposalCard(proposal);
    return;
  }

  mentorGstTopIssues = readBack.topIssues;
  mentorGstTopIssuesShowAll = false;
  mentorGstConfirmationPreamble =
    "On file — " + readBack.topIssues.length + " item(s) currently flagged across the generated 'GST - ...' sheets" +
    " (Priority 1 threshold: RCM/Ineligible ITC items ≥ " + mentorGstFormatAmount(readBack.materialityThreshold) + ").";

  mentorGstSetupIneligibleItcVendorTrigger(proposal);
  mentorPendingGstProposal = regenerateProposal;
  mentorRenderGstTopIssuesList();
}

/* ============================================================
   Point-and-explain — select a row on any of the 8 GST sheets, the
   sidebar shows which rule fired, its legal citation, and the row's
   Reviewer Status/Note. Everything here is a read of data already
   computed and written by writeGstReconciliationReport() (via
   explainGstRow() in gst-report-writer.js) — no LLM call, no
   recomputation, just surfacing what's already on the sheet.
   ============================================================ */

function mentorRenderGstExplainHtml(explanation) {
  const accent = "#00B0F0"; // same visual register as the general P&L-suggestion badge — this is informational, not a proposal needing accept/dismiss
  let html = "<div style='background:#1a1a2e;border-left:4px solid " + accent + ";padding:8px 10px;border-radius:6px;font-size:11px;color:#cfd2d8;line-height:1.5;'>";
  html += "<strong style='color:#fff;'>" + mentorGstEscapeHtml(explanation.ruleLabel) + "</strong><br/>";
  html += mentorGstEscapeHtml(explanation.whichRuleFired) + "<br/>";
  html += "<div style='margin-top:4px;color:#9aa0a6;'>" + mentorGstEscapeHtml(explanation.legalCitation) + "</div>";
  // Reviewer Status/Note only applies to a flagged row on one of the 8 rule
  // sheets — every caller there always passes real strings (possibly
  // empty), even when blank. A caller with no concept of "reviewer" at all
  // (e.g. Summary-row column traceability) leaves both fields genuinely
  // undefined, which suppresses this whole block rather than showing the
  // misleading "Not yet reviewed" nudge for something that isn't a
  // reviewable item.
  if (explanation.reviewerStatus !== undefined || explanation.reviewerNote !== undefined) {
    html += "<div style='margin-top:6px;padding-top:6px;border-top:1px solid #333;'>";
    if (explanation.reviewerStatus || explanation.reviewerNote) {
      html += "<strong>Reviewer:</strong> " + (explanation.reviewerStatus ? mentorGstEscapeHtml(explanation.reviewerStatus) : "(no status set)");
      if (explanation.reviewerNote) html += " — " + mentorGstEscapeHtml(explanation.reviewerNote);
    } else {
      html += "<span style='color:#7d828a;'>Not yet reviewed — set Reviewer Status/Note directly on this row.</span>";
    }
    html += "</div>";
  }
  html += "</div>";
  return html;
}

function mentorHideGstExplain() {
  const container = document.getElementById("mentor-gst-explain");
  if (container) {
    container.style.display = "none";
    container.innerHTML = "";
  }
}

// Part 3 — column traceability for "GST - Summary" rows. summaryValues:
// the Summary sheet's full used-range values, already read fresh by the
// caller. Delegates the actual explanation to mentor-gst-summary-explain.js
// (which re-derives the current resolved columns live — see that file's
// docstring for why) and renders it through the SAME mentorRenderGstExplainHtml
// every other GST sheet's point-and-explain already uses — no new UI.
async function mentorRenderGstSummaryExplain(context, container, summaryValues, selectedRowIndex) {
  const row = summaryValues[selectedRowIndex];
  if (!row) {
    mentorHideGstExplain();
    return;
  }
  // NOT trimmed — GST_SUMMARY_ROW_TRACEABILITY's keys deliberately include
  // the leading two-space indentation writeSummarySheet() itself writes
  // for sub-items, and Excel preserves leading whitespace in a text cell
  // verbatim, so the lookup must match that exactly.
  const rowLabel = String(row[0] === undefined || row[0] === null ? "" : row[0]);

  const gstr2aSheetName = summaryValues[1] && summaryValues[1][1]; // "2A sheet" row
  const booksSheetName = summaryValues[2] && summaryValues[2][1]; // "Books sheet" row

  let explanation = null;
  if (gstr2aSheetName && booksSheetName) {
    try {
      explanation = await mentorExplainGstSummaryRow(context, rowLabel, gstr2aSheetName, booksSheetName);
    } catch (error) {
      console.error("MENTOR GST summary explain error:", error);
    }
  }

  if (!explanation) {
    // Not a traceable total (a section header, the title, plain vendor
    // counts) or the source sheets couldn't currently be read — same
    // fallback message as before, slightly reworded now that some rows
    // ARE explainable.
    container.innerHTML =
      "<div style='background:#1a1a2e;border-left:4px solid #6c757d;padding:8px 10px;border-radius:6px;font-size:11px;color:#9aa0a6;'>" +
      "Not a traceable total — select one of the Summary sheet's total/subtotal lines for its source column(s), or a row on one of the other 'GST - ...' sheets for row-level detail.</div>";
    container.style.display = "block";
    return;
  }

  container.innerHTML = mentorRenderGstExplainHtml(explanation);
  container.style.display = "block";
}

async function mentorGstOnSelectionChanged(event) {
  const container = document.getElementById("mentor-gst-explain");
  if (!container) return;

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      await context.sync();

      const isExplainable = GST_EXPLAIN_SHEET_NAMES.has(sheet.name);
      const isGstSummary = sheet.name === "GST - Summary";
      if (!isExplainable && !isGstSummary) {
        mentorHideGstExplain(); // not one of the 8 GST sheets — stay out of the way entirely
        return;
      }

      const address = event && event.address;
      const rowMatch = address && address.match(/([A-Z]+)(\d+)/); // first row of the selection, even for a multi-cell drag
      if (!rowMatch) {
        mentorHideGstExplain();
        return;
      }
      const selectedRowIndex = parseInt(rowMatch[2], 10) - 1; // 0-based, absolute position on the sheet

      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("values, isNullObject");
      await context.sync();
      if (usedRange.isNullObject) {
        mentorHideGstExplain();
        return;
      }

      if (isGstSummary) {
        await mentorRenderGstSummaryExplain(context, container, usedRange.values, selectedRowIndex);
        return;
      }

      const explanation = explainGstRow(sheet.name, usedRange.values, selectedRowIndex);
      if (!explanation) {
        mentorHideGstExplain(); // header row, the (collapsed) explanation block, or past the last data row
        return;
      }

      container.innerHTML = mentorRenderGstExplainHtml(explanation);
      container.style.display = "block";
    });
  } catch (error) {
    console.error("MENTOR GST point-and-explain error:", error);
  }
}

function mentorShowGstProposalCard(proposal) {
  const container = document.getElementById("mentor-gst-reconciliation-prompt");
  if (!container) return;

  const accent = "#e08e0b"; // distinct from both the P&L-suggestion blue and the sheet-memory purple
  const verb = proposal.alreadyGenerated ? "Re-ran" : "Ran";
  const message =
    "Found GST data: '" + proposal.sheetNames.gstr2a + "' (GSTR-2A) vs '" + proposal.sheetNames.books + "' (Books). " + verb + " reconciliation — " +
    proposal.matchedVendors + " of " + proposal.totalVendors + " vendors matched; " +
    proposal.unresolvedItemCount + " item(s) need attention across " + proposal.categoriesWithFindings + " categor" + (proposal.categoriesWithFindings === 1 ? "y" : "ies") + ".";

  const actionVerb = proposal.alreadyGenerated ? "Regenerate" : "Generate";
  const regenerateNote = proposal.alreadyGenerated
    ? "The 'GST - ...' sheets already exist — regenerating replaces them with a fresh run (any manual edits made directly on those sheets will be overwritten, since they're recreated from scratch each time)."
    : "Nothing is filed or claimed — these are proposed for your review.";

  let html = "<div style='background:#1a1a2e;border-left:4px solid " + accent + ";padding:10px 10px 10px 12px;border-radius:6px;font-size:12px;line-height:1.5;color:#fff;'>";
  html += "<div style='margin-bottom:8px;'>" + mentorGstEscapeHtml(message) + "</div>";
  html += "<div style='color:#cfd2d8;margin-bottom:8px;'>" + actionVerb + " the review sheets (Extra in Books, Extra in 2A, Mismatch, Wrong Head, RCM, Ineligible ITC, Duplicate Invoices, Summary)? " + mentorGstEscapeHtml(regenerateNote) + "</div>";
  html += "<div>";
  html += "<button onclick='mentorAcceptGstReconciliation()' style='margin-right:6px;padding:3px 10px;background:#28a745;color:white;border:none;border-radius:4px;cursor:pointer;font-size:11px;'>" + actionVerb + " the sheets</button>";
  html += "<button onclick='mentorDismissGstReconciliation()' style='padding:3px 10px;background:#555;color:white;border:none;border-radius:4px;cursor:pointer;font-size:11px;'>Not now</button>";
  html += "</div></div>";

  container.innerHTML = html;
  container.style.display = "block";
}

function mentorHideGstProposalCard() {
  const container = document.getElementById("mentor-gst-reconciliation-prompt");
  if (container) {
    container.style.display = "none";
    container.innerHTML = "";
  }
}

const MENTOR_GST_TOP_ISSUES_HEADLINE_COUNT = 5;

function mentorGstFormatAmount(amount) {
  return "₹" + Math.round(Math.abs(amount)).toLocaleString("en-IN");
}

// TIER labels shown once as a small legend, not repeated per row — keeps
// each list entry to one line (amount + reason) as asked, while still
// making the ranking legible instead of just an unexplained order.
const MENTOR_GST_TIER_LABEL = { 1: "Priority 1", 2: "Priority 2", 3: "Priority 3" };

// Renders the "Get vendor-name suggestions" button/status — a separate,
// optional section appended below the Top Issues list. Returns "" when
// there's nothing to show (Ineligible ITC had real description-based
// findings, or a description field just doesn't apply to check at all).
function mentorRenderIneligibleItcSuggestionSectionHtml() {
  if (!mentorGstIneligibleItcVendors || mentorGstIneligibleItcVendors.length === 0) return "";

  let html = "<div style='margin-top:8px;padding-top:8px;border-top:1px solid #333;'>";

  if (mentorGstIneligibleItcSuggestionState === "loading") {
    html += "<div style='color:#9aa0a6;'>Checking " + mentorGstIneligibleItcVendors.length + " vendor name(s) against Section 17(5)…</div>";
  } else if (mentorGstIneligibleItcSuggestionState === "done") {
    const r = mentorGstIneligibleItcSuggestionResult;
    html +=
      "<div style='color:#cfd2d8;'>Vendor-name screening done — " + r.totalChecked + " checked, " + r.flaggedCount + " flagged, " + r.skippedCount +
      " no clear match. See the amber section on 'GST - Ineligible ITC'. Unconfirmed — review before treating as ineligible.</div>";
    html += "<a href='#' onclick='mentorGstFetchIneligibleItcVendorSuggestions(); return false;' style='color:#8ab4f8;font-size:11px;'>Run again</a>";
  } else if (mentorGstIneligibleItcSuggestionState === "error") {
    html += "<div style='color:#f28b82;margin-bottom:4px;'>Vendor-name suggestions failed: " + mentorGstEscapeHtml(mentorGstIneligibleItcSuggestionError) + "</div>";
    html += "<a href='#' onclick='mentorGstFetchIneligibleItcVendorSuggestions(); return false;' style='color:#8ab4f8;font-size:11px;'>Try again</a>";
  } else {
    html +=
      "<div style='color:#9aa0a6;margin-bottom:4px;'>Ineligible ITC found no genuine per-transaction description to check (" +
      mentorGstIneligibleItcVendors.length + " distinct vendor(s) on '" + "2A" + "'/'" + "Books" + "'). Optional — vendor-name-based suggestions from an LLM, clearly unconfirmed, for manual review only.</div>";
    html += "<button onclick='mentorGstFetchIneligibleItcVendorSuggestions()' style='padding:3px 10px;background:#b8860b;color:white;border:none;border-radius:4px;cursor:pointer;font-size:11px;'>Get vendor-name suggestions (uses AI, may take a few seconds)</button>";
  }

  html += "</div>";
  return html;
}

// Shown INSIDE the persistent summary card, above the Top Issues list,
// only when mentorPendingGstProposal is set for an ALREADY-generated
// pair (i.e. source data has changed since the last generation). The
// summary below it still reflects the LAST generation — labeled as such,
// since it's a pure read of the currently-written sheets, not the new
// data — so it's never presented as if it already matches what changed.
function mentorRenderGstRegenerateBannerHtml() {
  if (!mentorPendingGstProposal || !mentorPendingGstProposal.alreadyGenerated) return "";
  const proposal = mentorPendingGstProposal;
  const message =
    "Source data has changed since these sheets were last generated — re-checking now shows " +
    proposal.matchedVendors + " of " + proposal.totalVendors + " vendors matched, " + proposal.unresolvedItemCount +
    " item(s) across " + proposal.categoriesWithFindings + " categor" + (proposal.categoriesWithFindings === 1 ? "y" : "ies") + ".";

  let html = "<div style='margin-bottom:8px;padding-bottom:8px;border-bottom:1px solid #3a3d45;'>";
  html += "<div style='color:#f0b429;margin-bottom:6px;'>" + mentorGstEscapeHtml(message) + " The summary below still reflects the LAST generation.</div>";
  html += "<button onclick='mentorAcceptGstReconciliation()' style='margin-right:6px;padding:3px 10px;background:#28a745;color:white;border:none;border-radius:4px;cursor:pointer;font-size:11px;'>Regenerate the sheets</button>";
  html += "<button onclick='mentorDismissGstReconciliation()' style='padding:3px 10px;background:#555;color:white;border:none;border-radius:4px;cursor:pointer;font-size:11px;'>Not now</button>";
  html += "</div>";
  return html;
}

function mentorRenderGstTopIssuesList() {
  const container = document.getElementById("mentor-gst-reconciliation-prompt");
  if (!container) return;

  const banner = mentorRenderGstRegenerateBannerHtml();

  if (mentorGstTopIssues.length === 0) {
    container.innerHTML =
      "<div style='background:#1a1a2e;border-left:4px solid #28a745;padding:8px 10px;border-radius:6px;font-size:11px;color:#cfd2d8;'>" +
      banner +
      mentorGstEscapeHtml(mentorGstConfirmationPreamble) +
      " No issues found — every vendor matched and no line items were flagged." +
      mentorRenderIneligibleItcSuggestionSectionHtml() +
      "</div>";
    container.style.display = "block";
    return;
  }

  const visibleCount = mentorGstTopIssuesShowAll ? mentorGstTopIssues.length : Math.min(MENTOR_GST_TOP_ISSUES_HEADLINE_COUNT, mentorGstTopIssues.length);

  let html = "<div style='background:#1a1a2e;border-left:4px solid #28a745;padding:8px 10px;border-radius:6px;font-size:11px;color:#cfd2d8;'>";
  html += banner;
  html += mentorGstEscapeHtml(mentorGstConfirmationPreamble) + "<br/>";
  html += "<div style='margin-top:6px;color:#9aa0a6;'>Top issues (highest priority, then largest rupee amount, first):</div>";
  html += "<ol style='margin:6px 0 6px 0;padding-left:16px;'>";
  for (let i = 0; i < visibleCount; i++) {
    const issue = mentorGstTopIssues[i];
    html +=
      "<li style='margin-bottom:4px;cursor:pointer;color:#fff;' onclick='mentorGstSelectIssue(" + i + ")' title='" + mentorGstEscapeHtml(MENTOR_GST_TIER_LABEL[issue.tier] + " — click to select this row on \"" + issue.sheetName + "\"") + "'>" +
      "<strong>" + mentorGstFormatAmount(issue.amount) + "</strong> — " + mentorGstEscapeHtml(issue.reason) +
      " <span style='color:#7d828a;'>(" + mentorGstEscapeHtml(issue.sheetName) + ")</span></li>";
  }
  html += "</ol>";
  if (mentorGstTopIssues.length > MENTOR_GST_TOP_ISSUES_HEADLINE_COUNT) {
    html += mentorGstTopIssuesShowAll
      ? "<a href='#' onclick='mentorGstToggleShowAllIssues(); return false;' style='color:#8ab4f8;font-size:11px;'>Show top " + MENTOR_GST_TOP_ISSUES_HEADLINE_COUNT + " only</a>"
      : "<a href='#' onclick='mentorGstToggleShowAllIssues(); return false;' style='color:#8ab4f8;font-size:11px;'>Show all " + mentorGstTopIssues.length + "</a>";
  }
  html += mentorRenderIneligibleItcSuggestionSectionHtml();
  html += "</div>";

  container.innerHTML = html;
  container.style.display = "block";
}

// vendors: [{gstin, vendorName}, ...] already deduplicated by
// extractDistinctVendors(). Sends ONLY these two fields per vendor — no
// amounts, no other data — in a single batched request.
window.mentorGstFetchIneligibleItcVendorSuggestions = async function () {
  if (!mentorGstIneligibleItcVendors || mentorGstIneligibleItcVendors.length === 0) return;

  mentorGstIneligibleItcSuggestionState = "loading";
  mentorRenderGstTopIssuesList();

  try {
    const response = await fetch(MENTOR_SUGGESTIONS_PROXY_URL, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ vendors: mentorGstIneligibleItcVendors }),
    });
    let responseBody = null;
    try {
      responseBody = await response.json();
    } catch (parseError) {
      responseBody = null;
    }
    if (!response.ok) {
      throw new Error((responseBody && responseBody.error) || "Suggestions server returned status " + response.status);
    }

    let writeResult = null;
    await Excel.run(async (context) => {
      writeResult = await appendIneligibleItcVendorSuggestions(context, responseBody);
      if (mentorGstAppendAuditLogRow) {
        const description =
          (responseBody.mock ? "[MOCK — no real LLM call] " : "") +
          "Ran AI vendor-name screening for Ineligible ITC — " + writeResult.totalChecked + " vendor(s) checked, " + writeResult.flaggedCount +
          " flagged, " + writeResult.skippedCount + " no clear match. Unconfirmed — written to 'GST - Ineligible ITC' for manual review, not netted into the Summary sheet.";
        await mentorGstAppendAuditLogRow(context, "GST - Ineligible ITC", "(vendor-name suggestions section)", description);
      }
    });

    mentorGstIneligibleItcSuggestionResult = writeResult;
    mentorGstIneligibleItcSuggestionState = "done";
    if (mentorGstLogAction) mentorGstLogAction("Ran AI vendor-name screening for Ineligible ITC (" + (responseBody.mock ? "mock" : "live") + ")");
  } catch (error) {
    console.error("MENTOR vendor-name ITC suggestions error:", error);
    mentorGstIneligibleItcSuggestionState = "error";
    mentorGstIneligibleItcSuggestionError =
      error.message + (error.message && error.message.includes("fetch") ? " — is the suggestions proxy running? (npm run gst-suggestions-proxy)" : "");
  }

  mentorRenderGstTopIssuesList();
};

window.mentorGstToggleShowAllIssues = function () {
  mentorGstTopIssuesShowAll = !mentorGstTopIssuesShowAll;
  mentorRenderGstTopIssuesList();
};

// Switches to the issue's sheet and selects the exact row(s) it landed on
// (see attachTopIssueCandidates() in gst-report-writer.js for how rowIndex/
// rowSpan are computed) — entire row, not just column A, so it's obvious
// at a glance regardless of which column happens to be in view.
window.mentorGstSelectIssue = async function (index) {
  const issue = mentorGstTopIssues[index];
  if (!issue) return;
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItemOrNullObject(issue.sheetName);
      sheet.load("isNullObject");
      await context.sync();
      if (sheet.isNullObject) {
        console.error("MENTOR GST reconciliation: '" + issue.sheetName + "' no longer exists — can't select the row");
        return;
      }
      sheet.activate();
      const range = sheet.getRangeByIndexes(issue.rowIndex, 0, issue.rowSpan, 1).getEntireRow();
      range.select();
      await context.sync();
    });
  } catch (error) {
    console.error("MENTOR GST reconciliation: failed to select issue row", error);
  }
};

window.mentorAcceptGstReconciliation = async function () {
  if (!mentorPendingGstProposal) return;
  const proposal = mentorPendingGstProposal;
  mentorGstHandledThisSession.add(proposal.key);

  const actionVerb = proposal.alreadyGenerated ? "Regenerating" : "Generating";
  const container = document.getElementById("mentor-gst-reconciliation-prompt");
  if (container) container.innerHTML = "<div style='color:#9aa0a6;font-size:11px;padding:6px 0;'>" + actionVerb + " the review sheets…</div>";

  try {
    let summary = null;
    let reviewPreservation = null;
    let topIssues = [];
    let materialityThreshold = 0;
    await Excel.run(async (context) => {
      const result = await writeGstReconciliationReport(context, {
        sheetNames: proposal.sheetNames,
        comparisonSummary: proposal.comparisonSummary,
        invoiceLevelExtrasResult: proposal.invoiceLevelExtrasResult,
        wrongHeadResult: proposal.wrongHeadResult,
        crossSheetWrongHeadResult: proposal.crossSheetWrongHeadResult,
        rcmBySource: proposal.rcmBySource,
        itcBySource: proposal.itcBySource,
        dupBySource: proposal.dupBySource,
        rateMismatchBySource: proposal.rateMismatchBySource,
      });
      summary = result.summary;
      reviewPreservation = result.reviewPreservation;
      topIssues = result.topIssues;
      materialityThreshold = result.materialityThreshold;

      // Record what data this generation was run against — only after a
      // SUCCESSFUL write, so a failed generation never gets falsely
      // marked as "already processed." This is what lets the next scan
      // (including after a full Excel close/reopen) recognize "nothing's
      // changed" and stay silent instead of re-prompting.
      if (proposal.dataHashSettingKey) {
        context.workbook.settings.add(proposal.dataHashSettingKey, proposal.currentDataHash);
      }

      if (mentorGstAppendAuditLogRow) {
        const preservationNote =
          reviewPreservation.reappliedCount > 0 || reviewPreservation.orphanedCount > 0
            ? " Carried forward " + reviewPreservation.reappliedCount + " previous Reviewer Status/Note entr" + (reviewPreservation.reappliedCount === 1 ? "y" : "ies") + "." +
              (reviewPreservation.orphanedCount > 0 ? " " + reviewPreservation.orphanedCount + " earlier entr" + (reviewPreservation.orphanedCount === 1 ? "y" : "ies") + " could not be matched to a row in this run (that item is no longer flagged) — see console for details." : "")
            : "";
        const description =
          (proposal.alreadyGenerated ? "Re-ran" : "Ran") + " GST reconciliation ('" + proposal.sheetNames.gstr2a + "' vs '" + proposal.sheetNames.books + "') — " +
          proposal.matchedVendors + "/" + proposal.totalVendors + " vendors matched, " + proposal.unresolvedItemCount +
          " item(s) flagged across " + proposal.categoriesWithFindings + " categories. Net provisionally eligible ITC: ₹" + summary.netProvisionallyEligible +
          ". " + (proposal.alreadyGenerated ? "Regenerated" : "Generated") + " 8 review sheets (GST - Extra in Books/2A/Mismatch/Wrong Head/RCM/Ineligible ITC/Duplicate Invoices/Summary)." +
          preservationNote;
        await mentorGstAppendAuditLogRow(context, proposal.sheetNames.gstr2a, "(whole workbook)", description);
      }
    });

    const preservationLine =
      reviewPreservation && (reviewPreservation.reappliedCount > 0 || reviewPreservation.orphanedCount > 0)
        ? " Carried forward " + reviewPreservation.reappliedCount + " of your previous review note(s)" + (reviewPreservation.orphanedCount > 0 ? "; " + reviewPreservation.orphanedCount + " could no longer be matched (see the Audit Log)." : ".")
        : "";

    mentorGstTopIssues = topIssues;
    mentorGstTopIssuesShowAll = false; // always start collapsed to the headline top-5 on a fresh (re)generation
    mentorGstConfirmationPreamble =
      "Reconciliation complete — " + proposal.unresolvedItemCount + " item(s) flagged across " + proposal.categoriesWithFindings + " categories" +
      " (Priority 1 threshold: RCM/Ineligible ITC items ≥ " + mentorGstFormatAmount(materialityThreshold) + ")." + preservationLine;

    mentorGstSetupIneligibleItcVendorTrigger(proposal);
    mentorPendingGstProposal = null; // just written — nothing pending to regenerate
    mentorRenderGstTopIssuesList();

    if (mentorGstLogAction) mentorGstLogAction("Generated GST reconciliation review sheets");
  } catch (error) {
    console.error("MENTOR GST reconciliation: failed to write report sheets", error);
    if (container) {
      container.innerHTML = "<div style='background:#1a1a2e;border-left:4px solid #dc3545;padding:8px 10px;border-radius:6px;font-size:11px;color:#cfd2d8;'>Something went wrong generating the review sheets: " + mentorGstEscapeHtml(error.message) + "</div>";
    }
  }

  mentorPendingGstProposal = null;
};

window.mentorDismissGstReconciliation = function () {
  const hadPersistentSummary = mentorPendingGstProposal && mentorPendingGstProposal.alreadyGenerated;
  if (mentorPendingGstProposal) mentorGstHandledThisSession.add(mentorPendingGstProposal.key);
  mentorPendingGstProposal = null;
  if (hadPersistentSummary) {
    // Dismissing a regenerate banner on top of an already-generated
    // sheet set — the persistent summary underneath stays visible, only
    // the banner (which reads mentorPendingGstProposal) disappears.
    mentorRenderGstTopIssuesList();
  } else {
    // Never-generated case (mentorShowGstProposalCard) — nothing to fall
    // back to, hide the whole card.
    mentorHideGstProposalCard();
  }
};

if (typeof module !== "undefined" && module.exports) {
  // mentorComputeGstReconciliation is exported for reuse outside the live
  // UI (scripts/validate-real-files.js) — it's pure/synchronous (no Excel
  // calls of its own), so it's safe to call directly once its two sheet
  // inputs are already resolved.
  module.exports = { mentorGstReconciliationInit, mentorScanForGstReconciliation, mentorComputeGstReconciliation };
}
