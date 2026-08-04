/* global Excel */

/*
 * mentor-sheet-memory-ui.js
 * --------------------------
 * MENTOR Phase 2 — live UI wiring for the sheet-memory layer.
 *
 * Bridges the pure sheet-classifier modules (signal-extractor, classifier,
 * sheet-memory) to the actual Excel add-in: scans every worksheet, and for
 * any sheet the rule-based classifier can't place into one of its four
 * known types, checks sheet memory for a remembered label before falling
 * back to an inline "what is this sheet?" prompt.
 *
 * Scope note: this scans every sheet except MENTOR's own generated ones
 * ("MENTOR Audit Log", "MENTOR Sheet Memory"). There's no reliable structural
 * signal yet to distinguish "a genuinely novel raw-data sheet" from "a
 * report/output sheet someone built" (e.g. this demo workbook's Variance
 * Analysis) — both come back "unknown" from the classifier. In practice
 * this self-heals: answer once ("this is a report, not a data sheet") and
 * sheet memory remembers it, so it stops asking. A cleaner fix (e.g. an
 * explicit opt-out, or marking MENTOR-authored sheets) is a reasonable
 * follow-up, not attempted here.
 *
 * Runs at workbook open and whenever a new sheet is added — not on every
 * edit, since a sheet's fundamental shape doesn't usually change cell by
 * cell and re-scanning every column of every sheet on a 700ms debounce
 * would be wasteful.
 *
 * Storage: a WorkbookSheetMemoryStore (a dedicated "MENTOR Sheet Memory"
 * sheet inside the workbook) is authoritative, with a BrowserSheetMemoryStore
 * (localStorage) as a fallback only — localStorage turned out not to
 * reliably survive task pane close/reopen in this WebView2 hosting context
 * (Office appears to route task panes across multiple separate storage
 * profiles), so it can no longer be trusted as the primary source of truth.
 */

const { extractSheetSignals } = require("../sheet-classifier/signal-extractor.js");
const { classifySheet } = require("../sheet-classifier/classifier.js");
const { resolveSheetLabel, rememberSheetLabel } = require("../sheet-classifier/sheet-memory.js");
const { BrowserSheetMemoryStore } = require("../sheet-classifier/browser-sheet-memory-store.js");
const { FallbackSheetMemoryStore } = require("../sheet-classifier/fallback-sheet-memory-store.js");
const { WorkbookSheetMemoryStore, SHEET_NAME: MENTOR_SHEET_MEMORY_SHEET_NAME } = require("./mentor-workbook-sheet-memory-store.js");

const mentorSheetMemoryStore = new FallbackSheetMemoryStore(new WorkbookSheetMemoryStore(), new BrowserSheetMemoryStore());
let mentorSheetMemoryLogAction = null; // injected — updates the in-memory history counter/toggle text only
let mentorSheetMemoryAppendAuditLogRow = null; // injected — the actual durable write to the "MENTOR Audit Log" sheet
let mentorSheetMemoryQueue = [];
let mentorPendingSheetMemoryPrompt = null;
const mentorDismissedSheetMemoryThisSession = new Set();

const MENTOR_OWNED_SHEET_NAMES = new Set(["MENTOR Audit Log", MENTOR_SHEET_MEMORY_SHEET_NAME]);

// Placeholder client identifier, scoped to this workbook file — a real
// account/client system will replace this. Until then, one workbook file
// is treated as one "client" for sheet-memory scoping purposes.
function mentorGetSheetMemoryClientId(workbookName) {
  return "workbook:" + workbookName;
}

function mentorEscapeHtml(text) {
  return String(text).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

// logAction: mentorLogAction(description) — updates the in-memory counter/
// toggle text only, does NOT write to the Audit Log sheet by itself.
// appendAuditLogRow: mentorAppendAuditLogRow(context, sheetName, cellRef,
// description) — the actual durable Excel write. Every other flow in
// taskpane.js calls both together; this file does the same below. Both are
// passed in explicitly rather than read off `window`, so this file's
// dependencies on taskpane.js are explicit, not implicit global lookups.
async function mentorSheetMemoryInit(logAction, appendAuditLogRow) {
  mentorSheetMemoryLogAction = logAction;
  mentorSheetMemoryAppendAuditLogRow = appendAuditLogRow;

  try {
    await Excel.run(async (context) => {
      context.workbook.worksheets.onAdded.add(mentorScanForUnknownSheets);
      await context.sync();
    });
  } catch (error) {
    console.error("MENTOR sheet-memory: failed to register onAdded listener", error);
  }

  await mentorScanForUnknownSheets();
}

async function mentorScanForUnknownSheets() {
  try {
    await Excel.run(async (context) => {
      context.workbook.load("name");
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      const clientId = mentorGetSheetMemoryClientId(context.workbook.name);
      const candidates = [];

      for (const sheet of sheets.items) {
        if (MENTOR_OWNED_SHEET_NAMES.has(sheet.name)) continue;
        if (mentorDismissedSheetMemoryThisSession.has(sheet.name)) continue;

        // One sheet failing (a store error, a malformed range, anything)
        // must not silently abort the scan for every OTHER sheet — that's
        // exactly how a real bug here previously produced zero prompts
        // workbook-wide with no visible error. Every sheet gets its own
        // try/catch and a loud console.error if something goes wrong.
        try {
          const usedRange = sheet.getUsedRangeOrNullObject();
          usedRange.load("values, numberFormat, isNullObject");
          await context.sync();
          if (usedRange.isNullObject) {
            console.log("MENTOR sheet-memory: '" + sheet.name + "' has no used range — skipping");
            continue;
          }

          const sheetSignals = extractSheetSignals(sheet.name, usedRange.values, { numberFormats: usedRange.numberFormat });
          const classification = classifySheet(sheetSignals);
          const result = await resolveSheetLabel({ clientId, sheetName: sheet.name, sheetSignals, classification, store: mentorSheetMemoryStore });
          console.log("MENTOR sheet-memory: '" + sheet.name + "' -> " + result.status + (result.status === "remembered" ? " (\"" + result.label + "\", via " + result.matchedVia + ")" : result.status === "classified" ? " (" + result.type + ")" : ""));

          if (result.status === "needs_input") {
            candidates.push({ ...result, clientId, sheetSignals });
          }
        } catch (sheetError) {
          console.error("MENTOR sheet-memory: failed to check sheet '" + sheet.name + "' — skipping it this scan", sheetError);
        }
      }

      console.log("MENTOR sheet-memory: scan complete, " + candidates.length + " sheet(s) need input:", candidates.map((c) => c.sheetName));

      mentorSheetMemoryQueue = candidates;
      mentorShowNextSheetMemoryPrompt();
    });
  } catch (error) {
    console.error("MENTOR sheet-memory scan error:", error);
  }
}

function mentorShowNextSheetMemoryPrompt() {
  const container = document.getElementById("mentor-sheet-memory-prompt");
  if (!container) return;

  if (mentorSheetMemoryQueue.length === 0) {
    mentorPendingSheetMemoryPrompt = null;
    container.style.display = "none";
    container.innerHTML = "";
    return;
  }

  mentorPendingSheetMemoryPrompt = mentorSheetMemoryQueue.shift();
  const accent = "#9b59d0"; // distinct from the blue P&L-suggestion accent — reads as a different kind of card

  let html = "<div style='background:#1a1a2e;border-left:4px solid " + accent + ";padding:10px 10px 10px 12px;border-radius:6px;font-size:12px;line-height:1.5;color:#fff;'>";
  html += "<div style='margin-bottom:6px;'>" + mentorEscapeHtml(mentorPendingSheetMemoryPrompt.prompt) + "</div>";
  html +=
    "<input id='mentor-sheet-memory-input' type='text' placeholder=\"e.g. 'Stock levels', 'Marketing spend log'\" " +
    "style='width:100%;box-sizing:border-box;padding:5px 7px;border-radius:4px;border:1px solid #444;background:#22252b;color:#fff;font-size:12px;margin-bottom:6px;' " +
    "onkeydown='if(event.key===\"Enter\"){mentorSubmitSheetMemoryAnswer();}' />";
  html += "<div>";
  html += "<button onclick='mentorSubmitSheetMemoryAnswer()' style='margin-right:6px;padding:3px 10px;background:#28a745;color:white;border:none;border-radius:4px;cursor:pointer;font-size:11px;'>Save</button>";
  html += "<button onclick='mentorSkipSheetMemoryPrompt()' style='padding:3px 10px;background:#555;color:white;border:none;border-radius:4px;cursor:pointer;font-size:11px;'>Skip for now</button>";
  html += "</div></div>";

  container.innerHTML = html;
  container.style.display = "block";

  const input = document.getElementById("mentor-sheet-memory-input");
  if (input) input.focus();
}

window.mentorSubmitSheetMemoryAnswer = async function () {
  if (!mentorPendingSheetMemoryPrompt) return;
  const input = document.getElementById("mentor-sheet-memory-input");
  const value = input ? input.value.trim() : "";
  if (!value) {
    if (input) input.style.border = "1px solid #dc3545"; // flag the empty field instead of silently doing nothing
    return;
  }

  const { clientId, sheetName, sheetSignals } = mentorPendingSheetMemoryPrompt;
  try {
    console.log("MENTOR sheet-memory: submitting answer", { clientId, sheetName, value });
    const rememberResult = await rememberSheetLabel({ clientId, sheetName, sheetSignals, userProvidedLabel: value, store: mentorSheetMemoryStore });
    console.log("MENTOR sheet-memory: rememberSheetLabel returned", rememberResult);

    const description = "Learned that '" + sheetName + "' is \"" + value + "\" — won't ask again for a sheet with this shape";
    if (mentorSheetMemoryAppendAuditLogRow) {
      try {
        console.log("MENTOR sheet-memory: writing Audit Log row for", sheetName);
        await Excel.run(async (context) => {
          await mentorSheetMemoryAppendAuditLogRow(context, sheetName, "(whole sheet)", description);
        });
        console.log("MENTOR sheet-memory: Audit Log row write completed for", sheetName);
        if (mentorSheetMemoryLogAction) mentorSheetMemoryLogAction(description);
      } catch (writeError) {
        console.error("MENTOR sheet-memory: failed to write Audit Log row", writeError);
        if (mentorSheetMemoryLogAction) {
          mentorSheetMemoryLogAction("Learned that '" + sheetName + "' is \"" + value + "\", but writing it to the Audit Log sheet failed: " + writeError.message);
        }
      }
    } else {
      console.error("MENTOR sheet-memory: mentorSheetMemoryAppendAuditLogRow was never injected — mentorSheetMemoryInit may not have run correctly");
      if (mentorSheetMemoryLogAction) mentorSheetMemoryLogAction(description); // no audit-log writer wired in — still track it in-memory
    }
  } catch (error) {
    console.error("MENTOR sheet-memory: failed to save answer", error);
  }

  mentorShowNextSheetMemoryPrompt();
};

window.mentorSkipSheetMemoryPrompt = function () {
  if (!mentorPendingSheetMemoryPrompt) return;
  mentorDismissedSheetMemoryThisSession.add(mentorPendingSheetMemoryPrompt.sheetName);
  mentorShowNextSheetMemoryPrompt();
};

if (typeof module !== "undefined" && module.exports) {
  module.exports = { mentorSheetMemoryInit, mentorScanForUnknownSheets, MENTOR_OWNED_SHEET_NAMES };
}
