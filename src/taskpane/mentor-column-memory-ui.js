/* global Excel */

/*
 * mentor-column-memory-ui.js
 * -----------------------------
 * Live UI wiring for column-memory (column-memory.js) — the column-level
 * counterpart to mentor-sheet-memory-ui.js. Bridges
 * gst-column-ambiguity-rules.js + column-memory.js to the actual Excel
 * add-in: for a workflow's two GST sheets (2A and Books), resolves every
 * "in play" field either silently from memory, or by queuing a blocking
 * prompt.
 *
 * Key interaction difference from sheet-memory's prompt: this is NOT free
 * text. It's a pick-one-from-a-list-with-value-previews prompt — candidate
 * columns are shown with their header and a few sample cell values, and
 * the user clicks the real one. There's always a finite, inspectable list
 * of options (even the zero-match fallback is "pick from every column on
 * the sheet"), unlike "what is this sheet" which has no such list.
 *
 * Deliberately NOT given its own taskpane.js init/registration — this
 * module has no event listeners of its own to register. It's a pure
 * sub-step invoked only from inside mentor-gst-reconciliation-ui.js's
 * already-registered mentorScanForGstReconciliation(), which passes in
 * its own mentorGstAppendAuditLogRow/mentorGstLogAction explicitly rather
 * than this module reading them off window or needing its own init call.
 *
 * Storage: a WorkbookColumnMemoryStore (a dedicated "MENTOR Column Memory"
 * sheet inside the workbook) is authoritative, with a BrowserColumnMemory-
 * backed BrowserSheetMemoryStore (localStorage, a DISTINCT key from sheet-
 * memory's) as a fallback only — same rationale as mentor-sheet-memory-
 * ui.js's own store composition. That whole store composition is TIER 2
 * (per-client, exactly what Part 1 built).
 *
 * TIER 1 (shared software-pattern memory, Part 2): checked only when Tier 2
 * doesn't resolve, and only for genuinely ambiguous fields (2+ real
 * candidates) — not for zero-match-required fallback prompts, where the
 * offered "candidates" are just every column on the sheet, not a real
 * cross-client software-export pattern. Talks to the local proxy server
 * (mentor-proxy-column-pattern-store.js) — see that file for why every
 * call there must never throw or block. A Tier-1 hit is never written back
 * into Tier 2 (no promotion) — it's re-checked live on every scan; Tier 2,
 * once populated by an explicit user answer, is always checked first and
 * always wins for that client's own future scans.
 */

const { resolveColumnField, rememberColumnField } = require("../gst-reconciliation/column-memory.js");
const { fieldsInPlay } = require("../gst-reconciliation/gst-column-ambiguity-rules.js");
const { normalizeHeaderText } = require("../sheet-classifier/structural-signature.js");
const { BrowserSheetMemoryStore } = require("../sheet-classifier/browser-sheet-memory-store.js");
const { FallbackSheetMemoryStore } = require("../sheet-classifier/fallback-sheet-memory-store.js");
const { WorkbookColumnMemoryStore, SHEET_NAME: COLUMN_MEMORY_SHEET_NAME } = require("./mentor-workbook-column-memory-store.js");
const { lookupColumnPattern, rememberColumnPattern } = require("./mentor-proxy-column-pattern-store.js");

const mentorColumnMemoryStore = new FallbackSheetMemoryStore(
  new WorkbookColumnMemoryStore(),
  new BrowserSheetMemoryStore({ storageKey: "mentorColumnMemory" })
);

let mentorColumnMemoryQueue = [];
let mentorPendingColumnMemoryPrompt = null;
let mentorColumnMemoryOnAllResolved = null;
let mentorColumnMemoryAppendAuditLogRow = null;
let mentorColumnMemoryLogAction = null;
// Session-only dismissal (re-asked next scan/session, not persisted) — same
// convention as mentor-sheet-memory-ui.js's own "Skip for now".
const mentorDismissedColumnMemoryThisSession = new Set(); // keyed by sheetName + "::" + fieldName

function mentorColumnMemoryEscapeHtml(text) {
  return String(text).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

// Duplicated locally rather than imported from mentor-gst-reconciliation-
// ui.js (which already has an identical helper) — that file requires THIS
// one, so importing the other way would be circular; matches this
// codebase's existing convention of small trivial per-file helpers over a
// shared util for one function.
function mentorColumnMemoryDebounce(fn, ms) {
  let timer = null;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), ms);
  };
}

// Jumps the Excel view to a candidate column and selects it — the exact
// same activate()+select() pattern mentorGstSelectIssue() already uses for
// point-and-explain navigation on flagged rows. The one difference: a
// candidate's index is relative to THIS sheet's used range at scan time,
// which (unlike MENTOR's own generated report sheets) can start at a
// different row/column than A1 on a real uploaded file — so this resolves
// the column via usedRange.getColumn(index), relative to the used range's
// own corner, rather than assuming/computing an absolute column letter.
// Best-effort only: a hover-preview (or the bonus highlight-on-submit)
// failing silently (sheet renamed/deleted, or the sheet edited since the
// scan so the index no longer exists) must never surface as a user-facing
// error or block the actual answer submission — this is a convenience on
// top of the real flow, not part of it.
async function mentorHighlightColumnMemoryCandidate(sheetName, candidate) {
  if (!candidate) return;
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
      sheet.load("isNullObject");
      await context.sync();
      if (sheet.isNullObject) return;

      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("isNullObject");
      await context.sync();
      if (usedRange.isNullObject) return;

      const columnRange = usedRange.getColumn(candidate.index);
      sheet.activate();
      columnRange.select();
      await context.sync();
    });
  } catch (error) {
    console.log("MENTOR column-memory: couldn't preview candidate column (sheet may have changed since the scan)", error.message);
  }
}

const mentorDebouncedHighlightColumnMemoryCandidate = mentorColumnMemoryDebounce(mentorHighlightColumnMemoryCandidate, 150);

// For each of the two workflow sheets, resolves every field fieldsInPlay()
// flags — from memory where possible, otherwise collected as a pending
// prompt. Returns resolved plain columns objects (same shape
// identifyGstColumns() already produces) ready to shallow-merge over the
// existing role-recognition columns, plus the list of anything still
// unresolved (non-empty => the caller must block).
async function mentorResolveGstColumns({ clientId, gstr2aSheet, booksSheet }) {
  const perSheet = [
    { key: "gstr2a", sheet: gstr2aSheet, role: "gstr2a" },
    { key: "books", sheet: booksSheet, role: "purchase_register" },
  ];
  const resolvedColumns = { gstr2a: {}, books: {} };
  const pendingPrompts = [];

  for (const { key, sheet, role } of perSheet) {
    const sheetSignals = sheet.result.sheetSignals;
    const inPlay = fieldsInPlay(role, sheet.result.columns.candidates);
    const allColumnIndices = sheetSignals.columns.map((c) => c.index);

    for (const item of inPlay) {
      const resolved = await resolveColumnField({
        clientId,
        sheetName: sheet.sheetName,
        sheetSignals,
        fieldName: item.field,
        candidateIndices: item.candidateIndices,
        allColumnIndices,
        store: mentorColumnMemoryStore,
      });

      if (resolved.status === "resolved_from_memory") {
        resolvedColumns[key][item.field] = resolved.fieldIndex;
        continue;
      }

      // Tier 2 didn't resolve -- try Tier 1 (shared software-pattern
      // memory) before queuing a prompt, but only for genuine ambiguity
      // (see this file's docstring for why zero_match_* reasons are
      // excluded). Scoped identically on the remember side below.
      let tier1Resolved = false;
      if (item.reason === "ambiguous") {
        const candidateHeaders = resolved.candidates.map((c) => c.header);
        const tier1 = await lookupColumnPattern(item.field, candidateHeaders);
        if (tier1.found) {
          const match = sheetSignals.columns.find((c) => normalizeHeaderText(c.header) === tier1.chosenHeader);
          if (match) {
            resolvedColumns[key][item.field] = match.index;
            tier1Resolved = true;
          }
          // else: no column on THIS sheet currently has the remembered
          // header (shape drifted) -- fall through to needs_input below,
          // never error and never guess.
        }
      }

      if (!tier1Resolved) {
        pendingPrompts.push({ sheetKey: key, clientId, sheetSignals, reason: item.reason, ...resolved });
      }
    }
  }

  return { blocked: pendingPrompts.length > 0, resolvedColumns, pendingPrompts };
}

// prompts includes every field still unresolved, INCLUDING ones the user
// dismissed earlier this session — dismissing stops the nagging, it must
// NOT make the caller (mentorScanForGstReconciliation) think the field
// resolved and merge in a guessed value. So filtering happens only here,
// on what actually gets RENDERED — the caller's blocked/pendingPrompts
// count (computed before this filter) stays accurate.
function mentorShowColumnMemoryQueue(prompts, onAllResolved, appendAuditLogRow, logAction) {
  mentorColumnMemoryOnAllResolved = onAllResolved;
  mentorColumnMemoryAppendAuditLogRow = appendAuditLogRow;
  mentorColumnMemoryLogAction = logAction;
  mentorColumnMemoryQueue = prompts.filter((p) => !mentorDismissedColumnMemoryThisSession.has(p.sheetName + "::" + p.fieldName));
  if (mentorColumnMemoryQueue.length === 0) {
    // Everything still unresolved was dismissed this session — stay
    // completely silent (no card at all) rather than showing an empty
    // prompt box. Reconciliation itself is still blocked; the caller
    // already returned before rendering any proposal card.
    mentorHideColumnMemoryQueue();
    return;
  }
  mentorRenderNextColumnMemoryPrompt();
}

function mentorHideColumnMemoryQueue() {
  mentorColumnMemoryQueue = [];
  mentorPendingColumnMemoryPrompt = null;
  const container = document.getElementById("mentor-column-memory-prompt");
  if (container) {
    container.style.display = "none";
    container.innerHTML = "";
  }
}

function mentorRenderNextColumnMemoryPrompt() {
  const container = document.getElementById("mentor-column-memory-prompt");
  if (!container) return;

  if (mentorColumnMemoryQueue.length === 0) {
    mentorHideColumnMemoryQueue();
    return;
  }

  mentorPendingColumnMemoryPrompt = mentorColumnMemoryQueue.shift();
  const p = mentorPendingColumnMemoryPrompt;
  const accent = "var(--red)"; // distinct red — this is a BLOCKING prompt (no proposal card until resolved), not an FYI

  let html = "<div style=\"font-family:'Inter','Segoe UI',sans-serif;background:#FFFFFF;border:1px solid var(--line);border-left:4px solid " + accent + ";padding:10px;border-radius:6px;font-size:12px;color:var(--ink);\">";
  html += "<div style='margin-bottom:8px;'>" + mentorColumnMemoryEscapeHtml(p.prompt) + "</div>";
  p.candidates.forEach((c, i) => {
    const preview = (c.sampleValues || []).slice(0, 5).map((v) => mentorColumnMemoryEscapeHtml(String(v))).join(", ");
    const title = "Hover to see this column in '" + p.sheetName + "', click to choose it";
    html +=
      "<div onclick='mentorSubmitColumnMemoryAnswer(" + i + ")' onmouseenter='mentorHoverColumnMemoryCandidate(" + i + ")' title=\"" + mentorColumnMemoryEscapeHtml(title) + "\" " +
      "style='cursor:pointer;padding:6px;margin-bottom:4px;border:1px solid var(--line);border-radius:4px;background:var(--paper);'>";
    html += "<strong class='mentor-mono' style=\"font-family:'JetBrains Mono','Consolas',monospace;color:var(--ink);\">" + mentorColumnMemoryEscapeHtml(c.header || "(blank header)") + "</strong><br/>";
    html += "<span class='mentor-mono' style=\"font-family:'JetBrains Mono','Consolas',monospace;color:var(--ink-soft);font-size:11px;\">e.g. " + preview + "</span></div>";
  });
  html += "<button onclick='mentorSkipColumnMemoryPrompt()' style=\"font-family:'Space Grotesk','Segoe UI',sans-serif;font-weight:500;margin-top:4px;padding:3px 10px;background:var(--ink-soft);color:var(--paper);border:none;border-radius:4px;cursor:pointer;font-size:11px;\">Skip for now</button>";
  html += "</div>";

  container.innerHTML = html;
  container.style.display = "block";
}

// Synchronous entry point (fired on onmouseenter) — captures sheetName and
// the candidate object at THIS exact moment, before handing off to the
// debounced highlighter. Explicit args (not re-reading module state after
// the debounce delay) is what makes rapid mouse movement across several
// candidates safe: whichever hover was last wins, and it always highlights
// the column that was actually under the pointer, never a stale one from
// after the prompt has since advanced.
window.mentorHoverColumnMemoryCandidate = function (candidateIdx) {
  const p = mentorPendingColumnMemoryPrompt;
  if (!p) return;
  const candidate = p.candidates[candidateIdx];
  if (!candidate) return;
  mentorDebouncedHighlightColumnMemoryCandidate(p.sheetName, candidate);
};

window.mentorSubmitColumnMemoryAnswer = async function (candidateIdx) {
  const p = mentorPendingColumnMemoryPrompt;
  if (!p) return;
  const chosen = p.candidates[candidateIdx];
  if (!chosen) return;

  // Not awaited — this is a bonus confirmation of what was just chosen, not
  // part of the actual save. Must never slow down or be able to block the
  // real submit flow below.
  mentorHighlightColumnMemoryCandidate(p.sheetName, chosen);

  try {
    console.log("MENTOR column-memory: submitting answer", { sheetName: p.sheetName, fieldName: p.fieldName, chosen });
    await rememberColumnField({ clientId: p.clientId, sheetName: p.sheetName, sheetSignals: p.sheetSignals, fieldName: p.fieldName, chosenColumnIndex: chosen.index, store: mentorColumnMemoryStore });

    // Tier 1 (shared) -- best-effort, fire-and-forget (not awaited), never
    // blocks or fails the Tier-2 save or the Audit Log write below. Same
    // "ambiguous only" scoping as the lookup side in
    // mentorResolveGstColumns -- a zero-match fallback answer (picked from
    // "every column on the sheet") isn't a real cross-client software
    // pattern worth sharing.
    if (p.reason === "ambiguous") {
      const candidateHeaders = p.candidates.map((c) => c.header);
      rememberColumnPattern(p.fieldName, candidateHeaders, chosen.header);
    }

    const description = "Learned that '" + p.sheetName + "'s \"" + p.fieldName + "\" column is \"" + chosen.header + "\" — won't ask again for a sheet with this shape";
    if (mentorColumnMemoryAppendAuditLogRow) {
      try {
        await Excel.run(async (context) => {
          await mentorColumnMemoryAppendAuditLogRow(context, p.sheetName, "(whole sheet)", description);
        });
        if (mentorColumnMemoryLogAction) mentorColumnMemoryLogAction(description);
      } catch (writeError) {
        console.error("MENTOR column-memory: failed to write Audit Log row", writeError);
        if (mentorColumnMemoryLogAction) {
          mentorColumnMemoryLogAction("Learned that '" + p.sheetName + "'s \"" + p.fieldName + "\" column is \"" + chosen.header + "\", but writing it to the Audit Log sheet failed: " + writeError.message);
        }
      }
    } else if (mentorColumnMemoryLogAction) {
      mentorColumnMemoryLogAction(description);
    }
  } catch (error) {
    console.error("MENTOR column-memory: failed to save answer", error);
  }

  if (mentorColumnMemoryQueue.length === 0) {
    mentorHideColumnMemoryQueue();
    if (mentorColumnMemoryOnAllResolved) await mentorColumnMemoryOnAllResolved(); // re-run the scan — now resolves silently from memory
  } else {
    mentorRenderNextColumnMemoryPrompt();
  }
};

window.mentorSkipColumnMemoryPrompt = function () {
  const p = mentorPendingColumnMemoryPrompt;
  if (p) mentorDismissedColumnMemoryThisSession.add(p.sheetName + "::" + p.fieldName);
  if (mentorColumnMemoryQueue.length === 0) {
    mentorHideColumnMemoryQueue();
  } else {
    mentorRenderNextColumnMemoryPrompt();
  }
};

// Exposed so the "review learned answers" panel (mentor-memory-review-
// ui.js) writes to the SAME Audit Log entry point mentorSubmitColumnMemory-
// Answer above uses, instead of duplicating the injected-callback dance.
// Unlike that flow, the review panel isn't necessarily reached via a live
// mentorScanForGstReconciliation() call, so mentorColumnMemoryAppendAuditLog-
// Row may never have been injected — falls back to a plain console log.
async function mentorColumnMemoryWriteAuditLog(sheetName, description) {
  if (!mentorColumnMemoryAppendAuditLogRow) {
    console.log("MENTOR column-memory: " + description + " (no Audit Log writer injected yet this session)");
    if (mentorColumnMemoryLogAction) mentorColumnMemoryLogAction(description);
    return;
  }
  try {
    await Excel.run(async (context) => {
      await mentorColumnMemoryAppendAuditLogRow(context, sheetName, "(whole sheet)", description);
    });
    if (mentorColumnMemoryLogAction) mentorColumnMemoryLogAction(description);
  } catch (writeError) {
    console.error("MENTOR column-memory: failed to write Audit Log row", writeError);
    if (mentorColumnMemoryLogAction) mentorColumnMemoryLogAction(description + " (Audit Log write failed: " + writeError.message + ")");
  }
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    mentorResolveGstColumns,
    mentorShowColumnMemoryQueue,
    mentorHideColumnMemoryQueue,
    COLUMN_MEMORY_SHEET_NAME,
    mentorColumnMemoryStore,
    mentorColumnMemoryWriteAuditLog,
  };
}
