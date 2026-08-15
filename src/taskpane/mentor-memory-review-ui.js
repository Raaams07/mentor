/* global Excel */

/*
 * mentor-memory-review-ui.js
 * -------------------------------
 * User-facing "what has MENTOR learned about this workbook?" panel.
 *
 * Every previous way of fixing a wrong memory entry required unhiding
 * "MENTOR Sheet Memory" / "MENTOR Column Memory" and hand-editing cells —
 * not something a real pilot user would ever discover. This panel lists
 * every sheet-identity and column-mapping record MENTOR has stored for the
 * CURRENT workbook, in plain English, with Edit/Reset actions right there
 * in the sidebar. "Reset" doesn't try to guess a replacement — it just
 * deletes the record, so the next scan asks again through the normal
 * prompt (with real candidate headers/value previews), same as if MENTOR
 * had never seen this shape before.
 *
 * Reuses the SAME store instances mentor-sheet-memory-ui.js and mentor-
 * column-memory-ui.js already read/write from (imported from those
 * modules, not re-instantiated), and the SAME Audit Log write path they
 * use — so an edit made here is immediately consistent with what the next
 * scan resolves from memory, and shows up in the one Audit Log a user
 * already knows to check.
 *
 * Scope: per-workbook (Tier 2 / client-convention) memory only. The
 * shared Tier 1 software-pattern memory (server/column-pattern-store.js)
 * isn't scoped to "this workbook" — it's shared across every client that
 * talks to the local proxy — so a per-workbook review panel isn't the
 * right surface for editing it; nothing here touches it. A wrong Tier 1
 * entry still self-heals the normal way: the next ambiguous-column prompt
 * a user resolves differently overwrites it.
 *
 * Deliberately no "does this edit look right?" sanity check on save (unlike
 * the live sheet-memory prompt, which runs checkLabelPlausibility()): a
 * user opening this panel and typing a correction is already the
 * deliberate, looked-at-the-data case that check exists to catch mistakes
 * in — adding a second confirmation here would be friction without a
 * matching benefit.
 */

const { mentorSheetMemoryStore, mentorSheetMemoryWriteAuditLog } = require("./mentor-sheet-memory-ui.js");
const { mentorColumnMemoryStore, mentorColumnMemoryWriteAuditLog } = require("./mentor-column-memory-ui.js");
const { updateStoredSheetLabel, forgetSheetLabel } = require("../sheet-classifier/sheet-memory.js");
const { updateStoredColumnField, forgetColumnMemoryRecord, parseBlob } = require("../gst-reconciliation/column-memory.js");

let mentorMemoryReviewSheetRecords = [];
let mentorMemoryReviewColumnRecords = []; // each augmented with a parsed `.fields` object
// One in-progress edit/confirm at a time, or null. Shapes:
//   { kind: "edit-sheet", sig }
//   { kind: "confirm-reset-sheet", sig }
//   { kind: "edit-column-field", sig, field }
//   { kind: "confirm-reset-column-field", sig, field }
//   { kind: "confirm-reset-column-record", sig }
let mentorMemoryReviewUiState = null;

function mentorMemoryReviewEscapeHtml(text) {
  return String(text).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

// Same "one workbook file == one client" placeholder used by mentor-sheet-
// memory-ui.js / mentor-gst-reconciliation-ui.js — duplicated locally
// rather than imported, matching this codebase's existing convention of
// small trivial helpers living per-file instead of a shared util for one line.
function mentorGetMemoryReviewClientId(workbookName) {
  return "workbook:" + workbookName;
}

async function mentorRefreshMemoryReviewData() {
  try {
    let clientId;
    await Excel.run(async (context) => {
      context.workbook.load("name");
      await context.sync();
      clientId = mentorGetMemoryReviewClientId(context.workbook.name);
    });

    const [sheetRecords, columnRecords] = await Promise.all([mentorSheetMemoryStore.list(clientId), mentorColumnMemoryStore.list(clientId)]);

    mentorMemoryReviewSheetRecords = sheetRecords.slice().sort((a, b) => String(a.sheet_name_at_creation).localeCompare(String(b.sheet_name_at_creation)));
    mentorMemoryReviewColumnRecords = columnRecords
      .map((r) => ({ ...r, fields: parseBlob(r) }))
      .filter((r) => Object.keys(r.fields).length > 0) // shouldn't happen (an emptied blob is forgotten outright), but don't render a blank card if it ever does
      .sort((a, b) => String(a.sheet_name_at_creation).localeCompare(String(b.sheet_name_at_creation)));

    mentorMemoryReviewUiState = null;
    mentorRenderMemoryReviewPanel();
  } catch (error) {
    console.error("MENTOR memory review: failed to refresh", error);
  }
}

function mentorRenderMemoryReviewPanel() {
  const toggle = document.getElementById("mentor-memory-review-toggle");
  const container = document.getElementById("mentor-memory-review-list");
  if (!container) return;

  const totalCount = mentorMemoryReviewSheetRecords.length + mentorMemoryReviewColumnRecords.length;
  if (toggle) toggle.textContent = "Learned answers (" + totalCount + ")";

  if (totalCount === 0) {
    container.innerHTML = "<div style=\"font-family:'Inter','Segoe UI',sans-serif; color:var(--ink-soft); font-size:11px;\">Nothing learned yet for this workbook.</div>";
    return;
  }

  let html = "";
  if (mentorMemoryReviewSheetRecords.length > 0) {
    // "Sheet identities" plays the exact role mentor_landing.html's
    // .ledger-head (SUPPLIER / INVOICE / AMOUNT / STATUS) plays — a small
    // caption sitting above a list of rows — so it takes that style
    // verbatim (JetBrains Mono, uppercase, letter-spacing 0.04em), not the
    // Space Grotesk heading style, which is reserved for true headings.
    html += "<div class='mentor-label' style=\"font-family:'JetBrains Mono','Consolas',monospace;font-size:10px;color:var(--ink-soft);text-transform:uppercase;letter-spacing:0.04em;margin:6px 0 4px;\">Sheet identities</div>";
    html += "<div style=\"font-family:'Inter','Segoe UI',sans-serif;color:var(--ink-soft);font-size:10px;margin:-2px 0 6px;\">Recognized. Won't ask again.</div>";
    mentorMemoryReviewSheetRecords.forEach((r) => {
      html += mentorRenderSheetMemoryReviewRow(r);
    });
  }
  if (mentorMemoryReviewColumnRecords.length > 0) {
    html += "<div class='mentor-label' style=\"font-family:'JetBrains Mono','Consolas',monospace;font-size:10px;color:var(--ink-soft);text-transform:uppercase;letter-spacing:0.04em;margin:10px 0 4px;\">Column mappings</div>";
    html += "<div style=\"font-family:'Inter','Segoe UI',sans-serif;color:var(--ink-soft);font-size:10px;margin:-2px 0 6px;\">Columns matched to their fields.</div>";
    mentorMemoryReviewColumnRecords.forEach((r) => {
      html += mentorRenderColumnMemoryReviewRow(r);
    });
  }
  container.innerHTML = html;
}

// font-family + color:var(--ink) are load-bearing, not decorative — this
// card background never had either set explicitly, so every row's text was
// falling back to the browser default. Individual elements that set their
// own font/color (e.g. the mono ":var(--ink-soft)" sheet-name label) still
// correctly override this.
//
// Padding/border-radius/margin-bottom deliberately mirror mentor_landing.html's
// .card (border:1px solid var(--line);border-radius:10px;padding:26px;
// background:#fff — scaled down for the sidebar's ~360px width and 11-13px
// text) so each sheet-identity/shape record reads as its own distinct card,
// not a continuous run of text — the 12px margin-bottom is what actually
// separates one record from the next; a record's own internal element
// spacing (margin-bottom:4px between its own fields, etc.) is unrelated and
// left as-is.
const MENTOR_MEMORY_REVIEW_ROW_STYLE = "font-family:'Inter','Segoe UI',sans-serif;padding:14px;margin-bottom:12px;border:1px solid var(--line);border-radius:8px;background:#FFFFFF;color:var(--ink);";

function mentorRenderSheetMemoryReviewRow(record) {
  const sig = record.structural_signature;
  const state = mentorMemoryReviewUiState;

  if (state && state.kind === "confirm-reset-sheet" && state.sig === sig) {
    return (
      "<div style=\"" +
      MENTOR_MEMORY_REVIEW_ROW_STYLE +
      "\">" +
      "<div style='margin-bottom:6px;'>Forget that '" +
      mentorMemoryReviewEscapeHtml(record.sheet_name_at_creation) +
      "' is \"" +
      mentorMemoryReviewEscapeHtml(record.user_provided_label) +
      "\"? MENTOR will ask again next time it sees a sheet with this shape.</div>" +
      "<button onclick=\"mentorConfirmForgetSheetMemory('" +
      sig +
      "')\" style=\"font-family:'Space Grotesk','Segoe UI',sans-serif;font-weight:500;margin-right:6px;padding:3px 10px;background:var(--red);color:#FFFFFF;border:none;border-radius:4px;cursor:pointer;font-size:11px;\">Yes, forget it</button>" +
      "<button onclick=\"mentorCancelMemoryReviewAction()\" style=\"font-family:'Space Grotesk','Segoe UI',sans-serif;font-weight:500;padding:3px 10px;background:var(--ink-soft);color:var(--paper);border:none;border-radius:4px;cursor:pointer;font-size:11px;\">Cancel</button>" +
      "</div>"
    );
  }

  if (state && state.kind === "edit-sheet" && state.sig === sig) {
    return (
      "<div style=\"" +
      MENTOR_MEMORY_REVIEW_ROW_STYLE +
      "\">" +
      "<div class='mentor-mono' style=\"font-family:'JetBrains Mono','Consolas',monospace;color:var(--ink-soft);font-size:11px;margin-bottom:4px;\">'" +
      mentorMemoryReviewEscapeHtml(record.sheet_name_at_creation) +
      "'</div>" +
      "<input id='mentor-memory-review-edit-input' type='text' value=\"" +
      mentorMemoryReviewEscapeHtml(record.user_provided_label) +
      "\" style=\"font-family:'Inter','Segoe UI',sans-serif;width:100%;box-sizing:border-box;padding:5px 7px;border-radius:4px;border:1px solid var(--line);background:#FFFFFF;color:var(--ink);font-size:12px;margin-bottom:6px;\" " +
      "onkeydown='if(event.key===\"Enter\"){mentorSaveSheetMemoryReviewEdit(\"" +
      sig +
      "\");}' />" +
      "<button onclick=\"mentorSaveSheetMemoryReviewEdit('" +
      sig +
      "')\" style=\"font-family:'Space Grotesk','Segoe UI',sans-serif;font-weight:500;margin-right:6px;padding:3px 10px;background:var(--green-bg);color:var(--green);border:none;border-radius:4px;cursor:pointer;font-size:11px;\">Save</button>" +
      "<button onclick=\"mentorCancelMemoryReviewAction()\" style=\"font-family:'Space Grotesk','Segoe UI',sans-serif;font-weight:500;padding:3px 10px;background:var(--ink-soft);color:var(--paper);border:none;border-radius:4px;cursor:pointer;font-size:11px;\">Cancel</button>" +
      "</div>"
    );
  }

  return (
    "<div style=\"" +
    MENTOR_MEMORY_REVIEW_ROW_STYLE +
    "\">" +
    "<div><strong class='mentor-mono' style=\"font-family:'JetBrains Mono','Consolas',monospace;\">" +
    mentorMemoryReviewEscapeHtml(record.sheet_name_at_creation) +
    "</strong> — \"" +
    mentorMemoryReviewEscapeHtml(record.user_provided_label) +
    "\"</div>" +
    "<div style='margin-top:4px;'>" +
    "<a href=\"#\" onclick=\"mentorStartSheetMemoryReviewEdit('" +
    sig +
    "'); return false;\" style=\"font-family:'Inter','Segoe UI',sans-serif;font-size:11px;color:#9b59d0;text-decoration:underline;margin-right:10px;\">Edit</a>" +
    "<a href=\"#\" onclick=\"mentorStartMemoryReviewConfirmResetSheet('" +
    sig +
    "'); return false;\" style=\"font-family:'Inter','Segoe UI',sans-serif;font-size:11px;color:var(--red);text-decoration:underline;\">Reset</a>" +
    "</div></div>"
  );
}

function mentorRenderColumnMemoryReviewRow(record) {
  const sig = record.structural_signature;
  const state = mentorMemoryReviewUiState;

  let html = "<div style=\"" + MENTOR_MEMORY_REVIEW_ROW_STYLE + "\">";
  html += "<div class='mentor-mono' style=\"font-family:'JetBrains Mono','Consolas',monospace;color:var(--ink-soft);font-size:11px;margin-bottom:2px;\">'" + mentorMemoryReviewEscapeHtml(record.sheet_name_at_creation) + "' shape</div>";
  html += "<div style=\"font-family:'Inter','Segoe UI',sans-serif;color:var(--ink-soft);font-size:10px;margin-bottom:4px;\">Same layout recognized.</div>";

  Object.keys(record.fields).forEach((field) => {
    if (state && state.kind === "confirm-reset-column-field" && state.sig === sig && state.field === field) {
      html +=
        "<div style='margin-bottom:4px;'>Forget \"" +
        field +
        "\" → \"" +
        mentorMemoryReviewEscapeHtml(record.fields[field]) +
        "\"? " +
        "<button onclick=\"mentorConfirmForgetColumnField('" +
        sig +
        "','" +
        field +
        "')\" style=\"font-family:'Space Grotesk','Segoe UI',sans-serif;font-weight:500;margin-left:4px;margin-right:4px;padding:2px 8px;background:var(--red);color:#FFFFFF;border:none;border-radius:4px;cursor:pointer;font-size:10px;\">Yes</button>" +
        "<button onclick=\"mentorCancelMemoryReviewAction()\" style=\"font-family:'Space Grotesk','Segoe UI',sans-serif;font-weight:500;padding:2px 8px;background:var(--ink-soft);color:var(--paper);border:none;border-radius:4px;cursor:pointer;font-size:10px;\">Cancel</button></div>";
    } else if (state && state.kind === "edit-column-field" && state.sig === sig && state.field === field) {
      html +=
        "<div style='margin-bottom:4px;'><span class='mentor-mono' style=\"font-family:'JetBrains Mono','Consolas',monospace;color:var(--ink-soft);\">" +
        field +
        ":</span> " +
        "<input id='mentor-memory-review-edit-input' type='text' value=\"" +
        mentorMemoryReviewEscapeHtml(record.fields[field]) +
        "\" class='mentor-mono' style=\"font-family:'JetBrains Mono','Consolas',monospace;width:55%;box-sizing:border-box;padding:3px 5px;border-radius:4px;border:1px solid var(--line);background:#FFFFFF;color:var(--ink);font-size:11px;\" " +
        "onkeydown='if(event.key===\"Enter\"){mentorSaveColumnMemoryReviewEdit(\"" +
        sig +
        "\",\"" +
        field +
        "\");}' /> " +
        "<button onclick=\"mentorSaveColumnMemoryReviewEdit('" +
        sig +
        "','" +
        field +
        "')\" style=\"font-family:'Space Grotesk','Segoe UI',sans-serif;font-weight:500;padding:2px 8px;background:var(--green-bg);color:var(--green);border:none;border-radius:4px;cursor:pointer;font-size:10px;\">Save</button> " +
        "<button onclick=\"mentorCancelMemoryReviewAction()\" style=\"font-family:'Space Grotesk','Segoe UI',sans-serif;font-weight:500;padding:2px 8px;background:var(--ink-soft);color:var(--paper);border:none;border-radius:4px;cursor:pointer;font-size:10px;\">Cancel</button></div>";
    } else {
      html +=
        "<div class='mentor-mono' style=\"font-family:'JetBrains Mono','Consolas',monospace;margin-bottom:4px;color:var(--ink);\">" +
        field +
        " → \"" +
        mentorMemoryReviewEscapeHtml(record.fields[field]) +
        "\" " +
        "<a href=\"#\" onclick=\"mentorStartColumnMemoryReviewEdit('" +
        sig +
        "','" +
        field +
        "'); return false;\" style=\"font-family:'Inter','Segoe UI',sans-serif;font-size:11px;color:#9b59d0;text-decoration:underline;margin-left:6px;\">Edit</a> " +
        "<a href=\"#\" onclick=\"mentorStartMemoryReviewConfirmResetColumnField('" +
        sig +
        "','" +
        field +
        "'); return false;\" style=\"font-family:'Inter','Segoe UI',sans-serif;font-size:11px;color:var(--red);text-decoration:underline;margin-left:6px;\">Reset</a></div>";
    }
  });

  if (state && state.kind === "confirm-reset-column-record" && state.sig === sig) {
    html +=
      "<div style='margin-top:4px;'>Forget ALL column mappings learned for this sheet shape? " +
      "<button onclick=\"mentorConfirmForgetColumnRecord('" +
      sig +
      "')\" style=\"font-family:'Space Grotesk','Segoe UI',sans-serif;font-weight:500;margin-left:4px;margin-right:4px;padding:2px 8px;background:var(--red);color:#FFFFFF;border:none;border-radius:4px;cursor:pointer;font-size:10px;\">Yes</button>" +
      "<button onclick=\"mentorCancelMemoryReviewAction()\" style=\"font-family:'Space Grotesk','Segoe UI',sans-serif;font-weight:500;padding:2px 8px;background:var(--ink-soft);color:var(--paper);border:none;border-radius:4px;cursor:pointer;font-size:10px;\">Cancel</button></div>";
  } else {
    html += "<a href=\"#\" onclick=\"mentorStartMemoryReviewConfirmResetColumnRecord('" + sig + "'); return false;\" style=\"font-family:'Inter','Segoe UI',sans-serif;font-size:11px;color:var(--red);text-decoration:underline;\">Reset all for this shape</a>";
  }

  html += "</div>";
  return html;
}

window.mentorToggleMemoryReview = async function () {
  const list = document.getElementById("mentor-memory-review-list");
  if (!list) return;
  const opening = list.style.display === "none";
  list.style.display = opening ? "block" : "none";
  if (opening) await mentorRefreshMemoryReviewData();
};

window.mentorCancelMemoryReviewAction = function () {
  mentorMemoryReviewUiState = null;
  mentorRenderMemoryReviewPanel();
};

// ---- Sheet-identity actions ----

window.mentorStartSheetMemoryReviewEdit = function (sig) {
  mentorMemoryReviewUiState = { kind: "edit-sheet", sig };
  mentorRenderMemoryReviewPanel();
  const input = document.getElementById("mentor-memory-review-edit-input");
  if (input) {
    input.focus();
    input.select();
  }
};

window.mentorSaveSheetMemoryReviewEdit = async function (sig) {
  const input = document.getElementById("mentor-memory-review-edit-input");
  const value = input ? input.value.trim() : "";
  if (!value) {
    if (input) input.style.border = "1px solid var(--red)";
    return;
  }

  const record = mentorMemoryReviewSheetRecords.find((r) => r.structural_signature === sig);
  if (!record) return;
  const oldLabel = record.user_provided_label;

  if (value !== oldLabel) {
    try {
      await updateStoredSheetLabel({ clientId: record.client_id, structuralSignature: sig, newLabel: value, store: mentorSheetMemoryStore });
      await mentorSheetMemoryWriteAuditLog(
        record.sheet_name_at_creation,
        "Corrected learned sheet identity for '" + record.sheet_name_at_creation + "' from \"" + oldLabel + "\" to \"" + value + "\" (via review panel)"
      );
    } catch (error) {
      console.error("MENTOR memory review: failed to save sheet label edit", error);
    }
  }

  await mentorRefreshMemoryReviewData();
};

window.mentorStartMemoryReviewConfirmResetSheet = function (sig) {
  mentorMemoryReviewUiState = { kind: "confirm-reset-sheet", sig };
  mentorRenderMemoryReviewPanel();
};

window.mentorConfirmForgetSheetMemory = async function (sig) {
  const record = mentorMemoryReviewSheetRecords.find((r) => r.structural_signature === sig);
  if (!record) return;
  try {
    await forgetSheetLabel({ clientId: record.client_id, structuralSignature: sig, store: mentorSheetMemoryStore });
    await mentorSheetMemoryWriteAuditLog(
      record.sheet_name_at_creation,
      "Forgot the learned identity of '" + record.sheet_name_at_creation + "' (was \"" + record.user_provided_label + "\") via review panel — MENTOR will ask again"
    );
  } catch (error) {
    console.error("MENTOR memory review: failed to forget sheet label", error);
  }
  await mentorRefreshMemoryReviewData();
};

// ---- Column-mapping actions ----

window.mentorStartColumnMemoryReviewEdit = function (sig, field) {
  mentorMemoryReviewUiState = { kind: "edit-column-field", sig, field };
  mentorRenderMemoryReviewPanel();
  const input = document.getElementById("mentor-memory-review-edit-input");
  if (input) {
    input.focus();
    input.select();
  }
};

window.mentorSaveColumnMemoryReviewEdit = async function (sig, field) {
  const input = document.getElementById("mentor-memory-review-edit-input");
  const value = input ? input.value.trim() : "";
  if (!value) {
    if (input) input.style.border = "1px solid var(--red)";
    return;
  }

  const record = mentorMemoryReviewColumnRecords.find((r) => r.structural_signature === sig);
  if (!record) return;
  const oldValue = record.fields[field];

  if (value !== oldValue) {
    try {
      await updateStoredColumnField({ clientId: record.client_id, structuralSignature: sig, fieldName: field, newHeaderText: value, store: mentorColumnMemoryStore });
      await mentorColumnMemoryWriteAuditLog(
        record.sheet_name_at_creation,
        "Corrected learned column mapping for '" + record.sheet_name_at_creation + "'s \"" + field + "\" from \"" + oldValue + "\" to \"" + value + "\" (via review panel)"
      );
    } catch (error) {
      console.error("MENTOR memory review: failed to save column mapping edit", error);
    }
  }

  await mentorRefreshMemoryReviewData();
};

window.mentorStartMemoryReviewConfirmResetColumnField = function (sig, field) {
  mentorMemoryReviewUiState = { kind: "confirm-reset-column-field", sig, field };
  mentorRenderMemoryReviewPanel();
};

window.mentorConfirmForgetColumnField = async function (sig, field) {
  const record = mentorMemoryReviewColumnRecords.find((r) => r.structural_signature === sig);
  if (!record) return;
  const oldValue = record.fields[field];
  try {
    await updateStoredColumnField({ clientId: record.client_id, structuralSignature: sig, fieldName: field, newHeaderText: null, store: mentorColumnMemoryStore });
    await mentorColumnMemoryWriteAuditLog(
      record.sheet_name_at_creation,
      "Forgot learned column mapping for '" + record.sheet_name_at_creation + "'s \"" + field + "\" (was \"" + oldValue + "\") via review panel — MENTOR will ask again"
    );
  } catch (error) {
    console.error("MENTOR memory review: failed to forget column field", error);
  }
  await mentorRefreshMemoryReviewData();
};

window.mentorStartMemoryReviewConfirmResetColumnRecord = function (sig) {
  mentorMemoryReviewUiState = { kind: "confirm-reset-column-record", sig };
  mentorRenderMemoryReviewPanel();
};

window.mentorConfirmForgetColumnRecord = async function (sig) {
  const record = mentorMemoryReviewColumnRecords.find((r) => r.structural_signature === sig);
  if (!record) return;
  try {
    await forgetColumnMemoryRecord({ clientId: record.client_id, structuralSignature: sig, store: mentorColumnMemoryStore });
    await mentorColumnMemoryWriteAuditLog(
      record.sheet_name_at_creation,
      "Forgot ALL learned column mappings for '" + record.sheet_name_at_creation + "'s shape via review panel — MENTOR will ask again"
    );
  } catch (error) {
    console.error("MENTOR memory review: failed to forget column record", error);
  }
  await mentorRefreshMemoryReviewData();
};

// Called once at startup (mirrors mentorRefreshActionCount for History) so
// the toggle shows a real count immediately, not "(0)" until first opened.
async function mentorMemoryReviewInit() {
  await mentorRefreshMemoryReviewData();
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { mentorMemoryReviewInit };
}
