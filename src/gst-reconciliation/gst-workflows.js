/*
 * gst-workflows.js
 * -----------------
 * Registry of known GST reconciliation WORKFLOWS — each one a pair (or
 * set) of roles that, together, describe a specific comparison GST law
 * requires or that standard compliance practice performs.
 *
 * This is the piece that makes "GST reconciliation" more than one
 * hardcoded comparison: sheet recognition answers "what role is this
 * sheet playing," and this registry answers "given the roles present in
 * this workbook, which reconciliation workflow (if any) can actually be
 * run." Adding GSTR-1 vs GSTR-3B later means registering a new entry here
 * plus a new role pair in gst-role-recognizer.js and a new reconciler file
 * — this file and the recognizer never need to change for that addition,
 * and neither does gstr2a-vs-books-reconciler.js.
 */

const GST_WORKFLOWS = {
  gstr2a_vs_books: {
    label: "GSTR-2A vs Books (ITC reconciliation)",
    // GST law requires ITC claimed in a taxpayer's books to be supportable
    // by what suppliers have actually reported (CGST Rule 36(4) / Section
    // 16 conditions for availing input tax credit) — this is the standard
    // first-pass check for that: does the same supplier's GSTIN show the
    // same taxable value and tax in both the government's 2A and the
    // buyer's own purchase register.
    roles: { gstr2a: "gstr2a", purchaseRegister: "purchase_register" },
  },
  // Future extension point — NOT implemented yet:
  // gstr1_vs_gstr3b: {
  //   label: "GSTR-1 vs GSTR-3B (outward supply reconciliation)",
  //   roles: { gstr1: "gstr1", gstr3b: "gstr3b" },
  // },
};

// roleResultsBySheet: { sheetName: recognizeGstRole() result, ... }
// Returns the first fully-satisfied workflow — i.e. every role it requires
// is played by exactly one recognized sheet — or null if none is complete
// yet (e.g. only one of the two required sheets is present so far).
function identifyGstWorkflow(roleResultsBySheet) {
  const sheetNamesByRole = {};
  for (const [sheetName, result] of Object.entries(roleResultsBySheet)) {
    if (!result || result.role === "unknown") continue;
    // If two sheets recognize as the same role, prefer whichever scored
    // higher confidence — ambiguous, but a workflow needs exactly one
    // sheet per role to run at all.
    const existing = sheetNamesByRole[result.role];
    if (!existing || (result.confidence === "high" && existing.confidence !== "high")) {
      sheetNamesByRole[result.role] = { sheetName, confidence: result.confidence, result };
    }
  }

  for (const [workflowId, workflow] of Object.entries(GST_WORKFLOWS)) {
    const requiredRoles = Object.values(workflow.roles);
    const allPresent = requiredRoles.every((role) => sheetNamesByRole[role]);
    if (!allPresent) continue;

    const sheets = {};
    for (const [key, role] of Object.entries(workflow.roles)) {
      sheets[key] = { sheetName: sheetNamesByRole[role].sheetName, result: sheetNamesByRole[role].result };
    }
    return { workflowId, label: workflow.label, sheets };
  }

  return null;
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { GST_WORKFLOWS, identifyGstWorkflow };
}
