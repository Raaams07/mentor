/*
 * gst-validation.js
 * -------------------
 * The standing validation pass — every "does the math/shape actually make
 * sense" check that runs automatically on EVERY reconciliation, regardless
 * of which GST workflow is running. GSTR-2A vs Books routes through this
 * today; GSTR-1 vs GSTR-3B and GSTR-3B vs Books call the SAME
 * runStandingValidators() going forward, not their own copies — that's the
 * whole point of this module existing separately from gst-report-writer.js
 * (which is 2A/Books-specific report layout) and rate-mismatch-detector.js
 * (which is the one check currently registered here, not a workflow).
 *
 * Deliberately separate from column-ambiguity resolution (column-memory.js
 * / gst-column-ambiguity-rules.js) — that decides WHICH column plays a
 * role; this checks whether the numbers in the column it settled on
 * actually add up, catching a wrong pick even when header-matching was
 * fully confident. This distinction is exactly what the historical
 * incident this module exists for was: a real client file (Basai Steel
 * Traders) where CGST/SGST got summed from intermediate per-rate-tier
 * columns (e.g. "Output CGST@9%") instead of the final consolidated
 * column, understating claimable ITC by several lakh rupees. Column-
 * ambiguity resolution alone can't catch that class of bug — a rate-tier
 * column name still matches "cgst" vocabulary confidently. Only checking
 * the actual arithmetic (taxable value x rate% ~= tax picked up) does.
 *
 * VALIDATORS: an array of { name, run(source) }. source is a plain
 * { sheetName, role, values, headerRowIndex, columns } object — the SAME
 * shape rcmBySource/itcBySource/dupBySource/rateMismatchBySource already
 * use throughout gst-report-writer.js and mentor-gst-reconciliation-ui.js.
 * Each run() returns whatever shape that specific check produces
 * ({ applicable, flagged, reason? }, matching rate-mismatch-detector.js's
 * own return shape) — runStandingValidators() doesn't interpret or
 * transform results, it only fans a source list out to every registered
 * validator and hands back one `...bySource`-shaped array per validator,
 * so existing callers (writeRateMismatchSheet, sumRateMismatchAmount,
 * buildStepTwoTotals) need ZERO changes to their own shape assumptions.
 *
 * Adding a second standing validator later (for the new workflows, or a
 * future 2A/Books one) means adding one entry to VALIDATORS below — not a
 * new call site scattered into mentor-gst-reconciliation-ui.js the way
 * rate-mismatch used to be before this module existed.
 */

const { detectRateMismatch } = require("./rate-mismatch-detector.js");

const VALIDATORS = [
  {
    name: "rateMismatch",
    run: (source) => detectRateMismatch(source.values, source.headerRowIndex, source.columns),
  },
];

// sources: [{ sheetName, role, values, headerRowIndex, columns }, ...] —
// typically the workflow's recognized sheets (e.g. gstr2a + books), but
// genuinely generic: nothing here reads sheetName/role, only values/
// headerRowIndex/columns, which every registered validator's run()
// actually consumes.
//
// Returns { [validatorName]: [{ ...source, result }, ...], ... } — one
// key per registered validator, each an array in the same `...bySource`
// shape this codebase already threads through report-writing and totals.
function runStandingValidators(sources) {
  const resultsByValidator = {};
  for (const validator of VALIDATORS) {
    resultsByValidator[validator.name] = sources.map((source) => ({ ...source, result: validator.run(source) }));
  }
  return resultsByValidator;
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { runStandingValidators, VALIDATORS };
}
