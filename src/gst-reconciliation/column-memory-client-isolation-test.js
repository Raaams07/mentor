/*
 * column-memory-client-isolation-test.js
 * ------------------------------------------
 * Proves the two-tier design's core sharing/privacy guarantee end-to-end:
 * Tier 2 (per-client) never leaks a resolution from one client to
 * another, while Tier 1 (shared software-pattern) genuinely does share
 * across clients — exactly the Part 2 brief's "software pattern, not
 * client data" distinction, verified with two distinct clientIds rather
 * than the single-clientId scenarios column-memory-test.js already covers.
 *
 * A SINGLE Tier-2 store instance is used for both clients (not two
 * separate instances) — this correctly simulates real isolation, since
 * BaseSheetMemoryStore.findExact/findSimilar/remember filter by the
 * clientId PARAMETER itself (r.client_id === clientId), not by which
 * store instance is asked; one instance (or, live, one hidden workbook
 * sheet, or a future shared backend) already isolates every client this
 * way. A SINGLE ColumnPatternStore instance is shared between both
 * clients too — simulating one proxy shared by every client, exactly per
 * the Part 2 design.
 *
 * All data is synthetic/fictional.
 *
 * Run with: node src/gst-reconciliation/column-memory-client-isolation-test.js
 */

const { resolveColumnField, rememberColumnField } = require("./column-memory.js");
const { BaseSheetMemoryStore } = require("../sheet-classifier/sheet-memory-store-base.js");
const { ColumnPatternStore } = require("../../server/column-pattern-store.js");
const { extractSheetSignals } = require("../sheet-classifier/signal-extractor.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

const CLIENT_A = "workbook:ClientA.xlsx";
const CLIENT_B = "workbook:ClientB.xlsx";

// The same Tally-like ambiguous-CGST shape used elsewhere in this suite
// (gst-column-identifier-test.js, column-memory-test.js) — both "clients"
// happen to use the same export software, hence the identical shape.
function tallyLikeSignals(sheetName) {
  const values = [
    ["GSTIN", "Taxable Value", "Input CGST@9%", "Input CGST@14%", "Output CGST@9%", "CGST"],
    ["29AAAPL2356Q1Z8", 10000, 900, 1400, 100, 900],
  ];
  return extractSheetSignals(sheetName, values);
}

async function run() {
  console.log("-- Client A resolves the ambiguous CGST field (Tier 2 answer + Tier 1 share) --\n");

  const tier2Store = new BaseSheetMemoryStore(); // ONE instance, shared by both clientIds below
  const tier1Store = new ColumnPatternStore({ filePath: null }); // ONE instance, shared by both clients — simulates one proxy

  const sheetSignalsA = tallyLikeSignals("Books");
  const candidateHeaders = ["Input CGST@9%", "Input CGST@14%", "Output CGST@9%", "CGST"];

  // Client A answers the prompt: real CGST column is index 5.
  await rememberColumnField({ clientId: CLIENT_A, sheetName: "Books", sheetSignals: sheetSignalsA, fieldName: "cgst", chosenColumnIndex: 5, store: tier2Store });
  // Ambiguous answers write both tiers on the live path (mentor-column-memory-ui.js) — mirrored here.
  tier1Store.remember("cgst", candidateHeaders, "CGST");

  console.log("\n-- Client B (never answered anything) hits the identical shape --\n");

  const sheetSignalsB = tallyLikeSignals("Purchase Register"); // different sheet name, same shape -- same structural signature
  const tier2ResultForB = await resolveColumnField({
    clientId: CLIENT_B,
    sheetName: "Purchase Register",
    sheetSignals: sheetSignalsB,
    fieldName: "cgst",
    candidateIndices: [2, 3, 4, 5],
    allColumnIndices: [0, 1, 2, 3, 4, 5],
    store: tier2Store,
  });
  assert(tier2ResultForB.status === "needs_input", "Client A's Tier-2 (per-client) answer does NOT resolve Client B's lookup -- Tier 2 correctly stays isolated per client");

  const tier1ResultForB = tier1Store.lookup("cgst", candidateHeaders);
  assert(tier1ResultForB.found === true && tier1ResultForB.chosenHeader === "cgst", "Tier 1 (shared) DOES resolve for Client B, using Client A's answer -- proves the software-pattern layer genuinely shares across clients");

  console.log("\n-- Client B independently answers the SAME prompt (own Tier-2 record, doesn't touch Client A's) --\n");

  await rememberColumnField({ clientId: CLIENT_B, sheetName: "Purchase Register", sheetSignals: sheetSignalsB, fieldName: "cgst", chosenColumnIndex: 5, store: tier2Store });

  const tier2ResultForAAfterB = await resolveColumnField({
    clientId: CLIENT_A,
    sheetName: "Books",
    sheetSignals: sheetSignalsA,
    fieldName: "cgst",
    candidateIndices: [2, 3, 4, 5],
    allColumnIndices: [0, 1, 2, 3, 4, 5],
    store: tier2Store,
  });
  assert(tier2ResultForAAfterB.status === "resolved_from_memory" && tier2ResultForAAfterB.fieldIndex === 5, "Client A's OWN answer is still intact and independently resolvable after Client B wrote to the same shared Tier-2 store instance");

  const allTier2Records = await tier2Store.list(CLIENT_A);
  assert(allTier2Records.length === 1, "tier2Store.list(CLIENT_A) sees only Client A's own record, not Client B's, even though both live in the same store instance");

  console.log("");
}

run().then(() => {
  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All column-memory client-isolation checks passed.");
  }
});
