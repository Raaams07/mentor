/*
 * column-memory-test.js
 * ------------------------
 * Proves resolveColumnField()/rememberColumnField() in isolation, against
 * a plain in-memory BaseSheetMemoryStore (no Excel/DOM involved — its
 * default _load()/_persist() are no-ops over a plain array, exactly what's
 * needed for a fast, deterministic unit test).
 *
 * The most important scenario here is the storage-design regression test:
 * §3 of the implementation plan chose "one record per signature, JSON blob
 * of every field" specifically BECAUSE "one row per (signature, field)"
 * would silently let resolving field #2 overwrite field #1's row. That
 * failure mode is exactly what runTwoFieldsDontOverwrite() below checks.
 *
 * All data is synthetic/fictional.
 *
 * Run with: node src/gst-reconciliation/column-memory-test.js
 */

const { resolveColumnField, rememberColumnField } = require("./column-memory.js");
const { BaseSheetMemoryStore } = require("../sheet-classifier/sheet-memory-store-base.js");
const { extractSheetSignals } = require("../sheet-classifier/signal-extractor.js");

let failures = 0;
function pass(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

const CLIENT_ID = "workbook:test.xlsx";

function tallyLikeSignals(sheetName) {
  const values = [
    ["GSTIN", "Taxable Value", "Input CGST@9%", "Input CGST@14%", "Output CGST@9%", "CGST"],
    ["29AAAPL2356Q1Z8", 10000, 900, 1400, 100, 900],
  ];
  return extractSheetSignals(sheetName, values);
}

async function runNoStoredRecord() {
  console.log("-- No stored record: needs_input with real candidates --\n");
  const store = new BaseSheetMemoryStore();
  const sheetSignals = tallyLikeSignals("Books");

  const result = await resolveColumnField({
    clientId: CLIENT_ID,
    sheetName: "Books",
    sheetSignals,
    fieldName: "cgst",
    candidateIndices: [2, 3, 4, 5],
    allColumnIndices: [0, 1, 2, 3, 4, 5],
    store,
  });

  pass(result.status === "needs_input", "no record for this signature -> needs_input");
  pass(result.candidates.length === 4, "candidates list matches the 4 CGST-like columns passed in");
  pass(result.candidates[3].header === "CGST" && result.candidates[3].index === 5, "the real final CGST column is present among the candidates, with its header and index");
  pass(Array.isArray(result.candidates[0].sampleValues), "each candidate carries sampleValues for preview");

  console.log("");
}

async function runRememberThenResolve() {
  console.log("-- Remember, then re-resolve the SAME signature: resolves silently --\n");
  const store = new BaseSheetMemoryStore();
  const sheetSignals = tallyLikeSignals("Books");

  await rememberColumnField({ clientId: CLIENT_ID, sheetName: "Books", sheetSignals, fieldName: "cgst", chosenColumnIndex: 5, store });

  const result = await resolveColumnField({
    clientId: CLIENT_ID,
    sheetName: "Books",
    sheetSignals,
    fieldName: "cgst",
    candidateIndices: [2, 3, 4, 5],
    allColumnIndices: [0, 1, 2, 3, 4, 5],
    store,
  });

  pass(result.status === "resolved_from_memory", "second resolution for the same signature -> resolved_from_memory");
  pass(result.fieldIndex === 5, "resolves to the REMEMBERED column (5, real CGST), not the first-match guess (2)");
  pass(result.matchedVia === "exact", "matched via exact signature");

  console.log("");
}

async function runResolvesAcrossSecondStructurallyIdenticalSheet() {
  console.log("-- Same field remembered on sheet A resolves silently on a SECOND, structurally-identical sheet (different name) --\n");
  const store = new BaseSheetMemoryStore();
  const sheetA = tallyLikeSignals("Purchase Register Jan");
  const sheetB = tallyLikeSignals("Purchase Register Feb"); // same shape/headers, different name -- same structural signature

  await rememberColumnField({ clientId: CLIENT_ID, sheetName: "Purchase Register Jan", sheetSignals: sheetA, fieldName: "cgst", chosenColumnIndex: 5, store });

  const result = await resolveColumnField({
    clientId: CLIENT_ID,
    sheetName: "Purchase Register Feb",
    sheetSignals: sheetB,
    fieldName: "cgst",
    candidateIndices: [2, 3, 4, 5],
    allColumnIndices: [0, 1, 2, 3, 4, 5],
    store,
  });

  pass(result.status === "resolved_from_memory" && result.fieldIndex === 5, "a second workbook/sheet with the identical shape resolves from memory without asking again");

  console.log("");
}

async function runShapeDriftFuzzyMatch() {
  console.log("-- Shape drift within the fuzzy threshold still resolves; header text gone falls through to needs_input --\n");
  const store = new BaseSheetMemoryStore();
  const original = tallyLikeSignals("Books");
  await rememberColumnField({ clientId: CLIENT_ID, sheetName: "Books", sheetSignals: original, fieldName: "cgst", chosenColumnIndex: 5, store });

  // One extra unrelated text column appended -- close enough to fuzzy-match
  // (small column-count drift, same header wording throughout otherwise).
  const driftedValues = [
    ["GSTIN", "Taxable Value", "Input CGST@9%", "Input CGST@14%", "Output CGST@9%", "CGST", "Remarks"],
    ["29AAAPL2356Q1Z8", 10000, 900, 1400, 100, 900, "note"],
  ];
  const drifted = extractSheetSignals("Books v2", driftedValues);

  const result = await resolveColumnField({
    clientId: CLIENT_ID,
    sheetName: "Books v2",
    sheetSignals: drifted,
    fieldName: "cgst",
    candidateIndices: [2, 3, 4, 5],
    allColumnIndices: [0, 1, 2, 3, 4, 5, 6],
    store,
  });
  pass(result.status === "resolved_from_memory" && result.matchedVia === "fuzzy" && result.fieldIndex === 5, "structurally-drifted-but-similar sheet resolves via fuzzy match, still to the real CGST column");

  // Now rename the remembered header itself out of existence -- the fuzzy
  // record still matches structurally, but the stored header text no
  // longer appears anywhere on this sheet, so it must fall through, not
  // silently point at the wrong column or error.
  const renamedValues = [
    ["GSTIN", "Taxable Value", "Input CGST@9%", "Input CGST@14%", "Output CGST@9%", "Final CGST Column", "Remarks"],
    ["29AAAPL2356Q1Z8", 10000, 900, 1400, 100, 900, "note"],
  ];
  const renamed = extractSheetSignals("Books v3", renamedValues);
  const result2 = await resolveColumnField({
    clientId: CLIENT_ID,
    sheetName: "Books v3",
    sheetSignals: renamed,
    fieldName: "cgst",
    candidateIndices: [2, 3, 4, 5],
    allColumnIndices: [0, 1, 2, 3, 4, 5, 6],
    store,
  });
  pass(result2.status === "needs_input", "stored header text no longer exists on this (fuzzy-matched) sheet -> falls through to needs_input, not an error or a wrong guess");

  console.log("");
}

async function runTwoFieldsDontOverwrite() {
  console.log("-- REGRESSION: two different fields remembered for the SAME signature don't overwrite each other (the reason for the JSON-blob storage design) --\n");
  const store = new BaseSheetMemoryStore();
  const sheetSignals = tallyLikeSignals("Books");

  await rememberColumnField({ clientId: CLIENT_ID, sheetName: "Books", sheetSignals, fieldName: "gstin", chosenColumnIndex: 0, store });
  await rememberColumnField({ clientId: CLIENT_ID, sheetName: "Books", sheetSignals, fieldName: "cgst", chosenColumnIndex: 5, store });

  const gstinResolved = await resolveColumnField({ clientId: CLIENT_ID, sheetName: "Books", sheetSignals, fieldName: "gstin", candidateIndices: [0], allColumnIndices: [0, 1, 2, 3, 4, 5], store });
  const cgstResolved = await resolveColumnField({ clientId: CLIENT_ID, sheetName: "Books", sheetSignals, fieldName: "cgst", candidateIndices: [2, 3, 4, 5], allColumnIndices: [0, 1, 2, 3, 4, 5], store });

  pass(gstinResolved.status === "resolved_from_memory" && gstinResolved.fieldIndex === 0, "gstin, remembered FIRST, still resolves correctly...");
  pass(cgstResolved.status === "resolved_from_memory" && cgstResolved.fieldIndex === 5, "...AND cgst, remembered SECOND for the same signature, also resolves correctly -- neither overwrote the other");

  const records = await store.list(CLIENT_ID);
  pass(records.length === 1, "both fields live on a SINGLE record for this signature, not two separate rows");

  console.log("");
}

async function run() {
  await runNoStoredRecord();
  await runRememberThenResolve();
  await runResolvesAcrossSecondStructurallyIdenticalSheet();
  await runShapeDriftFuzzyMatch();
  await runTwoFieldsDontOverwrite();

  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All column-memory checks passed.");
  }
}

run();
