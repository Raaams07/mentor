/*
 * sheet-label-pattern-store-test.js
 * -------------------------------------
 * Proves SheetLabelPatternStore's lookup/remember semantics: absent-
 * returns-not-found, remember-then-lookup resolves, remember-again
 * updates in place (last-write-wins, single record), the label is stored
 * verbatim (NOT normalized, unlike the column-pattern case — see the
 * module's own docstring for why), and the real file-backed mode actually
 * persists to disk.
 *
 * Run with: node server/sheet-label-pattern-store-test.js
 */

const fs = require("fs");
const os = require("os");
const path = require("path");
const { SheetLabelPatternStore } = require("./sheet-label-pattern-store.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

function runInMemory() {
  console.log("-- In-memory mode (filePath: null) --\n");
  const store = new SheetLabelPatternStore({ filePath: null });

  const absent = store.lookup("6:TEXT_UNIQUE,NUM_HIGH_SIGNED,NUM_HIGH_SIGNED,NUM_HIGH_SIGNED,TEXT_CATEGORICAL,TEXT_MODERATE");
  assert(absent.found === false, "lookup on an absent signature -> { found: false }");

  const signature = "5:TEXT_UNIQUE,TEXT_CATEGORICAL,NUM_HIGH_SIGNED,NUM_HIGH_SIGNED,NUM_HIGH_SIGNED";
  store.remember(signature, "Purchase register (Books)");
  const found = store.lookup(signature);
  assert(found.found === true && found.label === "Purchase register (Books)", "remember-then-lookup resolves to the exact label stored");

  store.remember(signature, "GST purchase register");
  const updated = store.lookup(signature);
  assert(updated.label === "GST purchase register", "remember again on the SAME signature updates in place (last-write-wins)");
  assert(Object.keys(store.records).length === 1, "still exactly one record for this signature, not a second accumulating entry");

  console.log("");
}

function runLabelStoredVerbatim() {
  console.log("-- Label is stored verbatim, not normalized (unlike column-pattern's header text) --\n");
  const store = new SheetLabelPatternStore({ filePath: null });
  const signature = "3:NUM_HIGH_SIGNED,NUM_HIGH_SIGNED,TEXT_UNIQUE";

  store.remember(signature, "  GSTIN/IGST/CGST/SGST Pivot Table  ");
  const result = store.lookup(signature);
  assert(result.label === "  GSTIN/IGST/CGST/SGST Pivot Table  ", "capitalization, punctuation, and whitespace are preserved exactly as passed in — this label gets shown to a future user verbatim, not just matched against internally");

  console.log("");
}

function runFileRoundTrip() {
  console.log("-- Real file round-trip (temp path, auto-created parent dir) --\n");
  const tempDir = path.join(os.tmpdir(), "mentor-sheet-label-pattern-store-test-" + Date.now());
  const tempFile = path.join(tempDir, "nested", "sheet-label-pattern-memory.json"); // nested -- proves mkdir -p behavior

  try {
    assert(!fs.existsSync(path.dirname(tempFile)), "parent directory genuinely doesn't exist yet before the first write");

    const signature = "4:TEXT_UNIQUE,DATE_TRANSACTIONAL,NUM_HIGH_SIGNED,TEXT_CATEGORICAL";
    const writer = new SheetLabelPatternStore({ filePath: tempFile });
    writer.remember(signature, "Bank statement export");

    assert(fs.existsSync(tempFile), "remember() creates the file (and its parent directory) on disk");

    const reader = new SheetLabelPatternStore({ filePath: tempFile }); // independent instance, same path
    const result = reader.lookup(signature);
    assert(result.found === true && result.label === "Bank statement export", "a SECOND, independent store instance pointed at the same file sees the persisted record -- proves it's real disk persistence, not just in-process state");

    const raw = JSON.parse(fs.readFileSync(tempFile, "utf8"));
    const key = Object.keys(raw)[0];
    assert(!("sheetName" in raw[key]) && !("clientId" in raw[key]) && !("sampleValues" in raw[key]), "the persisted record carries no sheet name, client identifier, or sample values -- just the structural signature and the label");
  } finally {
    fs.rmSync(tempDir, { recursive: true, force: true });
  }

  console.log("");
}

function run() {
  runInMemory();
  runLabelStoredVerbatim();
  runFileRoundTrip();

  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All sheet-label-pattern-store checks passed.");
  }
}

run();
