/*
 * column-pattern-store-test.js
 * -------------------------------
 * Proves ColumnPatternStore's lookup/remember semantics: absent-returns-
 * not-found, remember-then-lookup resolves, remember-again updates in
 * place (last-write-wins, single record — not a growing log), and the
 * real file-backed mode actually persists to disk (a second, independent
 * store instance pointed at the same path sees the record) with the
 * parent directory auto-created when missing.
 *
 * Run with: node server/column-pattern-store-test.js
 */

const fs = require("fs");
const os = require("os");
const path = require("path");
const { ColumnPatternStore } = require("./column-pattern-store.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

function runInMemory() {
  console.log("-- In-memory mode (filePath: null) --\n");
  const store = new ColumnPatternStore({ filePath: null });

  const absent = store.lookup("cgst", ["Input CGST@9%", "CGST"]);
  assert(absent.found === false, "lookup on an absent pattern -> { found: false }");

  store.remember("cgst", ["Input CGST@9%", "CGST", "Input CGST@14%", "Output CGST @ 9%"], "CGST");
  const found = store.lookup("cgst", ["CGST", "Output CGST @ 9%", "Input CGST@14%", "Input CGST@9%"]); // reordered on purpose
  assert(found.found === true && found.chosenHeader === "cgst", "remember-then-lookup (reordered candidate list) resolves to the normalized chosen header");

  store.remember("cgst", ["Input CGST@9%", "CGST", "Input CGST@14%", "Output CGST @ 9%"], "Output CGST @ 9%");
  const updated = store.lookup("cgst", ["Input CGST@9%", "CGST", "Input CGST@14%", "Output CGST @ 9%"]);
  assert(updated.chosenHeader === "output cgst 9", "remember again on the SAME pattern with a different answer updates in place (last-write-wins)");
  assert(Object.keys(store.records).length === 1, "still exactly one record for this pattern, not a second accumulating entry");

  console.log("");
}

function runFileRoundTrip() {
  console.log("-- Real file round-trip (temp path, auto-created parent dir) --\n");
  const tempDir = path.join(os.tmpdir(), "mentor-column-pattern-store-test-" + Date.now());
  const tempFile = path.join(tempDir, "nested", "column-pattern-memory.json"); // nested -- proves mkdir -p behavior

  try {
    assert(!fs.existsSync(path.dirname(tempFile)), "parent directory genuinely doesn't exist yet before the first write");

    const writer = new ColumnPatternStore({ filePath: tempFile });
    writer.remember("gstin", ["Recipient GSTIN", "Supplier GSTIN"], "Supplier GSTIN");

    assert(fs.existsSync(tempFile), "remember() creates the file (and its parent directory) on disk");

    const reader = new ColumnPatternStore({ filePath: tempFile }); // independent instance, same path
    const result = reader.lookup("gstin", ["Supplier GSTIN", "Recipient GSTIN"]);
    assert(result.found === true && result.chosenHeader === "supplier gstin", "a SECOND, independent store instance pointed at the same file sees the persisted record -- proves it's real disk persistence, not just in-process state");

    const raw = JSON.parse(fs.readFileSync(tempFile, "utf8"));
    const key = Object.keys(raw)[0];
    assert(!("sampleValues" in raw[key]) && !("sheetName" in raw[key]) && !("clientId" in raw[key]), "the persisted record carries no sample values, sheet names, or client identifiers -- header text and field name only");
  } finally {
    fs.rmSync(tempDir, { recursive: true, force: true });
  }

  console.log("");
}

function run() {
  runInMemory();
  runFileRoundTrip();

  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All column-pattern-store checks passed.");
  }
}

run();
