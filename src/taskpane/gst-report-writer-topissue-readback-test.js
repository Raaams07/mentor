/*
 * gst-report-writer-topissue-readback-test.js
 * --------------------------------------------------
 * Unit tests for the Top-Issues "read back from already-generated sheets"
 * logic (readGstTopIssuesFromExistingSheets and its per-row helpers) in
 * gst-report-writer.js — the fix for the persistent summary card
 * disappearing after the source-data-hash "stay silent" change.
 *
 * These test the PURE row-parsing functions directly with synthetic rows
 * matching the real column layouts each write*Sheet function uses (see
 * VENDOR_HEADERS and each sheet's own headers array in gst-report-writer.js)
 * — no Excel dependency, no live workbook needed. The Excel-dependent
 * orchestration (readGstTopIssuesFromExistingSheets itself, which reads
 * real sheets via context) still needs live verification, same limitation
 * as every other Excel-side piece of this project.
 *
 * Run with: node src/taskpane/gst-report-writer-topissue-readback-test.js
 */

const {
  topIssueFromVendorRow,
  topIssueFromExtraInvoiceRow,
  topIssueFromPossibleMatchRow,
  topIssueFromWrongHeadRow,
  topIssueFromCrossSheetWrongHeadRow,
  topIssueFromRcmRow,
  topIssueFromIneligibleItcRow,
  topIssuesFromDuplicateDataRows,
  isDataRow,
  findGstSheetHeaderRowIndex,
  findWrongHeadSectionHeader,
} = require("./gst-report-writer.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

function runTests() {
  console.log("-- Mismatch (still vendor-level) --\n");

  // VENDOR_HEADERS layout: GSTIN(0), 2A TV/IGST/CGST/SGST(1-4), Books TV/IGST/CGST/SGST(5-8), Diff IGST/CGST/SGST(9-11)
  const mismatchRow = ["27ABCDE1234F1Z5", 100000, 18000, 0, 0, 90000, 16000, 0, 0, 100, 0, 0];
  const mm = topIssueFromVendorRow(mismatchRow);
  assert(mm.amount === 100, "Mismatch amount = |sum of diff columns| (100+0+0)");
  assert(mm.category === "Mismatch" && mm.tier === 2, "correct category/tier");

  console.log("\n-- Extra in Books / Extra in 2A (now invoice-level, matched by GSTIN + Invoice/Voucher Number) --\n");

  // headers: GSTIN(0), Invoice/Voucher Number(1), Row(2), Taxable Value(3), IGST(4), CGST(5), SGST(6)
  const extraInBooksRow = ["27ABCDE1234F1Z5", "INV-9", 12, 50000, 6000, 4500, 4500];
  const eib = topIssueFromExtraInvoiceRow(extraInBooksRow, "extraInBooks");
  assert(eib.amount === 15000, "Extra in Books amount = IGST+CGST+SGST (6000+4500+4500)");
  assert(eib.category === "Extra in Books" && eib.tier === 2, "correct category/tier");
  assert(eib.reason.includes("supplier"), "reason references the same invoice-level explanation used at write time");

  const extraIn2ARow = ["27ABCDE1234F1Z5", "INV-10", 30, 80000, 8000, 6000, 6000];
  const ei2a = topIssueFromExtraInvoiceRow(extraIn2ARow, "extraIn2A");
  assert(ei2a.amount === 20000, "Extra in 2A amount = IGST+CGST+SGST (8000+6000+6000)");
  assert(ei2a.category === "Extra in 2A" && ei2a.tier === 2, "correct category/tier");

  // The exact real-world pattern this whole invoice-level rewrite fixes:
  // a vendor with several invoices where only ONE matches should surface
  // each unmatched invoice as its OWN row/topIssue, not collapse into one
  // vendor-level figure.
  const partiallyMatchedVendorRows = [
    ["36AAACU2414K1ZG", "INV-A", 1, 10000, 0, 900, 900],
    ["36AAACU2414K1ZG", "INV-B", 2, 12000, 0, 1080, 1080],
    ["36AAACU2414K1ZG", "INV-C", 3, 8000, 0, 720, 720],
    ["36AAACU2414K1ZG", "INV-D", 4, 15000, 0, 1350, 1350],
  ];
  const perInvoiceIssues = partiallyMatchedVendorRows.map((row) => topIssueFromExtraInvoiceRow(row, "extraInBooks"));
  assert(perInvoiceIssues.length === 4, "4 separate invoice-level issues for the same GSTIN, not one collapsed vendor-level issue");
  assert(
    perInvoiceIssues.reduce((sum, i) => sum + i.amount, 0) === 900 + 900 + 1080 + 1080 + 720 + 720 + 1350 + 1350,
    "the 4 invoices' amounts sum correctly and independently — each invoice is its own item"
  );

  console.log("\n-- Possible Matches (Fix 4: amount+GSTIN fallback, no invoice-number match) --\n");

  // headers: GSTIN(0), 2A Invoice/Voucher Number(1), 2A Row(2), Books Invoice/Voucher Number(3), Books Row(4), 2A Taxable Value(5), 2A IGST(6), 2A CGST(7), 2A SGST(8), Books Taxable Value(9), Books IGST(10), Books CGST(11), Books SGST(12), Reason(13)
  const possibleMatchRow = ["36AAOFA8281B1ZE", "2018", 40, "1997", 88, 165738.16, 0, 14916.43, 14916.43, 165737.5, 0, 14916.38, 14916.38, "Invoice numbers don't match ('2018' vs '1997')"];
  const pm = topIssueFromPossibleMatchRow(possibleMatchRow);
  assert(pm.amount === 29832.86, "amount = 2A-side IGST+CGST+SGST (0+14916.43+14916.43)");
  assert(pm.category === "Possible Match" && pm.tier === 3, "correct category and tier -- low-confidence, ranked below confirmed Extra/Mismatch (tier 2)");
  assert(pm.reason.includes("2018") && pm.reason.includes("1997"), "reason names both mismatched identifiers, read straight off the sheet's own Reason column");

  console.log("\n-- Wrong Head (single-sheet section) --\n");

  // headers: Row(0), GSTIN(1), Invoice Number(2), Issue(3), Expected Head(4), Actual Head(5), Supplier State(6), PoS State(7), IGST(8), CGST(9), SGST(10)
  const wrongHeadRow = [5, "27ABCDE1234F1Z5", "INV-100", "wrong_head", "CGST+SGST", "IGST", "27", "27", 9000, 0, 0];
  const wh = topIssueFromWrongHeadRow(wrongHeadRow);
  assert(wh.amount === 9000, "Wrong Head amount = IGST+CGST+SGST (9000+0+0)");
  assert(wh.reason.includes("IGST") && wh.reason.includes("CGST+SGST"), "reason mentions both actual and expected heads");
  assert(wh.category === "Wrong Head" && wh.tier === 3, "correct category/tier");

  console.log("\n-- Wrong Head (cross-sheet section) — the real case this whole feature was built for --\n");

  // headers: GSTIN(0), Invoice/Voucher Number(1), 2A Row(2), Books Row(3), 2A Head(4), Books Head(5), Expected Head(6), Incorrect Side(7), 2A IGST/CGST/SGST(8-10), Books IGST/CGST/SGST(11-13)
  const crossRow = ["37AABCR0435L1ZD", "211101051790", 1443, 1600, "IGST", "CGST+SGST", "IGST", "Books", 345663, 0, 0, 0, 172831.5, 172831.5];
  const cross = topIssueFromCrossSheetWrongHeadRow(crossRow);
  assert(cross.amount === 345663, "cross-sheet amount = 2A IGST+CGST+SGST (345663), matching the real-file validated case");
  assert(cross.reason.includes("211101051790") && cross.reason.includes("Books"), "reason names the specific invoice and the incorrect side");

  console.log("\n-- RCM (materiality tiering) --\n");

  // headers: Source Sheet(0), Row(1), GSTIN(2), Invoice/Voucher Number(3), Taxable Value(4), IGST(5), CGST(6), SGST(7), Category(8), Notification(9), Basis(10), Needs Confirmation?(11)
  const bigRcmRow = ["2A", 5, "27ABCDE1234F1Z5", "INV-1", 100000, 18000, 0, 0, "GTA Freight", "Notification 13/2017", "reverse_charge_flag", "No"];
  const smallRcmRow = ["Books", 12, "27ABCDE1234F1Z5", "INV-2", 1000, 180, 0, 0, "Legal Services", "Notification 13/2017", "keyword_match", "No"];
  const bigRcm = topIssueFromRcmRow(bigRcmRow, 5000);
  const smallRcm = topIssueFromRcmRow(smallRcmRow, 5000);
  assert(bigRcm.amount === 18000 && bigRcm.tier === 1, "RCM amount >= materiality threshold -> tier 1");
  assert(smallRcm.amount === 180 && smallRcm.tier === 3, "RCM amount below materiality threshold -> tier 3, not silently dropped");
  assert(bigRcm.reason.includes("GTA Freight") && bigRcm.reason.includes("2A"), "reason includes category and source sheet");

  console.log("\n-- Ineligible ITC (materiality tiering) --\n");

  const itcRow = ["Books", 8, "27ABCDE1234F1Z5", "V-45", 50000, 0, 4500, 4500, "Motor Vehicles", "Section 17(5)(a)/(ab)", "car insurance"];
  const itc = topIssueFromIneligibleItcRow(itcRow, 5000);
  assert(itc.amount === 9000 && itc.tier === 1, "Ineligible ITC amount >= threshold -> tier 1");
  assert(itc.reason.includes("Motor Vehicles") && itc.reason.includes("Section 17(5)"), "reason includes category and section");

  console.log("\n-- Duplicate Invoices (per-cluster, not per-row) --\n");

  // headers: Source Sheet(0), Cluster #(1), GSTIN(2), Row(3), Identifier(4), Taxable Value(5), IGST(6), CGST(7), SGST(8), Match Reason(9), Date Spread(10)
  const dupDataRows = [
    { row: ["Books", 1, "27ABCDE1234F1Z5", 10, "INV-9", 10000, 900, 0, 0, "same_invoice_number", 3], absoluteIndex: 20 },
    { row: ["Books", 1, "27ABCDE1234F1Z5", 11, "INV-9", 10000, 900, 0, 0, "same_invoice_number", 3], absoluteIndex: 21 },
    { row: ["Books", 2, "07XYZAB5678C1Z1", 15, "V-100", 5000, 450, 0, 0, "same_amount", 1], absoluteIndex: 22 },
    { row: ["Books", 2, "07XYZAB5678C1Z1", 16, "V-101", 5000, 450, 0, 0, "same_amount", 1], absoluteIndex: 23 },
    { row: ["2A", 1, "27ABCDE1234F1Z5", 30, "INV-A", 2000, 360, 0, 0, "same_amount", 2], absoluteIndex: 24 },
  ];
  const clusters = topIssuesFromDuplicateDataRows(dupDataRows);
  assert(clusters.length === 3, "3 distinct clusters found (Books cluster 1, Books cluster 2, 2A cluster 1) — grouped by (Source Sheet, Cluster #), not flattened into 5 individual rows");
  const booksCluster1 = clusters.find((c) => c.rowIndex === 20);
  assert(booksCluster1.rowSpan === 2, "cluster 1 spans its 2 contiguous member rows");
  assert(Math.abs(booksCluster1.amount - 900) < 0.01, "over-claim for 2 identical ₹900 members = 900 (sum 1800 minus the max 900)");
  // A real 'GST - Duplicate Invoices' sheet never actually contains a
  // single-member cluster (findDuplicateInvoices only ever writes
  // clusters of 2+), but this checks the grouping logic doesn't misbehave
  // if it ever saw one, rather than assuming that guarantee holds forever.
  const twoACluster = clusters.find((c) => c.rowIndex === 24);
  assert(twoACluster && twoACluster.rowSpan === 1 && twoACluster.amount === 0, "a lone 1-member cluster is still represented (rowSpan 1), with zero over-claim (sum equals the max when there's only one member)");
  assert(clusters.some((c) => c.reason.includes("2A")) && clusters.some((c) => c.reason.includes("Books")), "reasons correctly distinguish which source sheet each cluster came from");

  console.log("\n-- isDataRow: distinguishing real data from marker/blank/explanation rows --\n");

  assert(isDataRow(["27ABCDE1234F1Z5", 1000, 2000]) === true, "a row with real content beyond column A is a data row");
  assert(isDataRow(["Some marker or explanation text only"]) === false, "a row with content ONLY in column A (marker/explanation/banner) is NOT a data row");
  assert(isDataRow(["", "", ""]) === false, "an all-blank row is not a data row");
  assert(isDataRow(["Label", "", "", 0]) === true, "a row with an explicit 0 beyond column A IS still a data row (0 is real content, not blank)");

  console.log("\n-- findGstSheetHeaderRowIndex: robust to varying explanation-block length --\n");

  const rcmSheetValues = [
    ["Explanation line 1"],
    ["Explanation line 2"],
    ["Explanation line 3"],
    [""],
    ["Source Sheet", "Row", "GSTIN", "Invoice/Voucher Number", "Taxable Value", "IGST", "CGST", "SGST", "Category", "Notification", "Basis", "Needs Manual Confirmation?", "Reviewer Status", "Reviewer Note"],
    ["2A", 5, "27ABCDE1234F1Z5", "INV-1", 100000, 18000, 0, 0, "GTA Freight", "Notification 13/2017", "reverse_charge_flag", "No", "", ""],
  ];
  assert(findGstSheetHeaderRowIndex("GST - RCM", rcmSheetValues) === 4, "finds the header row regardless of a 3-line explanation block above it");

  const extraInBooksSheetValues = [
    ["Explanation line 1"],
    ["Explanation line 2"],
    [""],
    ["GSTIN", "Invoice/Voucher Number", "Row (in Books)", "Taxable Value", "IGST", "CGST", "SGST", "Reviewer Status", "Reviewer Note"],
    ["27ABCDE1234F1Z5", "INV-9", 12, 50000, 6000, 4500, 4500, "", ""],
  ];
  assert(
    findGstSheetHeaderRowIndex("GST - Extra in Books", extraInBooksSheetValues) === 3,
    "finds the invoice-level Extra in Books header (still matched on 'GSTIN' at column 0, unaffected by the header layout change)"
  );

  console.log("\n-- findWrongHeadSectionHeader: two stacked sections with different layouts --\n");

  const wrongHeadSheetValues = [
    ["Single-sheet explanation..."],
    [""],
    ["Row (in 2A)", "GSTIN", "Invoice Number", "Issue", "Expected Head", "Actual Head", "Supplier State", "Place of Supply State", "IGST", "CGST", "SGST", "Reviewer Status", "Reviewer Note"],
    [7, "27ABCDE1234F1Z5", "INV-200", "wrong_head", "IGST", "CGST+SGST", "27", "36", 0, 4500, 4500, "", ""],
    [""],
    ["Cross-sheet check (2A vs Books, same invoice)"],
    ["Cross-sheet explanation..."],
    ["GSTIN", "Invoice/Voucher Number", "2A Row", "Books Row", "2A Head", "Books Head", "Expected Head", "Incorrect Side", "2A IGST", "2A CGST", "2A SGST", "Books IGST", "Books CGST", "Books SGST", "Reviewer Status", "Reviewer Note"],
    ["37AABCR0435L1ZD", "211101051790", 1443, 1600, "IGST", "CGST+SGST", "IGST", "Books", 345663, 0, 0, 0, 172831.5, 172831.5, "", ""],
  ];
  const singleSection = findWrongHeadSectionHeader(wrongHeadSheetValues, 3); // row 3 = the single-sheet data row
  assert(singleSection.sectionType === "single" && singleSection.headerRowIndex === 2, "row 3 (single-sheet data) resolves to the single-sheet header at row 2");
  const crossSection = findWrongHeadSectionHeader(wrongHeadSheetValues, 8); // row 8 = the cross-sheet data row
  assert(crossSection.sectionType === "cross" && crossSection.headerRowIndex === 7, "row 8 (cross-sheet data) resolves to the CROSS-sheet header at row 7, not the single-sheet one above it");

  console.log("");
}

function run() {
  runTests();
  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All Top-Issues read-back logic checks passed.");
  }
}

run();
