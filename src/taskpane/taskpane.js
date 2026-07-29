/* global Excel, Office */

const { mentorSheetMemoryInit } = require("./mentor-sheet-memory-ui.js");

/* ============================================================
   MENTOR — Background Workbook Index
   Watches the Raw Data sheets continuously in the background,
   so suggestions on the report sheet are instant, not a live scan.
   ============================================================ */

   const mentorWorkbookIndex = {
    invoices: [],
    vendorMap: {},
    bankPayments: [],
    posSales: [],
    payroll: [],
    lastBuilt: null
  };
  
  async function mentorInit() {
    await Excel.run(async (context) => {
      const rawSheetNames = [
        "Raw - Supplier Invoices",
        "Raw - Bank Statement",
        "Raw - POS Sales Report",
        "Raw - Payroll Summary"
      ];
      for (const name of rawSheetNames) {
        const sheet = context.workbook.worksheets.getItemOrNullObject(name);
        await context.sync();
        if (!sheet.isNullObject) {
          sheet.onChanged.add(mentorDebounce(() => mentorRebuildIndex(name), 700));
        }
      }
    });
    // Build once immediately at startup, don't wait for the first edit
    await mentorRebuildIndex("Raw - Supplier Invoices");
    await mentorRebuildIndex("Raw - Bank Statement");
    await mentorRebuildIndex("Raw - POS Sales Report");
    await mentorRebuildIndex("Raw - Payroll Summary");
  }
  
  async function mentorRebuildIndex(sheetName) {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
      await context.sync();
      if (sheet.isNullObject) return;
  
      const range = sheet.getUsedRange();
      range.load("values");
      await context.sync();
  
      if (sheetName === "Raw - Supplier Invoices") {
        // Columns: Invoice ID, Site, Vendor, Invoice Date, Category, Net, VAT Rate...
        const rows = range.values.slice(3).filter(r => r[0]).map(r => ({
          vendor: String(r[2]).trim(),
          category: String(r[4]).trim(),
          vatRate: Number(r[6]),
          notes: String(r[10]).trim()
        }));
        mentorWorkbookIndex.invoices = rows;
        mentorWorkbookIndex.vendorMap = mentorBuildVendorMap(rows);
      }
  
      if (sheetName === "Raw - Bank Statement") {
        // Columns: Date (parsed), Date (as entered), Site, Description, Amount, Category, Matched Invoice ID, Notes
        const rows = range.values.slice(3).filter(r => r[3]).map(r => ({
          date: r[0],
          dateAsEntered: String(r[1]).trim(),
          description: String(r[3]).trim(),
          amount: Number(r[4]),
          category: String(r[5]).trim(),
          matchedInvoiceId: String(r[6]).trim()
        }));
        mentorWorkbookIndex.bankPayments = rows;
      }
  
      if (sheetName === "Raw - POS Sales Report") {
        // Columns: Month, Site, Dine-In, Takeaway, Deliveroo Gross, UberEats Gross, JustEat Gross, Events, Total
        const rows = range.values.slice(3).filter(r => r[0]).map(r => ({
          month: String(r[0]).trim(),
          deliveroo: Number(r[4]),
          uberEats: Number(r[5]),
          justEat: Number(r[6])
        }));
        mentorWorkbookIndex.posSales = rows;
      }
  
      if (sheetName === "Raw - Payroll Summary") {
        // Columns: Month, Site, Basic Wages, Overtime, Tronc/Tips, Employer NI, Pension, Total
        const rows = range.values.slice(3).filter(r => r[0]).map(r => ({
          month: String(r[0]).trim(),
          tronc: Number(r[4])
        }));
        mentorWorkbookIndex.payroll = rows;
      }
      mentorWorkbookIndex.lastBuilt = Date.now();
    });
  }
  
  // Reuses YOUR existing similarityPercent() further down this file — no new matching logic needed.
  function mentorBuildVendorMap(invoiceRows) {
    const uniqueNames = [...new Set(invoiceRows.map(r => r.vendor))];
    const map = {};
    for (const name of uniqueNames) {
      const existingCanonical = Object.keys(map).find(canon => {
        if (canon === name) return false;
        const sim = similarityPercent(canon, name); // defined later in this file
        return sim >= 55 && sim < 100; // tune threshold once you see real vendor names
      });
      if (existingCanonical) {
        map[existingCanonical].aliases.push(name);
      } else if (!map[name]) {
        map[name] = { canonical: name, aliases: [name] };
      }
    }
    return map;
  }
  
  function mentorFindRenameNote(categoryLabel) {
    if (!mentorWorkbookIndex.lastBuilt) return null; // index not warm yet
    const relevant = mentorWorkbookIndex.invoices.filter(
      r => r.category.toLowerCase() === categoryLabel.toLowerCase()
    );
    if (relevant.length === 0) return null;
  
    // FIX: only match a renamed vendor if it's actually one of THIS category's
    // vendors — before, a Food-category rename could wrongly surface while
    // checking Beverage, since nothing verified the vendor belonged here.
    const relevantVendors = new Set(relevant.map(r => r.vendor));
    const renamed = Object.values(mentorWorkbookIndex.vendorMap)
      .filter(v => v.aliases.length > 1 && v.aliases.some(a => relevantVendors.has(a)));
    if (renamed.length === 0) return null;
  
    const names = renamed[0].aliases.map(a => "'" + a + "'").join(" and ");
    return names + " look like the same vendor, renamed partway through the period — treat as one when totalling?";
  }
  
  // Hard Case 3 — Beverage (alcohol) invoices should be standard-rated (20%),
  // unlike most food which is genuinely zero-rated. A Beverage invoice sitting
  // at 0% VAT is a likely compliance error worth a second look.
  function mentorFindVatIssueNote(categoryLabel) {
    if (!mentorWorkbookIndex.lastBuilt) return null;
    if (categoryLabel.toLowerCase() !== "beverage") return null; // only Beverage carries this expectation
  
    const suspect = mentorWorkbookIndex.invoices.filter(
      r => r.category.toLowerCase() === "beverage" && r.vatRate === 0
    );
    if (suspect.length === 0) return null;
  
    const extra = suspect.length > 1 ? " (and " + (suspect.length - 1) + " more like it)" : "";
    return "A Beverage invoice from '" + suspect[0].vendor + "'" + extra + " is showing 0% VAT — alcohol is normally standard-rated at 20%. Worth checking before this goes into your VAT return?";
  }
  // Hard Case 4 — a bank payment went out with nothing on file to justify it.
  // Cross-references the Bank Statement against the Supplier Invoices index.
  function mentorFindMissingInvoiceNote() {
    if (!mentorWorkbookIndex.lastBuilt || !mentorWorkbookIndex.bankPayments) return null;
    const missing = mentorWorkbookIndex.bankPayments.filter(
      p => p.category === "Supplier Payment" && p.matchedInvoiceId === ""
    );
    if (missing.length === 0) return null;
  
    const first = missing[0];
    const vendorName = first.description.replace(/^PAYMENT - /i, "");
    const extra = missing.length > 1 ? " (and " + (missing.length - 1) + " more like it)" : "";
    return "A payment to '" + vendorName + "'" + extra + " has no matching invoice on file — worth chasing the missing paperwork before this goes into your total?";
  }
  
  // Hard Case 5 — a genuine judgment call, not a clean error. A documented
  // price rise is real, but it's not the ONLY thing that could explain a food
  // cost increase. This deliberately surfaces both possibilities rather than
  // picking one — that restraint is the point, not a gap to fill in later.
  function mentorFindPriceRiseAmbiguityNote(categoryLabel) {
    if (!mentorWorkbookIndex.lastBuilt) return null;
    if (categoryLabel.toLowerCase() !== "food") return null;
  
    const priceRiseInvoices = mentorWorkbookIndex.invoices.filter(
      r => r.category.toLowerCase() === "food" && /price/i.test(r.notes)
    );
    if (priceRiseInvoices.length === 0) return null;
  
    return "'" + priceRiseInvoices[0].vendor + "' has a documented price increase this period — but a food cost rise could also reflect wastage. Both are plausible here; worth confirming with site managers rather than assuming either one.";
  }
  
  // Hard Case 6 — a commission rate change hidden inside net payouts. Instead
  // of searching for a hint someone left in a notes field, this actually
  // calculates the real commission rate each month (gross POS sales vs. net
  // bank payout) and flags when that rate moves — this works on any dataset
  // with sales and payout figures, not just this demo's planted wording.
  function mentorFindCommissionChangeNote() {
    if (!mentorWorkbookIndex.lastBuilt) return null;
    const pos = mentorWorkbookIndex.posSales || [];
    const bank = mentorWorkbookIndex.bankPayments || [];
    if (pos.length === 0 || bank.length === 0) return null;
  
    const monthOrder = [...new Set(pos.map(r => r.month))]; // preserves first-seen order
  
    for (const platform of ["deliveroo", "uberEats", "justEat"]) {
      const rateByMonth = {};
      for (const monthLabel of monthOrder) {
        const gross = pos.filter(r => r.month === monthLabel).reduce((sum, r) => sum + (r[platform] || 0), 0);
        const platformName = platform === "deliveroo" ? "Deliveroo" : platform === "uberEats" ? "UberEats" : "JustEat";
        const net = bank
          .filter(p => p.category === "Delivery Platform Payout" && p.description.startsWith(platformName))
          .filter(p => mentorSameMonth(p.date, monthLabel))
          .reduce((sum, p) => sum + p.amount, 0);
        if (gross > 0) rateByMonth[monthLabel] = 1 - (net / gross);
      }
  
      for (let i = 1; i < monthOrder.length; i++) {
        const prevRate = rateByMonth[monthOrder[i - 1]];
        const currRate = rateByMonth[monthOrder[i]];
        if (prevRate === undefined || currRate === undefined) continue;
        const diff = currRate - prevRate;
        if (Math.abs(diff) > 0.01) { // more than a 1 percentage point shift
          const platformName = platform === "deliveroo" ? "Deliveroo" : platform === "uberEats" ? "UberEats" : "JustEat";
          const direction = diff > 0 ? "risen" : "fallen";
          return platformName + "'s commission rate looks like it has " + direction + " from about " +
            Math.round(prevRate * 100) + "% to " + Math.round(currRate * 100) + "% between " +
            monthOrder[i - 1] + " and " + monthOrder[i] + " — worth confirming with the platform before assuming delivery revenue itself has changed.";
        }
      }
    }
    return null;
  }
  
  // Matches a bank transaction date to a POS month label like "Jun 2026".
  // Excel gives back date cells as plain serial numbers (e.g. 46200), not
  // real dates — same quirk your reconciliation code already handles
  // elsewhere in this file. Converting properly here fixes it for good.
  function mentorSameMonth(bankDate, monthLabel) {
    if (bankDate === null || bankDate === undefined || bankDate === "") return false;
    let d;
    if (typeof bankDate === "number") {
      const excelEpoch = new Date(1899, 11, 30);
      d = new Date(excelEpoch.getTime() + bankDate * 86400000);
    } else {
      d = new Date(bankDate);
    }
    if (isNaN(d.getTime())) return false;
    const monthNames = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
    const label = monthNames[d.getMonth()] + " " + d.getFullYear();
    return label === monthLabel;
  }
  
  // Hard Case 8 — not an error to catch, a design decision worth explaining
  // proactively. Tronc/tips are fully redistributed to staff, so excluding
  // them from Payroll Cost is correct — but a user seeing a lower number than
  // expected shouldn't have to wonder whether data went missing.
  function mentorFindTroncNote() {
    if (!mentorWorkbookIndex.lastBuilt) return null;
    const payroll = mentorWorkbookIndex.payroll || [];
    const hasTronc = payroll.some(r => r.tronc > 0);
    if (!hasTronc) return null;
    return "Payroll includes Tronc/tips redistributed to staff — deliberately excluded from this total, since it passes straight through and isn't a real cost to the business. If you'd rather include it, just say so.";
  }
  
  // Hard Case 7 — a genuinely messy data-entry artefact: one payout date was
  // typed in a different format (dots instead of slashes) than the rest.
  // Worth flagging as a data-hygiene issue, not something to silently "fix."
  function mentorFindMessyDateNote() {
    if (!mentorWorkbookIndex.lastBuilt) return null;
    const bank = mentorWorkbookIndex.bankPayments || [];
    const messy = bank.find(p => p.category === "Delivery Platform Payout" && /\d{2}\.\d{2}\.\d{4}/.test(p.dateAsEntered));
    if (!messy) return null;
    return "One payout ('" + messy.description + "', entered as '" + messy.dateAsEntered + "') was typed in a different date format than the rest of the sheet — worth a quick fix so date-based formulas and sorts don't trip over it.";
  }
  
  // Not a judgment call, just missing arithmetic — Total Revenue, Total COGS,
  // Gross Profit, Total Operating Expenses, and EBITDA should auto-fill once
  // their component rows have real figures in them.
  function mentorFindSubtotalRowIdx(values, headerRowIdx, label) {
    // Returns the LAST matching row, not the first — this sheet has section
    // headers (e.g. bold "Gross Profit") immediately above the real data row
    // of the same name, and the data row is always the later one.
    let found = -1;
    for (let i = headerRowIdx + 1; i < values.length; i++) {
      if (String(values[i][0]).toLowerCase().trim() === label) found = i;
    }
    return found;
  }
  
  function mentorRowHasAnyValue(values, rowIdx) {
    if (rowIdx === -1) return false;
    return values[rowIdx].slice(1).some(c => typeof c === "number" && c !== 0);
  }
  
  function mentorRowIsBlank(values, rowIdx) {
    if (rowIdx === -1) return true;
    return values[rowIdx].slice(1).every(c => c === "" || c === 0 || c === null || c === undefined);
  }
  
  function mentorFindSubtotalsNeeded(values, headerRowIdx) {
    const find = (label) => mentorFindSubtotalRowIdx(values, headerRowIdx, label);
    const hasAny = (labels) => labels.some(l => mentorRowHasAnyValue(values, find(l)));
    const isBlank = (label) => mentorRowIsBlank(values, find(label));
  
    if (hasAny(["dine-in", "takeaway", "delivery revenue (net of commission)", "events / catering"]) && isBlank("total revenue")) return true;
    if (hasAny(["food cost", "beverage cost"]) && isBlank("total cogs")) return true;
    if (!isBlank("total revenue") && !isBlank("total cogs") && isBlank("gross profit")) return true;
    if (hasAny(["payroll cost (excl. tronc)", "rent", "utilities", "other opex (cleaning, maintenance, marketing)"]) && isBlank("total operating expenses")) return true;
    if (!isBlank("gross profit") && !isBlank("total operating expenses") && isBlank("ebitda")) return true;
  
    return false;
  }
  
  // Row names this suggestion is actually relevant to — either a subtotal
  // row itself, or one of its component rows. FIX: this used to fire
  // unconditionally regardless of which row you were on, so it could show
  // while sitting on an unrelated cell like Delivery Revenue, or stay
  // silent while genuinely sitting on Total Revenue.
  const mentorSubtotalRelevantRows = new Set([
    "total revenue", "dine-in", "takeaway", "delivery revenue (net of commission)", "events / catering",
    "total cogs", "food cost", "beverage cost",
    "gross profit",
    "total operating expenses", "payroll cost (excl. tronc)", "rent", "utilities", "other opex (cleaning, maintenance, marketing)",
    "ebitda"
  ]);
  
  function mentorCheckSubtotalsForRow(values, headerRowIdx, address) {
    const match = (address || "").match(/([A-Z]+)(\d+)$/i);
    if (!match) return null;
    const rowIdx = parseInt(match[2], 10) - 1;
    if (rowIdx < 0 || rowIdx >= values.length) return null;
  
    const rowLabel = String(values[rowIdx][0]).toLowerCase().trim();
    if (!mentorSubtotalRelevantRows.has(rowLabel)) return null;
    if (!mentorFindSubtotalsNeeded(values, headerRowIdx)) return null;
    if (mentorDismissedThisSession.has("PNL_SUBTOTALS")) return null;
  
    return {
      id: "PNL_SUBTOTALS",
      title: "Fill in the subtotals?",
      message: "Some figures are in, but Total Revenue, Total COGS, Gross Profit, Total Operating Expenses, or EBITDA aren't calculating yet. Want me to add the formulas?",
      action: "insertPnlSubtotals"
    };
  }
  
  const mentorCategoryMap = { "food cost": "Food", "beverage cost": "Beverage" };
  
  // Shared by BOTH edits (onChanged) and clicks (onSelectionChanged) — checks
  // whichever row `address` points to for any of the hard cases. Returns a
  // stage object if one applies (and isn't already dismissed), or null.
  function mentorCheckHardCaseForRow(values, address) {
    const match = (address || "").match(/([A-Z]+)(\d+)$/i);
    if (!match) return null;
    const rowNum = parseInt(match[2], 10);
    const rowIdx = rowNum - 1;
    if (rowIdx < 0 || rowIdx >= values.length) return null;
  
    const rowLabel = String(values[rowIdx][0]).toLowerCase().trim();
  
    if (mentorCategoryMap[rowLabel]) {
      const note = mentorFindRenameNote(mentorCategoryMap[rowLabel]);
      const stageId = "VENDOR_RENAME_" + rowLabel;
      if (note && !mentorDismissedThisSession.has(stageId)) {
        return { id: stageId, title: "Same vendor, different name?", message: note, action: "documentVendorMerge", targetAddress: address };
      }
  
      const priceNote = mentorFindPriceRiseAmbiguityNote(mentorCategoryMap[rowLabel]);
      const priceStageId = "PRICE_AMBIGUITY_" + rowLabel;
      if (priceNote && !mentorDismissedThisSession.has(priceStageId)) {
        return { id: priceStageId, title: "Worth a closer look?", message: priceNote, action: "documentPriceAmbiguity", targetAddress: address };
      }
  
      const vatNote = mentorFindVatIssueNote(mentorCategoryMap[rowLabel]);
      const vatStageId = "VAT_ISSUE_" + rowLabel;
      if (vatNote && !mentorDismissedThisSession.has(vatStageId)) {
        return { id: vatStageId, title: "Possible VAT error?", message: vatNote, action: "documentVatIssue", targetAddress: address };
      }
    }
  
    if (rowLabel === "other opex (cleaning, maintenance, marketing)") {
      const missingNote = mentorFindMissingInvoiceNote();
      const missingStageId = "MISSING_INVOICE_" + rowLabel;
      if (missingNote && !mentorDismissedThisSession.has(missingStageId)) {
        return { id: missingStageId, title: "Missing invoice on file?", message: missingNote, action: "documentMissingInvoice", targetAddress: address };
      }
    }
  
    if (rowLabel === "delivery revenue (net of commission)") {
      const commissionNote = mentorFindCommissionChangeNote();
      const commissionStageId = "COMMISSION_CHANGE_" + rowLabel;
      if (commissionNote && !mentorDismissedThisSession.has(commissionStageId)) {
        return { id: commissionStageId, title: "Commission rate changed?", message: commissionNote, action: "documentCommissionChange", targetAddress: address };
      }
  
      const dateNote = mentorFindMessyDateNote();
      const dateStageId = "MESSY_DATE_" + rowLabel;
      if (dateNote && !mentorDismissedThisSession.has(dateStageId)) {
        return { id: dateStageId, title: "Inconsistent date format?", message: dateNote, action: "documentMessyDate", targetAddress: address };
      }
    }
  
    if (rowLabel === "payroll cost (excl. tronc)") {
      const troncNote = mentorFindTroncNote();
      const troncStageId = "TRONC_NOTE_" + rowLabel;
      if (troncNote && !mentorDismissedThisSession.has(troncStageId)) {
        return { id: troncStageId, title: "Why isn't tronc included?", message: troncNote, action: "documentTroncExclusion", targetAddress: address };
      }
    }
  
    return null;
  }
  
  function mentorDebounce(fn, ms) {
    let t;
    return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); };
  }
  
  // SHARED header-row finder — used by BOTH detectReportStage and
  // mentorAcceptStage's insertVariance action. Previously this logic was
  // copy-pasted in two places, and a fix applied to one copy didn't reach
  // the other — this single shared version prevents that happening again.
  function mentorFindHeaderRowIdx(values) {
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const rowLower = row.map(v => String(v).toLowerCase());
      const filledCells = rowLower.filter(v => v.trim() !== "").length;
      const looksLikeHeader = filledCells >= 3 && rowLower.some(h =>
        h.includes("category") || h.includes("amount") || h.includes("month") ||
        h.includes("revenue") || h.includes("cost") || h.includes("line item") ||
        /\d{4}/.test(h) // catches "May 2026", "Jun 2026" style headers too
      );
      if (looksLikeHeader) return i;
    }
    return -1;
  }
  
  window.onload = function() {
    document.getElementById("reconcile-btn").onclick = runReconciliation;
  
    if (typeof Office === "undefined") {
      document.getElementById("results").innerHTML = "<p style='color:red;'>Office JS not loaded</p>";
      return;
    }
  
    Office.onReady(function(info) {
      loadSheetNames();
      startMentorObserver();
      mentorInit();
      mentorSheetMemoryInit(mentorLogAction);
    });
  };
  
  async function loadSheetNames() {
    try {
      await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();
  
        const ledgerSelect = document.getElementById("ledger-select");
        const supplierSelect = document.getElementById("supplier-select");
  
        ledgerSelect.innerHTML = "";
        supplierSelect.innerHTML = "";
  
        sheets.items.forEach(function(sheet) {
          const opt1 = document.createElement("option");
          opt1.value = sheet.name;
          opt1.text = sheet.name;
          ledgerSelect.appendChild(opt1);
  
          const opt2 = document.createElement("option");
          opt2.value = sheet.name;
          opt2.text = sheet.name;
          supplierSelect.appendChild(opt2);
        });
  
        if (sheets.items.length >= 2) {
          supplierSelect.selectedIndex = 1;
        }
      });
    } catch (error) {
      document.getElementById("results").innerHTML = "<p style='color:red;'>Error loading sheets: " + error.message + "</p>";
    }
  }
  
  function levenshtein(a, b) {
    const matrix = [];
    for (let i = 0; i <= b.length; i++) matrix[i] = [i];
    for (let j = 0; j <= a.length; j++) matrix[0][j] = j;
    for (let i = 1; i <= b.length; i++) {
      for (let j = 1; j <= a.length; j++) {
        if (b.charAt(i - 1) === a.charAt(j - 1)) {
          matrix[i][j] = matrix[i - 1][j - 1];
        } else {
          matrix[i][j] = Math.min(
            matrix[i - 1][j - 1] + 1,
            matrix[i][j - 1] + 1,
            matrix[i - 1][j] + 1
          );
        }
      }
    }
    return matrix[b.length][a.length];
  }
  
  function similarityPercent(a, b) {
    const aLower = a.toLowerCase().trim();
    const bLower = b.toLowerCase().trim();
    if (aLower === bLower) return 100;
    const maxLen = Math.max(aLower.length, bLower.length);
    if (maxLen === 0) return 100;
    const dist = levenshtein(aLower, bLower);
    return Math.round((1 - dist / maxLen) * 100);
  }
  
  function findHeaderRow(values) {
    for (let i = 0; i < values.length; i++) {
      const cellValue = String(values[i][0]).toLowerCase().trim();
      if (cellValue.includes("invoice")) return i;
    }
    return 0;
  }
  
  async function runReconciliation() {
    const resultsDiv = document.getElementById("results");
    resultsDiv.innerHTML = "<p style='color:#666;'>Running reconciliation...</p>";
  
    const ledgerName = document.getElementById("ledger-select").value;
    const supplierName = document.getElementById("supplier-select").value;
  
    if (!ledgerName || !supplierName) {
      resultsDiv.innerHTML = "<p style='color:red;'>Please select both sheets first.</p>";
      return;
    }
  
    if (ledgerName === supplierName) {
      resultsDiv.innerHTML = "<p style='color:red;'>Please select two different sheets.</p>";
      return;
    }
  
    try {
      await Excel.run(async (context) => {
  
        const ledgerSheet = context.workbook.worksheets.getItem(ledgerName);
        const supplierSheet = context.workbook.worksheets.getItem(supplierName);
  
        const ledgerRange = ledgerSheet.getUsedRange();
        const supplierRange = supplierSheet.getUsedRange();
  
        ledgerRange.load("values, rowIndex");
        supplierRange.load("values, rowIndex");
  
        await context.sync();
  
        const ledgerValues = ledgerRange.values;
        const supplierValues = supplierRange.values;
  
        const ledgerStartRow = ledgerRange.rowIndex + 1;
        const supplierStartRow = supplierRange.rowIndex + 1;
  
        const ledgerHeaderIdx = findHeaderRow(ledgerValues);
        const supplierHeaderIdx = findHeaderRow(supplierValues);
  
        const ledgerRows = parseRows(ledgerValues, ledgerHeaderIdx);
        const supplierRows = parseRows(supplierValues, supplierHeaderIdx);
  
        const results = reconcile(ledgerRows, supplierRows);
  
        const excelColors = {
          MATCHED:          "#00B050",
          AMOUNT_DIFF:      "#FFC000",
          DATE_DIFF:        "#00B0F0",
          DUPLICATE:        "#FF0000",
          MISSING_SUPPLIER: "#FF6600",
          MISSING_LEDGER:   "#7030A0",
          VAT_ISSUE:        "#FFFF00",
          FUZZY_VENDOR:     "#00CED1"
        };
  
        ledgerSheet.getUsedRange().format.fill.clear();
        supplierSheet.getUsedRange().format.fill.clear();
  
        const ledgerRowMap = {};
        for (let i = ledgerHeaderIdx + 1; i < ledgerValues.length; i++) {
          const invNo = String(ledgerValues[i][0]).trim();
          if (invNo) {
            const excelRowNum = ledgerStartRow + i;
            if (!ledgerRowMap[invNo]) ledgerRowMap[invNo] = excelRowNum;
            if (!ledgerRowMap["all_" + invNo]) ledgerRowMap["all_" + invNo] = [];
            ledgerRowMap["all_" + invNo].push(excelRowNum);
          }
        }
  
        const supplierRowMap = {};
        for (let i = supplierHeaderIdx + 1; i < supplierValues.length; i++) {
          const invNo = String(supplierValues[i][0]).trim();
          if (invNo && !supplierRowMap[invNo]) {
            supplierRowMap[invNo] = supplierStartRow + i;
          }
        }
  
        results.forEach(function(r) {
          const color = excelColors[r.status];
  
          if (r.status === "DUPLICATE") {
            const allRows = ledgerRowMap["all_" + r.invoiceNo];
            if (allRows) {
              allRows.forEach(function(rowNum) {
                ledgerSheet.getRange("A" + rowNum + ":E" + rowNum).format.fill.color = color;
              });
            }
          } else if (r.status === "MISSING_LEDGER") {
            if (supplierRowMap[r.invoiceNo]) {
              const rowNum = supplierRowMap[r.invoiceNo];
              supplierSheet.getRange("A" + rowNum + ":D" + rowNum).format.fill.color = color;
            }
          } else if (r.status === "MISSING_SUPPLIER") {
            if (ledgerRowMap[r.invoiceNo]) {
              const rowNum = ledgerRowMap[r.invoiceNo];
              ledgerSheet.getRange("A" + rowNum + ":E" + rowNum).format.fill.color = color;
            }
          } else {
            if (ledgerRowMap[r.invoiceNo]) {
              const rowNum = ledgerRowMap[r.invoiceNo];
              ledgerSheet.getRange("A" + rowNum + ":E" + rowNum).format.fill.color = color;
            }
            if (supplierRowMap[r.invoiceNo]) {
              const rowNum = supplierRowMap[r.invoiceNo];
              supplierSheet.getRange("A" + rowNum + ":D" + rowNum).format.fill.color = color;
            }
          }
        });
  
        await context.sync();
  
        displayResults(results, resultsDiv);
  
      });
  
    } catch (error) {
      resultsDiv.innerHTML = "<p style='color:red;'>Error: " + error.message + "</p>";
      console.error("MENTOR error:", error);
    }
  }
  
  function parseRows(data, headerIdx) {
    const rows = [];
    for (let i = headerIdx + 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
  
      let dateVal = row[3];
      if (typeof dateVal === "number") {
        const excelEpoch = new Date(1899, 11, 30);
        const date = new Date(excelEpoch.getTime() + dateVal * 86400000);
        dateVal = date.toLocaleDateString("en-GB");
      } else {
        dateVal = String(dateVal).trim();
      }
  
      rows.push({
        invoiceNo: String(row[0]).trim(),
        supplier: String(row[1]).trim(),
        amount: parseFloat(row[2]),
        date: dateVal
      });
    }
    return rows;
  }
  
  function reconcile(ledger, supplier) {
    const results = [];
    const supplierMatched = new Set();
  
    const ledgerSignatures = {};
    ledger.forEach((row, idx) => {
      const sig = row.supplier.toLowerCase() + "|" + row.amount + "|" + row.date;
      if (!ledgerSignatures[sig]) ledgerSignatures[sig] = [];
      ledgerSignatures[sig].push(idx);
    });
  
    const ledgerInvoiceCount = {};
    ledger.forEach((row, idx) => {
      if (!ledgerInvoiceCount[row.invoiceNo]) ledgerInvoiceCount[row.invoiceNo] = [];
      ledgerInvoiceCount[row.invoiceNo].push(idx);
    });
  
    ledger.forEach((lRow, lIdx) => {
  
      if (ledgerInvoiceCount[lRow.invoiceNo].length > 1 &&
          ledgerInvoiceCount[lRow.invoiceNo][0] !== lIdx) {
        results.push({
          status: "DUPLICATE",
          invoiceNo: lRow.invoiceNo,
          message: lRow.invoiceNo + " appears more than once in your ledger. Check which entry is correct."
        });
        return;
      }
  
      const sig = lRow.supplier.toLowerCase() + "|" + lRow.amount + "|" + lRow.date;
      if (ledgerSignatures[sig].length > 1 && ledgerSignatures[sig][0] !== lIdx) {
        results.push({
          status: "DUPLICATE",
          invoiceNo: lRow.invoiceNo,
          message: lRow.invoiceNo + " — possible duplicate. Same supplier, amount and date as " +
            ledger[ledgerSignatures[sig][0]].invoiceNo + ". Check if this invoice was entered twice."
        });
        return;
      }
  
      const sIdx = supplier.findIndex(s => s.invoiceNo === lRow.invoiceNo);
  
      if (sIdx === -1) {
        const fuzzyMatch = supplier.find(s => {
          const sim = similarityPercent(lRow.supplier, s.supplier);
          return sim >= 80 && sim < 100;
        });
  
        if (fuzzyMatch) {
          const sim = similarityPercent(lRow.supplier, fuzzyMatch.supplier);
          results.push({
            status: "FUZZY_VENDOR",
            invoiceNo: lRow.invoiceNo,
            message: lRow.invoiceNo + " — Vendor name differs slightly. Your ledger says \"" + lRow.supplier + "\". Supplier says \"" + fuzzyMatch.supplier + "\". " + sim + "% match — possible same supplier, please verify."
          });
        } else {
          results.push({
            status: "MISSING_SUPPLIER",
            invoiceNo: lRow.invoiceNo,
            message: lRow.invoiceNo + " (" + lRow.supplier + " GBP " + lRow.amount + ") is in your ledger but not on the supplier statement. Possible overpayment or unconfirmed invoice."
          });
        }
        return;
      }
  
      const sRow = supplier[sIdx];
      supplierMatched.add(sIdx);
  
      const vendorSim = similarityPercent(lRow.supplier, sRow.supplier);
      if (vendorSim < 80) {
        results.push({
          status: "FUZZY_VENDOR",
          invoiceNo: lRow.invoiceNo,
          message: lRow.invoiceNo + " — Invoice matches but vendor names differ. Your ledger says \"" + lRow.supplier + "\". Supplier says \"" + sRow.supplier + "\". " + vendorSim + "% match — please verify."
        });
        return;
      }
  
      const amountDiff = Math.abs(lRow.amount - sRow.amount);
      if (amountDiff > 0.001) {
        const vatCheck = Math.abs(amountDiff / lRow.amount - 0.20);
        if (vatCheck < 0.01) {
          results.push({
            status: "VAT_ISSUE",
            invoiceNo: lRow.invoiceNo,
            message: lRow.invoiceNo + " - Amount differs by exactly 20%. Your ledger: GBP " + lRow.amount + ". Supplier: GBP " + sRow.amount + ". Likely a VAT posting error."
          });
        } else {
          results.push({
            status: "AMOUNT_DIFF",
            invoiceNo: lRow.invoiceNo,
            message: lRow.invoiceNo + " (" + lRow.supplier + ") - Amount mismatch. Your ledger: GBP " + lRow.amount + ". Supplier says: GBP " + sRow.amount + ". Difference: GBP " + amountDiff.toFixed(2) + "."
          });
        }
        return;
      }
  
      if (lRow.date !== sRow.date) {
        results.push({
          status: "DATE_DIFF",
          invoiceNo: lRow.invoiceNo,
          message: lRow.invoiceNo + " (" + lRow.supplier + ") - Dates differ. Your ledger: " + lRow.date + ". Supplier: " + sRow.date + ". Likely a timing difference."
        });
        return;
      }
  
      results.push({
        status: "MATCHED",
        invoiceNo: lRow.invoiceNo,
        message: lRow.invoiceNo + " (" + lRow.supplier + ") - Clean match. Amount and date agree."
      });
    });
  
    supplier.forEach((sRow, sIdx) => {
      if (!supplierMatched.has(sIdx)) {
        results.push({
          status: "MISSING_LEDGER",
          invoiceNo: sRow.invoiceNo,
          message: sRow.invoiceNo + " (" + sRow.supplier + " GBP " + sRow.amount + ") is on the supplier statement but missing from your ledger. Unrecorded liability - action required."
        });
      }
    });
  
    return results;
  }
  
  function displayResults(results, container) {
  
    const cardBg = "#22252b";
  
    const accentColors = {
      MATCHED:          "#00B050",
      AMOUNT_DIFF:      "#FFC000",
      DATE_DIFF:        "#00B0F0",
      DUPLICATE:        "#FF0000",
      MISSING_SUPPLIER: "#FF6600",
      MISSING_LEDGER:   "#7030A0",
      VAT_ISSUE:        "#FFFF00",
      FUZZY_VENDOR:     "#00CED1"
    };
  
    const labels = {
      MATCHED:          "MATCHED",
      AMOUNT_DIFF:      "AMOUNT DIFFERENCE",
      DATE_DIFF:        "DATE DIFFERENCE",
      DUPLICATE:        "DUPLICATE",
      MISSING_SUPPLIER: "MISSING FROM SUPPLIER",
      MISSING_LEDGER:   "UNRECORDED LIABILITY",
      VAT_ISSUE:        "POSSIBLE VAT ISSUE",
      FUZZY_VENDOR:     "VENDOR NAME MISMATCH"
    };
  
    const counts = {};
    results.forEach(function(r) {
      counts[r.status] = (counts[r.status] || 0) + 1;
    });
  
    let summary = "<div style='background:#1a1a2e;padding:10px;border-radius:8px;margin-bottom:12px;'>";
    summary += "<p style='color:white;font-size:11px;font-weight:bold;margin:0 0 8px 0;'>SUMMARY — " + results.length + " items reviewed</p>";
  
    Object.keys(counts).forEach(function(status) {
      const dot = "<span style='display:inline-block;width:8px;height:8px;border-radius:50%;background:" + accentColors[status] + ";margin-right:5px;vertical-align:middle;'></span>";
      summary += "<div style='color:white;font-size:11px;margin:3px 0;'>" + dot + counts[status] + " " + labels[status] + "</div>";
    });
  
    summary += "</div>";
  
    const order = ["DUPLICATE", "VAT_ISSUE", "AMOUNT_DIFF", "DATE_DIFF", "FUZZY_VENDOR", "MISSING_SUPPLIER", "MISSING_LEDGER", "MATCHED"];
  
    const grouped = {};
    order.forEach(function(status) {
      grouped[status] = results.filter(function(r) { return r.status === status; });
    });
  
    let html = summary;
    html += "<h3 style='margin-bottom:6px;font-size:13px;color:#222;'>Results</h3>";
  
    order.forEach(function(status) {
      if (grouped[status].length === 0) return;
  
      grouped[status].forEach(function(r) {
        const globalIdx = results.indexOf(r);
        const dot = "<span style='display:inline-block;width:9px;height:9px;border-radius:50%;background:" + accentColors[status] + ";margin-right:6px;vertical-align:middle;'></span>";
        html += "<div id='item-" + globalIdx + "' style='background:" + cardBg + ";border-left:4px solid " + accentColors[status] + ";padding:10px 10px 10px 12px;margin:6px 0;border-radius:6px;font-size:12px;line-height:1.5;color:#fff;'>";
        html += "<strong style='color:#fff;'>" + dot + labels[status] + "</strong><br/>";
        html += "<span style='color:#cfd2d8;'>" + r.message + "</span><br/>";
        html += "<div style='margin-top:6px;'>";
        html += "<button onclick='acceptItem(" + globalIdx + ")' style='margin-right:6px;padding:3px 10px;background:#28a745;color:white;border:none;border-radius:4px;cursor:pointer;font-size:11px;'>Accept</button>";
        html += "<button onclick='rejectItem(" + globalIdx + ")' style='padding:3px 10px;background:#dc3545;color:white;border:none;border-radius:4px;cursor:pointer;font-size:11px;'>Reject</button>";
        html += "</div></div>";
      });
    });
  
    container.innerHTML = html;
  }
  
  window.acceptItem = function(idx) {
    const item = document.getElementById("item-" + idx);
    if (item) {
      const buttons = item.querySelectorAll("button");
      buttons.forEach(function(b) { b.style.display = "none"; });
      item.style.opacity = "0.6";
      item.innerHTML += "<em style='color:#a8ffb0;font-size:11px;'>Accepted</em>";
    }
  };
  
  window.rejectItem = function(idx) {
    const item = document.getElementById("item-" + idx);
    if (item) {
      const buttons = item.querySelectorAll("button");
      buttons.forEach(function(b) { b.style.display = "none"; });
      item.style.opacity = "0.6";
      item.innerHTML += "<em style='color:#ffb3b8;font-size:11px;'>Rejected</em>";
    }
  };
  
  /* ============================================================
     MENTOR — Monthly Report Build Workflow (Day 1 build)
     Passive observation + suggestion badge + Accept/Reject
     ============================================================ */
  
  let mentorListenerAttached = false;
  let mentorDebounceTimer = null;
  const mentorDismissedThisSession = new Set();
  
  async function startMentorObserver() {
    if (mentorListenerAttached) return;
    try {
      await Excel.run(async (context) => {
        // FIX: listen at the WORKBOOK level (all sheets), not just whichever
        // sheet happened to be active when the task pane first loaded.
        // A sheet-specific onChanged only fires for edits on that one sheet —
        // that's why nothing ever fired when editing a different tab.
        context.workbook.worksheets.onChanged.add(onSheetChangedDebounced);
        context.workbook.worksheets.onSelectionChanged.add(mentorDebounce(onSelectionChanged, 120));
        await context.sync();
        mentorListenerAttached = true;
      });
    } catch (error) {
      console.error("MENTOR observer error:", error);
    }
  }
  
  // The proactive "notices the label, notices the matching raw data, and
  // proposes the formula" piece from the original roadmap. Fires when you
  // click into a still-blank cell whose row MENTOR recognises — before you've
  // typed anything — offering to pull the real total in from the raw sheets.
  let mentorPendingPreview = null; // { address, formula, stageId } — the ghost-preview currently sitting in a cell
  
  async function onSelectionChanged(event) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load("name");
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load("values, isNullObject");
        await context.sync();
  
        if (usedRange.isNullObject) { hideMentorBadge(); return; }
        if (sheet.name !== "Monthly P&L (Demo - Start Here)") { hideMentorBadge(); return; } // only the working sheet
  
        const address = event && event.address;
        if (!address) return;
  
        // Finalize any ghost preview left over from the PREVIOUS cell first,
        // regardless of what's relevant to the new selection. Moving away
        // without touching it = accepted; typing over it = your edit already
        // stands, nothing more to do.
        if (mentorPendingPreview && mentorPendingPreview.address !== address) {
          const pendingRange = sheet.getRange(mentorPendingPreview.address);
          pendingRange.load("formulas");
          await context.sync();
          const currentFormula = pendingRange.formulas[0][0];
          if (currentFormula === mentorPendingPreview.formula) {
            pendingRange.format.font.color = "#000000"; // was gray, now confirmed real
            await context.sync();
            mentorDismissedThisSession.add(mentorPendingPreview.stageId);
            mentorLogAction("Pulled a formula into " + mentorPendingPreview.address + " from the raw data sheets, on '" + sheet.name + "'");
  
            const confirmedRowMatch = mentorPendingPreview.address.match(/(\d+)$/);
            if (confirmedRowMatch) {
              await mentorAutoFillVarianceForRow(context, sheet, parseInt(confirmedRowMatch[1], 10) - 1);
            }
          }
          mentorPendingPreview = null;
        }
  
        await mentorCheckRowForSuggestion(context, sheet, usedRange.values, address);
      });
    } catch (error) {
      console.error("MENTOR selection error:", error);
    }
  }
  
  // Everything that decides what (if anything) to show for a given cell —
  // hard case, subtotals, or a ghost-preview offer. Shared by BOTH a real
  // selection change AND the post-accept re-check, so resolving an issue
  // immediately reveals the next thing (including a ghost preview) without
  // needing you to click away and back first.
  // Pure arithmetic, not judgment — if the Variance column already exists
  // and this row now has a real number where it didn't before, fill in its
  // Variance formula automatically. FIX: "Add a variance column?" only ever
  // swept whatever data existed at the moment you clicked it — anything
  // filled in afterward (via a ghost preview or later typing) never got
  // caught up. This runs every time a value lands, so nothing gets left behind.
  async function mentorAutoFillVarianceForRow(context, sheet, rowIdx) {
    const usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load("values, isNullObject");
    await context.sync();
    if (usedRange.isNullObject) return;
  
    const values = usedRange.values;
    const headerRowIdx = mentorFindHeaderRowIdx(values);
    if (headerRowIdx === -1 || rowIdx <= headerRowIdx || rowIdx >= values.length) return;
  
    const varianceColIdx = values[headerRowIdx].findIndex(v => String(v).toLowerCase().includes("variance"));
    if (varianceColIdx === -1) return; // no Variance column yet — nothing to catch up
  
    const bVal = values[rowIdx][1], cVal = values[rowIdx][2];
    const hasRealData = (typeof bVal === "number" && bVal !== 0) || (typeof cVal === "number" && cVal !== 0);
    if (!hasRealData) return;
  
    const currentVariance = values[rowIdx][varianceColIdx];
    if (currentVariance !== "" && currentVariance !== null && currentVariance !== undefined) return; // already filled
  
    const excelRow = rowIdx + 1;
    sheet.getRangeByIndexes(rowIdx, varianceColIdx, 1, 1).formulas = [["=B" + excelRow + "-C" + excelRow]];
    await context.sync();
  }
  
  async function mentorCheckRowForSuggestion(context, sheet, values, address) {
    const match = address.match(/([A-Z]+)(\d+)$/i);
    if (!match) return;
    const colLetter = match[1].toUpperCase();
    const rowNum = parseInt(match[2], 10);
    const rowIdx = rowNum - 1;
    if (rowIdx < 0 || rowIdx >= values.length) { hideMentorBadge(); return; }
  
    // FIX: only B, C, D (May/Jun/Jul) are cells you actually enter figures
    // into. Column A is just the row label, and E (Variance) is a computed
    // formula — neither is somewhere you're "working," so clicking there
    // shouldn't trigger a suggestion either.
    if (!["B", "C", "D"].includes(colLetter)) {
      hideMentorBadge();
      return;
    }
  
    // Check hard cases for THIS row first — same shared check used for
    // edits, so clicking around always reflects whatever's actually true
    // for the row you're currently on, not a leftover from somewhere else.
    const hardCaseStage = mentorCheckHardCaseForRow(values, address);
    if (hardCaseStage) {
      mentorLastChangedAddress = address;
      showMentorBadge(hardCaseStage);
      return;
    }
  
    // Same for the subtotals suggestion — only show it while you're
    // actually on a relevant row, and let a click alone trigger it too,
    // not just an edit.
    const headerRowIdx = mentorFindHeaderRowIdx(values);
    const subtotalsStage = mentorCheckSubtotalsForRow(values, headerRowIdx, address);
    if (subtotalsStage) {
      mentorLastChangedAddress = address;
      showMentorBadge(subtotalsStage);
      return;
    }
  
    const monthMap = { B: "May 2026", C: "Jun 2026", D: "Jul 2026" };
    const monthLabel = monthMap[colLetter];
    if (!monthLabel) {
      hideMentorBadge(); // not one of the three month columns — nothing to suggest here
      return;
    }
  
    const colIdx = colLetter.charCodeAt(0) - 65;
    const currentVal = values[rowIdx][colIdx];
    if (currentVal !== "" && currentVal !== null && currentVal !== undefined) {
      hideMentorBadge(); // already filled and no hard case applies — nothing to say here
      return;
    }
  
    const rowLabel = String(values[rowIdx][0]).toLowerCase().trim();
    const formulaInfo = mentorBuildSourceFormula(rowLabel, monthLabel);
    if (!formulaInfo) {
      hideMentorBadge();
      return;
    }
  
    const stageId = "PULL_FORMULA_" + rowLabel + "_" + colLetter;
    if (mentorDismissedThisSession.has(stageId)) {
      hideMentorBadge();
      return;
    }
  
    // Ghost preview — write the real formula immediately, but in gray,
    // so it visibly reads as a suggestion, not yet a real figure. Moving
    // on confirms it; typing over it replaces it, same as any suggestion.
    const targetAddress = colLetter + rowNum;
    const targetRange = sheet.getRange(targetAddress);
    targetRange.formulas = [[formulaInfo.formula]];
    targetRange.format.font.color = "#B0B0B0";
    await context.sync();
  
    mentorPendingPreview = { address: targetAddress, formula: formulaInfo.formula, stageId };
    mentorCurrentStage = null;
    const container = document.getElementById("mentor-suggestion");
    if (container) {
      const shortDesc = formulaInfo.description.replace(/^I can pull /, "Pulled ").replace(/ — want me to fill it in\?$/, "");
      container.innerHTML = "<div style='color:#9aa0a6;font-size:11px;'>" + shortDesc + " — move on to keep it, or type over it to replace.</div>";
    }
  }
  
  // Builds the real formula for a given P&L row + month, reading directly from
  // the raw data sheets — the exact same references already proven correct in
  // the Monthly P&L (Answer Key) sheet, just written live instead of upfront.
  function mentorBuildSourceFormula(rowLabel, monthLabel) {
    const monthDates = {
      "May 2026": ["DATE(2026,5,1)", "DATE(2026,5,31)"],
      "Jun 2026": ["DATE(2026,6,1)", "DATE(2026,6,30)"],
      "Jul 2026": ["DATE(2026,7,1)", "DATE(2026,7,31)"]
    };
    const dates = monthDates[monthLabel];
    if (!dates) return null;
    const [mStart, mEnd] = dates;
  
    const INV = "'Raw - Supplier Invoices'";
    const BANK = "'Raw - Bank Statement'";
    const POS = "'Raw - POS Sales Report'";
    const PR = "'Raw - Payroll Summary'";
  
    const commissionRates = {
      "May 2026": { deliveroo: 0.28, uberEats: 0.30, justEat: 0.14 },
      "Jun 2026": { deliveroo: 0.30, uberEats: 0.30, justEat: 0.14 },
      "Jul 2026": { deliveroo: 0.30, uberEats: 0.30, justEat: 0.14 }
    };
  
    switch (rowLabel) {
      case "dine-in":
        return {
          formula: "=SUMIF(" + POS + "!$A:$A,\"" + monthLabel + "\"," + POS + "!$C:$C)",
          description: "I can pull Dine-In sales for " + monthLabel + " straight from the POS report — want me to fill it in?"
        };
      case "takeaway":
        return {
          formula: "=SUMIF(" + POS + "!$A:$A,\"" + monthLabel + "\"," + POS + "!$D:$D)",
          description: "I can pull Takeaway sales for " + monthLabel + " from the POS report — want me to fill it in?"
        };
      case "events / catering":
        return {
          formula: "=SUMIF(" + POS + "!$A:$A,\"" + monthLabel + "\"," + POS + "!$H:$H)",
          description: "I can pull Events/Catering sales for " + monthLabel + " from the POS report — want me to fill it in?"
        };
      case "delivery revenue (net of commission)": {
        const r = commissionRates[monthLabel];
        return {
          formula: "=SUMIF(" + POS + "!$A:$A,\"" + monthLabel + "\"," + POS + "!$E:$E)*(1-" + r.deliveroo + ")" +
            "+SUMIF(" + POS + "!$A:$A,\"" + monthLabel + "\"," + POS + "!$F:$F)*(1-" + r.uberEats + ")" +
            "+SUMIF(" + POS + "!$A:$A,\"" + monthLabel + "\"," + POS + "!$G:$G)*(1-" + r.justEat + ")",
          description: "I can pull Delivery revenue for " + monthLabel + ", already netted for each platform's commission — want me to fill it in?"
        };
      }
      case "food cost":
        return {
          formula: "=SUMIFS(" + INV + "!$F:$F," + INV + "!$E:$E,\"Food\"," + INV + "!$D:$D,\">=\"&" + mStart + "," + INV + "!$D:$D,\"<=\"&" + mEnd + ")",
          description: "I can pull Food invoices for " + monthLabel + " from Supplier Invoices — want me to fill it in?"
        };
      case "beverage cost":
        return {
          formula: "=SUMIFS(" + INV + "!$F:$F," + INV + "!$E:$E,\"Beverage\"," + INV + "!$D:$D,\">=\"&" + mStart + "," + INV + "!$D:$D,\"<=\"&" + mEnd + ")",
          description: "I can pull Beverage invoices for " + monthLabel + " from Supplier Invoices — want me to fill it in?"
        };
      case "payroll cost (excl. tronc)":
        return {
          formula: "=SUMIF(" + PR + "!$A:$A,\"" + monthLabel + "\"," + PR + "!$C:$C)" +
            "+SUMIF(" + PR + "!$A:$A,\"" + monthLabel + "\"," + PR + "!$D:$D)" +
            "+SUMIF(" + PR + "!$A:$A,\"" + monthLabel + "\"," + PR + "!$F:$F)" +
            "+SUMIF(" + PR + "!$A:$A,\"" + monthLabel + "\"," + PR + "!$G:$G)",
          description: "I can pull Payroll cost for " + monthLabel + " (excluding tronc) from the Payroll Summary — want me to fill it in?"
        };
      case "rent":
        return {
          formula: "=SUMIFS(" + BANK + "!$E:$E," + BANK + "!$F:$F,\"Rent\"," + BANK + "!$A:$A,\">=\"&" + mStart + "," + BANK + "!$A:$A,\"<=\"&" + mEnd + ")*-1",
          description: "I can pull Rent for " + monthLabel + " from the Bank Statement — want me to fill it in?"
        };
      case "utilities":
        return {
          formula: "=SUMIFS(" + BANK + "!$E:$E," + BANK + "!$F:$F,\"Utilities\"," + BANK + "!$A:$A,\">=\"&" + mStart + "," + BANK + "!$A:$A,\"<=\"&" + mEnd + ")*-1",
          description: "I can pull Utilities for " + monthLabel + " from the Bank Statement — want me to fill it in?"
        };
      case "other opex (cleaning, maintenance, marketing)":
        return {
          formula: "=SUMIFS(" + INV + "!$F:$F," + INV + "!$E:$E,\"Cleaning/Supplies\"," + INV + "!$D:$D,\">=\"&" + mStart + "," + INV + "!$D:$D,\"<=\"&" + mEnd + ")" +
            "+SUMIFS(" + INV + "!$F:$F," + INV + "!$E:$E,\"Maintenance\"," + INV + "!$D:$D,\">=\"&" + mStart + "," + INV + "!$D:$D,\"<=\"&" + mEnd + ")" +
            "+SUMIFS(" + INV + "!$F:$F," + INV + "!$E:$E,\"Marketing\"," + INV + "!$D:$D,\">=\"&" + mStart + "," + INV + "!$D:$D,\"<=\"&" + mEnd + ")",
          description: "I can pull Cleaning, Maintenance & Marketing costs for " + monthLabel + " — want me to fill it in?"
        };
      default:
        return null;
    }
  }
  
  let mentorLastChangedAddress = null;
  
  function onSheetChangedDebounced(event) {
    clearTimeout(mentorDebounceTimer);
    const address = event && event.address;
    mentorDebounceTimer = setTimeout(() => onSheetChanged(address), 250); // fast enough to feel responsive, still avoids checking mid-keystroke
  }
  
  async function onSheetChanged(changedAddress) {
    if (changedAddress) mentorLastChangedAddress = changedAddress; // remember for re-checks after dismiss/accept
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load("values, rowCount, columnCount, isNullObject");
        await context.sync();
  
        if (changedAddress) {
          const rowMatch = changedAddress.match(/(\d+)$/);
          if (rowMatch) {
            await mentorAutoFillVarianceForRow(context, sheet, parseInt(rowMatch[1], 10) - 1);
          }
        }
  
        const stage = detectReportStage(usedRange, changedAddress);
  
        if (stage && !mentorDismissedThisSession.has(stage.id)) {
          showMentorBadge(stage);
        } else if (!stage) {
          hideMentorBadge();
        }
      });
    } catch (error) {
      console.error("MENTOR onSheetChanged error:", error);
    }
  }
  
  function detectReportStage(usedRange, changedAddress) {
  
    if (usedRange.isNullObject || usedRange.rowCount <= 1) {
      return {
        id: "SETUP_LAYOUT",
        title: "Start a report layout?",
        message: "This sheet looks empty. Want me to set up a standard monthly report layout — title, period, and section headers?",
        action: "insertLayout"
      };
    }
  
    const values = usedRange.values;
  
    // Search for the actual header row using the SHARED finder (see
    // mentorFindHeaderRowIdx above) — a real header row has MULTIPLE filled
    // columns, unlike a title row which only fills one cell.
    const headerRowIdx = mentorFindHeaderRowIdx(values);
    const headerRow = headerRowIdx === -1 ? [] : values[headerRowIdx].map(v => String(v).toLowerCase());
  
    // No recognizable header row found yet — nothing to suggest at this stage
    if (headerRowIdx === -1) {
      return null;
    }
  
    // Check only the DATA rows (below the header row) for a totals row —
    // ignore the title/header rows above it
    const dataRows = values.slice(headerRowIdx + 1);
  
    // MENTOR — hard-case checks for the specific row that was just edited.
    // Extracted into a shared function so BOTH edits (onChanged) and clicks
    // (onSelectionChanged) run the exact same check — previously hard-case
    // badges only reacted to typing, so clicking around left stale badges
    // behind, or hid ones that should still apply.
    if (changedAddress) {
      const hardCaseStage = mentorCheckHardCaseForRow(values, changedAddress);
      if (hardCaseStage) return hardCaseStage;
  
      // Pure arithmetic, not judgment — but only offer it while you're
      // actually on a relevant row, not as a leftover attached to whatever
      // cell you happen to be sitting on.
      const subtotalsStage = mentorCheckSubtotalsForRow(values, headerRowIdx, changedAddress);
      if (subtotalsStage) return subtotalsStage;
    }
  
    const hasDataRows = dataRows.some(row => row.some(cell => cell !== "" && cell !== null));
    const hasTotalsRow = dataRows.some(row => String(row[0]).toLowerCase().includes("total"));
    const hasVarianceCol = headerRow.some(h => h.includes("variance") || h.includes("vs") || h.includes("budget"));
  
    // Stage: headers + at least one real data row, but no totals row yet
    if (hasDataRows && !hasTotalsRow) {
      return {
        id: "ADD_TOTALS",
        title: "Add totals?",
        message: "I see figures without a totals row. Want me to add SUM totals at the bottom of each numeric column?",
        action: "insertTotals"
      };
    }
  
    // Stage: totals exist, no variance/comparison column in headers
    if (hasTotalsRow && !hasVarianceCol) {
      return {
        id: "ADD_VARIANCE",
        title: "Add a variance column?",
        message: "Totals look good. Want me to add a variance column comparing this to last period or budget?",
        action: "insertVariance"
      };
    }
  
    return null;
  }
  
  let mentorCurrentStage = null;
  
  function showMentorBadge(stage) {
    mentorCurrentStage = stage;
    const container = document.getElementById("mentor-suggestion");
    if (!container) {
      console.error("MENTOR: missing #mentor-suggestion div in taskpane.html");
      return;
    }
  
    const accent = "#00B0F0";
  
    let html = "<div onclick='mentorExpandBadge()' style='cursor:pointer;background:#1a1a2e;border-left:4px solid " + accent + ";";
    html += "padding:6px 10px;border-radius:6px;font-size:11px;color:#cfd2d8;display:flex;align-items:center;'>";
    html += "<span style='display:inline-block;width:7px;height:7px;border-radius:50%;background:" + accent + ";margin-right:7px;'></span>";
    html += "MENTOR noticed something — tap to see</div>";
  
    container.innerHTML = html;
  }
  
  function hideMentorBadge() {
    mentorCurrentStage = null;
    const container = document.getElementById("mentor-suggestion");
    if (container) container.innerHTML = "";
  }
  
  window.mentorExpandBadge = function() {
    if (!mentorCurrentStage) return;
    const stage = mentorCurrentStage;
    const container = document.getElementById("mentor-suggestion");
    const cardBg = "#22252b";
    const accent = "#00B0F0";
  
    let html = "<div style='background:" + cardBg + ";border-left:4px solid " + accent + ";";
    html += "padding:10px 10px 10px 12px;margin:6px 0;border-radius:6px;font-size:12px;line-height:1.5;color:#fff;'>";
    html += "<strong style='color:#fff;'>" + stage.title + "</strong><br/>";
    html += "<span style='color:#cfd2d8;'>" + stage.message + "</span><br/>";
    html += "<div style='margin-top:6px;'>";
    html += "<button onclick='mentorAcceptStage(\"" + stage.action + "\")' style='margin-right:6px;padding:3px 10px;background:#28a745;color:white;border:none;border-radius:4px;cursor:pointer;font-size:11px;'>Do it for me</button>";
    html += "<button onclick='mentorShowManualSteps(\"" + stage.action + "\")' style='margin-right:6px;padding:3px 10px;background:#555;color:white;border:none;border-radius:4px;cursor:pointer;font-size:11px;'>Show me how</button>";
    html += "<button onclick='mentorDismissStage()' style='padding:3px 10px;background:#dc3545;color:white;border:none;border-radius:4px;cursor:pointer;font-size:11px;'>Not now</button>";
    html += "</div></div>";
  
    container.innerHTML = html;
  };
  
  // After accepting or dismissing something, re-check the SAME cell —
  // including whether it's now a candidate for a ghost-preview offer.
  // Previously this only ran detectReportStage (hard cases / subtotals),
  // so a still-blank cell with nothing else flagged stayed silent until you
  // clicked away and back to trigger a real selection-change event.
  async function mentorRecheckAfterAction(address) {
    if (!address) return;
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load("values, rowCount, columnCount, isNullObject");
        await context.sync();
        if (usedRange.isNullObject) { hideMentorBadge(); return; }
  
        const stage = detectReportStage(usedRange, address);
        if (stage && !mentorDismissedThisSession.has(stage.id)) {
          showMentorBadge(stage);
          return;
        }
  
        // Nothing from the usual hard-case/subtotals check — see if this
        // still-blank cell now qualifies for a ghost-preview offer instead.
        await mentorCheckRowForSuggestion(context, sheet, usedRange.values, address);
      });
    } catch (error) {
      console.error("MENTOR recheck error:", error);
    }
  }
  
  window.mentorDismissStage = function() {
    if (mentorCurrentStage) {
      mentorDismissedThisSession.add(mentorCurrentStage.id);
    }
    hideMentorBadge();
    // Immediately re-check the same row — if a second issue exists (e.g. Food
    // Cost has both a vendor-rename AND a price-ambiguity note), it should
    // surface right away, without needing you to edit the cell again.
    mentorRecheckAfterAction(mentorLastChangedAddress);
  };
  
  const mentorActionHistory = [];
  
  function mentorLogAction(description) {
    const time = new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
    mentorActionHistory.push(time + " — " + description);
    renderMentorHistory();
  }
  
  function renderMentorHistory() {
    const toggle = document.getElementById("mentor-history-toggle");
    const list = document.getElementById("mentor-history-list");
    if (!toggle) return;
  
    // Simplified per your point — the Audit Log sheet is now the real,
    // permanent record, so this doesn't need to duplicate every entry
    // inline anymore. Just a count and a pointer to where the full detail
    // actually lives.
    toggle.textContent = mentorActionHistory.length === 0
      ? "Nothing logged yet"
      : mentorActionHistory.length + " action" + (mentorActionHistory.length === 1 ? "" : "s") + " logged — see 'MENTOR Audit Log' sheet for details";
  
    if (list) list.innerHTML = "";
  }
  
  window.mentorToggleHistory = function() {
    const list = document.getElementById("mentor-history-list");
    if (!list) return;
    list.style.display = list.style.display === "none" ? "block" : "none";
  };
  
  // Persistent, permanent record — a real sheet inside the workbook itself,
  // so the audit trail survives closing and reopening the file. Replaces
  // scattering native comments across the working sheet, which gets visually
  // noisy fast once more than a couple of rows have something flagged.
  async function mentorEnsureAuditLogSheet(context) {
    let logSheet = context.workbook.worksheets.getItemOrNullObject("MENTOR Audit Log");
    await context.sync();
    if (logSheet.isNullObject) {
      logSheet = context.workbook.worksheets.add("MENTOR Audit Log");
      logSheet.getRange("A1:D1").values = [["Timestamp", "Sheet", "Cell", "What MENTOR found"]];
      logSheet.getRange("A1:D1").format.font.bold = true;
      await context.sync();
    }
    return logSheet;
  }
  
  async function mentorAppendAuditLogRow(context, sheetName, cellRef, description) {
    const logSheet = await mentorEnsureAuditLogSheet(context);
    const usedRange = logSheet.getUsedRangeOrNullObject();
    usedRange.load("rowCount");
    await context.sync();
    const nextRow = usedRange.isNullObject ? 2 : usedRange.rowCount + 1;
    const timestamp = new Date().toLocaleString([], { dateStyle: "medium", timeStyle: "short" });
    logSheet.getRange("A" + nextRow + ":D" + nextRow).values = [[timestamp, sheetName, cellRef, description]];
    await context.sync();
  }
  
  window.mentorAcceptStage = async function(action) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load("name");
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load("rowCount, columnCount, values");
        await context.sync();
  
        if (action === "documentVendorMerge" || action === "documentMissingInvoice" || action === "documentVatIssue" || action === "documentPriceAmbiguity" || action === "documentCommissionChange" || action === "documentMessyDate" || action === "documentTroncExclusion") {
          const targetAddress = mentorCurrentStage && mentorCurrentStage.targetAddress;
          const noteText = (mentorCurrentStage && mentorCurrentStage.message) || "Noted by MENTOR.";
          const label = action === "documentVendorMerge" ? "vendor merge"
            : action === "documentMissingInvoice" ? "missing invoice"
            : action === "documentVatIssue" ? "possible VAT error"
            : action === "documentPriceAmbiguity" ? "price rise / wastage question"
            : action === "documentCommissionChange" ? "commission rate change"
            : action === "documentMessyDate" ? "inconsistent date format"
            : "tronc exclusion explanation";
  
          if (targetAddress) {
            try {
              await mentorAppendAuditLogRow(context, sheet.name, targetAddress, noteText);
              mentorLogAction("Logged a " + label + " for " + targetAddress + " ('" + sheet.name + "') in the MENTOR Audit Log sheet");
              // FIX: once handled, mark it resolved — otherwise this same
              // underlying condition (data hasn't changed) keeps matching
              // every time you click back into the cell, asking again forever.
              if (mentorCurrentStage && mentorCurrentStage.id) {
                mentorDismissedThisSession.add(mentorCurrentStage.id);
              }
            } catch (writeError) {
              mentorLogAction("Tried to log a " + label + " on " + targetAddress + " but the write failed: " + writeError.message);
            }
          } else {
            mentorLogAction("Tried to document a " + label + " but had no target cell — nothing changed");
          }
        }
  
        if (action === "insertSourceFormula") {
          const targetAddress = mentorCurrentStage && mentorCurrentStage.targetAddress;
          const formula = mentorCurrentStage && mentorCurrentStage.formula;
          if (targetAddress && formula) {
            sheet.getRange(targetAddress).formulas = [[formula]];
            mentorLogAction("Pulled a formula into " + targetAddress + " from the raw data sheets, on '" + sheet.name + "'");
          } else {
            mentorLogAction("Tried to pull a formula in but was missing the target cell or formula — nothing changed");
          }
        }
  
        if (action === "insertPnlSubtotals") {
          const values = usedRange.values;
          const headerRowIdx = mentorFindHeaderRowIdx(values);
          if (headerRowIdx === -1) {
            mentorLogAction("Tried to fill P&L subtotals but could not find a header row — nothing changed");
          } else {
            const find = (label) => mentorFindSubtotalRowIdx(values, headerRowIdx, label);
            const monthCols = [1, 2, 3]; // columns B, C, D — this P&L always has exactly 3 month columns
  
            const totalRevenueIdx = find("total revenue");
            const totalCogsIdx = find("total cogs");
            const grossProfitIdx = find("gross profit");
            const totalOpexIdx = find("total operating expenses");
            const ebitdaIdx = find("ebitda");
  
            const revenueRowIdxs = ["dine-in", "takeaway", "delivery revenue (net of commission)", "events / catering"].map(find).filter(i => i !== -1);
            const cogsRowIdxs = ["food cost", "beverage cost"].map(find).filter(i => i !== -1);
            const opexRowIdxs = ["payroll cost (excl. tronc)", "rent", "utilities", "other opex (cleaning, maintenance, marketing)"].map(find).filter(i => i !== -1);
  
            let filledCount = 0;
            monthCols.forEach(c => {
              const colLetter = String.fromCharCode(65 + c);
              if (totalRevenueIdx !== -1 && revenueRowIdxs.length) {
                const firstR = revenueRowIdxs[0] + 1, lastR = revenueRowIdxs[revenueRowIdxs.length - 1] + 1;
                const cell = sheet.getRangeByIndexes(totalRevenueIdx, c, 1, 1);
                cell.formulas = [["=SUM(" + colLetter + firstR + ":" + colLetter + lastR + ")"]];
                cell.format.font.bold = true;
                filledCount++;
              }
              if (totalCogsIdx !== -1 && cogsRowIdxs.length) {
                const firstR = cogsRowIdxs[0] + 1, lastR = cogsRowIdxs[cogsRowIdxs.length - 1] + 1;
                const cell = sheet.getRangeByIndexes(totalCogsIdx, c, 1, 1);
                cell.formulas = [["=SUM(" + colLetter + firstR + ":" + colLetter + lastR + ")"]];
                cell.format.font.bold = true;
                filledCount++;
              }
              if (grossProfitIdx !== -1 && totalRevenueIdx !== -1 && totalCogsIdx !== -1) {
                const cell = sheet.getRangeByIndexes(grossProfitIdx, c, 1, 1);
                cell.formulas = [["=" + colLetter + (totalRevenueIdx + 1) + "-" + colLetter + (totalCogsIdx + 1)]];
                cell.format.font.bold = true;
                filledCount++;
              }
              if (totalOpexIdx !== -1 && opexRowIdxs.length) {
                const firstR = opexRowIdxs[0] + 1, lastR = opexRowIdxs[opexRowIdxs.length - 1] + 1;
                const cell = sheet.getRangeByIndexes(totalOpexIdx, c, 1, 1);
                cell.formulas = [["=SUM(" + colLetter + firstR + ":" + colLetter + lastR + ")"]];
                cell.format.font.bold = true;
                filledCount++;
              }
              if (ebitdaIdx !== -1 && grossProfitIdx !== -1 && totalOpexIdx !== -1) {
                const cell = sheet.getRangeByIndexes(ebitdaIdx, c, 1, 1);
                cell.formulas = [["=" + colLetter + (grossProfitIdx + 1) + "-" + colLetter + (totalOpexIdx + 1)]];
                cell.format.font.bold = true;
                filledCount++;
              }
            });
  
            mentorLogAction("Filled in " + filledCount + " P&L subtotal formulas (Total Revenue, Total COGS, Gross Profit, Total Operating Expenses, EBITDA) on '" + sheet.name + "', all in bold");
  
            // Automatic formatting pass — no separate prompt, this just
            // happens the moment the subtotals are calculated, same as an
            // editor auto-indenting code. Standard accounting presentation:
            // currency, top borders above subtotals, double underline under
            // EBITDA, red-bracket negatives on Variance.
            if (revenueRowIdxs.length && ebitdaIdx !== -1) {
              const firstDataRow = revenueRowIdxs[0];
              const lastDataRow = ebitdaIdx;
              const rowSpan = lastDataRow - firstDataRow + 1;
  
              const valueRange = sheet.getRangeByIndexes(firstDataRow, 1, rowSpan, 3); // B:D
              valueRange.numberFormat = Array.from({ length: rowSpan }, () => ["£#,##0.00", "£#,##0.00", "£#,##0.00"]);
  
              const varianceColIdx = 4; // column E
              const varianceRange = sheet.getRangeByIndexes(firstDataRow, varianceColIdx, rowSpan, 1);
              varianceRange.numberFormat = Array.from({ length: rowSpan }, () => ["£#,##0.00;[Red](£#,##0.00)"]);
  
              [totalRevenueIdx, totalCogsIdx, grossProfitIdx, totalOpexIdx, ebitdaIdx].forEach(idx => {
                if (idx === -1) return;
                const borderRange = sheet.getRangeByIndexes(idx, 0, 1, 5); // A:E
                borderRange.format.borders.getItem("EdgeTop").style = "Continuous";
              });
  
              const ebitdaBorderRange = sheet.getRangeByIndexes(ebitdaIdx, 0, 1, 5);
              ebitdaBorderRange.format.borders.getItem("EdgeBottom").style = "Double";
  
              // Clear stray 0s on the bold SECTION-HEADER rows (Revenue, COGS,
              // Operating Expenses, and the header occurrences of "Gross
              // Profit"/"EBITDA") — these never have real numbers, so any
              // leftover Variance formula there just shows a meaningless 0.
              const mentorFindFirstRowIdx = (label) => {
                for (let i = headerRowIdx + 1; i < values.length; i++) {
                  if (String(values[i][0]).toLowerCase().trim() === label) return i;
                }
                return -1;
              };
              const sectionHeaderRows = [
                mentorFindFirstRowIdx("revenue"),
                mentorFindFirstRowIdx("cost of goods sold (cogs)"),
                mentorFindFirstRowIdx("gross profit"), // first occurrence = header, not the data row
                mentorFindFirstRowIdx("operating expenses"),
                mentorFindFirstRowIdx("ebitda") // first occurrence = header, not the data row
              ];
              sectionHeaderRows.forEach(idx => {
                if (idx === -1) return;
                sheet.getRangeByIndexes(idx, varianceColIdx, 1, 1).values = [[""]];
              });
  
              mentorLogAction("Applied accounting formatting (currency, bold totals, top borders, double underline under EBITDA, red-bracket negatives) on '" + sheet.name + "'");
            }
          }
        }
  
        if (action === "insertLayout") {
          sheet.getRange("A1").values = [["Monthly Report"]];
          sheet.getRange("A2").values = [["Period:"]];
          sheet.getRange("A4:D4").values = [["Category", "This Month", "Last Month", "Variance"]];
          sheet.getRange("A1").format.font.bold = true;
          sheet.getRange("A4:D4").format.font.bold = true;
          mentorLogAction("Set up blank report layout (title, period, headers) on '" + sheet.name + "'");
        }
  
        if (action === "insertTotals") {
          const lastRow = usedRange.rowCount;
          const colCount = usedRange.columnCount;
          const totalsRowIndex = lastRow + 1;
          const totalsRange = sheet.getRangeByIndexes(totalsRowIndex - 1, 0, 1, colCount);
          const formulaRow = [];
          for (let c = 0; c < colCount; c++) {
            if (c === 0) {
              formulaRow.push("Total");
            } else {
              const colLetter = String.fromCharCode(65 + c);
              formulaRow.push("=SUM(" + colLetter + "2:" + colLetter + lastRow + ")");
            }
          }
          totalsRange.formulas = [formulaRow];
          totalsRange.format.font.bold = true;
          mentorLogAction("Added a Totals row on '" + sheet.name + "' (row " + totalsRowIndex + ")");
        }
  
        if (action === "insertVariance") {
          const colCount = usedRange.columnCount;
          const rowCount = usedRange.rowCount;
          const values = usedRange.values;
  
          // Find the header row using the SAME shared finder as detectReportStage
          // (this used to be a separate copy that never got the year-header fix —
          // that mismatch was the actual bug tonight)
          const headerRowIdx = mentorFindHeaderRowIdx(values);
  
          if (headerRowIdx === -1) {
            console.error("MENTOR: could not find header row for variance calculation");
            mentorLogAction("Tried to add a Variance column on '" + sheet.name + "' but could not find a header row — nothing changed");
          } else {
            // Find existing "Variance" column, or add a new one if it's not there yet
            let varianceColIdx = values[headerRowIdx].findIndex(v => String(v).toLowerCase().includes("variance"));
            let addedNewColumn = false;
            // FIX: the header used to just say "Variance" with no indication
            // of which two months it's actually comparing — and since the
            // formula is always ThisMonth (B) minus LastMonth (C), July was
            // silently left out of the comparison with nothing telling you so.
            const monthBLabel = String(values[headerRowIdx][1]).trim() || "This period";
            const monthCLabel = String(values[headerRowIdx][2]).trim() || "Last period";
            const varianceHeaderText = "Variance (" + monthBLabel + " vs " + monthCLabel + ")";
            if (varianceColIdx === -1) {
              // FIX: don't trust usedRange's columnCount here — Excel counts a
              // column as "used" the moment ANY formatting has ever touched
              // it, even if it's visually blank. That's what pushed Variance
              // to F instead of E after earlier formatting passes. Instead,
              // find the actual rightmost header text and place it right
              // after that.
              let lastRealHeaderCol = -1;
              for (let c = 0; c < values[headerRowIdx].length; c++) {
                if (String(values[headerRowIdx][c]).trim() !== "") lastRealHeaderCol = c;
              }
              varianceColIdx = lastRealHeaderCol + 1;
              const headerCell = sheet.getRangeByIndexes(headerRowIdx, varianceColIdx, 1, 1);
              headerCell.values = [[varianceHeaderText]];
              headerCell.format.font.bold = true;
              addedNewColumn = true;
            } else {
              // Column already exists (e.g. from an earlier run) — make sure
              // its label is accurate too, in case it was left generic before.
              sheet.getRangeByIndexes(headerRowIdx, varianceColIdx, 1, 1).values = [[varianceHeaderText]];
            }
  
            // Fill in the actual formula for each data row: ThisMonth (col B) - LastMonth (col C)
            // Assumes standard layout: A=Category, B=This Month, C=Last Month, D=Variance
            let formulasFilled = 0;
            for (let r = headerRowIdx + 1; r < rowCount; r++) {
              const rowLabel = String(values[r][0]).toLowerCase();
              if (rowLabel === "total" || rowLabel === "") continue; // skip totals row and blanks
  
              // FIX: section-header rows (Revenue, COGS, Operating Expenses,
              // and the bold "Gross Profit"/"EBITDA" header rows) have no real
              // numbers in B or C at all — writing a formula there just shows
              // a stray 0-0=0. Skip any row where both are genuinely blank.
              const bVal = values[r][1], cVal = values[r][2];
              const hasRealData = (typeof bVal === "number" && bVal !== 0) || (typeof cVal === "number" && cVal !== 0);
              if (!hasRealData) continue;
  
              const excelRow = r + 1; // convert to 1-based Excel row number
              const varianceCell = sheet.getRangeByIndexes(r, varianceColIdx, 1, 1);
              varianceCell.formulas = [["=B" + excelRow + "-C" + excelRow]];
              formulasFilled++;
            }
            const colLetter = String.fromCharCode(65 + varianceColIdx);
            mentorLogAction(
              (addedNewColumn ? "Added a new 'Variance' column (" + colLetter + ") " : "Filled the existing Variance column (" + colLetter + ") ") +
              "on '" + sheet.name + "' with " + formulasFilled + " formulas"
            );
          }
        }
  
        await context.sync();
      });
    } catch (error) {
      console.error("MENTOR action error:", error);
      mentorLogAction("An action failed with an error: " + error.message);
    }
  
    hideMentorBadge();
    // Same as dismiss — re-check the same row right away in case a second
    // issue is waiting (e.g. the price-ambiguity note after accepting the
    // vendor-merge one on Food Cost), OR the cell is now a candidate for a
    // ghost-preview offer since selection never actually changed after
    // clicking "Do it for me."
    mentorRecheckAfterAction(mentorLastChangedAddress);
  };
  
  const mentorManualSteps = {
    insertLayout: "1. Type 'Monthly Report' in A1.\n2. Type 'Period:' in A2.\n3. Add column headers (Category, This Month, Last Month, Variance) in row 4.",
    insertTotals: "1. Go to the row below your last data row.\n2. Type 'Total' in column A.\n3. In each numeric column, use =SUM() over the data rows.",
    insertVariance: "1. Add a 'Variance' header in the next empty column.\n2. Use =ThisMonth-LastMonth for each row.",
    documentVendorMerge: "1. Right-click the Food Cost cell and insert a comment.\n2. Note that both vendor names refer to the same supplier, renamed mid-period.\n3. Keep invoices from both names included in your total — don't drop either one.",
    documentMissingInvoice: "1. Check the Bank Statement sheet for the payment in question.\n2. Contact the vendor and request the missing invoice.\n3. Note the follow-up directly on this cell until the invoice arrives, so it isn't forgotten.",
    documentVatIssue: "1. Open the supplier invoice in question.\n2. Confirm with the supplier or your VAT records whether 20% should apply.\n3. Correct the VAT rate on the invoice and note the correction here for your records.",
    documentPriceAmbiguity: "1. Check with the site manager whether wastage was unusually high this period.\n2. Compare against the supplier's price-increase letter to see how much of the rise it actually explains.\n3. Note whichever explanation you confirm — don't assume one without checking.",
    documentCommissionChange: "1. Check the platform's latest rate card or invoice for the confirmed commission rate.\n2. Recalculate gross vs net if the rate has genuinely changed.\n3. Note the confirmed rate here so next month's comparison is accurate.",
    documentMessyDate: "1. Open the Bank Statement sheet and find the payout in question.\n2. Retype the date in the same format as the surrounding rows.\n3. Double-check any date-based formulas still pick it up correctly.",
    documentTroncExclusion: "1. Confirm with your payroll provider that tronc is fully redistributed, not retained.\n2. If that's correct, no change needed — it's meant to sit outside Payroll Cost.\n3. If your business handles tronc differently, let MENTOR know so this rule can be adjusted.",
    insertPnlSubtotals: "1. In Total Revenue, use =SUM() over the Dine-In through Events rows.\n2. In Total COGS, use =SUM() over Food Cost and Beverage Cost.\n3. Gross Profit = Total Revenue - Total COGS.\n4. In Total Operating Expenses, use =SUM() over Payroll through Other Opex.\n5. EBITDA = Gross Profit - Total Operating Expenses.",
    insertSourceFormula: "1. Open the matching raw data sheet for this row.\n2. Filter or sum the rows for this specific month.\n3. Type the total (or build the same SUMIFS formula) into this cell yourself."
  };
  
  window.mentorShowManualSteps = function(action) {
    const container = document.getElementById("mentor-suggestion");
    const steps = mentorManualSteps[action] || "Steps not available yet.";
    container.innerHTML += "<div style='background:#1a1a2e;color:#cfd2d8;font-size:11px;padding:8px;border-radius:6px;margin-top:6px;white-space:pre-line;'>" + steps + "</div>";
  };