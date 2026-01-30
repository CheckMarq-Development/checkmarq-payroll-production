/****************************************************
 * CheckMarq Payroll Production
 * CLEAN FOUNDATION — SAFE IMPORT ONLY
 ****************************************************/

const SHEET_NAMES = {
  ADMIN: "Admin_Config",
  RAW: "Raw_All_Visits"
};

const PDF_STATE_KEY = "PHASE4_PDF_STATE";

/************ MENU ************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // 1️⃣ Payroll Automation
  ui.createMenu("Payroll Automation")
    .addItem("Reset for New Pay Period", "resetForNewPayPeriod")
    .addItem("Import Raw Data", "importRawData")
    .addItem("Recalculate Payroll & Invoices", "recalculateStub_")
    .addSeparator()
    .addItem("Generate Payroll PDFs", "generatePayrollPdfs")
    .addItem("Generate Invoice PDFs", "generateInvoicePdfs")
    .addSeparator()
    .addItem("Reset PDF Creation", "resetPdfState_")
    .addItem("Clean Temp Tabs", "cleanupTempSheets_")
    .addToUi();

  // 2️⃣ Email Automation
  addEmailMenus_();

  // 3️⃣ Final Export (PBPT lives here now)
  addFinalExportMenu_();
}

function addFinalExportMenu_() {
  SpreadsheetApp.getUi()
    .createMenu("Final Export")
    .addItem("PBPT – Export Final", "exportPBPTFinal_")
    .addSeparator()
    .addItem("Publish Final Snapshot", "publishFinalWorkbookCopy")
    .addToUi();
}


function addEmailMenus_() {
  SpreadsheetApp.getUi()
    .createMenu("Email Automation")
    .addItem("Create Payroll Email Drafts", "buildPayrollEmailDrafts")
    .addItem("Create Invoice Email Drafts", "buildInvoiceEmailDrafts")
    .addSeparator()
    .addItem("Send All Email Drafts", "sendPhase8Drafts") // ← add this line
    .addSeparator()
    .addItem("Open Email Audit", "openEmailAudit_")
    .addToUi();
}

function onOpen_SheetHelpers_() {
  SpreadsheetApp.getUi()
    .createMenu("Payroll")
    .addItem("Reset for New Pay Period", "resetForNewPayPeriod")
    .addSeparator()
    .addItem("Import Raw Visits", "importRawVisits") // existing
    .addItem("Rebuild Payroll & Invoices", "recalculatePayrollAndInvoices") // existing
    .addSeparator()
    .addItem("Publish Final Snapshot", "publishFinalWorkbookCopy")
    .addToUi();
}

/************ ADMIN CONFIG ************/
function getAdminConfig_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.ADMIN);
  if (!sh) throw new Error("Missing Admin_Config sheet");

  const data = sh.getDataRange().getValues();
  const cfg = {};
  data.slice(1).forEach(r => {
    if (r[0]) cfg[String(r[0]).trim()] = r[1];
  });

  const required = ["Payroll Reports Drive Folder ID", "Approved From", "Approved To"];
  required.forEach(k => {
    if (!cfg[k]) throw new Error(`Admin_Config missing: ${k}`);
  });

  cfg.approvedFrom = normalizeDate_(cfg["Approved From"]);
  cfg.approvedTo   = endOfDay_(normalizeDate_(cfg["Approved To"]));

  return cfg;
}

/************ IMPORT RAW DATA ************/
function importRawData() {
   assertPayPeriodConfigured_();
  const ss = SpreadsheetApp.getActive();
  const rawSheet = ss.getSheetByName(SHEET_NAMES.RAW);
  if (!rawSheet) throw new Error("Missing Raw_All_Visits sheet");

  const cfg = getAdminConfig_();
  const folder = DriveApp.getFolderById(cfg["Payroll Reports Drive Folder ID"]);

  // Get most recent Google Sheet
  let latestFile = null;
  let latestTime = 0;
  const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (files.hasNext()) {
    const f = files.next();
    const t = f.getLastUpdated().getTime();
    if (t > latestTime) {
      latestTime = t;
      latestFile = f;
    }
  }
  if (!latestFile) throw new Error("No Google Sheets found in folder");

  const src = SpreadsheetApp.open(latestFile).getSheets()[0];
  const range = src.getDataRange();
  const values = range.getValues();
  const headers = values[0];
  const rows = values.slice(1);

  const statusIdx = headers.indexOf("Visit status");
  const approvedIdx = headers.indexOf("Date when HA approved the Visit");

  if (statusIdx === -1 || approvedIdx === -1) {
    throw new Error("Required columns missing in source sheet");
  }

  const filtered = rows.filter(r => {
    const status = String(r[statusIdx] || "").toLowerCase();
    if (status === "rejected") return false;

    const approvedDate = normalizeDate_(r[approvedIdx]);
    if (!approvedDate) return false;

    return approvedDate >= cfg.approvedFrom && approvedDate <= cfg.approvedTo;
  });

  // Rewrite Raw_All_Visits
  rawSheet.clearContents();
  rawSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (filtered.length) {
    rawSheet.getRange(2, 1, filtered.length, headers.length).setValues(filtered);
  }

  applyColumnRules_(rawSheet);

  SpreadsheetApp.getUi().alert(
    `Imported ${filtered.length} rows from:\n${latestFile.getName()}`
  );
}

/************ COLUMN RULES ************/
function applyColumnRules_(sheet) {
  // Hide G, H, J, M, N (1-based)
  [7, 8, 10, 13, 14].forEach(c => sheet.hideColumns(c));

  // Delete P if present
  if (sheet.getLastColumn() >= 16) {
    sheet.deleteColumn(16);
  }
}

/************ DATE HELPERS ************/
function normalizeDate_(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === "number") {
    return new Date(Math.round((v - 25569) * 86400 * 1000));
  }
  if (!v) return null;
  const d = new Date(v);
  return isNaN(d) ? null : d;
}

function endOfDay_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 23, 59, 59);
}

/************ SAFE STUBS ************/
/************************************************************
 * MASTER RECALC — PAYROLL + INVOICE + SUMMARY + AUDIT
 * Raw_All_Visits is READ ONLY
 ************************************************************/
function recalculateStub_() {
   assertPayPeriodConfigured_();
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  // Guardrail
  if (!ss.getSheetByName("Raw_All_Visits")) {
    ui.alert("ERROR: Raw_All_Visits sheet not found.");
    return;
  }

  try {
    // 1. Payroll (D9 + D10)
    buildPayrollAndSummary();


    // 2. Summary (if you have it)
    if (typeof buildPayrollSummary_ === "function") {
      buildPayrollSummary_();
    }

    // 3. Invoices (if present)
    if (typeof buildInvoices === "function") {
      buildInvoices();
    }

    // 4. Audit (last, read-only)
    if (typeof buildAuditTab === "function") {
      buildAuditTab();
    }

    ui.alert(
      "Recalculation complete:\n\n" +
      "• Payroll rebuilt\n" +
      "• Summary updated\n" +
      "• Invoices rebuilt\n" +
      "• Audit refreshed"
    );

  } catch (err) {
    ui.alert("Recalculation FAILED:\n\n" + err.message);
    throw err;
  }
}


function resetPdfState_() {
  PropertiesService.getScriptProperties().deleteProperty(PDF_STATE_KEY);
  SpreadsheetApp.getUi().alert("PDF state reset.");
}
/****************************************************
 * PAYROLL BUILD — EXACT ORDER MATCH (PRODUCTION)
 * Depends ONLY on Raw_All_Visits
 ****************************************************/

function recalculatePayrollOnly() {
    assertPayPeriodConfigured_();
  const ss = SpreadsheetApp.getActive();
  buildPayrollBucket_(ss, "D9");
  buildPayrollBucket_(ss, "D10");
  SpreadsheetApp.getUi().alert("Payroll tabs rebuilt.");
}

function buildPayrollBucket_(ss, bucket) {
  const raw = ss.getSheetByName("Raw_All_Visits");
  if (!raw || raw.getLastRow() < 2) {
    throw new Error("Raw_All_Visits missing or empty");
  }

  const data = raw.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const idx = {
    first: headers.indexOf("Assigned Clinician First Name"),
    last: headers.indexOf("Assigned Clinician Last Name"),
    patient: headers.indexOf("Patient name"),
    visit: headers.indexOf("Visit type"),
    date: headers.indexOf("Visit scheduled date"),
    ha: headers.indexOf("HA Name"),
    pay: headers.indexOf("Price agreed between HA & Clinician")
  };

  Object.entries(idx).forEach(([k,v]) => {
    if (v === -1) throw new Error(`Missing column: ${k}`);
  });

  const out = [];

  rows.forEach(r => {
    const ha = String(r[idx.ha] || "").trim();
    if (!ha.startsWith(bucket + " ")) return;

    const d = r[idx.date];
    if (!(d instanceof Date) || isNaN(d)) return;

    out.push([
      r[idx.first],   // 0
      r[idx.last],    // 1
      r[idx.patient], // 2
      r[idx.visit],   // 3
      d,              // 4
      "",             // 5 Rate blank in payroll
      ha,             // 6
      num_(r[idx.pay])// 7 Pay (can be 0)
    ]);
  });

  stablePayrollSort_(out);

  writePayrollSheet_(
    ss,
    `${bucket}_Payroll`,
    ["First Name","Last Name","Patient Name","Visit Type","Date","Rate","HA Name","Pay"],
    out
  );
}

/************ STABLE PAYROLL SORT (MATCHES SCREENSHOT) ************/
function stablePayrollSort_(rows) {
  rows
    .map((r,i)=>({r,i}))
    .sort((a,b)=>{
      // 1. Scheduled Date
      const d = cmpDate_(a.r[4], b.r[4]);
      if (d) return d;

      // 2. Patient Name
      const p = cmpText_(a.r[2], b.r[2]);
      if (p) return p;

      // 3. Clinician Last
      const l = cmpText_(a.r[1], b.r[1]);
      if (l) return l;

      // 4. Clinician First
      const f = cmpText_(a.r[0], b.r[0]);
      if (f) return f;

      // 5. Visit Type
      const v = cmpText_(a.r[3], b.r[3]);
      if (v) return v;

      return a.i - b.i; // stability
    })
    .forEach((o,i)=> rows[i] = o.r);
}

function cmpDate_(a,b){
  const at = a instanceof Date ? a.getTime() : 0;
  const bt = b instanceof Date ? b.getTime() : 0;
  return at - bt;
}

function cmpText_(a,b){
  return String(a || "").trim().localeCompare(String(b || "").trim());
}

/************ WRITE + FORMAT ************/
function writePayrollSheet_(ss, name, headers, rows) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  if (sh.getFilter()) sh.getFilter().remove();
  sh.clearContents();

  sh.getRange(1,1,1,headers.length).setValues([headers]);
  if (rows.length) {
    sh.getRange(2,1,rows.length,headers.length).setValues(rows);
    sh.getRange(2,5,rows.length,1).setNumberFormat("m/d/yyyy");
    sh.getRange(2,8,rows.length,1).setNumberFormat("$#,##0.00").setBackground("#d9ead3");
  }

  sh.getRange(1,1,1,headers.length)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#eeeeee");

  sh.setFrozenRows(1);
  sh.getRange(1,1,1,headers.length).createFilter();
}


function resetPdfState_() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  Logger.log("PDF state reset");
}

function exportPBPTFinal_() {
  const ss = SpreadsheetApp.getActive();
  const admin = ss.getSheetByName("Admin_Config");
  if (!admin) throw new Error("Admin_Config not found");

  // ---- Pay period ----
  const { start, end } = getAdminDates_();
  const tz = Session.getScriptTimeZone();
  const startStr = Utilities.formatDate(start, tz, "MM-dd-yyyy");
  const endStr   = Utilities.formatDate(end, tz, "MM-dd-yyyy");

  const folderName = `${startStr} to ${endStr}`;
  const fileName = `PBPT_${startStr}_to_${endStr}`;

  // ---- Parent folder (given ID) ----
  const parent = DriveApp.getFolderById("1u0dpja4pU9qPhhjOAut9J7eqi_zsGKQP");

  // ---- Create / get period folder ----
  const periodFolder = getOrCreateFolder_(parent, folderName);

  // ---- Create destination spreadsheet ----
  const out = SpreadsheetApp.create(fileName);
  const pbptPayroll =
  out.getSheetByName("Payroll") || out.insertSheet("Payroll");

buildPBPTPayrollFromDerived_(pbptPayroll);

  const outFile = DriveApp.getFileById(out.getId());
  periodFolder.addFile(outFile);
  DriveApp.getRootFolder().removeFile(outFile); // keep Drive clean

  // ---- Tabs to export (values only) ----
  const TABS = [
    "D9_Payroll_Summary",
    "D9_Invoice",
    "D9_Invoice_Summary"
  ];

  TABS.forEach(name => {
    const src = ss.getSheetByName(name);
    if (!src) throw new Error(`Missing sheet: ${name}`);

    const dst = out.insertSheet(name);
    const data = src.getDataRange().getValues();

    if (data.length) {
      dst.getRange(1, 1, data.length, data[0].length).setValues(data);
    }

    dst.setFrozenRows(src.getFrozenRows());
    dst.autoResizeColumns(1, dst.getLastColumn());
  });

  // Remove default blank sheet
  const blank = out.getSheetByName("Sheet1");
  if (blank) out.deleteSheet(blank);

  SpreadsheetApp.getUi().alert(
    "PBPT export complete.\n\n" +
    `Folder: ${folderName}\n` +
    `File: ${fileName}`
  );
}
function buildPBPTPayrollFromDerived_(pbptSheet) {
  const ss = SpreadsheetApp.getActive();
  const raw = ss.getSheetByName("Raw_All_Visits_DERIVED");
  if (!raw) throw new Error("Raw_All_Visits_DERIVED not found");

  const data = raw.getDataRange().getValues();
  if (data.length <= 1) return;

  const headers = data[0].map(h => String(h).trim());
  const rows = data.slice(1);

  const idx = h => {
    const i = headers.findIndex(x => x.toLowerCase() === h.toLowerCase());
    if (i === -1) throw new Error(`Missing column in DERIVED: ${h}`);
    return i;
  };

  const COL = {
    FIRST: idx("Assigned Clinician First Name"),
    LAST: idx("Assigned Clinician Last Name"),
    PATIENT: idx("Patient name"),
    VISIT: idx("Visit type"),
    DATE: idx("Visit scheduled date"),
    HA: idx("HA Name"),
    PAY: idx("Price agreed between HA & Clinician")
  };

  const out = rows
    .filter(r => r.join("").trim() !== "")
    .map(r => [
      r[COL.FIRST],
      r[COL.LAST],
      r[COL.PATIENT],
      r[COL.VISIT],
      r[COL.DATE],
      r[COL.RATE],     // Rate
      r[COL.HA],
      r[COL.PAY]      // Pay
    ]);

  pbptSheet.clear();

  pbptSheet.getRange(1,1,1,8).setValues([[
    "First Name","Last Name","Patient name","Visit type",
    "Date","Rate","HA Name","Pay"
  ]]).setFontWeight("bold");

  if (out.length) {
    pbptSheet.getRange(2,1,out.length,8).setValues(out);
  }

  pbptSheet.getRange("E:E").setNumberFormat("mm/dd/yyyy");
  pbptSheet.getRange("F:F").setNumberFormat("$#,##0.00");
  pbptSheet.getRange("H:H").setNumberFormat("$#,##0.00");

  formatPBPTPayroll_(pbptSheet);
}
function formatPBPTPayroll_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1) return;

  sheet.getRange(2, 1, lastRow - 1, lastCol).sort([
    { column: 1, ascending: true }, // First
    { column: 2, ascending: true }, // Last
    { column: 3, ascending: true }, // Patient
    { column: 5, ascending: true }  // Date
  ]);

  buildPBPTClinicianTotals_(sheet);
}

function buildPBPTClinicianTotals_(sheet) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const FIRST = headers.indexOf("First Name");
  const LAST  = headers.indexOf("Last Name");
  const PAY   = headers.indexOf("Pay");

  const map = {};

  rows.forEach(r => {
    const key = `${r[FIRST]}|${r[LAST]}`;
    const pay = Number(String(r[PAY] || "").replace(/[$,]/g,"")) || 0;
    if (!map[key]) map[key] = { first: r[FIRST], last: r[LAST], total: 0 };
    map[key].total += pay;
  });

  const output = Object.values(map)
    .sort((a,b) =>
      a.last.localeCompare(b.last) || a.first.localeCompare(b.first)
    )
    .map(r => [r.first, r.last, r.total]);

  const startCol = sheet.getLastColumn() + 2;

  sheet.getRange(1, startCol, 1, 3)
    .setValues([["First Name", "Last Name", "Total"]])
    .setFontWeight("bold");

  if (output.length) {
    sheet.getRange(2, startCol, output.length, 3).setValues(output);
    sheet.getRange(2, startCol + 2, output.length, 1)
      .setNumberFormat("$#,##0.00");
  }

  sheet.autoResizeColumns(startCol, 3);
}

