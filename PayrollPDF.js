/***************************************************
 * PHASE 6 — CLINICIAN PAYROLL PDF GENERATION
 * - One PDF per clinician
 * - Source: rebuilt *_Payroll tabs only
 * - Pay period: Admin_Config!B11:B12
 ***************************************************/

const EXPORT_SLEEP_MS = 2500;
const BACKOFF_MS = 180000; // 3 minutes
const MANUAL_FULL_RUN = true;

/**
 * ENTRY POINT
 */
function generatePayrollPdfs() {
    cleanupTempSheets_();
    
  const ss = SpreadsheetApp.getActive();

  const admin = ss.getSheetByName("Admin_Config");
  if (!admin) throw new Error("Admin_Config tab not found");

const dates = getAdminDates_();
const tz = Session.getScriptTimeZone();

const periodStart = Utilities.formatDate(dates.start, tz, "MM/dd/yyyy");
const periodEnd   = Utilities.formatDate(dates.end, tz, "MM/dd/yyyy");

 const rootCell = admin
  .createTextFinder("Payroll Reports Drive Folder ID")
  .findNext();

if (!rootCell) {
  throw new Error("Admin_Config missing 'Payroll Reports Drive Folder ID'");
}

const rootId = admin.getRange(rootCell.getRow(), 2).getValue();
if (!rootId) {
  throw new Error("Payroll Reports Drive Folder ID value is blank");
}

const parent = DriveApp.getFolderById(rootId);



   const rootNameCell = admin
  .createTextFinder("Output PDFs Root Folder Name")
  .findNext();

if (!rootNameCell) {
  throw new Error("Admin_Config missing 'Output PDFs Root Folder Name'");
}

const rootFolderName = admin
  .getRange(rootNameCell.getRow(), 2)
  .getDisplayValue()
  .trim();

if (!rootFolderName) {
  throw new Error("Output PDFs Root Folder Name value is blank");
}

const root = getOrCreateFolder_(parent, rootFolderName);

  const periodFolder = getOrCreateFolder_(root, `${periodStart} - ${periodEnd}`);

  const payrollSheets = ss.getSheets().filter(s =>
    s.getName().endsWith("_Payroll")
  );

payrollSheets.forEach(sheet => {
  resetPayrollPdfCursor_(); // ← ADD THIS LINE

  const bucket = sheet.getName().split("_")[0];
  const bucketFolder = getOrCreateFolder_(periodFolder, bucket);
  const payrollFolder = getOrCreateFolder_(bucketFolder, "Payroll");

  processPayrollSheet_(ss, sheet, periodStart, periodEnd, bucketFolder);
});

}

/**
 * PROCESS ONE PAYROLL SHEET (GROUP BY CLINICIAN)
 */
function processPayrollSheet_(ss, sheet, periodStart, periodEnd, folder) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const headers = data[0];
  const rows = data.slice(1).filter(r => r.join("") !== "");
  if (!rows.length) return;

  const COL = indexColumns_(headers, {
    first:   "First Name",
    last:    "Last Name",
    patient: "Patient Name",
    visit:   "Visit Type",
    date:    "Visit Date",
    ha:      "HA Name",
    pay:     "Pay"
  });
["first","last","patient","visit","date","ha","pay"].forEach(k => {
  if (COL[k] === undefined) {
    throw new Error(`Missing required column "${k}" in sheet ${sheet.getName()}`);
  }
});

  const clinicianMap = {};
  rows.forEach(r => {
    const name = `${r[COL.first]} ${r[COL.last]}`.trim();
    if (!name) return;
    if (!clinicianMap[name]) clinicianMap[name] = [];
    clinicianMap[name].push(r);
  });

const names = Object.keys(clinicianMap);
const cursor = getPayrollPdfCursor_();
const MAX_PER_RUN = MANUAL_FULL_RUN ? Infinity : 10;

Logger.log(`Payroll PDF resume starting at index ${cursor} of ${names.length}`);

for (let i = cursor; i < names.length; i++) {
  const clinicianName = names[i];

  try {
    buildClinicianPayrollPdf_(
      ss,
      clinicianName,
      clinicianMap[clinicianName],
      COL,
      periodStart,
      periodEnd,
      folder
    );
  } catch (e) {
    logPayrollPdfFailure_(clinicianName, e);
  }

  setPayrollPdfCursor_(i + 1);
  Utilities.sleep(EXPORT_SLEEP_MS);

if (i - cursor + 1 >= MAX_PER_RUN) {
  clearPayrollPdfTriggers_(); // ← critical
  ScriptApp.newTrigger("generatePayrollPdfs")
    .timeBased()
    .after(1 * 60 * 1000)
    .create();
  return;
}
} // ← closes the for-loop

// Finished entire sheet
clearPayrollPdfTriggers_();
resetPayrollPdfCursor_();



}

/**
 * BUILD ONE CLINICIAN PAYROLL PDF
 */
function buildClinicianPayrollPdf_(
  ss,
  clinicianName,
  rows,
  COL,
  periodStart,
  periodEnd,
  folder
) {
  const totalVisits = rows.length;
  const totalPay = rows.reduce((s, r) => s + Number(r[COL.pay] || 0), 0);

  const safeName = clinicianName.replace(/[^\w\s-]/g, "").substring(0, 90);
  const tmp = ss.insertSheet(`__TMP_PAYROLL_${safeName}_${Date.now()}`);
  tmp.clear();

  /* ===== HEADER ===== */
  tmp.getRange("A1").setValue(clinicianName);
  tmp.getRange("A2").setValue(`Pay Period: ${periodStart} – ${periodEnd}`);

  tmp.getRange("F1").setValue(`Total Visits: ${totalVisits}`);
  tmp.getRange("F2").setValue(`Total Pay: ${formatCurrency_(totalPay)}`);

  tmp.getRange("A1:F2")
    .setFontWeight("bold")
    .setFontSize(14);

  tmp.getRange("F1:F2").setHorizontalAlignment("right");

  /* ===== TABLE HEADER ===== */
  const tableHeaders = [
    "Patient Name",
    "Visit Type",
    "Visit Date",
    "HA Name",
    "Pay"
  ];

  tmp.getRange("A4:E4")
  .setValues([tableHeaders])
  .setFontWeight("bold")
  .setHorizontalAlignment("center");


  /* ===== TABLE DATA ===== */
  const tableData = rows.map(r => [
    r[COL.patient],
    r[COL.visit],
    r[COL.date],
    r[COL.ha],
    r[COL.pay]
  ]);

  tmp.getRange(5, 1, tableData.length, tableHeaders.length)
    .setValues(tableData);

  tmp.getRange("C5:C")
  .setNumberFormat("MM/dd/yyyy")
  .setHorizontalAlignment("center");
  tmp.getRange("E5:E")
  .setNumberFormat("$#,##0.00")
  .setHorizontalAlignment("center");


  tmp.autoResizeColumns(1, 6);
  // Improve readability
tmp.setColumnWidth(1, 260); // Patient Name
tmp.setColumnWidth(2, 260); // Visit Type
tmp.setColumnWidth(3, 140); // Visit Date
tmp.setColumnWidth(4, 320); // HA Name
tmp.setColumnWidth(5, 120); // Pay

const payrollFolder = getOrCreateFolder_(folder, "Payroll");

// ===== EXPORT =====
SpreadsheetApp.flush();

let status = "UNKNOWN";
let notes = "";

try {
  const result = exportPayrollSheetToPdf_(tmp, payrollFolder, safeName, false);

  status = result || "UNKNOWN"; // critical: prevents blank
  if (status === "SKIPPED") notes = "PDF already existed";

} catch (err) {
  status = "FAILED";
  notes = err && err.message ? err.message : String(err);

  if (String(err).includes("429")) {
    Utilities.sleep(BACKOFF_MS);
    // continue on 429 (skip clinician)
    return;
  }
  throw err;

} finally {
  writePayrollPdfAudit_(
    clinicianName,
    folder.getName(),
    `${safeName}.pdf`,
    status,
    notes
  );

  ss.deleteSheet(tmp); // keep this here so it ALWAYS deletes
}


}

/**
 * EXPORT TEMP SHEET TO PDF
 * OPTION B — resume-only
 */
function exportPayrollSheetToPdf_(sheet, folder, filename, showGridlines) {
  const pdfName = `${filename}.pdf`;

const existing = folder.getFilesByName(pdfName);
while (existing.hasNext()) {
  const f = existing.next();
  if (!f.isTrashed()) {
    return "SKIPPED"; // a real, non-trashed copy exists
  }
}
// if we only found trashed copies (or none), proceed to create


  const ss = SpreadsheetApp.getActive();
  const url =
    `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?` +
    `format=pdf&gid=${sheet.getSheetId()}` +
    `&portrait=false&fitw=true` +
    `&sheetnames=false&pagenumbers=false` +
    `&gridlines=${showGridlines ? "true" : "false"}`;

  const blob = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` }
  }).getBlob().setName(pdfName);

  withBackoff_(() => {
    folder.createFile(blob);
  });

  return "CREATED"; // ⬅️ REQUIRED
}


function getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function indexColumns_(headers, map) {
  const out = {};
  headers.forEach((h, i) => {
    Object.keys(map).forEach(k => {
      if (h === map[k]) out[k] = i;
    });
  });
  return out;
}

function formatCurrency_(n) {
  return `$${Number(n).toFixed(2)}`;
}
function withBackoff_(fn, maxRetries = 5) {
  let delay = 1000; // 1 second

  for (let i = 0; i < maxRetries; i++) {
    try {
      return fn();
    } catch (e) {
      if (!String(e).includes("429")) throw e;
      Utilities.sleep(delay);
      delay *= 2; // exponential backoff
    }
  }

  throw new Error("Repeated 429 errors — aborting PDF generation");
}
function logPayrollPdfFailure_(clinician, error) {
  const ss = SpreadsheetApp.getActive();
  const sheet = getOrCreateSheet("Payroll_PDF_Errors", [
    "Timestamp",
    "Clinician",
    "Error"
  ]);

  sheet.appendRow([
    new Date(),
    clinician,
    String(error)
  ]);
}
function getPayrollPdfCursor_() {
  return Number(PropertiesService.getScriptProperties()
    .getProperty("PAYROLL_PDF_CURSOR") || 0);
}

function setPayrollPdfCursor_(n) {
  PropertiesService.getScriptProperties()
    .setProperty("PAYROLL_PDF_CURSOR", String(n));
}

function resetPayrollPdfCursor_() {
  PropertiesService.getScriptProperties()
    .deleteProperty("PAYROLL_PDF_CURSOR");
}
function writePayrollPdfAudit_(clinician, bucket, pdfName, status, notes) {
  const sheet = getOrCreateSheet("Payroll_PDF_Audit", [
    "Timestamp",
    "Clinician",
    "Bucket",
    "PDF Name",
    "Status",
    "Notes"
  ]);

const data = sheet.getDataRange().getValues();
for (let i = data.length - 1; i > 0; i--) {
  if (
    data[i][1] === clinician &&
    data[i][2] === bucket &&
    data[i][3] === pdfName
  ) {
    sheet.deleteRow(i + 1);
    break;
  }
}

sheet.appendRow([
  new Date(),
  clinician,
  bucket,
  pdfName,
  status,
  notes || ""
]);

}

function clearPayrollPdfTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "generatePayrollPdfs") {
      ScriptApp.deleteTrigger(t);
    }
  });
}
function __resetPayrollPdfState() {
  resetPayrollPdfCursor_();
  clearPayrollPdfTriggers_();
}

