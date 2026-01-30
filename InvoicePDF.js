/***************************************************
 * PHASE 7 — AGENCY INVOICE PDF GENERATION
 * - One PDF per HA
 * - Source: rebuilt *_Invoice tabs only
 * - Pay period: Admin_Config!B11:B12
 ***************************************************/


/**
 * ENTRY POINT
 */
function generateInvoicePdfs() {
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

  const invoiceSheets = ss.getSheets().filter(s =>
    s.getName().endsWith("_Invoice")
  );

  invoiceSheets.forEach(sheet => {
    processInvoiceSheet_(ss, sheet, periodStart, periodEnd, periodFolder);
  });
}

/**
 * PROCESS ONE INVOICE SHEET (GROUP BY HA)
 */
function processInvoiceSheet_(ss, sheet, periodStart, periodEnd, rootFolder) {
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

const headers = data[0];
const allRows = data.slice(1).filter(r => r.join("") !== "");

const COL = indexColumns_(headers, {
  first:   "Clinician First Name",
  last:    "Clinician Last Name",
  patient: "Patient Name",
  visit:   "Visit Type",
  date:    "Visit Date",
  rate:    "Rate",
  ha:      "HA Name"
});
["first","last","patient","visit","date","rate","ha"].forEach(k => {
  if (COL[k] === undefined) {
    throw new Error(`Missing required column "${k}" in sheet ${sheet.getName()}`);
  }
});

// 🔴 APPLY BUSINESS RULE HERE
const rows = allRows.filter(r => Number(r[COL.rate]) > 0);
if (!rows.length) return;

["first","last","patient","visit","date","rate","ha"].forEach(k => {
  if (COL[k] === undefined) {
    throw new Error(`Missing required column "${k}" in sheet ${sheet.getName()}`);
  }
});
  // --- determine bucket (D9 / D10) from sheet name ---
  const bucket = sheet.getName().split("_")[0];
  const bucketFolder = getOrCreateFolder_(rootFolder, bucket);

  const haMap = {};
  rows.forEach(r => {
    const ha = String(r[COL.ha] || "").trim();
    if (!ha) return;

    if (!haMap[ha]) haMap[ha] = [];
    haMap[ha].push(r);
  });

  Object.keys(haMap).forEach(haName => {
    buildInvoicePdf_(
      ss,
      haName,
      haMap[haName],
      COL,
      periodStart,
      periodEnd,
      bucketFolder
    );
  });
}

/**
 * BUILD ONE AGENCY INVOICE PDF
 */
function buildInvoicePdf_(
  ss,
  haName,
  rows,
  COL,
  periodStart,
  periodEnd,
  folder
) {
  /* ===== SORT ROWS (Patient → Date) ===== */
  const sortedRows = rows.slice().sort((a, b) => {
    const pA = String(a[COL.patient] || "").toUpperCase();
    const pB = String(b[COL.patient] || "").toUpperCase();
    if (pA < pB) return -1;
    if (pA > pB) return 1;

    const dA = a[COL.date] instanceof Date
      ? a[COL.date].getTime()
      : new Date(a[COL.date]).getTime();
    const dB = b[COL.date] instanceof Date
      ? b[COL.date].getTime()
      : new Date(b[COL.date]).getTime();
    return dA - dB;
  });

  const totalVisits = sortedRows.length;
  const totalAmount = sortedRows.reduce(
    (s, r) => s + Number(r[COL.rate] || 0), 0
  );

  const safeName = haName.replace(/[^\w\s-]/g, "").substring(0, 90);
  const tmp = ss.insertSheet(`__TMP_INVOICE_${safeName}_${Date.now()}`);
  tmp.clear();

  /* ===== HEADER ===== */
  tmp.getRange("A1").setValue(haName);
  tmp.getRange("A2").setValue(`Invoice Period: ${periodStart} – ${periodEnd}`);

  tmp.getRange("F1").setValue(`Total Visits: ${totalVisits}`);
  tmp.getRange("F2").setValue(`Total Amount: ${formatCurrency_(totalAmount)}`);

  tmp.getRange("A1:F2")
    .setFontWeight("bold")
    .setFontSize(14);

  tmp.getRange("F1:F2").setHorizontalAlignment("right");

  /* ===== TABLE HEADER ===== */
const tableHeaders = [
  "Patient Name",
  "Visit Type",
  "Visit Date",
  "First Name",
  "Last Name",
  "Rate"
];


  tmp.getRange("A4:F4")
    .setValues([tableHeaders])
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

const tableData = sortedRows.map(r => [
  r[COL.patient],
  r[COL.visit],
  r[COL.date],
  r[COL.first],
  r[COL.last],
  r[COL.rate]
]);


  tmp.getRange(5, 1, tableData.length, tableHeaders.length)
    .setValues(tableData);

  tmp.getRange("C5:C")
    .setNumberFormat("MM/dd/yyyy")
    .setHorizontalAlignment("center");

  tmp.getRange("F5:F")
    .setNumberFormat("$#,##0.00")
    .setHorizontalAlignment("center");
    // Text columns — left aligned for readability
tmp.getRange("A5:A").setHorizontalAlignment("left"); // Patient Name
tmp.getRange("B5:B").setHorizontalAlignment("left"); // Visit Type
tmp.getRange("D5:D").setHorizontalAlignment("left"); // First Name
tmp.getRange("E5:E").setHorizontalAlignment("left"); // Last Name


  // Column widths
tmp.setColumnWidth(1, 280); // Patient Name
tmp.setColumnWidth(2, 260); // Visit Type
tmp.setColumnWidth(3, 140); // Visit Date
tmp.setColumnWidth(4, 160); // First Name
tmp.setColumnWidth(5, 160); // Last Name
tmp.setColumnWidth(6, 120); // Rate


  /* ===== EXPORT (Phase-4 timing) ===== */
  SpreadsheetApp.flush();
  const invoiceFolder = getOrCreateFolder_(folder, "Invoices");


  try {
    exportSheetToPdf_(tmp, invoiceFolder, safeName, true); // invoices = gridlines ON
    Utilities.sleep(EXPORT_SLEEP_MS);
  } catch (err) {
    if (String(err).includes("429")) {
      Utilities.sleep(BACKOFF_MS);
      throw err;
    }
    throw err;
  } finally {
    ss.deleteSheet(tmp);
  }
}

/**
 * EXPORT TEMP SHEET TO PDF
 */
function exportSheetToPdf_(sheet, folder, filename, showGridlines) {
  // --- OPTION B: resume-only ---
// If PDF already exists, skip creation
const existing = folder.getFilesByName(`${filename}.pdf`);
if (existing.hasNext()) {
  return;
}

  
  // --- export fresh PDF ---
  const ss = SpreadsheetApp.getActive();
  const url =
    `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?` +
    `format=pdf&gid=${sheet.getSheetId()}` +
    `&portrait=false&fitw=true` +
    `&sheetnames=false&pagenumbers=false` +
    `&gridlines=${showGridlines ? "true" : "false"}`;

  const blob = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` }
  }).getBlob().setName(`${filename}.pdf`);

  folder.createFile(blob);
}



/**
 * HELPERS (shared with Phase 6)
 */
function getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}



