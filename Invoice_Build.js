/******************************************************
 * INVOICE BUILD + HA SUMMARY — D9 + D10
 * Standalone, read-only from Raw_All_Visits
 ******************************************************/

function buildInvoicesAndHASummariesOnly() {
   assertPayPeriodConfigured_();
  const ss = SpreadsheetApp.getActive();
  const raw = ss.getSheetByName("Raw_All_Visits_DERIVED");
if (!raw) throw new Error("Raw_All_Visits_DERIVED not found");

  const data = raw.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const COL = indexColumns_(headers, {
    patient: "Patient name",
    visitType: "Visit type",
    visitDate: "Visit scheduled date",
    first: "Assigned Clinician First Name",
    last: "Assigned Clinician Last Name",
    ha: "HA Name",
    rate: "HA Initial price"
  });

  buildInvoiceAndHASummary_(ss, "D9", rows, COL);
  buildInvoiceAndHASummary_(ss, "D10", rows, COL);

  SpreadsheetApp.getUi().alert("Invoice and HA summary tabs rebuilt.");
}

function buildInvoiceAndHASummary_(ss, bucket, rows, COL) {
  const invoiceName = `${bucket}_Invoice`;
  const summaryName = `${bucket}_Invoice_Summary`;

  let invoice = ss.getSheetByName(invoiceName) || ss.insertSheet(invoiceName);
  let summary = ss.getSheetByName(summaryName) || ss.insertSheet(summaryName);

  invoice.clear();
  summary.clear();

  invoice.setFrozenRows(1);
  summary.setFrozenRows(1);

  const output = rows
    .filter(r => {
      const ha = String(r[COL.ha] || "").trim().toUpperCase();
      return ha.startsWith(bucket + " ");
    })
    .map(r => [
      r[COL.patient],     // 0
      r[COL.visitType],   // 1
      r[COL.visitDate],   // 2
      r[COL.first],       // 3
      r[COL.last],        // 4
      num_(r[COL.rate]),  // 5
      r[COL.ha]           // 6
    ]);

  sortInvoiceRows_(output);

  // ---------- INVOICE TAB ----------
  invoice.getRange(1, 1, 1, 7).setValues([[
    "Patient Name",
    "Visit Type",
    "Visit Date",
    "Clinician First Name",
    "Clinician Last Name",
    "Rate",
    "HA Name"
  ]]);

  if (output.length) {
    invoice.getRange(2, 1, output.length, 7).setValues(output);
    invoice.getRange("C:C").setNumberFormat("m/d/yyyy");
    invoice.getRange("F:F").setNumberFormat("$#,##0.00");

    const totalRow = output.length + 2;
    invoice.getRange(totalRow, 5).setValue("TOTAL");
    invoice.getRange(totalRow, 6)
      .setFormula(`=SUBTOTAL(9,F2:F${totalRow - 1})`)
      .setFontWeight("bold");
  }

  if (invoice.getFilter()) invoice.getFilter().remove();
  invoice.getRange(1, 1, invoice.getLastRow(), 7).createFilter();

  // ---------- HA SUMMARY ----------
  const map = {};

  output.forEach(r => {
    const rate = r[5];
    if (rate === 0) return;

    const ha = r[6];
    if (!map[ha]) {
      map[ha] = { visits: 0, total: 0 };
    }

    map[ha].visits += 1;
    map[ha].total += rate;
  });

  const summaryRows = Object.entries(map).map(([ha, o]) => [
    ha,
    o.visits,
    o.total
  ]);

  summary.getRange(1, 1, 1, 3).setValues([[
    "HA Name",
    "Total Visits",
    "Invoice Total"
  ]]);

  if (summaryRows.length) {
    summary.getRange(2, 1, summaryRows.length, 3).setValues(summaryRows);
    summary.getRange("C:C").setNumberFormat("$#,##0.00");

    const end = summaryRows.length + 2;

    summary.getRange(end, 1).setValue("TOTAL");
    summary.getRange(end, 3)
      .setFormula(`=SUM(C2:C${end - 1})`)
      .setFontWeight("bold");
  }
}

/************ SORT ************/
function sortInvoiceRows_(rows) {
  rows
    .map((r, i) => ({ r, i }))
    .sort((a, b) => {
      const p = normKey_(a.r[0]).localeCompare(normKey_(b.r[0]));
if (p) return p;

const da = a.r[2] instanceof Date ? a.r[2].getTime() : 0;
const db = b.r[2] instanceof Date ? b.r[2].getTime() : 0;
if (da !== db) return da - db;


      const l = normKey_(a.r[4]).localeCompare(normKey_(b.r[4]));
      if (l) return l;

      const f = normKey_(a.r[3]).localeCompare(normKey_(b.r[3]));
      if (f) return f;

      const v = normKey_(a.r[1]).localeCompare(normKey_(b.r[1]));
      if (v) return v;

      return a.i - b.i;
    })
    .forEach((o, i) => (rows[i] = o.r));
}


// ---- Compatibility wrapper ----
// Called by Recalculate Payroll & Invoices controller
function buildInvoices() {
  buildInvoicesAndHASummariesOnly();
}
function buildD9AllAboutYouSpecialInvoice() {
  const ss = SpreadsheetApp.getActive();

  // ---- Source (DERIVED only) ----
  const raw = ss.getSheetByName("Raw_All_Visits_DERIVED");
  if (!raw) throw new Error("Raw_All_Visits_DERIVED not found");

  const data = raw.getDataRange().getValues();
  if (data.length < 2) return;

  const headers = data[0];
  const rows = data.slice(1);

  const COL = indexColumns_(headers, {
    patient: "Patient name",
    visitType: "Visit type",
    visitDate: "Visit scheduled date",
    first: "Assigned Clinician First Name",
    last: "Assigned Clinician Last Name",
    ha: "HA Name",
    pay: "Price agreed between HA & Clinician"
  });

  // ---- Filter D9 All About You only ----
  const output = rows
    .filter(r =>
      String(r[COL.ha] || "")
        .toUpperCase()
        .startsWith("D9 ALL ABOUT YOU")
    )
    .map(r => {
      const pay = Number(r[COL.pay]) || 0;
      const rate = Math.min(pay * 1.2, 89);

      return [
        r[COL.patient],   // Patient Name
        r[COL.visitType], // Visit Type
        r[COL.visitDate], // Visit Date
        r[COL.first],     // Clinician First Name
        r[COL.last],      // Clinician Last Name
        rate,             // Rate (special calc)
        r[COL.ha]         // HA Name
      ];
    });

  if (!output.length) {
    SpreadsheetApp.getUi().alert("No D9 All About You visits found.");
    return;
  }

  sortInvoiceRows_(output);

  // ---- Build invoice tab ----
  const invoiceName = "D9_All_About_You_Special_Invoice";
  let invoice = ss.getSheetByName(invoiceName);
  if (!invoice) invoice = ss.insertSheet(invoiceName);
  invoice.clear();
  invoice.setFrozenRows(1);

  invoice.getRange(1, 1, 1, 7).setValues([[
    "Patient Name",
    "Visit Type",
    "Visit Date",
    "Clinician First Name",
    "Clinician Last Name",
    "Rate",
    "HA Name"
  ]]);

  invoice.getRange(2, 1, output.length, 7).setValues(output);
  invoice.getRange("C:C").setNumberFormat("m/d/yyyy");
  invoice.getRange("F:F").setNumberFormat("$#,##0.00");

  const totalRow = output.length + 2;
  invoice.getRange(totalRow, 5).setValue("TOTAL");
  invoice.getRange(totalRow, 6)
    .setFormula(`=SUBTOTAL(9,F2:F${totalRow - 1})`)
    .setFontWeight("bold");

  if (invoice.getFilter()) invoice.getFilter().remove();
  invoice.getRange(1, 1, invoice.getLastRow(), 7).createFilter();

  // ---- PDF export (reuse existing invoice PDF logic) ----
  exportSpecialInvoicePdf_(invoice, output);

  // ---- Draft email (no audit) ----
  draftSpecialInvoiceEmail_(invoice);
}
function exportSpecialInvoicePdf_(invoiceSheet, outputRows) {
  const ss = SpreadsheetApp.getActive();
  const admin = ss.getSheetByName("Admin_Config");
  if (!admin) throw new Error("Admin_Config tab not found");

  const dates = getAdminDates_();
  const tz = Session.getScriptTimeZone();
  const periodStart = Utilities.formatDate(dates.start, tz, "MM/dd/yyyy");
  const periodEnd   = Utilities.formatDate(dates.end, tz, "MM/dd/yyyy");

  // --- Resolve Drive folders the SAME way your Payroll PDF script does (no .next() crashes) ---
  const rootCell = admin.createTextFinder("Payroll Reports Drive Folder ID").findNext();
  if (!rootCell) throw new Error("Admin_Config missing 'Payroll Reports Drive Folder ID'");
  const rootId = admin.getRange(rootCell.getRow(), 2).getValue();
  if (!rootId) throw new Error("Payroll Reports Drive Folder ID value is blank");
  const parent = DriveApp.getFolderById(rootId);

  const rootNameCell = admin.createTextFinder("Output PDFs Root Folder Name").findNext();
  if (!rootNameCell) throw new Error("Admin_Config missing 'Output PDFs Root Folder Name'");
  const rootFolderName = admin.getRange(rootNameCell.getRow(), 2).getDisplayValue().trim();
  if (!rootFolderName) throw new Error("Output PDFs Root Folder Name value is blank");

  const root = getOrCreateFolder_(parent, rootFolderName);
  const periodFolder = getOrCreateFolder_(root, `${periodStart} - ${periodEnd}`);

  // Put this PDF into: Period / D10 / Invoice  (per your spec)
  const d10Folder = getOrCreateFolder_(periodFolder, "D10");
  const invoiceFolder = getOrCreateFolder_(d10Folder, "Invoice");

  // --- Build a nicely formatted TEMP sheet like your Payroll PDFs ---
  const filename = "D9_All_About_You_Special_Invoice";
  const tmp = ss.insertSheet(`__TMP_INV_${filename}_${Date.now()}`);
  tmp.clear();

  const totalVisits = (outputRows && outputRows.length) ? outputRows.length : 0;
  const totalAmt = (outputRows || []).reduce((s, r) => s + Number(r[5] || 0), 0); // Rate is col 5 in output rows

  // HEADER (matches your payroll styling pattern)
  tmp.getRange("A1").setValue("D9 All About You — Special Invoice");
  tmp.getRange("A2").setValue(`Pay Period: ${periodStart} – ${periodEnd}`);

  tmp.getRange("E1").setValue(`Total Visits: ${totalVisits}`);
tmp.getRange("E2").setValue(`Total Invoice: ${formatCurrency_(totalAmt)}`);

tmp.getRange("A1:E2")
  .setFontWeight("bold")
  .setFontSize(14);

tmp.getRange("E1:E2").setHorizontalAlignment("right");


  // TABLE HEADER
const headers = [
  "Patient Name",
  "Visit Type",
  "Visit Date",
  "Clinician",
  "Rate"
];

tmp.getRange("A4:E4").setValues([headers]);

tmp.getRange("A4:E4")
  .setFontWeight("bold")
  .setFontSize(12)
  .setHorizontalAlignment("center");

tmp.getRange("A5:E")
  .setFontSize(11);

tmp.getRange("C5:C")
  .setNumberFormat("MM/dd/yyyy")
  .setHorizontalAlignment("center");

tmp.getRange("E5:E")
  .setNumberFormat("$#,##0.00")
  .setHorizontalAlignment("center");

  // TABLE DATA
  if (outputRows && outputRows.length) {
  const formattedRows = outputRows.map(r => [
    r[0],                 // Patient
    r[1],                 // Visit Type
    r[2],                 // Visit Date
    `${r[3]} ${r[4]}`,    // Clinician (First + Last)
    r[5]                  // Rate
  ]);

  tmp.getRange(5, 1, formattedRows.length, 5).setValues(formattedRows);
}


  // FORMATS (same idea as payroll)
  tmp.getRange("C5:C").setNumberFormat("MM/dd/yyyy").setHorizontalAlignment("center");
  tmp.getRange("F5:F").setNumberFormat("$#,##0.00").setHorizontalAlignment("center");

  // COLUMN WIDTHS (tuned for readability)
tmp.setColumnWidth(1, 280); // Patient Name
tmp.setColumnWidth(2, 240); // Visit Type
tmp.setColumnWidth(3, 140); // Visit Date
tmp.setColumnWidth(4, 260); // Clinician
tmp.setColumnWidth(5, 120); // Rate

// Remove unused columns to increase scale/readability
const maxCols = tmp.getMaxColumns();
if (maxCols > 5) {
  tmp.deleteColumns(6, maxCols - 5);
}

  SpreadsheetApp.flush();

  // EXPORT using your payroll-style exporter (landscape, fit-to-width, no gridlines)
  try {
    exportPayrollSheetToPdf_(tmp, invoiceFolder, filename, true);
  } finally {
    ss.deleteSheet(tmp); // always clean up
  }
}

