/******************************************************
 * AUDIT CHECKS — RAW vs PAYROLL vs INVOICE
 * Standalone, read-only
 ******************************************************/

function buildAuditTab() {
  const ss = SpreadsheetApp.getActive();

  const raw = ss.getSheetByName("Raw_All_Visits_DERIVED");
  if (!raw) throw new Error("Raw_All_Visits not found");

  const auditName = "Audit_Checks";
  let audit = ss.getSheetByName(auditName);
  if (!audit) audit = ss.insertSheet(auditName);
  audit.clear();
  // HARD reset sheet size to prevent phantom ranges
const maxRows = audit.getMaxRows();
if (maxRows > 10) {
  audit.deleteRows(2, maxRows - 10);
}

  audit.setFrozenRows(1);

  const rawData = raw.getDataRange().getValues().slice(1);

  const rawD9 = rawData.filter(r => isBucket_(r, "D9")).length;
  const rawD10 = rawData.filter(r => isBucket_(r, "D10")).length;
  const rawTotal = rawD9 + rawD10;

  const d9Payroll = countDataRows_("D9_Payroll");
  const d10Payroll = countDataRows_("D10_Payroll");
  const payrollTotal = d9Payroll + d10Payroll;

  const d9Invoice = countDataRows_("D9_Invoice");
  const d10Invoice = countDataRows_("D10_Invoice");
  const invoiceTotal = d9Invoice + d10Invoice;

  // RAW dollar totals
const rawPayrollTotal =
  sumRaw_("Price agreed between HA & Clinician", "D9") +
  sumRaw_("Price agreed between HA & Clinician", "D10");

const rawInvoiceTotal =
  sumRaw_("HA Initial price", "D9") +
  sumRaw_("HA Initial price", "D10");

// PAYROLL dollar totals
const d9PayrollAmt = sumSheetColumn_("D9_Payroll", "F");
const d10PayrollAmt = sumSheetColumn_("D10_Payroll", "F");
const payrollAmtTotal = d9PayrollAmt + d10PayrollAmt;

// INVOICE dollar totals
const d9InvoiceAmt = sumSheetColumn_("D9_Invoice", "F");
const d10InvoiceAmt = sumSheetColumn_("D10_Invoice", "F");
const invoiceAmtTotal = d9InvoiceAmt + d10InvoiceAmt;


const rows = [
  // ROW COUNT AUDIT — PAYROLL
  ["Payroll", "Raw", "D9 Payroll", rawD9, d9Payroll, match_(rawD9, d9Payroll)],
  ["Payroll", "Raw", "D10 Payroll", rawD10, d10Payroll, match_(rawD10, d10Payroll)],
  ["Payroll", "Raw", "D9 + D10 Payroll", rawTotal, payrollTotal, match_(rawTotal, payrollTotal)],

  // ROW COUNT AUDIT — INVOICE
  ["Invoice", "Raw", "D9 Invoice", rawD9, d9Invoice, match_(rawD9, d9Invoice)],
  ["Invoice", "Raw", "D10 Invoice", rawD10, d10Invoice, match_(rawD10, d10Invoice)],
  ["Invoice", "Raw", "D9 + D10 Invoice", rawTotal, invoiceTotal, match_(rawTotal, invoiceTotal)],

  // DOLLAR AUDIT — PAYROLL
  ["Payroll $", "Raw", "D9 + D10 Payroll", rawPayrollTotal, payrollAmtTotal, match_(rawPayrollTotal, payrollAmtTotal)],

  // DOLLAR AUDIT — INVOICE
  ["Invoice $", "Raw", "D9 + D10 Invoice", rawInvoiceTotal, invoiceAmtTotal, match_(rawInvoiceTotal, invoiceAmtTotal)]
];


// Write header
audit.getRange(1, 1, 1, 6).setValues([[
  "Check Type",
  "Source",
  "Target",
  "Raw Count",
  "Target Count",
  "Match"
]]);

// Write data — EXACT size only
if (rows.length) {
  audit.getRange(2, 1, rows.length, 6).setValues(rows);
}


 // Format Match column (ONLY data rows)
audit.getRange(2, 6, rows.length, 1).setFontWeight("bold");
audit.getRange(2, 6, rows.length, 1).setBackgrounds(
  audit.getRange(2, 6, rows.length, 1)
    .getValues()
    .map(r => [r[0] === "YES" ? "#c8e6c9" : "#ffcdd2"])
);

  audit.autoResizeColumns(1, 6);

  // Force currency formatting for $ audit rows
rows.forEach((r, i) => {
  if (String(r[0]).includes("$")) {
    audit.getRange(i + 2, 4, 1, 2)
      .setNumberFormat("$#,##0.00");
  }
});
}


function countDataRows_(sheetName) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return 0;

  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return 0;

  const data = values.slice(1); // drop header row

  return data.filter(r => {
    // skip fully blank rows
    if (r.join("").trim() === "") return false;

    // skip TOTAL/TOTALS rows regardless of which column they appear in
    const rowText = r.map(v => String(v || "").trim().toUpperCase()).join(" ");
    if (rowText.includes("TOTAL")) return false; // catches TOTAL and TOTALS

    return true;
  }).length;
}
/************ BUCKET HELPERS ************/

function isBucket_(row, bucket) {
  const haIdx = getHAIndex_();
  const ha = String(row[haIdx] || "").trim().toUpperCase();
  return ha.startsWith(bucket + " ");
}

let _haIndexCache = null;

function getHAIndex_() {
  if (_haIndexCache !== null) return _haIndexCache;

  const raw = SpreadsheetApp
    .getActive()
    .getSheetByName("Raw_All_Visits_DERIVED");

  if (!raw) throw new Error("Raw_All_Visits not found");

  const headers = raw
    .getRange(1, 1, 1, raw.getLastColumn())
    .getValues()[0];

  _haIndexCache = headers.findIndex(h =>
    String(h || "").trim().toLowerCase() === "ha name"
  );

  if (_haIndexCache === -1) {
    throw new Error("HA Name column not found in Raw_All_Visits");
  }

  return _haIndexCache;
}
function match_(a, b) {
  const ra = Math.round((Number(a) || 0) * 100);
  const rb = Math.round((Number(b) || 0) * 100);
  return ra === rb ? "YES" : "NO";
}

function sumRaw_(colName, bucket) {
  const ss = SpreadsheetApp.getActive();
  const raw = ss.getSheetByName("Raw_All_Visits_DERIVED");
  if (!raw) throw new Error("Raw_All_Visits not found");

  const data = raw.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const colIdx = headers.findIndex(h =>
    String(h || "").trim().toLowerCase() === colName.toLowerCase()
  );
  if (colIdx === -1) throw new Error(`Column not found: ${colName}`);

  return rows.reduce((sum, r) => {
    if (!isBucket_(r, bucket)) return sum;
    const v = Number(String(r[colIdx] || "").replace(/[$,]/g, ""));
    return sum + (isNaN(v) ? 0 : v);
  }, 0);
}

function sumSheetColumn_(sheetName, colLetter) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return 0;

  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return 0;

  const headers = data[0];
  const colIdx = colLetter.charCodeAt(0) - 65; // A=0, B=1, etc.
  const rows = data.slice(1); // drop header

  return rows.reduce((sum, r) => {
    // skip TOTAL / TOTALS rows anywhere
    const rowText = r.map(v => String(v || "").toUpperCase()).join(" ");
    if (rowText.includes("TOTAL")) return sum;

    const n = Number(String(r[colIdx] || "").replace(/[$,]/g, ""));
    return sum + (isNaN(n) ? 0 : n);
  }, 0);
}


