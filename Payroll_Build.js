/******************************************************
 * PAYROLL BUILD — D9 + D10
 * NO assumptions beyond approved spec
 ******************************************************/

function buildDerivedRawVisits_() {
  const ss = SpreadsheetApp.getActive();

  const raw = ss.getSheetByName("Raw_All_Visits");
  if (!raw) throw new Error("Raw_All_Visits_DERIVED not found");

  const derivedName = "Raw_All_Visits_DERIVED";
  let derived = ss.getSheetByName(derivedName);
  if (!derived) derived = ss.insertSheet(derivedName);

  derived.clear();

  const data = raw.getDataRange().getValues();
  if (data.length < 2) return;

  const headers = data[0];
  const rows = data.slice(1);

  const COL = indexColumns_(headers, {
    ha: "HA Name",
    pay: "Price agreed between HA & Clinician",
    rate: "HA Initial price"
  });

const output = rows.map(r => {
  const row = [...r]; // clone row defensively
  const haName = String(row[COL.ha] || "").toUpperCase();
  const pay    = Number(row[COL.pay]) || 0;

  if (haName.startsWith("D9 ALL ABOUT YOU")) {
    row[COL.rate] = pay;
  }

  return row;
});


  derived.getRange(1, 1, 1, headers.length).setValues([headers]);
  derived.getRange(2, 1, output.length, headers.length).setValues(output);
}

function buildPayrollAndSummary() {
  const ss = SpreadsheetApp.getActive();
  buildDerivedRawVisits_(); // ← ADD THIS LINE FIRST
const raw = ss.getSheetByName("Raw_All_Visits_DERIVED");
  if (!raw) throw new Error("Raw_All_Visits_DERIVED not found");

  const data = raw.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const COL = indexColumns_(headers, {
    first: "Assigned Clinician First Name",
    last: "Assigned Clinician Last Name",
    patient: "Patient name",
    visitType: "Visit type",
    visitDate: "Visit scheduled date",
    pay: "Price agreed between HA & Clinician",
    ha: "HA Name",
    rate: "HA Initial price"
  });

  buildPayroll_(ss, "D9", rows, COL);
  buildPayroll_(ss, "D10", rows, COL);
}

function buildPayroll_(ss, haName, rows, COL) {
  const payrollName = `${haName}_Payroll`;
  const summaryName = `${haName}_Payroll_Summary`;

  const payroll =
    ss.getSheetByName(payrollName) ||
    ss.insertSheet(payrollName);

  payroll.clear();
  payroll.setFrozenRows(1);

  // BUILD OUTPUT (NO SORT YET)
  let output = rows
    .filter(r => {
  const ha = String(r[COL.ha] || "").trim().toUpperCase();
  return ha.startsWith(haName.toUpperCase() + " ");
})
    .map(r => {
      const dt = r[COL.visitDate] instanceof Date ? r[COL.visitDate] : null;
      return [
        r[COL.first],
        r[COL.last],
        r[COL.patient],
        r[COL.visitType],
        dt,
        r[COL.pay],
        r[COL.ha],
        r[COL.rate]
      ];
    });

  output
  .map((r, i) => ({ r, i }))
  .sort((a, b) => {
    // 1. Visit Date
    const da = a.r[4] instanceof Date ? a.r[4].getTime() : 0;
    const db = b.r[4] instanceof Date ? b.r[4].getTime() : 0;
    if (da !== db) return da - db;

    // 2. Patient Name
    const p = normKey_(a.r[2]).localeCompare(normKey_(b.r[2]));
    if (p) return p;

    // 3. Clinician Last Name
    const l = normKey_(a.r[1]).localeCompare(normKey_(b.r[1]));
    if (l) return l;

    // 4. Clinician First Name
    const f = normKey_(a.r[0]).localeCompare(normKey_(b.r[0]));
    if (f) return f;

    // 5. Visit Type (FINAL tie-breaker — fixes D10)
    const v = normKey_(a.r[3]).localeCompare(normKey_(b.r[3]));
    if (v) return v;

    return a.i - b.i; // stability
  })
  .forEach((o, i) => (output[i] = o.r));


  payroll.getRange(1, 1, 1, 8).setValues([[
    "First Name",
    "Last Name",
    "Patient Name",
    "Visit Type",
    "Visit Date",
    "Pay",
    "HA Name",
    "Rate"
  ]]);

  if (output.length) {
    payroll.getRange(2, 1, output.length, 8).setValues(output);
  }

  payroll.getRange("E:E").setNumberFormat("m/d/yyyy");
  payroll.getRange("F:F").setNumberFormat("$#,##0.00");
  payroll.getRange("H:H").setNumberFormat("$#,##0.00");

  addTotals_(payroll);
  highlightDuplicates_(payroll);

  const existingFilter = payroll.getFilter();
  if (existingFilter) existingFilter.remove();
  payroll.getRange(1, 1, payroll.getLastRow(), 8).createFilter();

  buildSummary_(ss, payroll, summaryName);
}

function addTotals_(sheet) {
  const lastDataRow = sheet.getLastRow();
  const totalRow = lastDataRow + 1;

  sheet.getRange(totalRow, 5).setValue("TOTALS");
sheet.getRange(totalRow, 6)
  .setFormula(`=SUM(F2:F${lastDataRow})`);
sheet.getRange(totalRow, 8)
  .setFormula(`=SUM(H2:H${lastDataRow})`);

}

function highlightDuplicates_(sheet) {
  if (sheet.getLastRow() < 3) return;

  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8);
  const values = range.getValues();
  const seen = {};

  values.forEach((row, i) => {
    const key = row.join("||");
    if (seen[key]) {
      sheet.getRange(i + 2, 1, 1, 8).setBackground("#fff59d");
    } else {
      seen[key] = true;
    }
  });
}
function isW2Clinician_(first, last) {
  const key = `${first} ${last}`.toLowerCase().replace(/\s+/g, " ").trim();
  return [
    "grayson lambert",
    "adriana murrill",
    "emily steinman",
    "cameron lombardi",
    "glenda kiziltan",
    "sheila wilson"
  ].includes(key);
}

function buildSummary_(ss, payroll, summaryName) {
  const summary =
    ss.getSheetByName(summaryName) ||
    ss.insertSheet(summaryName);

  summary.clear();

  const payrollData = payroll.getDataRange().getValues();
  const totalRowIndex = payrollData.findIndex(r => r[4] === "TOTALS");
  const data = payrollData.slice(1, totalRowIndex);

  const map = {};

  data.forEach(r => {
    const key = `${r[0]}||${r[1]}`;
    if (!map[key]) {
      map[key] = { first: r[0], last: r[1], visits: 0, pay: 0 };
    }

    const pay = Number(r[5]) || 0;
    if (pay !== 0) map[key].visits += 1;
    map[key].pay += pay;
  });

const rows = Object.values(map).map(o => {
  const isW2 = isW2Clinician_(o.first, o.last);

  const pay1099 = (!isW2 && o.pay !== 0) ? o.pay : "";
  const payW2   = (isW2 && o.pay !== 0) ? o.pay : "";

  return [
    o.first,
    o.last,
    o.visits,
    pay1099, // 1099 (blank if 0)
    payW2,   // W2   (blank if 0)
    o.pay    // Total Pay (always numeric)
  ];
});


summary.getRange(1, 1, 1, 6).setValues([[
  "First Name",
  "Last Name",
  "Total Visits",
  "1099",
  "W2",
  "Total Pay"
]]);


  if (rows.length) {
    summary.getRange(2, 1, rows.length, 6).setValues(rows);
  }

  summary.getRange("D:F").setNumberFormat("$#,##0.00");

  const end = rows.length + 2;

  summary.getRange(end, 2).setValue("TOTALS");
summary.getRange(end, 3).setFormula(`=SUM(C2:C${end - 1})`);
summary.getRange(end, 4).setFormula(`=SUM(D2:D${end - 1})`); // 1099
summary.getRange(end, 5).setFormula(`=SUM(E2:E${end - 1})`); // W2
summary.getRange(end, 6).setFormula(`=SUM(F2:F${end - 1})`); // Total


  summary.getRange(end + 1, 2).setValue("MATCH VISITS?");
  summary.getRange(end + 1, 3)
    .setFormula(`=IF(C${end}=COUNTIF(${payroll.getName()}!F2:F${totalRowIndex},"<>0"),"YES","NO")`);

  summary.getRange(end + 2, 2).setValue("MATCH PAY?");
  summary.getRange(end + 2, 4)
     .setFormula(`=IF(F${end}=${payroll.getName()}!F${totalRowIndex + 1},"YES","NO")`);
}

function coerceToDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;

  const s = String(v == null ? "" : v).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;

  const mm = Number(m[1]), dd = Number(m[2]), yyyy = Number(m[3]);
  const d = new Date(yyyy, mm - 1, dd);
  return isNaN(d.getTime()) ? null : d;
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

function recalculatePayrollOnly() {
    assertPayPeriodConfigured_();
  const ss = SpreadsheetApp.getActive();
  buildPayrollBucket_(ss, "D9");
  buildPayrollBucket_(ss, "D10");
  SpreadsheetApp.getUi().alert("Payroll tabs rebuilt.");
}

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