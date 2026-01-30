function publishFinalWorkbookCopy() {
  const ss = SpreadsheetApp.getActive();
  const admin = ss.getSheetByName("Admin_Config");
  if (!admin) throw new Error("Admin_Config tab not found");

  // ---- Pay period ----
  const dates = getAdminDates_();
  const tz = Session.getScriptTimeZone();
  const periodStart = Utilities.formatDate(dates.start, tz, "MM/dd/yyyy");
  const periodEnd   = Utilities.formatDate(dates.end, tz, "MM/dd/yyyy");

  // ---- Folder resolution ----
  const rootCell = admin.createTextFinder("Payroll Reports Drive Folder ID").findNext();
  if (!rootCell) throw new Error("Missing Payroll Reports Drive Folder ID");

  const parent = DriveApp.getFolderById(
    admin.getRange(rootCell.getRow(), 2).getValue()
  );

  const rootNameCell = admin.createTextFinder("Output PDFs Root Folder Name").findNext();
  if (!rootNameCell) throw new Error("Missing Output PDFs Root Folder Name");

  const root = getOrCreateFolder_(
    parent,
    admin.getRange(rootNameCell.getRow(), 2).getDisplayValue().trim()
  );

  const periodFolder = getOrCreateFolder_(root, `${periodStart} - ${periodEnd}`);

  // ---- Final file name ----
  const finalName = `${periodStart} - ${periodEnd} FINAL (Snapshot)`;

  // ---- Overwrite if exists ----
  const existing = periodFolder.getFilesByName(finalName);
  if (existing.hasNext()) {
    existing.next().setTrashed(true);
  }

  // ---- Copy entire workbook ----
  const copy = DriveApp.getFileById(ss.getId()).makeCopy(finalName);

  // ---- Move into period folder ----
    periodFolder.addFile(copy);
  DriveApp.getRootFolder().removeFile(copy);
}  // <-- ADD THIS

/************ WRITE + FORMAT ************/


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
    PAY: idx("Price agreed between HA & Clinician"),
    RATE: idx("HA Initial price")
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