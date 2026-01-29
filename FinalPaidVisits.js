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
}
function onOpen_FinalPaidVisits_() {
  SpreadsheetApp.getUi()
    .createMenu("Payroll")
    .addItem("Publish Final Snapshot", "publishFinalWorkbookCopy")
    .addToUi();
}