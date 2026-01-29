function resetForNewPayPeriod() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "RESET FOR NEW PAY PERIOD",
    "This will permanently clear:\n\n" +
    "• Raw visit data\n" +
    "• Payroll & invoice calculations\n" +
    "• Audit logs\n" +
    "• Temporary tabs\n" +
    "• Pay period and approval dates\n\n" +
    "This CANNOT be undone.\n\nProceed?",
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActive();

  // ---- Clear RAW data ----
  const raw = ss.getSheetByName("Raw_All_Visits");
  if (raw && raw.getLastRow() > 1) {
    raw.getRange(2, 1, raw.getLastRow() - 1, raw.getLastColumn()).clearContent();
  }

  // ---- Clear calculated sheets ----
  ss.getSheets().forEach(s => {
    const name = s.getName();

    if (
      name.endsWith("_Payroll") ||
      name.endsWith("_Invoice") ||
      name.includes("Summary")
    ) {
      if (s.getLastRow() > 1) {
        s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).clearContent();
      }
    }

    // ---- Delete temp tabs ----
    if (name.startsWith("__TMP_")) {
      ss.deleteSheet(s);
    }
  });

// ---- Clear audit logs & derived data ----
[
  "Audit_Checks",
  "Raw_All_Visits_DERIVED",
  "Payroll_PDF_Audit",
  "Invoice_PDF_Audit",
  "Payroll_PDF_Errors"
].forEach(name => {
  const sheet = ss.getSheetByName(name);
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }
});



  // ---- Clear Admin_Config date fields ----
  const admin = ss.getSheetByName("Admin_Config");
  [
    "Approved From",
    "Approved To",
    "Pay Period Start",
    "Pay Period End"
  ].forEach(label => {
    const cell = admin.createTextFinder(label).findNext();
    if (cell) admin.getRange(cell.getRow(), 2).clearContent();
  });

  ui.alert("Reset complete. Enter new pay period details before importing data.");
}
