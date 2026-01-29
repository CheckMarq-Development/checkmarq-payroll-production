/************************************************************
 * BUTTON HANDLER
 * Recalculate Payroll & Invoices
 *
 * - Reads Raw_All_Visits only
 * - Rebuilds D9 + D10 Payroll
 * - Rebuilds D9 + D10 Invoices
 * - Does NOT import or overwrite Raw
 * - Does NOT touch PDFs
 * - Audit tab remains read-only
 ************************************************************/
function recalculatePayrollAndInvoices() {
  const ss = SpreadsheetApp.getActive();

  // ---- Hard safety check ----
  const raw = ss.getSheetByName("Raw_All_Visits");
  if (!raw) {
    SpreadsheetApp.getUi().alert("ERROR: Raw_All_Visits sheet not found.");
    return;
  }

  // ---- Required builders must already exist ----
  const requiredFns = [
    "buildPayrollAndSummary", // D9 + D10 Payroll
    "buildInvoices"          // D9 + D10 Invoices
  ];

  for (const fn of requiredFns) {
    if (typeof this[fn] !== "function") {
      SpreadsheetApp.getUi().alert(
        `ERROR: Required function missing:\n\n${fn}`
      );
      return;
    }
  }

  try {
    // ORDER IS INTENTIONAL. DO NOT SWAP.
    buildPayrollAndSummary(); // Payroll first
    buildInvoices();          // Invoices second

    SpreadsheetApp.getUi().alert(
      "Payroll and Invoices recalculated successfully."
    );

  } catch (err) {
    SpreadsheetApp.getUi().alert(
      "Recalculation failed:\n\n" + err.message
    );
    throw err; // surfaces stack trace in executions log
  }
}
function cleanupTempSheets_() {
  const ss = SpreadsheetApp.getActive();
  ss.getSheets()
    .filter(s => s.getName().startsWith("__TMP_"))
    .forEach(s => ss.deleteSheet(s));
}
