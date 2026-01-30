/************ CONSTANTS ************/


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



