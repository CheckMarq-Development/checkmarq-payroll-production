/***************************************************
 * PHASE 8 — EMAIL DRAFT GENERATION (DRAFT ONLY)
 * - Payroll + Invoice
 * - Resume-safe (Option B)
 * - Uses Admin_Config
 * - Writes Email_Audit
 ***************************************************/

/**
 * ENTRY — PAYROLL EMAIL DRAFTS
 */
function buildPayrollEmailDrafts() {
    buildEmailDrafts_({
    type: "PAYROLL",
    pdfSubfolder: "Payroll",
    emailLookup: lookupClinicianEmail_,
    subjectKey: "Payroll Email Subject",
    bodyKey: "Payroll Email Body",
    replyToKey: "Payroll Reply-To",
    ccKey: "Payroll CC",
    bccKey: "Payroll BCC"
  });
}


/**
 * ENTRY — INVOICE EMAIL DRAFTS
 */
function buildInvoiceEmailDrafts() {
    buildEmailDrafts_({
    type: "INVOICE",
    pdfSubfolder: "Invoice",
    emailLookup: lookupAgencyEmails_,
    subjectKey: "Invoice Email Subject",
    bodyKey: "Invoice Email Body",
    replyToKey: "Invoice Reply-To",
    ccKey: "Invoice CC",
    bccKey: "Invoice BCC"
  });
}


/**
 * CORE BUILDER (shared)
 */
function buildEmailDrafts_(cfg) {
  const ss = SpreadsheetApp.getActive();
  const admin = ss.getSheetByName("Admin_Config");
  Logger.log(JSON.stringify(cfg));
  const runId = generateRunId_();


  if (!admin) throw new Error("Admin_Config missing");

const audit = getOrCreateSheet("Email_Audit", [
  "Run ID",
  "Timestamp",
  "Type",
  "Recipients",
  "Subject",
  "PDF Name",
  "PDF File ID",
  "Status",
  "Notes"
]);


  const getCfg = label => {
    const cell = admin.createTextFinder(label).findNext();
    return cell ? admin.getRange(cell.getRow(), 2).getDisplayValue() : "";
  };


  const { start, end } = getAdminDates_();
  const tz = Session.getScriptTimeZone();
  const startStr = Utilities.formatDate(start, tz, "MM/dd/yyyy");
  const endStr = Utilities.formatDate(end, tz, "MM/dd/yyyy");

  const periodFolderName = `${startStr} - ${endStr}`;

  const rootId = getCfg("Payroll Reports Drive Folder ID");
  if (!rootId) throw new Error("Payroll Reports Drive Folder ID missing");

const parent = DriveApp.getFolderById(rootId);

const rootName = getCfg("Output PDFs Root Folder Name");
if (!rootName) throw new Error("Output PDFs Root Folder Name missing");

const rootIter = parent.getFoldersByName(rootName);
if (!rootIter.hasNext()) {
  throw new Error("Output PDFs root folder not found");
}

const root = rootIter.next();



  const periodIter = root.getFoldersByName(periodFolderName);
  if (!periodIter.hasNext()) return;

  const periodFolder = periodIter.next();
  const bucketFolders = periodFolder.getFolders();
  

  while (bucketFolders.hasNext()) {
    const bucketFolder = bucketFolders.next();
    const pdfIter = bucketFolder.getFoldersByName(cfg.pdfSubfolder);
    if (!pdfIter.hasNext()) continue;

    const pdfFolder = pdfIter.next();
    const files = pdfFolder.getFilesByType(MimeType.PDF);

    while (files.hasNext()) {
  const pdf = files.next();

    const baseName = pdf.getName().replace(/\.pdf$/i, "");


      const bucket = bucketFolder.getName(); // D9 or D10
const recipients = cfg
  .emailLookup(baseName, cfg.type === "PAYROLL" ? bucket : null)
  .map(e => String(e).trim())
  .filter(Boolean);


  if (!recipients.length) {
  audit.appendRow([
    runId,
    new Date(),
    cfg.type,
    "",
    "",
    pdf.getName(),
    pdf.getId(),
    "SKIPPED",
    "No email found"
  ]);
  continue;
}


      const subjectTpl = getCfg(cfg.subjectKey);
const htmlTpl = getCfg(cfg.bodyKey);

if (!subjectTpl || !htmlTpl) {
  throw new Error(`${cfg.type}: Email subject/body missing in Admin_Config`);
}

const replyTo = getCfg(cfg.replyToKey);
const cc = splitEmails_(getCfg(cfg.ccKey));
const bcc = splitEmails_(getCfg(cfg.bccKey));

const subject = `[${bucket}] ` + subjectTpl
  .replace("{START}", startStr)
  .replace("{END}", endStr)
  .replace("{AGENCY}", baseName);

if (draftAlreadyExists_(subject, pdf.getName())) {
  audit.appendRow([
    runId,
    new Date(),
    cfg.type,
    recipients.join(", "),
    subject,
    pdf.getName(),
    pdf.getId(),
    "SKIPPED",
    "Draft already exists"
  ]);

  continue;
}

const htmlBody = htmlTpl
  .replace(/{START}/g, startStr)
  .replace(/{END}/g, endStr)
  .replace(/{AGENCY}/g, baseName)
  .replace(/{BUCKET}/g, bucket);

GmailApp.createDraft(
  recipients.join(","),
  subject,
  "", // plain text intentionally empty
  {
    htmlBody,
    attachments: [pdf.getBlob()],
    replyTo: replyTo || undefined,
    cc: cc.length ? cc.join(",") : undefined,
    bcc: bcc.length ? bcc.join(",") : undefined
  }
);


     audit.appendRow([
  runId,
  new Date(),
  cfg.type,
  recipients.join(", "),
  subject,
  pdf.getName(),
  pdf.getId(),
  "DRAFTED",
  ""
]);

    }
  
      }

  // ===== RECONCILIATION SUMMARY =====
  let summary = null;

if (cfg.type === "PAYROLL") {
  summary = reconcileEmailAudit_(runId);
}

if (cfg.type === "INVOICE") {
  summary = reconcileInvoiceEmailAudit_(runId);
}


  if (!summary) return;

  audit.appendRow([""]);
  audit.appendRow(["SUMMARY"]);
audit.appendRow(["Expected", String(summary.expected)]);
audit.appendRow(["Drafted Emails", String(summary.drafted)]);
audit.appendRow(["Attempted Emails", String(summary.attempted)]);
audit.appendRow(["Missing Emails", String(summary.missing.length)]);


  if (summary.missing.length) {
    audit.appendRow([""]);
    audit.appendRow(["MISSING NAME", "BUCKET"]);
    summary.missing.forEach(r => audit.appendRow(r));
  }
}



/**
 * RESUME GUARD — PREVENT DUPLICATE DRAFTS
 */
function draftAlreadyExists_(subject, pdfName) {
  const threads = GmailApp.search(
    `in:drafts subject:"${subject}"`
  );

  for (const t of threads) {
    let messages;
    try {
      messages = t.getMessages();
    } catch (e) {
      continue;
    }

    for (const m of messages) {
      if (!m.isDraft()) continue;

      const atts = m.getAttachments({ includeInlineImages: false });
      if (atts.some(a => a.getName() === pdfName)) {
        return true; // exact match: subject + PDF
      }
    }
  }

  return false;
}

/**
 * HELPERS
 */
function splitEmails_(v) {
  return String(v || "")
    .split(",")
    .map(e => e.trim())
    .filter(Boolean);
}
/***************************************************
 * EMAIL LOOKUPS — AGENCIES (INVOICES)
 ***************************************************/
function lookupAgencyEmails_(pdfBaseName) {
  const target = normalizeName_(pdfBaseName);
  if (!target) return [];

  const emails = new Set();

  CONFIG.AGENCY_DIRECTORY.forEach(src => {
    const ss = SpreadsheetApp.openById(src.spreadsheetId);
    const sheet = ss.getSheetByName(src.sheetName);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    const headers = data.shift().map(h => String(h).trim());

    const nameCol = headers.indexOf(src.agencyNameHeader);
    const emailCol = headers.indexOf(src.emailHeader);

    if (nameCol < 0 || emailCol < 0) return;

    data.forEach(r => {
      if (normalizeName(r[nameCol]) !== target) return;

      String(r[emailCol] || "")
        .split(",")
        .map(e => e.trim())
        .filter(Boolean)
        .forEach(e => emails.add(e));
    });
  });

  return Array.from(emails);
}
/***************************************************
 * EMAIL LOOKUPS — CLINICIANS (PAYROLL)
 ***************************************************/
function lookupClinicianEmail_(pdfBaseName, bucket) {
  const admin = SpreadsheetApp.getActive().getSheetByName("Admin_Config");
  if (!admin) throw new Error("Admin_Config missing");

  const idLabel =
    bucket === "D10"
      ? "D10 Clinician Directory Spreadsheet ID"
      : "D9 Clinician Directory Spreadsheet ID";

  const cell = admin.createTextFinder(idLabel).findNext();
  if (!cell) {
    throw new Error(`Missing ${idLabel} in Admin_Config`);
  }

  const ssId = String(admin.getRange(cell.getRow(), 2).getValue()).trim();
  if (!ssId) {
    throw new Error(`${idLabel} value is blank`);
  }

  const dirSS = SpreadsheetApp.openById(ssId);
  const sheet = dirSS.getSheets()[0]; // first sheet only, by design

  const data = sheet.getDataRange().getValues();
  const headers = data.shift().map(h => String(h).trim().toLowerCase());

  const fnCol =
  headers.indexOf("first name") >= 0
    ? headers.indexOf("first name")
    : headers.indexOf("first");

const lnCol =
  headers.indexOf("last name") >= 0
    ? headers.indexOf("last name")
    : headers.indexOf("last");

const emCol = headers.indexOf("email");

if (fnCol < 0 || lnCol < 0 || emCol < 0) {
  throw new Error("Clinician directory must contain First / Last / Email columns");
}

  const target = normalizeName_(pdfBaseName);
  const emails = new Set();

  data.forEach(r => {
    const full = normalizeName_(`${r[fnCol]} ${r[lnCol]}`);
    if (full === target && r[emCol]) {
      r[emCol]
        .toString()
        .split(",")
        .map(e => e.trim())
        .filter(Boolean)
        .forEach(e => emails.add(e));
    }
  });

  return Array.from(emails);
}

/**
 * Read Pay Period dates from Admin_Config as real Date objects
 * HARD FAILS if values are not dates
 */
function getAdminDates_() {
  const ss = SpreadsheetApp.getActive();
  const admin = ss.getSheetByName("Admin_Config");
  if (!admin) throw new Error("Admin_Config missing");

  const rows = admin.getRange("A1:B100").getValues();
  const map = {};

  rows.forEach(r => {
  if (r[0]) map[String(r[0]).trim()] = r[1];
});


  const start = map["Pay Period Start"];
  const end   = map["Pay Period End"];

  if (!(start instanceof Date) || !(end instanceof Date)) {
    throw new Error(
      "Admin_Config: Pay Period Start / End must be DATE values"
    );
  }

  return { start, end };
}
/**
 * Normalize names for matching PDFs ↔ directory rows
 * - lowercase
 * - strip punctuation
 * - collapse whitespace
 */

/**
 * GLOBAL CONFIG — EMAIL LOOKUPS
 */

function normalizeName_(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/[^a-z\s]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}
function getPayrollClinicians_() {
  const ss = SpreadsheetApp.getActive();
  const tabs = [
    { tab: "D9_Payroll", bucket: "D9" },
    { tab: "D10_Payroll", bucket: "D10" }
  ];

  const clinicians = new Set(); // name|bucket

  tabs.forEach(({ tab, bucket }) => {
    const sheet = ss.getSheetByName(tab);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    const headers = data.shift().map(h => String(h).toLowerCase());

    const fnCol = headers.indexOf("first name");
    const lnCol = headers.indexOf("last name");

    if (fnCol < 0 || lnCol < 0) {
      throw new Error(`${tab} missing First Name / Last Name`);
    }

    data.forEach(r => {
      if (!r[fnCol] || !r[lnCol]) return;
      const name = normalizeName_(`${r[fnCol]} ${r[lnCol]}`);
      clinicians.add(`${name}|${bucket}`);
    });
  });

  return clinicians;
}

function reconcileEmailAudit_(runId) {
  const ss = SpreadsheetApp.getActive();
  const audit = ss.getSheetByName("Email_Audit");
  if (!audit) throw new Error("Email_Audit not found");

  const payrollClinicians = getPayrollClinicians_();
  const auditData = audit.getDataRange().getValues();
  auditData.shift(); // headers

  const attempted = new Set();
  const drafted = new Set();

auditData.forEach(r => {
  if (r[0] !== runId) return;
  if (r[2] !== "PAYROLL") return;

  const status = r[7];     // Status column
  const subject = r[4];    // Subject column
  const pdfName = r[5];    // PDF Name column
  if (!pdfName) return;

  // Bucket is encoded in subject like: "[D9] Payroll ...", "[D10] Payroll ..."
  const m = String(subject || "").match(/^\[(D9|D10)\]/i);
  const bucket = m ? m[1].toUpperCase() : null;
  if (!bucket) return; // if bucket missing, don't poison reconciliation

  const name = normalizeName_(String(pdfName).replace(/\.pdf$/i, ""));
  const key = `${name}|${bucket}`;

  attempted.add(key);
  if (status === "DRAFTED") drafted.add(key);
});


  const missing = [];
payrollClinicians.forEach(key => {
  if (!attempted.has(key)) {
    const [name, bucket] = key.split("|");
    missing.push([name, bucket]);
  }
});


  return {
    expected: payrollClinicians.size,
    drafted: drafted.size,
    attempted: attempted.size,
    missing
  };
}
function reconcileInvoiceEmailAudit_(runId) {
  const ss = SpreadsheetApp.getActive();
  const audit = ss.getSheetByName("Email_Audit");
  if (!audit) throw new Error("Email_Audit not found");

  const attempted = new Set();
  const drafted = new Set();

  const rows = audit.getDataRange().getValues();
  rows.shift();

rows.forEach(r => {
  if (r[0] !== runId) return;
  if (r[2] !== "INVOICE") return;


    const pdfName = r[5];
    if (!pdfName) return;

   const name = normalizeName_(pdfName.replace(/\.pdf$/i, ""));
attempted.add(name);


    if (r[7] === "DRAFTED") drafted.add(name);
  });

  return {
    expected: attempted.size,
    drafted: drafted.size,
    attempted: attempted.size,
    missing: [] // agencies don't need payroll-style missing logic
  };
}
function generateRunId_() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyyMMdd_HHmmss"
  );
}
function resetPhase8EmailRun() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "Reset Phase 8 Email Run",
    "This will:\n\n" +
    "• Delete Phase 8 email drafts\n" +
    "• Remove latest run rows from Email_Audit\n\n" +
    "PDFs will NOT be deleted.\n\nProceed?",
    ui.ButtonSet.YES_NO
  );

  if (resp !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActive();
  const audit = ss.getSheetByName("Email_Audit");
  if (!audit) throw new Error("Email_Audit not found");

  const data = audit.getDataRange().getValues();
  if (data.length < 2) return;

  const header = data.shift();

  // Find latest Run ID
  const runIds = data
    .map(r => r[0])
    .filter(Boolean)
    .sort();

  const lastRunId = runIds[runIds.length - 1];
  if (!lastRunId) return;

  // 1️⃣ Delete drafts
  deletePhase8Drafts_();

  // 2️⃣ Remove audit rows for that run
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] === lastRunId) {
      audit.deleteRow(i + 2);
    }
  }

  ui.alert(`Phase 8 reset complete.\nRun ID cleared: ${lastRunId}`);
}
function deletePhase8Drafts_() {
  const admin = SpreadsheetApp.getActive().getSheetByName("Admin_Config");
  if (!admin) throw new Error("Admin_Config missing");

  const getCfg = label => {
    const cell = admin.createTextFinder(label).findNext();
    return cell ? admin.getRange(cell.getRow(), 2).getDisplayValue() : "";
  };

  const payrollSubject = getCfg("Payroll Email Subject");
  const invoiceSubject = getCfg("Invoice Email Subject");

  const queries = [
    `in:drafts subject:"${payrollSubject}"`,
    `in:drafts subject:"${invoiceSubject}"`
  ];

  queries.forEach(q => {
    const threads = GmailApp.search(q);
    threads.forEach(t => t.moveToTrash());
  });
}function sendPhase8Drafts() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "Send Phase 8 Draft Emails",
    "This will SEND all Phase 8 payroll and invoice drafts.\n\n" +
    "This action CANNOT be undone.\n\nProceed?",
    ui.ButtonSet.YES_NO
  );

  if (resp !== ui.Button.YES) return;

  const admin = SpreadsheetApp.getActive().getSheetByName("Admin_Config");
  if (!admin) throw new Error("Admin_Config missing");

  const getCfg = label => {
    const cell = admin.createTextFinder(label).findNext();
    return cell ? admin.getRange(cell.getRow(), 2).getDisplayValue() : "";
  };

  const payrollSubject = getCfg("Payroll Email Subject");
  const invoiceSubject = getCfg("Invoice Email Subject");

  const queries = [
    `in:drafts subject:"${payrollSubject}"`,
    `in:drafts subject:"${invoiceSubject}"`
  ];

  let sent = 0;

  queries.forEach(q => {
    const threads = GmailApp.search(q);

    threads.forEach(t => {
      let messages;
      try {
        messages = t.getMessages();
      } catch (e) {
        return;
      }

      messages.forEach(m => {
        if (!m.isDraft()) return;

        try {
          m.send();
          sent++;
        } catch (e) {
          Logger.log(`Failed to send draft: ${e.message}`);
        }
      });
    });
  });

  ui.alert(`Phase 8 complete.\n\nSent ${sent} emails.`);
}


function draftSpecialInvoiceEmail_(invoiceSheet) {
  const admin = SpreadsheetApp.getActive().getSheetByName("Admin_Config");

  const subject = admin.createTextFinder("Invoice Email Subject").findNext();
  const body = admin.createTextFinder("Invoice Email Body").findNext();

  GmailApp.createDraft(
    admin.getRange(subject.getRow(), 2).getValue(),
    "[SPECIAL] D9 All About You Invoice",
    "",
    {
      htmlBody: admin.getRange(body.getRow(), 2).getValue()
    }
  );
}
