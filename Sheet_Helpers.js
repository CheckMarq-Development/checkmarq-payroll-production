/***************************************************
 * SHEET HELPERS — SHARED
 ***************************************************/
function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  if (headers && headers.length) {
    const existingHeaders =
      sheet.getLastRow() >= 1
        ? sheet.getRange(1, 1, 1, headers.length).getValues()[0]
        : [];

    const mismatch =
      existingHeaders.join("|") !== headers.join("|");

    if (mismatch) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }

  return sheet;
}
function assertPayPeriodConfigured_() {
  const ss = SpreadsheetApp.getActive();
  const admin = ss.getSheetByName("Admin_Config");
  if (!admin) throw new Error("Admin_Config tab not found");

  const requiredLabels = [
    "Approved From",
    "Approved To",
    "Pay Period Start",
    "Pay Period End"
  ];

  const missing = [];

  requiredLabels.forEach(label => {
    const cell = admin.createTextFinder(label).findNext();
    if (!cell) {
      missing.push(label + " (label not found)");
      return;
    }
    const value = admin.getRange(cell.getRow(), 2).getValue();
    if (!value) missing.push(label);
  });

  if (missing.length) {
    throw new Error(
      "Cannot proceed. The following Admin_Config fields are required:\n\n" +
      missing.join("\n")
    );
  }
}

function indexColumns_(headers, map) {
  const normalize = s =>
    String(s).toLowerCase().replace(/\u00a0/g, " ").trim();

  const normalizedHeaders = headers.map(normalize);
  const idx = {};

  Object.keys(map).forEach(k => {
    const target = normalize(map[k]);
    const i = normalizedHeaders.indexOf(target);
    if (i === -1) throw new Error(`Missing column: ${map[k]}`);
    idx[k] = i;
  });

  return idx;
}
function getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}