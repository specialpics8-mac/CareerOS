// ═══════════════════════════════════════════════════════════
// CAREEROS — Google Apps Script Backend V0.1
// Deploy: Deploy > New deployment > Web app
// Execute as: Me
// Who has access: Anyone
//
// Creates a private Google Sheet called "CareerOS" in the
// deploying user's Drive and remembers its ID in Script Properties.
// ═══════════════════════════════════════════════════════════

const SPREADSHEET_NAME = "CareerOS";

const SHEETS = {
  JOBS: "Jobs",
  APPLICATIONS: "Applications",
  COMPANIES: "Companies",
  CONTACTS: "Contacts",
  RESUMES: "Resumes",
  INTERVIEWS: "Interviews",
  FOLLOWUPS: "FollowUps",
  EMAIL_EVENTS: "EmailEvents",
  PORTALS: "Portals",
  SETTINGS: "Settings",
  META: "Meta"
};

const HEADERS = {
  Jobs: ["ID","Company","Title","Location","URL","Source","DateFound","Match","Priority","JD","Status","Notes"],
  Applications: ["ID","JobID","Company","Title","Status","DateApplied","ResumeID","Recruiter","NextAction","FollowUpDate","Notes"],
  Companies: ["ID","Name","Category","ATS","SourceURL","Active","LastChecked","Notes"],
  Contacts: ["ID","Name","Company","Title","Relationship","LinkedIn","Email","LastContact","NextAction","Notes"],
  Resumes: ["ID","Name","Text","Version","TargetRole","TargetCompany","CreatedAt","UpdatedAt"],
  Interviews: ["ID","ApplicationID","Company","Title","Round","Date","Time","Location","Notes"],
  FollowUps: ["ID","ApplicationID","Company","Title","DueDate","Type","Status","Notes"],
  EmailEvents: ["ID","MessageID","Date","Sender","Subject","EventType","Company","Title","ApplicationID","Confidence","RawSnippet"],
  Portals: ["ID","Name","URL","Notes","Active"],
  Settings: ["Key","Value"],
  Meta: ["Key","Value"]
};

function doGet(e) {
  return handleRequest_(e);
}

function doPost(e) {
  return handleRequest_(e);
}

function handleRequest_(e) {
  const action = (e.parameter && e.parameter.action) || "";
  try {
    let result;
    if (action === "bootstrap") result = bootstrap_();
    else if (action === "fetchAll") result = fetchAll_();
    else if (action === "save") result = saveRecord_(e.parameter);
    else if (action === "delete") result = deleteRecord_(e.parameter);
    else if (action === "seedCompanies") result = seedCompanies_();
    else if (action === "saveSetting") result = saveSetting_(e.parameter);
    else result = { error: "Unknown action: " + action };

    return json_(result);
  } catch (err) {
    return json_({ error: err.message, stack: err.stack });
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty("SPREADSHEET_ID");

  if (existingId) {
    try {
      return SpreadsheetApp.openById(existingId);
    } catch (err) {
      // Recreate if deleted or inaccessible.
    }
  }

  const ss = SpreadsheetApp.create(SPREADSHEET_NAME);
  props.setProperty("SPREADSHEET_ID", ss.getId());
  return ss;
}

function getSheet_(name) {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (HEADERS[name]) {
      sheet.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]);
      sheet.setFrozenRows(1);
    }
  }

  const defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  return sheet;
}

function bootstrap_() {
  Object.keys(SHEETS).forEach(k => getSheet_(SHEETS[k]));
  return {
    ok: true,
    spreadsheetId: getSpreadsheet_().getId(),
    sheets: Object.values(SHEETS)
  };
}

function rowsToObjects_(sheetName) {
  const sheet = getSheet_(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0].map(String);
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  }).filter(obj => obj.ID || obj.Key);
}

function fetchAll_() {
  const result = {};
  Object.values(SHEETS).forEach(name => {
    result[name] = rowsToObjects_(name);
  });
  return result;
}

function saveRecord_(p) {
  const sheetName = p.sheet;
  if (!HEADERS[sheetName]) throw new Error("Invalid sheet: " + sheetName);

  const sheet = getSheet_(sheetName);
  const headers = HEADERS[sheetName];
  const id = p.id || Utilities.getUuid();

  const row = headers.map(h => {
    if (h === "ID") return id;
    return p[h] !== undefined ? p[h] : "";
  });

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return { ok: true, id: id, updated: true };
    }
  }

  sheet.appendRow(row);
  return { ok: true, id: id, created: true };
}

function deleteRecord_(p) {
  const sheetName = p.sheet;
  const id = p.id;
  if (!HEADERS[sheetName]) throw new Error("Invalid sheet: " + sheetName);

  const sheet = getSheet_(sheetName);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }

  return { ok: false, error: "Record not found" };
}

function saveSetting_(p) {
  const sheet = getSheet_(SHEETS.SETTINGS);
  const data = sheet.getDataRange().getValues();
  const key = p.key || "";
  const value = p.value || "";

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return { ok: true };
    }
  }

  sheet.appendRow([key, value]);
  return { ok: true };
}

function seedCompanies_() {
  const companies = [
    ["HSBC","Global Bank"],["DBS","Global Bank"],["American Express","Global Bank"],["Bank of America","Global Bank"],["Citi","Global Bank"],["Standard Chartered","Global Bank"],["Deutsche Bank","Global Bank"],["Barclays","Global Bank"],["Goldman Sachs","Global Bank"],["JPMorgan","Global Bank"],
    ["Razorpay","Indian Fintech"],["PayU","Indian Fintech"],["Cashfree Payments","Indian Fintech"],["Pine Labs","Indian Fintech"],["Juspay","Indian Fintech"],["Poonawalla Fincorp","Indian Fintech"],["Jio Finance","Indian Fintech"],["Kiwi Credit Card","Indian Fintech"],["Kiwi Insurance","Indian Fintech"],["Fibe","Indian Fintech"],["Scapia","Indian Fintech"],["CRED","Indian Fintech"],["Groww","Indian Fintech"],["PhonePe","Indian Fintech"],["Jupiter","Indian Fintech"],["Navi","Indian Fintech"],["Slice","Indian Fintech"],["Fi Money","Indian Fintech"],["OneCard","Indian Fintech"],
    ["Stripe","Global Fintech"],["Airwallex","Global Fintech"],["Mastercard","Payments"],["NPCI","Payments"],["Visa","Payments"],["BlackRock","Finance/Data"],["Motilal Oswal","Finance/Data"],["CRISIL","Finance/Data"],["ABFC","Finance/Data"],["Google","Big Tech"],["Adobe","Big Tech"],["Amazon","Big Tech"],["Flipkart","Big Tech"],["Meesho","Big Tech"]
  ];

  const sheet = getSheet_(SHEETS.COMPANIES);
  const existing = rowsToObjects_(SHEETS.COMPANIES);
  const names = {};
  existing.forEach(x => names[String(x.Name).toLowerCase()] = true);

  let added = 0;
  companies.forEach(c => {
    if (!names[c[0].toLowerCase()]) {
      sheet.appendRow([Utilities.getUuid(), c[0], c[1], "", "", true, "", ""]);
      added++;
    }
  });

  return { ok: true, added: added, total: existing.length + added };
}
