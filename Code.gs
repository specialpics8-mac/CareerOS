// ═══════════════════════════════════════════════════════════
// CAREEROS — Google Apps Script Backend V0.2
// Complete backend for CareerOS
//
// Deploy:
//   Deploy > Manage deployments > Edit
//   Execute as: Me
//   Who has access: Anyone
//
// Creates / uses a Google Sheet called "CareerOS"
// and a Google Drive folder called "CareerOS Resume Library".
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
  META: "Meta",
  RESUME_LIBRARY: "ResumeLibrary"
};

const HEADERS = {
  Jobs: [
    "ID",
    "Company",
    "Title",
    "Location",
    "URL",
    "Source",
    "DateFound",
    "Match",
    "Priority",
    "JD",
    "Status",
    "Notes"
  ],

  Applications: [
    "ID",
    "JobID",
    "Company",
    "Title",
    "Status",
    "DateApplied",
    "ResumeID",
    "Recruiter",
    "NextAction",
    "FollowUpDate",
    "Notes"
  ],

  Companies: [
    "ID",
    "Name",
    "Category",
    "ATS",
    "SourceURL",
    "Active",
    "LastChecked",
    "Notes"
  ],

  Contacts: [
    "ID",
    "Name",
    "Company",
    "Title",
    "Relationship",
    "LinkedIn",
    "Email",
    "LastContact",
    "NextAction",
    "Notes"
  ],

  Resumes: [
    "ID",
    "Name",
    "Text",
    "Version",
    "TargetRole",
    "TargetCompany",
    "CreatedAt",
    "UpdatedAt"
  ],

  Interviews: [
    "ID",
    "ApplicationID",
    "Company",
    "Title",
    "Round",
    "Date",
    "Time",
    "Location",
    "Notes"
  ],

  FollowUps: [
    "ID",
    "ApplicationID",
    "Company",
    "Title",
    "DueDate",
    "Type",
    "Status",
    "Notes"
  ],

  EmailEvents: [
    "ID",
    "MessageID",
    "Date",
    "Sender",
    "Subject",
    "EventType",
    "Company",
    "Title",
    "ApplicationID",
    "Confidence",
    "RawSnippet"
  ],

  Portals: [
    "ID",
    "Name",
    "URL",
    "Notes",
    "Active"
  ],

  Settings: [
    "Key",
    "Value"
  ],

  Meta: [
    "Key",
    "Value"
  ],

  ResumeLibrary: [
    "ID",
    "Name",
    "Type",
    "FileId",
    "FileURL",
    "MimeType",
    "Size",
    "TargetRole",
    "TargetCompany",
    "CreatedAt",
    "UpdatedAt",
    "Notes"
  ]
};


// ═══════════════════════════════════════════════════════════
// WEB APP ENTRY POINTS
// ═══════════════════════════════════════════════════════════

function doGet(e) {
  return handleRequest_(e);
}

function doPost(e) {
  return handleRequest_(e);
}


// ═══════════════════════════════════════════════════════════
// REQUEST ROUTER
// ═══════════════════════════════════════════════════════════

function handleRequest_(e) {
  const action = (e.parameter && e.parameter.action) || "";

  try {
    let result;

    if (action === "bootstrap") {
      result = bootstrap_();

    } else if (action === "fetchAll") {
      result = fetchAll_();

    } else if (action === "save") {
      result = saveRecord_(e.parameter);

    } else if (action === "delete") {
      result = deleteRecord_(e.parameter);

    } else if (action === "seedCompanies") {
      result = seedCompanies_();

    } else if (action === "saveSetting") {
      result = saveSetting_(e.parameter);

    } else if (action === "uploadResumeFile") {
      result = uploadResumeFile_(e.parameter);

    } else if (action === "deleteResumeFile") {
      result = deleteResumeFile_(e.parameter);

    } else {
      result = {
        error: "Unknown action: " + action
      };
    }

    return json_(result);

  } catch (err) {

    return json_({
      error: err.message,
      stack: err.stack
    });
  }
}


// ═══════════════════════════════════════════════════════════
// JSON RESPONSE
// ═══════════════════════════════════════════════════════════

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// ═══════════════════════════════════════════════════════════
// GOOGLE SHEETS
// ═══════════════════════════════════════════════════════════

function getSpreadsheet_() {

  const props = PropertiesService.getScriptProperties();

  const existingId = props.getProperty("SPREADSHEET_ID");

  if (existingId) {

    try {
      return SpreadsheetApp.openById(existingId);

    } catch (err) {
      // Spreadsheet was deleted or inaccessible.
      // Recreate it below.
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

      sheet
        .getRange(1, 1, 1, HEADERS[name].length)
        .setValues([HEADERS[name]]);

      sheet.setFrozenRows(1);
    }
  }

  // Remove default Sheet1 once proper sheets exist.
  const defaultSheet = ss.getSheetByName("Sheet1");

  if (
    defaultSheet &&
    ss.getSheets().length > 1
  ) {
    ss.deleteSheet(defaultSheet);
  }

  return sheet;
}


// ═══════════════════════════════════════════════════════════
// BOOTSTRAP
// ═══════════════════════════════════════════════════════════

function bootstrap_() {

  Object.keys(SHEETS).forEach(function(key) {
    getSheet_(SHEETS[key]);
  });

  return {
    ok: true,
    spreadsheetId: getSpreadsheet_().getId(),
    sheets: Object.values(SHEETS)
  };
}


// ═══════════════════════════════════════════════════════════
// READ SHEET ROWS
// ═══════════════════════════════════════════════════════════

function rowsToObjects_(sheetName) {

  const sheet = getSheet_(sheetName);

  const data = sheet
    .getDataRange()
    .getValues();

  if (data.length <= 1) {
    return [];
  }

  const headers = data[0].map(String);

  return data
    .slice(1)
    .map(function(row) {

      const obj = {};

      headers.forEach(function(header, index) {
        obj[header] = row[index];
      });

      return obj;
    })
    .filter(function(obj) {
      return obj.ID || obj.Key;
    });
}


// ═══════════════════════════════════════════════════════════
// FETCH EVERYTHING
// ═══════════════════════════════════════════════════════════

function fetchAll_() {

  const result = {};

  Object.values(SHEETS).forEach(function(name) {
    result[name] = rowsToObjects_(name);
  });

  return result;
}


// ═══════════════════════════════════════════════════════════
// GENERIC SAVE
// ═══════════════════════════════════════════════════════════

function saveRecord_(p) {

  const sheetName = p.sheet;

  if (!HEADERS[sheetName]) {
    throw new Error("Invalid sheet: " + sheetName);
  }

  const sheet = getSheet_(sheetName);

  const headers = HEADERS[sheetName];

  const id = p.id || Utilities.getUuid();

  const row = headers.map(function(header) {

    if (header === "ID") {
      return id;
    }

    return p[header] !== undefined
      ? p[header]
      : "";
  });

  const data = sheet
    .getDataRange()
    .getValues();

  for (let i = 1; i < data.length; i++) {

    if (
      String(data[i][0]) === String(id)
    ) {

      sheet
        .getRange(
          i + 1,
          1,
          1,
          row.length
        )
        .setValues([row]);

      return {
        ok: true,
        id: id,
        updated: true
      };
    }
  }

  sheet.appendRow(row);

  return {
    ok: true,
    id: id,
    created: true
  };
}


// ═══════════════════════════════════════════════════════════
// GENERIC DELETE
// ═══════════════════════════════════════════════════════════

function deleteRecord_(p) {

  const sheetName = p.sheet;

  const id = p.id;

  if (!HEADERS[sheetName]) {
    throw new Error("Invalid sheet: " + sheetName);
  }

  const sheet = getSheet_(sheetName);

  const data = sheet
    .getDataRange()
    .getValues();

  for (let i = 1; i < data.length; i++) {

    if (
      String(data[i][0]) === String(id)
    ) {

      sheet.deleteRow(i + 1);

      return {
        ok: true
      };
    }
  }

  return {
    ok: false,
    error: "Record not found"
  };
}


// ═══════════════════════════════════════════════════════════
// SETTINGS
// ═══════════════════════════════════════════════════════════

function saveSetting_(p) {

  const sheet = getSheet_(SHEETS.SETTINGS);

  const data = sheet
    .getDataRange()
    .getValues();

  const key = p.key || "";

  const value = p.value || "";

  for (let i = 1; i < data.length; i++) {

    if (
      String(data[i][0]) === key
    ) {

      sheet
        .getRange(i + 1, 2)
        .setValue(value);

      return {
        ok: true
      };
    }
  }

  sheet.appendRow([
    key,
    value
  ]);

  return {
    ok: true
  };
}


// ═══════════════════════════════════════════════════════════
// RESUME & DOCUMENT LIBRARY
// ═══════════════════════════════════════════════════════════

function getResumeLibraryFolder_() {

  const folderName =
    "CareerOS Resume Library";

  const folders =
    DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  }

  return DriveApp.createFolder(folderName);
}


// ═══════════════════════════════════════════════════════════
// UPLOAD RESUME / DOCUMENT
// ═══════════════════════════════════════════════════════════

function uploadResumeFile_(p) {

  const name =
    String(p.name || "CareerOS Document");

  const mimeType =
    String(
      p.mimeType ||
      "application/octet-stream"
    );

  const base64 =
    String(p.base64 || "");

  const type =
    String(p.type || "Resume");

  const targetRole =
    String(p.targetRole || "");

  const targetCompany =
    String(p.targetCompany || "");

  const notes =
    String(p.notes || "");

  if (!base64) {
    throw new Error("No file data received.");
  }


  // Maximum decoded file size:
  // 5 MB
  const MAX_BYTES =
    5 * 1024 * 1024;

  const estimatedBytes =
    Math.floor(
      base64.length * 3 / 4
    );

  if (estimatedBytes > MAX_BYTES) {

    throw new Error(
      "File is too large. Maximum size is 5 MB."
    );
  }


  // Decode Base64.
  const bytes =
    Utilities.base64Decode(base64);


  // Supported extensions.
  const lowerName =
    name.toLowerCase();

  const allowed =
    lowerName.endsWith(".pdf") ||
    lowerName.endsWith(".doc") ||
    lowerName.endsWith(".docx") ||
    lowerName.endsWith(".txt");

  if (!allowed) {

    throw new Error(
      "Unsupported file type. Please upload PDF, DOC, DOCX or TXT."
    );
  }


  // Create Drive blob.
  const blob =
    Utilities.newBlob(
      bytes,
      mimeType,
      name
    );


  // Get CareerOS folder.
  const folder =
    getResumeLibraryFolder_();


  // Create file.
  const file =
    folder.createFile(blob);


  // Give file a sensible name.
  file.setName(name);


  const now =
    new Date();

  const id =
    Utilities.getUuid();


  // Build record.
  const record = {

    ID: id,

    Name: name,

    Type: type,

    FileId: file.getId(),

    FileURL: file.getUrl(),

    MimeType: mimeType,

    Size: bytes.length,

    TargetRole: targetRole,

    TargetCompany: targetCompany,

    CreatedAt: now,

    UpdatedAt: now,

    Notes: notes
  };


  // Save metadata in Sheet.
  const sheet =
    getSheet_(
      SHEETS.RESUME_LIBRARY
    );


  sheet.appendRow([

    record.ID,

    record.Name,

    record.Type,

    record.FileId,

    record.FileURL,

    record.MimeType,

    record.Size,

    record.TargetRole,

    record.TargetCompany,

    record.CreatedAt,

    record.UpdatedAt,

    record.Notes

  ]);


  return {

    ok: true,

    id: id,

    record: record

  };
}


// ═══════════════════════════════════════════════════════════
// DELETE RESUME / DOCUMENT
// ═══════════════════════════════════════════════════════════

function deleteResumeFile_(p) {

  const id =
    String(p.id || "");

  if (!id) {

    throw new Error(
      "Missing resume/document ID."
    );
  }


  const sheet =
    getSheet_(
      SHEETS.RESUME_LIBRARY
    );


  const data =
    sheet
      .getDataRange()
      .getValues();


  for (let i = 1; i < data.length; i++) {

    if (
      String(data[i][0]) === id
    ) {

      const fileId =
        String(data[i][3] || "");


      // Move Drive file to trash.
      if (fileId) {

        try {

          const file =
            DriveApp.getFileById(fileId);

          file.setTrashed(true);

        } catch (err) {

          // Continue deleting metadata
          // even if Drive file is missing.
        }
      }


      // Delete Sheet row.
      sheet.deleteRow(i + 1);


      return {
        ok: true
      };
    }
  }


  return {

    ok: false,

    error:
      "Resume/document not found."
  };
}


// ═══════════════════════════════════════════════════════════
// SEED TARGET COMPANIES
// ═══════════════════════════════════════════════════════════

function seedCompanies_() {

  const companies = [

    ["HSBC", "Global Bank"],
    ["DBS", "Global Bank"],
    ["American Express", "Global Bank"],
    ["Bank of America", "Global Bank"],
    ["Citi", "Global Bank"],
    ["Standard Chartered", "Global Bank"],
    ["Deutsche Bank", "Global Bank"],
    ["Barclays", "Global Bank"],
    ["Goldman Sachs", "Global Bank"],
    ["JPMorgan", "Global Bank"],

    ["Razorpay", "Indian Fintech"],
    ["PayU", "Indian Fintech"],
    ["Cashfree Payments", "Indian Fintech"],
    ["Pine Labs", "Indian Fintech"],
    ["Juspay", "Indian Fintech"],
    ["Poonawalla Fincorp", "Indian Fintech"],
    ["Jio Finance", "Indian Fintech"],
    ["Kiwi Credit Card", "Indian Fintech"],
    ["Kiwi Insurance", "Indian Fintech"],
    ["Fibe", "Indian Fintech"],
    ["Scapia", "Indian Fintech"],
    ["CRED", "Indian Fintech"],
    ["Groww", "Indian Fintech"],
    ["PhonePe", "Indian Fintech"],
    ["Jupiter", "Indian Fintech"],
    ["Navi", "Indian Fintech"],
    ["Slice", "Indian Fintech"],
    ["Fi Money", "Indian Fintech"],
    ["OneCard", "Indian Fintech"],

    ["Stripe", "Global Fintech"],
    ["Airwallex", "Global Fintech"],

    ["Mastercard", "Payments"],
    ["NPCI", "Payments"],
    ["Visa", "Payments"],

    ["BlackRock", "Finance/Data"],
    ["Motilal Oswal", "Finance/Data"],
    ["CRISIL", "Finance/Data"],
    ["ABFC", "Finance/Data"],

    ["Google", "Big Tech"],
    ["Adobe", "Big Tech"],
    ["Amazon", "Big Tech"],
    ["Flipkart", "Big Tech"],
    ["Meesho", "Big Tech"]
  ];


  const sheet =
    getSheet_(
      SHEETS.COMPANIES
    );


  const existing =
    rowsToObjects_(
      SHEETS.COMPANIES
    );


  const names = {};


  existing.forEach(function(x) {

    names[
      String(x.Name)
        .toLowerCase()
    ] = true;

  });


  let added = 0;


  companies.forEach(function(c) {

    const companyName =
      c[0].toLowerCase();


    if (!names[companyName]) {

      sheet.appendRow([

        Utilities.getUuid(),

        c[0],

        c[1],

        "",

        "",

        true,

        "",

        ""

      ]);

      added++;
    }

  });


  return {

    ok: true,

    added: added,

    total:
      existing.length + added

  };
}


// ═══════════════════════════════════════════════════════════
// PUBLIC BOOTSTRAP WRAPPER
// ═══════════════════════════════════════════════════════════

function bootstrap() {

  return bootstrap_();

}
