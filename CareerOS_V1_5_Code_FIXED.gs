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
  META: "Meta",
  RESUME_LIBRARY: "ResumeLibrary",
  JOB_ANALYSES: "JobAnalyses"
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
  Meta: ["Key","Value"],
  ResumeLibrary: ["ID","Name","Type","FileId","FileURL","MimeType","Size","TargetRole","TargetCompany","CreatedAt","UpdatedAt","Notes","ExtractedText","IndexedAt","IndexStatus"],
  JobAnalyses: ["ID","URL","Company","Title","JD","Match","BestResumeID","BestResumeName","Recommendation","CreatedAt","UpdatedAt"]
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
// REQUEST HANDLER
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

    } else if (action === "saveApplication") {
      result = saveApplication_(e.parameter);

    } else if (action === "delete") {
      result = deleteRecord_(e.parameter);

    } else if (action === "seedCompanies") {
      result = seedCompanies_();

    } else if (action === "saveSetting") {
      result = saveSetting_(e.parameter);

    } else if (action === "uploadResumeFile") {
      result = uploadResumeFile_(e.parameter);
    } else if (action === "indexResume") {
      result = indexResume_(e.parameter.id);
    } else if (action === "indexAllResumes") {
      result = indexAllResumes_();
    } else if (action === "repairResumeLibraryHeaders") {
      result = repairResumeLibraryHeaders_();

    } else if (action === "deleteResumeFile") {
      result = deleteResumeFile_(e.parameter);
    } else if (action === "deleteApplicationFast") {
      result = deleteRecord_(e.parameter);
    } else if (action === "analyzeJob") {
      result = analyzeJobV15_(e.parameter);
    } else if (action === "analyzePastedJob") {
      result = analyzePastedJobV15_(e.parameter);
    } else if (action === "getGmailStatus") {
      result = getGmailStatus_();
    } else if (action === "syncGmail") {
      result = syncGmailV15_(e.parameter);
    } else if (action === "getGmailEvents") {
      result = getGmailEvents_(e.parameter);

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
    .createTextOutput(
      JSON.stringify(obj)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );

}


// ═══════════════════════════════════════════════════════════
// GOOGLE SHEET
// ═══════════════════════════════════════════════════════════

function getSpreadsheet_() {

  const props =
    PropertiesService.getScriptProperties();

  const existingId =
    props.getProperty("SPREADSHEET_ID");

  if (existingId) {

    try {

      return SpreadsheetApp.openById(
        existingId
      );

    } catch (err) {

      // Recreate if deleted or inaccessible.

    }

  }

  const ss =
    SpreadsheetApp.create(
      SPREADSHEET_NAME
    );

  props.setProperty(
    "SPREADSHEET_ID",
    ss.getId()
  );

  return ss;
}


function ensureHeaders_(sheet, headerList) {
  if (!sheet || !headerList || !headerList.length) return;

  var lastColumn = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();

  if (lastRow === 0 || lastColumn === 0) {
    sheet.getRange(1, 1, 1, headerList.length).setValues([headerList]);
    sheet.setFrozenRows(1);
    return;
  }

  var currentCount = Math.max(lastColumn, 1);
  var current = sheet
    .getRange(1, 1, 1, currentCount)
    .getValues()[0]
    .map(String);

  var missing = headerList.filter(function(header) {
    return current.indexOf(header) === -1;
  });

  if (missing.length > 0) {
    var startColumn = current.length + 1;
    sheet
      .getRange(1, startColumn, 1, missing.length)
      .setValues([missing]);
  }

  sheet.setFrozenRows(1);
}

function getSheet_(name) {

  const ss =
    getSpreadsheet_();

  let sheet =
    ss.getSheetByName(name);

  if (!sheet) {

    sheet =
      ss.insertSheet(name);

    if (HEADERS[name]) {

      sheet
        .getRange(
          1,
          1,
          1,
          HEADERS[name].length
        )
        .setValues([
          HEADERS[name]
        ]);

      sheet.setFrozenRows(1);

    }

  }

  const defaultSheet =
    ss.getSheetByName("Sheet1");

  if (
    defaultSheet &&
    ss.getSheets().length > 1
  ) {

    ss.deleteSheet(
      defaultSheet
    );

  }

  return sheet;
}


// ═══════════════════════════════════════════════════════════
// BOOTSTRAP
// ═══════════════════════════════════════════════════════════

function bootstrap_() {

  Object.keys(SHEETS)
    .forEach(function(key) {

      getSheet_(
        SHEETS[key]
      );

    });

  return {

    ok: true,

    spreadsheetId:
      getSpreadsheet_().getId(),

    sheets:
      Object.values(SHEETS)

  };

}


// ═══════════════════════════════════════════════════════════
// READ SHEETS
// ═══════════════════════════════════════════════════════════

function rowsToObjects_(sheetName) {

  const sheet =
    getSheet_(sheetName);

  const data =
    sheet
      .getDataRange()
      .getValues();

  if (data.length <= 1) {
    return [];
  }

  const headers =
    data[0].map(String);

  return data
    .slice(1)
    .map(function(row) {

      const obj = {};

      headers.forEach(
        function(header, index) {

          obj[header] =
            row[index];

        }
      );

      return obj;

    })
    .filter(function(obj) {

      return obj.ID || obj.Key;

    });

}


// ═══════════════════════════════════════════════════════════
// FETCH ALL
// ═══════════════════════════════════════════════════════════

function fetchAll_() {

  const result = {};

  Object.values(SHEETS)
    .forEach(function(name) {

      result[name] =
        rowsToObjects_(name);

    });

  return result;
}


// ═══════════════════════════════════════════════════════════
// SAVE RECORD
// ═══════════════════════════════════════════════════════════

function saveRecord_(p) {

  const sheetName =
    p.sheet;

  if (!HEADERS[sheetName]) {

    throw new Error(
      "Invalid sheet: " +
      sheetName
    );

  }

  const sheet =
    getSheet_(sheetName);

  const headers =
    HEADERS[sheetName];

  const id =
    p.id ||
    Utilities.getUuid();

  const row =
    headers.map(function(header) {

      if (header === "ID") {
        return id;
      }

      return p[header] !== undefined
        ? p[header]
        : "";

    });

  const data =
    sheet
      .getDataRange()
      .getValues();

  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    if (
      String(data[i][0]) ===
      String(id)
    ) {

      sheet
        .getRange(
          i + 1,
          1,
          1,
          row.length
        )
        .setValues([
          row
        ]);

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
// DELETE RECORD
// ═══════════════════════════════════════════════════════════

function deleteRecord_(p) {

  const sheetName =
    p.sheet;

  const id =
    p.id;

  if (!HEADERS[sheetName]) {

    throw new Error(
      "Invalid sheet: " +
      sheetName
    );

  }

  const sheet =
    getSheet_(sheetName);

  const data =
    sheet
      .getDataRange()
      .getValues();

  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    if (
      String(data[i][0]) ===
      String(id)
    ) {

      sheet.deleteRow(
        i + 1
      );

      return {
        ok: true
      };

    }

  }

  return {

    ok: false,

    error:
      "Record not found"

  };

}


// ═══════════════════════════════════════════════════════════
// SETTINGS
// ═══════════════════════════════════════════════════════════

function saveSetting_(p) {

  const sheet =
    getSheet_(
      SHEETS.SETTINGS
    );

  const data =
    sheet
      .getDataRange()
      .getValues();

  const key =
    p.key || "";

  const value =
    p.value || "";

  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    if (
      String(data[i][0]) ===
      key
    ) {

      sheet
        .getRange(
          i + 1,
          2
        )
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

const RESUME_LIBRARY_FOLDER_NAME =
  "CareerOS Resume Library";

const MAX_RESUME_FILE_BYTES =
  5 * 1024 * 1024;


function getResumeLibraryFolder_() {

  const folders =
    DriveApp.getFoldersByName(
      RESUME_LIBRARY_FOLDER_NAME
    );

  if (folders.hasNext()) {

    return folders.next();

  }

  return DriveApp.createFolder(
    RESUME_LIBRARY_FOLDER_NAME
  );

}


// ═══════════════════════════════════════════════════════════
// UPLOAD RESUME / DOCUMENT
// ═══════════════════════════════════════════════════════════

function uploadResumeFile_(p) {

  const name =
    String(
      p.name || ""
    ).trim();

  const mimeType =
    String(
      p.mimeType ||
      "application/octet-stream"
    );

  const base64 =
    String(
      p.base64 || ""
    ).trim();

  const type =
    String(
      p.type ||
      "Resume"
    ).trim();

  const targetRole =
    String(
      p.targetRole ||
      ""
    ).trim();

  const targetCompany =
    String(
      p.targetCompany ||
      ""
    ).trim();

  const notes =
    String(
      p.notes ||
      ""
    ).trim();


  if (!name) {

    throw new Error(
      "Please select a file."
    );

  }

  if (!base64) {

    throw new Error(
      "File data was not received."
    );

  }


  const ext =
    name
      .toLowerCase()
      .split(".")
      .pop();

  const allowedExtensions = [
    "pdf",
    "doc",
    "docx",
    "txt"
  ];


  if (
    allowedExtensions.indexOf(
      ext
    ) === -1
  ) {

    throw new Error(
      "Only PDF, DOC, DOCX and TXT files are supported."
    );

  }


  const bytes =
    Utilities.base64Decode(
      base64
    );


  if (
    bytes.length >
    MAX_RESUME_FILE_BYTES
  ) {

    throw new Error(
      "File is larger than 5 MB. Please upload a smaller file."
    );

  }


  const folder =
    getResumeLibraryFolder_();


  const blob =
    Utilities.newBlob(
      bytes,
      mimeType,
      name
    );


  const file =
    folder.createFile(
      blob
    );


  file.setDescription(
    "CareerOS Resume & Document Library"
  );


  const now =
    new Date().toISOString();


  const recordId =
    Utilities.getUuid();


  const record = {

    ID: recordId,

    Name: name,

    Type: type,

    FileId:
      file.getId(),

    FileURL:
      file.getUrl(),

    MimeType:
      mimeType,

    Size:
      bytes.length,

    TargetRole:
      targetRole,

    TargetCompany:
      targetCompany,

    CreatedAt:
      now,

    UpdatedAt:
      now,

    Notes:
      notes

  };


  getSheet_(
    SHEETS.RESUME_LIBRARY
  ).appendRow([

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

    id:
      recordId,

    record:
      record

  };

}


// ═══════════════════════════════════════════════════════════
// DELETE RESUME / DOCUMENT
// ═══════════════════════════════════════════════════════════

function deleteResumeFile_(p) {

  const id =
    String(
      p.id || ""
    ).trim();


  if (!id) {

    throw new Error(
      "Document ID is required."
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


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    if (
      String(data[i][0]) ===
      id
    ) {

      const fileId =
        String(
          data[i][3] || ""
        );


      if (fileId) {

        try {

          DriveApp
            .getFileById(
              fileId
            )
            .setTrashed(true);

        } catch (err) {

          // File may already be missing/deleted.

        }

      }


      sheet.deleteRow(
        i + 1
      );


      return {
        ok: true
      };

    }

  }


  return {

    ok: false,

    error:
      "Document not found."

  };

}


// ═══════════════════════════════════════════════════════════
// SEED COMPANIES
// ═══════════════════════════════════════════════════════════


function saveApplication_(p) {
  if (!p.company || !p.title) {
    throw new Error("Company and role are required.");
  }
  return saveRecord_({
    sheet: "Applications",
    id: p.id || Utilities.getUuid(),
    JobID: p.jobId || "",
    Company: p.company || "",
    Title: p.title || "",
    Status: p.status || "Applied",
    DateApplied: p.dateApplied || "",
    ResumeID: p.resumeId || "",
    Recruiter: p.recruiter || "",
    NextAction: p.nextAction || "",
    FollowUpDate: p.followUpDate || "",
    Notes: p.notes || ""
  });
}

function seedCompanies_() {
  const companies = [
    ["HSBC","Global Bank","",""],
    ["DBS","Global Bank","",""],
    ["American Express","Global Bank","",""],
    ["Bank of America","Global Bank","",""],
    ["Citi","Global Bank","",""],
    ["Standard Chartered","Global Bank","",""],
    ["Deutsche Bank","Global Bank","",""],
    ["Barclays","Global Bank","",""],
    ["Goldman Sachs","Global Bank","",""],
    ["JPMorgan","Global Bank","",""],

    ["Razorpay","Indian Fintech","",""],
    ["PayU","Indian Fintech","",""],
    ["Cashfree Payments","Indian Fintech","",""],
    ["Pine Labs","Indian Fintech","",""],
    ["Juspay","Indian Fintech","",""],
    ["Poonawalla Fincorp","Indian Fintech","",""],
    ["Jio Finance","Indian Fintech","",""],
    ["Kiwi Credit Card","Indian Fintech","",""],
    ["Kiwi Insurance","Indian Fintech","",""],
    ["Fibe","Indian Fintech","",""],
    ["Scapia","Indian Fintech","",""],
    ["CRED","Indian Fintech","",""],
    ["Groww","Indian Fintech","",""],
    ["PhonePe","Indian Fintech","",""],
    ["Jupiter","Indian Fintech","",""],
    ["Navi","Indian Fintech","",""],
    ["Slice","Indian Fintech","",""],
    ["Fi Money","Indian Fintech","",""],
    ["OneCard","Indian Fintech","",""],

    ["Stripe","Global Fintech","",""],
    ["Airwallex","Global Fintech","",""],

    ["Mastercard","Payments","",""],
    ["NPCI","Payments","",""],
    ["Visa","Payments","",""],

    ["BlackRock","Finance/Data","",""],
    ["Motilal Oswal","Finance/Data","",""],
    ["CRISIL","Finance/Data","",""],
    ["ABFC","Finance/Data","",""],

    ["Google","Big Tech","",""],
    ["Adobe","Big Tech","",""],
    ["Amazon","Big Tech","",""],
    ["Flipkart","Big Tech","",""],
    ["Meesho","Big Tech","",""],

    ["Paytm","Indian Fintech","","https://paytm.com/careers"],
    ["Zerodha","Indian Fintech","",""],
    ["Policybazaar","Insurtech","","https://www.policybazaar.com/careers/"],
    ["MobiKwik","Indian Fintech","","https://www.mobikwik.com/careers/"],
    ["BharatPe","Indian Fintech","",""],
    ["Zeta","Fintech Infrastructure","","https://www.zeta.tech/in/careers/work-with-us"],
    ["KreditBee","Indian Fintech","","https://www.kreditbee.in/careers"],
    ["Intuit","Fintech / SaaS","","https://jobs.intuit.com/"],
    ["Swiggy","Consumer Tech","",""],
    ["Myntra","Consumer Tech","",""],
    ["Zomato","Consumer Tech","",""],
    ["Uber","Consumer Tech / Mobility","",""]
  ];

  const sheet = getSheet_(SHEETS.COMPANIES);
  const data = sheet.getDataRange().getValues();
  const existing = {};
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][1] || "").toLowerCase();
    if (key) existing[key] = i + 1;
  }

  let added = 0;
  let linksUpdated = 0;

  companies.forEach(function(c) {
    const name = c[0];
    const key = name.toLowerCase();
    const category = c[1];
    const careersUrl = c[3];

    if (!existing[key]) {
      sheet.appendRow([
        Utilities.getUuid(),
        name,
        category,
        "",
        careersUrl,
        true,
        "",
        ""
      ]);
      added++;
      return;
    }

    if (careersUrl) {
      const rowNumber = existing[key];
      const currentUrl = String(sheet.getRange(rowNumber, 5).getValue() || "");
      if (!currentUrl) {
        sheet.getRange(rowNumber, 5).setValue(careersUrl);
        linksUpdated++;
      }
    }
  });

  return {
    ok: true,
    added: added,
    linksUpdated: linksUpdated,
    total: rowsToObjects_(SHEETS.COMPANIES).length
  };
}



// ═══════════════════════════════════════════════════════════
// MANUAL BOOTSTRAP WRAPPER
// ═══════════════════════════════════════════════════════════

function bootstrap() {

  return bootstrap_();

}

// ═══════════════════════════════════════════════════════════
// CAREEROS V1.3 — JOB ANALYSIS + GMAIL SYNC
// ═══════════════════════════════════════════════════════════

function getGmailStatus_() {
  let email = "";
  try {
    email = Session.getEffectiveUser().getEmail() || "";
  } catch (err) {}

  return {
    ok: true,
    account: email,
    expected: "vishalnihalani2@gmail.com",
    isExpected: String(email).toLowerCase() === "vishalnihalani2@gmail.com"
  };
}

function syncGmail_(p) {
  const props = PropertiesService.getScriptProperties();
  const lastSyncRaw = props.getProperty("CAREEROS_GMAIL_LAST_SYNC_MS") || "0";
  const lastSyncMs = Number(lastSyncRaw);
  const now = Date.now();

  // Start with a bounded, recent search. Subsequent runs only inspect
  // messages received since the previous successful sync.
  const afterDate = new Date(Math.max(
    Date.now() - 30 * 24 * 60 * 60 * 1000,
    lastSyncMs - 2 * 60 * 1000
  ));

  const y = afterDate.getFullYear();
  const m = String(afterDate.getMonth() + 1).padStart(2, "0");
  const d = String(afterDate.getDate()).padStart(2, "0");
  const query = "after:" + y + "/" + m + "/" + d;

  let threads = [];
  try {
    threads = GmailApp.search(query, 0, 25);
  } catch (err) {
    throw new Error("Gmail access failed: " + err.message);
  }

  const apps = rowsToObjects_(SHEETS.APPLICATIONS);
  const eventsSheet = getSheet_(SHEETS.EMAIL_EVENTS);
  const existing = rowsToObjects_(SHEETS.EMAIL_EVENTS);
  const seen = {};
  existing.forEach(function(e) {
    if (e.MessageID) seen[String(e.MessageID)] = true;
  });

  let processed = 0;
  let matched = 0;
  let updated = 0;

  threads.forEach(function(thread) {
    const messages = thread.getMessages();

    messages.forEach(function(msg) {
      const msgId = String(msg.getId());
      const ts = msg.getDate().getTime();

      if (seen[msgId] || ts <= lastSyncMs) return;

      const subject = String(msg.getSubject() || "");
      const sender = String(msg.getFrom() || "");

      // Only read bodies for messages that look job-related.
      const subjectCandidate =
        /application|applied|interview|screen|recruit|hiring|assessment|offer|rejection|regret|opportunity|next step|career/i
        .test(subject);

      if (!subjectCandidate) {
        seen[msgId] = true;
        return;
      }

      const body = String(msg.getPlainBody() || "").slice(0, 5000);
      const combined = (subject + "\n" + body).toLowerCase();

      const eventType =
        /rejected|regret|not selected|unsuccessful|decline/i.test(combined) ? "Rejected" :
        /offer|pleased to offer/i.test(combined) ? "Offer" :
        /interview|round|schedule.*call|technical|case study/i.test(combined) ? "Interview" :
        /screen|recruiter|phone call|talent acquisition/i.test(combined) ? "Recruiter Screen" :
        /application.*received|application.*submitted|thank.*applying|successfully applied/i.test(combined) ? "Applied" :
        "Job-related";

      let company = "";
      let title = "";

      // Match to an existing application using company/title words.
      let best = null;
      let bestScore = 0;

      apps.forEach(function(a) {
        const ac = String(a.company || "").toLowerCase();
        const at = String(a.title || "").toLowerCase();
        let score = 0;
        if (ac && combined.indexOf(ac) >= 0) score += 3;
        if (at && combined.indexOf(at) >= 0) score += 3;

        const titleWords = at.split(/[^a-z0-9]+/).filter(function(w) {
          return w.length >= 4;
        });
        titleWords.slice(0, 8).forEach(function(w) {
          if (combined.indexOf(w) >= 0) score += 0.5;
        });

        if (score > bestScore) {
          bestScore = score;
          best = a;
        }
      });

      if (best && bestScore > 0) {
        matched++;
        company = best.company || "";
        title = best.title || "";

        let desiredStatus = best.status;
        if (eventType === "Rejected" || eventType === "Offer" ||
            eventType === "Interview" || eventType === "Recruiter Screen" ||
            eventType === "Applied") {
          desiredStatus = eventType;
        }

        if (desiredStatus && desiredStatus !== best.status) {
          best.status = desiredStatus;
          if (desiredStatus === "Applied" && !best.dateApplied) {
            best.dateApplied = msg.getDate().toISOString();
          }
          const saveResult = saveRecord_({
            sheet: "Applications",
            id: best.id,
            company: best.company || "",
            title: best.title || "",
            status: best.status || "",
            dateApplied: best.dateApplied || "",
            resumeId: best.resumeId || "",
            recruiter: best.recruiter || "",
            nextAction: best.nextAction || "",
            followUpDate: best.followUpDate || "",
            notes: best.notes || ""
          });
          if (saveResult && saveResult.ok) updated++;
        }
      }

      eventsSheet.appendRow([
        Utilities.getUuid(),
        msgId,
        msg.getDate().toISOString(),
        sender,
        subject,
        eventType,
        company,
        title,
        best ? best.id : "",
        best ? "matched" : "unmatched",
        body.slice(0, 800)
      ]);

      processed++;
      seen[msgId] = true;
    });
  });

  props.setProperty("CAREEROS_GMAIL_LAST_SYNC_MS", String(now));

  return {
    ok: true,
    account: getGmailStatus_().account,
    threadsScanned: threads.length,
    messagesProcessed: processed,
    applicationsMatched: matched,
    applicationsUpdated: updated,
    syncedAt: new Date(now).toISOString()
  };
}

function getGmailEvents_(p) {
  const limit = Math.min(Number(p.limit || 25), 100);
  const events = rowsToObjects_(SHEETS.EMAIL_EVENTS);
  return {
    ok: true,
    events: events.slice(Math.max(0, events.length - limit)).reverse()
  };
}


// ═══════════════════════════════════════════════════════════
// JOB ANALYSIS
// ═══════════════════════════════════════════════════════════

function analyzeJob_(p) {
  const url = String(p.url || "").trim();

  if (!url) throw new Error("Please paste a job URL.");
  if (!/^https?:\/\//i.test(url)) {
    throw new Error("Please use a full job URL.");
  }

  let html = "";
  try {
    const r = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { "User-Agent": "Mozilla/5.0 CareerOS" }
    });

    if (r.getResponseCode() < 200 || r.getResponseCode() >= 400) {
      throw new Error("HTTP " + r.getResponseCode());
    }

    html = r.getContentText();
  } catch (err) {
    return {
      ok: false,
      needsJdPaste: true,
      url: url,
      error: "This site blocked automatic JD reading. Paste the JD below."
    };
  }

  const jd = cleanCareerJobHtml_(html);

  if (jd.length < 150) {
    return {
      ok: false,
      needsJdPaste: true,
      url: url,
      error: "The page did not expose enough readable job text. Paste the JD below."
    };
  }

  return rankCareerResumes_(url, jd);
}

function analyzePastedJob_(p) {
  const url = String(p.url || "").trim();
  const jd = String(p.jd || "").trim();

  if (jd.length < 50) {
    throw new Error("Please paste a fuller job description.");
  }

  return rankCareerResumes_(url, jd);
}

function cleanCareerJobHtml_(html) {
  let t = String(html || "");

  t = t
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<noscript[\s\S]*?<\/noscript>/gi, " ")
    .replace(/<svg[\s\S]*?<\/svg>/gi, " ")
    .replace(/<!--[\s\S]*?-->/g, " ")
    .replace(/<\/(p|div|section|article|li|h1|h2|h3|h4|br)>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">");

  return t.replace(/\s+/g, " ").trim().slice(0, 30000);
}

function careerKeywords_(text) {
  const stop = {
    the:1,and:1,for:1,with:1,from:1,this:1,that:1,your:1,our:1,
    are:1,will:1,have:1,has:1,into:1,about:1,role:1,job:1,work:1,
    team:1,they:1,who:1,what:1,how:1,all:1,any:1,can:1,may:1,
    more:1,than:1,not:1,using:1,use:1,years:1,year:1,experience:1,
    skills:1,responsibilities:1,requirements:1,preferred:1
  };

  const out = {};

  String(text || "")
    .toLowerCase()
    .replace(/[^a-z0-9+#./ -]/g, " ")
    .split(/[\s,;|()]+/)
    .forEach(function(w) {
      if (w.length >= 3 && !stop[w]) out[w] = true;
    });

  return Object.keys(out);
}

function careerOverlap_(a, b) {
  if (!a.length || !b.length) return 0;

  const set = {};
  b.forEach(function(x) { set[x] = true; });

  let hits = 0;
  a.forEach(function(x) {
    if (set[x]) hits++;
  });

  return Math.min(
    100,
    Math.round((hits / Math.min(a.length, 80)) * 100)
  );
}

function rankCareerResumes_(url, jd) {
  const resumes = rowsToObjects_(SHEETS.RESUME_LIBRARY);

  if (!resumes.length) {
    return {
      ok: true,
      url: url,
      jd: jd,
      match: 0,
      bestResume: null,
      resumes: [],
      recommendation: "Upload at least one resume first."
    };
  }

  const jdWords = careerKeywords_(jd);

  const ranked = resumes.map(function(r) {

    const corpus = [
      r.Name || "",
      r.Type || "",
      r.TargetRole || "",
      r.TargetCompany || "",
      r.Notes || ""
    ].join(" ");

    return {
      id: r.ID,
      name: r.Name,
      targetRole: r.TargetRole || "",
      targetCompany: r.TargetCompany || "",
      score: careerOverlap_(
        jdWords,
        careerKeywords_(corpus)
      )
    };

  }).sort(function(a,b) {
    return b.score - a.score;
  });

  const best = ranked[0] || null;

  const recommendation =
    !best ? "Upload a resume first." :
    best.score >= 85 ? "Strong fit — use this resume." :
    best.score >= 70 ? "Good fit — minor tailoring may help." :
    "Weak fit — review the role carefully.";

  const analysisId =
    Utilities.getUuid();

  const now =
    new Date().toISOString();

  getSheet_(
    SHEETS.JOB_ANALYSES
  ).appendRow([
    analysisId,
    url,
    "",
    "",
    jd,
    best ? best.score : "",
    best ? best.id : "",
    best ? best.name : "",
    recommendation,
    now,
    now
  ]);

  return {
    ok: true,
    id: analysisId,
    url: url,
    jd: jd,
    match: best ? best.score : 0,
    bestResume: best,
    resumes: ranked,
    recommendation: recommendation,
    reliability: "CareerOS V1.3 uses deterministic ranking. Treat the percentage as a prioritization score, not a hiring probability."
  };
}


// ═══════════════════════════════════════════════════════════
// CAREEROS V1.5 — RESUME TEXT INDEXING
//
// Uses the Google Drive Advanced Service to import PDF/DOC/DOCX/TXT
// into a temporary Google Doc, reads the extracted text, then removes
// the temporary Doc. Enable: Services → + → Drive API.
// ═══════════════════════════════════════════════════════════

function extractResumeText_(fileId) {
  if (typeof Drive === "undefined" || !Drive.Files || !Drive.Files.create) {
    throw new Error(
      "Resume text extraction requires the Google Drive advanced service. " +
      "In Apps Script, add Services → Drive API, then authorize again."
    );
  }

  const original = DriveApp.getFileById(fileId);
  const blob = original.getBlob();
  const ext = String(original.getName() || "").toLowerCase().split(".").pop();
  const isText = ext === "txt";

  if (isText) {
    return blob.getDataAsString("UTF-8").trim();
  }

  const tempMetadata = {
    name: "CareerOS__INDEX__" + original.getName(),
    mimeType: "application/vnd.google-apps.document"
  };

  let tempDoc = null;

  try {
    tempDoc = Drive.Files.create(
      tempMetadata,
      blob,
      {
        fields: "id,name,mimeType",
        ocrLanguage: "en"
      }
    );

    Utilities.sleep(500);

    const text = DocumentApp
      .openById(tempDoc.id)
      .getBody()
      .getText()
      .trim();

    return text;
  } finally {
    if (tempDoc && tempDoc.id) {
      try {
        Drive.Files.remove(tempDoc.id);
      } catch (err) {
        try {
          DriveApp.getFileById(tempDoc.id).setTrashed(true);
        } catch (ignored) {}
      }
    }
  }
}

function getResumeLibraryHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const map = {};
  headers.forEach(function(h, i) { map[h] = i + 1; });
  return map;
}

function indexResume_(resumeId) {
  const sheet = getSheet_(SHEETS.RESUME_LIBRARY);
  const data = sheet.getDataRange().getValues();
  const header = getResumeLibraryHeaderMap_(sheet);

  let rowNumber = 0;
  let row = null;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(resumeId)) {
      rowNumber = i + 1;
      row = data[i];
      break;
    }
  }

  if (!rowNumber) throw new Error("Resume not found: " + resumeId);

  const fileId = String(row[(header.FileId || 4) - 1] || "");
  if (!fileId) throw new Error("Resume has no Drive file ID.");

  const now = new Date().toISOString();

  sheet.getRange(rowNumber, header.IndexStatus).setValue("Indexing");

  try {
    const text = extractResumeText_(fileId);

    sheet.getRange(rowNumber, header.ExtractedText).setValue(text);
    sheet.getRange(rowNumber, header.IndexedAt).setValue(now);
    sheet.getRange(rowNumber, header.IndexStatus).setValue("Indexed");

    return {
      ok: true,
      id: resumeId,
      chars: text.length,
      indexedAt: now
    };
  } catch (err) {
    sheet.getRange(rowNumber, header.IndexedAt).setValue(now);
    sheet.getRange(rowNumber, header.IndexStatus).setValue("Error");
    throw err;
  }
}

function indexAllResumes_() {
  const resumes = rowsToObjects_(SHEETS.RESUME_LIBRARY);
  const results = [];
  let indexed = 0;
  let failed = 0;

  resumes.forEach(function(r) {
    const text = String(r.ExtractedText || "").trim();
    const status = String(r.IndexStatus || "");

    if (text && status === "Indexed") {
      results.push({id: r.ID, name: r.Name, status: "Already indexed"});
      return;
    }

    try {
      const result = indexResume_(r.ID);
      indexed++;
      results.push({id: r.ID, name: r.Name, status: "Indexed", chars: result.chars});
    } catch (err) {
      failed++;
      results.push({id: r.ID, name: r.Name, status: "Error", error: err.message});
    }
  });

  return {
    ok: failed === 0,
    indexed: indexed,
    failed: failed,
    results: results
  };
}


// ═══════════════════════════════════════════════════════════
// CAREEROS V1.5 — EVIDENCE-BASED RESUME MATCHER
// Uses actual extracted resume text. Shows all active resumes.
// ═══════════════════════════════════════════════════════════

const CAREER_SYNONYMS = {
  "gtm": ["go to market","go-to-market","gtm","g2m"],
  "product marketing": ["product marketing","product marketer","pmm"],
  "growth marketing": ["growth marketing","growth"],
  "acquisition": ["acquisition","customer acquisition","user acquisition","performance marketing","paid media"],
  "lifecycle": ["lifecycle","crm","retention","churn","customer lifecycle"],
  "retention": ["retention","churn","lifecycle"],
  "crm": ["crm","clevertap","hubspot","salesforce","lifecycle"],
  "performance marketing": ["performance marketing","paid marketing","paid media","meta ads","google ads"],
  "analytics": ["analytics","google analytics","mixpanel","sql","superset"],
  "experimentation": ["a/b testing","ab testing","a/b","multivariate","experimentation","testing"],
  "conversion": ["conversion","funnel","conversion rate","cvr"],
  "b2b": ["b2b","business-to-business","enterprise"],
  "d2c": ["d2c","direct-to-consumer","consumer","b2c"],
  "fintech": ["fintech","financial services","banking","payments","lending","credit cards","bfsI"],
  "partnerships": ["partnerships","partner","alliances","api integrations","integrations"],
  "positioning": ["positioning","messaging","value proposition","brand messaging"],
  "research": ["market research","customer research","consumer research","research"],
  "sales enablement": ["sales enablement","enablement","sales collateral"],
  "pricing": ["pricing","monetization","packaging"],
  "seo": ["seo","search engine optimization"],
  "content": ["content marketing","content strategy","copywriting"],
  "product strategy": ["product strategy","product management","roadmap","product development"],
  "customer journey": ["customer journey","journey","onboarding","activation"]
};

function normalizeCareerTerm_(s) {
  return String(s || "").toLowerCase().replace(/[^a-z0-9+#./ -]/g, " ").replace(/\s+/g, " ").trim();
}

function expandCareerConcepts_(text) {
  const t = normalizeCareerTerm_(text);
  const concepts = {};
  Object.keys(CAREER_SYNONYMS).forEach(function(c) {
    CAREER_SYNONYMS[c].forEach(function(term) {
      if (t.indexOf(term) >= 0) concepts[c] = true;
    });
  });
  return Object.keys(concepts);
}

function jdRequirementSentences_(jd) {
  const sentences = String(jd || "")
    .replace(/\r/g, "\n")
    .split(/[\n.!?]+/)
    .map(function(s){return s.trim();})
    .filter(function(s){return s.length >= 8;});

  return sentences.filter(function(s) {
    return /required|requirements|must|should|you will|responsibilit|experience in|experience with|looking for|preferred|qualifications|skills/i.test(s);
  }).slice(0, 80);
}

function scoreResumeEvidence_(jd, resumeText, metadata) {
  const jdText = normalizeCareerTerm_(jd);
  const resume = normalizeCareerTerm_(resumeText);
  const meta = normalizeCareerTerm_(metadata);

  const jdConcepts = expandCareerConcepts_(jdText);
  const resumeConcepts = expandCareerConcepts_(resume + " " + meta);

  const conceptSet = {};
  resumeConcepts.forEach(function(c){ conceptSet[c] = true; });

  let conceptHits = 0;
  const conceptMatched = [];
  jdConcepts.forEach(function(c){
    if (conceptSet[c]) {
      conceptHits++;
      conceptMatched.push(c);
    }
  });

  const sentences = jdRequirementSentences_(jd);
  let reqTotal = 0;
  let reqHit = 0;
  const matchedRequirements = [];
  const missingRequirements = [];

  sentences.forEach(function(s) {
    const words = careerKeywords_(s);
    if (!words.length) return;
    reqTotal++;

    const hitWords = words.filter(function(w){
      return resume.indexOf(w) >= 0;
    });

    const ratio = hitWords.length / words.length;
    if (ratio >= 0.22) {
      reqHit++;
      matchedRequirements.push(s);
    } else if (ratio < 0.08) {
      missingRequirements.push(s);
    }
  });

  const jdKeywords = careerKeywords_(jdText);
  const resumeKeywords = careerKeywords_(resume);
  const keywordScore = careerOverlapWeighted_(jdKeywords, resumeKeywords);

  // Weight the evidence: requirements > concepts > broad keyword coverage.
  const reqScore = reqTotal ? Math.round((reqHit / reqTotal) * 100) : 0;
  const conceptScore = jdConcepts.length ? Math.round((conceptHits / jdConcepts.length) * 100) : 0;

  let score = Math.round(
    reqScore * 0.48 +
    conceptScore * 0.27 +
    keywordScore * 0.25
  );

  // Small title/target-role bonus, capped to keep it from dominating.
  const targetRole = normalizeCareerTerm_(metadata);
  if (targetRole) {
    const roleWords = careerKeywords_(targetRole);
    const jdRoleHits = roleWords.filter(function(w){return jdText.indexOf(w)>=0;}).length;
    if (jdRoleHits >= 2) score = Math.min(100, score + 4);
  }

  return {
    score: Math.max(0, Math.min(100, score)),
    matchedConcepts: conceptMatched.slice(0, 12),
    matchedRequirements: matchedRequirements.slice(0, 6),
    missingRequirements: missingRequirements.slice(0, 6),
    requirementScore: reqScore,
    conceptScore: conceptScore,
    keywordScore: keywordScore
  };
}

function careerOverlapWeighted_(a, b) {
  if (!a.length || !b.length) return 0;
  const set = {};
  b.forEach(function(x){set[x]=true;});
  let hits = 0;
  let weightedTotal = 0;

  a.forEach(function(x) {
    const w = x.length >= 7 ? 1.25 : 1;
    weightedTotal += w;
    if (set[x]) hits += w;
  });

  return weightedTotal ? Math.round((hits / weightedTotal) * 100) : 0;
}

function rankCareerResumesV15_(url, jd) {
  let resumes = rowsToObjects_(SHEETS.RESUME_LIBRARY);

  if (!resumes.length) {
    return {
      ok: true,
      url: url,
      jd: jd,
      match: 0,
      bestResume: null,
      resumes: [],
      recommendation: "Upload at least one resume first."
    };
  }

  // Index every resume that does not have extracted text.
  resumes.forEach(function(r) {
    const text = String(r.ExtractedText || "").trim();
    if (!text || String(r.IndexStatus || "") !== "Indexed") {
      try {
        indexResume_(r.ID);
      } catch (err) {
        r.IndexError = err.message;
      }
    }
  });

  // Re-read after indexing.
  resumes = rowsToObjects_(SHEETS.RESUME_LIBRARY);

  const ranked = resumes.map(function(r) {
    const text = String(r.ExtractedText || "");
    const metadata = [
      r.Name || "",
      r.Type || "",
      r.TargetRole || "",
      r.TargetCompany || "",
      r.Notes || ""
    ].join(" ");

    const evidence = scoreResumeEvidence_(jd, text, metadata);

    return {
      id: r.ID,
      name: r.Name,
      targetRole: r.TargetRole || "",
      targetCompany: r.TargetCompany || "",
      indexStatus: r.IndexStatus || "Not indexed",
      indexError: r.IndexError || "",
      score: evidence.score,
      requirementScore: evidence.requirementScore,
      conceptScore: evidence.conceptScore,
      keywordScore: evidence.keywordScore,
      strengths: evidence.matchedConcepts,
      matchedRequirements: evidence.matchedRequirements,
      gaps: evidence.missingRequirements
    };
  }).sort(function(a,b){
    return b.score - a.score;
  });

  const best = ranked[0] || null;

  const recommendation =
    !best ? "Upload a resume first." :
    best.score >= 85 ? "Strong fit — use this resume as your starting point." :
    best.score >= 72 ? "Good fit — minor tailoring may help." :
    best.score >= 60 ? "Possible fit — review the gaps before applying." :
    "Weak fit — another role/resume may be better.";

  const analysisId = Utilities.getUuid();
  const now = new Date().toISOString();

  getSheet_(SHEETS.JOB_ANALYSES).appendRow([
    analysisId, url, "", "", jd,
    best ? best.score : "",
    best ? best.id : "",
    best ? best.name : "",
    recommendation,
    best ? best.strengths.join(", ") : "",
    best ? best.gaps.join(" | ") : "",
    now, now
  ]);

  return {
    ok: true,
    id: analysisId,
    url: url,
    jd: jd,
    match: best ? best.score : 0,
    bestResume: best,
    resumes: ranked,
    recommendation: recommendation,
    reliability: "V1.5 compares the actual extracted resume text across every active library entry. It is a deterministic evidence score, not an AI hiring probability."
  };
}


function analyzeJobV15_(p) {
  const url = String(p.url || "").trim();
  if (!url) throw new Error("Please paste a job URL.");
  if (!/^https?:\/\//i.test(url)) throw new Error("Please use a full job URL.");

  let html = "";
  try {
    const r = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {"User-Agent":"Mozilla/5.0 CareerOS"}
    });

    if (r.getResponseCode() < 200 || r.getResponseCode() >= 400) {
      throw new Error("HTTP " + r.getResponseCode());
    }

    html = r.getContentText();
  } catch (err) {
    return {
      ok:false,
      needsJdPaste:true,
      url:url,
      error:"This site blocked automatic JD reading. Paste the JD below instead."
    };
  }

  const jd = cleanCareerJobHtml_(html);

  if (jd.length < 150) {
    return {
      ok:false,
      needsJdPaste:true,
      url:url,
      error:"The page did not expose enough readable job text. Paste the JD below instead."
    };
  }

  return rankCareerResumesV15_(url, jd);
}

function analyzePastedJobV15_(p) {
  const url = String(p.url || "").trim();
  const jd = String(p.jd || "").trim();
  if (jd.length < 50) throw new Error("Please paste a fuller job description.");
  return rankCareerResumesV15_(url, jd);
}


function syncGmailV15_(p) {
  const status = getGmailStatus_();
  if (!status.isExpected) {
    throw new Error(
      "CareerOS Gmail must run as vishalnihalani2@gmail.com. Current account: " +
      (status.account || "unknown")
    );
  }

  const props = PropertiesService.getScriptProperties();
  const last = Number(props.getProperty("CAREEROS_GMAIL_LAST_SYNC_MS") || "0");
  const now = Date.now();

  const after = new Date(Math.max(
    now - 14 * 24 * 60 * 60 * 1000,
    last - 5 * 60 * 1000
  ));

  const q =
    "after:" +
    after.getFullYear() + "/" +
    String(after.getMonth()+1).padStart(2,"0") + "/" +
    String(after.getDate()).padStart(2,"0") +
    " -in:chats " +
    "{application applied interview recruiter hiring assessment offer rejection regret \"next step\"}";

  let threads;
  try {
    threads = GmailApp.search(q, 0, 20);
  } catch (err) {
    throw new Error("Gmail search failed: " + err.message);
  }

  const eventsSheet = getSheet_(SHEETS.EMAIL_EVENTS);
  const existing = rowsToObjects_(SHEETS.EMAIL_EVENTS);
  const seen = {};
  existing.slice(Math.max(0, existing.length - 1000)).forEach(function(e) {
    if (e.MessageID) seen[String(e.MessageID)] = true;
  });

  const apps = rowsToObjects_(SHEETS.APPLICATIONS);

  let messagesProcessed = 0;
  let matched = 0;
  let updated = 0;

  threads.forEach(function(thread) {
    thread.getMessages().forEach(function(msg) {
      const id = String(msg.getId());
      const ts = msg.getDate().getTime();

      if (seen[id] || ts <= last) return;

      const subject = String(msg.getSubject() || "");
      const body = String(msg.getPlainBody() || "").slice(0, 4000);
      const text = (subject + "\n" + body).toLowerCase();

      let eventType =
        /rejected|regret|not selected|unsuccessful|decline|not moving forward/.test(text) ? "Rejected" :
        /offer|pleased to offer/.test(text) ? "Offer" :
        /interview|assessment|case study|technical round|schedule.*call/.test(text) ? "Interview" :
        /recruiter|talent acquisition|phone screen|screening/.test(text) ? "Recruiter Screen" :
        /application.*received|application.*submitted|successfully applied|thank.*applying/.test(text) ? "Applied" :
        "Job-related";

      let best = null;
      let bestScore = 0;

      apps.forEach(function(a) {
        const company = String(a.company || "").toLowerCase();
        const title = String(a.title || "").toLowerCase();
        let s = 0;

        if (company && text.indexOf(company) >= 0) s += 6;
        if (title && text.indexOf(title) >= 0) s += 6;

        title.split(/[^a-z0-9]+/)
          .filter(function(w){return w.length >= 5;})
          .slice(0, 10)
          .forEach(function(w){
            if (text.indexOf(w) >= 0) s += 1;
          });

        if (s > bestScore) {
          bestScore = s;
          best = a;
        }
      });

      if (best && bestScore >= 3) {
        matched++;

        if (["Applied","Recruiter Screen","Interview","Offer","Rejected"].indexOf(eventType) >= 0 &&
            eventType !== best.status) {

          const result = saveRecord_({
            sheet:"Applications",
            id:best.id,
            JobID:best.jobId || "",
            Company:best.company || "",
            Title:best.title || "",
            Status:eventType,
            DateApplied:best.dateApplied || (eventType === "Applied" ? msg.getDate().toISOString() : ""),
            ResumeID:best.resumeId || "",
            Recruiter:best.recruiter || "",
            NextAction:best.nextAction || "",
            FollowUpDate:best.followUpDate || "",
            Notes:best.notes || ""
          });

          if (result && result.ok) updated++;
        }
      }

      eventsSheet.appendRow([
        Utilities.getUuid(),
        id,
        msg.getDate().toISOString(),
        msg.getFrom() || "",
        subject,
        eventType,
        best ? best.company : "",
        best ? best.title : "",
        best ? best.id : "",
        best ? "matched" : "unmatched",
        body.slice(0, 800)
      ]);

      seen[id] = true;
      messagesProcessed++;
    });
  });

  props.setProperty("CAREEROS_GMAIL_LAST_SYNC_MS", String(now));

  return {
    ok:true,
    account:status.account,
    threadsScanned:threads.length,
    messagesProcessed:messagesProcessed,
    applicationsMatched:matched,
    applicationsUpdated:updated,
    syncedAt:new Date(now).toISOString()
  };
}

function migrateCareerOSV15_() {
  Object.keys(SHEETS).forEach(function(k) {
    getSheet_(SHEETS[k]);
  });
  return {
    ok: true,
    sheets: Object.values(SHEETS)
  };
}


function repairResumeLibraryHeaders_() {
  const sheet = getSheet_(SHEETS.RESUME_LIBRARY);
  ensureHeaders_(sheet, HEADERS[SHEETS.RESUME_LIBRARY]);
  return {
    ok: true,
    headers: sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
  };
}
