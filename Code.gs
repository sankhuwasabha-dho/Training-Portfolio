// ============================================================
//  Health Office Sankhuwasabha – Training Registration Backend
//  Two-sheet model:  submission_pending  →  submission_approved
//  Google Apps Script  |  Paste into script.google.com
// ============================================================

const PENDING_SHEET  = "submission_pending";
const APPROVED_SHEET = "submission_approved";

// Notification recipients (both will be emailed on every new submission)
const ADMIN_EMAILS = [
  "dpesh.stha2016@gmail.com",
  "dhosankhuwasabha@gmail.com"
];

// Column order — matches form.html field order exactly
const COLUMNS = [
  "submitted_at", "status",
  // Training Information
  "training_name", "role", "training_site", "training_province", "training_district",
  "start_date", "end_date", "fiscal_year",
  // Personal Information
  "name_english", "name_nepali", "sex", "sex_other", "dob_bs",
  // Permanent Address
  "perm_province", "perm_district", "perm_local_level", "perm_ward",
  "contact", "email",
  // Caste
  "caste", "caste_other",
  // Cadre
  "cadre", "cadre_other", "qualification",
  // Sponsored
  "sponsor", "sponsor_details",
  // Working Place
  "work_office", "work_district", "work_province", "work_local_level",
  "work_contact", "designation", "level",
  "pis_no", "citizenship", "council_reg"
];

const STATUS_COL = COLUMNS.indexOf("status") + 1; // 1-based column index

// ── HTTP entry points ──────────────────────────────────────
function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const action = (e && e.parameter && e.parameter.action) || "get";
  let result;
  try {
    if (action === "get")  result = getData();
    if (action === "add")  result = addRecord(e);
    if (action === "init") result = initSheets();
  } catch(err) {
    result = { error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Return all approved rows for public portfolio ──────────
function getData() {
  const sheet = getOrCreateSheet(APPROVED_SHEET);
  const rows  = sheet.getDataRange().getValues();
  if (rows.length < 2) return { data: [] };

  const headers = rows[0];
  const data = rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
  return { data };
}

// ── New submission → pending sheet ─────────────────────────
function addRecord(e) {
  let body = {};
  try {
    if (e.postData && e.postData.contents) body = JSON.parse(e.postData.contents);
    else if (e.parameter) body = e.parameter;
  } catch(_) { body = e.parameter || {}; }

  if (body.website || body.url_field) return { success: true }; // honeypot

  const sheet = getOrCreateSheet(PENDING_SHEET);
  const row   = COLUMNS.map(col => {
    if (col === "submitted_at") return new Date().toLocaleString("en-GB");
    if (col === "status")       return "Pending";
    return body[col] !== undefined ? body[col] : "";
  });

  sheet.appendRow(row);
  notifyAdmins(body);
  return { success: true };
}

// ── Email alert when a new submission arrives ──────────────
function notifyAdmins(data) {
  try {
    const subject = "New Training Registration – " + (data.name_english || "Unknown");
    const body = [
      "A new training registration has been submitted and is awaiting approval.",
      "",
      "Name        : " + (data.name_english || "–"),
      "Designation : " + (data.designation || "–"),
      "Office      : " + (data.work_office || "–"),
      "Local Level : " + (data.work_local_level || "–"),
      "Training    : " + (data.training_name || "–"),
      "Role        : " + (data.role || "–"),
      "Dates       : " + (data.start_date || "–") + " to " + (data.end_date || "–"),
      "Contact     : " + (data.contact || "–"),
      "Email       : " + (data.email || "–"),
      "",
      "To approve:",
      "1. Open the Google Sheet → submission_pending tab.",
      "2. Find this row.",
      "3. Change Status from 'Pending' to 'Approved'.",
      "",
      "The row will automatically move to submission_approved and appear on the public portfolio."
    ].join("\n");

    MailApp.sendEmail(ADMIN_EMAILS.join(","), subject, body);
  } catch(_) {}
}

// ── Auto-move row when status flipped to Approved ──────────
// Simple onEdit trigger fires automatically; no setup needed.
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== PENDING_SHEET) return;

  const row = e.range.getRow();
  if (row < 2) return; // header

  if (e.range.getColumn() !== STATUS_COL) return;
  const newValue = String(e.value || "").trim().toLowerCase();
  if (newValue !== "approved") return;

  // Read full row
  const rowData = sheet.getRange(row, 1, 1, COLUMNS.length).getValues()[0];
  // Force status to "Approved" (canonical capitalization)
  rowData[STATUS_COL - 1] = "Approved";

  // Append to approved sheet
  const approved = getOrCreateSheet(APPROVED_SHEET);
  approved.appendRow(rowData);

  // Delete row from pending
  sheet.deleteRow(row);
}

// ── Sheet helpers ──────────────────────────────────────────
function getOrCreateSheet(name) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
    styleHeader(sheet);
  }
  return sheet;
}

function styleHeader(sheet) {
  const hr = sheet.getRange(1, 1, 1, COLUMNS.length);
  hr.setBackground("#1a3c6e");
  hr.setFontColor("#ffffff");
  hr.setFontWeight("bold");
  sheet.setFrozenRows(1);
  sheet.getRange(1, STATUS_COL).setBackground("#c0392b");
  sheet.setColumnWidth(STATUS_COL, 100);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Pending", "Approved", "Rejected"], true)
    .build();
  sheet.getRange(2, STATUS_COL, 1000, 1).setDataValidation(rule);
}

function initSheets() {
  getOrCreateSheet(PENDING_SHEET);
  getOrCreateSheet(APPROVED_SHEET);
  return { success: true, message: "Initialized " + PENDING_SHEET + " and " + APPROVED_SHEET };
}

// ============================================================
//  ONE-TIME MIGRATION FROM OLD "Submissions" TAB → TWO SHEETS
// ============================================================
// USAGE: in the Apps Script editor, select function "migrateOnce"
//        from the dropdown, click Run, authorize if asked.
//        After it completes, you may delete the old "Submissions"
//        tab manually.
function migrateOnce() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const old = ss.getSheetByName("Submissions");
  if (!old) {
    SpreadsheetApp.getUi().alert("No 'Submissions' tab found. Nothing to migrate.");
    return;
  }

  const data = old.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert("'Submissions' tab is empty. Nothing to migrate.");
    return;
  }

  const headers = data[0];
  const statusIdx = headers.indexOf("status");
  if (statusIdx === -1) {
    SpreadsheetApp.getUi().alert("Couldn't find 'status' column in Submissions header.");
    return;
  }

  const pending  = getOrCreateSheet(PENDING_SHEET);
  const approved = getOrCreateSheet(APPROVED_SHEET);

  const pendingRows  = [];
  const approvedRows = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = String(row[statusIdx] || "").trim().toLowerCase();
    if (status === "approved")     approvedRows.push(row);
    else if (status === "pending") pendingRows.push(row);
    // Rejected / blank / other → skip (left in old sheet for review)
  }

  if (approvedRows.length > 0) {
    approved.getRange(approved.getLastRow() + 1, 1, approvedRows.length, COLUMNS.length).setValues(approvedRows);
  }
  if (pendingRows.length > 0) {
    pending.getRange(pending.getLastRow() + 1, 1, pendingRows.length, COLUMNS.length).setValues(pendingRows);
  }

  SpreadsheetApp.getUi().alert(
    "Migration complete.\n\n" +
    approvedRows.length + " rows → " + APPROVED_SHEET + "\n" +
    pendingRows.length  + " rows → " + PENDING_SHEET + "\n\n" +
    "You can now delete the old 'Submissions' tab if everything looks correct."
  );
}

// ============================================================
//  ONE-SHOT CSV IMPORT (historical bulk data, status=Approved)
// ============================================================
// USAGE:
//   1. Create a tab named exactly "RawImport" in this Sheet.
//   2. File → Import → Upload → "training information.csv" →
//      Import location: Replace current sheet → Comma separator.
//   3. In Apps Script editor, run function "importHistoricalCsv".
//   4. After completion you may delete the RawImport tab.
function importHistoricalCsv() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName("RawImport");
  if (!src) throw new Error('Create a sheet named "RawImport" and paste the CSV contents there first.');

  const data = src.getDataRange().getValues();
  if (data.length < 2) throw new Error("RawImport tab is empty.");

  const srcHeaders = data[0];
  const idx = {};
  srcHeaders.forEach((h, i) => { idx[String(h).trim()] = i; });

  const TRAININGS = [
    "PEN","CNSI","HMIS","DHIS2","STP","ENT","Implant","IUCD","CoFP",
    "MA","SBA","RUSG","Immunization","TB modular","PMTCT","PAMSv2","CBIMNCI","Mental Health"
  ];

  const dest = getOrCreateSheet(APPROVED_SHEET);
  const now  = "Historical Import";
  const newRows = [];

  for (let i = 1; i < data.length; i++) {
    const row  = data[i];
    const name = String(row[idx["Name of HW"]] || "").trim();
    if (!name) continue;

    const baseRow = {
      submitted_at:      now,
      status:            "Approved",
      role:              "Participant",
      training_province: "Koshi",
      training_district: "Sankhuwasabha",
      name_english:      name,
      perm_province:     "Koshi",
      perm_district:     "Sankhuwasabha",
      perm_local_level:  String(row[idx["Local Level"]] || ""),
      contact:           String(row[idx["Contact No"]]  || ""),
      email:             String(row[idx["Email ID"]]    || ""),
      qualification:     String(row[idx["Qualification"]] || ""),
      work_office:       String(row[idx["HF Name"]]     || ""),
      work_district:     "Sankhuwasabha",
      work_province:     "Koshi",
      work_local_level:  String(row[idx["Local Level"]] || ""),
      designation:       String(row[idx["Post"]]        || ""),
      level:             String(row[idx["Level"]]       || "")
    };

    TRAININGS.forEach(t => {
      const val = String(row[idx[t]] || "").toLowerCase().trim();
      if (val !== "yes") return;
      const merged = Object.assign({}, baseRow, { training_name: t });
      const arr = COLUMNS.map(c => merged[c] !== undefined ? merged[c] : "");
      newRows.push(arr);
    });
  }

  if (newRows.length === 0) {
    SpreadsheetApp.getUi().alert("No 'Yes'-marked trainings found in RawImport. Nothing imported.");
    return;
  }

  const startRow = dest.getLastRow() + 1;
  dest.getRange(startRow, 1, newRows.length, COLUMNS.length).setValues(newRows);

  SpreadsheetApp.getUi().alert("Imported " + newRows.length + " approved rows into " + APPROVED_SHEET + ".");
}
