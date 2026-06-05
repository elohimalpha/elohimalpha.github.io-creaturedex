// CREATURÉDEX — Google Apps Script sync backend
// ─────────────────────────────────────────────
// Setup:
//   1. Open your Google Sheet
//   2. Extensions → Apps Script → paste this whole file
//   3. Deploy → New deployment → Web app
//      - Execute as: Me
//      - Who has access: Anyone
//   4. Copy the deployment URL into the app's ⚙ config

const SHEET_NAME  = "creaturedex";
const FOLDER_NAME = "CREATURÉDEX";

// ── Helpers ───────────────────────────────────────────────────────────────────
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["key", "data", "updated_at"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateFolder() {
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(FOLDER_NAME);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Handle GET (pull) ─────────────────────────────────────────────────────────
function doGet(e) {
  const action = e?.parameter?.action;
  if (action === "pull") return handlePull();
  return jsonResponse({ status: "ok", message: "CREATURÉDEX sync endpoint running" });
}

// ── Handle POST (push / uploadImage) ─────────────────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === "push")        return handlePush(body.data);
    if (body.action === "uploadImage") return handleImageUpload(body);
    return jsonResponse({ status: "error", message: "Unknown action" });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

// ── Push creature data ────────────────────────────────────────────────────────
function handlePush(data) {
  try {
    const sheet = getOrCreateSheet();
    const json  = JSON.stringify(data);
    const now   = new Date().toISOString();

    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === "main") {
        sheet.getRange(i + 1, 2).setValue(json);
        sheet.getRange(i + 1, 3).setValue(now);
        return jsonResponse({ status: "ok", message: "Updated" });
      }
    }

    sheet.appendRow(["main", json, now]);
    return jsonResponse({ status: "ok", message: "Created" });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

// ── Pull creature data ────────────────────────────────────────────────────────
function handlePull() {
  try {
    const sheet  = getOrCreateSheet();
    const values = sheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === "main") {
        const data = JSON.parse(values[i][1]);
        return jsonResponse({ status: "ok", data });
      }
    }

    return jsonResponse({ status: "error", message: "No data found — push first from any device." });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

// ── Upload image to Drive, return public URL ──────────────────────────────────
function handleImageUpload(body) {
  try {
    const { filename, mimeType, base64 } = body;
    if (!filename || !mimeType || !base64) {
      return jsonResponse({ status: "error", message: "Missing filename, mimeType, or base64" });
    }

    const folder  = getOrCreateFolder();
    const bytes   = Utilities.base64Decode(base64);
    const blob    = Utilities.newBlob(bytes, mimeType, filename);
    const file    = folder.createFile(blob);

    // Make the file publicly readable so the app can display it
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Convert to a direct-display URL (bypasses the Drive preview page)
    const fileId  = file.getId();
    const url     = `https://drive.google.com/thumbnail?id=${fileId}&sz=w800`;

    return jsonResponse({ status: "ok", url, fileId });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}
